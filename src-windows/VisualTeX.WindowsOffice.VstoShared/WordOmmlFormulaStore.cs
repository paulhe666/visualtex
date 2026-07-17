using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static class WordOmmlFormulaStore
{
    internal const string NamespaceUri = "urn:visualtex:word-omml:1";
    internal const string BookmarkPrefix = "VTOMML_";

    private static readonly XNamespace VisualTeXNamespace = NamespaceUri;

    internal static string BookmarkName(string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var parsed))
            throw new InvalidDataException("VisualTeX OMML formulaId must be a UUID.");
        return BookmarkPrefix + parsed.ToString("N");
    }

    internal static bool TryGetFormulaId(Bookmark? bookmark, out string formulaId)
    {
        formulaId = string.Empty;
        if (bookmark is null) return false;
        string name;
        try { name = bookmark.Name ?? string.Empty; }
        catch { return false; }
        if (!name.StartsWith(BookmarkPrefix, StringComparison.Ordinal)) return false;
        var candidate = name.Substring(BookmarkPrefix.Length);
        if (!Guid.TryParseExact(candidate, "N", out var parsed)) return false;
        formulaId = parsed.ToString();
        return true;
    }

    internal static Bookmark? FindAtRange(Document document, Range selectionRange)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            Bookmark? best = null;
            var bestLength = int.MaxValue;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? equationRange = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryGetFormulaId(bookmark, out _)) continue;
                    try { equationRange = GetEquationRange(bookmark); }
                    catch { continue; }
                    var containsCaret = selectionRange.Start == selectionRange.End
                        && selectionRange.Start >= equationRange.Start
                        && selectionRange.Start <= equationRange.End;
                    var overlaps = selectionRange.Start < equationRange.End
                        && selectionRange.End > equationRange.Start;
                    if (!containsCaret && !overlaps) continue;
                    var length = Math.Max(0, equationRange.End - equationRange.Start);
                    if (length >= bestLength) continue;
                    Release(best);
                    best = bookmark;
                    bookmark = null;
                    bestLength = length;
                }
                finally
                {
                    Release(equationRange);
                    Release(bookmark);
                }
            }
            return best;
        }
        finally { Release(bookmarks); }
    }

    internal static Bookmark? FindByFormulaId(Document document, string formulaId)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = BookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return null;
            return bookmarks[name];
        }
        finally { Release(bookmarks); }
    }

    internal static IReadOnlyList<string> FormulaIds(Document document)
    {
        var result = new List<string>();
        var staleFormulaIds = new List<string>();
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? bookmarkRange = null;
                Range? equationRange = null;
                OMaths? maths = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryGetFormulaId(bookmark, out var formulaId)) continue;
                    bookmarkRange = bookmark.Range;
                    try
                    {
                        equationRange = GetEquationRange(bookmark);
                        maths = equationRange.OMaths;
                        var anchorDistance = equationRange.Start - bookmarkRange.Start;
                        if (maths.Count == 1 && anchorDistance >= 0 && anchorDistance <= 1)
                            result.Add(formulaId);
                        else
                            staleFormulaIds.Add(formulaId);
                    }
                    catch
                    {
                        // Deleting an OMML formula with Word's Delete key leaves
                        // its collapsed VisualTeX bookmark and custom XML part.
                        // Treat that anchor as stale so Update Equation Numbers
                        // can remove the old visible number and renumber the rest.
                        staleFormulaIds.Add(formulaId);
                    }
                }
                finally
                {
                    Release(maths);
                    Release(equationRange);
                    Release(bookmarkRange);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }

        foreach (var formulaId in staleFormulaIds.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            Bookmark? staleBookmark = null;
            try
            {
                staleBookmark = FindByFormulaId(document, formulaId);
                staleBookmark?.Delete();
                Delete(document, formulaId);
            }
            catch { }
            finally { Release(staleBookmark); }
        }
        return result;
    }

    internal static FormulaMetadata? TryRead(Document document, Bookmark bookmark)
    {
        if (!TryGetFormulaId(bookmark, out var formulaId)) return null;
        return TryRead(document, formulaId);
    }

    internal static FormulaMetadata? TryRead(Document document, string formulaId)
    {
        object? part = null;
        try
        {
            part = FindPart(document, formulaId);
            if (part is null) return null;
            var partXml = (string?)((dynamic)part).XML;
            return TryDecodePartXml(partXml, out var metadata)
                && string.Equals(metadata.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase)
                    ? metadata
                    : null;
        }
        catch
        {
            return null;
        }
        finally { Release(part); }
    }

    internal static void Save(Document document, FormulaMetadata metadata)
    {
        metadata.Validate();
        var xml = BuildPartXml(metadata);
        object? existing = null;
        object? parts = null;
        object? added = null;
        try
        {
            existing = FindPart(document, metadata.FormulaId);
            parts = ((dynamic)document).CustomXMLParts;
            added = ((dynamic)parts).Add(xml);
            if (added is null)
                throw new InvalidOperationException("Word did not create the VisualTeX OMML metadata part.");

            // Word may reject CustomXMLPart.LoadXML after an OMath rebuild.
            // Add the replacement first, then remove the old part so a failed
            // update never destroys the last valid metadata copy.
            if (existing is not null) ((dynamic)existing).Delete();
        }
        finally
        {
            Release(added);
            Release(parts);
            Release(existing);
        }
    }

    internal static void Delete(Document document, string formulaId)
    {
        object? part = null;
        try
        {
            part = FindPart(document, formulaId);
            if (part is not null) ((dynamic)part).Delete();
        }
        finally { Release(part); }
    }

    internal static Bookmark Wrap(
        Document document,
        Range equationRange,
        FormulaMetadata metadata)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? anchorRange = null;
        Range? preceding = null;
        try
        {
            var anchorPosition = equationRange.Start;
            if (anchorPosition > 0)
            {
                object precedingStart = anchorPosition - 1;
                object precedingEnd = anchorPosition;
                preceding = document.Range(ref precedingStart, ref precedingEnd);
                if (string.Equals(preceding.Text, "\v", StringComparison.Ordinal))
                    anchorPosition--;
            }

            object anchorStart = anchorPosition;
            object anchorEnd = anchorPosition;
            anchorRange = document.Range(ref anchorStart, ref anchorEnd);
            bookmarks = document.Bookmarks;
            var name = BookmarkName(metadata.FormulaId);
            if (bookmarks.Exists(name)) bookmarks[name].Delete();
            bookmark = bookmarks.Add(name, anchorRange);
            var result = bookmark;
            bookmark = null;
            return result;
        }
        finally
        {
            Release(preceding);
            Release(anchorRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static Range GetEquationRange(Bookmark bookmark)
    {
        Range? bookmarkRange = null;
        Document? document = null;
        OMaths? maths = null;
        OMath? bestMath = null;
        Range? bestRange = null;
        try
        {
            bookmarkRange = bookmark.Range;
            document = bookmarkRange.Document;
            maths = document.OMaths;
            var anchor = bookmarkRange.Start;
            var bestDistance = int.MaxValue;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    var containsAnchor = anchor >= range.Start && anchor <= range.End;
                    var distance = containsAnchor ? 0 : range.Start - anchor;
                    if (distance < 0 || distance > 8 || distance >= bestDistance) continue;
                    Release(bestRange);
                    Release(bestMath);
                    bestRange = range.Duplicate;
                    bestMath = math;
                    math = null;
                    bestDistance = distance;
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
            if (bestRange is null)
                throw new InvalidDataException(
                    "The VisualTeX OMML anchor is no longer adjacent to a Word equation.");
            var result = bestRange;
            bestRange = null;
            return result;
        }
        finally
        {
            Release(bestRange);
            Release(bestMath);
            Release(maths);
            Release(document);
            Release(bookmarkRange);
        }
    }

    internal static float EstimateHeightPoints(Bookmark bookmark)
    {
        Range? equationRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            equationRange = GetEquationRange(bookmark);
            font = equationRange.Font;
            var size = 11f;
            try { size = font.Size; } catch { }
            if (float.IsNaN(size) || float.IsInfinity(size) || size <= 0 || size > 256)
                size = 11f;
            return Math.Max(11f, size * 1.5f);
        }
        finally
        {
            Release(font);
            Release(equationRange);
        }
    }

    internal static string BuildPartXml(FormulaMetadata metadata)
    {
        metadata.Validate();
        var encoded = FormulaMetadataCodec.Encode(metadata);
        return new XDocument(
            new XElement(
                VisualTeXNamespace + "formula",
                new XAttribute("formulaId", metadata.FormulaId),
                new XElement(VisualTeXNamespace + "metadata", encoded)))
            .ToString(SaveOptions.DisableFormatting);
    }

    internal static bool TryDecodePartXml(string? xml, out FormulaMetadata metadata)
    {
        metadata = null!;
        if (string.IsNullOrWhiteSpace(xml)) return false;
        try
        {
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreWhitespace = true,
            };
            using var stringReader = new StringReader(xml);
            using var reader = XmlReader.Create(stringReader, settings);
            var document = XDocument.Load(reader, LoadOptions.None);
            var root = document.Root;
            if (root?.Name != VisualTeXNamespace + "formula") return false;
            var formulaId = (string?)root.Attribute("formulaId");
            var encoded = root.Element(VisualTeXNamespace + "metadata")?.Value;
            var decoded = FormulaMetadataCodec.Decode(encoded);
            if (decoded is null
                || !string.Equals(decoded.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase))
                return false;
            decoded.Validate();
            metadata = decoded;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static object? FindPart(Document document, string formulaId)
    {
        object? parts = null;
        object? selected = null;
        try
        {
            parts = ((dynamic)document).CustomXMLParts;
            selected = ((dynamic)parts).SelectByNamespace(NamespaceUri);
            var count = (int)((dynamic)selected).Count;
            for (var index = 1; index <= count; index++)
            {
                object? part = null;
                try
                {
                    part = ((dynamic)selected)[index];
                    var partXml = (string?)((dynamic)part).XML;
                    if (!TryDecodePartXml(partXml, out var metadata)
                        || !string.Equals(
                            metadata.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    var result = part;
                    part = null;
                    return result;
                }
                finally { Release(part); }
            }
            return null;
        }
        finally
        {
            Release(selected);
            Release(parts);
        }
    }

    internal static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
