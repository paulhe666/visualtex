using System.Linq;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static class WordOmmlNativeSource
{
    internal static FormulaMetadata CreateForNative(
        Document document,
        Range equationRange)
    {
        var formulaId = Guid.NewGuid().ToString("D");
        var wordOpenXml = ReadCompleteEquationWordOpenXml(
            document,
            equationRange,
            formulaId);
        var displayMode = ReadDisplayMode(equationRange);
        var numbered = string.Equals(
                displayMode,
                "block",
                StringComparison.Ordinal)
            && WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                wordOpenXml);
        var semanticWordOpenXml = numbered
            ? WordOmmlConverter.StripManagedVisualTeXNativeEquationNumber(
                wordOpenXml)
            : wordOpenXml;
        var mathMl = WordOmmlConverter.TransformOmmlToMathMl(
            semanticWordOpenXml,
            display: string.Equals(displayMode, "block", StringComparison.Ordinal));
        var latex = SanitizeFormulaBoundaryArtifacts(
            MathMlToLatexConverter.Convert(mathMl));
        if (string.IsNullOrWhiteSpace(latex))
            throw new InvalidDataException(
                "The Word-native OMML equation could not be converted back to editable LaTeX.");

        var fontSize = ReadFontSize(equationRange);
        var now = DateTimeOffset.UtcNow.ToString("O");
        var metadata = new FormulaMetadata
        {
            FormulaId = formulaId,
            Title = "Word Formula",
            Latex = latex,
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
            },
            CodeFormat = "raw",
            DisplayMode = displayMode,
            Numbered = numbered,
            FontSizePt = fontSize,
            RenderFontSizePt = fontSize,
            NativeOmmlFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                semanticWordOpenXml),
            CreatedWithVersion = "1.2.5",
            UpdatedWithVersion = "1.2.5",
            CreatedAt = now,
            UpdatedAt = now,
        };
        metadata.Validate();
        return metadata;
    }

    internal static FormulaMetadata RefreshForVisualTeX(
        Document document,
        Bookmark bookmark,
        FormulaMetadata stored)
    {
        Range? equationRange = null;
        try
        {
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var sanitizedStored = Clone(stored);
            SanitizeMetadataBoundaryArtifacts(sanitizedStored);
            var wordOpenXml = ReadCompleteEquationWordOpenXml(
                document,
                equationRange,
                stored.FormulaId);
            var fingerprint = WordOmmlConverter.ComputeOmmlFingerprint(wordOpenXml);
            if (string.Equals(
                    stored.NativeOmmlFingerprint,
                    fingerprint,
                    StringComparison.OrdinalIgnoreCase))
                return sanitizedStored;

            var mathMl = WordOmmlConverter.TransformOmmlToMathMl(
                wordOpenXml,
                display: string.Equals(
                    stored.DisplayMode,
                    "block",
                    StringComparison.Ordinal));
            var latex = SanitizeFormulaBoundaryArtifacts(
                MathMlToLatexConverter.Convert(mathMl));
            if (string.IsNullOrWhiteSpace(latex))
                throw new InvalidDataException(
                    "The Word-native OMML equation could not be converted back to editable LaTeX.");

            var refreshed = Clone(stored);
            var lineId = refreshed.Lines.FirstOrDefault()?.Id;
            if (string.IsNullOrWhiteSpace(lineId)) lineId = Guid.NewGuid().ToString();
            refreshed.Latex = latex;
            refreshed.Lines = new List<FormulaLine>
            {
                new() { Id = lineId!, Latex = latex },
            };
            refreshed.CodeFormat = "raw";
            refreshed.NativeOmmlFingerprint = fingerprint;
            refreshed.Validate();

            // Older VisualTeX builds persisted the converter-side OMML
            // fingerprint before Word finished normalizing the native equation.
            // Once an explicit VisualTeX read has resolved the physical OMath,
            // persist the live fingerprint so a later VTOMML bookmark drift can
            // be recovered without relying on the old anchor coordinates.
            if (!document.ReadOnly)
            {
                try { WordOmmlFormulaStore.Save(document, refreshed); }
                catch
                {
                    // The current edit can still use the refreshed in-memory
                    // metadata. Commit-time range resolution also prefers this
                    // Session snapshot, so a transient CustomXML write refusal
                    // must not make the editor unavailable.
                }
            }
            return refreshed;
        }
        finally
        {
            Release(equationRange);
        }
    }

    internal static void StampFingerprint(FormulaMetadata metadata, Range equationRange)
    {
        Document? document = null;
        try
        {
            document = equationRange.Document;
            metadata.NativeOmmlFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                ReadCompleteEquationWordOpenXml(
                    document,
                    equationRange,
                    metadata.FormulaId));
        }
        finally { Release(document); }
    }

    internal static void StampFingerprintFromResolvedRange(
        FormulaMetadata metadata,
        Range equationRange)
    {
        // Callers use this only while they still own the complete live OMath
        // Range returned by WordOmmlConverter/OMath.Range. Avoid constructing a
        // second document probe and enumerating bookmarks merely to serialize the
        // same equation again; this path runs after Word's final normalization and
        // is important for large numbered/redraw workloads.
        metadata.NativeOmmlFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
            equationRange.WordOpenXML);
    }

    internal static int RefreshFingerprintsFromDocumentOpenXml(
        Document document,
        IReadOnlyCollection<string> formulaIds)
    {
        if (formulaIds is null || formulaIds.Count == 0) return 0;
        Range? content = null;
        try
        {
            content = document.Content;
            var wordOpenXml = content.WordOpenXML ?? string.Empty;
            if (string.IsNullOrWhiteSpace(wordOpenXml))
                throw new InvalidDataException(
                    "Word returned empty document XML while finalizing OMML fingerprints.");

            var package = XDocument.Parse(wordOpenXml, LoadOptions.PreserveWhitespace);
            XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var formulaByBookmark = formulaIds
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    WordOmmlFormulaStore.BookmarkName,
                    formulaId => formulaId,
                    StringComparer.OrdinalIgnoreCase);
            var fingerprints = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string? pendingFormulaId = null;

            foreach (var element in package.Descendants())
            {
                if (element.Name == word + "bookmarkStart")
                {
                    var name = (string?)element.Attribute(word + "name") ?? string.Empty;
                    if (name.StartsWith("VTOMML_", StringComparison.OrdinalIgnoreCase))
                    {
                        pendingFormulaId = formulaByBookmark.TryGetValue(name, out var formulaId)
                            ? formulaId
                            : null;
                    }
                    continue;
                }
                if (pendingFormulaId is null || element.Name != math + "oMath")
                    continue;

                fingerprints[pendingFormulaId] =
                    WordOmmlConverter.ComputeOmmlFingerprint(
                        element.ToString(SaveOptions.DisableFormatting));
                pendingFormulaId = null;
                if (fingerprints.Count == formulaByBookmark.Count)
                    break;
            }

            if (fingerprints.Count != formulaByBookmark.Count)
            {
                var missing = formulaByBookmark.Values
                    .Where(formulaId => !fingerprints.ContainsKey(formulaId))
                    .Take(5)
                    .ToArray();
                throw new InvalidDataException(
                    $"Word document XML exposed {fingerprints.Count}/{formulaByBookmark.Count} converted OMML formulas while finalizing fingerprints. Missing: {string.Join(", ", missing)}");
            }

            var updated = 0;
            foreach (var pair in fingerprints)
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, pair.Key)
                    ?? throw new InvalidDataException(
                        $"Converted OMML metadata '{pair.Key}' disappeared before fingerprint finalization.");
                metadata.NativeOmmlFingerprint = pair.Value;
                WordOmmlFormulaStore.Save(document, metadata);
                updated++;
            }
            return updated;
        }
        finally { Release(content); }
    }

    internal static string ReadCompleteEquationWordOpenXml(
        Document document,
        Range equationRange,
        string formulaId)
    {
        Range? content = null;
        Range? probe = null;
        Bookmarks? bookmarks = null;
        Bookmark? boundaryBookmark = null;
        Range? boundaryRange = null;
        try
        {
            content = document.Content;
            var probeEnd = equationRange.End;
            if (Guid.TryParse(formulaId, out var parsed))
            {
                bookmarks = document.Bookmarks;
                var boundaryName = "VTBL_" + parsed.ToString("N");
                if (bookmarks.Exists(boundaryName))
                {
                    boundaryBookmark = bookmarks[boundaryName];
                    boundaryRange = boundaryBookmark.Range;
                    // VTBL owns the ordinary-text typing anchor immediately after
                    // an inline formula. The anchor is a hard serialization bound,
                    // never part of the native equation.
                    probeEnd = Math.Max(probeEnd, boundaryRange.Start);
                }
            }

            // A field boundary can span dozens of Word structure characters.
            // Include the complete field, not merely its first marker; otherwise
            // Word may serialize only the leading fragment of a compound OMath.
            object start = equationRange.Start;
            object end = Math.Min(content.End, Math.Max(probeEnd, equationRange.End));
            probe = document.Range(ref start, ref end);
            var xml = probe.WordOpenXML;
            WordOmmlConverter.ExtractSingleOMath(xml);
            return xml;
        }
        finally
        {
            Release(boundaryRange);
            Release(boundaryBookmark);
            Release(bookmarks);
            Release(probe);
            Release(content);
        }
    }

    private static string ReadDisplayMode(Range equationRange)
    {
        OMaths? maths = null;
        OMath? selected = null;
        try
        {
            maths = equationRange.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? candidate = null;
                Range? candidateRange = null;
                try
                {
                    candidate = maths[index];
                    candidateRange = candidate.Range;
                    if (selected is null
                        || candidateRange.Start == equationRange.Start
                            && candidateRange.End == equationRange.End)
                    {
                        Release(selected);
                        selected = candidate;
                        candidate = null;
                        if (candidateRange.Start == equationRange.Start
                            && candidateRange.End == equationRange.End)
                            break;
                    }
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }
            return selected?.Type == WdOMathType.wdOMathDisplay
                ? "block"
                : "inline";
        }
        catch { return "inline"; }
        finally
        {
            Release(selected);
            Release(maths);
        }
    }

    private static double ReadFontSize(Range equationRange)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = equationRange.Font;
            var size = font.Size;
            return size > 0 && !float.IsNaN(size) && !float.IsInfinity(size)
                ? FormulaFontSize.Normalize(size)
                : FormulaFontSize.DefaultPt;
        }
        catch { return FormulaFontSize.DefaultPt; }
        finally { Release(font); }
    }

    private static void SanitizeMetadataBoundaryArtifacts(FormulaMetadata metadata)
    {
        foreach (var line in metadata.Lines)
            line.Latex = SanitizeFormulaBoundaryArtifacts(line.Latex);
        metadata.Latex = metadata.Lines.Count > 0
            ? string.Join("\n", metadata.Lines.Select(line => line.Latex))
            : SanitizeFormulaBoundaryArtifacts(metadata.Latex);
    }

    private static string SanitizeFormulaBoundaryArtifacts(string? latex)
    {
        if (string.IsNullOrEmpty(latex)) return string.Empty;
        static bool IsBoundaryArtifact(char character) =>
            character is '\u200B' or '\u200C' or '\u2060' or '\uFEFF';

        var value = latex!;
        for (var index = 0; index < value.Length; index++)
        {
            if (!IsBoundaryArtifact(value[index])) continue;
            var runEnd = index + 1;
            while (runEnd < value.Length && IsBoundaryArtifact(value[runEnd]))
                runEnd++;
            var left = value.Substring(0, index).Trim();
            var right = value.Substring(runEnd).Trim();
            if (left.Length > 0 && string.Equals(left, right, StringComparison.Ordinal))
                return left;
            index = runEnd - 1;
        }

        return new string(value.Where(character => !IsBoundaryArtifact(character)).ToArray());
    }

    private static FormulaMetadata Clone(FormulaMetadata metadata)
    {
        var clone = FormulaMetadataCodec.Decode(FormulaMetadataCodec.Encode(metadata));
        return clone
            ?? throw new InvalidDataException("Unable to clone VisualTeX formula metadata.");
    }

    private static void Release(object? value)
    {
        if (value is null || !System.Runtime.InteropServices.Marshal.IsComObject(value)) return;
        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(value); } catch { }
    }
}
