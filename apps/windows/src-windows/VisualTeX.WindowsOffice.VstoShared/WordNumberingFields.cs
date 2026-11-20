using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordEquationNumbering;
using static VisualTeX.WordVsto.WordEquationSequenceFields;

namespace VisualTeX.WordVsto;

/// <summary>
/// Field and bookmark primitives of equation numbering: sequence fields, bookmark names and formula ids, deleting
/// bookmarked ranges.
/// </summary>
internal static class WordNumberingFields
{
    private static List<int> GetNativeEquationFieldPositions(
        Document document,
        string nativeSequenceName)
    {
        var positions = new List<int>();
        Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName)) continue;
                    result = field.Result;
                    positions.Add(result.Start);
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }
        positions.Sort();
        return positions;
    }

    private static bool TryGetNativeCaptionInfo(
        Document document,
        string formulaId,
        string nativeSequenceName,
        out int position,
        out string numberText)
    {
        position = -1;
        numberText = string.Empty;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Field? field = null;
        Range? result = null;
        try
        {
            bookmarks = document.Bookmarks;
            var bookmarkName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(bookmarkName)) return false;
            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range;
            field = FindNativeEquationFieldAtRange(document, range, nativeSequenceName);
            if (field is null) return false;
            field.Update();
            result = field.Result;
            position = result.Start;
            numberText = (range.Text ?? string.Empty).Trim();
            return !string.IsNullOrWhiteSpace(numberText);
        }
        finally
        {
            Release(result);
            Release(field);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static Field? FindNativeEquationFieldInRange(
        Range scopeRange,
        string nativeSequenceName)
    {
        Fields? fields = null;
        try
        {
            fields = scopeRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName))
                        continue;
                    var found = field;
                    field = null;
                    return found;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            return null;
        }
        finally { Release(fields); }
    }

    internal static Field? FindNativeEquationFieldAtRange(
        Document document,
        Range targetRange,
        string nativeSequenceName)
    {
        // Word can invalidate a live Range RCW while the document-wide Fields
        // collection is materialized. Freeze the coordinates before enumerating
        // fields so later comparisons never re-enter a deleted COM proxy.
        var targetStart = targetRange.Start;
        var targetEnd = targetRange.End;
        Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName)) continue;
                    result = field.Result;
                    var overlaps = result.Start < targetEnd
                        && result.End > targetStart;
                    var sameCollapsedPosition = result.Start == targetStart
                        && result.End == targetEnd;
                    if (!overlaps && !sameCollapsedPosition) continue;
                    var found = field;
                    field = null;
                    return found;
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
            return null;
        }
        finally { Release(fields); }
    }

    internal static string GetNativeEquationSequenceName(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));

        // Every VisualTeX-owned equation host must participate in one locale-neutral
        // Word sequence. Native OMML #(SEQ), table-free VisualTeX OLE and converted
        // MathType formulas can coexist in the same document; using Word's localized
        // built-in “Equation” caption name for only the OLE side would split F9 into
        // independent sequences on non-English Word installations. Older localized
        // caption fields are migration input: EnsureNativeCaption removes/recreates
        // their VisualTeX-owned VTEqCap/VTEqNum scaffold with this fixed sequence.
        return LegacyEquationSequenceName;
    }

    internal static string EquationBookmarkName(string formulaId) =>
        BookmarkName(EquationBookmarkPrefix, formulaId);

    internal static string NativeCaptionBookmarkName(string formulaId) =>
        BookmarkName(NativeCaptionBookmarkPrefix, formulaId);

    internal static string NativeNumberBookmarkName(string formulaId) =>
        BookmarkName(NativeNumberBookmarkPrefix, formulaId);

    internal static string NativeCaptionFlowTailBookmarkName(string formulaId) =>
        BookmarkName(NativeCaptionFlowTailBookmarkPrefix, formulaId);

    private static string BookmarkName(string prefix, string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var value))
            throw new InvalidOperationException("VisualTeX formulaId must be a UUID.");
        return $"{prefix}{value:N}";
    }

    internal static bool TryFormulaIdFromEquationBookmark(
        string? bookmarkName,
        out string formulaId) =>
        TryFormulaIdFromBookmark(bookmarkName, EquationBookmarkPrefix, out formulaId);

    internal static bool TryFormulaIdFromBookmark(
        string? bookmarkName,
        string prefix,
        out string formulaId)
    {
        formulaId = string.Empty;
        if (string.IsNullOrWhiteSpace(bookmarkName)) return false;
        var name = bookmarkName!;
        if (!name.StartsWith(prefix, StringComparison.Ordinal)
            || !Guid.TryParseExact(name.Substring(prefix.Length), "N", out var value))
            return false;
        formulaId = value.ToString();
        return true;
    }

    internal static bool IsVisualTeXSequenceFieldCode(string? code) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf(
            $"SEQ {LegacyEquationSequenceName}",
            StringComparison.OrdinalIgnoreCase) >= 0;

    internal static bool IsNativeEquationSequenceFieldCode(
        string? code,
        string nativeSequenceName)
    {
        if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(nativeSequenceName))
            return false;
        return IsVisualTeXSequenceFieldCode(code)
            || code!.IndexOf(
                   $"SEQ {nativeSequenceName}",
                   StringComparison.OrdinalIgnoreCase) >= 0
            || code.IndexOf(
                   $"SEQ \"{nativeSequenceName}\"",
                   StringComparison.OrdinalIgnoreCase) >= 0;
    }

    internal static bool TryResolveVisualTeXReferenceBookmark(
        Document document,
        string? fieldCode,
        out string bookmarkName)
    {
        bookmarkName = string.Empty;
        if (string.IsNullOrWhiteSpace(fieldCode)) return false;
        var match = Regex.Match(
            fieldCode!,
            @"^\s*REF\s+(?:""(?<quoted>[^""]+)""|(?<plain>[^\s\\]+))",
            RegexOptions.IgnoreCase);
        if (!match.Success) return false;
        var targetName = match.Groups["quoted"].Success
            ? match.Groups["quoted"].Value
            : match.Groups["plain"].Value;
        if (targetName.StartsWith(NativeNumberBookmarkPrefix, StringComparison.Ordinal))
        {
            bookmarkName = targetName;
            return true;
        }

        Bookmarks? bookmarks = null;
        Bookmark? targetBookmark = null;
        Range? targetRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(targetName)) return false;
            targetBookmark = bookmarks[targetName];
            targetRange = targetBookmark.Range;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? candidateBookmark = null;
                Range? candidateRange = null;
                try
                {
                    candidateBookmark = bookmarks[index];
                    if (!candidateBookmark.Name.StartsWith(
                            NativeNumberBookmarkPrefix,
                            StringComparison.Ordinal))
                        continue;
                    candidateRange = candidateBookmark.Range;
                    var overlaps = candidateRange.Start < targetRange.End
                        && candidateRange.End > targetRange.Start;
                    var sameCollapsedPosition = candidateRange.Start == targetRange.Start
                        && candidateRange.End == targetRange.End;
                    if (!overlaps && !sameCollapsedPosition) continue;
                    bookmarkName = candidateBookmark.Name;
                    return true;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidateBookmark);
                }
            }
            return false;
        }
        finally
        {
            Release(targetRange);
            Release(targetBookmark);
            Release(bookmarks);
        }
    }

    internal static bool IsReferenceFieldCode(string? code)
    {
        if (string.IsNullOrWhiteSpace(code)) return false;
        var trimmed = code!.TrimStart();
        return trimmed.StartsWith("REF ", StringComparison.OrdinalIgnoreCase);
    }

    internal static bool IsReferenceToBookmark(string? code, string bookmarkName) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf(
            $"REF {bookmarkName}",
            StringComparison.OrdinalIgnoreCase) >= 0;

    private static void DeleteBookmarkedRange(Document document, string name)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            range = bookmark.Range;
            range.Delete();
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static void DeleteBookmarkOnly(Document document, string name)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            bookmark.Delete();
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
    }
    private static void RewriteSequenceFieldCodeOutsideMath(
        Field field,
        Range code,
        string replacementCode)
    {
        EnsureSequenceFieldCodeCanBeRewritten(field);
        code.Text = replacementCode;
    }

    internal static string ExtractNativeHashSequenceResultFromBookmarkXml(
        string bookmarkSegment)
    {
        if (string.IsNullOrEmpty(bookmarkSegment)) return string.Empty;
        var options = RegexOptions.IgnoreCase
            | RegexOptions.CultureInvariant
            | RegexOptions.Singleline;
        var begin = Regex.Match(
            bookmarkSegment,
            @"<w:fldChar\b[^>]*\bw:fldCharType=""begin""[^>]*/?>",
            options);
        var separate = Regex.Match(
            bookmarkSegment,
            @"<w:fldChar\b[^>]*\bw:fldCharType=""separate""[^>]*/?>",
            options);
        var end = Regex.Match(
            bookmarkSegment,
            @"<w:fldChar\b[^>]*\bw:fldCharType=""end""[^>]*/?>",
            options);
        if (!begin.Success || !separate.Success || !end.Success
            || begin.Index >= separate.Index
            || separate.Index >= end.Index)
            return string.Empty;

        static int LastMathRunStart(string xml, int beforeIndex)
        {
            var bare = xml.LastIndexOf(
                "<m:r>",
                beforeIndex,
                StringComparison.OrdinalIgnoreCase);
            var attributed = xml.LastIndexOf(
                "<m:r ",
                beforeIndex,
                StringComparison.OrdinalIgnoreCase);
            return Math.Max(bare, attributed);
        }

        var beginRunStart = LastMathRunStart(bookmarkSegment, begin.Index);
        var separateRunEnd = bookmarkSegment.IndexOf(
            "</m:r>",
            separate.Index,
            StringComparison.OrdinalIgnoreCase);
        var endRunStart = LastMathRunStart(bookmarkSegment, end.Index);
        if (beginRunStart < 0 || separateRunEnd < separate.Index
            || endRunStart <= separateRunEnd)
            return string.Empty;
        separateRunEnd += "</m:r>".Length;

        // VTEqNum wraps the complete displayed number. Static chapter text lives
        // before fldChar(begin), while the live SEQ result lives between
        // fldChar(separate) and fldChar(end). The instruction run between begin and
        // separate must never be interpreted as rendered number text.
        var visibleXml = bookmarkSegment.Substring(0, beginRunStart)
            + bookmarkSegment.Substring(
                separateRunEnd,
                endRunStart - separateRunEnd);
        var text = string.Concat(
            Regex.Matches(
                    visibleXml,
                    @"<(?:w|m):t(?:\s[^>]*)?>(?<text>.*?)</(?:w|m):t>",
                    options)
                .Cast<Match>()
                .Select(match => System.Net.WebUtility.HtmlDecode(
                    match.Groups["text"].Value)));
        return NormalizeNativeEquationNumberText(text);
    }

    internal static bool HasLegacyNumberedOmmlShapeArtifact(Document document)
    {
        Range? content = null;
        try
        {
            content = document.Content;
            var xml = WordDocumentXml.Read(document) ?? string.Empty;
            return Regex.IsMatch(
                    xml,
                    @"\bVTEqShape_[0-9A-F]{32}\b",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                || xml.IndexOf(
                    NativeDisplayNumberShapeAlternativeTextPrefix,
                    StringComparison.Ordinal) >= 0;
        }
        catch
        {
            return false;
        }
        finally { Release(content); }
    }
}
