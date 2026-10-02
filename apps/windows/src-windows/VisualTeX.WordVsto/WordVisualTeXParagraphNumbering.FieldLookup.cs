using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static partial class WordVisualTeXParagraphNumbering
{
    private const string SequenceName = "VisualTeXEquation";
    private const string ChapterSequenceName = "VisualTeXChapter";
    private const string SectionSequenceName = "VisualTeXSection";

    // Operation-owned proof of one canonical body number. General discovery does
    // not retain these mutation handles. Numeric positions are never deletion owners.
    internal sealed class BoundBodyNumber : IDisposable
    {
        private Range? _host;
        private Range? _field;
        private Range? _separator;
        private Range? _code;
        private readonly List<(int Order, Field Field)> _headings = new();

        internal BoundBodyNumber(Range host) { _host = host.Duplicate; }
        internal Range Code => _code!;
        internal void AddHeading(int order, Field field) => _headings.Add((order, field));
        internal void SetOuter(Range field, Range separator, Range code)
        { _field = field; _separator = separator; _code = code; }

        internal Range Remove()
        {
            if (_field is null || _separator is null || _host is null)
                throw new InvalidOperationException("The prepared VisualTeX body number was already consumed.");
            _field.Delete();
            _separator.Delete();
            foreach (var heading in _headings.OrderByDescending(item => item.Order))
                heading.Field.Delete();
            var host = _host;
            _host = null;
            try { Dispose(); return host; }
            catch { Release(host); throw; }
        }

        public void Dispose()
        {
            var host = _host; _host = null;
            var field = _field; _field = null;
            var separator = _separator; _separator = null;
            var code = _code; _code = null;
            try { Release(code); }
            finally
            {
                try { Release(separator); }
                finally
                {
                    try { Release(field); }
                    finally
                    {
                        try { Release(host); }
                        finally
                        {
                            while (_headings.Count > 0)
                            {
                                var index = _headings.Count - 1;
                                var heading = _headings[index].Field;
                                _headings.RemoveAt(index);
                                Release(heading);
                            }
                        }
                    }
                }
            }
        }
    }

    private static bool IsHeadingStateRestartCode(
        string? instruction)
    {
        if (string.IsNullOrWhiteSpace(
                instruction))
            return false;
        var normalized =
            instruction!
                .TrimStart();
        var isOwnedSequence =
            normalized.StartsWith(
                "SEQ "
                + ChapterSequenceName
                + " ",
                StringComparison.OrdinalIgnoreCase)
            || normalized.StartsWith(
                "SEQ "
                + SectionSequenceName
                + " ",
                StringComparison.OrdinalIgnoreCase);
        return isOwnedSequence
            && normalized.IndexOf(
                   "\\r",
                   StringComparison.OrdinalIgnoreCase) >= 0
            && normalized.IndexOf(
                   "\\h",
                   StringComparison.OrdinalIgnoreCase) >= 0;
    }


    internal static Range? FindVisibleLabelRange(
        Document document,
        string formulaId)
    {
        Range? hostRange = null;
        Field? placeRef = null;
        Range? visible = null;
        try
        {
            hostRange =
                ResolveVisualTeXHostRangeByFormulaId(
                    document,
                    formulaId);
            if (hostRange is null
                || !TryFindPlaceRef(
                    document,
                    hostRange,
                    out placeRef)
                || placeRef is null
                || !TryGetVisibleNumberRange(
                    document,
                    placeRef,
                    includeParentheses: true,
                    out visible)
                || visible is null)
                return null;

            var result =
                visible.Duplicate;
            return result;
        }
        finally
        {
            Release(visible);
            Release(placeRef);
            Release(hostRange);
        }
    }

    internal static Range? FindOwnerParagraphRange(
        Document document,
        string formulaId)
    {
        Range? hostRange = null;
        Field? placeRef = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? owner = null;
        try
        {
            hostRange =
                ResolveVisualTeXHostRangeByFormulaId(
                    document,
                    formulaId);
            if (hostRange is null
                || !TryFindPlaceRef(
                    document,
                    hostRange,
                    out placeRef)
                || placeRef is null)
                return null;

            paragraphs =
                hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                return null;
            paragraph =
                paragraphs[1];
            owner =
                paragraph.Range.Duplicate;
            var result = owner;
            owner = null;
            return result;
        }
        finally
        {
            Release(owner);
            Release(paragraph);
            Release(paragraphs);
            Release(placeRef);
            Release(hostRange);
        }
    }

    private static bool TryFindPlaceRef(
        Document document,
        Range hostRange,
        out Field? placeRef,
        BoundBodyNumber? retainedBody = null)
    {
        placeRef = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        Range? separator = null;
        try
        {
            paragraphs =
                hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                return false;
            paragraph =
                paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            fields =
                paragraphRange.Fields;

            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(result);
                result = null;
                Release(code);
                code = null;
                Release(field);
                field =
                    fields[index];
                if (retainedBody is not null)
                {
                    code = field.Code.Duplicate;
                    if (code.Start - 1 < hostRange.Start
                        && IsHeadingStateRestartCode(code.Text))
                    {
                        retainedBody.AddHeading(code.Start, field);
                        field = null;
                        continue;
                    }
                }
                if (placeRef is not null
                    || field.Type != WdFieldType.wdFieldMacroButton)
                    continue;
                code ??= field.Code.Duplicate;
                if (!IsPlaceRefCode(code.Text))
                    continue;

                var fieldStart =
                    code.Start - 1;
                if (fieldStart < hostRange.End)
                    return false;
                separator =
                    document.Range(
                        hostRange.End,
                        fieldStart);
                if (!string.Equals(
                        separator.Text,
                        "\t",
                        StringComparison.Ordinal))
                    return false;

                result =
                    field.Result.Duplicate;
                var fieldEnd =
                    ResolveOuterFieldEndExclusive(
                        document,
                        code,
                        result);
                var editableEnd =
                    Math.Max(
                        paragraphRange.Start,
                        paragraphRange.End - 1);
                if (fieldEnd > editableEnd)
                    return false;

                Range? suffix = null;
                try
                {
                    suffix =
                        document.Range(
                            fieldEnd,
                            editableEnd);
                    if (!ContainsOnlyScaffoldWhitespace(
                            suffix.Text))
                        return false;
                }
                finally
                {
                    Release(suffix);
                }

                if (retainedBody is not null)
                {
                    var fullField = document.Range(fieldStart, fieldEnd);
                    retainedBody.SetOuter(fullField, separator, code);
                    separator = null;
                    code = null;
                }
                var found = field;
                field = null;
                placeRef = found;
                if (retainedBody is null) return true;
            }

            return placeRef is not null;
        }
        finally
        {
            Release(separator);
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool TryGetVisibleNumberRange(
        Document document,
        Field placeRef,
        bool includeParentheses,
        out Range? numberRange,
        Range? preparedOuterCode = null)
    {
        numberRange = null;
        Range? outerCode = null;
        Fields? nested = null;
        Field? field = null;
        Range? fieldCode = null;
        Range? hiddenResult = null;
        try
        {
            outerCode = preparedOuterCode ?? placeRef.Code.Duplicate;
            if (preparedOuterCode is null && !IsPlaceRefCode(outerCode.Text))
                return false;
            nested =
                outerCode.Fields;
            for (var index = 1;
                 index <= nested.Count;
                 index++)
            {
                Release(fieldCode);
                fieldCode = null;
                Release(field);
                field =
                    nested[index];
                fieldCode =
                    field.Code.Duplicate;
                if (!IsHiddenIncrement(
                        document,
                        fieldCode.Text))
                    continue;

                hiddenResult =
                    field.Result.Duplicate;
                var start =
                    hiddenResult.End;
                var end =
                    outerCode.End;

                var outerText =
                    outerCode.Text
                    ?? string.Empty;
                var trimmedLength =
                    outerText
                        .TrimEnd()
                        .Length;
                end -=
                    outerText.Length
                    - trimmedLength;
                if (!includeParentheses)
                {
                    start++;
                    end--;
                }
                if (end <= start)
                    return false;

                numberRange =
                    document.Range(
                        start,
                        end);
                return true;
            }

            return false;
        }
        finally
        {
            Release(hiddenResult);
            Release(fieldCode);
            Release(field);
            Release(nested);
            if (preparedOuterCode is null) Release(outerCode);
        }
    }

    private static bool IsHiddenIncrement(
        Document document,
        string? instruction)
    {
        if (string.IsNullOrWhiteSpace(
                instruction))
            return false;
        var normalized =
            " "
            + instruction!
                .Replace('\t', ' ')
                .Replace('\r', ' ')
                .Replace('\n', ' ')
            + " ";
        return normalized.IndexOf(
                   " SEQ "
                   + SequenceName
                   + " ",
                   StringComparison.OrdinalIgnoreCase) >= 0
            && normalized.IndexOf(
                   "\\h",
                   StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static int ResolveOuterFieldEndExclusive(
        Document document,
        Range code,
        Range result)
    {
        var codeBoundary =
            ReadCharacter(
                document,
                code.End);
        if (codeBoundary ==
            '\u0015')
            return code.End + 1;
        if (codeBoundary ==
                '\u0014'
            && ReadCharacter(
                    document,
                    result.End)
                == '\u0015')
            return result.End + 1;
        throw new InvalidDataException(
            $"VisualTeXPlaceRef has an invalid outer field boundary at code={code.End}, result={result.End}.");
    }

    private static char ReadCharacter(
        Document document,
        int position)
    {
        Range? probe = null;
        try
        {
            if (position <
                    document.Content.Start
                || position >=
                    document.Content.End)
                return '\0';
            probe =
                document.Range(
                    position,
                    position + 1);
            var text =
                probe.Text
                ?? string.Empty;
            return text.Length == 1
                ? text[0]
                : '\0';
        }
        finally
        {
            Release(probe);
        }
    }

    private static bool ContainsOnlyScaffoldWhitespace(
        string? text)
    {
        if (string.IsNullOrEmpty(
                text))
            return true;
        foreach (var character in text!)
        {
            if (character is
                    '\t'
                    or '\r'
                    or '\n'
                    or '\v'
                    or '\u0013'
                    or '\u0014'
                    or '\u0015'
                || char.IsWhiteSpace(
                    character))
                continue;
            return false;
        }
        return true;
    }

    private static Range? ResolveVisualTeXHostRangeByFormulaId(
        Document document,
        string formulaId)
    {
        if (!Guid.TryParse(
                formulaId,
                out var parsed))
            return null;

        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? owner = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            bookmarks =
                document.Bookmarks;
            var identityName =
                WordFormulaMetadataReader
                    .IdentityBookmarkName(
                        parsed.ToString("D"));
            if (!bookmarks.Exists(
                    identityName))
                return null;

            bookmark =
                bookmarks[identityName];
            owner =
                bookmark.Range.Duplicate;
            shapes =
                owner.InlineShapes;
            if (shapes.Count != 1)
                return null;
            shape =
                shapes[1];
            if (!WordFormulaMetadataReader.IsNativeOle(
                    shape))
                return null;
            shapeRange =
                shape.Range.Duplicate;
            if (shapeRange.StoryType !=
                    owner.StoryType
                || shapeRange.Start !=
                    owner.Start
                || shapeRange.End !=
                    owner.End)
                return null;

            var result =
                shapeRange;
            shapeRange = null;
            return result;
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(owner);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void Release(
        object? value)
    {
        if (value is null
            || !Marshal.IsComObject(
                value))
            return;
        try
        {
            Marshal.ReleaseComObject(
                value);
        }
        catch { }
    }
}
