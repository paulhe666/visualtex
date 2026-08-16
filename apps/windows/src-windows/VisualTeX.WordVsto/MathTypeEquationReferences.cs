using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Native MathType equation-reference compatibility for Word.
///
/// MathType owns numbered equations with an outer MACROBUTTON MTPlaceRef field.
/// Its hidden first SEQ MTEqn \h field increments the equation counter; only the
/// visible portion after that hidden field is bookmarked.  Native MathType
/// references are a GOTOBUTTON containing a nested REF ... \! field.  Keeping
/// exactly that structure lets VisualTeX reference numbered equations created by
/// either MathType itself or VisualTeX without inventing a second numbering
/// system.
/// </summary>
internal static class MathTypeEquationReferences
{
    private const string PlaceRefMarker = "MACROBUTTON MTPlaceRef";
    private const string EquationBookmarkPrefix = "ZEqnNum";

    internal static IReadOnlyList<EquationReferenceTarget> GetTargets(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));

        var targets = new List<EquationReferenceTarget>();
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (!IsMathTypePlaceRefCode(code.Text)) continue;

                var position = Math.Max(document.Content.Start, code.Start - 1);
                Range? numberRange = null;
                try
                {
                    if (!TryGetVisibleNumberRange(document, field, out numberRange)
                        || numberRange is null)
                        continue;
                    if (!TryFindMathTypeEquationLatex(field, out var latexPreview))
                        continue;

                    var numberText = ReadVisibleNumberText(field);
                    if (string.IsNullOrWhiteSpace(numberText))
                        numberText = (numberRange.Text ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(numberText)) continue;

                    targets.Add(new EquationReferenceTarget(
                        $"mathtype:{position}",
                        -1,
                        numberText,
                        latexPreview,
                        position,
                        EquationReferenceSource.MathType));
                }
                catch
                {
                    // A malformed legacy MTPlaceRef must not hide other valid
                    // numbered MathType equations from the reference picker.
                }
                finally { Release(numberRange); }
            }
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }

        return targets
            .OrderBy(target => target.Position)
            .ToArray();
    }

    internal static void InsertReference(
        Document document,
        Selection selection,
        EquationReferenceTarget target)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (selection is null) throw new ArgumentNullException(nameof(selection));
        if (target is null) throw new ArgumentNullException(nameof(target));
        if (target.Source != EquationReferenceSource.MathType)
            throw new InvalidOperationException("The selected equation is not a MathType reference target.");
        if (document.ReadOnly)
            throw new UnauthorizedAccessException("当前 Word 文档为只读状态。");

        Field? placeRef = null;
        Range? numberRange = null;
        Range? insertion = null;
        Field? goToField = null;
        Range? goToCode = null;
        Range? nestedInsertion = null;
        Field? refField = null;
        Range? goToResult = null;
        try
        {
            placeRef = ResolvePlaceRef(document, target)
                ?? throw new InvalidOperationException(
                    "找不到目标 MathType 公式编号。文档内容可能已在引用窗口打开后发生变化。");
            if (!TryGetVisibleNumberRange(document, placeRef, out numberRange)
                || numberRange is null)
                throw new InvalidDataException("MathType 公式编号缺少可引用的可见编号范围。");

            var bookmarkName = EnsureNativeMathTypeNumberBookmark(document, numberRange);

            insertion = selection.Range;
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            goToField = document.Fields.Add(
                insertion,
                WdFieldType.wdFieldGoToButton,
                bookmarkName + " ",
                true);

            // MathType inserts the REF field *inside the GOTOBUTTON field code*.
            // Word then renders the nested REF result as the visible reference,
            // while double-clicking the outer field navigates back to the number.
            goToCode = goToField.Code;
            nestedInsertion = document.Range(goToCode.End, goToCode.End);
            refField = document.Fields.Add(
                nestedInsertion,
                WdFieldType.wdFieldRef,
                bookmarkName + " \\* Charformat \\!",
                true);
            try { refField.Update(); } catch { }
            try { refField.ShowCodes = false; } catch { }
            try { goToField.ShowCodes = false; } catch { }

            Release(goToResult);
            goToResult = goToField.Result;
            var after = Math.Min(document.Content.End, goToResult.End + 1);
            selection.SetRange(after, after);
        }
        finally
        {
            Release(goToResult);
            Release(refField);
            Release(nestedInsertion);
            Release(goToCode);
            Release(goToField);
            Release(insertion);
            Release(numberRange);
            Release(placeRef);
        }
    }

    private static Field? ResolvePlaceRef(
        Document document,
        EquationReferenceTarget target)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Field? best = null;
        var bestDistance = int.MaxValue;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (!IsMathTypePlaceRefCode(code.Text)) continue;
                var position = Math.Max(document.Content.Start, code.Start - 1);
                var distance = Math.Abs(position - target.Position);
                if (distance > 8 || distance >= bestDistance) continue;

                Range? numberRange = null;
                try
                {
                    if (!TryGetVisibleNumberRange(document, field, out numberRange)
                        || numberRange is null)
                        continue;
                    var numberText = ReadVisibleNumberText(field);
                    if (!string.Equals(
                            numberText.Trim(),
                            target.NumberText.Trim(),
                            StringComparison.Ordinal))
                        continue;
                }
                finally { Release(numberRange); }

                Release(best);
                best = field;
                field = null;
                bestDistance = distance;
                if (distance == 0) break;
            }
            return best;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static bool TryGetVisibleNumberRange(
        Document document,
        Field placeRef,
        out Range? numberRange)
    {
        numberRange = null;
        Range? outerCode = null;
        Fields? nestedFields = null;
        Field? nested = null;
        Range? nestedCode = null;
        Range? hiddenResult = null;
        try
        {
            outerCode = placeRef.Code;
            if (!IsMathTypePlaceRefCode(outerCode.Text)) return false;
            nestedFields = outerCode.Fields;
            for (var index = 1; index <= nestedFields.Count; index++)
            {
                Release(nestedCode);
                nestedCode = null;
                Release(nested);
                nested = nestedFields[index];
                nestedCode = nested.Code;
                if (!IsHiddenMathTypeEquationIncrement(nestedCode.Text)) continue;
                hiddenResult = nested.Result;
                var start = hiddenResult.End;
                var end = outerCode.End;
                if (end <= start) return false;
                numberRange = document.Range(start, end);
                return true;
            }
            return false;
        }
        finally
        {
            Release(hiddenResult);
            Release(nestedCode);
            Release(nested);
            Release(nestedFields);
            Release(outerCode);
        }
    }

    private static string ReadVisibleNumberText(Field placeRef)
    {
        Range? outerCode = null;
        Fields? nestedFields = null;
        Field? nested = null;
        Range? nestedCode = null;
        Range? nestedResult = null;
        try
        {
            outerCode = placeRef.Code;
            var stream = outerCode.Text ?? string.Empty;
            nestedFields = outerCode.Fields;
            if (nestedFields.Count == 0) return string.Empty;

            var hiddenControlStart = -1;
            var hiddenControlEnd = -1;
            var nestedOrdinal = 0;
            for (var index = 0; index < stream.Length; index++)
            {
                if (stream[index] != '\u0013') continue;
                nestedOrdinal++;
                var end = stream.IndexOf('\u0015', index + 1);
                if (end < 0) return string.Empty;
                if (nestedOrdinal <= nestedFields.Count)
                {
                    Release(nestedCode);
                    nestedCode = null;
                    Release(nested);
                    nested = nestedFields[nestedOrdinal];
                    nestedCode = nested.Code;
                    if (IsHiddenMathTypeEquationIncrement(nestedCode.Text))
                    {
                        hiddenControlStart = index;
                        hiddenControlEnd = end;
                        break;
                    }
                }
                index = end;
            }
            if (hiddenControlStart < 0 || hiddenControlEnd < 0) return string.Empty;

            var visible = new System.Text.StringBuilder();
            var fieldOrdinal = nestedOrdinal;
            for (var index = hiddenControlEnd + 1; index < stream.Length;)
            {
                if (stream[index] != '\u0013')
                {
                    if (stream[index] is not '\u0014' and not '\u0015')
                        visible.Append(stream[index]);
                    index++;
                    continue;
                }

                var end = stream.IndexOf('\u0015', index + 1);
                if (end < 0) break;
                fieldOrdinal++;
                if (fieldOrdinal <= nestedFields.Count)
                {
                    Release(nestedResult);
                    nestedResult = null;
                    Release(nested);
                    nested = nestedFields[fieldOrdinal];
                    nestedResult = nested.Result;
                    visible.Append(nestedResult.Text ?? string.Empty);
                }
                index = end + 1;
            }
            return visible.ToString().Trim();
        }
        catch { return string.Empty; }
        finally
        {
            Release(nestedResult);
            Release(nestedCode);
            Release(nested);
            Release(nestedFields);
            Release(outerCode);
        }
    }

    private static bool TryFindMathTypeEquationLatex(Field placeRef, out string latex)
    {
        latex = "MathType 公式";
        Range? code = null;
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        try
        {
            code = placeRef.Code;
            var document = code.Document;
            var position = Math.Max(document.Content.Start, code.Start - 1);
            probe = document.Range(position, position);
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            shapes = paragraphRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                try
                {
                    var mathMl = MathTypeOleStorage.ReadMathMl(shape);
                    var converted = MathMlToLatexConverter.Convert(mathMl).Trim();
                    if (!string.IsNullOrWhiteSpace(converted)) latex = converted;
                }
                catch { }
                return true;
            }
            return false;
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
            Release(code);
        }
    }

    private static string EnsureNativeMathTypeNumberBookmark(
        Document document,
        Range numberRange)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmarkRange);
                bookmarkRange = null;
                Release(bookmark);
                bookmark = bookmarks[index];
                if (!bookmark.Name.StartsWith(EquationBookmarkPrefix, StringComparison.OrdinalIgnoreCase))
                    continue;
                bookmarkRange = bookmark.Range;
                if (bookmarkRange.Start == numberRange.Start
                    && bookmarkRange.End == numberRange.End)
                    return bookmark.Name;
            }

            var seed = 100000 + Math.Abs(numberRange.Start % 900000);
            for (var offset = 0; offset < 900000; offset++)
            {
                var value = 100000 + ((seed - 100000 + offset) % 900000);
                var name = EquationBookmarkPrefix
                    + value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (bookmarks.Exists(name)) continue;
                Release(bookmark);
                bookmark = bookmarks.Add(name, numberRange);
                return name;
            }
            throw new InvalidOperationException("无法为 MathType 公式编号创建唯一引用书签。");
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static bool IsMathTypePlaceRefCode(string? code) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf(PlaceRefMarker, StringComparison.OrdinalIgnoreCase) >= 0;

    private static bool IsHiddenMathTypeEquationIncrement(string? code)
    {
        if (string.IsNullOrWhiteSpace(code)) return false;
        var normalized = " " + code!.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ') + " ";
        return normalized.IndexOf(" SEQ MTEqn ", StringComparison.OrdinalIgnoreCase) >= 0
            && normalized.IndexOf("\\h", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.FinalReleaseComObject(value); }
        catch { }
    }
}
