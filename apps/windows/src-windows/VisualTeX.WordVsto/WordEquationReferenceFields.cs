using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Creates a real Word equation reference: an outer GOTOBUTTON field containing
/// one nested REF field. The nested REF keeps the displayed number live, while
/// the outer field makes double-click navigation work consistently for VisualTeX,
/// OMML and MathType numbering bookmarks.
/// </summary>
internal static class WordEquationReferenceFields
{
    private const string MathTypeSectionStyleName = "MTEquationSection";

    private sealed class CharacterFormatting
    {
        internal int? Bold { get; set; }
        internal int? Italic { get; set; }
        internal WdUnderline? Underline { get; set; }
        internal WdColor? Color { get; set; }
        internal float? Size { get; set; }
    }

    internal static void InsertNavigableReference(
        Document document,
        Selection selection,
        string bookmarkName,
        string prefix,
        string suffix,
        WdColor? preferredInsertionColor = null)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (selection is null) throw new ArgumentNullException(nameof(selection));
        if (string.IsNullOrWhiteSpace(bookmarkName))
            throw new ArgumentException("Equation reference bookmark is required.", nameof(bookmarkName));
        if (document.ReadOnly)
            throw new UnauthorizedAccessException("当前 Word 文档为只读状态。");

        Bookmarks? bookmarks = null;
        Range? sourceFormattingRange = null;
        Range? insertion = null;
        Field? goToField = null;
        Range? goToCode = null;
        Range? nestedInsertion = null;
        Field? refField = null;
        Fields? finalNestedFields = null;
        Field? finalRefField = null;
        Range? finalRefCode = null;
        Range? finalRefResult = null;
        Range? goToResult = null;
        Range? selectionRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(bookmarkName))
                throw new InvalidDataException(
                    $"公式引用目标书签“{bookmarkName}”已不存在。文档内容可能已发生变化。");

            sourceFormattingRange = selection.Range.Duplicate;
            sourceFormattingRange.Collapse(WdCollapseDirection.wdCollapseStart);
            var formatting = CaptureFormatting(sourceFormattingRange);
            if (preferredInsertionColor.HasValue)
            {
                var requested = preferredInsertionColor.Value;
                formatting.Color = requested == WdColor.wdColorAutomatic
                    || (int)requested >= 0
                        ? requested
                        : WdColor.wdColorAutomatic;
            }

            if (!string.IsNullOrEmpty(prefix))
                selection.TypeText(prefix);

            insertion = selection.Range.Duplicate;
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            goToField = document.Fields.Add(
                insertion,
                WdFieldType.wdFieldGoToButton,
                bookmarkName + " ",
                true);

            // The nested REF is deliberately placed inside GOTOBUTTON.Code. Word
            // renders its result as the visible number and routes a double-click
            // on the enclosing field to the bookmark.
            goToCode = goToField.Code;
            NormalizeInternalStyle(goToCode);
            ApplyFormatting(goToCode, formatting);
            nestedInsertion = document.Range(goToCode.End, goToCode.End);
            refField = document.Fields.Add(
                nestedInsertion,
                WdFieldType.wdFieldRef,
                bookmarkName + " \\* CHARFORMAT \\!",
                true);
            try { refField.ShowCodes = false; } catch { }
            try { goToField.ShowCodes = false; } catch { }

            // Updating the nested field can rematerialize the complete outer field
            // and reapply Word's default red GOTOBUTTON formatting. Reacquire and
            // normalize the final field tree after the update.
            try { refField.Update(); } catch { }
            Release(goToCode);
            goToCode = goToField.Code;
            NormalizeInternalStyle(goToCode);
            ApplyFormatting(goToCode, formatting);

            finalNestedFields = goToCode.Fields;
            if (finalNestedFields.Count != 1)
                throw new InvalidDataException(
                    "公式引用未能保留一个完整的嵌套 REF 字段。");
            finalRefField = finalNestedFields[1];
            finalRefCode = finalRefField.Code;
            NormalizeInternalStyle(finalRefCode);
            ApplyFormatting(finalRefCode, formatting);
            try { finalRefField.Update(); } catch { }

            Release(goToCode);
            goToCode = goToField.Code;
            NormalizeInternalStyle(goToCode);
            ApplyFormatting(goToCode, formatting);

            Release(finalRefResult);
            finalRefResult = null;
            Release(finalRefCode);
            finalRefCode = null;
            Release(finalRefField);
            finalRefField = null;
            Release(finalNestedFields);
            finalNestedFields = null;
            finalNestedFields = goToCode.Fields;
            if (finalNestedFields.Count != 1)
                throw new InvalidDataException(
                    "公式引用在刷新后丢失了嵌套 REF 字段。");
            finalRefField = finalNestedFields[1];
            finalRefResult = finalRefField.Result;
            NormalizeInternalStyle(finalRefResult);
            ApplyFormatting(finalRefResult, formatting);

            goToResult = goToField.Result;
            var after = Math.Max(goToResult.End + 1, goToCode.End + 2);
            after = Math.Max(
                document.Content.Start,
                Math.Min(after, Math.Max(document.Content.Start, document.Content.End - 1)));
            selection.SetRange(after, after);
            selectionRange = selection.Range;
            NormalizeInternalStyle(selectionRange);
            ApplyFormatting(selectionRange, formatting);

            if (!string.IsNullOrEmpty(suffix))
                selection.TypeText(suffix);
        }
        finally
        {
            Release(selectionRange);
            Release(goToResult);
            Release(finalRefResult);
            Release(finalRefCode);
            Release(finalRefField);
            Release(finalNestedFields);
            Release(refField);
            Release(nestedInsertion);
            Release(goToCode);
            Release(goToField);
            Release(insertion);
            Release(sourceFormattingRange);
            Release(bookmarks);
        }
    }

    private static CharacterFormatting CaptureFormatting(Range range)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            var formatting = new CharacterFormatting();
            try
            {
                var value = font.Bold;
                if (value != (int)WdConstants.wdUndefined) formatting.Bold = value;
            }
            catch { }
            try
            {
                var value = font.Italic;
                if (value != (int)WdConstants.wdUndefined) formatting.Italic = value;
            }
            catch { }
            try
            {
                var value = font.Underline;
                if ((int)value != (int)WdConstants.wdUndefined) formatting.Underline = value;
            }
            catch { }
            try
            {
                var value = font.Color;
                if (value == WdColor.wdColorAutomatic || (int)value >= 0)
                    formatting.Color = value;
            }
            catch { }
            try
            {
                var value = font.Size;
                if (value > 0 && value < 1000) formatting.Size = value;
            }
            catch { }
            return formatting;
        }
        finally { Release(font); }
    }

    private static void ApplyFormatting(Range range, CharacterFormatting formatting)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            if (formatting.Bold.HasValue)
            {
                try { font.Bold = formatting.Bold.Value; } catch { }
            }
            if (formatting.Italic.HasValue)
            {
                try { font.Italic = formatting.Italic.Value; } catch { }
            }
            if (formatting.Underline.HasValue)
            {
                try { font.Underline = formatting.Underline.Value; } catch { }
            }
            if (formatting.Color.HasValue)
            {
                try { font.Color = formatting.Color.Value; } catch { }
            }
            if (formatting.Size.HasValue)
            {
                try { font.Size = formatting.Size.Value; } catch { }
            }
            try { font.Hidden = 0; } catch { }
        }
        finally { Release(font); }
    }

    private static void NormalizeInternalStyle(Range range)
    {
        Style? style = null;
        try
        {
            try { style = range.get_Style() as Style; }
            catch { }
            var styleName = string.Empty;
            try { styleName = style?.NameLocal ?? string.Empty; }
            catch { }
            if (!string.Equals(
                    styleName,
                    MathTypeSectionStyleName,
                    StringComparison.OrdinalIgnoreCase))
                return;

            object defaultParagraphFont = WdBuiltinStyle.wdStyleDefaultParagraphFont;
            try { range.set_Style(ref defaultParagraphFont); } catch { }
            try { range.Font.Hidden = 0; } catch { }
        }
        finally { Release(style); }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
