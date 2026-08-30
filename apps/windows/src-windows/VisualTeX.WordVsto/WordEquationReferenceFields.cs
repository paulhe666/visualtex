using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
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
        internal int? Position { get; set; }
        internal string? Name { get; set; }
        internal string? NameAscii { get; set; }
        internal string? NameFarEast { get; set; }
        internal string? NameBi { get; set; }
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

    internal static int FreezeNavigableReferences(
        Document document,
        string targetBookmarkName)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(targetBookmarkName)) return 0;

        Fields? outerFields = null;
        var frozen = 0;
        try
        {
            outerFields = document.Fields;
            for (var outerIndex = outerFields.Count; outerIndex >= 1; outerIndex--)
            {
                Field? outerField = null;
                Range? outerCode = null;
                Fields? nestedFields = null;
                Field? nestedField = null;
                Range? nestedCode = null;
                try
                {
                    outerField = outerFields[outerIndex];
                    if (outerField.Type != WdFieldType.wdFieldGoToButton)
                        continue;
                    outerCode = outerField.Code;
                    nestedFields = outerCode.Fields;
                    var matches = false;
                    for (var nestedIndex = 1; nestedIndex <= nestedFields.Count; nestedIndex++)
                    {
                        Release(nestedCode);
                        nestedCode = null;
                        Release(nestedField);
                        nestedField = nestedFields[nestedIndex];
                        if (nestedField.Type != WdFieldType.wdFieldRef)
                            continue;
                        nestedCode = nestedField.Code;
                        if (!TryReadVisualTeXNumberBookmark(
                                nestedCode.Text,
                                out var bookmarkName)
                            || !string.Equals(
                                bookmarkName,
                                targetBookmarkName,
                                StringComparison.OrdinalIgnoreCase))
                            continue;
                        try { nestedField.Update(); } catch { }
                        matches = true;
                        break;
                    }
                    if (!matches) continue;

                    // The nested REF lives in GOTOBUTTON.Code, so document.Fields
                    // never enumerates it as an ordinary top-level REF. Unlinking
                    // only after the nested result is current lets Word replace the
                    // complete navigable field tree with exactly the visible number.
                    // This is the required semantic when the target equation itself
                    // is restored to plain LaTeX and its bookmark is about to vanish.
                    outerField.Unlink();
                    frozen++;
                }
                finally
                {
                    Release(nestedCode);
                    Release(nestedField);
                    Release(nestedFields);
                    Release(outerCode);
                    Release(outerField);
                }
            }
        }
        finally { Release(outerFields); }
        return frozen;
    }

    internal static int UpdateNavigableReferences(
        Document document,
        ISet<string>? targetBookmarkNames = null)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));

        Fields? outerFields = null;
        var updated = 0;
        try
        {
            outerFields = document.Fields;
            for (var outerIndex = 1; outerIndex <= outerFields.Count; outerIndex++)
            {
                Field? outerField = null;
                Range? outerCode = null;
                Fields? nestedFields = null;
                Field? nestedField = null;
                Range? nestedCode = null;
                Range? nestedResult = null;
                try
                {
                    outerField = outerFields[outerIndex];
                    if (outerField.Type != WdFieldType.wdFieldGoToButton)
                        continue;

                    outerCode = outerField.Code;
                    nestedFields = outerCode.Fields;
                    for (var nestedIndex = 1;
                         nestedIndex <= nestedFields.Count;
                         nestedIndex++)
                    {
                        Release(nestedResult);
                        nestedResult = null;
                        Release(nestedCode);
                        nestedCode = null;
                        Release(nestedField);
                        nestedField = nestedFields[nestedIndex];
                        if (nestedField.Type != WdFieldType.wdFieldRef)
                            continue;

                        nestedCode = nestedField.Code;
                        if (!TryReadVisualTeXNumberBookmark(
                                nestedCode.Text,
                                out var bookmarkName))
                            continue;
                        if (targetBookmarkNames is not null
                            && !targetBookmarkNames.Contains(bookmarkName))
                            continue;

                        nestedResult = nestedField.Result;
                        var formatting = CaptureReferenceHostFormatting(
                            document,
                            outerField,
                            nestedResult);

                        // The REF is nested inside GOTOBUTTON.Code and therefore is
                        // not part of document.Fields' top-level enumeration. Update
                        // it directly. Never touch the target #(SEQ) field or its
                        // mathematical Field.Code.Text.
                        nestedField.Update();

                        // Word may rematerialize the outer field tree after a nested
                        // update. Reacquire it before restoring the user-visible
                        // character formatting inherited when the reference was
                        // inserted.
                        Release(nestedResult);
                        nestedResult = null;
                        Release(nestedCode);
                        nestedCode = null;
                        Release(nestedField);
                        nestedField = null;
                        Release(nestedFields);
                        nestedFields = null;
                        Release(outerCode);
                        outerCode = outerField.Code;
                        NormalizeInternalStyle(outerCode);
                        nestedFields = outerCode.Fields;
                        for (var refreshedIndex = 1;
                             refreshedIndex <= nestedFields.Count;
                             refreshedIndex++)
                        {
                            Release(nestedCode);
                            nestedCode = null;
                            Release(nestedField);
                            nestedField = nestedFields[refreshedIndex];
                            if (nestedField.Type != WdFieldType.wdFieldRef)
                                continue;
                            nestedCode = nestedField.Code;
                            if (!TryReadVisualTeXNumberBookmark(
                                    nestedCode.Text,
                                    out var refreshedBookmark)
                                || !string.Equals(
                                    refreshedBookmark,
                                    bookmarkName,
                                    StringComparison.OrdinalIgnoreCase))
                                continue;
                            NormalizeInternalStyle(nestedCode);
                            ApplyFormatting(nestedCode, formatting);
                            nestedResult = nestedField.Result;
                            NormalizeInternalStyle(nestedResult);
                            ApplyFormatting(nestedResult, formatting);
                            break;
                        }
                        try { outerField.ShowCodes = false; } catch { }
                        updated++;
                        break;
                    }
                }
                catch (COMException)
                {
                    // A protected or temporarily busy field must not prevent the
                    // remaining references from refreshing. Ordinary REF fields are
                    // still handled by WordEquationNumbering's normal pass.
                }
                finally
                {
                    Release(nestedResult);
                    Release(nestedCode);
                    Release(nestedField);
                    Release(nestedFields);
                    Release(outerCode);
                    Release(outerField);
                }
            }
        }
        finally { Release(outerFields); }
        return updated;
    }

    private static bool TryReadVisualTeXNumberBookmark(
        string? code,
        out string bookmarkName)
    {
        bookmarkName = string.Empty;
        if (string.IsNullOrWhiteSpace(code)) return false;
        var match = Regex.Match(
            code!,
            @"^\s*REF\s+(?:""(?<quoted>[^""]+)""|(?<plain>[^\s\\]+))",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        if (!match.Success) return false;
        var candidate = match.Groups["quoted"].Success
            ? match.Groups["quoted"].Value
            : match.Groups["plain"].Value;
        if (!candidate.StartsWith("VTEqNum_", StringComparison.OrdinalIgnoreCase))
            return false;
        bookmarkName = candidate;
        return true;
    }

    private static CharacterFormatting CaptureReferenceHostFormatting(
        Document document,
        Field outerField,
        Range fallbackRange)
    {
        Range? outerCode = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? probe = null;
        try
        {
            outerCode = outerField.Code;
            if (outerCode.StoryType != WdStoryType.wdMainTextStory)
                return CaptureFormatting(fallbackRange);
            paragraphs = outerCode.Paragraphs;
            if (paragraphs.Count != 1)
                return CaptureFormatting(fallbackRange);
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            var fieldStart = Math.Max(
                paragraphRange.Start,
                outerCode.Start - 1);

            // The visible character immediately before GOTOBUTTON (for example
            // the user-typed prefix or '(') is the most faithful source of the
            // surrounding body formatting. It is outside Word's field tree, so a
            // field update cannot silently make it bold or switch its typeface.
            if (fieldStart > paragraphRange.Start)
            {
                probe = document.Range(fieldStart - 1, fieldStart);
                var text = probe.Text ?? string.Empty;
                if (text.Length > 0 && text[0] != '\r' && text[0] != '\a')
                    return CaptureFormatting(probe);
                Release(probe);
                probe = null;
            }

            // At paragraph start there may be no preceding character. The paragraph
            // mark still carries the body style and is likewise outside the nested
            // REF result that Word can rematerialize during renumbering.
            if (paragraphRange.End > paragraphRange.Start)
            {
                probe = paragraphRange.Duplicate;
                probe.SetRange(paragraphRange.End - 1, paragraphRange.End);
                return CaptureFormatting(probe);
            }
            return CaptureFormatting(fallbackRange);
        }
        catch
        {
            return CaptureFormatting(fallbackRange);
        }
        finally
        {
            Release(probe);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(outerCode);
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
            try
            {
                var value = font.Position;
                if (value != (int)WdConstants.wdUndefined) formatting.Position = value;
            }
            catch { }
            static string? ReadName(Func<string?> read)
            {
                try
                {
                    var value = read();
                    return string.IsNullOrWhiteSpace(value) ? null : value;
                }
                catch { return null; }
            }
            formatting.Name = ReadName(() => font.Name);
            formatting.NameAscii = ReadName(() => font.NameAscii);
            formatting.NameFarEast = ReadName(() => font.NameFarEast);
            formatting.NameBi = ReadName(() => font.NameBi);
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
            if (formatting.Position.HasValue)
            {
                try { font.Position = formatting.Position.Value; } catch { }
            }
            if (!string.IsNullOrWhiteSpace(formatting.Name))
            {
                try { font.Name = formatting.Name; } catch { }
            }
            if (!string.IsNullOrWhiteSpace(formatting.NameAscii))
            {
                try { font.NameAscii = formatting.NameAscii; } catch { }
            }
            if (!string.IsNullOrWhiteSpace(formatting.NameFarEast))
            {
                try { font.NameFarEast = formatting.NameFarEast; } catch { }
            }
            if (!string.IsNullOrWhiteSpace(formatting.NameBi))
            {
                try { font.NameBi = formatting.NameBi; } catch { }
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
