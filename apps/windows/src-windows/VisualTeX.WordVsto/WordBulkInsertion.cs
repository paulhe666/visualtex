using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using static VisualTeX.WordVsto.WordFormulaService;
using static VisualTeX.WordVsto.WordDisplayInsertion;
using static VisualTeX.WordVsto.WordFormulaHostSetup;
using static VisualTeX.WordVsto.WordOmmlMathFont;

namespace VisualTeX.WordVsto;

/// <summary>
/// Writing many formulas and the text between them in one pass: OMML typography for the batch, native text runs and
/// the paragraph formatting around them.
/// </summary>
internal static class WordBulkInsertion
{
    internal static void ApplyBulkOmmlTypographyXml(
        XElement equation,
        double fontSizePt,
        FormulaMetadata metadata,
        XNamespace word,
        XNamespace math)
    {
        var halfPoints = ((int)Math.Round(
            FormulaFontSize.NormalizeWordOmmlSize(fontSizePt) * 2.0,
            MidpointRounding.AwayFromZero))
            .ToString(System.Globalization.CultureInfo.InvariantCulture);

        // Keep ordinary variables, digits, Greek letters, operators and all native
        // structures as genuine Office Math. Their glyph selection and spacing must
        // come from the document-level m:mathFont and its OpenType MATH table—not
        // from per-run w:rFonts, run splitting, or m:nor normal-text conversion.
        foreach (var mathRun in equation.DescendantsAndSelf(math + "r"))
        {
            var mathProperties = mathRun.Element(math + "rPr");
            var normalText = mathProperties?.Element(math + "nor") is not null;
            var properties = mathRun.Element(word + "rPr");
            if (properties is null)
            {
                properties = new XElement(word + "rPr");
                if (mathProperties is not null) mathProperties.AddAfterSelf(properties);
                else mathRun.AddFirst(properties);
            }

            if (!normalText)
            {
                // A native math run must inherit the document's Office Math font.
                // Explicit run fonts bypass that mechanism and can flatten MATH-table
                // italic correction, operator spacing and extensible constructions.
                properties.Element(word + "rFonts")?.Remove();
            }
            else if (ContainsChineseOmmlText(mathRun, math))
            {
                // m:nor is retained only when it already carries real text semantics
                // (for example MathML mtext / LaTeX \text{中文}). Select only the
                // East-Asian text face; do not override the mathematical ASCII runs.
                var fonts = properties.Element(word + "rFonts");
                if (fonts is null)
                {
                    fonts = new XElement(word + "rFonts");
                    properties.AddFirst(fonts);
                }
                fonts.SetAttributeValue(word + "eastAsiaTheme", null);
                fonts.SetAttributeValue(
                    word + "eastAsia",
                    ResolveOmmlChineseFont(metadata.FormulaChineseFont));
            }

            ApplyNativeOmmlSizeAndPosition(properties, halfPoints, word);
        }

        // Fraction bars, radicals, delimiters and other structures may carry
        // m:ctrlPr/w:rPr independently of visible runs. Size them consistently,
        // while removing any control-level font override so the same MATH font
        // drives glyph assembly and spacing throughout the complete OMath tree.
        foreach (var controlProperties in equation.DescendantsAndSelf(math + "ctrlPr"))
        {
            var properties = controlProperties.Element(word + "rPr");
            if (properties is null)
            {
                properties = new XElement(word + "rPr");
                controlProperties.Add(properties);
            }
            properties.Element(word + "rFonts")?.Remove();
            ApplyNativeOmmlSizeAndPosition(properties, halfPoints, word);
        }

        // Genuine Word display OMML obtains fraction, radical, matrix and large-
        // operator sizing from the OpenType MATH table. Never apply the retired
        // numbered-inline 1.5x numerator/denominator compensation here.

        // Numbering is paragraph structure, never mathematical content. Keep the
        // OMath tree free of generated '#(...)' wrappers so Word does not turn a
        // normal formula into a full-width equation-array control. The visible REF
        // number is created outside OMath by WordEquationNumbering, exactly like the
        // accepted VisualTeX/MathType OLE tab layout.
    }

    private static bool ContainsChineseOmmlText(
        XElement mathRun,
        XNamespace math)
    {
        var text = string.Concat(mathRun.Elements(math + "t").Select(item => item.Value));
        for (var index = 0; index < text.Length;)
        {
            var width = char.IsHighSurrogate(text[index])
                && index + 1 < text.Length
                && char.IsLowSurrogate(text[index + 1])
                    ? 2
                    : 1;
            var codePoint = width == 2
                ? char.ConvertToUtf32(text[index], text[index + 1])
                : text[index];
            if (IsChineseOmmlCodePoint(codePoint)) return true;
            index += width;
        }
        return false;
    }

    private static void ApplyNativeOmmlSizeAndPosition(
        XElement properties,
        string halfPoints,
        XNamespace word)
    {
        var size = properties.Element(word + "sz");
        if (size is null)
        {
            size = new XElement(word + "sz");
            properties.Add(size);
        }
        size.SetAttributeValue(word + "val", halfPoints);

        var complexSize = properties.Element(word + "szCs");
        if (complexSize is null)
        {
            complexSize = new XElement(word + "szCs");
            properties.Add(complexSize);
        }
        complexSize.SetAttributeValue(word + "val", halfPoints);

        var position = properties.Element(word + "position");
        if (position is null)
        {
            position = new XElement(word + "position");
            properties.Add(position);
        }
        position.SetAttributeValue(word + "val", "0");
    }

    internal static void MoveSelectionAfterBulkParagraph(
        Document document,
        Selection selection,
        int paragraphStart)
    {
        var paragraph = WordCharacterFormatting.ResolveMainStoryParagraphAtPosition(document, paragraphStart);
        try { selection.SetRange(paragraph.End, paragraph.End); }
        finally { Release(paragraph); }
    }

    internal static void EnsureWritableParagraph(Selection selection)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        try
        {
            paragraphs = selection.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            range = paragraph.Range;
            if (!ContainsVisibleBodyText(range.Text))
                return;
            if (selection.Start >= range.End - 1)
                selection.TypeParagraph();
        }
        finally
        {
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static void InsertNativeTextRun(
        Document document,
        Selection selection,
        WordBulkRun run)
    {
        if (string.IsNullOrEmpty(run.Text)) return;
        using (WordOperationMetrics.Measure("Bulk.TextRun.ResetBefore"))
            ResetSelectionTransientFormatting(selection);
        var start = selection.Start;
        using (WordOperationMetrics.Measure("Bulk.TextRun.Type"))
            selection.TypeText(run.Text);
        Range? inserted = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        using var fontMetric = WordOperationMetrics.Measure("Bulk.TextRun.Font");
        try
        {
            inserted = document.Range(start, selection.Start);
            font = inserted.Font;
            // Write only what differs (see ResetTransientFont). Word reads True
            // back as -1; the value written is unchanged (1 or 0).
            if (font.Bold != (run.Bold ? -1 : 0)) font.Bold = run.Bold ? 1 : 0;
            if (font.Italic != (run.Italic ? -1 : 0)) font.Italic = run.Italic ? 1 : 0;
            if (font.StrikeThrough != (run.Strike ? -1 : 0)) font.StrikeThrough = run.Strike ? 1 : 0;
            try { if (font.DoubleStrikeThrough != 0) font.DoubleStrikeThrough = 0; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            var underline = run.Underline
                ? WdUnderline.wdUnderlineSingle
                : WdUnderline.wdUnderlineNone;
            if (font.Underline != underline) font.Underline = underline;
            if (font.Hidden != 0) font.Hidden = 0;
            if (font.Subscript != 0) font.Subscript = 0;
            if (font.Superscript != 0) font.Superscript = 0;
            if (font.Position != 0) font.Position = 0;
            try { if (font.AllCaps != 0) font.AllCaps = 0; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            try { if (font.SmallCaps != 0) font.SmallCaps = 0; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            if (run.Code)
            {
                font.Name = "Consolas";
                try { font.NameAscii = "Consolas"; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                try { font.NameFarEast = "Microsoft YaHei UI"; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            }
        }
        finally
        {
            Release(font);
            Release(inserted);
            // Range formatting can update Word's collapsed typing state. Keep
            // placeholders, following runs and the next paragraph neutral; the
            // next source run reapplies its own explicit semantics.
            ResetSelectionTransientFormatting(selection);
        }
    }

    internal static void ApplyBulkParagraphFormatting(
        Document document,
        int start,
        int end,
        WordBulkBlock block)
    {
        if (end < start) return;
        Range? range = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ListFormat? listFormat = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = document.Range(start, end);
            paragraphs = range.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            switch (block.Kind)
            {
                case WordBulkBlockKind.Heading:
                    object heading = block.Level switch
                    {
                        <= 1 => WdBuiltinStyle.wdStyleHeading1,
                        2 => WdBuiltinStyle.wdStyleHeading2,
                        3 => WdBuiltinStyle.wdStyleHeading3,
                        4 => WdBuiltinStyle.wdStyleHeading4,
                        5 => WdBuiltinStyle.wdStyleHeading5,
                        6 => WdBuiltinStyle.wdStyleHeading6,
                        7 => WdBuiltinStyle.wdStyleHeading7,
                        8 => WdBuiltinStyle.wdStyleHeading8,
                        _ => WdBuiltinStyle.wdStyleHeading9,
                    };
                    range.set_Style(ref heading);
                    break;
                case WordBulkBlockKind.Bullet:
                    listFormat = range.ListFormat;
                    listFormat.ApplyBulletDefault();
                    for (var level = 0; level < Math.Min(block.Level, 8); level++)
                        listFormat.ListIndent();
                    break;
                case WordBulkBlockKind.Numbered:
                    listFormat = range.ListFormat;
                    listFormat.ApplyNumberDefault();
                    for (var level = 0; level < Math.Min(block.Level, 8); level++)
                        listFormat.ListIndent();
                    break;
                case WordBulkBlockKind.Quote:
                    paragraph.LeftIndent = 18f;
                    paragraph.RightIndent = 9f;
                    font = range.Font;
                    font.Italic = 1;
                    break;
                case WordBulkBlockKind.Code:
                    font = range.Font;
                    font.Name = "Consolas";
                    try { font.NameAscii = "Consolas"; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                    paragraph.LeftIndent = 18f;
                    paragraph.SpaceBefore = 3f;
                    paragraph.SpaceAfter = 3f;
                    break;
            }

        }
        finally
        {
            Release(font);
            Release(listFormat);
            Release(paragraph);
            Release(paragraphs);
            Release(range);
        }
    }

    internal static void ResetNextParagraphFormatting(
        Selection selection,
        WordBulkBlockKind current,
        WordBulkBlockKind? next)
    {
        var continuingList =
            current == WordBulkBlockKind.Bullet && next == WordBulkBlockKind.Bullet
            || current == WordBulkBlockKind.Numbered && next == WordBulkBlockKind.Numbered;
        if (continuingList)
        {
            // Preserve list numbering/indentation, but never carry the previous
            // item's final bold/italic run into the next list item.
            ResetSelectionTransientFormatting(selection);
            return;
        }
        Range? caret = null;
        Range? nextParagraph = null;
        ListFormat? list = null;
        ParagraphFormat? format = null;
        try
        {
            caret = selection.Range;
            nextParagraph = WordCharacterFormatting.ResolveParagraphAtPosition(caret, selection.Start);
            // An import in the middle of existing text can leave user content in
            // the following paragraph. Only initialize an empty typing paragraph.
            if ((nextParagraph.Text ?? string.Empty).Trim('\r', '\a').Length != 0) return;
            // As in ResetTransientFont: change only what is not already reset.
            list = nextParagraph.ListFormat;
            if (list.ListType != WdListType.wdListNoNumbering) list.RemoveNumbers();
            format = nextParagraph.ParagraphFormat;
            if (format.Alignment != WdParagraphAlignment.wdAlignParagraphLeft)
                format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            if (format.LeftIndent != 0) format.LeftIndent = 0;
            if (format.RightIndent != 0) format.RightIndent = 0;
            if (format.FirstLineIndent != 0) format.FirstLineIndent = 0;
            object normal = WdBuiltinStyle.wdStyleNormal;
            using (WordOperationMetrics.Measure("Bulk.ResetNextParagraph.Style"))
                nextParagraph.set_Style(ref normal);
            using (WordOperationMetrics.Measure("Bulk.ResetNextParagraph.Selection"))
                ResetSelectionTransientFormatting(selection);
        }
        finally
        {
            Release(format);
            Release(list);
            Release(nextParagraph);
            Release(caret);
        }
    }
}
