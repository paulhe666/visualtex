using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using VisualTeX.WindowsOffice.VstoShared;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using View = Microsoft.Office.Interop.PowerPoint.View;

namespace VisualTeX.PowerPointVsto;

// Agents' insert_rich_text on a slide (PowerPoint, WPS Presentation): text with formulas in a new text box, or at
// the text caret of one. The text is written with its LaTeX ($...$, \(...\), $$...$$, \[...\]) as it stands; the
// LaTeX redraw then makes native equations of it in place (ThisAddIn.Automation), as latex_text_to_formulas does.
// Each line is a paragraph; "- " / "* " / "1. " lines are list items, "#" lines bold headings, **...** bold.
public sealed partial class PowerPointFormulaService
{
    internal sealed class SlideParagraph
    {
        internal string Text { get; set; } = string.Empty;
        /// <summary>"text", "bullet", "numbered", "heading" or "display" (a formula paragraph).</summary>
        internal string Kind { get; set; } = "text";
        internal List<(int Start, int Length)> Bold { get; } = new();
    }

    internal sealed class AutomationRichText
    {
        internal int SlideId { get; set; }
        internal int ShapeId { get; set; }
        internal bool NewBox { get; set; }
        internal bool HasLatex { get; set; }
        /// <summary>The written text in the box (1-based); length 0: the whole box.</summary>
        internal int TextStart { get; set; }
        internal int TextLength { get; set; }
    }

    /// <summary>The shape (and text) the next format conversion works on instead of the selection (agents' writes).</summary>
    internal (int SlideId, int ShapeId, int TextStart, int TextLength)? AutomationConversionTarget { get; set; }

    private static readonly Regex ListItem = new(@"^\s*(?:([-*+])|(\d+)[.)])\s+(.*)$");
    private static readonly Regex Heading = new(@"^\s*#{1,6}\s+(.*)$");
    private static readonly Regex BoldSpan = new(@"\*\*(.+?)\*\*");

    /// The paragraphs of an agent's text: a $$...$$ or \[...\] block over several lines is one paragraph.
    internal static List<SlideParagraph> ParseSlideText(string text)
    {
        var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var result = new List<SlideParagraph>();
        for (var index = 0; index < lines.Length; index++)
        {
            var line = lines[index].TrimEnd();
            if (line.Trim().Length == 0) continue;
            var trimmed = line.Trim();
            foreach (var (open, close) in new[] { ("$$", "$$"), ("\\[", "\\]") })
            {
                if (!trimmed.StartsWith(open, StringComparison.Ordinal)) continue;
                var block = new StringBuilder(trimmed);
                while (!(block.Length > open.Length + 1 && block.ToString().EndsWith(close, StringComparison.Ordinal))
                       && index + 1 < lines.Length)
                    block.Append(' ').Append(lines[++index].Trim());
                trimmed = null!;
                result.Add(new SlideParagraph { Text = block.ToString(), Kind = "display" });
                break;
            }
            if (trimmed is null) continue;
            var paragraph = new SlideParagraph();
            Match match;
            if ((match = Heading.Match(line)).Success)
            {
                paragraph.Kind = "heading";
                line = match.Groups[1].Value;
            }
            else if ((match = ListItem.Match(line)).Success)
            {
                paragraph.Kind = match.Groups[1].Success ? "bullet" : "numbered";
                line = match.Groups[3].Value;
            }
            else line = line.Trim();
            // **bold** (outside formulas): the markers go, the span is remembered
            var plain = new StringBuilder();
            var last = 0;
            foreach (Match bold in BoldSpan.Matches(line))
            {
                plain.Append(line, last, bold.Index - last);
                paragraph.Bold.Add((plain.Length, bold.Groups[1].Length));
                plain.Append(bold.Groups[1].Value);
                last = bold.Index + bold.Length;
            }
            plain.Append(line, last, line.Length - last);
            paragraph.Text = plain.ToString();
            result.Add(paragraph);
        }
        if (result.Count == 0)
            throw new OfficeAutomationException("invalid_arguments", "text is empty.");
        return result;
    }

    /// Writes the paragraphs: at the text caret ("cursor" with a caret in a text box) or into a new text box on the
    /// slide the insertion was placed on. The written text is left selected for the LaTeX redraw.
    internal AutomationRichText AutomationWriteRichText(
        string? key,
        AutomationInsertion placed,
        JsonElement location,
        List<SlideParagraph> paragraphs,
        double? fontSizePt)
    {
        DocumentWindow? window = null;
        View? view = null;
        Selection? selection = null;
        Slide? slide = null;
        Shapes? shapes = null;
        Shape? shape = null;
        try
        {
            window = AutomationActivate(key);
            view = window.View;
            selection = window.Selection;
            var result = new AutomationRichText
            {
                SlideId = placed.SlideId,
                HasLatex = paragraphs.Any(paragraph => PowerPointLatexText.MayContainLatex(paragraph.Text)),
            };
            if (placed.AtCursor && selection.Type == PpSelectionType.ppSelectionText)
            {
                // into the text box the caret is in, a paragraph break at a time (WPS keeps a "\r" written inside
                // a longer insertion as a character of one paragraph)
                dynamic caret = selection.TextRange;
                dynamic inserted = caret.InsertAfter(paragraphs[0].Text);
                var start = (int)inserted.Start;
                for (var index = 1; index < paragraphs.Count; index++)
                {
                    inserted = inserted.InsertAfter("\r");
                    inserted = inserted.InsertAfter(paragraphs[index].Text);
                }
                var length = (int)inserted.Start + (int)inserted.Length - start;
                shape = selection.ShapeRange[1];
                result.ShapeId = shape.Id;
                result.SlideId = ((Slide)view.Slide).SlideID;
                result.TextStart = start;
                result.TextLength = length;
                ApplyRichFormatting(shape, paragraphs, startCharacter: start, wholeBox: false, fontSizePt);
                ((dynamic)shape.TextFrame.TextRange).Characters(start, length).Select();
                return result;
            }

            slide = (Slide)view.Slide;
            result.SlideId = slide.SlideID;
            (double Width, double Height)? slideSize;
            var presentation = FindAutomationPresentation(key);
            try { slideSize = SlideSize(presentation); }
            finally { Release(presentation); }
            var left = (float)(OfficeAutomationTextNumber(location, "left_pt") ?? placed.Left ?? 60);
            var top = (float)(OfficeAutomationTextNumber(location, "top_pt") ?? placed.Top ?? 60);
            var width = (float)(OfficeAutomationTextNumber(location, "width_pt")
                ?? Math.Max(120, (slideSize?.Width ?? 960) - left - 60));
            shapes = slide.Shapes;
            shape = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, 40);
            result.ShapeId = shape.Id;
            result.NewBox = true;
            var frame = shape.TextFrame;
            frame.WordWrap = MsoTriState.msoTrue;
            frame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            // all at once, single breaks (no paragraph is empty): WPS makes paragraphs of the "\r"s of a Text
            // assignment, but keeps one appended with text as a character, and "\r\r" as one break and a CR
            frame.TextRange.Text = string.Join("\r", paragraphs.Select(paragraph => paragraph.Text));
            ApplyRichFormatting(shape, paragraphs, startCharacter: 1, wholeBox: true, fontSizePt);
            selection.Unselect();
            shape.Select();
            return result;
        }
        catch (System.Runtime.InteropServices.COMException error)
        {
            throw new OfficeAutomationException("host_error", "The text could not be written: " + error.Message);
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(slide);
            Release(selection);
            Release(view);
            Release(window);
        }
    }

    private static double? OfficeAutomationTextNumber(JsonElement location, string name) =>
        location.ValueKind == JsonValueKind.Object ? OfficeAutomationText.Number(location, name) : null;

    // Lists, headings, bold spans and the size, paragraph by paragraph (the written paragraphs start at startCharacter).
    private static void ApplyRichFormatting(Shape shape, List<SlideParagraph> paragraphs, int startCharacter, bool wholeBox, double? fontSizePt)
    {
        dynamic box = shape.TextFrame.TextRange;
        var position = startCharacter;
        var paragraphNumber = 0;
        foreach (var paragraph in paragraphs)
        {
            paragraphNumber++;
            dynamic range = box.Characters(position, paragraph.Text.Length);
            if (fontSizePt is { } size && size > 0) range.Font.Size = (float)size;
            if (wholeBox)
            {
                // by paragraph number: WPS gives a character range's ParagraphFormat to the whole box
                dynamic format = box.Paragraphs(paragraphNumber, 1).ParagraphFormat;
                switch (paragraph.Kind)
                {
                    case "bullet":
                        format.Bullet.Visible = MsoTriState.msoTrue;
                        format.Bullet.Type = PpBulletType.ppBulletUnnumbered;
                        HangingIndent(shape, paragraphNumber);
                        break;
                    case "numbered":
                        format.Bullet.Visible = MsoTriState.msoTrue;
                        format.Bullet.Type = PpBulletType.ppBulletNumbered;
                        HangingIndent(shape, paragraphNumber);
                        break;
                    case "display":
                        format.Alignment = PpParagraphAlignment.ppAlignCenter;
                        break;
                }
            }
            if (paragraph.Kind == "heading")
            {
                range.Font.Bold = MsoTriState.msoTrue;
                range.Font.Size = (float)((double)range.Font.Size * 1.2);
            }
            foreach (var (start, length) in paragraph.Bold)
                box.Characters(position + start, length).Font.Bold = MsoTriState.msoTrue;
            position += paragraph.Text.Length + 1;
        }
    }

    // A list item's text 18 pt in, its bullet at the margin (a text box's default puts the text a wide tab away).
    private static void HangingIndent(Shape shape, int paragraph)
    {
        try
        {
            dynamic format = ((dynamic)shape.TextFrame2).TextRange.Paragraphs(paragraph).ParagraphFormat;
            format.LeftIndent = 18f;
            format.FirstLineIndent = -18f;
        }
        catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
    }

    /// The text box an insert_rich_text wrote, with its frame and the formulas now in it.
    internal object AutomationDescribeRichText(string? key, AutomationRichText written)
    {
        Presentation? presentation = null;
        Slides? slides = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            presentation = FindAutomationPresentation(key);
            slides = presentation.Slides;
            slide = FindSlideById(slides, written.SlideId);
            for (var index = 1; index <= slide.Shapes.Count && shape is null; index++)
            {
                var candidate = slide.Shapes[index];
                if (candidate.Id == written.ShapeId) shape = candidate;
                else Release(candidate);
            }
            var formulas = ScanAutomationFormulas(presentation, new HashSet<string>(), 40, written.SlideId)
                .Where(formula => formula.ShapeId == written.ShapeId)
                .Select(formula => DescribeAutomationFormula(formula, full: true))
                .ToList();
            return new
            {
                slide = slide.SlideIndex,
                text_box = shape is null ? null : shape.Name,
                new_text_box = written.NewBox,
                frame = shape is null ? null : DescribeFrame(ReadFrame(shape)),
                text_preview = shape is null ? null : AutomationTextPreview(shape),
                formulas,
            };
        }
        finally
        {
            Release(shape);
            Release(slide);
            Release(slides);
            Release(presentation);
        }
    }
}
