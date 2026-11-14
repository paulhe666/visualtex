using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>Preserves the identified paragraph's style and direct paragraph format
/// while its generated formula host is replaced. Character formatting is separate.</summary>
internal sealed class WordParagraphFormatting : IDisposable
{
    private readonly string? styleName;
    private ParagraphFormat? format;

    private WordParagraphFormatting(string? styleName, ParagraphFormat format)
    {
        this.styleName = styleName;
        this.format = format;
    }

    internal static WordParagraphFormatting Capture(Range owner)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        Style? style = null;
        Range? paragraphMark = null;
        ParagraphFormat? paragraphFormat = null;
        try
        {
            paragraphs = owner.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException("Paragraph formatting requires one identified owner.");
            paragraph = paragraphs[1];
            range = paragraph.Range;
            object? styleObject = range.get_Style();
            style = styleObject as Style;

            // Range.Style is null when one paragraph contains mixed character
            // styles (for example an EMBED field plus a REF result), even though
            // the paragraph itself still has one stable paragraph style. Resolve
            // the final paragraph mark separately before deciding the style is
            // unavailable.
            if (style is null && range.End > range.Start)
            {
                paragraphMark = range.Document.Range(
                    range.End - 1,
                    range.End);
                styleObject = paragraphMark.get_Style();
                style = styleObject as Style;
            }

            paragraphFormat = range.ParagraphFormat;
            return new WordParagraphFormatting(
                style?.NameLocal,
                paragraphFormat.Duplicate);
        }
        finally
        {
            Release(paragraphFormat);
            Release(paragraphMark);
            Release(style);
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal void Apply(Range owner)
    {
        if (format is null) throw new ObjectDisposedException(nameof(WordParagraphFormatting));
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? mark = null;
        Style? currentStyle = null;
        ParagraphFormat? current = null;
        try
        {
            paragraphs = owner.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException("The replacement no longer owns one paragraph.");
            // Writing an unchanged format back is not neutral: WPS (and Word for
            // list/heading values) stores every inherited value as direct
            // formatting. When the retained paragraph mark still has the same
            // style and visible layout, leave the paragraph alone.
            paragraph = paragraphs[1];
            mark = paragraph.Range;
            if (mark.End > mark.Start) mark.SetRange(mark.End - 1, mark.End);
            currentStyle = mark.get_Style() as Style;
            current = paragraph.Format;
            if ((string.IsNullOrWhiteSpace(styleName)
                    || string.Equals(currentStyle?.NameLocal, styleName, StringComparison.Ordinal))
                && SameLayout(current, format))
                return;
            // All current replacement callers retain the original
            // paragraph mark. If Word still cannot expose a Style object after the
            // paragraph-mark fallback, leaving Style untouched is safer than
            // coercing user content to Normal; the retained mark already carries
            // the original style. Direct paragraph formatting is always restored.
            if (!string.IsNullOrWhiteSpace(styleName))
            {
                object style = styleName!;
                owner.set_Style(ref style);
            }
            owner.ParagraphFormat = format;
        }
        finally
        {
            Release(current);
            Release(currentStyle);
            Release(mark);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool SameLayout(ParagraphFormat a, ParagraphFormat b)
    {
        static bool Near(float x, float y) => Math.Abs(x - y) <= 0.05f;
        TabStops? tabsA = null, tabsB = null;
        try
        {
            if (a.Alignment != b.Alignment
                || !Near(a.LeftIndent, b.LeftIndent)
                || !Near(a.RightIndent, b.RightIndent)
                || !Near(a.FirstLineIndent, b.FirstLineIndent)
                || !Near(a.SpaceBefore, b.SpaceBefore)
                || !Near(a.SpaceAfter, b.SpaceAfter)
                || a.SpaceBeforeAuto != b.SpaceBeforeAuto
                || a.SpaceAfterAuto != b.SpaceAfterAuto
                || a.LineSpacingRule != b.LineSpacingRule
                || !Near(a.LineSpacing, b.LineSpacing)
                || a.BaseLineAlignment != b.BaseLineAlignment
                || a.OutlineLevel != b.OutlineLevel
                || a.KeepTogether != b.KeepTogether
                || a.KeepWithNext != b.KeepWithNext
                || a.WidowControl != b.WidowControl
                || a.PageBreakBefore != b.PageBreakBefore)
                return false;
            try
            {
                if (!Near(a.CharacterUnitLeftIndent, b.CharacterUnitLeftIndent)
                    || !Near(a.CharacterUnitRightIndent, b.CharacterUnitRightIndent)
                    || !Near(a.CharacterUnitFirstLineIndent, b.CharacterUnitFirstLineIndent))
                    return false;
            }
            catch (COMException) { }
            tabsA = a.TabStops;
            tabsB = b.TabStops;
            return tabsA.Count == tabsB.Count;
        }
        catch (COMException) { return false; }
        finally
        {
            Release(tabsB);
            Release(tabsA);
        }
    }

    public void Dispose()
    {
        Release(format);
        format = null;
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
            Marshal.ReleaseComObject(value);
    }
}
