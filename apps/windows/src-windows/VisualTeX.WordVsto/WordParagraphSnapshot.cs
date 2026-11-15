using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>The user's paragraph around a formula that is being converted.
/// A conversion replaces only the formula: its paragraph keeps the style,
/// indents, spacing and page flow it had. The target writer owns just the
/// alignment and tab stops that place the new formula; every other value
/// it changed is put back. Values are compared before writing, so nothing
/// inherited from a style becomes direct formatting.</summary>
internal sealed class WordParagraphSnapshot
{
    private const float Tolerance = 0.05f;

    private string? styleName;
    private float leftIndent, rightIndent, firstLineIndent;
    private float characterLeft, characterRight, characterFirstLine;
    private float spaceBefore, spaceAfter, lineUnitBefore, lineUnitAfter;
    private int spaceBeforeAuto, spaceAfterAuto;
    private WdLineSpacing lineSpacingRule;
    private float lineSpacing;
    private int keepTogether, keepWithNext, widowControl, pageBreakBefore;
    private bool numbered;

    internal bool UsesDefaultStyle { get; private set; }
    internal bool WithinTable { get; private set; }
    internal float LeftIndent => leftIndent;
    internal float RightIndent => rightIndent;

    private WordParagraphSnapshot() { }

    /// <summary>Reads the paragraph that contains <paramref name="position"/>.
    /// Returns null when Word cannot expose it; the conversion then keeps the
    /// writer's own layout instead of failing.</summary>
    internal static WordParagraphSnapshot? TryCaptureAt(Document document, int position)
    {
        Range? probe = null;
        try
        {
            probe = document.Range(position, position);
            return TryCapture(probe);
        }
        catch (COMException) { return null; }
        finally { Release(probe); }
    }

    internal static WordParagraphSnapshot? TryCapture(Range owner)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        ParagraphFormat? format = null;
        try
        {
            paragraphs = owner.Paragraphs;
            if (paragraphs.Count < 1) return null;
            paragraph = paragraphs[1];
            range = paragraph.Range;
            format = paragraph.Format;
            var snapshot = new WordParagraphSnapshot
            {
                styleName = ReadStyleName(range, out var usesDefault),
                leftIndent = format.LeftIndent,
                rightIndent = format.RightIndent,
                firstLineIndent = format.FirstLineIndent,
                spaceBefore = format.SpaceBefore,
                spaceAfter = format.SpaceAfter,
                spaceBeforeAuto = format.SpaceBeforeAuto,
                spaceAfterAuto = format.SpaceAfterAuto,
                lineSpacingRule = format.LineSpacingRule,
                lineSpacing = format.LineSpacing,
                keepTogether = format.KeepTogether,
                keepWithNext = format.KeepWithNext,
                widowControl = format.WidowControl,
                pageBreakBefore = format.PageBreakBefore,
                WithinTable = IsWithinTable(range),
                numbered = IsNumbered(range),
            };
            snapshot.UsesDefaultStyle = usesDefault;
            // East Asian character/line units exist in Word and WPS but not
            // every host build exposes them; points remain authoritative then.
            try
            {
                snapshot.characterLeft = format.CharacterUnitLeftIndent;
                snapshot.characterRight = format.CharacterUnitRightIndent;
                snapshot.characterFirstLine = format.CharacterUnitFirstLineIndent;
                snapshot.lineUnitBefore = format.LineUnitBefore;
                snapshot.lineUnitAfter = format.LineUnitAfter;
            }
            catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
            return snapshot;
        }
        catch (COMException) { return null; }
        finally
        {
            Release(format);
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    /// <summary>Puts back what the writer changed in the paragraph that now
    /// holds the formula. <paramref name="keepStyle"/> is false only when the
    /// target deliberately applied its own native paragraph style.</summary>
    internal void Restore(Range formula, bool keepStyle = true)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        ParagraphFormat? format = null;
        try
        {
            paragraphs = formula.Paragraphs;
            if (paragraphs.Count < 1) return;
            paragraph = paragraphs[1];
            range = paragraph.Range;
            // A numbered target can move the formula into its own layout
            // table; that table's cells are not the user's paragraph.
            if (!WithinTable && IsWithinTable(range))
            {
                Trace($"skip-layout-table formula={formula.Start}:{formula.End} paragraph={range.Start}:{range.End}", null);
                return;
            }
            format = paragraph.Format;
            Trace("before", format);

            if (keepStyle && !string.IsNullOrWhiteSpace(styleName)
                && !string.Equals(ReadStyleName(range, out _), styleName, StringComparison.Ordinal))
            {
                // Applying a paragraph style resets direct paragraph formatting,
                // so it comes first; the values below are compared afterwards.
                object style = styleName!;
                range.set_Style(ref style);
            }
            // A style can bring its own list numbering back, which the user
            // had switched off for this paragraph.
            if (keepStyle && !numbered && IsNumbered(range))
            {
                ListFormat? list = null;
                try { list = range.ListFormat; list.RemoveNumbers(); }
                catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
                finally { Release(list); }
            }
            Release(format);
            format = paragraph.Format;
            if (keepStyle) RestoreSpacing(format);
            RestoreIndent(format, keepStyle);
            Trace("after", format);
        }
        finally
        {
            Release(format);
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private void RestoreIndent(ParagraphFormat format, bool includeFirstLine)
    {
        // A character-unit indent wins over its point value in both hosts;
        // writing points would freeze it to the current font size.
        SetIndent(characterLeft, leftIndent,
            () => format.CharacterUnitLeftIndent, value => format.CharacterUnitLeftIndent = value,
            () => format.LeftIndent, value => format.LeftIndent = value);
        SetIndent(characterRight, rightIndent,
            () => format.CharacterUnitRightIndent, value => format.CharacterUnitRightIndent = value,
            () => format.RightIndent, value => format.RightIndent = value);
        if (includeFirstLine)
            SetIndent(characterFirstLine, firstLineIndent,
                () => format.CharacterUnitFirstLineIndent, value => format.CharacterUnitFirstLineIndent = value,
                () => format.FirstLineIndent, value => format.FirstLineIndent = value);
    }

    private static void SetIndent(
        float characters, float points,
        Func<float> readCharacters, Action<float> writeCharacters,
        Func<float> readPoints, Action<float> writePoints)
    {
        if (Math.Abs(characters) > Tolerance)
        {
            try
            {
                if (Math.Abs(readCharacters() - characters) > Tolerance) writeCharacters(characters);
                return;
            }
            catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
        }
        if (Math.Abs(readPoints() - points) > Tolerance) writePoints(points);
    }

    private void RestoreSpacing(ParagraphFormat format)
    {
        if (format.SpaceBeforeAuto != spaceBeforeAuto) format.SpaceBeforeAuto = spaceBeforeAuto;
        if (format.SpaceAfterAuto != spaceAfterAuto) format.SpaceAfterAuto = spaceAfterAuto;
        SetSpace(lineUnitBefore, spaceBefore, spaceBeforeAuto,
            () => format.LineUnitBefore, value => format.LineUnitBefore = value,
            () => format.SpaceBefore, value => format.SpaceBefore = value);
        SetSpace(lineUnitAfter, spaceAfter, spaceAfterAuto,
            () => format.LineUnitAfter, value => format.LineUnitAfter = value,
            () => format.SpaceAfter, value => format.SpaceAfter = value);
        if (format.LineSpacingRule != lineSpacingRule) format.LineSpacingRule = lineSpacingRule;
        if (Math.Abs(format.LineSpacing - lineSpacing) > Tolerance)
            try { format.LineSpacing = lineSpacing; } catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
        if (format.KeepTogether != keepTogether) format.KeepTogether = keepTogether;
        if (format.KeepWithNext != keepWithNext) format.KeepWithNext = keepWithNext;
        if (format.WidowControl != widowControl) format.WidowControl = widowControl;
        if (format.PageBreakBefore != pageBreakBefore) format.PageBreakBefore = pageBreakBefore;
    }

    private static void SetSpace(
        float lines, float points, int auto,
        Func<float> readLines, Action<float> writeLines,
        Func<float> readPoints, Action<float> writePoints)
    {
        if (auto != 0) return;
        if (Math.Abs(lines) > Tolerance)
        {
            try
            {
                if (Math.Abs(readLines() - lines) > Tolerance) writeLines(lines);
                return;
            }
            catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
        }
        if (Math.Abs(readPoints() - points) > Tolerance) writePoints(points);
    }

    private static string? ReadStyleName(Range paragraphRange, out bool usesDefault)
    {
        usesDefault = false;
        Range? mark = null;
        Style? style = null;
        Document? document = null;
        Styles? styles = null;
        Style? normal = null;
        try
        {
            // Range.Style is null for a paragraph with mixed character styles;
            // the paragraph mark always carries the paragraph style.
            mark = paragraphRange.Duplicate;
            if (mark.End > mark.Start) mark.SetRange(mark.End - 1, mark.End);
            style = mark.get_Style() as Style;
            if (style is null) return null;
            var name = style.NameLocal;
            try
            {
                document = paragraphRange.Document;
                styles = document.Styles;
                object normalId = WdBuiltinStyle.wdStyleNormal;
                normal = styles[ref normalId];
                usesDefault = string.Equals(normal.NameLocal, name, StringComparison.Ordinal);
            }
            catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
            return name;
        }
        catch (COMException) { return null; }
        finally
        {
            Release(normal);
            Release(styles);
            Release(document);
            Release(style);
            Release(mark);
        }
    }

    private void Trace(string stage, ParagraphFormat? format)
    {
        if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH")))
            return;
        var current = string.Empty;
        try
        {
            if (format is not null)
                current = $" now=L{format.LeftIndent:0.#}/R{format.RightIndent:0.#}/F{format.FirstLineIndent:0.#}/B{format.SpaceBefore:0.#}/A{format.SpaceAfter:0.#}";
        }
        catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
        WordDoubleClickHook.TraceMessage(
            $"paragraph-restore {stage} style={styleName} default={UsesDefaultStyle} numbered={numbered}"
            + $" want=L{leftIndent:0.#}/R{rightIndent:0.#}/F{firstLineIndent:0.#}/B{spaceBefore:0.#}/A{spaceAfter:0.#}{current}");
    }

    private static bool IsNumbered(Range range)
    {
        ListFormat? list = null;
        try
        {
            list = range.ListFormat;
            return list.ListType != WdListType.wdListNoNumbering;
        }
        catch (COMException) { return false; }
        finally { Release(list); }
    }

    private static bool IsWithinTable(Range range)
    {
        // Range.Tables also reports a table that starts right after the
        // paragraph mark; ask about the paragraph's first character only.
        Range? start = null;
        try
        {
            start = range.Duplicate;
            start.SetRange(range.Start, Math.Min(range.End, range.Start + 1));
            return start.get_Information(WdInformation.wdWithInTable) is true;
        }
        catch (COMException) { return false; }
        finally { Release(start); }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
            Marshal.ReleaseComObject(value);
    }
}
