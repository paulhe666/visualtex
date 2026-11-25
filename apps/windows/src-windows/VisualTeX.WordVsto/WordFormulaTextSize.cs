using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// The text size at existing formulas, for 「跟随正文 / Follow text」 (Ribbon 行内 / 行间): an inline formula takes
/// the size of the nearest visible text beside it in its paragraph (the character before it first, then after it);
/// a display formula takes the size of the nearest visible body text outside its paragraph (the paragraphs before
/// it first, then after). Formulas are not text: every formula of the document, and the whole paragraph of a
/// display formula (its tabs and its number), is stepped over, so a formula never takes the size of a neighbouring
/// formula or of a number another batch has just resized. Searching outside a paragraph steps over headings too.
/// Nothing found: the document's Normal style size.
/// The rule for a formula that LaTeX redraw writes from its source text is
/// <see cref="WordSourceFontSize.ResolveSourceFormulaFontSize"/>; both give the same size beside plain text, and a
/// display formula right below a heading takes the body text's size there too (<see cref="TrySizeBelowHeading"/>).
/// </summary>
internal static class WordFormulaTextSize
{
    // How far a probe walks in one direction before giving up (characters actually read: spans are jumped).
    private const int MaximumProbeCharacters = 400;

    internal static Dictionary<WordFormulaHostDescriptor, float> FollowSizes(
        Document document,
        IReadOnlyCollection<WordFormulaHostDescriptor> targets,
        IReadOnlyCollection<WordFormulaHostDescriptor> allFormulas)
    {
        var result = new Dictionary<WordFormulaHostDescriptor, float>();
        if (targets.Count == 0) return result;
        var paragraphs = new Dictionary<int, (int Start, int End)>();
        (int Start, int End) ParagraphOf(WordFormulaRangeAddress address)
        {
            if (paragraphs.TryGetValue(address.Start, out var known)) return known;
            Range? range = null;
            try
            {
                range = document.Range(address.Start, address.Start);
                range.Expand(WdUnits.wdParagraph);
                var span = (range.Start, range.End);
                paragraphs[address.Start] = span;
                return span;
            }
            finally { Release(range); }
        }

        // Main-story spans that are not text: every formula, and a display formula's whole paragraph.
        var spans = new List<(int Start, int End)>();
        foreach (var formula in allFormulas)
        {
            if (formula.Range.StoryType != WdStoryType.wdMainTextStory) continue;
            spans.Add((formula.Range.Start, formula.Range.End));
            if (formula.Display || formula.Numbering.Numbered)
                spans.Add(ParagraphOf(formula.Range));
        }
        var merged = Merge(spans);
        float? normalSize = null;
        float NormalSize()
        {
            if (normalSize.HasValue) return normalSize.Value;
            Styles? styles = null;
            Style? normal = null;
            Microsoft.Office.Interop.Word.Font? font = null;
            try
            {
                styles = document.Styles;
                object id = WdBuiltinStyle.wdStyleNormal;
                normal = styles[ref id];
                font = normal.Font;
                normalSize = WordFormulaService.TryResolveWordFontSize(font.Size, out var size)
                    ? size
                    : FormulaFontSize.DefaultPt;
            }
            catch (COMException) { normalSize = FormulaFontSize.DefaultPt; }
            finally { Release(font); Release(normal); Release(styles); }
            return normalSize.Value;
        }

        foreach (var target in targets)
        {
            var paragraph = ParagraphOf(target.Range);
            float? size = null;
            if (!(target.Display || target.Numbering.Numbered))
            {
                size = Probe(document, merged, target.Range.Start - 1, -1, paragraph.Start, skipHeadings: false)
                    ?? Probe(document, merged, target.Range.End, 1, paragraph.End - 1, skipHeadings: false);
            }
            size ??= Probe(document, merged, paragraph.Start - 1, -1, 0, skipHeadings: true)
                ?? Probe(document, merged, paragraph.End, 1, int.MaxValue, skipHeadings: true);
            result[target] = FormulaFontSize.Normalize(size ?? NormalSize());
        }
        return result;
    }

    /// <summary>
    /// For a display formula that LaTeX redraw writes alone on its line (<paramref name="position"/> in its
    /// paragraph) while 「行间」 follows the text: when the text before it is a heading, the size of the body text
    /// before the heading (or after the formula), as 「跟随正文」 gives an existing display formula. False when no
    /// heading is in the way: the redraw's own rule (WordSourceFontSize) stands. <paramref name="sources"/>: the
    /// LaTeX of every formula the redraw writes (<see cref="Spans"/>), stepped over as the redraw's rule does.
    /// </summary>
    internal static bool TrySizeBelowHeading(
        Document document,
        int position,
        IReadOnlyList<(int Start, int End)> sources,
        out float size)
    {
        size = FormulaFontSize.DefaultPt;
        Range? paragraph = null;
        try
        {
            paragraph = document.Range(position, position);
            paragraph.Expand(WdUnits.wdParagraph);
            var skippedHeading = false;
            var found = Probe(document, sources, paragraph.Start - 1, -1, 0, skipHeadings: true, ref skippedHeading);
            if (!skippedHeading) return false;
            found ??= Probe(document, sources, paragraph.End, 1, int.MaxValue, skipHeadings: true, ref skippedHeading);
            if (found is not { } value) return false;
            size = FormulaFontSize.Normalize(value);
            return true;
        }
        catch (COMException) { return false; }
        finally { Release(paragraph); }
    }

    private static float? Probe(
        Document document,
        IReadOnlyList<(int Start, int End)> spans,
        int start,
        int step,
        int bound,
        bool skipHeadings)
    {
        var skippedHeading = false;
        return Probe(document, spans, start, step, bound, skipHeadings, ref skippedHeading);
    }

    // The size of the first visible text character from `start` in `step` direction up to `bound` (inclusive),
    // jumping over the non-text spans (and over heading paragraphs when asked; `skippedHeading` then tells
    // whether one was jumped).
    private static float? Probe(
        Document document,
        IReadOnlyList<(int Start, int End)> spans,
        int start,
        int step,
        int bound,
        bool skipHeadings,
        ref bool skippedHeading)
    {
        Range? content = null;
        try
        {
            content = document.Content;
            var lower = step < 0 ? Math.Max(content.Start, bound) : content.Start;
            var upper = step > 0 ? Math.Min(content.End - 1, bound) : content.End - 1;
            var position = start;
            for (var read = 0; read < MaximumProbeCharacters && position >= lower && position <= upper; read++)
            {
                var span = Containing(spans, position);
                if (span.HasValue)
                {
                    position = step < 0 ? span.Value.Start - 1 : span.Value.End;
                    continue;
                }
                Range? probe = null;
                Microsoft.Office.Interop.Word.Font? font = null;
                Paragraphs? owners = null;
                Paragraph? owner = null;
                Range? ownerRange = null;
                try
                {
                    probe = document.Range(position, position + 1);
                    if (!WordFormulaHostSetup.ContainsVisibleBodyText(probe.Text))
                    {
                        position += step;
                        continue;
                    }
                    if (skipHeadings)
                    {
                        owners = probe.Paragraphs;
                        owner = owners[1];
                        if (owner.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
                        {
                            skippedHeading = true;
                            ownerRange = owner.Range;
                            position = step < 0 ? ownerRange.Start - 1 : ownerRange.End;
                            continue;
                        }
                    }
                    font = probe.Font;
                    if (WordFormulaService.TryResolveWordFontSize(font.Size, out var size))
                        return size;
                    position += step;
                }
                catch (COMException ignored)
                {
                    IgnoredErrors.Ignore(ignored);
                    position += step;
                }
                finally
                {
                    Release(ownerRange);
                    Release(owner);
                    Release(owners);
                    Release(font);
                    Release(probe);
                }
            }
            return null;
        }
        finally { Release(content); }
    }

    /// Main-story spans (start, end) sorted and merged, for <see cref="TrySizeBelowHeading"/>.
    internal static IReadOnlyList<(int Start, int End)> Spans(IEnumerable<(int Start, int End)> spans) =>
        Merge(spans.ToList());

    private static List<(int Start, int End)> Merge(List<(int Start, int End)> spans)
    {
        spans.Sort((left, right) => left.Start != right.Start
            ? left.Start.CompareTo(right.Start)
            : left.End.CompareTo(right.End));
        var merged = new List<(int Start, int End)>();
        foreach (var span in spans)
        {
            if (span.End <= span.Start) continue;
            if (merged.Count > 0 && span.Start <= merged[merged.Count - 1].End)
            {
                var last = merged[merged.Count - 1];
                merged[merged.Count - 1] = (last.Start, Math.Max(last.End, span.End));
            }
            else merged.Add(span);
        }
        return merged;
    }

    private static (int Start, int End)? Containing(IReadOnlyList<(int Start, int End)> spans, int position)
    {
        int low = 0, high = spans.Count - 1;
        while (low <= high)
        {
            var middle = (low + high) / 2;
            var span = spans[middle];
            if (position < span.Start) high = middle - 1;
            else if (position >= span.End) low = middle + 1;
            else return span;
        }
        return null;
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
            Marshal.ReleaseComObject(value);
    }
}
