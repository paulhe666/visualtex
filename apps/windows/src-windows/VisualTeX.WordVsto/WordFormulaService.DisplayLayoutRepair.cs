using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>What the display-formula repair found wrong and put back (formulas, then each kind of fault).</summary>
internal sealed class WordDisplayLayoutRepairReport
{
    /// <summary>Display formulas looked at (VisualTeX, Word equations, MathType).</summary>
    internal int Checked { get; set; }
    /// <summary>Formulas whose line was put back (one or more faults).</summary>
    internal int Repaired { get; set; }
    /// <summary>Formulas Word could not describe; left as they are.</summary>
    internal int Skipped { get; set; }
    /// <summary>First-line, left or right indents on the line (Ctrl+A + Paragraph, 首行缩进 2 字符).</summary>
    internal int Indents { get; set; }
    /// <summary>Centre/right tab stops missing or not those of the column (Ctrl+Q, a style, the format painter, margins).</summary>
    internal int TabStops { get; set; }
    /// <summary>Exact line spacing (固定值), which cuts off formulas taller than the line.</summary>
    internal int LineSpacing { get; set; }
    /// <summary>A paragraph alignment the line's layout does not use.</summary>
    internal int Alignment { get; set; }
    /// <summary>Bullets or list numbering on the line.</summary>
    internal int ListNumbering { get; set; }
    /// <summary>The tab before the formula, gone.</summary>
    internal int LeadingTab { get; set; }
    /// <summary>A VisualTeX or MathType object moved off its baseline (character formatting cleared).</summary>
    internal int Baseline { get; set; }
}

// One click puts every display formula's line back the way VisualTeX writes it, after the changes users make to a
// whole document (measured in Word 16, 2026-11-22, docs/office-display-formula-layout.md):
// - a first-line or left indent moves a formula laid out on tab stops (VisualTeX, MathType) and its number off the
//   centre by the indent; Ctrl+Q, Clear All Formatting, the Normal style or the format painter take the line's tab
//   stops away (the formula goes to the left margin, the number right behind it or onto the next line);
// - exact line spacing (固定值) cuts off every formula taller than the line, Word equations included;
// - bullets or numbering put a list's indent on the line; Clear All Formatting and Ctrl+Space also drop the
//   offset that lowers a VisualTeX or MathType object onto the line's baseline.
// A Word equation (OMML) is centred by Word itself: only its indents, alignment, spacing and list numbering count.
// Line spacing stays the paragraph's own (the body's: Word's and MathType's display equations keep it); only exact
// spacing becomes "at least" the same value, so a formula taller than the line gets its room and a shorter one keeps
// the line's height. A formula is changed only where it is wrong, all of it in one Undo step.
internal sealed partial class WordFormulaService
{
    private enum DisplayLineKind
    {
        /// <summary>A VisualTeX object on the centre/right tab stops of its column.</summary>
        VisualTeXOle,
        /// <summary>A MathType object on the centre/right tab stops of its column (MathType's own layout).</summary>
        MathTypeOle,
        /// <summary>A Word equation, centred (and numbered) by Word itself.</summary>
        Omml,
    }

    private sealed class DisplayLineTarget
    {
        internal DisplayLineKind Kind { get; set; }
        internal WordFormulaRangeAddress Range { get; set; } = new();
        internal FormulaMetadata? Metadata { get; set; }
        /// <summary>A MathType object's offset on its line as MathType places it (null: not checked).</summary>
        internal int? ObjectPosition { get; set; }
    }

    /// <summary>The repair of the active document's display formulas, in one Undo step (none when nothing is wrong).</summary>
    internal WordDisplayLayoutRepairReport RepairDisplayFormulaLayout()
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(NoActiveDocumentMessage);
            EnsureWritable(document);
            using var metric = WordOperationMetrics.Measure("Display.LayoutRepair");
            var targets = CaptureDisplayLineTargets(document);
            var report = new WordDisplayLayoutRepairReport { Checked = targets.Count };
            if (targets.Count == 0) return report;
            ResolveMathTypeObjectPositions(document, targets);
            // the sections read once (a formula's stops come from its column); positions are the ones of the start,
            // valid while the repair goes from the end of the document to its start
            var columns = WordDisplayColumns.BatchCache.BySectionPosition(document);
            // Reading first: a document without a fault gets no Undo record.
            var faulty = targets.Where(target => RepairDisplayLine(document, columns, target, apply: false, report: null)).ToList();
            if (faulty.Count == 0) return report;
            WordFormulaMutationTransaction.Execute(
                _application,
                document,
                "VisualTeX Repair Display Formulas",
                () =>
                {
                    // from the end: the one text change (a missing leading tab) moves only what follows it
                    foreach (var target in faulty.OrderByDescending(item => item.Range.Start))
                    {
                        if (RepairDisplayLine(document, columns, target, apply: true, report))
                            report.Repaired++;
                    }
                    return report.Repaired;
                });
            return report;
        }
        finally { Release(document); }
    }

    /// <summary>The document's display formulas: VisualTeX's and Word's from the formula index, MathType's own.</summary>
    private static List<DisplayLineTarget> CaptureDisplayLineTargets(Document document)
    {
        var targets = new List<DisplayLineTarget>();
        var index = WordFormulaHostResolver.CaptureDocumentIndex(document);
        foreach (var host in index.VisualTeX)
        {
            if (!host.Display || host.Range.StoryType != WdStoryType.wdMainTextStory) continue;
            targets.Add(new DisplayLineTarget { Kind = DisplayLineKind.VisualTeXOle, Range = host.Range, Metadata = host.Metadata });
        }
        foreach (var host in index.Omml)
        {
            if (!host.Display || host.Range.StoryType != WdStoryType.wdMainTextStory) continue;
            targets.Add(new DisplayLineTarget { Kind = DisplayLineKind.Omml, Range = host.Range });
        }
        if (index.HasForeignOle)
        {
            foreach (var host in WordMathTypeHostAdapter.CaptureLayoutIndex(document))
            {
                if (!host.Display || host.Range.StoryType != WdStoryType.wdMainTextStory) continue;
                targets.Add(new DisplayLineTarget { Kind = DisplayLineKind.MathTypeOle, Range = host.Range });
            }
        }
        return targets;
    }

    /// <summary>
    /// Whether the line of one display formula is wrong; with <paramref name="apply"/>, it is put back and the faults
    /// counted in <paramref name="report"/>. A formula Word cannot describe is skipped.
    /// </summary>
    private static bool RepairDisplayLine(
        Document document,
        WordDisplayColumns.BatchCache columns,
        DisplayLineTarget target,
        bool apply,
        WordDisplayLayoutRepairReport? report)
    {
        Range? range = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        ListFormat? list = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        try
        {
            range = WordFormulaHostSemanticReader.CreateRange(document, target.Range);
            paragraphs = range.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            format = paragraph.Format;
            list = paragraphRange.ListFormat;
            var tabLaid = target.Kind != DisplayLineKind.Omml;
            var faults = 0;

            // bullets or list numbering (they bring a hanging indent too)
            var listed = list.ListType != WdListType.wdListNoNumbering;
            if (listed)
            {
                faults++;
                if (apply)
                {
                    list.RemoveNumbers(WdNumberType.wdNumberParagraph);
                    report!.ListNumbering++;
                }
            }

            // indents, in points or in characters (首行缩进 2 字符): the line starts at the margin
            if (HasIndent(format))
            {
                faults++;
                if (apply)
                {
                    WordEquationParagraphLayout.ZeroIndents(format);
                    report!.Indents++;
                }
            }

            // alignment: a tab-laid line is justified (Word, MathType), a Word equation's paragraph centred
            var alignment = format.Alignment;
            var wantedAlignment = tabLaid
                ? WdParagraphAlignment.wdAlignParagraphJustify
                : WdParagraphAlignment.wdAlignParagraphCenter;
            var alignmentFine = tabLaid
                ? alignment is WdParagraphAlignment.wdAlignParagraphJustify or WdParagraphAlignment.wdAlignParagraphLeft
                : alignment == WdParagraphAlignment.wdAlignParagraphCenter;
            if (!alignmentFine)
            {
                faults++;
                if (apply)
                {
                    format.Alignment = wantedAlignment;
                    report!.Alignment++;
                }
            }

            // the centre and right stops of the column (or table cell) the line is in, and no stop before them
            if (tabLaid)
            {
                var expected = columns.TabStopsFor(document, paragraphRange, target.Range.Start)
                    ?? WordDisplayColumns.TabStops(document, paragraphRange);
                if (!HasDisplayTabStops(format, expected))
                {
                    faults++;
                    if (apply)
                    {
                        SetDisplayTabStops(format, expected);
                        report!.TabStops++;
                    }
                }
            }

            // exact line spacing: "at least" the same value
            if (format.LineSpacingRule == WdLineSpacing.wdLineSpaceExactly)
            {
                faults++;
                if (apply)
                {
                    var value = format.LineSpacing;
                    format.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast;
                    format.LineSpacing = value;
                    report!.LineSpacing++;
                }
            }

            // the tab the formula stands behind (a tab-laid line starts with it)
            if (tabLaid && !HasLeadingTab(document, range, paragraphRange))
            {
                faults++;
                if (apply)
                {
                    WordEquationParagraphLayout.EnsureLeadingEquationTab(document, range);
                    report!.LeadingTab++;
                }
            }

            // a VisualTeX object's baseline offset (from its own render baseline)
            if (target.Kind == DisplayLineKind.VisualTeXOle && target.Metadata is { } metadata && HasRenderBaseline(metadata))
            {
                shapes = range.InlineShapes;
                if (shapes.Count == 1)
                {
                    shape = shapes[1];
                    var height = shape.Height;
                    var current = WordInlineOlePosition.ReadInlineOleWordPosition(shape);
                    var wanted = WordNumberedDisplayLayout.DisplayOlePosition(height, metadata);
                    // VisualTeX lowers the object for the size it gives it, which Word keeps only to the
                    // twip: the offset is right for that size too (one point apart for some heights).
                    if (current != wanted && WrittenHeight(metadata, height) is { } written)
                        wanted = WordNumberedDisplayLayout.DisplayOlePosition(written, metadata);
                    if (current != wanted)
                    {
                        faults++;
                        if (apply && WordNumberedDisplayLayout.TryMoveEmbeddedOleObject(range, wanted))
                            report!.Baseline++;
                    }
                }
            }

            // a MathType object's offset, where MathType's own placement differs (see ResolveMathTypeObjectPositions)
            if (target.Kind == DisplayLineKind.MathTypeOle && target.ObjectPosition is { } position)
            {
                shapes = range.InlineShapes;
                if (shapes.Count == 1)
                {
                    shape = shapes[1];
                    if (WordInlineOlePosition.ReadInlineOleWordPosition(shape) != position)
                    {
                        faults++;
                        if (apply && WordNumberedDisplayLayout.TryMoveEmbeddedOleObject(range, position))
                            report!.Baseline++;
                    }
                }
            }
            return faults > 0;
        }
        catch (Exception error) when (error is COMException or InvalidCastException or InvalidDataException or ArgumentException)
        {
            // a formula Word cannot describe now (or a host that moved): left as it is
            IgnoredErrors.Ignore(error);
            if (apply && report is not null) report.Skipped++;
            return false;
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(list);
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(range);
        }
    }

    /// <summary>
    /// MathType lowers a display equation by its depth, which only MathType's layout knows. An object that is not
    /// lowered (Clear All Formatting and Ctrl+Space put it at 0) gets MathType's own offset from VisualTeX's native
    /// renderer, scaled to the size Word shows it at; an object that is lowered keeps its offset (MathType's own, or
    /// the user's). Without the renderer, or for an equation it cannot draw, the offset is not checked.
    /// </summary>
    private static void ResolveMathTypeObjectPositions(Document document, List<DisplayLineTarget> targets)
    {
        var pending = new Dictionary<string, (DisplayLineTarget Target, float HeightScale)>(StringComparer.Ordinal);
        var mtefs = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var target in targets)
        {
            if (target.Kind != DisplayLineKind.MathTypeOle) continue;
            Range? range = null;
            InlineShapes? shapes = null;
            InlineShape? shape = null;
            try
            {
                range = WordFormulaHostSemanticReader.CreateRange(document, target.Range);
                shapes = range.InlineShapes;
                if (shapes.Count != 1) continue;
                shape = shapes[1];
                if (WordInlineOlePosition.ReadInlineOleWordPosition(shape) < 0) continue;
                var fragment = MathTypeWordOpenXml.Read(shape);
                var key = pending.Count.ToString(CultureInfo.InvariantCulture);
                mtefs[key] = WordFormulaOle.ReadMathTypeMtefFromCompoundFile(fragment.CompoundFile);
                pending[key] = (target, WordMathTypePresentation.CalculateMathTypeNativePresentationScale(
                    fragment.HeightPt, fragment.OriginalHeightPt));
            }
            catch (Exception error) when (error is COMException or InvalidDataException or InvalidCastException
                or OverflowException or System.Xml.XmlException)
            {
                IgnoredErrors.Ignore(error);
            }
            finally
            {
                Release(shape);
                Release(shapes);
                Release(range);
            }
        }
        if (pending.Count == 0 || !MathTypeNativePreviewRenderer.IsAvailable) return;
        MathTypeNativePreviewRenderer.TryRenderBatchWithRetry(mtefs, Path.GetTempPath(), out var results);
        try
        {
            foreach (var result in results)
            {
                if (!pending.TryGetValue(result.Key, out var item)) continue;
                item.Target.ObjectPosition = (int)Math.Round(
                    result.Value.WordPosition * item.HeightScale, MidpointRounding.AwayFromZero);
            }
        }
        finally
        {
            foreach (var result in results.Values) result.Dispose();
        }
    }

    private static bool HasIndent(ParagraphFormat format)
    {
        static bool Set(Func<float> read)
        {
            try { return Math.Abs(read()) > 0.05f; }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); return false; }
        }
        return Set(() => format.FirstLineIndent)
            || Set(() => format.LeftIndent)
            || Set(() => format.RightIndent)
            || Set(() => format.CharacterUnitFirstLineIndent)
            || Set(() => format.CharacterUnitLeftIndent)
            || Set(() => format.CharacterUnitRightIndent);
    }

    /// <summary>
    /// The line's own tab stops are the centre stop and the right stop of its column, and no other stop comes before
    /// the right one (a stop a style or the format painter left would take the tab before the formula).
    /// </summary>
    private static bool HasDisplayTabStops(ParagraphFormat format, (float Center, float Right) expected)
    {
        TabStops? tabs = null;
        TabStop? tab = null;
        try
        {
            tabs = format.TabStops;
            var center = false;
            var right = false;
            for (var position = 1; position <= tabs.Count; position++)
            {
                Release(tab);
                tab = tabs[position];
                if (!IsCustomTab(tab)) continue;
                var at = tab.Position;
                if (tab.Alignment == WdTabAlignment.wdAlignTabCenter && Math.Abs(at - expected.Center) <= 0.5f)
                    center = true;
                else if (tab.Alignment == WdTabAlignment.wdAlignTabRight && Math.Abs(at - expected.Right) <= 0.5f)
                    right = true;
                else if (at < expected.Right - 0.5f)
                    return false;
            }
            return center && right;
        }
        finally
        {
            Release(tab);
            Release(tabs);
        }
    }

    private static bool IsCustomTab(TabStop tab)
    {
        try { return tab.CustomTab; }
        catch (Exception ignored) { IgnoredErrors.Ignore(ignored); return true; }
    }

    /// <summary>The centre and right stops, alone (the collection is read again after each change: Word drops a
    /// stale one when the cleared stops came from the paragraph's style).</summary>
    private static void SetDisplayTabStops(ParagraphFormat format, (float Center, float Right) expected)
    {
        TabStops? tabs = null;
        TabStop? tab = null;
        try
        {
            tabs = format.TabStops;
            tabs.ClearAll();
            Release(tabs);
            tabs = format.TabStops;
            tab = tabs.Add(expected.Center, WdTabAlignment.wdAlignTabCenter, WdTabLeader.wdTabLeaderSpaces);
            Release(tab);
            Release(tabs);
            tabs = format.TabStops;
            tab = tabs.Add(expected.Right, WdTabAlignment.wdAlignTabRight, WdTabLeader.wdTabLeaderSpaces);
        }
        finally
        {
            Release(tab);
            Release(tabs);
        }
    }

    /// <summary>Whether a tab stands right before the formula, and nothing else of the line comes before it.</summary>
    private static bool HasLeadingTab(Document document, Range formulaRange, Range paragraphRange)
    {
        Range? before = null;
        try
        {
            if (formulaRange.Start <= paragraphRange.Start) return false;
            before = document.Range(paragraphRange.Start, formulaRange.Start);
            var text = (before.Text ?? string.Empty).Replace("\v", string.Empty);
            return text.Length > 0 && text.Trim('\t').Length == 0;
        }
        finally { Release(before); }
    }

    // The height VisualTeX gave the object at its font size, when Word still shows that size (an object the
    // user resized by hand has its own height, and is lowered for that one).
    private static float? WrittenHeight(FormulaMetadata metadata, float height)
    {
        if (metadata.FontSizePt is not > 0 || metadata.RenderWidthPx is not > 0 || metadata.RenderHeightPx is not > 0)
            return null;
        try
        {
            var size = FormulaFontSize.OleSizeAt(metadata, metadata.FontSizePt.Value);
            return Math.Abs(size.Height - height) <= 0.5f ? size.Height : null;
        }
        catch (InvalidOperationException)
        {
            return null;
        }
    }

    private static bool HasRenderBaseline(FormulaMetadata metadata) =>
        metadata.RenderHeightPx is > 0
        && metadata.Baseline.HasValue
        && metadata.Baseline.Value >= 0
        && metadata.Baseline.Value < metadata.RenderHeightPx.Value;
}
