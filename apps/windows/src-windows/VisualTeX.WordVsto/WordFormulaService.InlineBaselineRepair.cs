using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed class WordInlineBaselineRepairTarget
{
    internal int Start { get; set; }
    internal int End { get; set; }
    internal int CurrentPosition { get; set; }
    internal int TargetPosition { get; set; }
    internal float WidthPt { get; set; }
    internal float HeightPt { get; set; }
    internal WordInlineHostAlignment HostAlignment { get; set; }
    internal string FormulaId { get; set; } = string.Empty;
    internal string ProgId { get; set; } = string.Empty;
    internal bool IsVisualTeX { get; set; }
}

internal sealed class WordInlineBaselineRepairReport
{
    internal string DocumentId { get; set; } = string.Empty;
    internal int ScannedOleCount { get; set; }
    internal int VisualTeXInlineCount { get; set; }
    internal int MathTypeInlineCount { get; set; }
    internal int SkippedCount { get; set; }
    internal List<WordInlineBaselineRepairTarget> Targets { get; } = new();
    internal int TotalRepairCount => Targets.Count;
    internal int VisualTeXRepairCount => Targets.Count(item => item.IsVisualTeX);
    internal int MathTypeRepairCount => Targets.Count(item => !item.IsVisualTeX);
}

internal sealed partial class WordFormulaService
{
    // A baseline repair is a Word character-format operation. It never renders,
    // activates an OLE server, updates EMF/PNG, changes an extent, rewrites a
    // paragraph, or infers a baseline from the majority of unrelated equations.
    internal WordInlineBaselineRepairReport ScanInlineBaselineRepairs()
    {
        Document? document = null;
        InlineShapes? shapes = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            shapes = document.InlineShapes;
            var report = new WordInlineBaselineRepairReport
            {
                DocumentId = DocumentIdentity(document),
                ScannedOleCount = shapes.Count,
            };
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                OLEFormat? format = null;
                try
                {
                    shape = shapes[index];
                    if (shape.Type != WdInlineShapeType.wdInlineShapeEmbeddedOLEObject)
                        continue;
                    format = shape.OLEFormat;
                    var progId = format.ProgID ?? string.Empty;
                    var isVisualTeX = string.Equals(progId, FormulaOleContract.ProgId,
                        StringComparison.OrdinalIgnoreCase);
                    var isMathType = progId.StartsWith("Equation.", StringComparison.OrdinalIgnoreCase);
                    if (!isVisualTeX && !isMathType) continue;
                    range = shape.Range;
                    var metadata = isVisualTeX
                        ? WordFormulaMetadataReader.TryReadCached(shape)
                        : null;
                    if (isVisualTeX && metadata is null)
                    {
                        report.SkippedCount++;
                        continue;
                    }
                    if (isVisualTeX
                        ? !string.Equals(metadata!.DisplayMode, "inline", StringComparison.OrdinalIgnoreCase)
                        : !HasVisibleSurroundingText(range))
                        continue;
                    if (isVisualTeX) report.VisualTeXInlineCount++;
                    else report.MathTypeInlineCount++;

                    var host = ReadInlineHostAlignment(range);
                    int expected;
                    if (isVisualTeX)
                    {
                        if (host is WordInlineHostAlignment.Automatic or WordInlineHostAlignment.Baseline
                            && (metadata!.RenderHeightPx is not > 0
                                || !metadata.Baseline.HasValue
                                || metadata.Baseline.Value < 0
                                || metadata.Baseline.Value >= metadata.RenderHeightPx.Value))
                        {
                            report.SkippedCount++;
                            continue;
                        }
                        var semanticSize = FormulaFontSize.ResolveSemanticFontSize(metadata!);
                        expected = CalculateVisualTeXInlinePosition(
                            range, shape.Height, (float)(metadata!.RenderHeightPx ?? 0),
                            metadata.Baseline.HasValue ? (float?)metadata.Baseline.Value : null,
                            existingFontPosition: null,
                            sourceSemanticFontSizePoints: semanticSize,
                            targetSemanticFontSizePoints: semanticSize);
                    }
                    else
                    {
                        // MathType's native baseline/presentation remains owned by
                        // MathType. Only explicit character-center anchoring gives
                        // a source-independent, non-rendering repair in this family.
                        if (host != WordInlineHostAlignment.Center)
                        {
                            report.SkippedCount++;
                            continue;
                        }
                        expected = 0;
                    }
                    var current = ReadInlineOleWordPosition(shape);
                    if (current == expected) continue;
                    report.Targets.Add(new WordInlineBaselineRepairTarget
                    {
                        Start = range.Start, End = range.End,
                        CurrentPosition = current, TargetPosition = expected,
                        WidthPt = shape.Width, HeightPt = shape.Height,
                        HostAlignment = host, FormulaId = metadata?.FormulaId ?? string.Empty,
                        ProgId = progId, IsVisualTeX = isVisualTeX,
                    });
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    report.SkippedCount++;
                }
                finally { Release(format); Release(range); Release(shape); }
            }
            return report;
        }
        finally { Release(shapes); Release(document); }
    }

    internal int RepairInlineBaselineRepairs(WordInlineBaselineRepairReport report)
    {
        if (report is null) throw new ArgumentNullException(nameof(report));
        if (report.Targets.Count == 0) return 0;
        Document? document = null;
        UndoRecord? undo = null;
        var targets = new List<(InlineShape Shape, WordInlineBaselineRepairTarget Target)>();
        var applied = new List<(InlineShape Shape, WordInlineBaselineRepairTarget Target)>();
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, report.DocumentId);
            // Validate the complete scan before making the first change.
            foreach (var target in report.Targets)
            {
                var shape = ResolveInlineBaselineRepairTarget(document, target)
                    ?? throw new InvalidOperationException("An equation changed after the baseline scan.");
                targets.Add((shape, target));
                Range? range = null;
                try
                {
                    range = shape.Range;
                    if (range.End != target.End
                        || ReadInlineHostAlignment(range) != target.HostAlignment
                        || ReadInlineOleWordPosition(shape) != target.CurrentPosition
                        || Math.Abs(shape.Width - target.WidthPt) > 0.25f
                        || Math.Abs(shape.Height - target.HeightPt) > 0.25f)
                        throw new InvalidOperationException("An equation or its paragraph changed after the baseline scan.");
                    if (target.IsVisualTeX)
                    {
                        var metadata = WordFormulaMetadataReader.TryReadCached(shape);
                        var semanticSize = metadata is null
                            ? FormulaFontSize.DefaultPt
                            : FormulaFontSize.ResolveSemanticFontSize(metadata);
                        if (metadata?.FormulaId != target.FormulaId
                            || CalculateVisualTeXInlinePosition(range, shape.Height,
                                (float)(metadata.RenderHeightPx ?? 0),
                                metadata.Baseline.HasValue ? (float?)metadata.Baseline.Value : null,
                                existingFontPosition: null,
                                sourceSemanticFontSizePoints: semanticSize,
                                targetSemanticFontSizePoints: semanticSize)
                                != target.TargetPosition)
                            throw new InvalidOperationException("An equation's render geometry changed after the baseline scan.");
                    }
                }
                finally { Release(range); }
            }
            undo = BeginUndoRecord("VisualTeX Repair Inline Formula Baselines")
                ?? throw new InvalidOperationException("Word could not establish a baseline repair undo record.");
            foreach (var item in targets)
            {
                applied.Add(item);
                SetInlineOleResultPositionOnly(document, item.Shape, item.Target.TargetPosition);
                if (ReadInlineOleWordPosition(item.Shape) != item.Target.TargetPosition
                    || Math.Abs(item.Shape.Width - item.Target.WidthPt) > 0.25f
                    || Math.Abs(item.Shape.Height - item.Target.HeightPt) > 0.25f)
                    throw new InvalidOperationException("Word did not preserve equation geometry during baseline repair.");
            }
            WordDoubleClickHook.TraceMessage(
                $"inline-baseline-repair-complete scanned={report.ScannedOleCount} "
                + $"visualTeX={report.VisualTeXRepairCount} mathType={report.MathTypeRepairCount} "
                + $"repaired={applied.Count} previewUpdates=0 extentWrites=0 paragraphWrites=0");
            return applied.Count;
        }
        catch
        {
            if (document is not null)
                foreach (var item in applied.AsEnumerable().Reverse())
                    SetInlineOleResultPositionOnly(document, item.Shape, item.Target.CurrentPosition);
            throw;
        }
        finally
        {
            EndUndoRecord(undo);
            Release(undo);
            foreach (var item in targets) Release(item.Shape);
            Release(document);
        }
    }

    private static void SetInlineOleResultPositionOnly(Document document, InlineShape shape, int position)
    {
        Range? range = null;
        Range? character = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = shape.Range;
            for (var index = range.Start; index < range.End; index++)
            {
                Release(character);
                character = document.Range(index, index + 1);
                if (character.Text != "\u0001") continue;
                font = character.Font;
                if (font.Position != position) font.Position = position;
                return;
            }
            throw new InvalidDataException("The inline OLE object has no painted result character.");
        }
        finally { Release(font); Release(character); Release(range); }
    }

    private static InlineShape? ResolveInlineBaselineRepairTarget(
        Document document, WordInlineBaselineRepairTarget target)
    {
        Range? window = null;
        InlineShapes? shapes = null;
        try
        {
            window = document.Range(Math.Max(0, target.Start - 1), target.End + 1);
            shapes = window.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                var shape = shapes[index];
                Range? range = null;
                OLEFormat? format = null;
                var keep = false;
                try
                {
                    range = shape.Range;
                    if (range.Start != target.Start || range.End != target.End) continue;
                    format = shape.OLEFormat;
                    if (!string.Equals(format.ProgID, target.ProgId, StringComparison.OrdinalIgnoreCase)) continue;
                    keep = true;
                    return shape;
                }
                finally { Release(format); Release(range); if (!keep) Release(shape); }
            }
            return null;
        }
        finally { Release(shapes); Release(window); }
    }
}
