using System.Diagnostics;
using System.Text;
using System.Text.Json;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // Real owned Word COM selection events, not physical mouse/keyboard timing.
    private static void RunWordProseSelectionCostAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord, "The prose probe must never attach to user Word.");
        AssertTrue(!string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(
            "VISUALTEX_ACCEPTANCE_OWNED_WORD_ROOT")), "Explicit owned Word root required.");
        Directory.CreateDirectory(artifactRoot);
        var samples = new List<object>();
        using var host = new WordPerformanceHost(documentPath: null);
        var document = host.Document;
        var application = host.Application;
        document.Content.Text = "Ordinary prose alpha beta gamma delta epsilon zeta.\r"
            + "Second ordinary text paragraph without formula objects.\r";
        var ribbon = new AcceptanceRibbonProbe();
        host.AddIn.OnRibbonLoad(ribbon);
        PumpWordUi();
        var selection = application.Selection;
        try
        {
            MeasureProseSelectionCost(host, selection, "empty-prose", 0, 8,
                i => 3 + i * 3, samples, ribbon);
            WordDoubleClickHook.TraceMessage("prose-cost phase=fixture index=0");
            var assetRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "VisualTeX", "office", "temp", "prose-cost-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(assetRoot);
            var previous = Environment.GetEnvironmentVariable("VISUALTEX_PERF_CREATE_LIMIT");
            try
            {
                Environment.SetEnvironmentVariable("VISUALTEX_PERF_CREATE_LIMIT", "20");
                CreatePerformanceCorpus(host, assetRoot);
            }
            finally { Environment.SetEnvironmentVariable("VISUALTEX_PERF_CREATE_LIMIT", previous); }
            AssertEqual(20, document.InlineShapes.Count, "Twenty real VisualTeX OLE fixtures required.");
            selection.SetRange(3, 3);
            PumpWordUi();
            MeasureProseSelectionCost(host, selection, "20-ole-distant-prose", 20, 12,
                i => 4 + i * 3, samples, ribbon);

            // Each boundary refers to a different object; includes the cold entry.
            var starts = new List<int>();
            var ends = new List<int>();
            Word.InlineShapes? shapes = null;
            Word.InlineShape? shape = null;
            Word.Range? range = null;
            try
            {
                shapes = document.InlineShapes;
                for (var i = 1; i <= 20; i++)
                {
                    shape = shapes[i]; range = shape.Range;
                    starts.Add(range.Start); ends.Add(range.End);
                    Release(range); range = null; Release(shape); shape = null;
                }
            }
            finally { Release(range); Release(shape); Release(shapes); }
            MeasureProseSelectionCost(host, selection, "20-ole-leading-boundary", 20, 10,
                i => starts[i * 2 + 1], samples, ribbon);
            MeasureProseSelectionCost(host, selection, "20-ole-trailing-boundary", 20, 10,
                i => ends[i * 2 + 1], samples, ribbon);
            MeasureProseSelectionCost(host, selection, "20-ole-prose-after", 20, 10,
                i => ends[i * 2 + 1] + 2, samples, ribbon);
            MeasureProseSelectionCost(host, selection, "20-ole-block-leading", 20, 5,
                i => starts[i * 2], samples, ribbon);
            MeasureProseSelectionCost(host, selection, "20-ole-block-trailing", 20, 5,
                i => ends[i * 2], samples, ribbon);
            WordDoubleClickHook.TraceMessage("prose-cost phase=20-ole-type-prose index=0");
            selection.SetRange(8, 8); PumpWordUi();
            for (var i = 0; i < 10; i++)
            {
                var watch = Stopwatch.StartNew();
                selection.TypeText("a");
                watch.Stop();
                samples.Add(new { Phase = "20-ole-type-prose", Index = i, ElapsedMs = watch.Elapsed.TotalMilliseconds });
                PumpWordUi();
            }
            AssertEqual(20, document.InlineShapes.Count, "Prose interaction changed formula inventory.");
            AssertTrue(document.Content.Text.Contains("aaaaaaaaaa"), "Typed prose was lost.");
            host.Save(Path.Combine(artifactRoot, "prose-20-ole.docx"));
            if (Environment.GetEnvironmentVariable("VISUALTEX_PROSE_SCALE100") == "1")
            {
                WordDoubleClickHook.TraceMessage("prose-cost phase=fixture-100 index=0");
                var previousLimit = Environment.GetEnvironmentVariable("VISUALTEX_PERF_CREATE_LIMIT");
                try
                {
                    Environment.SetEnvironmentVariable("VISUALTEX_PERF_CREATE_LIMIT", "20");
                    for (var cohort = 0; cohort < 4; cohort++)
                    {
                        var cohortAssets = Path.Combine(assetRoot, "cohort-" + cohort);
                        Directory.CreateDirectory(cohortAssets);
                        CreatePerformanceCorpus(host, cohortAssets);
                    }
                }
                finally { Environment.SetEnvironmentVariable("VISUALTEX_PERF_CREATE_LIMIT", previousLimit); }
                AssertEqual(100, document.InlineShapes.Count, "Scaling fixture must contain exactly100 OLEs.");
                selection.SetRange(3, 3); PumpWordUi();
                MeasureProseSelectionCost(host, selection, "100-ole-distant-prose", 100, 12,
                    i => 4 + i * 3, samples, ribbon);
                var lastStarts = new List<int>();
                var lastEnds = new List<int>();
                try
                {
                    shapes = document.InlineShapes;
                    for (var i = 82; i <= 100; i += 2)
                    {
                        shape = shapes[i]; range = shape.Range;
                        lastStarts.Add(range.Start); lastEnds.Add(range.End);
                        Release(range); range = null; Release(shape); shape = null;
                    }
                }
                finally { Release(range); Release(shape); Release(shapes); }
                MeasureProseSelectionCost(host, selection, "100-ole-leading-boundary", 100, 10,
                    i => lastStarts[i], samples, ribbon);
                MeasureProseSelectionCost(host, selection, "100-ole-trailing-boundary", 100, 10,
                    i => lastEnds[i], samples, ribbon);
                MeasureProseSelectionCost(host, selection, "100-ole-prose-after", 100, 10,
                    i => lastEnds[i] + 2, samples, ribbon);
                AssertEqual(100, document.InlineShapes.Count, "Scaling selections changed OLE inventory.");
                host.Save(Path.Combine(artifactRoot, "prose-100-ole.docx"));
            }
        }
        finally
        {
            Release(selection);
            File.WriteAllText(Path.Combine(artifactRoot, "selection-cost.json"),
                JsonSerializer.Serialize(samples, new JsonSerializerOptions { WriteIndented = true }), new UTF8Encoding(false));
        }
        Console.WriteLine("Prose selection probe completed: owned Word COM events, 0/20 real OLEs, distant prose, leading/trailing boundaries, typing; see exact samples and optional selection trace.");
    }

    private static void MeasureProseSelectionCost(WordPerformanceHost host,
        Word.Selection selection, string phase, int count, int repetitions,
        Func<int, int> position, List<object> samples, AcceptanceRibbonProbe ribbon)
    {
        ribbon.Reset();
        host.AddIn.ResetSelectionChangeDiagnosticsForAcceptance();
        for (var i = 0; i < repetitions; i++)
        {
            Console.WriteLine($"[PROSE COST] phase={phase} index={i} before");
            WordDoubleClickHook.TraceMessage($"prose-cost phase={phase} index={i}");
            var watch = Stopwatch.StartNew();
            var target = position(i);
            selection.SetRange(target, target);
            watch.Stop();
            samples.Add(new { Phase = phase, FormulaCount = count, Index = i,
                Position = target, ElapsedMs = watch.Elapsed.TotalMilliseconds });
            PumpWordUi();
            AssertEqual(target, selection.Start, "Passive caret work moved the requested insertion point.");
            AssertEqual(target, selection.End, "Passive caret work expanded the insertion point.");
            if (phase.EndsWith("-ole-trailing-boundary", StringComparison.Ordinal))
                AssertNear(0, selection.Font.Position, 0.1f, "Inline formula tail retained formula baseline.");
        }
        var state = host.AddIn.ReadSelectionChangeDiagnosticsForAcceptance();
        Console.WriteLine($"[PROSE COST] phase={phase} events={state.SelectionChanges} formulaReads={state.FormulaStateReads} deferred={state.DeferredCaretPasses} ribbon={ribbon.FormulaInvalidations}");
        if (phase.EndsWith("prose", StringComparison.Ordinal))
        {
            AssertEqual(0, state.FormulaStateReads, "Distant prose must not read formula state.");
            AssertEqual(0, state.DeferredCaretPasses, "Distant prose must not defer caret work.");
        }
    }
}
