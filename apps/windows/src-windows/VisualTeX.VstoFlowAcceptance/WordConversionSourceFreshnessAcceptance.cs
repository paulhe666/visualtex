using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // Focused three-formula repro, not the daily twenty-formula regression.
    // Register as: word-conversion-source-freshness
    // Required: VISUALTEX_SOURCE_FRESHNESS_SOURCE = owned source-MathType.docx.
    // No frontend, native preview generation, installed add-in, or live user Word.
    private static void RunWordConversionSourceFreshnessAcceptance(string artifactRoot)
    {
        var source = Environment.GetEnvironmentVariable("VISUALTEX_SOURCE_FRESHNESS_SOURCE");
        if (string.IsNullOrWhiteSpace(source) || !File.Exists(source))
            throw new InvalidDataException(
                "Set VISUALTEX_SOURCE_FRESHNESS_SOURCE to an existing owned 20-formula source-MathType.docx.");
        Directory.CreateDirectory(artifactRoot);
        var unchanged = RunSourceFreshnessCase(artifactRoot, source!, changeSource: false);
        var changed = RunSourceFreshnessCase(artifactRoot, source!, changeSource: true);
        File.WriteAllText(Path.Combine(artifactRoot, "source-freshness-result.json"),
            JsonSerializer.Serialize(new
            {
                Status = "Passed",
                SelectedFormulaCount = 3,
                SourceFixture = Path.GetFullPath(source!),
                Unchanged = unchanged,
                Changed = changed,
            }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine("[SOURCE FRESHNESS PASS] unchanged selection=3 converted; "
            + "changed selection=3 rejected before target writes/Undo; save/reopen and single Undo/Redo verified.");
    }

    private static object RunSourceFreshnessCase(string artifactRoot, string source, bool changeSource)
    {
        const int selectedCount = 3;
        const int fixtureCount = 20;
        const string expectedError =
            "The MathType source changed while its conversion was being prepared.";
        var label = changeSource ? "changed-source" : "unchanged-source";
        var output = Path.Combine(artifactRoot, label + ".docx");
        var tracePath = Path.Combine(artifactRoot, label + "-apply.log");
        // Refuse to overwrite an existing output, even on a rerun. The caller
        // must provide a fresh artifact directory; the source is never opened.
        File.Copy(source, output, overwrite: false);
        var previousTrace = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        var previousPerf = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF");
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateFreshFormatScalingWord(artifactRoot, "source-freshness-" + label);
            document = application.Documents.Open(output,
                ReadOnly: false, AddToRecentFiles: false, Visible: false);
            document.Activate();
            var service = new WordFormulaService(application);
            var original = service.CaptureFormulaFormatConversionPlan(true,
                FormulaOleContract.MathTypeOleMode, FormulaOleContract.WordOmmlMode);
            AssertEqual(fixtureCount, original.Targets.Count, label + ": expected daily-20 owned fixture.");
            AssertEqual(0, document.OMaths.Count, label + ": source fixture already contains OMML.");
            AssertEqual(0, CountVisualTeXNativeOleShapes(document),
                label + ": source fixture already contains VisualTeX OLE.");
            var expectedSelected = original.Targets.Take(selectedCount).ToArray();
            SelectSourceFreshnessScope(document, expectedSelected);
            var plan = service.CaptureFormulaFormatConversionPlan(false,
                FormulaOleContract.MathTypeOleMode, FormulaOleContract.WordOmmlMode);
            AssertEqual(selectedCount, plan.Targets.Count, label + ": selection must capture exactly three formulas.");
            AssertTrue(plan.Targets.Select(item => item.SourceObjectId)
                    .SequenceEqual(expectedSelected.Select(item => item.SourceObjectId)),
                label + ": selection did not capture the first three exact source objects.");

            // These evidence objects are detached scalar values, not live COM
            // ranges. The prepared plan/metadata are never edited to fake staleness.
            var before = ReadSourceFreshnessEvidence(application, document, plan);
            var prepared = PrepareSourceFreshnessOmml(plan);
            var paragraphsBefore = document.Paragraphs.Count;
            var topologyBefore = LargeRedrawTableTopology(document);
            var textBefore = document.Content.Text;
            SourceFreshnessEvidence[] edited = before;
            if (changeSource)
            {
                AssertTrue(before[0].SemanticSignature != before[1].SemanticSignature,
                    "Freshness mutation requires mathematically different first and second source objects.");
                AssertSourceFreshnessPresentationEqual(before[0], before[1],
                    "Freshness donor must have the same display/number/font presentation.");
                ReplaceSourceFreshnessOle(document, plan.Targets[0], plan.Targets[1]);
                edited = ReadSourceFreshnessEvidence(application, document, plan);
                for (var index = 0; index < selectedCount; index++)
                {
                    AssertSourceFreshnessAddressEqual(before[index], edited[index],
                        "Mutation changed captured location " + index + "; this is not a semantic-freshness repro.");
                    AssertSourceFreshnessPresentationEqual(before[index], edited[index],
                        "Mutation changed display/number/font " + index + "; this repro must isolate source MathML.");
                }
                AssertEqual(before[1].SemanticSignature, edited[0].SemanticSignature,
                    "Word did not replace the first native equation with the donor's different mathematics.");
                AssertTrue(before[0].SemanticSignature != edited[0].SemanticSignature,
                    "The real Word source must change after Capture and before Apply.");
                for (var index = 1; index < selectedCount; index++)
                    AssertEqual(before[index].MathMl, edited[index].MathMl,
                        "Mutation changed a non-target selected source " + index);
                AssertEqual(textBefore, document.Content.Text,
                    "Exact-owner OLE replacement changed surrounding Word text.");
                AssertEqual(paragraphsBefore, document.Paragraphs.Count,
                    "Exact-owner OLE replacement changed paragraph structure.");
                AssertEqual(topologyBefore, LargeRedrawTableTopology(document),
                    "Exact-owner OLE replacement changed table structure.");
                AssertEqual(fixtureCount, CountMathTypeOleShapes(document),
                    "Exact-owner OLE replacement changed source count.");
            }

            // Clear fixture preparation/mutation history on the owned copy only.
            // A failing Apply must leave no undoable write, and the trace must
            // independently prove it never entered even an empty custom record.
            document.UndoClear();
            AssertSourceFreshnessUndoIdle(application, label + ": before Apply");
            File.WriteAllText(tracePath, string.Empty);
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF", "1");
            WordFormulaMutationTransaction.TraceBoundary(application, "freshness-before-apply");
            Exception? failure = null;
            WordFormulaFormatConversionResult? conversion = null;
            IReadOnlyDictionary<string, FormatScalingMetric> metrics;
            using (var operationMetrics = new WordOperationMetrics())
            {
                try { conversion = service.ApplyFormulaFormatConversionPlan(plan, prepared); }
                catch (Exception error) { failure = error; }
                metrics = SnapshotFormatScalingMetrics(operationMetrics);
            }
            WordFormulaMutationTransaction.TraceBoundary(application, "freshness-after-apply");
            AssertSourceFreshnessUndoIdle(application, label + ": after Apply");
            var trace = File.ReadAllText(tracePath);
            AssertTrue(trace.Contains("stage=freshness-before-apply")
                    && trace.Contains("stage=freshness-after-apply"),
                label + ": Undo trace unavailable; absence of transaction events would be inconclusive.");
            var transactionStarts = trace.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Count(line => line.Contains("stage=transaction-start "));
            var targetWriteCalls = metrics.Where(pair =>
                    pair.Key.StartsWith("FormatConversion.Replace", StringComparison.Ordinal)
                    || pair.Key.StartsWith("FormatConversion.OmmlHostWriter.", StringComparison.Ordinal))
                .Sum(pair => pair.Value.Calls);
            File.WriteAllText(Path.Combine(artifactRoot, label + "-evidence.json"),
                JsonSerializer.Serialize(new
                {
                    Case = label,
                    SelectedFormulaCount = selectedCount,
                    Original = before,
                    Edited = edited,
                    ErrorType = failure?.GetType().FullName,
                    Error = failure?.Message,
                    TransactionStarts = transactionStarts,
                    TargetWriteCalls = targetWriteCalls,
                    Metrics = metrics.ToDictionary(pair => pair.Key,
                        pair => new { pair.Value.Calls, pair.Value.Milliseconds }),
                }, new JsonSerializerOptions { WriteIndented = true }));

            if (changeSource)
            {
                AssertTrue(failure is InvalidOperationException && failure.InnerException is null,
                    "Changed source must yield one direct meaningful preflight exception, not rollback errors: "
                    + (failure is null ? "no exception" : FormatScalingExceptionTree(failure)));
                AssertEqual(expectedError, failure!.Message, "Changed-source diagnostic.");
                AssertTrue(conversion is null, "Changed source unexpectedly returned a conversion result.");
                AssertEqual(0, transactionStarts, "Changed source opened a new conversion Undo record.");
                AssertEqual(0L, targetWriteCalls, "Changed source entered a target writer.");
                AssertSourceFreshnessRejectedState(application, document, plan, edited,
                    paragraphsBefore, topologyBefore, textBefore, fixtureCount, "after rejected Apply");
                AssertTrue(!document.Undo(1),
                    "Rejected Apply left an undoable action after owned-copy UndoClear.");
                AssertSourceFreshnessRejectedState(application, document, plan, edited,
                    paragraphsBefore, topologyBefore, textBefore, fixtureCount, "after empty Undo");
                document.Save();
                document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                Release(document);
                document = application.Documents.Open(output,
                    ReadOnly: false, AddToRecentFiles: false, Visible: false);
                document.Activate();
                AssertSourceFreshnessRejectedState(application, document, plan, edited,
                    paragraphsBefore, topologyBefore, textBefore, fixtureCount, "saved/reopened rejection");
                Console.WriteLine("[SOURCE FRESHNESS CHANGED PASS] exact start/end/story unchanged; "
                    + "native MathML changed; zero target writers; zero transaction starts; source edit persisted.");
                return new
                {
                    Status = "Passed", Selected = selectedCount, SourceObjects = fixtureCount,
                    Error = failure.Message, TargetWriteCalls = targetWriteCalls,
                    TransactionStarts = transactionStarts, UndoAvailableAfterApply = false,
                    EditedSourcePreserved = true, SaveReopen = true, Output = output,
                };
            }

            if (failure is not null)
                throw new InvalidDataException("Unchanged-source control failed: "
                    + FormatScalingExceptionTree(failure), failure);
            AssertTrue(conversion is not null, "Unchanged-source control returned no result.");
            AssertEqual(selectedCount, conversion!.FormulaCount, "Unchanged-source conversion count.");
            AssertEqual(0, conversion.FailedFormulaCount, "Unchanged-source failed formula count.");
            AssertEqual(0, conversion.Failures.Count, "Unchanged-source reported failure details.");
            AssertEqual(1, transactionStarts, "Unchanged-source conversion must open exactly one batch Undo record.");
            AssertSourceFreshnessConvertedState(application, document, original,
                paragraphsBefore, topologyBefore, "unchanged Apply");
            AssertTrue(document.Undo(1), "Unchanged-source conversion did not support one Undo.");
            AssertSourceFreshnessRejectedState(application, document, plan, before,
                paragraphsBefore, topologyBefore, textBefore, fixtureCount, "unchanged single Undo");
            AssertTrue(document.Redo(1), "Unchanged-source conversion did not support one Redo.");
            AssertSourceFreshnessConvertedState(application, document, original,
                paragraphsBefore, topologyBefore, "unchanged single Redo");
            document.Save();
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(document);
            document = application.Documents.Open(output,
                ReadOnly: false, AddToRecentFiles: false, Visible: false);
            document.Activate();
            AssertSourceFreshnessConvertedState(application, document, original,
                paragraphsBefore, topologyBefore, "unchanged saved/reopened");
            Console.WriteLine("[SOURCE FRESHNESS CONTROL PASS] selected=3 converted; "
                + "17 unselected sources preserved; single Undo/Redo and save/reopen verified.");
            return new
            {
                Status = "Passed", Selected = selectedCount, Converted = conversion.FormulaCount,
                RemainingMathType = fixtureCount - selectedCount, TransactionStarts = transactionStarts,
                SingleUndoRedo = true, SaveReopen = true, Output = output,
            };
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", previousTrace);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF", previousPerf);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private sealed class SourceFreshnessEvidence
    {
        public int Start { get; set; }
        public int End { get; set; }
        public int Story { get; set; }
        public string MathMl { get; set; } = string.Empty;
        public string SemanticSignature { get; set; } = string.Empty;
        public string DisplayMode { get; set; } = string.Empty;
        public bool Numbered { get; set; }
        public double SemanticFontSize { get; set; }
    }

    private static SourceFreshnessEvidence[] ReadSourceFreshnessEvidence(
        Word.Application application, Word.Document document, WordFormulaFormatConversionPlan plan)
    {
        var result = new List<SourceFreshnessEvidence>();
        foreach (var target in plan.Targets)
        {
            Word.Range? range = null;
            try
            {
                range = WordFormulaOperationLocator.ResolveCapturedRange(document, target.SourceObjectId);
                using var binding = WordMathTypeHostAdapter.BindCapturedHost(application, document, range);
                var host = binding.Host;
                AssertEqual(range.Start, host.Range.Start, "Freshness evidence must bind exact start.");
                AssertEqual(range.End, host.Range.End, "Freshness evidence must bind exact end.");
                AssertEqual(range.StoryType, host.Range.StoryType, "Freshness evidence must bind exact story.");
                var mathMl = binding.MathMl
                    ?? throw new InvalidDataException("Freshly bound MathType source has no native MathML.");
                result.Add(new SourceFreshnessEvidence
                {
                    Start = host.Range.Start,
                    End = host.Range.End,
                    Story = (int)host.Range.StoryType,
                    MathMl = mathMl,
                    SemanticSignature = WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(mathMl),
                    DisplayMode = host.DisplayMode,
                    Numbered = host.Numbering.Numbered,
                    SemanticFontSize = FormulaFontSize.ResolveSemanticFontSize(host.Metadata!),
                });
            }
            finally { Release(range); }
        }
        return result.ToArray();
    }

    private static Dictionary<string, PreparedWordBulkFormula> PrepareSourceFreshnessOmml(
        WordFormulaFormatConversionPlan plan)
    {
        var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
        foreach (var target in plan.Targets)
        {
            var mathMl = target.SourceMathMl
                ?? throw new InvalidDataException("Freshness fixture source has no MathML.");
            prepared.Add(target.Id, new PreparedWordBulkFormula
            {
                Run = new WordBulkRun
                {
                    Id = target.Id, IsFormula = true, Latex = target.Latex, DisplayMode = target.DisplayMode,
                },
                Session = CreateSimpleFormatTargetSession(target, FormulaOleContract.WordOmmlMode, mathMl),
                MathMl = mathMl,
            });
        }
        return prepared;
    }

    private static void SelectSourceFreshnessScope(Word.Document document,
        IReadOnlyList<WordFormulaFormatConversionTarget> targets)
    {
        Word.Range? first = null;
        Word.Range? last = null;
        Word.Range? scope = null;
        try
        {
            first = WordFormulaOperationLocator.ResolveCapturedRange(document, targets[0].SourceObjectId);
            last = WordFormulaOperationLocator.ResolveCapturedRange(document, targets[targets.Count - 1].SourceObjectId);
            AssertEqual(first.StoryType, last.StoryType, "Freshness selection crosses Word stories.");
            scope = document.Range(first.Start, last.End);
            scope.Select();
        }
        finally { Release(scope); Release(last); Release(first); }
    }

    private static void ReplaceSourceFreshnessOle(Word.Document document,
        WordFormulaFormatConversionTarget destination, WordFormulaFormatConversionTarget donor)
    {
        Word.Range? targetRange = null;
        Word.Range? donorRange = null;
        Word.Range? formatted = null;
        try
        {
            targetRange = WordFormulaOperationLocator.ResolveCapturedRange(document, destination.SourceObjectId);
            donorRange = WordFormulaOperationLocator.ResolveCapturedRange(document, donor.SourceObjectId);
            // InlineShape.Range can span its EMBED field code, not just one
            // character. Prove the physical owner instead of assuming a length.
            AssertSourceFreshnessExactOwner(targetRange, "Destination");
            AssertSourceFreshnessExactOwner(donorRange, "Donor");
            // This mutates Word's native OLE payload, never the captured plan.
            // The subsequent exact binding/address assertions are mandatory:
            // if Word moves/expands the host, fail rather than claim freshness.
            formatted = donorRange.FormattedText;
            targetRange.FormattedText = formatted;
        }
        finally { Release(formatted); Release(donorRange); Release(targetRange); }
    }

    private static void AssertSourceFreshnessExactOwner(Word.Range captured, string label)
    {
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        Word.Range? owned = null;
        try
        {
            shapes = captured.InlineShapes;
            AssertEqual(1, shapes.Count, label + " must contain exactly one physical InlineShape owner.");
            shape = shapes[1];
            owned = shape.Range;
            AssertEqual(captured.Start, owned.Start, label + " must match the exact InlineShape start.");
            AssertEqual(captured.End, owned.End, label + " must match the exact InlineShape end.");
            AssertEqual(captured.StoryType, owned.StoryType, label + " must match the exact InlineShape story.");
        }
        finally { Release(owned); Release(shape); Release(shapes); }
    }

    private static void AssertSourceFreshnessAddressEqual(
        SourceFreshnessEvidence expected, SourceFreshnessEvidence actual, string message)
    {
        AssertEqual(expected.Start, actual.Start, message + " start");
        AssertEqual(expected.End, actual.End, message + " end");
        AssertEqual(expected.Story, actual.Story, message + " story");
    }

    private static void AssertSourceFreshnessPresentationEqual(
        SourceFreshnessEvidence expected, SourceFreshnessEvidence actual, string message)
    {
        AssertEqual(expected.DisplayMode, actual.DisplayMode, message + " display");
        AssertEqual(expected.Numbered, actual.Numbered, message + " numbered");
        AssertEqual(expected.SemanticFontSize, actual.SemanticFontSize, message + " font");
    }

    private static void AssertSourceFreshnessUndoIdle(Word.Application application, string label)
    {
        Word.UndoRecord? record = null;
        try
        {
            record = application.UndoRecord;
            AssertTrue(!record.IsRecordingCustomRecord && record.CustomRecordLevel == 0,
                label + ": a custom Undo record is active.");
        }
        finally { Release(record); }
    }

    private static void AssertSourceFreshnessRejectedState(Word.Application application, Word.Document document,
        WordFormulaFormatConversionPlan plan, SourceFreshnessEvidence[] expected,
        int paragraphs, string topology, string text, int sourceCount, string label)
    {
        AssertEqual(0, document.OMaths.Count, label + ": target OMML must remain absent.");
        AssertEqual(0, CountVisualTeXNativeOleShapes(document), label + ": unexpected target VisualTeX OLE.");
        AssertEqual(sourceCount, CountMathTypeOleShapes(document), label + ": MathType source count.");
        AssertEqual(paragraphs, document.Paragraphs.Count, label + ": paragraphs.");
        AssertEqual(topology, LargeRedrawTableTopology(document), label + ": table topology.");
        AssertEqual(text, document.Content.Text, label + ": surrounding Word text.");
        var actual = ReadSourceFreshnessEvidence(application, document, plan);
        AssertEqual(expected.Length, actual.Length, label + ": source evidence count.");
        for (var index = 0; index < expected.Length; index++)
        {
            AssertSourceFreshnessAddressEqual(expected[index], actual[index], label + " source " + index);
            AssertSourceFreshnessPresentationEqual(expected[index], actual[index], label + " source " + index);
            AssertEqual(expected[index].MathMl, actual[index].MathMl, label + ": native MathML " + index);
        }
    }

    private static void AssertSourceFreshnessConvertedState(Word.Application application, Word.Document document,
        WordFormulaFormatConversionPlan original, int paragraphs, string topology, string label)
    {
        var service = new WordFormulaService(application);
        var targets = service.CaptureFormulaFormatConversionPlan(true,
            FormulaOleContract.WordOmmlMode, FormulaOleContract.MathTypeOleMode).Targets;
        var remaining = service.CaptureFormulaFormatConversionPlan(true,
            FormulaOleContract.MathTypeOleMode, FormulaOleContract.WordOmmlMode).Targets;
        AssertEqual(3, targets.Count, label + ": converted target count.");
        AssertEqual(original.Targets.Count - 3, remaining.Count, label + ": unselected source count.");
        AssertEqual(0, CountVisualTeXNativeOleShapes(document), label + ": unexpected VisualTeX targets.");
        AssertEqual(paragraphs, document.Paragraphs.Count, label + ": paragraph boundaries.");
        AssertEqual(topology, LargeRedrawTableTopology(document), label + ": table topology.");
        for (var index = 0; index < targets.Count; index++)
        {
            var expected = original.Targets[index];
            var actual = targets[index];
            AssertEqual(expected.DisplayMode, actual.DisplayMode, label + ": target display " + index);
            AssertEqual(expected.Numbered, actual.Numbered, label + ": target numbering " + index);
            AssertTrue(Math.Abs(expected.FontSizePt - actual.FontSizePt) < 0.01,
                label + ": target font " + index);
            AssertEqual(
                WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(expected.SourceMathMl!),
                WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(actual.SourceMathMl!),
                label + ": target mathematics " + index);
        }
        for (var index = 0; index < remaining.Count; index++)
        {
            var expected = original.Targets[index + 3];
            var actual = remaining[index];
            AssertEqual(expected.SourceMathMl, actual.SourceMathMl, label + ": unselected mathematics " + index);
            AssertEqual(expected.DisplayMode, actual.DisplayMode, label + ": unselected display " + index);
            AssertEqual(expected.Numbered, actual.Numbered, label + ": unselected numbering " + index);
            AssertEqual(expected.FontSizePt, actual.FontSizePt, label + ": unselected font " + index);
        }
    }
}
