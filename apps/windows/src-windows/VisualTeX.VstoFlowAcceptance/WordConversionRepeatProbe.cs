using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")]
    private static extern uint GetConversionProbeWindowThreadProcessId(IntPtr window, out uint processId);

    private static void RunWordConversionRepeatProbe(string artifactRoot)
    {
        var source = Environment.GetEnvironmentVariable("VISUALTEX_REPEAT_SOURCE")
            ?? throw new InvalidOperationException("Repeat probe requires an owned MathType source fixture.");
        var owned = Path.Combine(artifactRoot, "owned-repeat.docx");
        File.Copy(source, owned);
        var trace = Path.Combine(artifactRoot, "repeat.log");
        var previousTrace = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", trace);
        var results = new List<object>();
        var persistedExpectations = new List<FormulaMetadata[]>();
        try
        {
            using var host = new WordPerformanceHost(owned);
            var window = host.Document.ActiveWindow;
            int wordWindow;
            try
            {
                if (Environment.GetEnvironmentVariable("VISUALTEX_REPEAT_VISIBLE") == "1")
                {
                    host.Application.Visible = true;
                    window.Visible = true;
                    host.Document.Activate();
                    host.Application.Activate();
                }
                wordWindow = window.Hwnd;
            }
            finally { Release(window); }
            GetConversionProbeWindowThreadProcessId(new IntPtr(wordWindow), out var wordPid);
            var service = new WordFormulaService(host.Application);
            using var wordProcess = Process.GetProcessById((int)wordPid);
            using var harnessProcess = Process.GetCurrentProcess();
            for (var round = 1; round <= 2; round++)
            {
                var before = service.CaptureFormulaFormatConversionPlan(true,
                    FormulaOleContract.MathTypeOleMode, FormulaOleContract.NativeOleMode);
                AssertEqual(20, before.Targets.Count, "Repeat source count");
                AssertEqual(10, before.Targets.Count(item => item.DisplayMode == "inline"), "Repeat inline count");
                var numberCount = before.Targets.Count(item => item.Numbered);
                if (numberCount == 0) AssertNoConversionNumberFields(host.Document, "Repeat source");
                using var metrics = new WordOperationMetrics();
                wordProcess.Refresh(); harnessProcess.Refresh();
                var wordCpuBefore = wordProcess.TotalProcessorTime;
                var harnessCpuBefore = harnessProcess.TotalProcessorTime;
                var retriesBefore = System.Threading.Interlocked.Read(ref _officeComRetryCount);
                var retryDelayBefore = System.Threading.Interlocked.Read(ref _officeComRetryRequestedMilliseconds);
                var watch = Stopwatch.StartNew();
                host.AddIn.OnConvertMathTypeToVisualTeXDocument(new object());
                WaitForAddInIdleFast(host.AddIn, TimeSpan.FromMinutes(3));
                watch.Stop();
                wordProcess.Refresh(); harnessProcess.Refresh();
                var wordCpuMs = (wordProcess.TotalProcessorTime - wordCpuBefore).TotalMilliseconds;
                var harnessCpuMs = (harnessProcess.TotalProcessorTime - harnessCpuBefore).TotalMilliseconds;
                var retryCount = System.Threading.Interlocked.Read(ref _officeComRetryCount) - retriesBefore;
                var retryRequestedMs = System.Threading.Interlocked.Read(ref _officeComRetryRequestedMilliseconds) - retryDelayBefore;
                File.WriteAllText(Path.Combine(artifactRoot, "round-" + round + "-metrics.json"),
                    JsonSerializer.Serialize(metrics.Entries, new JsonSerializerOptions { WriteIndented = true }));
                var after = service.CaptureFormulaFormatConversionPlan(true,
                    FormulaOleContract.NativeOleMode, FormulaOleContract.MathTypeOleMode);
                AssertEqual(20, after.Targets.Count, "Repeat target count");
                persistedExpectations.Add(after.Targets.Select(item => item.Metadata).ToArray());
                AssertEqual(0, service.CountFormulaObjectsForLatex(true, FormulaOleContract.MathTypeOleMode), "Repeat retired source count");
                AssertTrue(before.Targets.Select(item => item.DisplayMode).SequenceEqual(after.Targets.Select(item => item.DisplayMode)), "Repeat display modes");
                AssertTrue(before.Targets.Select(item => item.Numbered).SequenceEqual(after.Targets.Select(item => item.Numbered)), "Repeat numbered flags");
                for (var index = 0; index < 20; index++) AssertRenderMatrixFormula(after.Targets[index], index + 1, "Repeat target");
                AssertConversionNumberSequence(host, service, FormulaOleContract.NativeOleMode, numberCount, "Repeat target");
                if (numberCount == 0) AssertNoConversionNumberFields(host.Document, "Repeat target");
                host.Document.Save();
                File.Copy(owned, Path.Combine(artifactRoot, "round-" + round + ".docx"));
                AssertTrue(host.Document.Undo(1), "Repeat one Undo");
                AssertEqual(20, service.CountFormulaObjectsForLatex(true, FormulaOleContract.MathTypeOleMode), "Repeat Undo restored source");
                AssertEqual(0, service.CountFormulaObjectsForLatex(true, FormulaOleContract.NativeOleMode), "Repeat Undo removed target");
                var limit = numberCount == 0 ? 12000 : 15000;
                results.Add(new { Round = round, WordProcessId = wordPid, WordWindow = wordWindow,
                    Visible = host.Application.Visible, ScreenUpdating = host.Application.ScreenUpdating,
                    Count = 20, NumberedCount = numberCount, ConversionMs = watch.ElapsedMilliseconds,
                    WordCpuMs = wordCpuMs, HarnessCpuMs = harnessCpuMs,
                    ComRetryCount = retryCount, ComRetryRequestedMilliseconds = retryRequestedMs,
                    LimitMs = limit, PerformancePassed = watch.ElapsedMilliseconds <= limit,
                    CorrectnessPassed = true, RestoredBySingleUndo = true });
                File.WriteAllText(Path.Combine(artifactRoot, "results.json"), JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));
                Console.WriteLine("[REPEAT PROBE] round=" + round + " ms=" + watch.ElapsedMilliseconds + " wordPid=" + wordPid);
            }
            host.Dispose();
            for (var round = 1; round <= 2; round++)
            {
                using var reopened = new WordPerformanceHost(Path.Combine(artifactRoot, "round-" + round + ".docx"));
                var shapes = reopened.Document.InlineShapes;
                try
                {
                    AssertEqual(20, shapes.Count, "Persisted OLE count");
                    for (var index = 1; index <= shapes.Count; index++)
                    {
                        var shape = shapes[index];
                        try
                        {
                            var actual = WordFormulaMetadataReader.TryReadEmbeddedNativeOle(shape)
                                ?? throw new InvalidDataException("Saved OLE has no authoritative native payload.");
                            var expected = persistedExpectations[round - 1][index - 1];
                            AssertEqual(expected.FormulaId, actual.FormulaId, "Persisted formula identity");
                            AssertEqual(expected.Latex, actual.Latex, "Persisted LaTeX");
                            AssertEqual(expected.DisplayMode, actual.DisplayMode, "Persisted display mode");
                            AssertEqual(expected.Numbered, actual.Numbered, "Persisted numbering intent");
                            AssertEqual(expected.FontSizePt, actual.FontSizePt, "Persisted semantic font size");
                        }
                        finally { Release(shape); }
                    }
                }
                finally { Release(shapes); }
            }
            File.WriteAllText(Path.Combine(artifactRoot, "persistence.json"),
                JsonSerializer.Serialize(new { Passed = true, Rounds = 2, NativePayloads = 40 }));
        }
        finally { Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", previousTrace); }
    }
}
