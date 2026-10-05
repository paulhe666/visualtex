using System.Diagnostics;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void RunOwnedSixWayPerformance(string root)
    {
        var tracePath = Path.Combine(root, "conversion-trace.log");
        Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
        var corpus = LoadConversionCorpus(Environment.GetEnvironmentVariable("VISUALTEX_PERF_CORPUS")
            ?? throw new InvalidOperationException("Explicit performance corpus required."));
        var source = Environment.GetEnvironmentVariable("VISUALTEX_PERF_SOURCE") ?? "VisualTeX";
        var target = Environment.GetEnvironmentVariable("VISUALTEX_PERF_TARGET") ?? "Omml";
        var seed = Environment.GetEnvironmentVariable("VISUALTEX_PERF_SEED");
        var output = Path.Combine(root, source + "-to-" + target + ".docx");
        if (string.IsNullOrWhiteSpace(seed))
        {
            var seedOutput = Path.Combine(root, "source-" + source + ".docx");
            var watch = Stopwatch.StartNew();
            if (source == "VisualTeX") CreateVisualTeXConversionCorpusFixture(seedOutput, corpus);
            else if (source == "Omml") CreateConversionCorpusFixture(root, seedOutput, corpus);
            else throw new InvalidOperationException("MathType source must be an explicitly saved, inventoried fixture.");
            File.WriteAllText(Path.Combine(root, "seed-result.json"), JsonSerializer.Serialize(new {
                source, count = corpus.Count, seedOutput, milliseconds = watch.Elapsed.TotalMilliseconds }));
            return;
        }
        string Mode(string name) => name switch {
            "VisualTeX" => FormulaOleContract.NativeOleMode,
            "MathType" => FormulaOleContract.MathTypeOleMode,
            "Omml" => FormulaOleContract.WordOmmlMode,
            _ => throw new InvalidOperationException("Unknown formula format: " + name),
        };
        var sourceMode = Mode(source); var targetMode = Mode(target);
        if (sourceMode == targetMode) throw new InvalidOperationException("Conversion requires different formats.");
        File.Copy(seed, output);
        var failures = new List<object>();
        var sourceDifferences = new List<object>();
        var timings = new Dictionary<string, double>();
        var comparisons = new List<object>();
        var converted = 0; var complete = false;
        WordFormulaFormatConversionPlan? before = null;
        void Validate(WordFormulaFormatConversionPlan after, string phase)
        {
            if (after.Targets.Count != corpus.Count)
                failures.Add(new { phase, kind="count", expected=corpus.Count, actual=after.Targets.Count });
            for (var i = 0; i < Math.Min(before!.Targets.Count, after.Targets.Count); i++)
            {
                try {
                    AssertEqual(before.Targets[i].DisplayMode, after.Targets[i].DisplayMode, "display mode " + i);
                    AssertEqual(before.Targets[i].Numbered, after.Targets[i].Numbered, "numbered " + i);
                    AssertConversionCorpusFormula(corpus[i], before.Targets[i], after.Targets[i], root, phase, i);
                    comparisons.Add(new { phase, index=i+1, passed=true });
                } catch (Exception error) {
                    failures.Add(new { phase, index=i+1, error=error.Message });
                    comparisons.Add(new { phase, index=i+1, passed=false });
                }
            }
        }
        try
        {
            var watch = Stopwatch.StartNew();
            using (var host = new WordPerformanceHost(output))
            {
                timings["hostOpenAndAddInMs"] = watch.Elapsed.TotalMilliseconds;
                var service = new WordFormulaService(host.Application);
                watch.Restart();
                before = service.CaptureFormulaFormatConversionPlan(true, sourceMode, targetMode);
                timings["sourceInspectionOutsideConversionMs"] = watch.Elapsed.TotalMilliseconds;
                AssertEqual(corpus.Count, before.Targets.Count, "Source corpus count");
                for (var i=0; i<corpus.Count; i++)
                {
                    var mathMl=before.Targets[i].SourceMathMl;
                    if (mathMl is null) continue;
                    var expected=WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(corpus[i].MathMl);
                    var actual=WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(mathMl);
                    if (expected!=actual) sourceDifferences.Add(new { index=i+1, expected, actual });
                }
                using (var metrics = new WordOperationMetrics())
                {
                    watch.Restart();
                    typeof(ThisAddIn).GetMethod("OnConvert" + source + "To" + target + "Document")!
                        .Invoke(host.AddIn, new object[] { new object() });
                    WaitForAddInIdleFast(host.AddIn, TimeSpan.FromMinutes(5));
                    timings["conversionMs"] = watch.Elapsed.TotalMilliseconds;
                    File.WriteAllText(Path.Combine(root,"operation-metrics.json"), JsonSerializer.Serialize(metrics.Entries,
                        new JsonSerializerOptions {WriteIndented=true}));
                }
                var trace=File.ReadAllText(tracePath);
                complete=trace.Contains("format-conversion-complete") && trace.Contains($"converted={corpus.Count} failed=0");
                if (!complete) failures.Add(new { phase="conversion", error="Production conversion did not complete all requested formulas; see trace." });
                watch.Restart(); host.Document.Save(); timings["saveMs"]=watch.Elapsed.TotalMilliseconds;
                watch.Restart();
                var after=service.CaptureFormulaFormatConversionPlan(true,targetMode,sourceMode);
                converted=after.Targets.Count;
                Validate(after,"converted");
                var remaining=service.CountFormulaObjectsForLatex(true,sourceMode);
                if(remaining!=0) failures.Add(new {phase="conversion", remainingSource=remaining});
                AssertTrue(host.Document.Content.Text.StartsWith("VisualTeX format conversion scaling source"),"Surrounding prose");
                timings["validationOutsideConversionMs"]=watch.Elapsed.TotalMilliseconds;
            }
            watch.Restart();
            using(var reopened = new WordPerformanceHost(output))
            {
                var after=new WordFormulaService(reopened.Application).CaptureFormulaFormatConversionPlan(true,targetMode,sourceMode);
                Validate(after,"reopened");
            }
            timings["reopenAndValidationMs"]=watch.Elapsed.TotalMilliseconds;
        }
        catch(Exception error) { failures.Add(new { phase="runner",error=error.ToString() }); }
        finally
        {
            File.WriteAllText(Path.Combine(root,"performance-result.json"),JsonSerializer.Serialize(new {
                source,target,requested=corpus.Count,converted,conversionCompleted=complete,
                fidelityPassed=failures.Count==0,sourceDifferences,failures,comparisons,timings,
                scope="Conversion wall time after explicit source inspection; validation, save, reopen and startup excluded. Source differences cannot be recovered by a round trip. No pixel acceptance.",
                wordDll=typeof(ThisAddIn).Assembly.Location,
                wordMvid=typeof(ThisAddIn).Assembly.ManifestModule.ModuleVersionId
            },new JsonSerializerOptions{WriteIndented=true}));
        }
        Console.WriteLine($"[SIX-WAY PERFORMANCE] {source}->{target} requested={corpus.Count} converted={converted} complete={complete} fidelityFailures={failures.Count}");
        if(failures.Count!=0) throw new InvalidDataException("Performance recorded with conversion/fidelity failures; see performance-result.json.");
    }
}
