using System.Diagnostics;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordFormatConversionRenderMatrix(string artifactRoot)
    {
        var externalSeed = Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_EXTERNAL_SEED");
        var matchedSeed = Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_MATCHED_SEED");
        var numberingSetting = Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_NUMBERED");
        var numbered = numberingSetting != "0";
        var mixedDisplay = Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_MIXED_DISPLAY") == "1";
        var corpusPath = Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_CORPUS");
        var corpus = string.IsNullOrWhiteSpace(corpusPath) ? null : LoadConversionCorpus(corpusPath);
        var captureScreenshots = Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_CAPTURE") != "0";
        if (corpus is not null && !captureScreenshots)
            Console.WriteLine("[SCREEN CAPTURE NOT RUN] Semantic-only acceptance; visual acceptance remains separately blocked.");
        var count = corpus?.Count ?? int.Parse(Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_COUNT") ?? "4");
        var previewRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX", "office", "temp", "render-matrix-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(previewRoot);
        var png = Path.Combine(previewRoot, "source.png");
        var svg = Path.Combine(previewRoot, "source.svg");
        WriteAcceptancePng(png, "x+y", 260, 96);
        File.WriteAllText(svg, "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"96\" viewBox=\"0 0 260 96\"><text x=\"4\" y=\"64\" font-family=\"Times New Roman\" font-size=\"44\">x+y</text></svg>");
        var emf = OfficeOlePreview.CreateVectorEmfFromSvg(svg, 260, 96);
        var modes = new[] { (Mode: FormulaOleContract.WordOmmlMode, Name: "Omml"),
            (Mode: FormulaOleContract.NativeOleMode, Name: "VisualTeX"),
            (Mode: FormulaOleContract.MathTypeOleMode, Name: "MathType") };
        var oldTrace = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        var rows = new List<string> { "source,target,count,conversion_ms,restore_latex_ms,render_bypassed,restore_exercised" };
        var requestedSource = Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_SOURCE");
        try
        {
            foreach (var source in modes)
            {
                if (requestedSource is not null && source.Name != requestedSource) continue;
                var fixture = Path.Combine(artifactRoot, "source-" + source.Name + ".docx");
                if (!string.IsNullOrWhiteSpace(matchedSeed))
                    File.Copy(matchedSeed, fixture);
                else if (!string.IsNullOrWhiteSpace(externalSeed))
                    File.Copy(externalSeed, fixture);
                else if (source.Mode == FormulaOleContract.MathTypeOleMode)
                    File.Copy(Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_MATHTYPE_SEED")
                        ?? Path.Combine(artifactRoot, corpus is null ? "Omml-to-MathType.docx" : "VisualTeX-to-MathType.docx"), fixture);
                else if (corpus is not null && source.Mode == FormulaOleContract.NativeOleMode)
                    CreateVisualTeXConversionCorpusFixture(fixture, corpus);
                else if (corpus is not null && !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_OMML_SEED")))
                    File.Copy(Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_OMML_SEED")!, fixture);
                else if (corpus is not null)
                    CreateConversionCorpusFixture(artifactRoot, fixture, corpus);
                else
                    CreateFormatScalingSourceFixture(artifactRoot, fixture, source.Mode, count, png, emf, numbered, mixedDisplay);
                var requestedTarget = Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_TARGET");
                foreach (var target in modes.Where(mode => mode.Mode != source.Mode && (string.IsNullOrWhiteSpace(requestedTarget) || mode.Name == requestedTarget)))
                {
                    var label = source.Name + "-to-" + target.Name;
                    var output = Path.Combine(artifactRoot, label + ".docx");
                    var tracePath = Path.Combine(artifactRoot, label + ".log");
                    File.Copy(fixture, output);
                    WordFormulaFormatConversionPlan? sourceSnapshot = null;
                    string? sourceTopology = null;
                    var sourceParagraphs = 0;
                    var conversionMs = 0L;
                    Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
                    using (var host = new WordPerformanceHost(output))
                    {
                        var service = new WordFormulaService(host.Application);
                        var before = service.CaptureFormulaFormatConversionPlan(true, source.Mode, target.Mode);
                        sourceSnapshot = before;
                        if (!string.IsNullOrWhiteSpace(externalSeed))
                        {
                            sourceTopology = LargeRedrawTableTopology(host.Document);
                            sourceParagraphs = host.Document.Paragraphs.Count;
                        }
                        AssertEqual(count, before.Targets.Count, label + " source count");
                        if (numberingSetting is not null)
                        {
                            AssertEqual(numbered ? before.Targets.Count(item => item.DisplayMode == "block") : 0,
                                before.Targets.Count(item => item.Numbered), label + " requested source numbering");
                            if (!numbered) AssertNoConversionNumberFields(host.Document, label + " source");
                            if (mixedDisplay) AssertEqual(count / 2, before.Targets.Count(item => item.DisplayMode == "inline"), label + " source inline count");
                        }
                        if (corpus is null && string.IsNullOrWhiteSpace(externalSeed))
                            for (var index = 0; index < count; index++)
                                AssertRenderMatrixFormula(before.Targets[index], index + 1, label + " source");
                        if (corpus is not null && captureScreenshots)
                            CaptureConversionCorpusScreenshots(host, before.Targets, artifactRoot, label + "-source");
                        using var operationMetrics = new WordOperationMetrics();
                        var watch = Stopwatch.StartNew();
                        typeof(ThisAddIn).GetMethod("OnConvert" + source.Name + "To" + target.Name + "Document")!
                            .Invoke(host.AddIn, new object[] { new object() });
                        WaitForAddInIdleFast(host.AddIn, TimeSpan.FromMinutes(20));
                        watch.Stop();
                        conversionMs = watch.ElapsedMilliseconds;
                        File.WriteAllText(Path.Combine(artifactRoot, label + "-timing.json"), System.Text.Json.JsonSerializer.Serialize(new { Source = source.Name, Target = target.Name, Count = count, ConversionMs = conversionMs }));
                        Console.WriteLine($"[CONVERSION WRITTEN] {label} count={count} conversionMs={conversionMs}");
                        File.WriteAllText(Path.Combine(artifactRoot, label + "-metrics.json"), System.Text.Json.JsonSerializer.Serialize(operationMetrics.Entries, new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
                        host.Document.Save();
                        var trace = File.ReadAllText(tracePath);
                        AssertTrue(trace.Contains("format-conversion-complete") && trace.Contains($"converted={count} failed=0"), label + " did not complete: " + trace);
                        Console.WriteLine("[CONVERSION VALIDATING] " + label);
                        var after = service.CaptureFormulaFormatConversionPlan(true, target.Mode, source.Mode);
                        AssertEqual(count, after.Targets.Count, label + " target count");
                        AssertTrue(after.Targets.Select(item => item.Numbered).SequenceEqual(before.Targets.Select(item => item.Numbered)), label + " lost numbering");
                        AssertTrue(after.Targets.Select(item => item.DisplayMode).SequenceEqual(before.Targets.Select(item => item.DisplayMode)), label + " changed inline/display shape");
                        if (numberingSetting == "0") AssertNoConversionNumberFields(host.Document, label + " target");
                        AssertConversionNumberSequence(host, service, target.Mode, before.Targets.Count(item => item.Numbered), label);
                        if (corpus is not null && captureScreenshots)
                            CaptureConversionCorpusScreenshots(host, after.Targets, artifactRoot, label);
                        for (var index = 0; index < count; index++)
                            if (!string.IsNullOrWhiteSpace(externalSeed)) AssertExternalConversionFormula(before.Targets[index], after.Targets[index], artifactRoot, label, index);
                            else if (corpus is null) AssertRenderMatrixFormula(after.Targets[index], index + 1, label + " target");
                            else AssertConversionCorpusFormula(corpus[index], before.Targets[index], after.Targets[index], artifactRoot, label, index);
                        if (!string.IsNullOrWhiteSpace(externalSeed))
                            AssertExternalConversionStructure(host, service, before, target.Mode, sourceTopology!, sourceParagraphs, label);
                        AssertEqual(0, service.CountFormulaObjectsForLatex(true, source.Mode), label + " retained source hosts");
                        AssertTrue(host.Document.Content.Text.StartsWith(string.IsNullOrWhiteSpace(externalSeed) ? "VisualTeX format conversion scaling source" : "LARGE-REDRAW-BEGIN"), label + " damaged surrounding prose");
                        host.Document.Save();

                        if (Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_HISTORY") == "1")
                        {
                            AssertTrue(host.Document.Undo(1), label + " one Undo");
                            AssertEqual(count, service.CountFormulaObjectsForLatex(true, source.Mode), label + " Undo source count");
                            AssertEqual(0, service.CountFormulaObjectsForLatex(true, target.Mode), label + " Undo target removed");
                            AssertTrue(host.Document.Redo(1), label + " one Redo");
                            AssertEqual(count, service.CountFormulaObjectsForLatex(true, target.Mode), label + " Redo target count");
                        }
                        var restoreExercised = Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_RESTORE") == "1";
                        watch.Reset();
                        if (restoreExercised)
                        {
                        watch.Start();
                        var reverseName = target.Mode == FormulaOleContract.NativeOleMode ? "Ole" : target.Name;
                        typeof(ThisAddIn).GetMethod("OnRedrawDocument" + reverseName + "ToLatex")!
                            .Invoke(host.AddIn, new object[] { new object() });
                        WaitForAddInIdleFast(host.AddIn, TimeSpan.FromMinutes(20));
                        watch.Stop();
                        AssertEqual(0, service.CountFormulaObjectsForLatex(true, target.Mode), label + " restore-to-LaTeX retained hosts");
                        AssertEqual(before.Targets.Count(item => item.DisplayMode == "block") * 2, System.Text.RegularExpressions.Regex.Matches(host.Document.Content.Text, @"\$\$").Count, label + " restore-to-LaTeX lost display sources");
                        }
                        rows.Add($"{source.Name},{target.Name},{count},{conversionMs},{watch.ElapsedMilliseconds},{trace.Contains("format-conversion-render-bypass")},{restoreExercised}");
                    }
                    Console.WriteLine("[CONVERSION REOPEN] " + label);
                    using (var reopened = new WordPerformanceHost(output))
                    {
                        var service = new WordFormulaService(reopened.Application);
                        AssertEqual(count, service.CountFormulaObjectsForLatex(true, target.Mode), label + " save/reopen count");
                        if (numberingSetting == "0") AssertNoConversionNumberFields(reopened.Document, label + " reopened");
                        AssertConversionNumberSequence(reopened, service, target.Mode,
                            sourceSnapshot!.Targets.Count(item => item.Numbered), label + " reopened");
                        if (!string.IsNullOrWhiteSpace(externalSeed))
                        {
                            AssertExternalConversionStructure(reopened, service, sourceSnapshot!, target.Mode, sourceTopology!, sourceParagraphs, label + " reopened");
                            var targets = service.CaptureFormulaFormatConversionPlan(true, target.Mode, source.Mode).Targets;
                            for (var index = 0; index < count; index++) AssertExternalConversionFormula(sourceSnapshot!.Targets[index], targets[index], artifactRoot, label + " reopened", index);
                        }
                        else if (corpus is not null)
                        {
                            var targets = service.CaptureFormulaFormatConversionPlan(true, target.Mode, source.Mode).Targets;
                            AssertEqual(count, targets.Count, label + " reopened semantic count");
                            for (var index = 0; index < count; index++)
                                AssertConversionCorpusFormula(corpus[index], sourceSnapshot!.Targets[index], targets[index],
                                    artifactRoot, label + " reopened", index);
                        }
                    }
                    File.WriteAllLines(Path.Combine(artifactRoot, "conversion-render-matrix.csv"), rows);
                    if (count == 20 && numberingSetting is "0" or "1")
                    {
                        var limitMs = numbered ? 15000 : 12000;
                        var performancePassed = conversionMs <= limitMs;
                        File.WriteAllText(Path.Combine(artifactRoot, label + "-gate.json"),
                            System.Text.Json.JsonSerializer.Serialize(new { Source = source.Name, Target = target.Name,
                                Count = count, Numbered = numbered, ConversionMs = conversionMs, LimitMs = limitMs,
                                CorrectnessPassed = true, PerformancePassed = performancePassed }));
                        Console.WriteLine("[PERFORMANCE GATE] " + (performancePassed ? "PASS " : "FAIL ")
                            + label + " " + conversionMs + " <= " + limitMs);
                    }
                    Console.WriteLine("[CONVERSION RENDER MATRIX] CORRECTNESS PASS " + rows.Last());
                }
            }
        }
        finally { Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", oldTrace); }
    }

    private static void AssertNoConversionNumberFields(Microsoft.Office.Interop.Word.Document document, string label)
    {
        using var fields = WordNumberingFieldCollection.Capture(document);
        for (var index = 1; index <= fields.Count; index++)
        {
            using var entry = fields.TakeEntry(index);
            var instruction = entry.Instruction.Trim();
            AssertTrue(entry.Type != Microsoft.Office.Interop.Word.WdFieldType.wdFieldSequence
                && entry.Type != Microsoft.Office.Interop.Word.WdFieldType.wdFieldStyleRef
                && instruction.IndexOf("PlaceRef", StringComparison.OrdinalIgnoreCase) < 0,
                label + " has numbering field: " + instruction);
        }
    }

    private static string RenderMatrixMathTokens(string mathMl) => string.Join("|",
        XDocument.Parse(mathMl).Descendants()
            .Where(node => node.Name.LocalName is "mi" or "mn" or "mo")
            .Select(node => node.Name.LocalName + ":" + node.Value));

    private static void AssertRenderMatrixFormula(WordFormulaFormatConversionTarget formula, int index, string label)
    {
        if (formula.SourceMathMl is not null)
            AssertEqual(RenderMatrixMathTokens(FormatScalingMathMl(index)), RenderMatrixMathTokens(formula.SourceMathMl), label + " mathematical tokens");
        else
            AssertEqual($"x_{index}+y_{index}={index}", System.Text.RegularExpressions.Regex.Replace(formula.Latex, @"[\s{}]", ""), label + " LaTeX metadata");
    }
}
