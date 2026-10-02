using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordLargeRedrawLayout(string artifactRoot)
    {
        var count = int.Parse(Environment.GetEnvironmentVariable("VISUALTEX_LARGE_REDRAW_COUNT") ?? "200");
        var layout = Environment.GetEnvironmentVariable("VISUALTEX_LARGE_REDRAW_LAYOUT") ?? "mixed";
        var mode = Environment.GetEnvironmentVariable("VISUALTEX_LARGE_REDRAW_MODE") ?? FormulaOleContract.NativeOleMode;
        var journal = Environment.GetEnvironmentVariable("VISUALTEX_LARGE_RENDER_JOURNAL")!;
        var label = mode + "-" + layout + "-" + count;
        var logPath = Path.Combine(artifactRoot, "redraw.log");
        var oldLog = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG");
        var oldNumber = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_NUMBER_DISPLAY_FORMULAS");
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG", logPath);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_NUMBER_DISPLAY_FORMULAS", "1");
        var sourceLatex = new List<string>();
        string topology;
        int sourceParagraphs, sourcePages;
        long wallMs, validationMs;
        var output = Path.Combine(artifactRoot, "result.docx");
        try
        {
            using (var host = new WordPerformanceHost(null))
            {
                host.Document.Content.Font.Name = "宋体";
                host.Document.Content.Font.Size = 10.5f;
                WordEquationNumbering.SetEquationNumberFormatPreference(host.Document, EquationNumberFormat.ContinuousId);
                host.Document.Content.Text = "LARGE-REDRAW-BEGIN\r";

                string FormulaText()
                {
                    var index = sourceLatex.Count + 1;
                    var latex = LargeRedrawLatex(index);
                    sourceLatex.Add(latex);
                    return index % 2 == 1
                        ? $"BEGIN[{index:0000}] 前文 ${latex}$ 后文 END[{index:0000}]\r"
                        : $"BEGIN[{index:0000}] 前文\r$${latex}$$\r后文 END[{index:0000}]\r";
                }
                void AddBody(int formulas)
                {
                    var text = new StringBuilder();
                    for (var index = 0; index < formulas; index++) text.Append(FormulaText());
                    var tail = host.Document.Range(host.Document.Content.End - 1, host.Document.Content.End - 1);
                    try { tail.InsertBefore(text.ToString()); } finally { Release(tail); }
                }
                void AddTable(int formulas, int tableIndex)
                {
                    const int perCell = 5;
                    var cells = (formulas + perCell - 1) / perCell;
                    var rows = 1 + (cells + 1) / 2;
                    var anchor = host.Document.Range(host.Document.Content.End - 1, host.Document.Content.End - 1);
                    Word.Table? table = null;
                    try
                    {
                        table = host.Document.Tables.Add(anchor, rows, 2);
                        table.AllowAutoFit = false;
                        table.Rows.AllowBreakAcrossPages = -1;
                        var left = table.Cell(1, 1); var right = table.Cell(1, 2);
                        try { left.Merge(right); } finally { Release(left); Release(right); }
                        var header = table.Cell(1, 1); var headerRange = header.Range.Duplicate;
                        try { headerRange.End--; headerRange.Text = $"MERGED TABLE {tableIndex} HEADER\r"; } finally { Release(headerRange); Release(header); }
                        table.Rows[1].HeadingFormat = -1;
                        var remaining = formulas;
                        for (var row = 2; row <= rows; row++)
                        for (var column = 1; column <= 2; column++)
                        {
                            var cell = table.Cell(row, column); var range = cell.Range.Duplicate;
                            try
                            {
                                range.End--;
                                var text = new StringBuilder($"CELL[{tableIndex},{row},{column}]\r");
                                for (var index = 0; index < perCell && remaining > 0; index++, remaining--) text.Append(FormulaText());
                                text.Append($"CELL-END[{tableIndex},{row},{column}]\r");
                                range.Text = text.ToString();
                            }
                            finally { Release(range); Release(cell); }
                        }
                    }
                    finally { Release(table); Release(anchor); }
                    var tail = host.Document.Range(host.Document.Content.End - 1, host.Document.Content.End - 1);
                    try { tail.InsertBefore("AFTER-TABLE-" + tableIndex + "\r"); } finally { Release(tail); }
                }

                if (layout == "body") AddBody(count);
                else if (layout == "table") AddTable(count, 1);
                else
                {
                    AddBody(count / 4); AddTable(count / 4, 1);
                    AddBody(count / 4); AddTable(count - 3 * (count / 4), 2);
                }
                var end = host.Document.Range(host.Document.Content.End - 1, host.Document.Content.End - 1);
                try { end.InsertBefore("LARGE-REDRAW-END\r"); } finally { Release(end); }
                AssertEqual(count, sourceLatex.Count, "Large fixture formula count");
                sourceParagraphs = host.Document.Paragraphs.Count;
                topology = LargeRedrawTableTopology(host.Document);
                sourcePages = host.Document.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                File.WriteAllText(Path.Combine(artifactRoot, "expected-latex.json"), JsonSerializer.Serialize(sourceLatex));
                host.Save(Path.Combine(artifactRoot, "source.docx"));
                var service = new WordFormulaService(host.Application);
                var captured = service.CaptureLatexRedrawPlan(true);
                AssertEqual(count, captured.Targets.Count, "Large fixture source capture");
                AssertTrue(sourceLatex.SequenceEqual(captured.Targets.Select(target => target.Latex)), "Large fixture source order");
                Console.WriteLine($"[LARGE REDRAW START] {label} pages={sourcePages} paragraphs={sourceParagraphs}");

                var watch = Stopwatch.StartNew();
                using var operationMetrics = new WordOperationMetrics();
                if (mode == FormulaOleContract.WordOmmlMode) host.AddIn.OnRedrawDocumentToOmml(new object());
                else if (mode == FormulaOleContract.MathTypeOleMode) host.AddIn.OnRedrawDocumentToMathType(new object());
                else host.AddIn.OnRedrawDocumentToOle(new object());
                var log = WaitForLatexRedraw(logPath, TimeSpan.FromMinutes(20));
                WaitForAddInIdleFast(host.AddIn, TimeSpan.FromMinutes(5));
                watch.Stop(); wallMs = watch.ElapsedMilliseconds;
                File.WriteAllText(Path.Combine(artifactRoot, "operation-metrics.json"), JsonSerializer.Serialize(
                    operationMetrics.Entries, new JsonSerializerOptions { WriteIndented = true }));
                AssertTrue(log.Contains($"redraw-complete formulas={count} unique={count} skipped=0"), "Large redraw incomplete: " + log);
                Console.WriteLine($"[LARGE REDRAW WRITTEN] {label} wallMs={wallMs}");
                watch.Restart();
                ValidateLargeRedraw(host, mode, sourceLatex, topology, sourceParagraphs, journal, artifactRoot);
                watch.Stop(); validationMs = watch.ElapsedMilliseconds;
                host.Save(output);
            }
            var reopen = Stopwatch.StartNew();
            using (var host = new WordPerformanceHost(output))
                ValidateLargeRedraw(host, mode, sourceLatex, topology, sourceParagraphs, journal, artifactRoot);
            reopen.Stop();
            var phaseLog = File.ReadAllText(logPath);
            long Phase(string name) => long.Parse(Regex.Match(phaseLog, @"\b" + name + @"=(\d+)").Groups[1].Value);
            var result = new { Status = "Passed", Mode = mode, Layout = layout, Count = count, Inline = (count + 1) / 2,
                NumberedDisplay = count / 2, Unique = count, SourcePages = sourcePages, SourceParagraphs = sourceParagraphs,
                WallMs = wallMs, CaptureMs = Phase("captureMs"), RenderMs = Phase("renderBatchMs"), WordApplyMs = Phase("wordApplyMs"),
                ValidationMs = validationMs, ReopenValidationMs = reopen.ElapsedMilliseconds, TableTopology = topology,
                ProsePreserved = true, SemanticReadback = true, SaveReopen = true };
            File.WriteAllText(Path.Combine(artifactRoot, "result.json"), JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine("[LARGE REDRAW PASS] " + JsonSerializer.Serialize(result));
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG", oldLog);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_NUMBER_DISPLAY_FORMULAS", oldNumber);
        }
    }

    private static void RunLargeRedrawNumberingProfile(string artifactRoot)
    {
        var sourceRoot = Environment.GetEnvironmentVariable("VISUALTEX_LARGE_PROFILE_SOURCE")!;
        var results = new List<object>();
        foreach (var mode in new[] { FormulaOleContract.NativeOleMode, FormulaOleContract.WordOmmlMode, FormulaOleContract.MathTypeOleMode })
        {
            using var host = new WordPerformanceHost(Path.Combine(sourceRoot, "cp146-large-body-" + mode + "-200", "result.docx"));
            void Measure(string operation, Action action)
            {
                using var metrics = new WordOperationMetrics();
                var watch = Stopwatch.StartNew();
                WordFormulaMutationTransaction.Execute(host.Application, host.Document, "Redraw numbering profile", () => { action(); return true; });
                watch.Stop();
                var result = new { Mode = mode, Operation = operation, WallMs = watch.ElapsedMilliseconds, Metrics = metrics.Entries };
                results.Add(result);
                File.WriteAllText(Path.Combine(artifactRoot, "numbering-profile.json"), JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));
                Console.WriteLine("[NUMBERING PROFILE] " + JsonSerializer.Serialize(result));
            }
            if (mode == FormulaOleContract.MathTypeOleMode)
                Measure("current", () => MathTypeEquationNumbering.UpdateEquationNumbers(host.Document));
            else
            {
                Measure("current", () => WordFormulaNumberingKernel.RefreshCanonicalNumbers(host.Document));
                Measure("target-family", () => WordFormulaNumberingKernel.RefreshCanonicalNumbers(host.Document,
                    kindFilter: mode == FormulaOleContract.NativeOleMode ? WordFormulaHostKind.VisualTeX : WordFormulaHostKind.Omml));
            }
        }
    }

    private static string LargeRedrawLatex(int n) => ((n - 1) % 6) switch
    {
        0 => $"x_{{{n}}}+y_{{{n}}}={n}",
        1 => $"\\frac{{x_{{{n}}}+1}}{{y_{{{n}}}+2}}={n}",
        2 => $"\\sqrt{{x_{{{n}}}^2+y_{{{n}}}^2}}={n}",
        3 => $"\\sum_{{k=1}}^{{{n}}}k^2=S_{{{n}}}",
        4 => $"\\int_{{0}}^{{1}}x^{{{n}}}\\,\\mathrm{{d}}x=I_{{{n}}}",
        _ => $"\\begin{{pmatrix}}{n}&1\\\\0&{n}\\end{{pmatrix}}",
    };

    private static string LargeRedrawTableTopology(Word.Document document)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var xml = XDocument.Parse(document.Content.WordOpenXML);
        return string.Join(";", xml.Descendants(w + "tbl").Select(table =>
            string.Join("/", table.Elements(w + "tr").Select(row => string.Join(",", row.Elements(w + "tc").Select(cell =>
                (string?)cell.Element(w + "tcPr")?.Element(w + "gridSpan")?.Attribute(w + "val") ?? "1"))))));
    }

    private static void ValidateLargeRedraw(WordPerformanceHost host, string mode, List<string> latex, string topology,
        int paragraphs, string journal, string artifactRoot)
    {
        var document = host.Document;
        AssertEqual(topology, LargeRedrawTableTopology(document), "Large redraw changed table rows/cells/merged spans");
        AssertEqual(paragraphs, document.Paragraphs.Count, "Large redraw changed paragraph boundaries");
        var text = document.Content.Text;
        AssertTrue(text.Contains("LARGE-REDRAW-BEGIN") && text.Contains("LARGE-REDRAW-END"), "Large redraw lost document bounds");
        AssertTrue(!text.Contains("$"), "Large redraw left LaTeX delimiters");
        for (var index = 1; index <= latex.Count; index++)
        {
            AssertEqual(1, Regex.Matches(text, Regex.Escape($"BEGIN[{index:0000}] 前文")).Count, "Large redraw damaged leading prose " + index);
            AssertEqual(1, Regex.Matches(text, Regex.Escape($"后文 END[{index:0000}]")).Count, "Large redraw damaged trailing prose " + index);
        }
        var service = new WordFormulaService(host.Application);
        var readback = service.CaptureFormulaFormatConversionPlan(true, mode,
            mode == FormulaOleContract.WordOmmlMode ? FormulaOleContract.MathTypeOleMode : FormulaOleContract.WordOmmlMode);
        AssertEqual(latex.Count, readback.Targets.Count, "Large redraw host count");
        if (mode == FormulaOleContract.NativeOleMode)
        {
            var numbers = service.GetCanonicalEquationReferenceTargets(document)
                .Where(target => target.Source == EquationReferenceSource.VisualTeX)
                .OrderBy(target => target.Position).Select(target => target.NumberText).ToArray();
            AssertTrue(numbers.SequenceEqual(Enumerable.Range(1, latex.Count / 2).Select(number => number.ToString())),
                "Large redraw visible number sequence: " + string.Join(",", numbers));
        }
        var expectedMath = new Dictionary<string, string>();
        foreach (var line in File.ReadLines(journal))
        {
            using var record = JsonDocument.Parse(line);
            expectedMath[record.RootElement.GetProperty("latex").GetString()!] = record.RootElement.GetProperty("mathMl").GetString()!;
        }
        for (var index = 0; index < latex.Count; index++)
        {
            var target = readback.Targets[index];
            AssertEqual(index % 2 == 1, target.Numbered, "Large redraw numbering state " + (index + 1));
            AssertEqual(index % 2 == 1 ? "block" : "inline", target.DisplayMode, "Large redraw display state " + (index + 1));
            if (mode == FormulaOleContract.NativeOleMode)
                AssertEqual(latex[index], target.Latex, "Large redraw durable LaTeX " + (index + 1));
            else
            {
                var expected = expectedMath[latex[index]];
                var expectedSignature = WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(expected);
                var actualSignature = WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(target.SourceMathMl!);
                if (expectedSignature != actualSignature)
                    File.WriteAllText(Path.Combine(artifactRoot, $"semantic-mismatch-{index + 1}.json"), JsonSerializer.Serialize(new { expected, actual = target.SourceMathMl, expectedSignature, actualSignature }));
                AssertEqual(expectedSignature, actualSignature, "Large redraw mathematical content " + (index + 1));
            }
        }
    }
}
