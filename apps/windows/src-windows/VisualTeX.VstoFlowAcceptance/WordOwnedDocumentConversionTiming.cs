using System.Diagnostics;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    // Times the production whole-document conversion on a copy of a real user
    // document. VISUALTEX_PERF_LIMIT=N first copies the document's first N OMML
    // formulas into a new empty document and times that instead.
    private static string DescribeTables(Word.Document document) =>
        $"count={document.Tables.Count} omaths={document.OMaths.Count} " + string.Join(" ",
            document.Tables.Cast<Word.Table>().Select(table => $"[{table.Range.Start}-{table.Range.End} rows={table.Rows.Count}]"));

    private static void RunOwnedDocumentConversionTiming(string root)
    {
        var tracePath = Path.Combine(root, "conversion-trace.log");
        Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
        var sourceDocument = Environment.GetEnvironmentVariable("VISUALTEX_PERF_DOCUMENT")
            ?? throw new InvalidOperationException("VISUALTEX_PERF_DOCUMENT is required.");
        var source = Environment.GetEnvironmentVariable("VISUALTEX_PERF_SOURCE") ?? "Omml";
        var target = Environment.GetEnvironmentVariable("VISUALTEX_PERF_TARGET") ?? "VisualTeX";
        var limitText = Environment.GetEnvironmentVariable("VISUALTEX_PERF_LIMIT");
        var output = Path.Combine(root, source + "-to-" + target + ".docx");
        if (File.Exists(output)) File.Delete(output);
        var timings = new Dictionary<string, double>();
        var watch = Stopwatch.StartNew();
        if (int.TryParse(limitText, out var limit) && limit > 0)
        {
            using var seedHost = new WordPerformanceHost(sourceDocument);
            var sourceMaths = seedHost.Document.OMaths;
            var count = Math.Min(limit, sourceMaths.Count);
            var copy = seedHost.Application.Documents.Add();
            for (var i = 1; i <= count; i++)
            {
                var end = copy.Content;
                end.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                end.InsertAfter("Formula " + i + ": ");
                end.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                end.FormattedText = sourceMaths[i].Range.FormattedText;
                copy.Content.InsertParagraphAfter();
            }
            copy.SaveAs2(output, Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
            copy.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            timings["seedMs"] = watch.Elapsed.TotalMilliseconds;
        }
        else
        {
            File.Copy(sourceDocument, output);
        }

        string Mode(string name) => name switch {
            "VisualTeX" => FormulaOleContract.NativeOleMode,
            "MathType" => FormulaOleContract.MathTypeOleMode,
            "Omml" => FormulaOleContract.WordOmmlMode,
            _ => throw new InvalidOperationException("Unknown formula format: " + name),
        };
        var sourceMode = Mode(source);
        var targetMode = Mode(target);
        int requested, converted, remaining;
        if (File.Exists(tracePath)) File.Delete(tracePath);
        watch.Restart();
        using (var host = new WordPerformanceHost(output))
        {
            timings["hostOpenAndAddInMs"] = watch.Elapsed.TotalMilliseconds;
            // VISUALTEX_PERF_EXTRA_DOC=1 opens a second document with its own
            // OMML formula in the same Word, then re-activates the target.
            Word.Document? extra = null;
            if (Environment.GetEnvironmentVariable("VISUALTEX_PERF_EXTRA_DOC") == "1")
            {
                extra = host.Application.Documents.Add();
                extra.Content.Text = "Second document x=1";
                var formula = extra.Range(extra.Content.End - 4, extra.Content.End - 1);
                extra.OMaths.Add(formula);
                host.Document.Activate();
                Console.WriteLine($"[EXTRA DOC] omaths={extra.OMaths.Count} inlineShapes={extra.InlineShapes.Count} windowState={extra.ActiveWindow.WindowState}");
            }
            Console.WriteLine($"[VIEW] at open: {host.Document.ActiveWindow.View.Type} documents={host.Application.Documents.Count} active={host.Application.ActiveDocument.Name}");
            var service = new WordFormulaService(host.Application);
            requested = service.CountFormulaObjectsForLatex(true, sourceMode);
            // VISUALTEX_PERF_SELECT_FIRST=N converts only the first N formulas of
            // the full document, so the long tail after them is still laid out.
            var scope = "Document";
            if (int.TryParse(Environment.GetEnvironmentVariable("VISUALTEX_PERF_SELECT_FIRST"), out var first) && first > 0)
            {
                var maths = host.Document.OMaths;
                var end = maths[Math.Min(first, maths.Count)].Range.End;
                host.Document.Range(0, end).Select();
                requested = Math.Min(first, maths.Count);
                scope = "Selection";
            }
            // VISUALTEX_PERF_METHOD names another ribbon handler, for example
            // OnRedrawDocumentToOmml, instead of the format conversion.
            var method = Environment.GetEnvironmentVariable("VISUALTEX_PERF_METHOD");
            if (string.IsNullOrWhiteSpace(method)) method = "OnConvert" + source + "To" + target + scope;
            Console.WriteLine($"[TABLES] before: {DescribeTables(host.Document)}");
            using (var metrics = new WordOperationMetrics())
            {
                watch.Restart();
                typeof(ThisAddIn).GetMethod(method)!
                    .Invoke(host.AddIn, new object[] { new object() });
                WaitForAddInIdleFast(host.AddIn, TimeSpan.FromMinutes(15));
                timings["conversionMs"] = watch.Elapsed.TotalMilliseconds;
                File.WriteAllText(Path.Combine(root, "operation-metrics.json"), JsonSerializer.Serialize(
                    metrics.Entries, new JsonSerializerOptions { WriteIndented = true }));
            }
            if (extra is not null)
            {
                Console.WriteLine($"[EXTRA DOC] after: omaths={extra.OMaths.Count} inlineShapes={extra.InlineShapes.Count} text='{extra.Content.Text.Trim()}' windowState={extra.ActiveWindow.WindowState}");
                Console.WriteLine($"[TARGET DOC] after: windowState={host.Document.ActiveWindow.WindowState} active={host.Application.ActiveDocument.FullName == host.Document.FullName}");
                host.Document.Activate();
            }
            Console.WriteLine($"[VIEW] after conversion: {host.Document.ActiveWindow.View.Type} windowState={host.Document.ActiveWindow.WindowState} active={host.Application.ActiveDocument.Name}");
            Console.WriteLine($"[TABLES] after: {DescribeTables(host.Document)}");
            converted = service.CountFormulaObjectsForLatex(true, targetMode);
            Console.WriteLine($"[VIEW] after count target: {host.Document.ActiveWindow.View.Type} converted={converted}");
            remaining = service.CountFormulaObjectsForLatex(true, sourceMode);
            Console.WriteLine($"[VIEW] after count source: {host.Document.ActiveWindow.View.Type} remaining={remaining} omaths={host.Document.OMaths.Count}");
            extra?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            host.Document.Save();
        }
        File.WriteAllText(Path.Combine(root, "timing-result.json"), JsonSerializer.Serialize(new {
            sourceDocument, source, target, limit = limitText, requested, converted, remaining, timings,
            wordDll = typeof(ThisAddIn).Assembly.Location,
        }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"[DOCUMENT TIMING] {source}->{target} requested={requested} converted={converted} remaining={remaining} conversionMs={timings["conversionMs"]:F0}");
    }
}
