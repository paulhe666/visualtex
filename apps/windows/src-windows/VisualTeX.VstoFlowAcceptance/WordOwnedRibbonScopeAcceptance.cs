using System.Diagnostics;
using System.Text.Json;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // Word ribbon rows: every conversion direction owns a "selection" and a
    // "whole document" button. This acceptance drives both callbacks of all six
    // rows in an owned Word instance: the selection button must convert exactly
    // the selected formula, the document button the remaining ones.
    private static void RunWordOwnedRibbonScopeAcceptance(string root)
    {
        var tracePath = Path.Combine(root, "conversion-trace.log");
        Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
        var latex = new[] { @"x^{2}+y^{2}=z^{2}", @"\frac{a+b}{c}", @"\sqrt{x+1}" };
        var rows = new (string Source, string Target)[]
        {
            ("VisualTeX", "MathType"), ("MathType", "Omml"), ("Omml", "VisualTeX"),
            ("VisualTeX", "Omml"), ("Omml", "MathType"), ("MathType", "VisualTeX"),
        };
        string Mode(string name) => name switch
        {
            "VisualTeX" => FormulaOleContract.NativeOleMode,
            "MathType" => FormulaOleContract.MathTypeOleMode,
            "Omml" => FormulaOleContract.WordOmmlMode,
            _ => throw new InvalidOperationException(name),
        };
        var results = new List<object>();
        using var host = new WordPerformanceHost(null);
        var ribbon = XDocument.Parse(host.AddIn.GetCustomUI("Microsoft.Word.Document"));
        XNamespace ns = ribbon.Root!.Name.Namespace;
        var group = ribbon.Descendants(ns + "group")
            .Single(item => (string?)item.Attribute("id") == "VisualTeX.WordVsto.FormatConversionGroup");
        foreach (var (source, target) in rows)
        {
            foreach (var scope in new[] { "Selection", "Document" })
            {
                var callback = "OnConvert" + source + "To" + target + scope;
                var buttons = group.Descendants(ns + "button")
                    .Where(item => (string?)item.Attribute("onAction") == callback).ToList();
                AssertEqual(1, buttons.Count, "Ribbon button for " + callback);
                AssertTrue(typeof(ThisAddIn).GetMethod(callback) is not null, "Callback method " + callback);
            }
        }
        AssertEqual(12, group.Descendants(ns + "button").Count(), "Word conversion group button count");
        AssertTrue(!ribbon.Descendants(ns + "splitButton").Any(item => ((string?)item.Attribute("id") ?? "").Contains("Convert")),
            "Retired split buttons must be gone.");

        host.Document.Content.Text = "VisualTeX ribbon scope\r"
            + string.Join("\r", latex.Select(item => "$$" + item + "$$")) + "\r";
        host.AddIn.OnRedrawDocumentToOle(new object());
        WaitForAddInIdleFast(host.AddIn, TimeSpan.FromMinutes(5));
        var service = new WordFormulaService(host.Application);
        AssertEqual(latex.Length, service.CountFormulaObjectsForLatex(true, FormulaOleContract.NativeOleMode), "VisualTeX seed");

        int Count(string mode) => service.CountFormulaObjectsForLatex(true, mode);
        string Trace()
        {
            if (!File.Exists(tracePath)) return string.Empty;
            using var stream = new FileStream(tracePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }
        void Run(string callback)
        {
            var before = Trace().Length;
            var watch = Stopwatch.StartNew();
            typeof(ThisAddIn).GetMethod(callback)!.Invoke(host.AddIn, new object[] { new object() });
            WaitForAddInIdleFast(host.AddIn, TimeSpan.FromMinutes(5));
            var added = Trace().Substring(before);
            AssertTrue(!added.Contains("format-conversion-failed"), callback + " failed: " + added);
            Console.WriteLine($"[RIBBON SCOPE] {callback} {watch.ElapsedMilliseconds} ms");
        }
        void SelectMiddleFormula(string sourceMode, string targetMode)
        {
            var plan = service.CaptureFormulaFormatConversionPlan(true, sourceMode, targetMode);
            AssertEqual(latex.Length, plan.Targets.Count, "Source formulas before selection conversion");
            var start = plan.Targets.Select(item => item.SourceStart).OrderBy(value => value).ElementAt(1);
            Word.Range? anchor = null; Word.Paragraphs? paragraphs = null; Word.Paragraph? paragraph = null; Word.Range? range = null;
            try
            {
                anchor = host.Document.Range(start, start);
                paragraphs = anchor.Paragraphs;
                paragraph = paragraphs[1];
                range = paragraph.Range;
                range.Select();
            }
            finally { Release(range); Release(paragraph); Release(paragraphs); Release(anchor); }
        }

        foreach (var (source, target) in rows)
        {
            var sourceMode = Mode(source);
            var targetMode = Mode(target);
            AssertEqual(latex.Length, Count(sourceMode), source + " count before row " + source + "->" + target);
            SelectMiddleFormula(sourceMode, targetMode);
            Run("OnConvert" + source + "To" + target + "Selection");
            AssertEqual(latex.Length - 1, Count(sourceMode), "Selection must leave the other sources: " + source + "->" + target);
            AssertEqual(1, Count(targetMode), "Selection must convert exactly one: " + source + "->" + target);
            var converted = service.CaptureFormulaFormatConversionPlan(true, targetMode, sourceMode).Targets.Single().SourceStart;
            var remaining = service.CaptureFormulaFormatConversionPlan(true, sourceMode, targetMode).Targets
                .Select(item => item.SourceStart).OrderBy(value => value).ToArray();
            AssertTrue(remaining[0] < converted && converted < remaining[1], "Selection converted the selected (middle) formula: " + source + "->" + target);
            Run("OnConvert" + source + "To" + target + "Document");
            AssertEqual(0, Count(sourceMode), "Document must convert every source: " + source + "->" + target);
            AssertEqual(latex.Length, Count(targetMode), "Document target count: " + source + "->" + target);
            results.Add(new { row = source + "->" + target, selection = "1 of 3", document = "3 of 3" });
        }
        var output = Path.Combine(root, "ribbon-scope.docx");
        host.Save(output);
        File.WriteAllText(Path.Combine(root, "ribbon-scope-result.json"),
            JsonSerializer.Serialize(new { status = "Passed", results }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine("PASS Word ribbon selection/document scope acceptance");
    }
}
