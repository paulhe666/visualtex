using System.Text.Json;
using System.Text.RegularExpressions;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordVisualTeXMixedNumbering(string artifactRoot)
    {
        var source = Environment.GetEnvironmentVariable("VISUALTEX_MIXED_NUMBERING_SOURCE")!;
        var output = Path.Combine(artifactRoot, "mixed-number-formats.docx");
        var previousDefault = WordEquationNumbering.GetDefaultEquationNumberFormatId();
        var formats = new[] { EquationNumberFormat.Heading1DotId, EquationNumberFormat.Heading2DotId,
            EquationNumberFormat.Heading1DashId, EquationNumberFormat.Heading2DashId, EquationNumberFormat.ContinuousId };
        var results = new List<object>();
        string[] expected = Enumerable.Range(1, 6).Select(n => n.ToString()).ToArray();
        try
        {
            using (var host = new WordPerformanceHost(source))
            {
                var doc = host.Document;
                var service = new WordFormulaService(host.Application);
                var headings = doc.Range(0, 0);
                try { headings.InsertBefore("4 Test chapter\r4.2 Test section\r"); }
                finally { Release(headings); }
                for (var index = 1; index <= 2; index++)
                {
                    var range = doc.Paragraphs[index].Range;
                    try { object style = index == 1 ? Word.WdBuiltinStyle.wdStyleHeading1 : Word.WdBuiltinStyle.wdStyleHeading2; range.set_Style(ref style); }
                    finally { Release(range); }
                }
                var targets = service.GetCanonicalEquationReferenceTargets(doc).OrderBy(target => target.Position).ToArray();
                AssertEqual(6, targets.Length, "Mixed numbering reference seed count");
                for (var index = 0; index < targets.Length; index++)
                {
                    var selection = host.Application.Selection;
                    try
                    {
                        selection.SetRange(doc.Content.End - 1, doc.Content.End - 1);
                        selection.TypeText($"REF[{index + 1}]:");
                        service.InsertEquationReferenceCore(doc, selection, targets[index], EquationReferenceStyle.NumberOnly, Word.WdColor.wdColorAutomatic);
                        selection.SetRange(doc.Content.End - 1, doc.Content.End - 1);
                        selection.TypeText($":ENDREF[{index + 1}]\r");
                    }
                    finally { Release(selection); }
                }
                var topology = LargeRedrawTableTopology(doc);
                AssertVisualTeXMixedNumbers(host, expected);
                foreach (var formatId in formats)
                {
                    var before = expected;
                    var format = EquationNumberFormat.Resolve(formatId);
                    var prefix = format.UsesHeading ? (format.HeadingLevel == 1 ? "4" : "4.2") + format.Separator : "";
                    expected = Enumerable.Range(1, 6).Select(n => prefix + n).ToArray();
                    AssertEqual(6, service.SetEquationNumberFormat(formatId), "Mixed number format count");
                    AssertVisualTeXMixedNumbers(host, expected);
                    AssertTrue(doc.Undo(1), "Mixed number format Undo");
                    AssertVisualTeXMixedNumbers(host, before);
                    AssertTrue(doc.Redo(1), "Mixed number format Redo");
                    AssertVisualTeXMixedNumbers(host, expected);
                    service.UpdateEquationNumbers();
                    AssertVisualTeXMixedNumbers(host, expected);
                    AssertEqual(topology, LargeRedrawTableTopology(doc), "Mixed number format changed user tables");
                    results.Add(new { Format = formatId, Numbers = expected, References = true, UndoRedo = true, Refresh = true });
                    File.WriteAllText(Path.Combine(artifactRoot, "result.json"), JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));
                    Console.WriteLine("[MIXED NUMBER FORMAT PASS] " + formatId);
                }
                host.Save(output);
            }
            using (var host = new WordPerformanceHost(output)) AssertVisualTeXMixedNumbers(host, expected);
            Console.WriteLine("[MIXED NUMBER SAVE/REOPEN PASS]");
        }
        finally { WordEquationNumbering.SetDefaultEquationNumberFormatPreference(previousDefault); }
    }

    private static void AssertVisualTeXMixedNumbers(WordPerformanceHost host, string[] expected)
    {
        var service = new WordFormulaService(host.Application);
        var targets = service.GetCanonicalEquationReferenceTargets(host.Document).OrderBy(target => target.Position).ToArray();
        AssertTrue(expected.SequenceEqual(targets.Select(target => target.NumberText)),
            "Mixed number labels: " + string.Join(",", targets.Select(target => target.NumberText)));
        var text = host.Document.Content.Text;
        for (var index = 0; index < expected.Length; index++)
        {
            var match = Regex.Match(text, @"REF\[" + (index + 1) + @"\]:(.*?):ENDREF\[" + (index + 1) + @"\]", RegexOptions.Singleline);
            AssertTrue(match.Success, "Mixed reference marker missing");
            AssertEqual(expected[index], match.Groups[1].Value.Trim().Trim('(', ')'), "Mixed reference value " + (index + 1));
        }
        var formulas = service.CaptureFormulaFormatConversionPlan(true, FormulaOleContract.NativeOleMode, FormulaOleContract.WordOmmlMode);
        AssertTrue(formulas.Targets.Select(target => target.Latex).SequenceEqual(Enumerable.Range(1, 12).Select(LargeRedrawLatex)), "Mixed number operation changed formula content");
    }
}
