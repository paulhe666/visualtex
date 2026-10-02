using System.Text.Json;
using System.Text.RegularExpressions;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void RunConversionExistingTargets(string artifactRoot)
    {
        var source = Environment.GetEnvironmentVariable("VISUALTEX_CONVERSION_EXISTING_SOURCE")!;
        var results = new List<object>();
        foreach (var formatId in new[] { EquationNumberFormat.ContinuousId, EquationNumberFormat.Heading1DotId, EquationNumberFormat.Heading2DotId })
        {
            var output = Path.Combine(artifactRoot, formatId + ".docx");
            File.Copy(source, output);
            string[] expectedLatex;
            int expectedNumberCount;
            int[] referenceOrdinals;
            var format = EquationNumberFormat.Resolve(formatId);
            var prefix = format.UsesHeading ? (format.HeadingLevel == 1 ? "4." : "4.2.") : "";
            using (var host = new WordPerformanceHost(output))
            {
                var doc = host.Document;
                var service = new WordFormulaService(host.Application);
                if (format.UsesHeading)
                {
                    var start = doc.Range(0, 0);
                    try { start.InsertBefore("4 Test chapter\r4.2 Test section\r"); } finally { Release(start); }
                    for (var i = 1; i <= 2; i++)
                    {
                        var heading = doc.Paragraphs[i].Range;
                        try { object style = i == 1 ? Word.WdBuiltinStyle.wdStyleHeading1 : Word.WdBuiltinStyle.wdStyleHeading2; heading.set_Style(ref style); }
                        finally { Release(heading); }
                    }
                }
                WordEquationNumbering.SetEquationNumberFormatPreference(doc, formatId);
                var sourceTargets = service.CaptureFormulaFormatConversionPlan(true, FormulaOleContract.WordOmmlMode, FormulaOleContract.NativeOleMode).Targets;
                expectedLatex = sourceTargets.Select(item => item.Latex).ToArray();
                AssertEqual(20, expectedLatex.Length, "Existing-target daily seed count");
                expectedNumberCount = sourceTargets.Count(item => item.Numbered);
                var existingIndices = new[] { 19, 9, 1 };
                AssertTrue(existingIndices.All(index => sourceTargets[index].Numbered), "Selected existing targets are numbered");
                referenceOrdinals = existingIndices.OrderBy(index => index).Select(index => sourceTargets.Take(index + 1).Count(item => item.Numbered)).ToArray();
                foreach (var index in existingIndices)
                {
                    var plan = service.CaptureFormulaFormatConversionPlan(true, FormulaOleContract.WordOmmlMode, FormulaOleContract.NativeOleMode);
                    var range = WordFormulaOperationLocator.ResolveCapturedRange(doc, plan.Targets[index].SourceObjectId);
                    try { range.Select(); } finally { Release(range); }
                    host.AddIn.OnConvertOmmlToVisualTeXSelection(new object());
                    WaitForAddInIdleFast(host.AddIn, TimeSpan.FromMinutes(3));
                }
                AssertEqual(3, service.CountFormulaObjectsForLatex(true, FormulaOleContract.NativeOleMode), "Prepared existing targets");
                var refs = service.GetCanonicalEquationReferenceTargets(doc).Where(item => item.Source == EquationReferenceSource.VisualTeX).OrderBy(item => item.Position).ToArray();
                AssertEqual(3, refs.Length, "Existing numbered targets");
                for (var i = 0; i < refs.Length; i++)
                {
                    var selection = host.Application.Selection;
                    try
                    {
                        selection.SetRange(doc.Content.End - 1, doc.Content.End - 1);
                        selection.TypeText("EXISTREF[" + i + "]:");
                        service.InsertEquationReferenceCore(doc, selection, refs[i], EquationReferenceStyle.NumberOnly, Word.WdColor.wdColorAutomatic);
                        selection.SetRange(doc.Content.End - 1, doc.Content.End - 1);
                        selection.TypeText(":ENDEXISTREF[" + i + "]\r");
                    }
                    finally { Release(selection); }
                }
                var topology = LargeRedrawTableTopology(doc);
                host.AddIn.OnConvertOmmlToVisualTeXDocument(new object());
                WaitForAddInIdleFast(host.AddIn, TimeSpan.FromMinutes(3));
                AssertConversionExistingFinal(host, prefix, expectedLatex, expectedNumberCount, referenceOrdinals);
                AssertEqual(topology, LargeRedrawTableTopology(doc), "Existing target table topology");
                AssertTrue(doc.Undo(1), "Existing-target single Undo");
                AssertEqual(17, service.CountFormulaObjectsForLatex(true, FormulaOleContract.WordOmmlMode), "Existing-target Undo source count");
                AssertEqual(3, service.CountFormulaObjectsForLatex(true, FormulaOleContract.NativeOleMode), "Existing-target Undo existing count");
                AssertTrue(doc.Redo(1), "Existing-target single Redo");
                AssertConversionExistingFinal(host, prefix, expectedLatex, expectedNumberCount, referenceOrdinals);
                doc.Save();
            }
            using (var host = new WordPerformanceHost(output))
                AssertConversionExistingFinal(host, prefix, expectedLatex, expectedNumberCount, referenceOrdinals);
            results.Add(new { Format = formatId, Status = "Passed", Existing = 3, Converted = 17, Final = 20, Numbers = expectedNumberCount, References = 3, UndoRedo = true, SaveReopen = true });
            File.WriteAllText(Path.Combine(artifactRoot, "result.json"), JsonSerializer.Serialize(results));
            Console.WriteLine("[EXISTING TARGET PASS] " + formatId);
        }
    }

    private static void AssertConversionExistingFinal(WordPerformanceHost host, string prefix, string[] expectedLatex, int expectedNumberCount, int[] referenceOrdinals)
    {
        var service = new WordFormulaService(host.Application);
        AssertEqual(0, service.CountFormulaObjectsForLatex(true, FormulaOleContract.WordOmmlMode), "Existing-target retired source count");
        var formulas = service.CaptureFormulaFormatConversionPlan(true, FormulaOleContract.NativeOleMode, FormulaOleContract.WordOmmlMode).Targets;
        AssertTrue(expectedLatex.SequenceEqual(formulas.Select(item => item.Latex)), "Existing-target mathematical sources");
        var numbers = service.GetCanonicalEquationReferenceTargets(host.Document).Where(item => item.Source == EquationReferenceSource.VisualTeX)
            .OrderBy(item => item.Position).Select(item => item.NumberText).ToArray();
        AssertTrue(numbers.SequenceEqual(Enumerable.Range(1, expectedNumberCount).Select(i => prefix + i)), "Existing-target numbers: " + string.Join(",", numbers));
        var text = host.Document.Content.Text;
        var ordinals = referenceOrdinals;
        for (var i = 0; i < ordinals.Length; i++)
        {
            var match = Regex.Match(text, @"EXISTREF\[" + i + @"\]:(.*?):ENDEXISTREF\[" + i + @"\]", RegexOptions.Singleline);
            AssertTrue(match.Success, "Existing-target reference marker");
            AssertEqual(prefix + ordinals[i], match.Groups[1].Value.Trim().Trim('(', ')'), "Existing-target reference value");
        }
    }
}
