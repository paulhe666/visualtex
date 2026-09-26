using System.Globalization;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeInlineParagraphOwnershipAcceptance(string artifactRoot)
    {
        // Supplemental live Word structural coverage. The actual user documents
        // are also exercised separately through their visible installed Ribbon.
        var modes = new[] {
            Word.WdBaselineAlignment.wdBaselineAlignAuto,
            Word.WdBaselineAlignment.wdBaselineAlignBaseline,
            Word.WdBaselineAlignment.wdBaselineAlignCenter,
            Word.WdBaselineAlignment.wdBaselineAlignTop,
            Word.WdBaselineAlignment.wdBaselineAlignFarEast50,
        };
        var assetRoot = Path.Combine(artifactRoot, "paragraph-ownership-assets");
        Directory.CreateDirectory(assetRoot);
        var svgPath = Path.Combine(assetRoot, "fallback.svg");
        File.WriteAllText(svgPath, "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"60\" height=\"40\" viewBox=\"0 0 60 40\"><text x=\"2\" y=\"24\">a/b</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 60, 40);
        Word.Application? application = null;
        try
        {
            application = CreateFreshAcceptanceAutomationWord(artifactRoot);
            foreach (var compatibility in new[] { 14, 15 })
            foreach (var inTable in new[] { false, true })
            {
                Word.Document? document = null;
                try
                {
                    document = application.Documents.Add();
                    document.Activate();
                    document.SetCompatibilityMode(compatibility);
                    document.Content.Font.Name = "Times New Roman";
                    document.Content.Font.Size = 12;
                    Word.Table? table = null;
                    if (inTable)
                    {
                        var site = document.Range(0, 0);
                        try { table = document.Tables.Add(site, modes.Length, 1); }
                        finally { Release(site); }
                    }
                    var sourceText = "prefix $x$ middle $\\frac{a}{b}$ suffix $y$ END";
                    for (var index = 0; index < modes.Length; index++)
                    {
                        Word.Range? range = null;
                        Word.Cell? cell = null;
                        try
                        {
                            if (table is not null)
                            {
                                cell = table.Cell(index + 1, 1);
                                range = cell.Range;
                                range.End--;
                                range.Text = sourceText;
                            }
                            else
                            {
                                range = document.Range(document.Content.End - 1, document.Content.End - 1);
                                range.Text = sourceText + "\r";
                            }
                        }
                        finally { Release(range); Release(cell); }
                    }
                    Release(table);
                    var expected = new List<string>();
                    for (var index = 0; index < modes.Length; index++)
                    {
                        var paragraph = InlineOwnershipFixtureParagraph(document, index, inTable);
                        var range = paragraph.Range;
                        var format = range.ParagraphFormat;
                        try
                        {
                            format.BaseLineAlignment = modes[index];
                            format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                            format.LineSpacing = 19;
                            format.SpaceBefore = 3;
                            format.SpaceAfter = 5;
                            format.LeftIndent = 9;
                            format.FirstLineIndent = 2;
                            format.KeepTogether = -1;
                            format.TabStops.Add(140, Word.WdTabAlignment.wdAlignTabRight);
                            var prefix = document.Range(range.Start, range.Start + 6);
                            try { prefix.Font.Bold = -1; prefix.Font.Position = -2; }
                            finally { Release(prefix); }
                            expected.Add(InlineParagraphSignature(range));
                        }
                        finally { Release(format); Release(range); Release(paragraph); }
                    }
                    document.SaveAs2(Path.Combine(artifactRoot, $"ownership-{compatibility}-table-{inTable}-source.docx"));
                    Console.WriteLine($"[OWNERSHIP SOURCE] compatibility={compatibility}; table={inTable}; paragraphs={document.Paragraphs.Count}; text='{document.Content.Text.Replace("\r", "<CR>").Replace("\a", "<CELL>")}'");
                    var service = new WordFormulaService(application);
                    var plan = service.CaptureLatexRedrawPlan(wholeDocument: true);
                    AssertEqual(15, plan.Targets.Count, "Inline ownership fixture source count changed.");
                    var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
                    foreach (var target in plan.Targets)
                    {
                        var mathMl = target.Latex == "x" ? "<math><mi>x</mi></math>"
                            : target.Latex == "y" ? "<math><mi>y</mi></math>"
                            : "<math><mfrac><mi>a</mi><mi>b</mi></mfrac></math>";
                        var session = CreateOmmlMathTypeAcceptanceSession(mathMl, "inline", false, FormulaOleContract.MathTypeOleMode);
                        prepared.Add(target.Id, new PreparedWordBulkFormula {
                            Run = new WordBulkRun { Id = target.Id, IsFormula = true, Latex = target.Latex, DisplayMode = "inline" },
                            Session = session, MathMl = mathMl, EmfPath = emfPath,
                        });
                    }
                    var result = service.ApplyLatexRedrawPlan(plan, prepared);
                    document.SaveAs2(Path.Combine(artifactRoot, $"ownership-{compatibility}-table-{inTable}-result.docx"));
                    Console.WriteLine($"[OWNERSHIP RESULT] compatibility={compatibility}; table={inTable}; paragraphs={document.Paragraphs.Count}; text='{document.Content.Text.Replace("\r", "<CR>").Replace("\a", "<CELL>")}'");
                    AssertEqual(15, result.FormulaCount, "Inline ownership redraw skipped targets.");
                    AssertEqual(0, result.Failures.Count, "Inline ownership redraw reported failures.");
                    for (var index = 0; index < modes.Length; index++)
                    {
                        var paragraph = InlineOwnershipFixtureParagraph(document, index, inTable);
                        var range = paragraph.Range;
                        try
                        {
                            AssertEqual(expected[index], InlineParagraphSignature(range),
                                $"Paragraph format changed: compatibility={compatibility}, table={inTable}, alignment={modes[index]}.");
                            AssertEqual("prefix \u0001 middle \u0001 suffix \u0001 END" + (inTable ? "\r\a" : "\r"), range.Text,
                                "Inline insertion changed surrounding text or paragraph boundaries.");
                            var prefix = document.Range(range.Start, range.Start + 6);
                            try
                            {
                                AssertEqual(-1, prefix.Font.Bold, "Surrounding bold text changed.");
                                AssertEqual(-2, prefix.Font.Position, "Surrounding text vertical position changed.");
                            }
                            finally { Release(prefix); }
                            var baselineAligned = modes[index] is Word.WdBaselineAlignment.wdBaselineAlignAuto
                                or Word.WdBaselineAlignment.wdBaselineAlignBaseline;
                            var fractionPosition = 0;
                            for (var s = 1; s <= range.InlineShapes.Count; s++)
                            {
                                var shape = range.InlineShapes[s];
                                var shapeRange = shape.Range;
                                var field = shapeRange.Fields[1];
                                var fieldResult = field.Result;
                                try
                                {
                                    if (!baselineAligned) AssertEqual(0, fieldResult.Font.Position, "Non-baseline paragraph received baseline descent twice.");
                                    if (s == 2) fractionPosition = fieldResult.Font.Position;
                                }
                                finally { Release(fieldResult); Release(field); Release(shapeRange); Release(shape); }
                            }
                            if (baselineAligned && fractionPosition >= 0)
                                throw new InvalidDataException("Baseline paragraph lost the fraction's native descent correction.");
                        }
                        finally { Release(range); Release(paragraph); }
                    }
                    Console.WriteLine($"[INLINE OWNERSHIP PASS] compatibility={compatibility}; table={inTable}; alignments=5; formulas=15; prose and paragraph properties preserved.");
                }
                finally
                {
                    if (document is not null) document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    Release(document);
                }
            }
        }
        finally { try { QuitWordApplicationIfOwned(application); } finally { Release(application); } }
    }

    private static Word.Paragraph InlineOwnershipFixtureParagraph(Word.Document document, int index, bool inTable)
    {
        if (!inTable) return document.Paragraphs[index + 1];
        var tables = document.Tables;
        var table = tables[1];
        var cell = table.Cell(index + 1, 1);
        var range = cell.Range;
        var paragraphs = range.Paragraphs;
        try { return paragraphs[1]; }
        finally { Release(paragraphs); Release(range); Release(cell); Release(table); Release(tables); }
    }

    private static string InlineParagraphSignature(Word.Range range)
    {
        var format = range.ParagraphFormat;
        try
        {
            var tabs = format.TabStops;
            var positions = new List<string>();
            try
            {
                foreach (Word.TabStop stop in tabs)
                {
                    positions.Add($"{stop.Position.ToString(CultureInfo.InvariantCulture)}:{(int)stop.Alignment}:{(int)stop.Leader}");
                    Release(stop);
                }
            }
            finally { Release(tabs); }
            return string.Join("|", (int)format.BaseLineAlignment, (int)format.LineSpacingRule,
                format.LineSpacing, format.SpaceBefore, format.SpaceAfter, format.LeftIndent,
                format.FirstLineIndent, format.KeepTogether, string.Join(",", positions));
        }
        finally { Release(format); }
    }
}
