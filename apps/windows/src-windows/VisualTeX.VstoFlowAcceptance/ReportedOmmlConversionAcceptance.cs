using System.Text;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    // Invoked directly by the audited registry-free runner, never through Main.
    private static void RunReportedOmmlConversions(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord, "This test never attaches user Word.");
        Directory.CreateDirectory(artifactRoot);
        using var filter = OfficeComMessageFilter.Register();
        var rows = new List<object>();
        var scenario = Environment.GetEnvironmentVariable("VISUALTEX_CP267_SCENARIO") ?? "basic";
        Word.Application? app = null;
        Word.Document? doc = null;
        try
        {
            var owned = Path.Combine(artifactRoot, "owned");
            Directory.CreateDirectory(owned);
            app = CreateFreshAcceptanceAutomationWord(owned);
            foreach (var display in scenario == "tracked" ? new[] { true } : new[] { false, true })
            foreach (var colored in scenario == "unicode" ? new[] { true } : new[] { false, true })
            foreach (var toOle in new[] { false, true })
            {
                var name = (display ? "display" : "inline") + (colored ? "-blue" : "-plain") + (toOle ? "-ole" : "-latex");
                string? error = null; string? expected = null; string? before = null; string? after = null;
                string? recovered = null; bool undo = false, redo = false, persisted = false;
                Word.Range? insertion = null; Word.Range? equation = null;
                try
                {
                    Console.WriteLine("START " + name);
                    doc = app.Documents.Add();
                    doc.Activate();
                    doc.Content.Text = display ? "LEFT\r\rRIGHT\r" : "LEFT RIGHT\r";
                    insertion = doc.Range(5, 5);
                    var mathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow"
                        + (colored ? " mathcolor=\"#0070C0\"" : "")
                        + "><mi>" + (scenario == "unicode" ? "\U00020000" : "x") + "</mi><mo>+</mo><mi>y</mi></mrow></math>";
                    equation = WordOmmlConverter.ReplaceWithPreparedOmmlDirect(doc, insertion,
                        WordOmmlConverter.TransformMathMlToOmml(mathMl), display, doc.OMathFontName);
                    var host = WordFormulaHostLookup.ResolveLocal(doc, equation, WordFormulaHostKind.Omml)
                        ?? throw new InvalidDataException("No exact native source.");
                    AssertEqual(display, host.Display, "Fixture display state");
                    var native = WordFormulaHostSemanticReader.Read(doc, host);
                    recovered = native.Latex;
                    if (colored) AssertTrue(recovered.Contains("\\textcolor{#0070C0}"), "Native source lost exact blue before conversion");
                    before = doc.Content.Text;
                    File.WriteAllText(Path.Combine(artifactRoot, name + "-source.xml"), doc.Content.WordOpenXML);
                    doc.SaveAs2(Path.Combine(artifactRoot, name + "-source.docx"), Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
                    if (scenario == "tracked") doc.TrackRevisions = true;
                    var service = new WordFormulaService(app);
                    if (toOle)
                    {
                        var plan = service.CaptureFormulaFormatConversionPlan(true, FormulaOleContract.WordOmmlMode, FormulaOleContract.NativeOleMode);
                        AssertEqual(1, plan.Targets.Count, "One conversion target");
                        var target = plan.Targets[0];
                        expected = target.Latex;
                        var assetRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                            "VisualTeX", "office", "temp", "cp267-" + Guid.NewGuid().ToString("N"));
                        Directory.CreateDirectory(assetRoot);
                        // Synthetic previews deliberately test structural writeback, not visual rendering.
                        var assets = CreatePerformanceOleAssets(assetRoot, 1, expected, PerformanceSvg(1));
                        var prepared = new Dictionary<string, PreparedWordBulkFormula> {
                            [target.Id] = new PreparedWordBulkFormula {
                                Run = new WordBulkRun { Id=target.Id, IsFormula=true, Latex=target.Latex, DisplayMode=target.DisplayMode },
                                Session = CreateSimpleFormatTargetSession(target, FormulaOleContract.NativeOleMode, target.SourceMathMl!),
                                MathMl=target.SourceMathMl, PngPath=assets.PngPath, EmfPath=assets.EmfPath
                            }
                        };
                        var result = service.ApplyFormulaFormatConversionPlan(plan, prepared);
                        AssertEqual(1, result.FormulaCount, "Converted OLE count");
                        AssertEqual(0, result.FailedFormulaCount, string.Join(" | ", result.Failures));
                        AssertEqual(1, CountVisualTeXNativeOleShapes(doc), "Physical OLE count");
                        var ole = doc.InlineShapes[1];
                        try {
                            var range = ole.Range;
                            try {
                                var targetHost = WordFormulaHostLookup.ResolveLocal(doc, range, WordFormulaHostKind.VisualTeX)
                                    ?? throw new InvalidDataException("No converted OLE host.");
                                var payload = WordFormulaHostSemanticReader.Read(doc, targetHost);
                                AssertEqual(expected, payload.Latex, "OLE embedded exact source");
                            } finally { Release(range); }
                        } finally { Release(ole); }
                        after = doc.Content.Text;
                        AssertEqual(display ? "LEFT\r\t\u0001\rRIGHT\r\r" : "LEFT \u0001RIGHT\r\r", after,
                            "Exact OLE and surrounding prose, including owned display TAB and paragraph marks");
                    }
                    else
                    {
                        expected = (display ? "$$" : "$") + recovered + (display ? "$$" : "$");
                        var result = service.ConvertFormulaObjectsToLatex(true, FormulaOleContract.WordOmmlMode);
                        AssertEqual(1, result.FormulaCount, "Converted LaTeX count");
                        after = doc.Content.Text;
                        AssertEqual(display ? "LEFT\r" + expected + "\rRIGHT\r\r" : "LEFT " + expected + "RIGHT\r\r", after, "Exact source with surrounding prose");
                    }
                    AssertEqual(0, doc.OMaths.Count, "Source OMath removed");
                    File.WriteAllText(Path.Combine(artifactRoot, name + "-converted.xml"), doc.Content.WordOpenXML);
                    doc.Undo(1); AssertEqual(before, doc.Content.Text, "Single Undo exact source text");
                    AssertEqual(1, doc.OMaths.Count, "Single Undo native source"); undo=true;
                    doc.Redo(1); AssertEqual(after, doc.Content.Text, "Single Redo exact target text"); redo=true;
                    var path = Path.Combine(artifactRoot, name + "-converted.docx");
                    doc.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles:false);
                    doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges); Release(doc); doc=null;
                    doc=app.Documents.Open(path, ReadOnly:false, AddToRecentFiles:false, Visible:false);
                    AssertEqual(after, doc.Content.Text, "Saved reopen exact target text");
                    AssertEqual(0, doc.OMaths.Count, "Saved reopen no source OMath");
                    AssertEqual(toOle ? 1 : 0, CountVisualTeXNativeOleShapes(doc), "Saved reopen target inventory");
                    persisted=true;
                }
                catch (Exception ex) { error=ex.ToString(); Console.WriteLine("FAIL " + name + ": " + ex.Message);
                    if (doc is not null) try { File.WriteAllText(Path.Combine(artifactRoot,name+"-failed.xml"),doc.Content.WordOpenXML); } catch {} }
                finally {
                    Release(equation); Release(insertion);
                    if (doc is not null) { try {doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);} catch {} Release(doc); doc=null; }
                    rows.Add(new {name,display,colored,toOle,expected,recovered,before,after,undo,redo,persisted,error});
                    File.WriteAllText(Path.Combine(artifactRoot,"results.json"),JsonSerializer.Serialize(rows,new JsonSerializerOptions {WriteIndented=true}));
                }
            }
        }
        finally { if(doc is not null) {try{doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);}catch{} Release(doc);}
            QuitWordApplicationIfOwned(app); Release(app); ForceComCleanup(); }
    }
}
