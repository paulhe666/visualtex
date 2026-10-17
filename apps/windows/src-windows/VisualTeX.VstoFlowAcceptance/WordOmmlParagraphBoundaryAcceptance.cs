using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // A LaTeX redraw replaces formulas paragraph by paragraph and must keep every
    // paragraph mark. Word's paragraph-level InsertXML drops the replaced
    // paragraph's own mark when the Normal style carries no paragraph properties
    // (English templates, Pandoc/python-docx documents) and joins the paragraph
    // to the next one, or to the first cell of a following table. Both kinds of
    // Normal style are checked, around text, inline formulas, an empty
    // paragraph and a table.
    private static void RunOwnedOmmlParagraphBoundaryAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>E</mi><mo>=</mo>"
            + "<mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></math>";
        var layouts = new (string Name, string Text, bool TableAfterDisplay)[]
        {
            ("text-display-text", "Before.\r$$E=mc^2$$\rAfter.", false),
            ("inline-display-inline", "Case $E=mc^2$ end.\r$$E=mc^2$$\rCase $E=mc^2$ end.", false),
            ("display-empty-text", "Before.\r$$E=mc^2$$\r\rAfter.", false),
            ("display-table", "Before $E=mc^2$ text.\r$$E=mc^2$$\rTABLE\rAfter.", true),
        };

        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            var service = new WordFormulaService(application);
            foreach (var plainNormal in new[] { true, false })
            foreach (var (name, text, tableAfterDisplay) in layouts)
            {
                var label = (plainNormal ? "plain Normal " : "justified Normal ") + name;
                Word.Document? document = null;
                Word.Style? normal = null;
                Word.Range? content = null;
                Word.Range? tableAnchor = null;
                Word.Table? table = null;
                try
                {
                    document = application.Documents.Add();
                    document.Activate();
                    normal = document.Styles[Word.WdBuiltinStyle.wdStyleNormal];
                    normal.ParagraphFormat.Alignment = plainNormal
                        ? Word.WdParagraphAlignment.wdAlignParagraphLeft
                        : Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    normal.ParagraphFormat.WidowControl = plainNormal ? -1 : 0;
                    content = document.Content;
                    content.Text = text;
                    if (tableAfterDisplay)
                    {
                        tableAnchor = document.Paragraphs[3].Range;
                        tableAnchor.Text = string.Empty;
                        table = document.Tables.Add(tableAnchor, 1, 2);
                        table.Cell(1, 1).Range.Text = "c1";
                        table.Cell(1, 2).Range.Text = "c2";
                    }
                    var paragraphsBefore = document.Paragraphs.Count;
                    var tablesBefore = document.Tables.Count;

                    var plan = service.CaptureLatexRedrawPlan(wholeDocument: true);
                    var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
                    foreach (var target in plan.Targets)
                    {
                        var session = CreateOmmlMathTypeAcceptanceSession(
                            mathMl, target.DisplayMode, numbered: false, FormulaOleContract.WordOmmlMode);
                        session.Lines[0].Latex = target.Latex;
                        prepared[target.Id] = new PreparedWordBulkFormula
                        {
                            Run = new WordBulkRun
                            {
                                Id = target.Id,
                                IsFormula = true,
                                Latex = target.Latex,
                                DisplayMode = target.DisplayMode,
                            },
                            Session = session,
                            MathMl = mathMl,
                        };
                    }
                    var result = service.ApplyLatexRedrawPlan(plan, prepared);

                    AssertTrue(plan.Targets.Count > 0 && result.FormulaCount == plan.Targets.Count,
                        $"{label}: every formula is redrawn ({result.FormulaCount}/{plan.Targets.Count}).");
                    AssertTrue(document.OMaths.Count == plan.Targets.Count,
                        $"{label}: one OMath per formula ({document.OMaths.Count}).");
                    AssertTrue(document.Paragraphs.Count == paragraphsBefore,
                        $"{label}: paragraphs kept ({paragraphsBefore} -> {document.Paragraphs.Count}).");
                    AssertTrue(document.Tables.Count == tablesBefore,
                        $"{label}: tables kept ({tablesBefore} -> {document.Tables.Count}).");
                    var finalText = document.Content.Text ?? string.Empty;
                    AssertTrue(finalText.IndexOf('\v') < 0,
                        $"{label}: no paragraph turned into a line break.");
                    if (tableAfterDisplay)
                    {
                        var cellText = (document.Tables[1].Cell(1, 1).Range.Text ?? string.Empty)
                            .TrimEnd('\a', '\r');
                        AssertTrue(cellText == "c1",
                            $"{label}: the table's first cell is unchanged ('{cellText}').");
                    }
                    Console.WriteLine($"PASS {label}: {plan.Targets.Count} formulas, {paragraphsBefore} paragraphs.");
                }
                finally
                {
                    Release(table);
                    Release(tableAnchor);
                    Release(content);
                    Release(normal);
                    try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                    Release(document);
                }
            }
            Console.WriteLine("PASS OMML paragraph boundaries: every redraw kept its paragraph marks.");
        }
        finally
        {
            QuitWordApplicationIfOwned(application);
            Release(application);
        }
    }
}
