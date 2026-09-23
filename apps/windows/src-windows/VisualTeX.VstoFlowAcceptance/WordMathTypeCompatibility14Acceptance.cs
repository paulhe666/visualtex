using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeCompatibility14Acceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "mathtype-compat14-preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"96\" viewBox=\"0 0 240 96\"><text x=\"4\" y=\"64\" font-family=\"Times New Roman\" font-size=\"48\">x+1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 96);
        var documentPath = Path.Combine(artifactRoot, "VisualTeX-MathType-Compat14.docx");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "compatibility 14 MathType insertion\r";
            document.SaveAs2(
                FileName: documentPath,
                FileFormat: Word.WdSaveFormat.wdFormatXMLDocument,
                CompatibilityMode: 14);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(documentPath, ReadOnly: false, Visible: false);
            document.Activate();

            AssertEqual(14, document.CompatibilityMode,
                "Compatibility-mode fixture did not reopen as Word compatibility mode 14.");

            var service = new WordFormulaService(application);

            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: false,
                    latex: @"x+y"),
                SimpleMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "Compatibility-mode unnumbered MathType insertion failed.");

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "right"),
                FirstNumberedMathMl,
                emfPath);
            AssertEqual(2, document.InlineShapes.Count,
                "Compatibility-mode numbered MathType insertion failed.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Compatibility-mode numbered MathType insertion did not create MTPlaceRef.");

            Console.WriteLine(
                "[MATHTYPE COMPAT14 PASS] Unnumbered then numbered Equation.DSMT4 insertion completed in Word compatibility mode 14.");

            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(document);
            document = application.Documents.Add();
            document.Content.Text = "before\r$$a+b$$\r$$c+d$$\rafter\r";
            document.SaveAs2(
                FileName: documentPath,
                FileFormat: Word.WdSaveFormat.wdFormatXMLDocument,
                CompatibilityMode: 14);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(documentPath, ReadOnly: false, Visible: false);
            document.Activate();

            foreach (var redraw in new[]
                     {
                         (Source: "$$c+d$$", Latex: @"c+d", MathMl: SecondNumberedMathMl),
                         (Source: "$$a+b$$", Latex: @"a+b", MathMl: FirstNumberedMathMl),
                     })
            {
                var text = document.Content.Text ?? string.Empty;
                var sourceOffset = text.IndexOf(redraw.Source, StringComparison.Ordinal);
                AssertTrue(sourceOffset >= 0,
                    $"Compatibility-mode redraw fixture lost source {redraw.Source}.");
                var absoluteStart = document.Content.Start + sourceOffset;
                var source = document.Range(
                    absoluteStart,
                    absoluteStart + redraw.Source.Length);
                try
                {
                    source.Delete();
                }
                finally { Release(source); }

                var redrawSession = CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: redraw.Latex,
                    mathTypeNumberPosition: "right");
                redrawSession.SourceObjectId =
                    $"visualtex-word-vsto-range:{absoluteStart}:{absoluteStart}";
                service.InsertMathTypeOle(
                    redrawSession,
                    redraw.MathMl,
                    emfPath,
                    preserveExistingDisplayParagraphBoundary: true,
                    preserveCapturedInsertion: true);
            }
            AssertEqual(2, document.InlineShapes.Count,
                "Compatibility-mode reverse-order numbered MathType redraw insertion failed.");
            AssertEqual(2, CountMathTypePlaceRefFields(document),
                "Compatibility-mode reverse-order numbered MathType redraw lost MTPlaceRef.");
            Console.WriteLine(
                "[MATHTYPE COMPAT14 REDRAW PASS] Two dedicated display sources were replaced in reverse document order with numbered MathType equations.");
        }
        finally
        {
            Release(range);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }
}
