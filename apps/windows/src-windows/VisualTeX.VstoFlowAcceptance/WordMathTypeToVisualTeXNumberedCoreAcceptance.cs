using System.Text;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeToVisualTeXNumberedCoreAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var tempRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(tempRoot);
        var svgPath = Path.Combine(tempRoot, $"{Guid.NewGuid():N}.svg");
        var pngPath = Path.Combine(tempRoot, $"{Guid.NewGuid():N}.png");
        string? emfPath = null;
        Word.Application? application = null;
        Word.Document? document = null;
        var previousPreviewDisable = Environment.GetEnvironmentVariable(
            "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW");
        var previousPerf = Environment.GetEnvironmentVariable(
            "VISUALTEX_NUMBERED_PERF_TRACE");
        try
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                "1");
            Environment.SetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE", "1");
            File.WriteAllText(
                svgPath,
                CreateFontAcceptanceSvg("Times New Roman", "SimSun"),
                new UTF8Encoding(false));
            WriteAcceptancePng(pngPath, "x=1", 260, 96);
            emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 260, 96);

            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);
            var service = new WordFormulaService(application);

            var sources = new[]
            {
                (Latex: @"\hbar\omega+1", MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mi>ℏ</mi><mi>ω</mi><mo>+</mo><mn>1</mn></math>"),
                (Latex: @"\int_0^1 x^2\,dx", MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup><msup><mi>x</mi><mn>2</mn></msup><mi>d</mi><mi>x</mi></math>"),
            };

            for (var index = 0; index < sources.Length; index++)
            {
                application.Selection.EndKey(Word.WdUnits.wdStory);
                var insertion = application.Selection.Range;
                var session = CreateSimpleVisualTeXSourceSession(
                    sources[index].Latex,
                    numbered: true);
                session.ObjectMode = FormulaOleContract.MathTypeOleMode;
                session.SourceDocumentId = document.FullName;
                session.SourceObjectId = WordRangeReference(insertion.Start, insertion.End);
                session.MathTypeNumberPosition = "right";
                Release(insertion);
                service.InsertMathTypeOle(
                    session,
                    sources[index].MathMl,
                    emfPath,
                    updateCreatedMathTypeNumberFields: true);
                application.Selection.EndKey(Word.WdUnits.wdStory);
                application.Selection.TypeParagraph();
            }

            AssertEqual(2, CountMathTypeOleShapes(document),
                "Numbered MathType→VisualTeX core fixture did not create two Equation.DSMT4 sources.");
            AssertEqual(2, CountMathTypePlaceRefFields(document),
                "Numbered MathType→VisualTeX core fixture did not create two MTPlaceRef fields.");

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.NativeOleMode);
            AssertEqual(2, plan.Targets.Count,
                "Numbered MathType→VisualTeX core capture did not find two sources.");
            AssertTrue(plan.Targets.All(target => target.Numbered),
                "Numbered MathType→VisualTeX core capture lost numbered state.");

            var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
            for (var index = 0; index < plan.Targets.Count; index++)
            {
                var target = plan.Targets[index];
                var source = sources[index];
                prepared[target.Id] = new PreparedWordBulkFormula
                {
                    Run = new WordBulkRun
                    {
                        Id = target.Id,
                        IsFormula = true,
                        Latex = target.Latex,
                        DisplayMode = target.DisplayMode,
                    },
                    Session = CreateSimpleFormatTargetSession(
                        target,
                        FormulaOleContract.NativeOleMode,
                        source.MathMl),
                    MathMl = source.MathMl,
                    PngPath = pngPath,
                    EmfPath = emfPath,
                };
            }

            var result = service.ApplyFormulaFormatConversionPlan(plan, prepared);
            Console.WriteLine(
                $"[MT→VT NUMBERED CORE] converted={result.FormulaCount} failed={result.FailedFormulaCount} failures={string.Join(" | ", result.Failures)}");
            AssertEqual(2, result.FormulaCount,
                "Numbered MathType→VisualTeX core conversion did not convert both formulas.");
            AssertEqual(0, result.FailedFormulaCount,
                "Numbered MathType→VisualTeX core conversion failed: "
                + string.Join(" | ", result.Failures));
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Numbered MathType→VisualTeX core conversion left MathType sources behind.");
            AssertEqual(2, CountVisualTeXNativeOleShapes(document),
                "Numbered MathType→VisualTeX core conversion did not create two VisualTeX OLE targets.");
            AssertEqual(2, CountInstalledVisualTeXNumberedFormulaHosts(document),
                "Numbered MathType→VisualTeX core conversion lost numbered VisualTeX hosts.");
            AssertEqual(0, CountMathTypePlaceRefFields(document),
                "Numbered MathType→VisualTeX core conversion left MTPlaceRef fields behind.");

            var outputPath = Path.Combine(
                artifactRoot,
                "MathType-To-VisualTeX-Numbered-Core.docx");
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Reopened numbered MathType→VisualTeX core document restored MathType sources.");
            AssertEqual(2, CountVisualTeXNativeOleShapes(document),
                "Reopened numbered MathType→VisualTeX core document lost VisualTeX targets.");
            AssertEqual(2, CountInstalledVisualTeXNumberedFormulaHosts(document),
                "Reopened numbered MathType→VisualTeX core document lost numbered state.");
            Console.WriteLine(
                "Numbered MathType→VisualTeX core acceptance passed: genuine MTPlaceRef sources converted through the production batch path to two numbered VisualTeX OLE hosts and survived save/reopen.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                previousPreviewDisable);
            Environment.SetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE", previousPerf);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            try { if (!string.IsNullOrWhiteSpace(emfPath)) File.Delete(emfPath); } catch { }
            try { File.Delete(pngPath); } catch { }
            try { File.Delete(svgPath); } catch { }
        }
    }
}
