using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordLatexRedrawNumberingOrderAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            "latex-redraw-number-order-"
            + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(assetRoot);
        try
        {
            // The native VisualTeX OLE server intentionally accepts preview
            // assets only from VisualTeX's private Office temp root. Keep this
            // acceptance on the same product-safe path instead of the evidence
            // directory, otherwise Initialize() correctly returns E_ACCESSDENIED
            // before numbering is reached.
            var svgPath = Path.Combine(
                assetRoot,
                "latex-redraw-number-order.svg");
            var pngPath = Path.Combine(
                assetRoot,
                "latex-redraw-number-order.png");
            File.WriteAllText(
                svgPath,
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"96\" viewBox=\"0 0 240 96\"><text x=\"4\" y=\"64\" font-family=\"Times New Roman\" font-size=\"48\">x+1</text></svg>");
            var emfPath =
                OfficeOlePreview.CreateVectorEmfFromSvg(
                    svgPath,
                    240,
                    96);
            var pngDataUrl =
                CreatePngDataUrl(
                    "x+1",
                    240,
                    96);
            var comma =
                pngDataUrl.IndexOf(',');
            if (comma < 0)
                throw new InvalidDataException(
                    "Acceptance PNG data URL is malformed.");
            File.WriteAllBytes(
                pngPath,
                Convert.FromBase64String(
                    pngDataUrl.Substring(
                        comma + 1)));

            foreach (var mode in new[]
                     {
                         FormulaOleContract.WordOmmlMode,
                         FormulaOleContract.MathTypeOleMode,
                         FormulaOleContract.NativeOleMode,
                     })
            {
                RunWordLatexRedrawNumberingOrderScenario(
                    mode,
                    pngPath,
                    emfPath,
                    artifactRoot);
            }

            Console.WriteLine(
                "[LATEX REDRAW NUMBER ORDER PASS] Batch redraw committed final 1/2/3 numbering for OMML, MathType and VisualTeX without a manual number refresh.");
        }
        finally
        {
            try
            {
                Directory.Delete(
                    assetRoot,
                    recursive: true);
            }
            catch { }
        }
    }

    private static void RunWordLatexRedrawNumberingOrderScenario(
        string objectMode,
        string pngPath,
        string emfPath,
        string artifactRoot)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application =
                string.Equals(
                    objectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal)
                    ? CreateFreshAcceptanceAutomationWord(
                        artifactRoot)
                    : CreateWordApplication(
                        visible: false);
            document = application.Documents.Add();
            document.Activate();
            document.Content.Text =
                "$$a+b$$\r$$c+d$$\r$$x+y$$\r";
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            var service = new WordFormulaService(application);
            var plan = service.CaptureLatexRedrawPlan(
                wholeDocument: true);
            AssertEqual(
                3,
                plan.Targets.Count,
                "Number-order redraw fixture did not capture three display formulas.");
            plan.NumberDisplayFormulas = true;

            var mathMlByIndex = new[]
            {
                FirstNumberedMathMl,
                SecondNumberedMathMl,
                SimpleMathMl,
            };
            var prepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal);
            for (var index = 0; index < plan.Targets.Count; index++)
            {
                var target = plan.Targets[index];
                var mathMl = mathMlByIndex[index];
                prepared[target.Id] = new PreparedWordBulkFormula
                {
                    Run = new WordBulkRun
                    {
                        Id = target.Id,
                        IsFormula = true,
                        Latex = target.Latex,
                        DisplayMode = target.DisplayMode,
                    },
                    Session = CreateOmmlMathTypeAcceptanceSession(
                        mathMl,
                        target.DisplayMode,
                        numbered: false,
                        objectMode),
                    MathMl = mathMl,
                    PngPath = string.Equals(
                        objectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal)
                        ? pngPath
                        : null,
                    EmfPath = string.Equals(
                        objectMode,
                        FormulaOleContract.WordOmmlMode,
                        StringComparison.Ordinal)
                        ? null
                        : emfPath,
                };
            }

            var result = service.ApplyLatexRedrawPlan(
                plan,
                prepared);
            AssertEqual(
                3,
                result.FormulaCount,
                "Number-order redraw did not insert all three formulas.");

            if (string.Equals(
                    objectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal))
            {
                AssertMathTypeNumberTexts(
                    document,
                    "(1)",
                    "(2)",
                    "(3)");
                return;
            }

            var expectedSource = string.Equals(
                    objectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal)
                ? EquationReferenceSource.VisualTeX
                : EquationReferenceSource.WordOmml;
            var targets = service
                .GetCanonicalEquationReferenceTargets(document)
                .Where(target => target.Source == expectedSource)
                .OrderBy(target => target.Position)
                .ToArray();
            AssertEqual(
                3,
                targets.Length,
                "Canonical number-order redraw did not expose three numbered targets.");
            AssertEqual(
                "1",
                targets[0].NumberText,
                "First redrawn equation number is not 1.");
            AssertEqual(
                "2",
                targets[1].NumberText,
                "Second redrawn equation number is not 2.");
            AssertEqual(
                "3",
                targets[2].NumberText,
                "Third redrawn equation number is not 3.");
        }
        finally
        {
            if (document is not null)
            {
                try
                {
                    document.Close(
                        Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }
}
