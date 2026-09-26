using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordLatexRedrawPartialFailureAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            "latex-redraw-partial-failure-"
            + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(assetRoot);
        try
        {
            var svgPath = Path.Combine(
                assetRoot,
                "latex-redraw-partial-failure.svg");
            var pngPath = Path.Combine(
                assetRoot,
                "latex-redraw-partial-failure.png");
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
            var comma = pngDataUrl.IndexOf(',');
            if (comma < 0)
                throw new InvalidDataException(
                    "Partial-failure acceptance PNG data URL is malformed.");
            File.WriteAllBytes(
                pngPath,
                Convert.FromBase64String(
                    pngDataUrl.Substring(comma + 1)));

            foreach (var objectMode in new[]
                     {
                         FormulaOleContract.WordOmmlMode,
                         FormulaOleContract.MathTypeOleMode,
                         FormulaOleContract.NativeOleMode,
                     })
            {
                RunWordLatexRedrawPartialFailureScenario(
                    objectMode,
                    pngPath,
                    emfPath,
                    artifactRoot);
            }

            Console.WriteLine(
                "[LATEX REDRAW PARTIAL FAILURE PASS] OMML, MathType and VisualTeX each rolled back one failed target to exact LaTeX while later formulas continued.");
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

    private static void RunWordLatexRedrawPartialFailureScenario(
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
                "$a+b$ text $c+d$ text $x+y$\r";

            var service = new WordFormulaService(application);
            var plan = service.CaptureLatexRedrawPlan(
                wholeDocument: true);
            AssertEqual(
                3,
                plan.Targets.Count,
                "Partial-failure redraw fixture did not capture three formulas.");

            var middleTarget = plan.Targets[1];
            var middleSource =
                document.Range(
                    middleTarget.AbsoluteStart,
                    middleTarget.AbsoluteEnd)
                    .Text;

            var prepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal);
            for (var index = 0; index < plan.Targets.Count; index++)
            {
                var target = plan.Targets[index];
                var failThisTarget = index == 1;
                prepared[target.Id] =
                    new PreparedWordBulkFormula
                    {
                        Run = new WordBulkRun
                        {
                            Id = target.Id,
                            IsFormula = true,
                            Latex = target.Latex,
                            DisplayMode = target.DisplayMode,
                        },
                        Session =
                            CreateOmmlMathTypeAcceptanceSession(
                                SimpleMathMl,
                                target.DisplayMode,
                                numbered: false,
                                objectMode),
                        // Inject the failure after source deletion, inside the
                        // per-target Word transaction. OMML and MathType reject a
                        // missing MathML payload; VisualTeX rejects a missing PNG
                        // preview. Successful neighbors use valid assets.
                        MathMl =
                            failThisTarget
                            && !string.Equals(
                                objectMode,
                                FormulaOleContract.NativeOleMode,
                                StringComparison.Ordinal)
                                ? null
                                : SimpleMathMl,
                        PngPath =
                            string.Equals(
                                objectMode,
                                FormulaOleContract.NativeOleMode,
                                StringComparison.Ordinal)
                                ? failThisTarget
                                    ? null
                                    : pngPath
                                : null,
                        EmfPath =
                            string.Equals(
                                objectMode,
                                FormulaOleContract.WordOmmlMode,
                                StringComparison.Ordinal)
                                ? null
                                : emfPath,
                    };
            }

            var result =
                service.ApplyLatexRedrawPlan(
                    plan,
                    prepared);

            AssertEqual(
                2,
                result.FormulaCount,
                $"{objectMode}: partial-failure redraw did not continue after the failed middle formula.");
            AssertEqual(
                1,
                result.Failures.Count,
                $"{objectMode}: partial-failure redraw did not report exactly one failed target.");
            AssertEqual(
                middleTarget.Id,
                result.Failures[0].TargetId,
                $"{objectMode}: partial-failure redraw reported the wrong failed target.");

            if (string.Equals(
                    objectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                AssertEqual(
                    2,
                    document.OMaths.Count,
                    "OMML partial-failure redraw created the wrong successful formula count.");
            }
            else
            {
                AssertEqual(
                    2,
                    document.InlineShapes.Count,
                    $"{objectMode}: partial-failure redraw created the wrong successful OLE count.");
            }

            var remainingText = document.Content.Text ?? string.Empty;
            if (string.IsNullOrEmpty(middleSource)
                || remainingText.IndexOf(
                    middleSource,
                    StringComparison.Ordinal) < 0)
            {
                throw new InvalidDataException(
                    $"{objectMode}: partial-failure redraw did not preserve the failed formula's exact LaTeX source.");
            }
            if (remainingText.IndexOf(
                    "$a+b$",
                    StringComparison.Ordinal) >= 0
                || remainingText.IndexOf(
                    "$x+y$",
                    StringComparison.Ordinal) >= 0)
            {
                throw new InvalidDataException(
                    $"{objectMode}: partial-failure redraw left a successfully redrawn formula as LaTeX.");
            }

            Console.WriteLine(
                $"[LATEX REDRAW PARTIAL FAILURE {objectMode}] successes={result.FormulaCount}; failures={result.Failures.Count}; preserved='{middleSource}'.");
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
