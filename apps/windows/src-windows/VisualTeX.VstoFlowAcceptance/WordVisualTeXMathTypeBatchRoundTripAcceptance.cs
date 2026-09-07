using System.Text;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordVisualTeXMathTypeBatchRoundTripAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var tempRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            "vt-mt-vt-roundtrip-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempRoot);
        var pngPath = Path.Combine(tempRoot, "roundtrip.png");
        var svgPath = Path.Combine(tempRoot, "roundtrip.svg");
        string? emfPath = null;
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            WriteAcceptancePng(pngPath, "roundtrip", 320, 112);
            File.WriteAllText(
                svgPath,
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"320\" height=\"112\" viewBox=\"0 0 320 112\"><text x=\"8\" y=\"76\" font-family=\"Cambria Math\" font-size=\"48\">x+1</text></svg>",
                new UTF8Encoding(false));
            emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 320, 112);

            var sources = new[]
            {
                (
                    Latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                    MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>"),
                (
                    Latex: @"\mathrm{e}^{\mathrm{i}\pi}+1=0",
                    MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msup><mi mathvariant=\"normal\">e</mi><mrow><mi mathvariant=\"normal\">i</mi><mi>π</mi></mrow></msup><mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>"),
                (
                    Latex: @"a^2+b^2=c^2",
                    MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msup><mi>a</mi><mn>2</mn></msup><mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo><msup><mi>c</mi><mn>2</mn></msup></math>"),
            };
            var mathMlByLatex = sources.ToDictionary(
                source => source.Latex,
                source => source.MathMl,
                StringComparer.Ordinal);

            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DashId);
            var service = new WordFormulaService(application);

            foreach (var source in sources)
            {
                application.Selection.EndKey(Word.WdUnits.wdStory);
                var insertion = application.Selection.Range;
                try { insertion.Select(); }
                finally { Release(insertion); }
                service.InsertOle(
                    CreateSimpleVisualTeXSourceSession(source.Latex, numbered: true),
                    pngPath,
                    emfPath);
            }

            AssertEqual(3, CountVisualTeXNativeOleShapes(document),
                "VT→MT→VT setup did not create three VisualTeX OLE sources.");
            AssertEqual(3, CountInstalledVisualTeXNumberedFormulaHosts(document),
                "VT→MT→VT setup did not create three numbered VisualTeX hosts.");

            var toMathTypePlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.NativeOleMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(3, toMathTypePlan.Targets.Count,
                "VT→MT→VT first leg did not capture three VisualTeX formulas.");
            var toMathTypePrepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
            foreach (var target in toMathTypePlan.Targets)
            {
                if (!mathMlByLatex.TryGetValue(target.Latex, out var mathMl))
                    throw new InvalidDataException("No MathML fixture for VisualTeX source '" + target.Latex + "'.");
                toMathTypePrepared[target.Id] = new PreparedWordBulkFormula
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
                        FormulaOleContract.MathTypeOleMode,
                        mathMl),
                    MathMl = mathMl,
                    EmfPath = emfPath,
                };
            }
            var toMathTypeResult = service.ApplyFormulaFormatConversionPlan(
                toMathTypePlan,
                toMathTypePrepared);
            Console.WriteLine(
                $"[VT→MT→VT LEG1] converted={toMathTypeResult.FormulaCount} failed={toMathTypeResult.FailedFormulaCount} failures={string.Join(" | ", toMathTypeResult.Failures)}");
            AssertEqual(3, toMathTypeResult.FormulaCount,
                "VT→MT→VT first leg did not convert all three formulas.");
            AssertEqual(0, toMathTypeResult.FailedFormulaCount,
                "VT→MT→VT first leg failed: " + string.Join(" | ", toMathTypeResult.Failures));
            AssertEqual(3, CountMathTypeOleShapes(document),
                "VT→MT→VT first leg did not leave three MathType OLE objects.");
            AssertEqual(0, CountVisualTeXNativeOleShapes(document),
                "VT→MT→VT first leg left VisualTeX OLE objects behind.");
            AssertEqual(3, CountMathTypePlaceRefFields(document),
                "VT→MT→VT first leg did not leave three MTPlaceRef fields.");

            var intermediatePath = Path.Combine(artifactRoot, "VisualTeX-To-MathType-Intermediate.docx");
            document.SaveAs2(intermediatePath, Word.WdSaveFormat.wdFormatXMLDocument);

            Dictionary<string, PreparedWordBulkFormula> PrepareVisualTeXTargets(
                WordFormulaFormatConversionPlan conversionPlan)
            {
                var result = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
                foreach (var target in conversionPlan.Targets)
                {
                    var mathMl = target.SourceMathMl;
                    if (string.IsNullOrWhiteSpace(mathMl))
                        throw new InvalidDataException("Converted MathType source has no readable MathML for '" + target.Latex + "'.");
                    result[target.Id] = new PreparedWordBulkFormula
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
                            mathMl),
                        MathMl = mathMl,
                        PngPath = pngPath,
                        EmfPath = emfPath,
                    };
                }
                return result;
            }

            var rollbackPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.NativeOleMode);
            AssertEqual(3, rollbackPlan.Targets.Count,
                "VT→MT→VT rollback probe did not capture three MathType formulas.");
            var rollbackPrepared = PrepareVisualTeXTargets(rollbackPlan);
            var previousInjectedFailure = Environment.GetEnvironmentVariable(
                "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE");
            try
            {
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE",
                    "numbered");
                var rollbackResult = service.ApplyFormulaFormatConversionPlan(
                    rollbackPlan,
                    rollbackPrepared);
                AssertEqual(0, rollbackResult.FormulaCount,
                    "Injected MathType→VisualTeX rollback unexpectedly committed a target.");
                AssertEqual(3, rollbackResult.FailedFormulaCount,
                    "Injected MathType→VisualTeX atomic rollback did not report the full batch as unconverted.");
                AssertEqual(3, CountMathTypeOleShapes(document),
                    "Atomic MathType→VisualTeX rollback did not restore all three MathType sources.");
                AssertEqual(0, CountVisualTeXNativeOleShapes(document),
                    "Atomic MathType→VisualTeX rollback left a provisional VisualTeX target behind.");
                AssertEqual(3, CountMathTypePlaceRefFields(document),
                    "Atomic MathType→VisualTeX rollback did not restore all three MTPlaceRef fields.");
                Console.WriteLine(
                    "[VT→MT→VT ROLLBACK] Injected post-delete failure restored the complete Equation.DSMT4 + MTPlaceRef owner.");
            }
            finally
            {
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE",
                    previousInjectedFailure);
            }

            var toVisualTeXPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.NativeOleMode);
            AssertEqual(3, toVisualTeXPlan.Targets.Count,
                "VT→MT→VT second leg did not capture three MathType formulas after rollback.");
            var toVisualTeXPrepared = PrepareVisualTeXTargets(toVisualTeXPlan);

            var toVisualTeXResult = service.ApplyFormulaFormatConversionPlan(
                toVisualTeXPlan,
                toVisualTeXPrepared);
            Console.WriteLine(
                $"[VT→MT→VT LEG2] converted={toVisualTeXResult.FormulaCount} failed={toVisualTeXResult.FailedFormulaCount} failures={string.Join(" | ", toVisualTeXResult.Failures)}");
            AssertEqual(3, toVisualTeXResult.FormulaCount,
                "VT→MT→VT second leg did not convert all three formulas.");
            AssertEqual(0, toVisualTeXResult.FailedFormulaCount,
                "VT→MT→VT second leg failed: " + string.Join(" | ", toVisualTeXResult.Failures));
            AssertEqual(0, CountMathTypeOleShapes(document),
                "VT→MT→VT second leg left MathType source objects behind.");
            AssertEqual(3, CountVisualTeXNativeOleShapes(document),
                "VT→MT→VT second leg did not recreate three VisualTeX OLE objects.");
            AssertEqual(3, CountInstalledVisualTeXNumberedFormulaHosts(document),
                "VT→MT→VT second leg did not recreate three numbered VisualTeX hosts.");
            AssertEqual(0, CountMathTypePlaceRefFields(document),
                "VT→MT→VT second leg left MTPlaceRef fields behind.");
            var expectedVisualTeXNumbers = new[] { "(0-1)", "(0-2)", "(0-3)" };
            var actualVisualTeXNumbers = new List<string>();
            for (var index = 1; index <= document.InlineShapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                Word.Range? numberRange = null;
                try
                {
                    shape = document.InlineShapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata?.Numbered != true) continue;
                    numberRange = WordEquationNumbering.FindVisibleEquationNumberRange(
                        document,
                        metadata.FormulaId)
                        ?? throw new InvalidDataException(
                            "Converted VisualTeX formula lost its visible equation number range.");
                    actualVisualTeXNumbers.Add(
                        (numberRange.Text ?? string.Empty).TrimStart('\t').TrimEnd('\r', '\a'));
                }
                finally
                {
                    Release(numberRange);
                    Release(shape);
                }
            }
            AssertEqual(expectedVisualTeXNumbers.Length, actualVisualTeXNumbers.Count,
                "VT→MT→VT second leg returned the wrong number of visible VisualTeX labels.");
            for (var index = 0; index < expectedVisualTeXNumbers.Length; index++)
                AssertEqual(expectedVisualTeXNumbers[index], actualVisualTeXNumbers[index],
                    $"VT→MT→VT visible VisualTeX number {index + 1} changed ordinal or prefix.");

            var outputPath = Path.Combine(artifactRoot, "VisualTeX-MathType-VisualTeX-Roundtrip.docx");
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                "[VT→MT→VT] Three numbered VisualTeX formulas survived a whole-document VisualTeX→MathType→VisualTeX roundtrip under heading1-dash numbering.");
        }
        finally
        {
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
            try { Directory.Delete(tempRoot, recursive: true); } catch { }
        }
    }
}
