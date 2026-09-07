using System.Text;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordActiveMathTypeVisualTeXRoundTripCloneAcceptance(string artifactRoot)
    {
        AssertTrue(AttachActiveWord,
            "Active MathType→VisualTeX clone acceptance must attach to the user's current Word instance.");
        Directory.CreateDirectory(artifactRoot);
        var tempRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            "active-mt-vt-clone-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempRoot);
        var pngPath = Path.Combine(tempRoot, "active-clone.png");
        var svgPath = Path.Combine(tempRoot, "active-clone.svg");
        string? emfPath = null;
        Word.Application? application = null;
        Word.Document? sourceDocument = null;
        Word.Document? clone = null;
        try
        {
            WriteAcceptancePng(pngPath, "clone", 320, 112);
            File.WriteAllText(
                svgPath,
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"320\" height=\"112\" viewBox=\"0 0 320 112\"><text x=\"8\" y=\"76\" font-family=\"Cambria Math\" font-size=\"48\">x+1</text></svg>",
                new UTF8Encoding(false));
            emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 320, 112);

            application = CreateWordApplication(visible: false);
            sourceDocument = application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document is available.");
            var sourceName = sourceDocument.Name;
            var sourceFormatId = WordEquationNumbering.GetEquationNumberFormatId(sourceDocument);
            Console.WriteLine(
                $"[ACTIVE MT→VT CLONE] source='{sourceName}' shapes={sourceDocument.InlineShapes.Count} fields={sourceDocument.Fields.Count} bookmarks={sourceDocument.Bookmarks.Count} format={sourceFormatId}.");

            clone = application.Documents.Add(Visible: false);
            clone.Content.FormattedText = sourceDocument.Content.FormattedText;
            WordEquationNumbering.SetEquationNumberFormat(clone, sourceFormatId);
            clone.Activate();
            var service = new WordFormulaService(application);
            Console.WriteLine(
                $"[ACTIVE MT→VT CLONE] cloned shapes={clone.InlineShapes.Count} mathType={CountMathTypeOleShapes(clone)} visualTeX={CountVisualTeXNativeOleShapes(clone)} fields={clone.Fields.Count} bookmarks={clone.Bookmarks.Count}.");

            // The user's failure state contains the already-converted last formula as
            // VisualTeX plus the untouched earlier MathType formulas. Recreate the
            // immediate pre-failure state by converting only that VisualTeX target
            // back to MathType inside the clone.
            var restorePlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.NativeOleMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(1, restorePlan.Targets.Count,
                "Active failure clone does not contain exactly one already-converted VisualTeX target.");
            var restoreTarget = restorePlan.Targets[0];
            var restoreMathMl = KnownRoundTripMathMl(restoreTarget.Latex);
            var restorePrepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal)
            {
                [restoreTarget.Id] = new PreparedWordBulkFormula
                {
                    Run = new WordBulkRun
                    {
                        Id = restoreTarget.Id,
                        IsFormula = true,
                        Latex = restoreTarget.Latex,
                        DisplayMode = restoreTarget.DisplayMode,
                    },
                    Session = CreateSimpleFormatTargetSession(
                        restoreTarget,
                        FormulaOleContract.MathTypeOleMode,
                        restoreMathMl),
                    MathMl = restoreMathMl,
                    EmfPath = emfPath,
                },
            };
            var restoreResult = service.ApplyFormulaFormatConversionPlan(
                restorePlan,
                restorePrepared);
            Console.WriteLine(
                $"[ACTIVE MT→VT CLONE RESTORE] converted={restoreResult.FormulaCount} failed={restoreResult.FailedFormulaCount} failures={string.Join(" | ", restoreResult.Failures)}");
            AssertEqual(1, restoreResult.FormulaCount,
                "Active failure clone could not reconstruct the third MathType source.");
            AssertEqual(0, restoreResult.FailedFormulaCount,
                "Active failure clone reconstruction failed: " + string.Join(" | ", restoreResult.Failures));
            AssertEqual(3, CountMathTypeOleShapes(clone),
                "Active failure clone did not reconstruct three MathType formulas.");
            AssertEqual(0, CountVisualTeXNativeOleShapes(clone),
                "Active failure clone still contains a VisualTeX formula after reconstruction.");

            var preFailurePath = Path.Combine(artifactRoot, "Active-Document1-Reconstructed-Three-MathType.docx");
            clone.SaveAs2(preFailurePath, Word.WdSaveFormat.wdFormatXMLDocument);
            DumpActiveCloneStructure(clone, "reconstructed");

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.NativeOleMode);
            AssertEqual(3, plan.Targets.Count,
                "Active reconstructed clone did not capture three MathType formulas.");
            var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
            foreach (var target in plan.Targets)
            {
                var mathMl = target.SourceMathMl;
                if (string.IsNullOrWhiteSpace(mathMl))
                    throw new InvalidDataException("MathType source has no readable MathML: " + target.Latex);
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
                        mathMl),
                    MathMl = mathMl,
                    PngPath = pngPath,
                    EmfPath = emfPath,
                };
            }

            var result = service.ApplyFormulaFormatConversionPlan(plan, prepared);
            Console.WriteLine(
                $"[ACTIVE MT→VT CLONE APPLY] converted={result.FormulaCount} failed={result.FailedFormulaCount} failures={string.Join(" | ", result.Failures)}");
            DumpActiveCloneStructure(clone, "after-apply");
            AssertEqual(3, result.FormulaCount,
                "Active reconstructed MathType→VisualTeX clone did not convert all three formulas.");
            AssertEqual(0, result.FailedFormulaCount,
                "Active reconstructed MathType→VisualTeX clone failed: " + string.Join(" | ", result.Failures));
            AssertEqual(0, CountMathTypeOleShapes(clone),
                "Active reconstructed MathType→VisualTeX clone left MathType formulas behind.");
            AssertEqual(3, CountVisualTeXNativeOleShapes(clone),
                "Active reconstructed MathType→VisualTeX clone did not leave three VisualTeX formulas.");

            var outputPath = Path.Combine(artifactRoot, "Active-Document1-MathType-To-VisualTeX-Result.docx");
            clone.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine("[ACTIVE MT→VT CLONE] Exact current document structure roundtrip passed.");
        }
        finally
        {
            if (clone is not null)
            {
                try { clone.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(clone);
            if (sourceDocument is not null)
            {
                try { sourceDocument.Activate(); } catch { }
            }
            Release(sourceDocument);
            Release(application);
            ForceComCleanup();
            try { if (!string.IsNullOrWhiteSpace(emfPath)) File.Delete(emfPath); } catch { }
            try { File.Delete(pngPath); } catch { }
            try { File.Delete(svgPath); } catch { }
            try { Directory.Delete(tempRoot, recursive: true); } catch { }
        }
    }

    private static string KnownRoundTripMathMl(string latex)
    {
        if (latex.IndexOf(@"\mathrm{e}", StringComparison.Ordinal) >= 0)
            return "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msup><mi mathvariant=\"normal\">e</mi><mrow><mi mathvariant=\"normal\">i</mi><mi>π</mi></mrow></msup><mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>";
        if (latex.IndexOf("a^2", StringComparison.Ordinal) >= 0)
            return "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msup><mi>a</mi><mn>2</mn></msup><mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo><msup><mi>c</mi><mn>2</mn></msup></math>";
        return "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mi>x</mi><mo>+</mo><mn>1</mn></math>";
    }

    private static void DumpActiveCloneStructure(Word.Document document, string stage)
    {
        Console.WriteLine(
            $"[ACTIVE MT→VT CLONE {stage}] shapes={document.InlineShapes.Count} fields={document.Fields.Count} bookmarks={document.Bookmarks.Count} paragraphs={document.Paragraphs.Count}.");
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            Word.Range? range = null;
            try
            {
                shape = document.InlineShapes[index];
                range = shape.Range;
                string progId;
                try { progId = shape.OLEFormat.ProgID ?? string.Empty; }
                catch { progId = string.Empty; }
                var isMathType = false;
                var mathMlState = "mathMl=<not-read>";
                var metadataState = "metadata=<not-read>";
                var sideState = "side=<not-read>";
                try { isMathType = MathTypeOleInterop.IsMathTypeOle(shape); }
                catch (Exception error) { mathMlState = "isMathType-error=" + error.GetType().Name + ":" + error.Message; }
                try
                {
                    var mathMl = MathTypeOleStorage.ReadMathMl(shape);
                    mathMlState = $"mathMl={mathMl.Length}";
                    try
                    {
                        var metadata = MathTypeOleInterop.ReadMetadata(document.Application, shape, mathMl);
                        metadataState = $"metadata={metadata.DisplayMode}/{metadata.Numbered}/{metadata.Latex}";
                    }
                    catch (Exception error)
                    {
                        metadataState = "metadata-error=" + error.GetType().Name + ":" + error.Message;
                    }
                }
                catch (Exception error)
                {
                    mathMlState = "mathMl-error=" + error.GetType().Name + ":" + error.Message;
                    metadataState = "metadata=skipped";
                }
                try
                {
                    sideState = MathTypeOleInterop.TryReadDisplayNumberPosition(shape, out var side)
                        ? "side=" + side
                        : "side=<none>";
                }
                catch (Exception error)
                {
                    sideState = "side-error=" + error.GetType().Name + ":" + error.Message;
                }

                Console.WriteLine(
                    $"  shape#{index} {progId} isMathType={isMathType} range={range.Start}:{range.End} paragraph={range.Paragraphs[1].Range.Start}:{range.Paragraphs[1].Range.End} {mathMlState} {metadataState} {sideState}");
            }
            finally
            {
                Release(range);
                Release(shape);
            }
        }
    }
}
