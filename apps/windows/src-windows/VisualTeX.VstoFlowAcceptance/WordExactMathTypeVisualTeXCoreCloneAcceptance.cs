using System.Text;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunExactMathTypeVisualTeXCoreCloneAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourcePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_EXACT_MT_VISUALTEX_SOURCE");
        if (string.IsNullOrWhiteSpace(sourcePath))
            throw new InvalidOperationException(
                "VISUALTEX_EXACT_MT_VISUALTEX_SOURCE must point to the real MathType source document.");
        sourcePath = Path.GetFullPath(sourcePath);
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException(
                "The real MathType source document is missing.",
                sourcePath);

        var inputPath = Path.Combine(
            artifactRoot,
            "Exact-MathType-To-VisualTeX-Core-Input.docx");
        var resultPath = Path.Combine(
            artifactRoot,
            "Exact-MathType-To-VisualTeX-Core-Result.docx");
        var tempRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            "exact-mt-vt-core-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempRoot);
        var pngPath = Path.Combine(tempRoot, "preview.png");
        var svgPath = Path.Combine(tempRoot, "preview.svg");
        string? emfPath = null;
        File.Copy(sourcePath, inputPath, overwrite: true);

        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            WriteAcceptancePng(
                pngPath,
                "real-mt-vt",
                260,
                96);
            File.WriteAllText(
                svgPath,
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"96\" viewBox=\"0 0 260 96\"><text x=\"8\" y=\"68\" font-family=\"Cambria Math\" font-size=\"42\">x+1</text></svg>",
                new UTF8Encoding(false));
            emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(
                svgPath,
                260,
                96);

            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(
                inputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();

            var sourceMathType = CountMathTypeOleShapes(document);
            AssertTrue(
                sourceMathType > 0,
                "Exact MathType→VisualTeX source contains no MathType formulas.");

            var sourcePresentationSizes =
                CaptureMathTypePresentationSizesByStart(document);
            AssertEqual(
                sourceMathType,
                sourcePresentationSizes.Count,
                "Exact MathType→VisualTeX source presentation-size capture is incomplete.");

            var service = new WordFormulaService(application);
            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.NativeOleMode);
            AssertEqual(
                sourceMathType,
                plan.Targets.Count,
                "Exact MathType→VisualTeX plan did not capture every MathType source.");

            foreach (var target in plan.Targets)
            {
                if (!sourcePresentationSizes.TryGetValue(
                        target.SourceStart,
                        out var expectedFontSize))
                {
                    throw new InvalidDataException(
                        $"MathType source at {target.SourceStart} has no captured Word presentation size.");
                }
                AssertNear(
                    expectedFontSize,
                    (float)target.FontSizePt,
                    0.01f,
                    $"MathType→VisualTeX plan at {target.SourceStart} did not adopt the Word-visible font size.");
            }

            var sampleIndexes = new[]
            {
                0,
                1,
                2,
                plan.Targets.Count / 4,
                plan.Targets.Count / 2,
                plan.Targets.Count * 3 / 4,
                plan.Targets.Count - 3,
                plan.Targets.Count - 2,
                plan.Targets.Count - 1,
            }
            .Where(index => index >= 0 && index < plan.Targets.Count)
            .Distinct()
            .OrderBy(index => index)
            .ToArray();
            var sampledTargets = sampleIndexes
                .Select(index => plan.Targets[index])
                .OrderBy(target => target.SourceStart)
                .ToList();
            plan.Targets = sampledTargets;

            var prepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal);
            foreach (var target in plan.Targets)
            {
                var mathMl = target.SourceMathMl
                    ?? throw new InvalidDataException(
                        $"MathType source at {target.SourceStart} has no source MathML.");
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

            Console.WriteLine(
                $"[EXACT CORE MT→VT BEFORE] source={sourcePath}; "
                + $"mathType={sourceMathType}; sampled={sampledTargets.Count}; "
                + $"sampleIndexes=[{string.Join(",", sampleIndexes)}]; "
                + $"sizes=[{string.Join(",", sourcePresentationSizes.Values.GroupBy(value => value).OrderBy(group => group.Key).Select(group => $"{group.Key:0.##}:{group.Count()}"))}]");

            var watch = System.Diagnostics.Stopwatch.StartNew();
            var result = service.ApplyFormulaFormatConversionPlan(
                plan,
                prepared);
            watch.Stop();
            if (result.FailedFormulaCount != 0)
                throw new InvalidDataException(
                    "Exact MathType→VisualTeX conversion reported failures: "
                    + string.Join(" || ", result.Failures));

            AssertEqual(
                sampledTargets.Count,
                result.FormulaCount,
                "Exact MathType→VisualTeX sampled conversion did not report every selected source.");
            AssertEqual(
                sourceMathType - sampledTargets.Count,
                CountMathTypeOleShapes(document),
                "Exact MathType→VisualTeX sampled conversion changed the wrong number of MathType sources.");
            AssertEqual(
                sampledTargets.Count,
                CountVisualTeXNativeOleShapes(document),
                "Exact MathType→VisualTeX sampled conversion did not create one VisualTeX host per selected source.");

            var visualShapes = new List<Word.InlineShape>();
            Word.InlineShapes? shapes = null;
            try
            {
                shapes = document.InlineShapes;
                for (var index = 1; index <= shapes.Count; index++)
                {
                    Word.InlineShape? shape = null;
                    try
                    {
                        shape = shapes[index];
                        if (!WordFormulaMetadataReader.IsNativeOle(shape))
                            continue;
                        visualShapes.Add(shape);
                        shape = null;
                    }
                    finally
                    {
                        Release(shape);
                    }
                }
            }
            finally
            {
                Release(shapes);
            }

            AssertEqual(
                sampledTargets.Count,
                visualShapes.Count,
                "Exact MathType→VisualTeX metadata verification did not capture every sampled target.");

            try
            {
                for (var index = 0; index < visualShapes.Count; index++)
                {
                    var shape = visualShapes[index];
                    var metadata =
                        WordFormulaMetadataReader.TryReadEmbeddedNativeOle(
                            shape)
                        ?? throw new InvalidDataException(
                            $"VisualTeX target #{index + 1} has no embedded metadata.");
                    AssertNear(
                        10.5f,
                        (float)(metadata.FontSizePt ?? 0),
                        0.01f,
                        $"VisualTeX target #{index + 1} did not persist 10.5 pt internally.");
                    AssertNear(
                        10.5f,
                        (float)(metadata.RenderFontSizePt ?? 0),
                        0.01f,
                        $"VisualTeX target #{index + 1} did not persist 10.5 pt render metadata.");
                }

                for (var sampleIndex = 0;
                     sampleIndex < visualShapes.Count;
                     sampleIndex++)
                {
                    var shape = visualShapes[sampleIndex];
                    shape.Range.Select();
                    var reopened = service.ReadSelection();
                    AssertNear(
                        10.5f,
                        (float)(reopened.Metadata?.FontSizePt ?? 0),
                        0.01f,
                        $"Reopened VisualTeX target #{sampleIndex + 1} restored a hidden MathType font size.");
                }
            }
            finally
            {
                foreach (var shape in visualShapes)
                    Release(shape);
            }

            document.SaveAs2(
                resultPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            Console.WriteLine(
                $"[EXACT CORE MT→VT PASS] captured10.5={sourceMathType}/{sourceMathType}; "
                + $"converted={result.FormulaCount}; metadata10.5={visualShapes.Count}/{visualShapes.Count}; "
                + $"reopenSamples={visualShapes.Count}; elapsedMs={watch.ElapsedMilliseconds}; result={resultPath}");
        }
        finally
        {
            try
            {
                document?.Close(
                    Word.WdSaveOptions.wdDoNotSaveChanges);
            }
            catch { }
            try
            {
                QuitWordApplicationIfOwned(
                    application);
            }
            catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
            try
            {
                if (!string.IsNullOrWhiteSpace(emfPath))
                    File.Delete(emfPath);
            }
            catch { }
            try { File.Delete(pngPath); } catch { }
            try { File.Delete(svgPath); } catch { }
            try { Directory.Delete(tempRoot, recursive: true); } catch { }
        }
    }

    private static Dictionary<int, float>
        CaptureMathTypePresentationSizesByStart(
            Word.Document document)
    {
        var result = new Dictionary<int, float>();
        Word.InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                Word.InlineShape? shape = null;
                Word.Range? shapeRange = null;
                Word.Range? probe = null;
                Word.Font? font = null;
                try
                {
                    shape = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(shape))
                        continue;
                    shapeRange = shape.Range.Duplicate;
                    float? size = null;
                    for (var position = shapeRange.Start;
                         position < shapeRange.End;
                         position++)
                    {
                        Release(font);
                        font = null;
                        Release(probe);
                        probe = document.Range(
                            position,
                            position + 1);
                        if (!string.Equals(
                                probe.Text,
                                "\u0001",
                                StringComparison.Ordinal))
                            continue;
                        font = probe.Font;
                        if (font.Size > 0
                            && font.Size < 200)
                            size = font.Size;
                        break;
                    }
                    if (size is null)
                        throw new InvalidDataException(
                            $"MathType source at {shapeRange.Start} has no readable U+0001 Word font size.");
                    result.Add(
                        shapeRange.Start,
                        size.Value);
                }
                finally
                {
                    Release(font);
                    Release(probe);
                    Release(shapeRange);
                    Release(shape);
                }
            }
        }
        finally
        {
            Release(shapes);
        }
        return result;
    }
}
