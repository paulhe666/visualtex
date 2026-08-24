using VisualTeX.WordVsto;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunMathTypeNativePreviewSinglePerformanceAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        const string mathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\"><mfrac><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow><mrow><mi>y</mi><mo>−</mo><mn>2</mn></mrow></mfrac></math>";
        var generated = MathTypeMtefCodec.CreateEquationNative(mathMl, inline: true);
        var timings = new List<long>();
        for (var iteration = 1; iteration <= 3; iteration++)
        {
            var input = new Dictionary<string, byte[]>(StringComparer.Ordinal)
            {
                ["single"] = generated.Mtef,
            };
            var watch = System.Diagnostics.Stopwatch.StartNew();
            var success = MathTypeNativePreviewRenderer.TryRenderBatch(
                input,
                artifactRoot,
                out var previews);
            watch.Stop();
            timings.Add(watch.ElapsedMilliseconds);
            try
            {
                if (!success || !previews.TryGetValue("single", out var preview))
                    throw new InvalidDataException(
                        $"Single-formula native preview failed on iteration {iteration}.");
                Console.WriteLine(
                    $"[MathType native single] iteration={iteration} elapsedMs={watch.ElapsedMilliseconds} "
                    + $"width={preview.WidthPt} height={preview.HeightPt} baseline={preview.WordPosition}");
            }
            finally
            {
                foreach (var preview in previews.Values)
                    preview.Dispose();
            }
        }
        Console.WriteLine(
            "[MathType native single] timingsMs=" + string.Join(",", timings));
    }

    private static void RunMathTypeNativePreviewSharedLifecycleAcceptance()
    {
        AssertTrue(
            MathTypeNativePreviewRenderer.IsMathTypeRpcCommandLine(
                "\"C:\\Program Files (x86)\\MathType\\MathType.exe\" -mtrpc"),
            "MathType helper classifier did not recognize the native -mtrpc command line.");
        AssertTrue(
            !MathTypeNativePreviewRenderer.IsMathTypeRpcCommandLine(
                "\"C:\\Program Files (x86)\\MathType\\MathType.exe\""),
            "MathType helper classifier would treat a normal user launch as an owned RPC helper.");
        AssertTrue(
            !MathTypeNativePreviewRenderer.IsMathTypeRpcCommandLine(
                "\"C:\\Program Files (x86)\\MathType\\MathType.exe\" -Embedding"),
            "MathType helper classifier would treat an OLE -Embedding server as the preview RPC helper.");

        var baseline = SnapshotMathTypeProcessIds();
        AssertEqual(0, baseline.Count,
            "Shared MathType preview lifecycle acceptance requires no MathType process at start.");

        MathTypeNativePreviewRenderer.AcquireSharedSession();
        MathTypeNativePreviewRenderer.ReleaseSharedSession();

        // The cold MathPage startup continues off-thread after Word has already
        // disconnected. Give that intentionally non-blocking prewarm enough time
        // to finish, then require its completion-side cleanup to remove every
        // helper that was not present in the original baseline.
        Thread.Sleep(7_000);
        var remaining = SnapshotMathTypeProcessIds().Except(baseline).ToArray();
        AssertEqual(0, remaining.Length,
            "Closing Word during MathType preview prewarm left a windowless MathType helper behind: "
            + string.Join(",", remaining));
        Console.WriteLine(
            "[MathType shared lifecycle] Immediate session release during cold prewarm left no MathType process behind.");
    }

    private static void RunMathTypeNativePreviewComplexAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var cases = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["partial-mi"] = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mfrac><mrow><msup><mi>∂</mi><mn>2</mn></msup><mi>u</mi></mrow><mrow><mi>∂</mi><msup><mi>x</mi><mn>2</mn></msup></mrow></mfrac><mo>+</mo><mfrac><mrow><msup><mi>∂</mi><mn>2</mn></msup><mi>u</mi></mrow><mrow><mi>∂</mi><msup><mi>y</mi><mn>2</mn></msup></mrow></mfrac><mo>=</mo><mn>0</mn></math>",
            ["partial-mo"] = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mfrac><mrow><msup><mo>∂</mo><mn>2</mn></msup><mi>u</mi></mrow><mrow><mo>∂</mo><msup><mi>x</mi><mn>2</mn></msup></mrow></mfrac><mo>+</mo><mfrac><mrow><msup><mo>∂</mo><mn>2</mn></msup><mi>u</mi></mrow><mrow><mo>∂</mo><msup><mi>y</mi><mn>2</mn></msup></mrow></mfrac><mo>=</mo><mn>0</mn></math>",
            ["contour-simple"] = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msub><mo data-mjx-texclass=\"OP\">∮</mo><mi>C</mi></msub><mi>f</mi><mi>d</mi><mi>z</mi></math>",
            ["contour-fraction"] = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msub><mo data-mjx-texclass=\"OP\">∮</mo><mi>C</mi></msub><mfrac><mrow><mi>f</mi><mo stretchy=\"false\">(</mo><mi>z</mi><mo stretchy=\"false\">)</mo></mrow><mrow><mi>z</mi><mo>−</mo><msub><mi>z</mi><mn>0</mn></msub></mrow></mfrac><mi>d</mi><mi>z</mi></math>",
            ["contour-full"] = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mfrac><mn>1</mn><mrow><mn>2</mn><mi>π</mi><mi>i</mi></mrow></mfrac><msub><mo data-mjx-texclass=\"OP\">∮</mo><mi>C</mi></msub><mfrac><mrow><mi>f</mi><mo stretchy=\"false\">(</mo><mi>z</mi><mo stretchy=\"false\">)</mo></mrow><mrow><mi>z</mi><mo>−</mo><msub><mi>z</mi><mn>0</mn></msub></mrow></mfrac><mspace width=\"0.167em\"/><mi>d</mi><mi>z</mi><mo>=</mo><mi>f</mi><mo stretchy=\"false\">(</mo><msub><mi>z</mi><mn>0</mn></msub><mo stretchy=\"false\">)</mo></math>",
        };

        foreach (var pair in cases)
        {
            var generated = MathTypeMtefCodec.CreateEquationNative(pair.Value, inline: false);
            File.WriteAllBytes(Path.Combine(artifactRoot, pair.Key + ".mtef"), generated.Mtef);
            var success = MathTypeNativePreviewRenderer.TryRender(
                generated.Mtef,
                artifactRoot,
                out var preview);
            using (preview)
            {
                Console.WriteLine($"{pair.Key}: success={success}; width={preview.WidthPt}; height={preview.HeightPt}; baseline={preview.WordPosition}");
                if (!success)
                    throw new InvalidDataException($"MathType native preview failed for {pair.Key}.");
                File.Copy(preview.WmfPath, Path.Combine(artifactRoot, pair.Key + ".wmf"), true);
            }
        }

    }
}
