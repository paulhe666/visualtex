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

    private static void RunMathTypeNativePreviewPrivateBoundaryAcceptance(string artifactRoot)
    {
        var baseline = SnapshotInteractiveMathTypeIdentities();
        AssertTrue(MathTypeNativePreviewRenderer.IsAvailable, "The private render service is not running.");
        // An already-open native editor is valid. Rendering must not start or
        // terminate an engine in the interactive session.
        RunMathTypeNativePreviewSinglePerformanceAcceptance(artifactRoot);
        var after = SnapshotInteractiveMathTypeIdentities();
        AssertTrue(baseline.Count == after.Count && baseline.All(pair =>
                after.TryGetValue(pair.Key, out var created) && created == pair.Value),
            "Private rendering changed the interactive MathType process identities.");
        Console.WriteLine("[MathType private boundary] Native rendering preserved interactive process identities.");
    }

    private static Dictionary<int, long> SnapshotInteractiveMathTypeIdentities()
    {
        using var current = System.Diagnostics.Process.GetCurrentProcess();
        var result = new Dictionary<int, long>();
        foreach (var name in new[] { "MathType", "MathTypeLib" })
        foreach (var process in System.Diagnostics.Process.GetProcessesByName(name))
        {
            using (process)
            {
                try
                {
                    if (process.SessionId == current.SessionId)
                        result.Add(process.Id, process.StartTime.ToUniversalTime().Ticks);
                }
                catch (InvalidOperationException) { }
            }
        }
        return result;
    }

    private static HashSet<int> SnapshotInteractiveMathTypeProcessIds() =>
        SnapshotInteractiveMathTypeIdentities().Keys.ToHashSet();

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
            ["nuclear-prescripts"] = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msubsup><mi/><mrow><mn>27</mn></mrow><mrow><mn>59</mn></mrow></msubsup><mi>C</mi><mi>o</mi><msubsup><mo>+</mo><mrow><mn>0</mn></mrow><mrow><mn>1</mn></mrow></msubsup><mi>n</mi><msubsup><mo stretchy=\"false\">→</mo><mrow><mn>27</mn></mrow><mrow><mn>60</mn></mrow></msubsup><mi>C</mi><mi>o</mi><mo>+</mo><mi>γ</mi></math>",
            ["east-asian-text"] = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msubsup><mi/><mn>6</mn><mn>12</mn></msubsup><mi>C</mi><mo>+</mo><mtext>中文氢燃烧</mtext></math>",
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
