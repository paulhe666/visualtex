using VisualTeX.WordVsto;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
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
