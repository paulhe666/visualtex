using System.Text.Json;
using VisualTeX.WordVsto;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class StandaloneRootProbe
    {
        public string Name { get; set; } = string.Empty;
        public bool Inline { get; set; }
        public string Latex { get; set; } = string.Empty;
        public string ExpectedSignature { get; set; } = string.Empty;
        public string ActualSignature { get; set; } = string.Empty;
        public string ReadbackLatex { get; set; } = string.Empty;
        public int EquationNativeBytes { get; set; }
        public int CompoundFileBytes { get; set; }
    }

    private static void RunMathTypeStandaloneRootRegressionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);

        var cases = new[]
        {
            (
                Name: "percent-before-script",
                Latex: @"100%\mathrm{e}^{\mathrm{i}\pi}+1=0",
                MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                    + "<mn>100</mn><mtext>%</mtext>"
                    + "<msup><mi mathvariant=\"normal\">e</mi><mrow><mi mathvariant=\"normal\">i</mi><mi>π</mi></mrow></msup>"
                    + "<mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>"
            ),
            (
                Name: "bold-italic-ddot-eta",
                Latex: @"\boldsymbol{\ddot{\eta}}",
                MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                    + "<mover accent=\"true\"><mi mathvariant=\"bold-italic\">η</mi><mo>¨</mo></mover></math>"
            ),
        };

        var rows = new List<StandaloneRootProbe>();
        foreach (var item in cases)
        foreach (var inline in new[] { false, true })
        {
            var expected = MathTypeMtefCodec.SemanticSignature(item.MathMl);
            var prepared = WordFormulaService.PrepareStandaloneMathTypeOleData(
                item.MathMl,
                inline,
                fontSizePt: 12,
                item.Latex);
            var actualMathMl = MathTypeOleStorage.ReadMathMl(prepared.CompoundFile);
            var actual = MathTypeMtefCodec.SemanticSignature(actualMathMl);
            AssertEqual(
                expected,
                actual,
                $"{item.Name} standalone MathType semantic signature changed.");
            AssertEqual(
                expected,
                prepared.SemanticSignature,
                $"{item.Name} product preflight returned a different semantic signature.");

            var latex = MathMlToLatexConverter.Convert(actualMathMl);
            if (item.Name == "percent-before-script")
            {
                AssertTrue(
                    latex.IndexOf(@"100\text{\%}\mathrm{e}^{\mathrm{i}\pi}", StringComparison.Ordinal) >= 0
                    || latex.IndexOf(@"100\%\mathrm{e}^{\mathrm{i}\pi}", StringComparison.Ordinal) >= 0,
                    "Percent-before-script readback did not keep percent separate from the scripted base.");
            }
            else
            {
                AssertTrue(
                    latex.IndexOf(@"\ddot{\boldsymbol{\eta}}", StringComparison.Ordinal) >= 0
                    || latex.IndexOf(@"\boldsymbol{\ddot{\eta}}", StringComparison.Ordinal) >= 0,
                    "Bold italic eta lost bold+italic semantics under double-dot embellishment.");
            }

            rows.Add(new StandaloneRootProbe
            {
                Name = item.Name,
                Inline = inline,
                Latex = item.Latex,
                ExpectedSignature = expected,
                ActualSignature = actual,
                ReadbackLatex = latex,
                EquationNativeBytes = prepared.Generated.EquationNative.Length,
                CompoundFileBytes = prepared.CompoundFile.Length,
            });

            Console.WriteLine(
                $"PASS standalone-root {item.Name} inline={inline} latex=[{latex}]");
        }

        File.WriteAllText(
            Path.Combine(artifactRoot, "summary.json"),
            JsonSerializer.Serialize(
                new
                {
                    Time = DateTimeOffset.Now,
                    Passed = rows.Count,
                    Rows = rows,
                },
                new JsonSerializerOptions { WriteIndented = true }));

        Console.WriteLine(
            $"MathType standalone root regression passed: {rows.Count}/{rows.Count}.");
    }
}
