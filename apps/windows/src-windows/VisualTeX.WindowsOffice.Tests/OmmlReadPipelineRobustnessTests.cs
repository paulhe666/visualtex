using VisualTeX.WordVsto;
using Xunit;
using Xunit.Abstractions;

namespace VisualTeX.WindowsOffice.Tests;

// Any well-formed OMML a user's document holds must read back: to MathML, to
// editable LaTeX, to a content fingerprint and to MathType, without an error.
public sealed class OmmlReadPipelineRobustnessTests
{
    private readonly ITestOutputHelper output;

    public OmmlReadPipelineRobustnessTests(ITestOutputHelper output) => this.output = output;

    [Fact]
    public void EveryGeneratedFormulaReadsBackWithoutError()
    {
        var count = int.TryParse(Environment.GetEnvironmentVariable("VT_PROJECTION_CASES"), out var n) ? n : 3000;
        var failures = new Dictionary<string, (int Count, int Seed, string Omml)>();
        void Record(string stage, Exception error, int seed, string omml)
        {
            var frames = string.Join(" <- ", (error.StackTrace ?? string.Empty)
                .Split('\n').Select(line => line.Trim()).Where(line => line.Contains("VisualTeX.")).Take(3));
            var key = stage + ": " + error.GetType().Name + ": " + error.Message + "\n  at " + frames;
            failures[key] = failures.TryGetValue(key, out var seen)
                ? (seen.Count + 1, seen.Seed, seen.Omml)
                : (1, seed, omml);
        }

        for (var seed = 1; seed <= count; seed++)
        {
            var omml = new OmmlMathMlProjectionOracleTests.Generator(seed, astral: true, equations: 1).Document();
            var display = seed % 2 == 0;
            string mathMl;
            try { mathMl = WordOmmlConverter.TransformOmmlToMathMl(omml, display); }
            catch (Exception error) { Record("MathML", error, seed, omml); continue; }
            try { MathMlToLatexConverter.Convert(mathMl); }
            catch (Exception error) { Record("LaTeX", error, seed, omml); }
            try { WordOmmlConverter.ComputeOmmlFingerprint(omml); }
            catch (Exception error) { Record("Fingerprint", error, seed, omml); }
            try { MathTypeMtefCodec.CreateEquationNative(mathMl, inline: !display); }
            catch (Exception error) { Record("MathType", error, seed, omml); }
        }

        foreach (var failure in failures.OrderByDescending(item => item.Value.Count))
            output.WriteLine($"{failure.Value.Count}x {failure.Key}\n  first seed {failure.Value.Seed}");
        Assert.True(failures.Count == 0, $"{failures.Count} distinct read failures; see test output.");
    }
}
