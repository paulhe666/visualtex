using System.Diagnostics;
using System.Xml.Xsl;
using VisualTeX.WordVsto;
using Xunit;
using Xunit.Abstractions;

namespace VisualTeX.WindowsOffice.Tests;

// Timing of the owned OMML projection against Office's classic stylesheet
// route it replaced. Runs only with VT_BENCH=1; it measures, never asserts.
public sealed class OmmlProjectionBenchmark
{
    private readonly ITestOutputHelper output;

    public OmmlProjectionBenchmark(ITestOutputHelper output) => this.output = output;

    [Fact]
    public void CompareProjectionCost()
    {
        if (Environment.GetEnvironmentVariable("VT_BENCH") != "1") return;

        // First use in a process: stylesheet load + compile vs. first projection.
        var sample = new OmmlMathMlProjectionOracleTests.Generator(1, astral: false, equations: 1).Document();
        var clock = Stopwatch.StartNew();
        var oracle = ClassicOfficeReverseStylesheet.Load(ClassicOfficeReverseStylesheet.ResolvePath());
        OmmlMathMlProjectionOracleTests.OracleProject(oracle, sample);
        var oracleFirst = clock.Elapsed.TotalMilliseconds;
        clock.Restart();
        OmmlMathMlProjection.Project(sample);
        var ownFirst = clock.Elapsed.TotalMilliseconds;
        output.WriteLine($"first use: office stylesheet {oracleFirst:F1} ms, owned projection {ownFirst:F1} ms");

        var documents = Enumerable.Range(1, 3000)
            .Select(seed => new OmmlMathMlProjectionOracleTests.Generator(seed, astral: false, equations: 1).Document())
            .ToArray();
        // Warm both paths.
        foreach (var document in documents.Take(200))
        {
            OmmlMathMlProjectionOracleTests.OracleProject(oracle, document);
            OmmlMathMlProjection.Project(document);
        }

        var buckets = new[] { (0, 1000, "under 1 KB"), (1000, 4000, "1-4 KB"), (4000, 16000, "4-16 KB"), (16000, int.MaxValue, "over 16 KB") };
        var rows = documents.Select(document => (document.Length, Oracle: Time(() => OmmlMathMlProjectionOracleTests.OracleProject(oracle, document)),
            Own: Time(() => OmmlMathMlProjection.Project(document)),
            Full: Time(() => WordOmmlConverter.TransformOmmlToMathMl(document, false)))).ToArray();
        foreach (var (low, high, label) in buckets)
        {
            var bucket = rows.Where(row => row.Length >= low && row.Length < high).ToArray();
            if (bucket.Length == 0) continue;
            output.WriteLine($"{label,-11} n={bucket.Length,4}  office median {Median(bucket.Select(r => r.Oracle)):F3} ms"
                + $"  owned median {Median(bucket.Select(r => r.Own)):F3} ms"
                + $"  whole read median {Median(bucket.Select(r => r.Full)):F3} ms");
        }
        output.WriteLine($"all        n={rows.Length}  office total {rows.Sum(r => r.Oracle):F0} ms  owned total {rows.Sum(r => r.Own):F0} ms"
            + $"  whole read total {rows.Sum(r => r.Full):F0} ms");
    }

    [Fact]
    public void CompareInsertionProjectionCost()
    {
        if (Environment.GetEnvironmentVariable("VT_BENCH") != "1") return;

        var sample = new MathMlOmmlProjectionOracleTests.Generator(1, astral: false).Document();
        var clock = Stopwatch.StartNew();
        var oracle = ClassicOfficeForwardStylesheet.Load(ClassicOfficeForwardStylesheet.ResolvePath());
        MathMlOmmlProjectionOracleTests.OracleProject(oracle, sample);
        var oracleFirst = clock.Elapsed.TotalMilliseconds;
        clock.Restart();
        MathMlOmmlProjection.Project(sample);
        var ownFirst = clock.Elapsed.TotalMilliseconds;
        output.WriteLine($"insertion first use: office stylesheet {oracleFirst:F1} ms, owned projection {ownFirst:F1} ms");

        var documents = Enumerable.Range(1, 3000)
            .Select(seed => new MathMlOmmlProjectionOracleTests.Generator(seed, astral: false).Document())
            .ToArray();
        foreach (var document in documents.Take(200))
        {
            MathMlOmmlProjectionOracleTests.OracleProject(oracle, document);
            MathMlOmmlProjection.Project(document);
        }
        var buckets = new[] { (0, 1000, "under 1 KB"), (1000, 4000, "1-4 KB"), (4000, 16000, "4-16 KB"), (16000, int.MaxValue, "over 16 KB") };
        var rows = documents.Select(document => (document.Length,
            Oracle: Time(() => MathMlOmmlProjectionOracleTests.OracleProject(oracle, document)),
            Own: Time(() => MathMlOmmlProjection.Project(document)))).ToArray();
        foreach (var (low, high, label) in buckets)
        {
            var bucket = rows.Where(row => row.Length >= low && row.Length < high).ToArray();
            if (bucket.Length == 0) continue;
            output.WriteLine($"insertion {label,-11} n={bucket.Length,4}  office median {Median(bucket.Select(r => r.Oracle)):F3} ms"
                + $"  owned median {Median(bucket.Select(r => r.Own)):F3} ms");
        }
        output.WriteLine($"insertion all n={rows.Length}  office total {rows.Sum(r => r.Oracle):F0} ms  owned total {rows.Sum(r => r.Own):F0} ms");
    }

    private static double Time(Action action)
    {
        var clock = Stopwatch.StartNew();
        action();
        return clock.Elapsed.TotalMilliseconds;
    }

    private static double Median(IEnumerable<double> values)
    {
        var sorted = values.OrderBy(value => value).ToArray();
        return sorted[sorted.Length / 2];
    }
}
