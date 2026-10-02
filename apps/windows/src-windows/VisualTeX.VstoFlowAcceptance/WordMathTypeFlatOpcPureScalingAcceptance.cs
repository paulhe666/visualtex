using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class FlatOpcPureScalingRow
    {
        public int Index { get; set; }
        public int MathMlBytes { get; set; }
        public int CompoundBytes { get; set; }
        public int WmfBytes { get; set; }
        public int FlatOpcChars { get; set; }
        public double ElapsedMilliseconds { get; set; }
    }

    private static void RunWordMathTypeFlatOpcPureScalingAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "flatopc-pure-scaling.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"160\" height=\"48\" viewBox=\"0 0 160 48\"><text x=\"4\" y=\"34\" font-family=\"Cambria Math\" font-size=\"28\">x₁ = 1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 160, 48);
        var previewWmf = MathTypeWordOpenXml.ConvertEnhancedMetafileToPlaceableWmf(
            emfPath,
            widthPt: 120,
            heightPt: 36);

        const int count = 60;
        var rows = new List<FlatOpcPureScalingRow>(count);
        for (var index = 1; index <= count; index++)
        {
            var mathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msub><mi>x</mi><mn>"
                + index.ToString(CultureInfo.InvariantCulture)
                + "</mn></msub><mo>=</mo><mn>"
                + index.ToString(CultureInfo.InvariantCulture)
                + "</mn></math>";
            var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(
                mathMl,
                inline: true,
                fontSizePt: 12d);
            var compound = MathTypeOleStorage.CreateStandaloneCompoundFile(generated);
            var bookmark = WordFormulaIdentityStore.MathTypeBookmarkName(Guid.NewGuid().ToString("D"));
            var watch = Stopwatch.StartNew();
            var xml = MathTypeWordOpenXml.CreateWithPlaceableWmf(
                compound,
                previewWmf,
                widthPt: 120,
                heightPt: 36,
                display: false,
                numberTemplate: null,
                mathTypeNumberPosition: "right",
                locatorBookmarkName: bookmark);
            watch.Stop();
            rows.Add(new FlatOpcPureScalingRow
            {
                Index = index,
                MathMlBytes = System.Text.Encoding.UTF8.GetByteCount(mathMl),
                CompoundBytes = compound.Length,
                WmfBytes = previewWmf.Length,
                FlatOpcChars = xml.Length,
                ElapsedMilliseconds = watch.Elapsed.TotalMilliseconds,
            });
        }

        static double Median(IEnumerable<double> values)
        {
            var ordered = values.OrderBy(value => value).ToArray();
            if (ordered.Length == 0) return 0;
            var middle = ordered.Length / 2;
            return ordered.Length % 2 == 0
                ? (ordered[middle - 1] + ordered[middle]) / 2d
                : ordered[middle];
        }

        static double Slope(IReadOnlyList<double> values)
        {
            var meanX = (values.Count + 1) / 2d;
            var meanY = values.Average();
            double numerator = 0;
            double denominator = 0;
            for (var index = 0; index < values.Count; index++)
            {
                var x = index + 1;
                numerator += (x - meanX) * (values[index] - meanY);
                denominator += (x - meanX) * (x - meanX);
            }
            return denominator > 0 ? numerator / denominator : 0;
        }

        var elapsed = rows.Select(row => row.ElapsedMilliseconds).ToArray();
        var report = new
        {
            Time = DateTimeOffset.Now,
            Count = rows.Count,
            First10MedianMs = Median(elapsed.Take(10)),
            Last10MedianMs = Median(elapsed.Skip(elapsed.Length - 10)),
            SlopeMsPerFormula = Slope(elapsed),
            TotalMs = elapsed.Sum(),
            MinCompoundBytes = rows.Min(row => row.CompoundBytes),
            MaxCompoundBytes = rows.Max(row => row.CompoundBytes),
            MinFlatOpcChars = rows.Min(row => row.FlatOpcChars),
            MaxFlatOpcChars = rows.Max(row => row.FlatOpcChars),
            Rows = rows,
        };
        var reportPath = Path.Combine(artifactRoot, "flatopc-pure-scaling.json");
        File.WriteAllText(
            reportPath,
            JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine(
            $"[MATHTYPE FLATOPC PURE SCALING] count={rows.Count} "
            + $"first10={report.First10MedianMs:0.000}ms "
            + $"last10={report.Last10MedianMs:0.000}ms "
            + $"slope={report.SlopeMsPerFormula:0.000}ms/formula "
            + $"total={report.TotalMs:0.0}ms "
            + $"compound={report.MinCompoundBytes}-{report.MaxCompoundBytes}B "
            + $"flatopc={report.MinFlatOpcChars}-{report.MaxFlatOpcChars}chars.");
        Console.WriteLine($"Report: {reportPath}");
    }
}
