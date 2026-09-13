using System.Diagnostics;
using System.Text.Json;
using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeMathVariantRoundTripTests
{
    public static IEnumerable<object[]> SemanticVariantCases()
    {
        yield return new object[] { "normal" };
        yield return new object[] { "bold" };
        yield return new object[] { "italic" };
        yield return new object[] { "sans-serif" };
        yield return new object[] { "monospace" };
        yield return new object[] { "double-struck" };
        yield return new object[] { "script" };
        yield return new object[] { "fraktur" };
        yield return new object[] { "bold-italic" };
        yield return new object[] { "bold-script" };
        yield return new object[] { "bold-fraktur" };
        yield return new object[] { "bold-sans-serif" };
        yield return new object[] { "sans-serif-italic" };
        yield return new object[] { "sans-serif-bold-italic" };
    }

    public static IEnumerable<object[]> ExplicitVariantCases()
    {
        foreach (var variant in new[]
                 {
                     "bold",
                     "sans-serif",
                     "monospace",
                     "double-struck",
                     "script",
                     "fraktur",
                     "bold-italic",
                     "bold-script",
                     "bold-fraktur",
                     "bold-sans-serif",
                     "sans-serif-italic",
                     "sans-serif-bold-italic",
                 })
            yield return new object[] { variant };
    }

    [Theory]
    [MemberData(nameof(SemanticVariantCases))]
    public void StyledMultiLetterIdentifierSurvivesNativeCharacterRunSplitting(string sourceVariant)
    {
        var source = $"<math><mi mathvariant='{sourceVariant}'>ABC</mi></math>";
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, inline: true, 10.5);
        var read = MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(read));
    }

    [Theory]
    [InlineData("monospace", "normal")]
    [InlineData("bold", "normal")]
    [InlineData("bold-italic", "italic")]
    [InlineData("script", "fraktur")]
    public void StyledRunNormalizationStillRejectsDifferentTypography(string left, string right)
        => Assert.NotEqual(
            MathTypeMtefCodec.SemanticSignature($"<math><mi mathvariant='{left}'>ABC</mi></math>"),
            MathTypeMtefCodec.SemanticSignature($"<math><mi mathvariant='{right}'>ABC</mi></math>"));

    [Fact]
    public void StyledRunNormalizationRetainsCharacterOrder()
        => Assert.NotEqual(
            MathTypeMtefCodec.SemanticSignature("<math><mi mathvariant='monospace'>ABC</mi></math>"),
            MathTypeMtefCodec.SemanticSignature("<math><mi mathvariant='monospace'>ACB</mi></math>"));

    [Theory]
    [MemberData(nameof(SemanticVariantCases))]
    public void StandaloneEquationNativePreservesSupportedIdentifierSemantics(string sourceVariant)
    {
        var source = $"<math><mi mathvariant='{sourceVariant}'>A</mi></math>";
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, inline: false, 10.5);
        var roundTrip = MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);

        Assert.Equal(
            MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(roundTrip));
    }

    [Theory]
    [MemberData(nameof(ExplicitVariantCases))]
    public void StandaloneEquationNativePreservesExplicitIdentifierVariant(string sourceVariant)
    {
        var source = $"<math><mi mathvariant='{sourceVariant}'>A</mi></math>";
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, inline: false, 10.5);
        var roundTrip = MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);
        var identifier = XDocument.Parse(roundTrip)
            .Descendants()
            .Single(element => element.Name.LocalName == "mi" && element.Value == "A");

        Assert.Equal(sourceVariant, (string?)identifier.Attribute("mathvariant"));
    }

    [Theory]
    [InlineData("normal", @"\mathrm{A}")]
    [InlineData("bold", @"\mathbf{A}")]
    [InlineData("italic", "A")]
    [InlineData("sans-serif", @"\mathsf{A}")]
    [InlineData("monospace", @"\mathtt{A}")]
    [InlineData("double-struck", @"\mathbb{A}")]
    [InlineData("script", @"\mathcal{A}")]
    [InlineData("fraktur", @"\mathfrak{A}")]
    [InlineData("bold-italic", @"\mathbfit{A}")]
    [InlineData("bold-script", @"\boldsymbol{\mathcal{A}}")]
    [InlineData("bold-fraktur", @"\boldsymbol{\mathfrak{A}}")]
    [InlineData("bold-sans-serif", @"\boldsymbol{\mathsf{A}}")]
    public void EquationNativeReadbackKeepsLatexFontMeaning(
        string sourceVariant,
        string expectedLatex)
    {
        var mathJaxShape = $"<math><mrow data-mjx-texclass='ORD'><mi mathvariant='{sourceVariant}'>A</mi></mrow></math>";
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(mathJaxShape, inline: false, 10.5);
        var roundTrip = MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);

        Assert.Equal(expectedLatex, MathMlToLatexConverter.Convert(roundTrip));
    }

    [Fact]
    public void ActualMathJaxMathttAShapeRoundTripsWithoutLosingMonospace()
    {
        const string source = "<math xmlns='http://www.w3.org/1998/Math/MathML'><mrow data-mjx-texclass='ORD'><mi mathvariant='monospace'>A</mi></mrow></math>";
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, inline: true, 10.5);
        var roundTrip = MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);

        Assert.Equal(
            MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(roundTrip));
        Assert.Equal(@"\mathtt{A}", MathMlToLatexConverter.Convert(roundTrip));
    }

    [Theory]
    [InlineData("α", "bold-italic", @"\boldsymbol{\alpha}")]
    [InlineData("Γ", "bold", @"\boldsymbol{\Gamma}")]
    public void EquationNativeReadbackPreservesBoldGreekLatexMeaning(
        string token,
        string variant,
        string expectedLatex)
    {
        var source = $"<math><mi mathvariant='{variant}'>{token}</mi></math>";
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, inline: false, 10.5);
        var roundTrip = MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);

        Assert.Equal(
            MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(roundTrip));
        Assert.Equal(expectedLatex, MathMlToLatexConverter.Convert(roundTrip));
    }

    [Fact]
    public void PrivateMathPageAcceptsGeneratedExplicitFontVariantsWhenEnabled()
    {
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_RUN_PRIVATE_MATHTYPE_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            return;

        var windowsRoot = FindWindowsRoot();
        var bridge = Path.Combine(
            windowsRoot,
            "src-tauri",
            "target",
            "release",
            "visualtex-windows-office-bridge.exe");
        var mathPage = Path.Combine(
            windowsRoot,
            "src-tauri",
            "target",
            "release",
            "mathtype-runtime",
            "MathPage",
            "64",
            "MathPage.wll");
        Assert.True(File.Exists(bridge), $"Missing native preview bridge: {bridge}");
        Assert.True(File.Exists(mathPage), $"Missing private MathPage runtime: {mathPage}");

        var root = Path.Combine(
            Path.GetTempPath(),
            "VisualTeX",
            "mathtype-variant-acceptance",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var resultPath = Path.Combine(root, "result.json");
            var items = new List<object>();
            foreach (var variant in new[]
                     {
                         "sans-serif",
                         "monospace",
                         "double-struck",
                         "script",
                         "fraktur",
                         "bold-italic",
                         "bold-script",
                         "bold-fraktur",
                         "bold-sans-serif",
                         "sans-serif-italic",
                         "sans-serif-bold-italic",
                     })
            {
                var source = $"<math><mi mathvariant='{variant}'>A</mi></math>";
                var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(
                    source,
                    inline: false,
                    10.5);
                var mtefPath = Path.Combine(root, $"{variant}.mtef");
                var wmfPath = Path.Combine(root, $"{variant}.wmf");
                File.WriteAllBytes(mtefPath, generated.Mtef);
                items.Add(new { Id = variant, MtefPath = mtefPath, WmfPath = wmfPath });
            }
            {
                const string id = "bold-italic-greek";
                const string source = "<math><mi mathvariant='bold-italic'>&#x3B1;</mi></math>";
                var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(
                    source,
                    inline: false,
                    10.5);
                var mtefPath = Path.Combine(root, $"{id}.mtef");
                var wmfPath = Path.Combine(root, $"{id}.wmf");
                File.WriteAllBytes(mtefPath, generated.Mtef);
                items.Add(new { Id = id, MtefPath = mtefPath, WmfPath = wmfPath });
            }

            var manifestPath = Path.Combine(root, "request.json");
            File.WriteAllText(
                manifestPath,
                JsonSerializer.Serialize(new
                {
                    ResultPath = resultPath,
                    MathTypeServerPath = string.Empty,
                    Items = items,
                }));

            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = bridge,
                Arguments = $"--mathtype-preview-manifest \"{manifestPath}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
            });
            Assert.NotNull(process);
            Assert.True(process!.WaitForExit(30_000), "Private MathPage native preview timed out.");
            Assert.Equal(0, process.ExitCode);
            Assert.True(File.Exists(resultPath), "Private MathPage did not write its result manifest.");

            using var response = JsonDocument.Parse(File.ReadAllText(resultPath));
            Assert.Equal(string.Empty, response.RootElement.GetProperty("Error").GetString());
            var results = response.RootElement.GetProperty("Items").EnumerateArray().ToArray();
            Assert.Equal(items.Count, results.Length);
            var expectedWmfFonts = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["sans-serif"] = "Arial",
                ["monospace"] = "Courier New",
                ["double-struck"] = "Euclid Math Two",
                ["script"] = "Euclid Math One",
                ["fraktur"] = "Euclid Fraktur",
                ["bold-italic"] = "Times New Roman",
                ["bold-script"] = "Euclid Math One",
                ["bold-fraktur"] = "Euclid Fraktur",
                ["bold-sans-serif"] = "Arial",
                ["sans-serif-italic"] = "Arial",
                ["sans-serif-bold-italic"] = "Arial",
                ["bold-italic-greek"] = "Times New Roman",
            };
            foreach (var result in results)
            {
                var id = result.GetProperty("Id").GetString();
                Assert.True(
                    result.GetProperty("Success").GetBoolean(),
                    id + ": " + result.GetProperty("Error").GetString());
                var wmfPath = result.GetProperty("WmfPath").GetString();
                Assert.False(string.IsNullOrWhiteSpace(wmfPath));
                Assert.True(new FileInfo(wmfPath!).Length > 22);
                Assert.True(result.GetProperty("WidthPt").GetSingle() > 0);
                Assert.True(result.GetProperty("HeightPt").GetSingle() > 0);
                Assert.NotNull(id);
                Assert.True(expectedWmfFonts.TryGetValue(id!, out var expectedFont));
                var wmfStrings = System.Text.Encoding.Latin1.GetString(File.ReadAllBytes(wmfPath!));
                Assert.Contains(expectedFont, wmfStrings, StringComparison.OrdinalIgnoreCase);
            }
        }
        finally
        {
            try { Directory.Delete(root, recursive: true); } catch { }
        }
    }

    private static string FindWindowsRoot()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory);
             directory is not null;
             directory = directory.Parent)
        {
            if (File.Exists(Path.Combine(directory.FullName, "package.json"))
                && Directory.Exists(Path.Combine(directory.FullName, "src-tauri")))
                return directory.FullName;
        }
        throw new DirectoryNotFoundException("Could not locate the VisualTeX Windows workspace root.");
    }
}
