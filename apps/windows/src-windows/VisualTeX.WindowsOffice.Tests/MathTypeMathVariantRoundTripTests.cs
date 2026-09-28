using VisualTeX.WindowsOffice.Contracts;
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
    public async Task PrivateServiceAcceptsGeneratedExplicitFontVariantsWhenEnabled()
    {
        if (Environment.GetEnvironmentVariable("VISUALTEX_RUN_PRIVATE_MATHTYPE_ACCEPTANCE") != "1")
            return;

        Assert.True(NativeRenderClient.IsAvailable, "The private render service must be installed for this live test.");
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
        var request = new NativeRenderRequest();
        foreach (var id in expectedWmfFonts.Keys)
        {
            var variant = id == "bold-italic-greek" ? "bold-italic" : id;
            var character = id == "bold-italic-greek" ? "&#x3B1;" : "A";
            var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(
                $"<math><mi mathvariant='{variant}'>{character}</mi></math>", inline: false, 10.5);
            request.Items.Add(new NativeRenderInput { Id = id, Mtef = generated.Mtef });
        }
        var response = await NativeRenderClient.RenderAsync(request, CancellationToken.None);
        Assert.Equal(string.Empty, response.Error);
        Assert.Equal(request.Items.Count, response.Items.Count);
        foreach (var result in response.Items)
        {
            Assert.True(result.WidthPt > 0 && result.HeightPt > 0);
            Assert.True(result.Wmf.Length > 22);
            Assert.True(expectedWmfFonts.TryGetValue(result.Id, out var expectedFont));
            Assert.Contains(expectedFont, System.Text.Encoding.Latin1.GetString(result.Wmf),
                StringComparison.OrdinalIgnoreCase);
        }
    }
}
