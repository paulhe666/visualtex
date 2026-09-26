using System.Xml.Linq;
using VisualTeX.WordVsto;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeRealDocumentRegressionTests
{
    private static string RoundTrip(string source)
    {
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, true, 12);
        return MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);
    }

    private static void AssertRoundTrip(string source)
    {
        var expected = MathTypeMtefCodec.SemanticSignature(source);
        var actual = RoundTrip(source);
        Assert.Equal(expected, MathTypeMtefCodec.SemanticSignature(actual));
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, true, 12);
        var rewritten = MathTypeMtefCodec.RewriteEquationNativeAtFontSize(
            generated.EquationNative, actual, true, 18);
        Assert.Equal(expected, MathTypeMtefCodec.SemanticSignature(
            MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative)));
    }

    public static IEnumerable<object[]> EmptyBases()
    {
        foreach (var basis in new[] { "<mi/>", "<mrow/>", "<mrow><mi/></mrow>", "<mstyle><mrow/></mstyle>" })
        foreach (var script in new[] { "msub", "msup", "msubsup", "munder", "mover", "munderover", "mmultiscripts" })
        foreach (var prefix in new[] { "", "<mi>x</mi><mo>+</mo>" })
        {
            var pair = script is "msubsup" or "munderover" or "mmultiscripts";
            yield return new object[] { "<math>" + prefix + "<" + script + ">" + basis
                + "<mn>2</mn>" + (pair ? "<mn>4</mn>" : "")
                + "</" + script + "><mi>H</mi><mi>e</mi></math>" };
        }
    }

    [Theory]
    [MemberData(nameof(EmptyBases))]
    public void EmptyScriptBaseIsAnObjectNotThePreviousSibling(string source)
    {
        AssertRoundTrip(source);
        var actual = XElement.Parse(RoundTrip(source));
        var script = actual.Descendants().First(e => e.Name.LocalName is "msub" or "msup" or "msubsup");
        Assert.Empty(script.Elements().First().Value);
    }

    [Theory]
    [InlineData("<math><mo>∘</mo></math>")]
    [InlineData("<math><mi>f</mi><mo>∘</mo><mi>g</mi></math>")]
    [InlineData("<math><msup><mn>90</mn><mo>∘</mo></msup></math>")]
    public void CircleMtCodeRemainsCircleIndependentOfLegacyFontPosition(string source)
    {
        AssertRoundTrip(source);
        Assert.Contains("∘", RoundTrip(source));
        Assert.DoesNotContain("°", RoundTrip(source));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(source.Replace("∘", "°")));
    }

    [Theory]
    [InlineData("%")]
    [InlineData("#")]
    [InlineData("!")]
    [InlineData("@")]
    [InlineData("123")]
    [InlineData("a.b")]
    public void UprightOrdinaryIdentifierCanUseNativeTextTypeface(string literal)
    {
        var escaped = new XElement("mi", new XAttribute("mathvariant", "normal"), literal).ToString();
        AssertRoundTrip("<math>" + escaped + "</math>");
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature("<math>" + escaped + "</math>"),
            MathTypeMtefCodec.SemanticSignature("<math>" + escaped.Replace("normal", "bold") + "</math>"));
    }

    [Theory]
    [InlineData("\u200B")]
    [InlineData("\u2060")]
    [InlineData("\u200B\u2060")]
    [InlineData("\u200D\u200C")]
    public void UnrenderableFormatControlsCannotBeAcceptedByRelaxingTokenComparison(string controls)
    {
        // MathPage renders these unsupported CHARs as question-mark diamonds.
        // Word boundary artifacts must be sanitized before MathJax, not made to
        // pass merely by declaring their mtext and mo representations equivalent.
        var separated = "<math><mi>α</mi>" + string.Concat(controls.Select(c => "<mo>" + c + "</mo>")) + "</math>";
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(separated),
            MathTypeMtefCodec.SemanticSignature(RoundTrip(separated)));
    }

    [Theory]
    [InlineData("\\alpha \u200B\u2060", "\\alpha ")]
    [InlineData("x\u200Cy\uFEFF", "xy")]
    [InlineData("x\u200Bx", "x")]
    [InlineData("x\u2061y\u200D", "x\u2061y\u200D")]
    public void ConverterSourceUsesTheSameBoundaryCleanupAsMetadata(string source, string expected)
    {
        var cleaned = OfficeSessionDocument.SanitizeFormulaBoundaryArtifacts(source);
        Assert.Equal(expected, cleaned);
        var session = new OfficeSessionDocument { Lines = new List<FormulaLine> { new() { Id = "line", Latex = source } } };
        Assert.Equal(session.ToMetadata().Latex, cleaned);
    }

    [Fact]
    public void ActualSymbolAndStyleLossMustRemainDetectable()
    {
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature("<math><mo>+</mo></math>"),
            MathTypeMtefCodec.SemanticSignature("<math><mtext>+</mtext></math>"));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature("<math><mi mathvariant='bold'>x</mi></math>"),
            MathTypeMtefCodec.SemanticSignature("<math><mi mathvariant='normal'>x</mi></math>"));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature("<math><mo>\u200B</mo></math>"),
            MathTypeMtefCodec.SemanticSignature("<math><mo>\u2060</mo></math>"));
    }
}
