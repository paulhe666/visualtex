using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeLargeCorpusRegressionTests
{
    private static string RoundTrip(string source, bool inline = false)
    {
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, inline, 8);
        return MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SansSerifTransposeSurvivesEquationNativeAndRewrite(bool inline)
    {
        const string source = "<math><msub><mi mathvariant='bold'>Q</mi><mn>4</mn></msub>"
            + "<msup><mn>96</mn><mrow><mrow><mi mathvariant='sans-serif'>T</mi></mrow></mrow></msup>"
            + "<mo>=</mo><mi mathvariant='bold'>I</mi></math>";
        var actual = RoundTrip(source, inline);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(actual));
        var transpose = XDocument.Parse(actual).Descendants().Single(e => e.Name.LocalName == "mi" && e.Value == "T");
        Assert.Equal("sans-serif", (string?)transpose.Attribute("mathvariant"));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(actual.Replace("sans-serif", "normal")));
        var original = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, inline, 8);
        var edited = MathTypeMtefCodec.RewriteEquationNativeAtFontSize(original.EquationNative, source, inline, 8);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(edited.EquationNative)));
    }

    [Theory]
    [InlineData("&#x27E8;", "&#x27E9;")]
    [InlineData("(", ")")]
    [InlineData("[", "]")]
    public void FixedInnerDelimitersDoNotBecomeScalableFences(string open, string close)
    {
        var source = "<math><msup><mrow data-mjx-texclass='INNER'><mo data-mjx-texclass='OPEN'>|</mo>"
            + $"<mo fence='false' stretchy='false'>{open}</mo>"
            + "<msub><mi>f</mi><mn>4</mn></msub><mn>92</mn><mo stretchy='false'>|</mo>"
            + "<msub><mover><mi>V</mi><mo stretchy='false'>^</mo></mover><mn>4</mn></msub><mn>92</mn>"
            + "<mo stretchy='false'>|</mo><msub><mi>i</mi><mn>4</mn></msub><mn>92</mn>"
            + $"<mo fence='false' stretchy='false'>{close}</mo>"
            + "<mo data-mjx-texclass='CLOSE'>|</mo></mrow><mn>2</mn></msup></math>";
        var actual = RoundTrip(source);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(actual));
        Assert.Single(XDocument.Parse(actual).Descendants(), e => e.Name.LocalName == "mfenced");
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(actual.Replace(">V<", ">W<")));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(actual.Replace(">2<", ">3<")));
    }
}
