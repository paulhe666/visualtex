using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeRuleLessStackRoundTripTests
{
    private const string Ns = "http://www.w3.org/1998/Math/MathML";

    [Fact]
    public void WordOmmlBinomialWithFixedLiteralParenthesesSurvivesTheMtefRoundTrip()
    {
        // Office's OMML2MML output for a VisualTeX-inserted \binom: fixed
        // literal parentheses around a zero-rule fraction. MTEF writes the
        // fraction as a PILE, which reads back as a two-row table.
        var source = $"<math display='block' xmlns='{Ns}'>"
            + "<munderover><mo stretchy='true'>∑</mo><mrow><mi>k</mi><mo stretchy='false' mathvariant='normal'>=</mo><mn>0</mn></mrow><mrow><mi>n</mi></mrow></munderover>"
            + "<mrow><mo stretchy='false' mathvariant='normal'>(</mo>"
            + "<mfrac linethickness='0'><mrow><mi>n</mi></mrow><mrow><mi>k</mi></mrow></mfrac>"
            + "<mo stretchy='false' mathvariant='normal'>)</mo></mrow></math>";
        var generated = MathTypeMtefCodec.CreateEquationNative(source, inline: false);
        var roundTrip = MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);
        Assert.Contains("data-mtef-pile", roundTrip, StringComparison.Ordinal);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(roundTrip));
    }

    [Fact]
    public void RuleLessFractionAndTwoRowColumnShareOneStackSignature()
    {
        var fraction = $"<math xmlns='{Ns}'><mfrac linethickness='0'><mi>a</mi><mi>b</mi></mfrac></math>";
        var pile = $"<math xmlns='{Ns}'><mtable data-mtef-pile='true'><mtr><mtd><mi>a</mi></mtd></mtr><mtr><mtd><mi>b</mi></mtd></mtr></mtable></math>";
        var column = $"<math xmlns='{Ns}'><mtable><mtr><mtd><mi>a</mi></mtd></mtr><mtr><mtd><mi>b</mi></mtd></mtr></mtable></math>";
        var ruled = $"<math xmlns='{Ns}'><mfrac><mi>a</mi><mi>b</mi></mfrac></math>";
        var expected = MathTypeMtefCodec.SemanticSignature(fraction);
        Assert.Equal(expected, MathTypeMtefCodec.SemanticSignature(pile));
        Assert.Equal(expected, MathTypeMtefCodec.SemanticSignature(column));
        // A ruled fraction is a different object and must stay distinct.
        Assert.NotEqual(expected, MathTypeMtefCodec.SemanticSignature(ruled));
    }
}
