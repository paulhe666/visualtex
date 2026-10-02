using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeFenceRoundTripTests
{
    [Theory]
    [InlineData("{", "")]
    [InlineData("", "}")]
    [InlineData("(", "")]
    [InlineData("", "]")]
    [InlineData("|", "")]
    [InlineData("{", "}")]
    public void ExplicitFencesPreserveSemanticsAndExpandableTemplate(string open, string close)
    {
        var source = $"<math><mrow><mo fence='true' stretchy='true'>{open}</mo>"
            + "<mtable><mtr><mtd><mfrac><mi>x</mi><mi>y</mi></mfrac></mtd></mtr>"
            + "<mtr><mtd><mi>z</mi></mtd></mtr></mtable>"
            + $"<mo fence='true' stretchy='true'>{close}</mo></mrow></math>";
        var generated = MathTypeMtefCodec.CreateEquationNative(source, inline: false);
        var roundTrip = MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(roundTrip));
        var fence = XDocument.Parse(roundTrip).Descendants().Single(e => e.Name.LocalName == "mfenced");
        Assert.Equal(open, (string?)fence.Attribute("open"));
        Assert.Equal(close, (string?)fence.Attribute("close"));
    }

    [Fact]
    public void EmptyTokenAddsNoGlyphButCannotInventOrEraseARealClosingFence()
    {
        // MathML 3 section3.2.1 permits zero-character tokens: their glyph
        // content is empty, not an unspecified missing operator. This older
        // assertion incorrectly required an empty token to carry mathematics.
        // Preserve the real invariant: a nonempty closing fence is distinct.
        const string source = "<math><mrow><mo>{</mo><mi>x</mi><mo/></mrow></math>";
        const string withoutEmpty = "<math><mrow><mo>{</mo><mi>x</mi></mrow></math>";
        const string withClosing = "<math><mrow><mo>{</mo><mi>x</mi><mo>}</mo></mrow></math>";
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(withoutEmpty));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(withClosing));
        var native=MathTypeMtefCodec.CreateEquationNative(source,false);
        var actual=MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source),MathTypeMtefCodec.SemanticSignature(actual));
        Assert.DoesNotContain("}",XElement.Parse(actual).Value);
    }

    [Fact]
    public void ARealClosingFenceIsNotEquivalentToAnEmptyOne()
    {
        const string leftOnly = "<math><mfenced open='{' close=''><mi>x</mi></mfenced></math>";
        const string paired = "<math><mfenced open='{' close='}'><mi>x</mi></mfenced></math>";
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(leftOnly), MathTypeMtefCodec.SemanticSignature(paired));
    }

    [Fact]
    public void OneSidedWordFenceEqualsLooseLeftDelimiterAndTable()
    {
        const string loose =
            "<math><mrow><mo fence='true' stretchy='true'>{</mo>"
            + "<mtable><mtr><mtd><mi>x</mi></mtd><mtd><mi>y</mi></mtd></mtr>"
            + "<mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr></mtable>"
            + "<mo fence='true' stretchy='true'></mo></mrow></math>";
        const string word =
            "<math><mfenced open='{' close=''><mtable>"
            + "<mtr><mtd><mi>x</mi></mtd><mtd><mi>y</mi></mtd></mtr>"
            + "<mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr>"
            + "</mtable></mfenced></math>";
        Assert.Equal(
            MathTypeMtefCodec.SemanticSignature(loose),
            MathTypeMtefCodec.SemanticSignature(word));
    }

    [Fact]
    public void NamedOperatorEqualsWordUprightLetterRun()
    {
        const string named =
            "<math><msub><mi>x</mi><mo>max</mo></msub></math>";
        const string upright =
            "<math><msub><mi>x</mi><mrow>"
            + "<mi mathvariant='normal'>m</mi>"
            + "<mi mathvariant='normal'>a</mi>"
            + "<mi mathvariant='normal'>x</mi>"
            + "</mrow></msub></math>";
        Assert.Equal(
            MathTypeMtefCodec.SemanticSignature(named),
            MathTypeMtefCodec.SemanticSignature(upright));
    }
}
