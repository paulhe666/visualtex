using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeUprightIdentifierSpacingTests
{
    [Theory]
    [InlineData("a", "b")]
    [InlineData("AB", "cd")]
    [InlineData("α", "b")]
    public void UprightIdentifiersAroundMathSpaceCannotBecomeProse(string left, string right)
    {
        var source = $"<math><mstyle mathvariant='normal'><mi>{left}</mi><mspace width='1em'/><mi>{right}</mi></mstyle></math>";
        var generated = MathTypeMtefCodec.CreateEquationNative(source, false);
        var actual = MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);
        Assert.DoesNotContain(XDocument.Parse(actual).Descendants(), e => e.Name.LocalName == "mtext");
        Assert.Equal(left + right, string.Concat(XDocument.Parse(actual).Descendants().Where(e => e.Name.LocalName == "mi").Select(e => e.Value)));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(actual));
        var rewritten = MathTypeMtefCodec.RewriteEquationNative(generated.EquationNative, actual, false);
        Assert.Equal(generated.Mtef, rewritten.Mtef);
    }

    [Fact]
    public void ParallelRelationSequenceKeepsEachUprightIdentifierOutsideProse()
    {
        // Exact production frontend export for representative corpus F01.
        const string source = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">\n  <mrow data-mjx-texclass=\"ORD\">\n    <mi mathvariant=\"normal\">a</mi>\n    <mo>&#x2225;</mo>\n    <mi mathvariant=\"normal\">b</mi>\n    <mstyle scriptlevel=\"0\">\n      <mspace width=\"1em\"></mspace>\n    </mstyle>\n    <mi mathvariant=\"normal\">a</mi>\n    <mo data-mjx-alternate=\"1\">&#x2225;</mo>\n    <mi mathvariant=\"normal\">b</mi>\n    <mstyle scriptlevel=\"0\">\n      <mspace width=\"1em\"></mspace>\n    </mstyle>\n    <mi mathvariant=\"normal\">a</mi>\n    <mo>&#x2226;</mo>\n    <mi mathvariant=\"normal\">b</mi>\n    <mstyle scriptlevel=\"0\">\n      <mspace width=\"1em\"></mspace>\n    </mstyle>\n    <mi mathvariant=\"normal\">a</mi>\n    <mo data-mjx-alternate=\"1\">&#x2226;</mo>\n    <mi mathvariant=\"normal\">b</mi>\n  </mrow>\n</math>";
        var generated = MathTypeMtefCodec.CreateEquationNative(source, false);
        var actual = MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);
        Assert.Equal(8, XDocument.Parse(actual).Descendants().Count(e => e.Name.LocalName == "mi" && (e.Value == "a" || e.Value == "b")));
        Assert.DoesNotContain(XDocument.Parse(actual).Descendants(), e => e.Name.LocalName == "mtext");
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(actual));
    }

    [Fact]
    public void ProseStillCoalescesItsOwnSpacesWithoutAbsorbingUprightIdentifiers()
    {
        const string source = "<math><mtext>two words</mtext><mi mathvariant='normal'>a</mi><mspace width='1em'/><mi mathvariant='normal'>b</mi><mtext>more words</mtext></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var actual = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative));
        Assert.Equal(new[] { "two words", "more words" }, actual.Descendants().Where(e => e.Name.LocalName == "mtext").Select(e => e.Value).ToArray());
        Assert.Equal(new[] { "a", "b" }, actual.Descendants().Where(e => e.Name.LocalName == "mi").Select(e => e.Value).ToArray());
    }
}
