using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlFixedRunProjectionTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace P = "http://www.w3.org/1998/Math/MathML";
    private static string Run(string text, string properties = "") =>
        "<m:r><m:rPr>" + properties + "</m:rPr><m:t>" + new XText(text) + "</m:t></m:r>";
    private static XDocument Reverse(string body) => XDocument.Parse(
        WordOmmlConverter.TransformOmmlToMathMl("<m:oMath xmlns:m='" + M + "'>" + body + "</m:oMath>", false));
    private static string Fence(string body, string properties = "") =>
        "<m:d><m:dPr>" + properties + "</m:dPr><m:e>" + body + "</m:e></m:d>";
    private static void Fixed(XElement token) => Assert.Equal("false", (string?)token.Attribute("stretchy"));

    [Theory]
    [InlineData("p", "normal")][InlineData("i", "italic")]
    [InlineData("b", "bold")][InlineData("bi", "bold-italic")]
    public void MergedNativeRunHasFixedOperatorsAndKeepsStyle(string style, string expected)
    {
        var output = Reverse(Run("(x+2)→", "<m:sty m:val='" + style + "'/>"));
        var operators = output.Descendants(P + "mo").ToArray();
        Assert.Equal(new[] { "(", "+", ")", "→" }, operators.Select(t => t.Value));
        Assert.All(operators, Fixed);
        Assert.All(operators, token => Assert.Equal(expected, MathMlTokenStyleContract.ResolveEffectiveVariant(token)));
        Assert.All(operators, token => Assert.Null(token.Attribute("fence")));
    }

    [Theory]
    [InlineData(false)][InlineData(true)]
    public void ScriptedBareParenthesesBecomeOneNativeGroupWithItsPower(bool display)
    {
        // A bracket pair typed without \left/\right is written as one native
        // group, so it grows with its content and the power sits on the group
        // instead of on a lone closing character.
        var source = "<math xmlns='" + P + "'><mo stretchy='false'>(</mo><mi>x</mi><mo>+</mo><mi>y</mi>"
            + "<msup><mo stretchy='false'>)</mo><mn>2</mn></msup></math>";
        var native = WordOmmlConverter.TransformMathMlToOmml(source);
        var actual = WordOmmlConverter.TransformOmmlToMathMl(native, display);
        Assert.Equal(
            "sup(fence((,),mi[](x)o(+)mi[](y)),n(2))",
            MathTypeMtefCodec.SemanticSignature(actual));
    }

    [Theory]
    [InlineData("0")][InlineData("1")]
    public void RealDelimiterUsesItsStructuralProjectionWithoutFixedRunAnnotation(string grow)
    {
        var output = Reverse(Fence(Run("x"), "<m:grow m:val='" + grow + "'/>"));
        var fence = Assert.Single(output.Descendants(P + "mfenced"));
        Assert.Null(fence.Attribute("stretchy"));
        Assert.Empty(output.Descendants(P + "mo"));
        // This is a noninterference control, not a claim that existing mfenced
        // projection retains the independent grow=0 delimiter capability.
    }

    [Theory]
    [InlineData("(", "")][InlineData("", "]")]
    public void OneSidedDelimiterKeepsItsEmptyEndpoint(string open, string close)
    {
        var output = Reverse(Fence(Run("x"), "<m:begChr m:val='" + open + "'/><m:endChr m:val='" + close + "'/>"));
        var fence = Assert.Single(output.Descendants(P + "mfenced"));
        Assert.Equal(open, (string?)fence.Attribute("open") ?? "(");
        Assert.Equal(close, (string?)fence.Attribute("close") ?? ")");
        Assert.Null(fence.Attribute("stretchy"));
    }

    [Fact]
    public void FixedRunInsideRealDelimiterStillHasItsOwnGeometry()
    {
        var output = Reverse(Fence(Run("(x)")));
        Assert.Single(output.Descendants(P + "mfenced"));
        Assert.Equal(2, output.Descendants(P + "mo").Count());
        Assert.All(output.Descendants(P + "mo"), Fixed);
    }

    [Fact]
    public void AccentTokenSharingParseMtDoesNotAcquireRunGeometry()
    {
        var output = Reverse("<m:acc><m:accPr><m:chr m:val='→'/></m:accPr><m:e>" + Run("(x)") + "</m:e></m:acc>");
        var accent = Assert.Single(output.Descendants(P + "mover"));
        Assert.All(accent.Elements().First().Descendants(P + "mo"), Fixed);
        Assert.DoesNotContain(accent.Elements().Last().DescendantsAndSelf(), t => (string?)t.Attribute("stretchy") == "false");
    }

    [Fact]
    public void BarDecorationIsNotAnOrdinaryRun()
    {
        var output = Reverse("<m:bar><m:barPr><m:pos m:val='top'/></m:barPr><m:e>" + Run("x") + "</m:e></m:bar>");
        Assert.DoesNotContain(output.Descendants(P + "mo"), t => (string?)t.Attribute("stretchy") == "false");
    }

    [Fact]
    public void GroupCharacterIsNotAnOrdinaryRun()
    {
        var output = Reverse("<m:groupChr><m:groupChrPr><m:chr m:val='⏞'/><m:pos m:val='top'/><m:vertJc m:val='bot'/></m:groupChrPr><m:e>" + Run("x") + "</m:e></m:groupChr>");
        Assert.DoesNotContain(output.Descendants(P + "mo"), t => (string?)t.Attribute("stretchy") == "false");
    }

    [Theory]
    [InlineData("0", "false")][InlineData("1", "true")]
    public void NaryGrowthRemainsOwnedByItsNativeStructure(string grow, string expected)
    {
        var output = Reverse("<m:nary><m:naryPr><m:chr m:val='∑'/><m:grow m:val='" + grow + "'/><m:subHide m:val='1'/><m:supHide m:val='1'/></m:naryPr><m:sub/><m:sup/><m:e>" + Run("x") + "</m:e></m:nary>");
        Assert.Equal(expected, (string?)Assert.Single(output.Descendants(P + "mo")).Attribute("stretchy"));
    }

    [Fact]
    public void EquationArrayRunKeepsFixedOperatorsAndAlignmentMarkers()
    {
        var output = Reverse("<m:eqArr><m:e>" + Run("(x)&=y") + "</m:e></m:eqArr>");
        Assert.All(output.Descendants(P + "mo"), Fixed);
        Assert.Single(output.Descendants(P + "mtable"));
        Assert.NotEmpty(output.Descendants(P + "malignmark"));
    }

    [Fact]
    public void RepeatedSameTextInDifferentScriptSlotsKeepsOwnership()
    {
        string Sup(string body, string exponent) => "<m:sSup><m:e>" + body + "</m:e><m:sup>" + Run(exponent) + "</m:sup></m:sSup>";
        var output = Reverse(Sup(Run(")"), "2") + Sup(Fence(Run("x")), "3") + Sup(Run(")"), "4"));
        var scripts = output.Descendants(P + "msup").ToArray();
        Assert.Equal(3, scripts.Length);
        Fixed(Assert.Single(scripts[0].Elements().First().Descendants(P + "mo")));
        Assert.Single(scripts[1].Elements().First().Descendants(P + "mfenced"));
        Fixed(Assert.Single(scripts[2].Elements().First().Descendants(P + "mo")));
        Assert.Equal(new[] { "2", "3", "4" }, scripts.Select(s => s.Elements().Last().Value));
    }

    [Fact]
    public void NormalTextDoesNotAcquireOperatorGeometry()
    {
        var output = Reverse(Run("(words)", "<m:nor/>"));
        var text = Assert.Single(output.Descendants(P + "mtext"));
        Assert.Equal("(words)", text.Value);
        Assert.Null(text.Attribute("stretchy"));
    }

    [Fact]
    public void OperatorEmulatorRunDoesNotBecomeScalableFence()
    {
        var output = Reverse("<m:box><m:boxPr><m:opEmu m:val='1'/></m:boxPr><m:e>"
            + Run("(x)", "<m:sty m:val='p'/>") + "</m:e></m:box>");
        Assert.Empty(output.Descendants(P + "mfenced"));
        Assert.All(output.Descendants(P + "mo"), Fixed);
    }
}
