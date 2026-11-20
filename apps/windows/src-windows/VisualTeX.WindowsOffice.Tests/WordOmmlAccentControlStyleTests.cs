using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public class WordOmmlAccentControlStyleTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace ML = "http://www.w3.org/1998/Math/MathML";

    [Theory]
    [InlineData("normal","normal","0","0")]
    [InlineData("bold","normal","1","0")]
    [InlineData("normal","bold","0","0")]
    [InlineData("italic","normal","0","1")]
    [InlineData("bold-italic","normal","1","1")]
    [InlineData("bold","bold","1","0")]
    public void ProjectsAccentStyleIndependentlyOfBase(string mark, string basis, string bold, string italic)
    {
        var input = $"<math xmlns='{ML}'><mover><mi mathvariant='{basis}'>x</mi><mo mathvariant='{mark}'>^</mo></mover></math>";
        var omml = WordOmmlConverter.TransformMathMlToOmml(input);
        var root = XElement.Parse(omml);
        var control = Assert.Single(root.Descendants(M + "accPr")).Element(M + "ctrlPr")!.Element(W + "rPr")!;
        Assert.Equal(bold, (string?)control.Element(W + "b")!.Attribute(W + "val"));
        Assert.Equal(italic, (string?)control.Element(W + "i")!.Attribute(W + "val"));
        var actual = XElement.Parse(WordOmmlConverter.TransformOmmlToMathMl(omml, false));
        Assert.Equal(mark, (string?)Assert.Single(actual.Descendants(ML + "mo")).Attribute("mathvariant"));
        Assert.Equal(basis, (string?)Assert.Single(actual.Descendants(ML + "mi")).Attribute("mathvariant"));
    }

    [Theory]
    [InlineData("1","bold")]
    [InlineData("true","bold")]
    [InlineData("on","bold")]
    [InlineData("0","normal")]
    [InlineData("false","normal")]
    [InlineData("off","normal")]
    public void ReadsNativeControlOnOffValues(string value, string expected)
    {
        var omml = $"<m:oMath xmlns:m='{M}' xmlns:w='{W}'><m:acc><m:accPr><m:ctrlPr><w:rPr><w:b w:val='{value}'/></w:rPr></m:ctrlPr></m:accPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc></m:oMath>";
        var actual = XElement.Parse(WordOmmlConverter.TransformOmmlToMathMl(omml, false));
        Assert.Equal(expected, (string?)Assert.Single(actual.Descendants(ML + "mo")).Attribute("mathvariant"));
    }

    [Fact]
    public void NestedAccentsKeepEachControlStyleAndBase()
    {
        var input = $"<math xmlns='{ML}'><mover><mover><mi>x</mi><mo mathvariant='bold'>~</mo></mover><mo mathvariant='normal'>^</mo></mover></math>";
        var actual = XElement.Parse(WordOmmlConverter.TransformOmmlToMathMl(WordOmmlConverter.TransformMathMlToOmml(input), false));
        var outer = Assert.Single(actual.Elements(ML + "mover"));
        Assert.Equal("normal", (string?)outer.Elements().Last().Attribute("mathvariant"));
        var inner = Assert.Single(outer.Descendants(ML + "mover"));
        Assert.Equal("bold", (string?)inner.Elements().Last().Attribute("mathvariant"));
        Assert.Equal("x", Assert.Single(inner.Descendants(ML + "mi")).Value);
    }

    [Theory]
    [InlineData("bold", @"\mathbf{\hat{\mathnormal{x}}}")]
    [InlineData("italic", @"\mathit{\hat{\mathnormal{x}}}")]
    [InlineData("bold-italic", @"\mathbfit{\hat{\mathnormal{x}}}")]
    public void RecoversAccentScopeWithoutRestylingBase(string variant, string expected)
    {
        var input = $"<math xmlns='{ML}'><mover><mi>x</mi><mo mathvariant='{variant}' stretchy='false'>^</mo></mover></math>";
        Assert.Equal(expected, MathMlToLatexConverter.Convert(input));
    }

    [Fact]
    public void ExplicitNonAccentKeepsLimitStructure()
    {
        var input = $"<math xmlns='{ML}'><mover accent='false'><mi>x</mi><mo mathvariant='bold'>^</mo></mover></math>";
        var root = XElement.Parse(WordOmmlConverter.TransformMathMlToOmml(input));
        Assert.Empty(root.Descendants(M + "acc"));
        Assert.Single(root.Descendants(M + "limUpp"));
    }
}
