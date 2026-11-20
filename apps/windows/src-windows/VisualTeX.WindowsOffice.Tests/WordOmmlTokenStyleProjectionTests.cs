using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;
public sealed class WordOmmlTokenStyleProjectionTests
{
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    [Theory]
    [InlineData("⋅", "b", "bold")]
    [InlineData("×", "b", "bold")]
    [InlineData("∥", "i", "italic")]
    [InlineData("∥", "bi", "bold-italic")]
    [InlineData("∠", "i", "italic")]
    [InlineData("∇", "b", "bold")]
    [InlineData("2", "b", "bold")]
    [InlineData("2", "i", "italic")]
    [InlineData("2", "bi", "bold-italic")]
    public void NativeRunStyleReachesMathMlToken(string glyph, string style, string expected)
    {
        var omml = $"<m:oMath xmlns:m=\"{M}\"><m:r><m:rPr><m:sty m:val=\"{style}\"/></m:rPr><m:t>{glyph}</m:t></m:r></m:oMath>";
        var math = XDocument.Parse(WordOmmlConverter.TransformOmmlToMathMl(omml, false));
        var token = Assert.Single(math.Descendants().Where(node => node.Name.LocalName is "mi" or "mo" or "mn"));
        Assert.Equal(glyph, token.Value);
        var variant = (string?)token.Attribute("mathvariant")
            ?? (token.Name.LocalName == "mi" && token.Value.Length == 1 ? "italic" : null);
        Assert.Equal(expected, variant);
    }

    [Theory]
    [InlineData("")]
    [InlineData("<m:rPr><m:sty m:val=\"p\"/></m:rPr>")]
    public void OrdinaryDigitsKeepDefaultSourceContract(string properties)
    {
        var omml = $"<m:oMath xmlns:m=\"{M}\"><m:r>{properties}<m:t>123</m:t></m:r></m:oMath>";
        var xml = WordOmmlConverter.TransformOmmlToMathMl(omml, false);
        var token = Assert.Single(XDocument.Parse(xml).Descendants().Where(node => node.Name.LocalName == "mn"));
        Assert.Null(token.Attribute("mathvariant"));
        Assert.Null(token.Attribute("fontstyle"));
        Assert.Equal("123", MathMlToLatexConverter.Convert(xml).Trim());
    }

    [Fact]
    public void StyledScriptRunRetainsBaseAndExponentOwnership()
    {
        var omml = $"<m:oMath xmlns:m=\"{M}\"><m:sSup><m:e><m:r><m:rPr><m:sty m:val=\"b\"/></m:rPr><m:t>∥</m:t></m:r></m:e><m:sup><m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>";
        var root = XDocument.Parse(WordOmmlConverter.TransformOmmlToMathMl(omml, false));
        var script = Assert.Single(root.Descendants().Where(node => node.Name.LocalName == "msup"));
        var slots = script.Elements().ToArray();
        Assert.Equal(2, slots.Length);
        var baseToken = Assert.Single(slots[0].DescendantsAndSelf().Where(node => node.Name.LocalName == "mo"));
        Assert.Equal("bold", (string?)baseToken.Attribute("mathvariant"));
        Assert.Equal("2", slots[1].Value);
    }
}
