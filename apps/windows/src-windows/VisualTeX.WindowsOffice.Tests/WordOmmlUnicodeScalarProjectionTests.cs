using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlUnicodeScalarProjectionTests
{
    private readonly ClassicOfficeReverseStylesheet.UnicodeScalarXPathFunctions _functions = new();

    [Theory]
    [InlineData("abc", 3)]
    [InlineData("𝓁", 1)]
    [InlineData("a𝓁b", 3)]
    [InlineData("𝓁𝔸", 2)]
    [InlineData("", 0)]
    public void LengthCountsScalars(string text, int expected)
        => Assert.Equal(expected, _functions.StringLength(text));

    [Theory]
    [InlineData(1, 1, "a")]
    [InlineData(2, 1, "𝓁")]
    [InlineData(3, 1, "b")]
    [InlineData(1.5, 1, "𝓁")]
    [InlineData(0, 2, "a")]
    [InlineData(-2, 5, "a𝓁")]
    [InlineData(2, 0, "")]
    public void SubstringUsesXPathRoundingAndWholeScalars(double start, double length, string expected)
        => Assert.Equal(expected, _functions.Substring("a𝓁b", start, length));

    [Fact]
    public void ExceptionalXPathNumbersKeepTheirDefinedSemantics()
    {
        Assert.Equal("", _functions.Substring("a𝓁b", double.NaN, 3));
        Assert.Equal("", _functions.Substring("a𝓁b", 1, double.NaN));
        Assert.Equal("", _functions.Substring("a𝓁b", double.NegativeInfinity, double.PositiveInfinity));
        Assert.Equal("a𝓁b", _functions.Substring("a𝓁b", -42, double.PositiveInfinity));
        Assert.Equal("𝓁b", _functions.Substring("a𝓁b", 2));
    }

    [Fact]
    public void ValuesImmediatelyBelowHalfDoNotRoundUpFromFloatingPointAddition()
    {
        var belowHalf = Math.BitDecrement(0.5);
        Assert.Equal("", _functions.Substring("ab", belowHalf, 1));
        Assert.Equal("", _functions.Substring("ab", 1, belowHalf));
        Assert.Equal("a", _functions.Substring("ab", 0.5, 1));
    }

    [Fact]
    public void XPathRewritePreservesQuotedStringsAndQualifiedNames()
    {
        Assert.Equal("vtUnicode:Substring($text, 1, vtUnicode:StringLength($text))", ClassicOfficeReverseStylesheet.RewriteUnicodeScalarXPath("substring($text, 1, string-length($text))"));
        const string expression = "concat('substring(x)', \"string-length(y)\", custom:substring($text,1), $substring)";
        Assert.Equal(expression, ClassicOfficeReverseStylesheet.RewriteUnicodeScalarXPath(expression));
        Assert.Equal("vtUnicode:Substring ($text,2)", ClassicOfficeReverseStylesheet.RewriteUnicodeScalarXPath("substring ($text,2)"));
    }

    [Theory]
    [InlineData("𝓁")]
    [InlineData("𝔸")]
    [InlineData("𝟙")]
    [InlineData("𠀀")]
    [InlineData("😀")]
    [InlineData("a𝓁b𝔸")]
    [InlineData("a𝓁+𝔸=𝟙")]
    public void ActualReverseProjectionNeverSplitsValidSupplementaryCharacters(string text)
    {
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var source = new XElement(m + "oMath", new XElement(m + "r",
            new XElement(m + "rPr", new XElement(m + "sty", new XAttribute(m + "val", "p"))),
            new XElement(m + "t", text)));
        var result = XElement.Parse(WordOmmlConverter.TransformOmmlToMathMl(source.ToString(), false));
        Assert.Equal(text, result.Value);
        Assert.All(result.Descendants().Where(x => x.Name.LocalName == "mi"), token => Assert.Equal("normal", (string?)token.Attribute("mathvariant")));
    }
}
