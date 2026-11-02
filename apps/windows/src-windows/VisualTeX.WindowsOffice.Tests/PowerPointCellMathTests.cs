using System.Xml.Linq;
using VisualTeX.PowerPointVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class PowerPointCellMathTests
{
    private const string Namespace = "http://www.w3.org/1998/Math/MathML";

    private static XElement[] Tokens(string body) =>
        XDocument.Parse(PowerPointCellMath.Prepare($"<math xmlns=\"{Namespace}\">{body}</math>"))
            .Root!.Descendants().Where(element => element.Name.LocalName is "mi" or "mn" or "mtext").ToArray();

    [Fact]
    public void UprightLettersBecomeNormalText()
    {
        var tokens = Tokens("<msub><mi mathvariant=\"normal\">M</mi><mrow><mi mathvariant=\"normal\">Na</mi></mrow></msub><mi>abc</mi>");
        Assert.All(tokens, token => Assert.Equal("mtext", token.Name.LocalName));
        Assert.Equal(new[] { "M", "Na", "abc" }, tokens.Select(token => token.Value));
        Assert.All(tokens, token => Assert.Null(token.Attribute("mathvariant")));
    }

    [Fact]
    public void FunctionNamesAndItalicLettersStayIdentifiers()
    {
        var tokens = Tokens("<mi>sin</mi><mo>⁡</mo><mi>x</mi><mi mathvariant=\"normal\">cos</mi><mo>⁡</mo><mi>y</mi>");
        Assert.Equal(new[] { "mi", "mi", "mi", "mi" }, tokens.Select(token => token.Name.LocalName));
        Assert.Equal(new[] { "sin", "x", "cos", "y" }, tokens.Select(token => token.Value));
    }

    [Fact]
    public void StyledLettersAreSpelledInMathematicalAlphanumerics()
    {
        var tokens = Tokens("<mi mathvariant=\"bold\">v</mi><mi mathvariant=\"double-struck\">R</mi><mi mathvariant=\"script\">B</mi>"
            + "<mi mathvariant=\"fraktur\">g</mi><mn mathvariant=\"bold\">12</mn><mstyle mathvariant=\"bold\"><mi>A</mi></mstyle>");
        Assert.Equal(new[] { "\U0001D42F", "ℝ", "ℬ", "\U0001D524", "\U0001D7CF\U0001D7D0", "\U0001D400" },
            tokens.Select(token => token.Value));
        // A single styled letter is not italicized again.
        Assert.Equal("normal", (string?)tokens[0].Attribute("mathvariant"));
    }

    [Fact]
    public void NonLettersAndUnknownVariantsAreKept()
    {
        Assert.Equal("α+", PowerPointCellMath.Styled("α+", "bold"));
        Assert.Null(PowerPointCellMath.Styled("x", "stretched"));
        Assert.Throws<InvalidDataException>(() => PowerPointCellMath.Prepare("<math"));
    }
}
