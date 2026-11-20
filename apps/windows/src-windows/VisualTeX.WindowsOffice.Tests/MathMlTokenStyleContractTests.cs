using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathMlTokenStyleContractTests
{
    [Theory]
    [InlineData("")]
    [InlineData("http://www.w3.org/1998/Math/MathML")]
    public void NearestDeclarationWinsWithoutChangingTheTree(string ns)
    {
        var root = XElement.Parse($"<math xmlns='{ns}' mathvariant='normal'><mstyle mathvariant='bold'><mrow mathvariant='italic'><mi>x</mi><mi mathvariant='bold-italic'>x</mi></mrow></mstyle></math>");
        var before = root.ToString(SaveOptions.DisableFormatting);
        var tokens = root.Descendants().Where(e => e.Name.LocalName == "mi").ToArray();
        Assert.Equal("italic", MathMlTokenStyleContract.ResolveDeclaredVariant(tokens[0]));
        Assert.Equal("bold-italic", MathMlTokenStyleContract.ResolveDeclaredVariant(tokens[1]));
        Assert.Equal("italic", MathMlTokenStyleContract.ResolveEffectiveVariant(tokens[0]));
        Assert.Equal("bold-italic", MathMlTokenStyleContract.ResolveEffectiveVariant(tokens[1]));
        Assert.Equal(before, root.ToString(SaveOptions.DisableFormatting));
    }

    [Fact]
    public void MissingDeclarationIsNotAnExplicitDefault()
    {
        var implicitToken = XElement.Parse("<mi>x</mi>");
        var explicitToken = XElement.Parse("<mi mathvariant='italic'>x</mi>");
        Assert.Null(MathMlTokenStyleContract.ResolveDeclaredVariant(implicitToken));
        Assert.Equal("italic", MathMlTokenStyleContract.ResolveDeclaredVariant(explicitToken));
        Assert.Equal("italic", MathMlTokenStyleContract.ResolveEffectiveVariant(implicitToken));
        Assert.Equal("italic", MathMlTokenStyleContract.ResolveEffectiveVariant(explicitToken));
        Assert.Equal("normal", MathMlTokenStyleContract.ResolveEffectiveVariant(XElement.Parse("<mi mathvariant='normal'>x</mi>")));
    }

    [Fact]
    public void EmptyDeclarationDoesNotHideNearestNonemptyDeclaration()
    {
        var token = XElement.Parse("<mstyle mathvariant=' BOLD-ITALIC '><mi mathvariant=' '>x</mi></mstyle>").Elements().Single();
        Assert.Equal("bold-italic", MathMlTokenStyleContract.ResolveDeclaredVariant(token));
        Assert.Equal("bold-italic", MathMlTokenStyleContract.ResolveEffectiveVariant(token));
    }

    [Theory]
    [InlineData("mi", "x", "italic")]
    [InlineData("mi", "Γ", "italic")]
    [InlineData("mi", "sin", "normal")]
    [InlineData("mi", "", "normal")]
    [InlineData("mo", "∥", "normal")]
    [InlineData("mo", "lim", "normal")]
    [InlineData("mn", "2", "normal")]
    [InlineData("mtext", "x", "normal")]
    [InlineData("ms", "x", "normal")]
    public void DefaultsBelongToTokenRole(string kind, string text, string expected)
    {
        var token = new XElement(kind, text);
        Assert.Null(MathMlTokenStyleContract.ResolveDeclaredVariant(token));
        Assert.Equal(expected, MathMlTokenStyleContract.ResolveEffectiveVariant(token));
    }

    [Theory]
    [InlineData("normal")]
    [InlineData("italic")]
    [InlineData("bold")]
    [InlineData("bold-italic")]
    public void ExplicitGreekStyleIsNeverReplacedWithATexCommandDefault(string variant)
    {
        var token = new XElement("mi", new XAttribute("mathvariant", variant), "Γ");
        Assert.Equal(variant, MathMlTokenStyleContract.ResolveDeclaredVariant(token));
        Assert.Equal(variant, MathMlTokenStyleContract.ResolveEffectiveVariant(token));
    }

    [Fact]
    public void FunctionApplicationDoesNotInventUprightStyle()
    {
        var row = XElement.Parse("<mrow><mi>f</mi><mo>&#x2061;</mo><mi>x</mi></mrow>");
        Assert.Equal("italic", MathMlTokenStyleContract.ResolveEffectiveVariant(row.Elements().First()));
        Assert.Equal("normal", MathMlTokenStyleContract.ResolveEffectiveVariant(row.Elements().Skip(1).First()));
    }

    [Fact]
    public void ProseAndNativeMetadataDoNotChangeTokenRoleOrStyleAuthority()
    {
        var prose = XElement.Parse("<mtext data-mtef-character-style='3'>sin</mtext>");
        var identifier = XElement.Parse("<mi data-mtef-character-style='0' data-mtef-explicit-font='Symbol'>Γ</mi>");
        Assert.Equal("normal", MathMlTokenStyleContract.ResolveEffectiveVariant(prose));
        Assert.Equal("italic", MathMlTokenStyleContract.ResolveEffectiveVariant(identifier));
        Assert.Equal("mtext", prose.Name.LocalName);
        Assert.Equal("mi", identifier.Name.LocalName);
    }

    [Fact]
    public void ForeignXmlCannotSupplyADeclaration()
    {
        var root = XElement.Parse("<mstyle mathvariant='bold'><foreign xmlns='urn:foreign' mathvariant='normal'><mi xmlns=''>x</mi></foreign></mstyle>");
        var token = root.Descendants().Single(e => e.Name.LocalName == "mi");
        Assert.Null(MathMlTokenStyleContract.ResolveDeclaredVariant(token));
        Assert.Equal("italic", MathMlTokenStyleContract.ResolveEffectiveVariant(token));
        var foreignToken = XElement.Parse("<mi xmlns='urn:foreign' mathvariant='italic'>x</mi>");
        Assert.Null(MathMlTokenStyleContract.ResolveDeclaredVariant(foreignToken));
        Assert.Null(MathMlTokenStyleContract.ResolveEffectiveVariant(foreignToken));
    }

    [Fact]
    public void HostXmlOutsideTheMathRootCannotSupplyADeclaration()
    {
        var root = XElement.Parse("<host mathvariant='bold'><math><mi>x</mi></math></host>");
        var token = root.Descendants().Single(e => e.Name.LocalName == "mi");
        Assert.Null(MathMlTokenStyleContract.ResolveDeclaredVariant(token));
        Assert.Equal("italic", MathMlTokenStyleContract.ResolveEffectiveVariant(token));
    }

    [Fact]
    public void UnsupportedDeclarationsRemainVisibleInsteadOfBecomingDefaults()
    {
        var token = XElement.Parse("<mi mathvariant='future-variant'>x</mi>");
        Assert.Equal("future-variant", MathMlTokenStyleContract.ResolveDeclaredVariant(token));
        Assert.Equal("future-variant", MathMlTokenStyleContract.ResolveEffectiveVariant(token));
        Assert.Null(MathMlTokenStyleContract.ResolveEffectiveVariant(XElement.Parse("<mrow><mi>x</mi></mrow>")));
        Assert.Null(MathMlTokenStyleContract.ResolveEffectiveVariant(XElement.Parse("<mi><mglyph/></mi>")));
    }

    [Theory]
    [InlineData("x", true)]
    [InlineData("\U0001D465", true)]
    [InlineData("x\u0302", true)]
    [InlineData("x\u0302\u20D7", true)]
    [InlineData("\U0001D465\u0302", true)]
    [InlineData("x\u1AB0", true)]
    [InlineData("x\u1ABE", true)]
    [InlineData("x\u1DC0", true)]
    [InlineData("x\u20EF", true)]
    [InlineData("xy", false)]
    [InlineData("x\u1ABF", false)]
    [InlineData("x\uFE0F", false)]
    [InlineData("\U0001F469\u200D\U0001F4BB", false)]
    [InlineData("", false)]
    [InlineData("x\n", false)]
    public void SingleCharacterMatchesSupportedMathJaxCombiningContract(string text, bool expected)
    {
        Assert.Equal(expected, MathMlTokenStyleContract.IsSingleMathMlCharacter(text));
    }

    [Fact]
    public void IsolatedSurrogatesAreNotUnicodeCharacters()
    {
        Assert.False(MathMlTokenStyleContract.IsSingleMathMlCharacter(new string((char)0xD835, 1)));
        Assert.False(MathMlTokenStyleContract.IsSingleMathMlCharacter(new string((char)0xDC00, 1)));
    }

    [Theory]
    [InlineData("x\u0302")]
    [InlineData("\U0001D465")]
    [InlineData("\U0001D465\u0302")]
    public void CombiningAndSupplementaryIdentifiersHaveOneConsistentDefault(string text)
    {
        Assert.Equal("italic", MathMlTokenStyleContract.ResolveEffectiveVariant(new XElement("mi", text)));
    }
}
