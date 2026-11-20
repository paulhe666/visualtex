using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathMlItalicSourceAlphabetTests
{
    [Theory]
    [InlineData("A", "A")]
    [InlineData("x", "x")]
    [InlineData("α", @"\alpha")]
    [InlineData("π", @"\pi")]
    public void MaterializedMathItalicDoesNotInventTexMathitAlphabet(string token, string expected)
    {
        var input = $"<math><mi mathvariant='italic'>{token}</mi></math>";
        var element = XDocument.Parse(input).Root!.Elements().Single();
        Assert.Equal("italic", MathMlTokenStyleContract.ResolveEffectiveVariant(element));
        Assert.Equal(expected, MathMlToLatexConverter.Convert(input));
        // A native italic character-style bit materializes the same MathML
        // style; it is not evidence of TeX's distinct text-italic alphabet.
        element.SetAttributeValue("data-mtef-character-style", 2);
        Assert.Equal(expected, MathMlToLatexConverter.Convert(element.Parent!.ToString()));
        element.SetAttributeValue("mathvariant", "normal");
        Assert.NotEqual("italic", MathMlTokenStyleContract.ResolveEffectiveVariant(element));
    }

    [Theory]
    [InlineData("A")]
    [InlineData("x")]
    [InlineData("α")]
    [InlineData("π")]
    public void GenuineTexMathitTokenProvenanceIsNotDiscarded(string token)
    {
        var input = $"<math><mi mathvariant='italic' data-mjx-variant='-tex-mathit'>{token}</mi></math>";
        Assert.Equal(@"\mathit{" + token + "}", MathMlToLatexConverter.Convert(input));
    }

    [Fact]
    public void OrdinaryStyleAndAlphabetProvenanceHaveSeparateScopes()
    {
        Assert.Equal("A", MathMlToLatexConverter.Convert(
            "<math><mstyle mathvariant='italic' data-mjx-variant='-tex-mathit'><mi>A</mi></mstyle></math>"));
        Assert.Equal(@"\mathrm{A}", MathMlToLatexConverter.Convert(
            "<math><mi mathvariant='normal'>A</mi></math>"));
        Assert.Equal(@"\mathbfit{A}", MathMlToLatexConverter.Convert(
            "<math><mi mathvariant='bold-italic'>A</mi></math>"));
    }
}
