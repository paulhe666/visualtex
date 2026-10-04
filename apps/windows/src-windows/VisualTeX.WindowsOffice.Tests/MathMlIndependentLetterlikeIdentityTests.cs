using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathMlIndependentLetterlikeIdentityTests
{
    public static IEnumerable<object[]> DistinctStyles()
    {
        var styles = new[] { "normal", "italic", "bold", "bold-italic" };
        for (var i = 0; i < styles.Length; i++)
            for (var j = i + 1; j < styles.Length; j++) yield return new object[] { styles[i], styles[j] };
    }

    [Theory]
    [MemberData(nameof(DistinctStyles))]
    public void IndependentEllNeverLosesADeclaredStyle(string first, string second)
        => Assert.NotEqual(Signature(first), Signature(second));

    [Theory]
    [InlineData("normal", "script")]
    [InlineData("italic", "script")]
    [InlineData("bold", "bold-script")]
    [InlineData("bold-italic", "bold-script")]
    public void IndependentEllIsNotAMathematicalAlphabetAlias(string style, string alphabet)
    {
        Assert.NotEqual(Signature(style), MathTypeMtefCodec.SemanticSignature($"<math><mi mathvariant='{alphabet}'>l</mi></math>"));
        Assert.NotEqual(Signature(style), MathTypeMtefCodec.SemanticSignature($"<math><mi mathvariant='{style}'>𝓁</mi></math>"));
    }

    [Theory]
    [InlineData("normal")]
    [InlineData("italic")]
    [InlineData("bold")]
    [InlineData("bold-italic")]
    public void ExistingNativeScalarAndStyleRemainAnExactRoundTrip(string style)
    {
        var input = $"<math><mi mathvariant='{style}'>ℓ</mi></math>";
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(input, false, 12);
        var readback = MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);
        var token = Assert.Single(XElement.Parse(readback).Descendants(), e => e.Name.LocalName == "mi");
        Assert.Equal("ℓ", token.Value);
        Assert.Equal("8467", (string?)token.Attribute("data-mtef-mtcode"));
        Assert.Equal("108", (string?)token.Attribute("data-mtef-encoded8"));
        Assert.Equal(style, MathMlTokenStyleContract.ResolveEffectiveVariant(token));
        Assert.Equal(Signature(style), MathTypeMtefCodec.SemanticSignature(readback));
    }

    [Fact]
    public void InheritedAndDirectStylesAgreeWithoutChangingTheScalar()
        => Assert.Equal(Signature("bold"), MathTypeMtefCodec.SemanticSignature("<math><mstyle mathvariant='bold'><mi>ℓ</mi></mstyle></math>"));

    [Theory]
    [InlineData("<mi>ℓ</mi>", "italic")]
    [InlineData("<mstyle mathvariant='italic'><mi>ℓ</mi></mstyle>", "italic")]
    [InlineData("<mstyle mathvariant='bold'><mi mathvariant='normal'>ℓ</mi></mstyle>", "normal")]
    [InlineData("<mstyle mathvariant='normal'><mi mathvariant='italic'>ℓ</mi></mstyle>", "italic")]
    public void EffectiveStyleDoesNotDependOnSerializerAttributeOmission(string body, string expected)
    {
        var input = "<math>" + body + "</math>";
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(input, false, 12);
        var readback = MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);
        var token = Assert.Single(XElement.Parse(readback).Descendants(), e => e.Name.LocalName == "mi");
        Assert.Equal("ℓ", token.Value);
        Assert.Equal(expected, MathMlTokenStyleContract.ResolveEffectiveVariant(token));
        Assert.Equal(Signature(expected), MathTypeMtefCodec.SemanticSignature(readback));
    }

    private static string Signature(string style)
        => MathTypeMtefCodec.SemanticSignature($"<math><mi mathvariant='{style}'>ℓ</mi></math>");
}
