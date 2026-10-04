using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordNativeStyleComparisonTests
{
    private static string Signature(string body) => WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature("<math xmlns='http://www.w3.org/1998/Math/MathML'>" + body + "</math>");

    [Theory]
    [InlineData("bold", "bold-italic")]
    [InlineData("sans-serif", "sans-serif-italic")]
    [InlineData("bold-sans-serif", "sans-serif-bold-italic")]
    public void IdentifierItalicIsNotDiscarded(string plain, string italic)
    {
        Assert.NotEqual(Signature("<mi mathvariant='" + plain + "'>x</mi>"), Signature("<mi mathvariant='" + italic + "'>x</mi>"));
    }

    [Fact]
    public void MathematicalSymbolBoldAndBoldItalicRemainDifferent()
    {
        Assert.NotEqual(Signature("<mi mathvariant='bold'>∠</mi>"), Signature("<mi mathvariant='bold-italic'>∠</mi>"));
    }

    [Fact]
    public void PlainRelationIsNotAnImplicitItalicIdentifier()
    {
        Assert.NotEqual(Signature("<mo>∥</mo>"), Signature("<mi>∥</mi>"));
    }

    [Fact]
    public void ImplicitAndExplicitSingleIdentifierItalicAgree()
    {
        Assert.Equal(Signature("<mi>∡</mi>"), Signature("<mi mathvariant='italic'>∡</mi>"));
    }

    [Fact]
    public void InheritedBoldItalicMatchesExplicitBoldItalic()
    {
        Assert.Equal(Signature("<mstyle mathvariant='bold-italic'><mi>x</mi></mstyle>"), Signature("<mi mathvariant='bold-italic'>x</mi>"));
    }
}
