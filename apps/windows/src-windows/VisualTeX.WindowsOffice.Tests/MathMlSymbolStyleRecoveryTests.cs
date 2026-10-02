using VisualTeX.WordVsto;
namespace VisualTeX.WindowsOffice.Tests;
public sealed class MathMlSymbolStyleRecoveryTests
{
    [Theory]
    [InlineData("∠", @"\mathit{\angle}")]
    [InlineData("∇", @"\mathit{\nabla}")]
    [InlineData("∞", @"\mathit{\infty}")]
    [InlineData("℘", @"\mathit{\wp}")]
    public void PreservesImplicitSingleSymbolIdentifierItalic(string symbol, string expected)
        => Assert.Equal(expected, MathMlToLatexConverter.Convert("<math><mi>" + symbol + "</mi></math>"));

    [Theory]
    [InlineData("italic", @"\mathit{\parallel}")]
    [InlineData("bold", @"\mathbf{\parallel}")]
    [InlineData("bold-italic", @"\mathbfit{\parallel}")]
    public void PreservesInheritedOperatorStyle(string variant, string expected)
        => Assert.Equal(expected, MathMlToLatexConverter.Convert("<math><mstyle mathvariant='" + variant + "'><mo>∥</mo></mstyle></math>"));

    [Fact]
    public void PreservesInnerNormalReset()
        => Assert.Equal(@"\parallel", MathMlToLatexConverter.Convert("<math><mstyle mathvariant='italic'><mo mathvariant='normal'>∥</mo></mstyle></math>"));

    [Fact]
    public void LeavesOrdinaryIdentifierDefaultsUnwrapped()
        => Assert.Equal(@"x\alpha", MathMlToLatexConverter.Convert("<math><mi>x</mi><mi>α</mi></math>"));
}
