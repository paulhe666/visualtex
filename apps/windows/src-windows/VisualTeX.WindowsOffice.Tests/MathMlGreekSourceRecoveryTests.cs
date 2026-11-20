using VisualTeX.WordVsto;
using Xunit;
namespace VisualTeX.WindowsOffice.Tests;
public sealed class MathMlGreekSourceRecoveryTests
{
    [Theory]
    [InlineData("α", "normal", @"\mathrm{α}")]
    [InlineData("Γ", "normal", @"\mathrm{Γ}")]
    [InlineData("α", "upright", @"\mathrm{α}")]
    [InlineData("α", "italic", @"\alpha")]
    [InlineData("Γ", "italic", @"\mathit{Γ}")]
    [InlineData("α", "bold", @"\mathbf{α}")]
    [InlineData("Γ", "bold", @"\boldsymbol{\Gamma}")]
    [InlineData("α", "bold-italic", @"\boldsymbol{\alpha}")]
    [InlineData("Γ", "bold-italic", @"\boldsymbol{\mathit{Γ}}")]
    [InlineData("αβ", "normal", @"\mathrm{αβ}")]
    [InlineData("Γα", "italic", @"\mathit{Γα}")]
    [InlineData("ΓΔ", "bold", @"\mathbf{ΓΔ}")]
    [InlineData("Α", "bold", @"\mathbf{Α}")]
    public void ExplicitGreekStyleUsesFrontendCompatibleSource(string text,string variant,string expected)
        => Assert.Equal(expected,MathMlToLatexConverter.Convert($"<math><mi mathvariant='{variant}'>{text}</mi></math>"));

    [Theory]
    [InlineData("α", @"\alpha")]
    [InlineData("Γ", @"\mathit{Γ}")]
    [InlineData("ΓΔ", @"\mathrm{ΓΔ}")]
    [InlineData("αβ", @"\mathrm{αβ}")]
    public void UnstyledGreekPreservesEffectiveMathMlDefault(string text,string expected)
        => Assert.Equal(expected,MathMlToLatexConverter.Convert($"<math><mi>{text}</mi></math>"));

    [Fact]
    public void NearestInheritedVariantAndChildOverridesRemainDistinct()
        => Assert.Equal(@"\mathrm{α}\beta \mathbf{γ}",
            MathMlToLatexConverter.Convert("<math><mstyle mathvariant='normal'><mi>α</mi><mstyle mathvariant='italic'><mi>β</mi></mstyle><mi mathvariant='bold'>γ</mi></mstyle></math>"));

    [Theory]
    [InlineData(0, @"\mathrm{α}")]
    [InlineData(1, @"\mathbf{α}")]
    [InlineData(2, @"\mathit{α}")]
    [InlineData(3, @"\boldsymbol{\alpha}")]
    public void NativeCharacterStyleBitsAreAuthoritative(int bits,string expected)
        => Assert.Equal(expected,MathMlToLatexConverter.Convert($"<math><mi mathvariant='normal' data-mtef-character-style='{bits}'>α</mi></math>"));
}
