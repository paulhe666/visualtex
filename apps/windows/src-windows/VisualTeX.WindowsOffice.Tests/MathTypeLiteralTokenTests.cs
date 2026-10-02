using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeLiteralTokenTests
{
    [Theory]
    [InlineData("mn", "100%", @"100\%")]
    [InlineData("mi", "rate%", @"rate\%")]
    [InlineData("mo", "%%", @"\%\%")]
    [InlineData("mi", "a_b", @"a\_b")]
    [InlineData("mo", "a&amp;b", @"a\&b")]
    [InlineData("mi", "a#b", @"a\#b")]
    [InlineData("mi", "a{b}", @"a\{b\}")]
    public void LiteralCharactersInsideOneTokenCannotBecomeTexSyntax(string tag, string text, string expected)
    {
        Assert.Equal(expected + "+2", MathMlToLatexConverter.Convert(
            "<math><" + tag + ">" + text + "</" + tag + "><mo>+</mo><mn>2</mn></math>"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void NativeNumberTokenPercentPreservesTheFollowingEdit(bool inline)
    {
        const string mathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mn>100%</mn><mo>+</mo><mn>2</mn></math>";
        var native = MathTypeMtefCodec.CreateEquationNativeAtFontSize(mathMl, inline, 12);
        var readback = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        Assert.Equal(@"100\%+2", MathMlToLatexConverter.Convert(readback));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(mathMl), MathTypeMtefCodec.SemanticSignature(readback));
    }

    [Fact]
    public void SplitTextPercentInsideARealExpressionMustSurviveStandaloneMtef()
    {
        const string mathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mn>100</mn><mtext>%</mtext>"
            + "<msup><mi mathvariant=\"normal\">e</mi><mrow><mi mathvariant=\"normal\">i</mi><mi>π</mi></mrow></msup>"
            + "<mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>";
        var native = MathTypeMtefCodec.CreateEquationNativeAtFontSize(mathMl, inline: false, 12);
        var readback = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);

        Assert.Equal(
            MathTypeMtefCodec.SemanticSignature(mathMl),
            MathTypeMtefCodec.SemanticSignature(readback));
        // mtext must remain text, so its escaped spelling is deliberately
        // different from the original grouped mn token covered above.
        Assert.Equal(@"100\text{\%}\mathrm{e}^{\mathrm{i}\pi}+1=0",
            MathMlToLatexConverter.Convert(readback));
    }

    [Fact]
    public void ActualLatexCommandsAndStructuralAlignmentRemainGeneratedByTheirOwnNodes()
    {
        Assert.Equal(@"\alpha +100\%", MathMlToLatexConverter.Convert("<math><mi>α</mi><mo>+</mo><mn>100</mn><mo>%</mo></math>"));
        Assert.Equal(@"\begin{matrix}x & 100\%\end{matrix}", MathMlToLatexConverter.Convert("<math><mtable><mtr><mtd><mi>x</mi></mtd><mtd><mn>100%</mn></mtd></mtr></mtable></math>"));
    }
}
