using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// Formula OCR writes numbers digit by digit ("1 0 0", "3 . 1 4"); MathJax keeps
// such digits as separate <mn> tokens, while MathType's read-back groups adjacent
// digits into one number. Saving them as MathType must not be refused as a
// content change (PowerPoint "cannot be saved as MathType without changing its
// content" after appending an OCR line), and real digit or style changes must
// still be caught.
public sealed class MathTypeSplitNumberRoundTripTests
{
    private static string Sig(string mathMl) => MathTypeMtefCodec.SemanticSignature(mathMl);

    private static void AssertRoundTrip(string mathMl)
    {
        var native = MathTypeMtefFontSize.CreateEquationNativeAtFontSize(mathMl, inline: false, 24);
        var readBack = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        Assert.Equal(Sig(mathMl), Sig(readBack));
        var rewritten = MathTypeMtefFontSize.RewriteEquationNativeAtFontSize(native.EquationNative, mathMl, false, 24);
        Assert.Equal(Sig(mathMl), Sig(MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative)));
    }

    [Fact]
    public void DigitsWrittenApartRoundTrip()
        => AssertRoundTrip("<math display='block'><mi>R</mi><mo>=</mo><mn>1</mn><mn>0</mn><mi mathvariant='normal'>&#x3A9;</mi></math>");

    [Fact]
    public void DecimalWrittenApartRoundTrips()
        => AssertRoundTrip("<math display='block'><mi>&#x3C0;</mi><mo>&#x2248;</mo><mn>3</mn><mo>.</mo><mn>1</mn><mn>4</mn></math>");

    [Fact]
    public void PercentAndScriptsWrittenApartRoundTrip()
        => AssertRoundTrip(
            "<math display='block'><mi>&#x3B7;</mi><mo>=</mo><mn>1</mn><mn>0</mn><mn>0</mn><mi mathvariant='normal'>%</mi>"
            + "<mo>&#xD7;</mo><mn>1</mn><msup><mn>0</mn><mrow><mn>2</mn><mn>3</mn></mrow></msup></math>");

    [Fact]
    public void NuclideWrittenApartRoundTrips()
        => AssertRoundTrip(
            "<math display='block'><msubsup><mrow></mrow><mrow><mn>9</mn><mn>2</mn></mrow><mrow><mn>2</mn><mn>3</mn><mn>8</mn></mrow></msubsup>"
            + "<mrow><mi mathvariant='normal'>U</mi></mrow><mo stretchy='false'>&#x2192;</mo>"
            + "<msubsup><mrow></mrow><mrow><mn>9</mn><mn>0</mn></mrow><mrow><mn>2</mn><mn>3</mn><mn>4</mn></mrow></msubsup>"
            + "<mrow><mi mathvariant='normal'>T</mi><mi mathvariant='normal'>h</mi></mrow></math>");

    [Fact]
    public void SplitAndJoinedDigitsAreTheSameNumber()
        => Assert.Equal(Sig("<math><mn>100</mn></math>"), Sig("<math><mn>1</mn><mn>0</mn><mn>0</mn></math>"));

    [Fact]
    public void ChangedOrLostDigitsStillDiffer()
    {
        Assert.NotEqual(Sig("<math><mn>100</mn></math>"), Sig("<math><mn>10</mn></math>"));
        Assert.NotEqual(Sig("<math><mn>12</mn></math>"), Sig("<math><mn>21</mn></math>"));
        Assert.NotEqual(Sig("<math><mn>1</mn><mo>.</mo><mn>5</mn></math>"), Sig("<math><mn>15</mn></math>"));
    }

    [Fact]
    public void DigitStyleStillCounts()
        => Assert.NotEqual(
            Sig("<math><mn mathvariant='bold'>1</mn><mn>0</mn></math>"),
            Sig("<math><mn>10</mn></math>"));
}
