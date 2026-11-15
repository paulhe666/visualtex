using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// Editing a one-line MathType equation into several lines (PowerPoint: reopen an
// inserted formula and append an OCR line) rewrites its root LINE as an aligned
// PILE. The source's own font definitions sat at the start of that LINE; dropped
// with it, every later style index pointed one definition too far and the save
// was refused ("cannot be saved as MathType without changing its content"):
// upper Greek read back as Latin letters, upright letters as italic.
public sealed class MathTypeRewriteIntoPileTests
{
    private static string Sig(string mathMl) => MathTypeMtefCodec.SemanticSignature(mathMl);

    private const string Units =
        "<mrow><mi>kg</mi></mrow><mo>&#x22C5;</mo><mrow><mi mathvariant='normal'>m</mi></mrow>"
        + "<mrow><mo>/</mo></mrow><msup><mrow><mi mathvariant='normal'>s</mi></mrow><mrow><mn>2</mn></mrow></msup>";

    private const string Discriminant =
        "<mi mathvariant='normal'>&#x394;</mi><mo>=</mo><msup><mi>b</mi><mrow><mn>2</mn></mrow></msup>"
        + "<mo>&#x2212;</mo><mn>4</mn><mi>a</mi><mi>c</mi>";

    private const string Iron =
        "<mrow><mi mathvariant='normal'>F</mi><mi mathvariant='normal'>e</mi></mrow><mo>+</mo>"
        + "<msub><mrow><mi mathvariant='normal'>C</mi><mi mathvariant='normal'>u</mi><mi mathvariant='normal'>S</mi>"
        + "<mi mathvariant='normal'>O</mi></mrow><mrow><mn>4</mn></mrow></msub>";

    private const string UpperGreek =
        "<mi mathvariant='normal'>&#x393;</mi><mi mathvariant='normal'>&#x394;</mi><mi mathvariant='normal'>&#x398;</mi>"
        + "<mi mathvariant='normal'>&#x3A3;</mi><mi mathvariant='normal'>&#x3A9;</mi>";

    private static string Line(string row) => "<math display='block'>" + row + "</math>";

    private static string Rows(params string[] rows) =>
        "<math display='block'><mtable displaystyle='true' columnalign='right left' columnspacing='0em' rowspacing='3pt'>"
        + string.Concat(rows.Select(row => "<mtr><mtd/><mtd>" + row + "</mtd></mtr>"))
        + "</mtable></math>";

    private static void AssertRewrite(string sourceMathMl, string editedMathMl)
    {
        var source = MathTypeMtefFontSize.CreateEquationNativeAtFontSize(sourceMathMl, inline: false, 24);
        Assert.Equal(Sig(sourceMathMl), Sig(MathTypeMtefCodec.ReadEquationNativeMathMl(source.EquationNative)));
        var edited = MathTypeMtefFontSize.RewriteEquationNativeAtFontSize(source.EquationNative, editedMathMl, false, 24);
        Assert.Equal(Sig(editedMathMl), Sig(MathTypeMtefCodec.ReadEquationNativeMathMl(edited.EquationNative)));
    }

    [Fact]
    public void AppendedLineKeepsTheSourceFonts()
        => AssertRewrite(Line(Units), Rows(Units, Discriminant));

    [Fact]
    public void PrependedLineKeepsTheSourceFonts()
        => AssertRewrite(Line(Units), Rows(Discriminant, Units));

    [Fact]
    public void UprightGreekSourceGainsAnUprightLatinLine()
        => AssertRewrite(Line(Discriminant), Rows(Discriminant, Iron));

    [Fact]
    public void UpperGreekLineAfterAnUprightSource()
        => AssertRewrite(Line(Iron), Rows(Iron, UpperGreek));

    [Fact]
    public void PileBackToOneLine()
    {
        var source = MathTypeMtefFontSize.CreateEquationNativeAtFontSize(Rows(Units, Discriminant), inline: false, 24);
        var edited = MathTypeMtefFontSize.RewriteEquationNativeAtFontSize(source.EquationNative, Line(Discriminant), false, 24);
        Assert.Equal(Sig(Line(Discriminant)), Sig(MathTypeMtefCodec.ReadEquationNativeMathMl(edited.EquationNative)));
    }
}
