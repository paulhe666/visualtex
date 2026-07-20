using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathMlToLatexConverterTests
{
    [Theory]
    [InlineData(
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>",
        @"\frac{a}{b}")]
    [InlineData(
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msqrt><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><msup><mi>y</mi><mn>2</mn></msup></mrow></msqrt></math>",
        @"\sqrt{x^{2}+y^{2}}")]
    [InlineData(
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup><msup><mi>t</mi><mn>2</mn></msup><mi>d</mi><mi>t</mi></math>",
        @"\int_{0}^{1}t^{2}dt")]
    public void ConvertsCommonWordMathMlToEditableLatex(string mathMl, string expected)
    {
        Assert.Equal(expected, MathMlToLatexConverter.Convert(mathMl));
    }

    [Fact]
    public void ConvertsMatricesAndGreekSymbols()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mfenced open=\"(\" close=\")\"><mtable>"
            + "<mtr><mtd><mi>α</mi></mtd><mtd><mi>β</mi></mtd></mtr>"
            + "<mtr><mtd><mi>γ</mi></mtd><mtd><mi>δ</mi></mtd></mtr>"
            + "</mtable></mfenced></math>";

        var latex = MathMlToLatexConverter.Convert(mathMl);

        Assert.Contains(@"\begin{matrix}", latex);
        Assert.Contains(@"\alpha", latex);
        Assert.Contains(@"\delta", latex);
        Assert.StartsWith(@"\left(", latex);
        Assert.EndsWith(@"\right)", latex);
    }
}
