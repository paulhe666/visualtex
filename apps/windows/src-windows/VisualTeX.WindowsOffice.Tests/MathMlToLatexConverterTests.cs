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

    [Theory]
    [InlineData("lim", @"\lim_{x\to 0}")]
    [InlineData("max", @"\max_{x\in A}")]
    [InlineData("min", @"\min_{x\in A}")]
    [InlineData("sup", @"\sup_{x\in A}")]
    [InlineData("inf", @"\inf_{x\in A}")]
    [InlineData("limsup", @"\limsup_{n\to \infty}")]
    [InlineData("liminf", @"\liminf_{n\to \infty}")]
    public void ConvertsMathMlUnderScriptsOnLimitLikeOperatorsToNativeLatexSubscripts(
        string operatorName,
        string expected)
    {
        var script = operatorName.StartsWith("lim", StringComparison.Ordinal)
            ? "<mrow><mi>n</mi><mo>→</mo><mi>∞</mi></mrow>"
            : "<mrow><mi>x</mi><mo>∈</mo><mi>A</mi></mrow>";
        if (operatorName == "lim")
            script = "<mrow><mi>x</mi><mo>→</mo><mn>0</mn></mrow>";
        var mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + $"<munder><mi mathvariant=\"normal\">{operatorName}</mi>{script}</munder></math>";

        Assert.Equal(expected, MathMlToLatexConverter.Convert(mathMl));
    }

    [Fact]
    public void KeepsOrdinaryMathMlUnderScriptAsUnderset()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<munder><mi>x</mi><mi>i</mi></munder></math>";

        Assert.Equal(@"\underset{i}{x}", MathMlToLatexConverter.Convert(mathMl));
    }

    [Fact]
    public void ConvertsLooseNoBarFractionFenceToBinomial()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mo>(</mo><mfrac linethickness=\"0\"><mi>n</mi><mi>k</mi></mfrac><mo>)</mo>"
            + "<mo>=</mo><mn>1</mn></math>";

        Assert.Equal(
            @"\binom{n}{k}=1",
            MathMlToLatexConverter.Convert(mathMl));
    }

    [Fact]
    public void ConvertsMathJaxWrappedOpenCloseNoBarFractionToBinomial()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mrow data-mjx-texclass=\"ORD\">"
            + "<mrow data-mjx-texclass=\"OPEN\"><mo minsize=\"2.047em\" maxsize=\"2.047em\">(</mo></mrow>"
            + "<mfrac linethickness=\"0\"><mi>n</mi><mi>k</mi></mfrac>"
            + "<mrow data-mjx-texclass=\"CLOSE\"><mo minsize=\"2.047em\" maxsize=\"2.047em\">)</mo></mrow>"
            + "</mrow></math>";

        Assert.Equal(@"\binom{n}{k}", MathMlToLatexConverter.Convert(mathMl));
    }

    [Fact]
    public void ConvertsMathJaxFencedNoBarFractionThroughTransparentRowToBinomial()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mfenced open=\"(\" close=\")\"><mrow><mfrac linethickness=\"0\"><mi>n</mi><mi>k</mi></mfrac></mrow></mfenced>"
            + "</math>";

        Assert.Equal(@"\binom{n}{k}", MathMlToLatexConverter.Convert(mathMl));
    }

    [Fact]
    public void ConvertsLooseMathTypePileFenceToBinomial()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mo>(</mo><mtable data-mtef-pile=\"true\">"
            + "<mtr><mtd><mi>n</mi></mtd></mtr>"
            + "<mtr><mtd><mi>k</mi></mtd></mtr>"
            + "</mtable><mo>)</mo><mo>=</mo><mn>1</mn></math>";

        Assert.Equal(
            @"\binom{n}{k}=1",
            MathMlToLatexConverter.Convert(mathMl));
    }

    [Fact]
    public void DoesNotConvertOrdinaryFencedFractionToBinomial()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mo>(</mo><mfrac><mi>n</mi><mi>k</mi></mfrac><mo>)</mo>"
            + "</math>";

        Assert.Equal(
            @"(\frac{n}{k})",
            MathMlToLatexConverter.Convert(mathMl));
    }

    [Theory]
    [InlineData("'", @"f^{\prime}")]
    [InlineData("′", @"f^{\prime}")]
    [InlineData("″", @"f^{\prime\prime}")]
    public void ConvertsWordPrimeSpellingsToLatexPrime(string prime, string expected)
    {
        var mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + $"<msup><mi>f</mi><mi>{prime}</mi></msup></math>";

        Assert.Equal(expected, MathMlToLatexConverter.Convert(mathMl));
    }

    [Fact]
    public void PreservesExplicitUprightLatinIdentifiers()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mi mathvariant=\"normal\">e</mi>"
            + "<msup><mi mathvariant=\"normal\">i</mi><mi>π</mi></msup>"
            + "<mi>x</mi></math>";

        Assert.Equal(
            @"\mathrm{e}\mathrm{i}^{\pi}x",
            MathMlToLatexConverter.Convert(mathMl));
    }

    [Theory]
    [InlineData("mover", "⏞", @"\overbrace{a+b}^{n}")]
    [InlineData("munder", "⏟", @"\underbrace{a+b}_{n}")]
    public void ConvertsAnnotatedHorizontalBracesThroughSingleTransparentWrapper(
        string innerName,
        string marker,
        string expected)
    {
        var outerName = innerName;
        var mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + $"<{outerName}><mrow><{innerName}><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mo stretchy=\"true\">{marker}</mo></{innerName}></mrow><mi>n</mi></{outerName}></math>";

        Assert.Equal(expected, MathMlToLatexConverter.Convert(mathMl));
    }

    [Theory]
    [InlineData("mover", "―", @"\overline{AB}")]
    [InlineData("munder", "―", @"\underline{AB}")]
    public void ConvertsMathJaxHorizontalBarAccentToLineCommands(
        string elementName,
        string marker,
        string expected)
    {
        var mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + $"<{elementName}><mrow><mi>A</mi><mi>B</mi></mrow><mo>{marker}</mo></{elementName}></math>";

        Assert.Equal(expected, MathMlToLatexConverter.Convert(mathMl));
    }

    [Theory]
    [InlineData("|", "|", @"\left| x\right|")]
    [InlineData("‖", "‖", @"\left\| x\right\|")]
    public void ConvertsFencedAbsoluteAndNormDelimitersWithMatchingSides(
        string open,
        string close,
        string expected)
    {
        var mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + $"<mfenced open=\"{open}\" close=\"{close}\"><mi>x</mi></mfenced></math>";

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
