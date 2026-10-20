using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// Word equation (OMML) -> MathML -> LaTeX for structures users write every day:
// function names, Word limits, single-character accents and aligned equations.
public sealed class OmmlLatexRoundTripTests
{
    private const string Namespaces =
        "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" "
        + "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    private static string Run(string text) => $"<m:r><m:t>{text}</m:t></m:r>";
    private static string Upright(string text) => $"<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>{text}</m:t></m:r>";
    private static string Function(string name, string argument) =>
        $"<m:func><m:fName>{name}</m:fName><m:e>{argument}</m:e></m:func>";
    private static string Accent(string character, string body) =>
        $"<m:acc><m:accPr><m:chr m:val=\"{character}\"/></m:accPr><m:e>{body}</m:e></m:acc>";

    private static string Latex(string body, bool display = false)
    {
        var omml = display
            ? $"<m:oMathPara {Namespaces}><m:oMath>{body}</m:oMath></m:oMathPara>"
            : $"<m:oMath {Namespaces}>{body}</m:oMath>";
        return MathMlToLatexConverter.Convert(WordOmmlConverter.TransformOmmlToMathMl(omml, display));
    }

    [Fact]
    public void KnownFunctionNamesComeBackAsTheirCommandsWithoutTheInvisibleApplication()
    {
        Assert.Equal(@"\sin x", Latex(Function(Upright("sin"), Run("x"))));
        Assert.Equal(@"\sin^{2} \theta", Latex(Function($"<m:sSup><m:e>{Upright("sin")}</m:e><m:sup>{Run("2")}</m:sup></m:sSup>", Run("θ"))));
        Assert.Equal(@"\log_{2} x", Latex(Function($"<m:sSub><m:e>{Upright("log")}</m:e><m:sub>{Run("2")}</m:sub></m:sSub>", Run("x"))));
        Assert.DoesNotContain("⁡", Latex(Function(Upright("cos"), Run("x"))), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("ker")]
    [InlineData("dim")]
    [InlineData("deg")]
    [InlineData("arg")]
    [InlineData("hom")]
    [InlineData("arcsin")]
    [InlineData("arccos")]
    [InlineData("arctan")]
    [InlineData("coth")]
    [InlineData("lg")]
    public void AllTeXFunctionNamesComeBackAsTheirCommands(string name)
    {
        Assert.Equal($@"\{name} x", Latex(Function(Upright(name), Run("x"))));
    }

    [Theory]
    [InlineData("∑", @"\sum_{k}x")]
    [InlineData("∏", @"\prod_{k}x")]
    [InlineData("⋃", @"\bigcup_{k}x")]
    public void WordSumsWithOnlyALowerLimitUseTheScriptForm(string glyph, string expected)
    {
        var nary = $"<m:nary><m:naryPr><m:chr m:val=\"{glyph}\"/><m:limLoc m:val=\"undOvr\"/><m:supHide m:val=\"1\"/></m:naryPr>"
            + $"<m:sub>{Run("k")}</m:sub><m:sup/><m:e>{Run("x")}</m:e></m:nary>";
        Assert.Equal(expected, Latex(nary, display: true));
    }

    [Fact]
    public void AnUprightWordStaysOneGroup()
    {
        Assert.Equal(@"\mathrm{Tr}x", Latex(Upright("Tr") + Run("x")));
        Assert.Equal(@"\mathrm{im}x", Latex(Upright("im") + Run("x")));
        Assert.Equal(@"H_{\mathrm{d}R}", Latex($"<m:sSub><m:e>{Run("H")}</m:e><m:sub>{Upright("d")}{Run("R")}</m:sub></m:sSub>"));
    }

    [Fact]
    public void UnknownFunctionNamesKeepTheirApplicationMarker()
    {
        Assert.Contains("⁡", Latex(Function(Upright("foo"), Run("x"))), StringComparison.Ordinal);
    }

    [Fact]
    public void WordLimitsComeBackAsLimitOperatorsWithTheLimitUnderneath()
    {
        var limit = $"<m:limLow><m:e>{Upright("lim")}</m:e><m:lim>{Run("x")}{Upright("→0")}</m:lim></m:limLow>";
        Assert.Equal(@"\lim\limits_{x\to 0}", Latex(limit));
        // Word's own equation editor wraps the limit in a function.
        Assert.Equal(@"\lim\limits_{x\to 0} f", Latex(Function(limit, Run("f"))));
    }

    [Theory]
    [InlineData("⃗", "F", @"\vec{F}")]
    [InlineData("⃗", "AB", @"\overrightarrow{AB}")]
    [InlineData("̂", "x", @"\hat{x}")]
    [InlineData("̂", "AB", @"\widehat{AB}")]
    [InlineData("̃", "y", @"\tilde{y}")]
    public void WordAccentsOverOneCharacterAreTheFixedAccent(string character, string body, string expected)
    {
        Assert.Equal(expected, Latex(Accent(character, Run(body))));
    }

    [Fact]
    public void AlignedMatrixComesBackAsAligned()
    {
        const string columns =
            "<m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val=\"1\"/><m:mcJc m:val=\"right\"/></m:mcPr></m:mc>"
            + "<m:mc><m:mcPr><m:count m:val=\"1\"/><m:mcJc m:val=\"left\"/></m:mcPr></m:mc></m:mcs></m:mPr>";
        var matrix = $"<m:m>{columns}<m:mr><m:e>{Run("a")}</m:e><m:e>{Run("=b")}</m:e></m:mr>"
            + $"<m:mr><m:e/><m:e>{Run("=c")}</m:e></m:mr></m:m>";
        Assert.Equal(@"\begin{aligned}a & =b \\ & =c\end{aligned}", Latex(matrix, display: true));
    }

    [Fact]
    public void CenteredMatrixStaysAMatrix()
    {
        var matrix = $"<m:m><m:mr><m:e>{Run("1")}</m:e><m:e>{Run("2")}</m:e></m:mr><m:mr><m:e>{Run("3")}</m:e><m:e>{Run("4")}</m:e></m:mr></m:m>";
        Assert.Equal(@"\begin{matrix}1 & 2 \\ 3 & 4\end{matrix}", Latex(matrix, display: true));
    }

    [Fact]
    public void EquationArrayAlignmentPointsBecomeAlignedColumns()
    {
        var array = $"<m:eqArr><m:e>{Run("a")}{Run("&amp;=b+c")}</m:e><m:e>{Run("&amp;=d")}</m:e></m:eqArr>";
        Assert.Equal(@"\begin{aligned}a&=b+c \\ &=d\end{aligned}", Latex(array, display: true));
    }

    [Theory]
    [InlineData("'", @"f^{\prime}")]
    [InlineData("''", @"f^{\prime \prime}")]
    [InlineData("′", @"f^{\prime}")]
    public void UprightPrimeSuperscriptsComeBackAsPrimes(string primes, string expected)
    {
        // VisualTeX for Mac writes f' as an upright ASCII apostrophe superscript.
        Assert.Equal(expected, Latex($"<m:sSup><m:e>{Run("f")}</m:e><m:sup>{Upright(primes)}</m:sup></m:sSup>"));
    }

    // Word equations store no VisualTeX source, so what comes back is written
    // the way people type it. omml_latex/source_style.rs uses the same table.
    [Theory]
    [InlineData(@"x^{2}+y^{2}=z^{2}", @"x^2+y^2=z^2")]
    [InlineData(@"\int_{0}^{1}x^{2}\mathrm{d}x=\frac{1}{3}", @"\int_0^1 x^2\mathrm{d}x=\frac{1}{3}")]
    [InlineData(@"\sin^{2} \theta +\cos^{2} \theta =1", @"\sin^2\theta+\cos^2\theta=1")]
    [InlineData(@"x=\frac{-b\pm \sqrt{b^{2}-4ac}}{2a}", @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}")]
    [InlineData(@"f^{\prime}(x)", @"f'(x)")]
    [InlineData(@"f^{\prime \prime}", @"f''")]
    [InlineData(@"\left( \begin{matrix}1 & 2 \\ 3 & 4\end{matrix}\right)", @"\begin{pmatrix}1 & 2 \\ 3 & 4\end{pmatrix}")]
    [InlineData(@"\left[ \begin{matrix}a\end{matrix}\right]", @"\begin{bmatrix}a\end{bmatrix}")]
    [InlineData(@"f(x)=\left\{ \begin{matrix}x, & x\ge 0 \\ -x, & x<0\end{matrix}\right.", @"f(x)=\begin{cases}x, & x\ge 0 \\ -x, & x<0\end{cases}")]
    [InlineData(@"\left( \begin{matrix}1\end{matrix}\right\}", @"\left(\begin{matrix}1\end{matrix}\right\}")]
    [InlineData(@"\left( \frac{a}{b}\right)^{2}+\binom{n}{k}", @"\left(\frac{a}{b}\right)^2+\binom{n}{k}")]
    [InlineData(@"\lim\limits_{x\to 0} \frac{\sin x}{x}=1", @"\lim\limits_{x\to 0}\frac{\sin x}{x}=1")]
    [InlineData(@"\alpha +\beta =\gamma", @"\alpha+\beta=\gamma")]
    [InlineData(@"a_{ij}+x^{10}+e^{i\pi}", @"a_{ij}+x^{10}+e^{i\pi}")]
    [InlineData(@"\mathrm{H}_{2}\mathrm{O}", @"\mathrm{H}_2\mathrm{O}")]
    [InlineData(@"\text{a b} +1", @"\text{a b}+1")]
    [InlineData(@"\begin{aligned}a&=b+c \\ &=d\end{aligned}", @"\begin{aligned}a&=b+c \\ &=d\end{aligned}")]
    // PowerPoint stores math letters as mathematical alphanumerics (𝑎).
    [InlineData(@"\frac{𝑎}{𝑏}", @"\frac{a}{b}")]
    [InlineData(@"\sqrt{{𝑏}^2-4𝑎𝑐}", @"\sqrt{b^2-4ac}")]
    [InlineData(@"\int_{0}^{1}{𝑥}^{2}𝑑𝑥=1", @"\int_0^1 x^2 dx=1")]
    [InlineData(@"\sin𝑥+\cos^{2}𝜃", @"\sin x+\cos^2\theta")]
    [InlineData(@"{𝑎}_{𝑖𝑗}", @"a_{ij}")]
    [InlineData(@"𝐀+𝔸+ℎ+𝟏", @"\mathbf{A}+\mathbb{A}+h+\mathbf{1}")]
    [InlineData(@"𝜋𝑟^2", @"\pi r^2")]
    [InlineData(@"\mathrm{x}^{2}", @"\mathrm{x}^2")]
    public void WordEquationLatexIsWrittenTheWayPeopleTypeIt(string converted, string expected)
    {
        Assert.Equal(expected, MathMlToLatexConverter.WordEquationSourceStyle(converted));
    }

    [Fact]
    public void WordEquationConversionUsesTheSourceStyle()
    {
        var omml = $"<m:oMath {Namespaces}><m:sSup><m:e>{Run("x")}</m:e><m:sup>{Run("2")}</m:sup></m:sSup></m:oMath>";
        Assert.Equal("x^2", MathMlToLatexConverter.ConvertWordEquation(
            WordOmmlConverter.TransformOmmlToMathMl(omml, false)));
    }
}
