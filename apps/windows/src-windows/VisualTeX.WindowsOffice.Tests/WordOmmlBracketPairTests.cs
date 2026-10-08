using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlBracketPairTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string Ns = "http://www.w3.org/1998/Math/MathML";

    private static XDocument Omml(string body) =>
        XDocument.Parse(WordOmmlConverter.TransformMathMlToOmml(
            $"<math xmlns='{Ns}' display='block'>{body}</math>"));

    private static int Groups(XDocument omml) => omml.Descendants(M + "d").Count();

    private static bool HasLiteralBracket(XDocument omml) =>
        omml.Descendants(M + "t").Any(text => text.Value.Length > 0
            && text.Value.All(character => "()[]{}".IndexOf(character) >= 0));

    [Fact]
    public void ClosingBracketThatCarriesAPowerIsGroupedWithItsOpeningBracket()
    {
        // MathJax output for (a+b)^n: the closing bracket is the script's base.
        var omml = Omml(
            "<mo>(</mo><mi>a</mi><mo>+</mo><mi>b</mi>"
            + "<msup><mrow><mo>)</mo></mrow><mi>n</mi></msup>");
        Assert.Equal(1, Groups(omml));
        Assert.False(HasLiteralBracket(omml));
        var power = Assert.Single(omml.Descendants(M + "sSup"));
        Assert.NotNull(power.Element(M + "e")?.Element(M + "d"));
    }

    [Fact]
    public void FunctionArgumentBracketsBecomeAGroup()
    {
        var omml = Omml("<mi>f</mi><mo>(</mo><mi>x</mi><mo>)</mo>");
        Assert.Equal(1, Groups(omml));
        Assert.False(HasLiteralBracket(omml));
    }

    [Fact]
    public void NestedBracketsBecomeNestedGroups()
    {
        var omml = Omml(
            "<mo>(</mo><mo>(</mo><mi>a</mi><mo>+</mo><mi>b</mi><mo>)</mo><mo>+</mo><mi>c</mi><mo>)</mo>");
        Assert.Equal(2, Groups(omml));
        Assert.False(HasLiteralBracket(omml));
    }

    [Fact]
    public void BracketsOfOtherKindsAreGrouped()
    {
        var omml = Omml("<mo>[</mo><mi>a</mi><mo>]</mo><mo>+</mo><mo>{</mo><mi>b</mi><mo>}</mo>");
        Assert.Equal(2, Groups(omml));
    }

    [Fact]
    public void HandSizedBracketsBecomeAdaptiveGroups()
    {
        var omml = Omml(
            "<mrow data-mjx-texclass='OPEN'><mo minsize='1.2em' maxsize='1.2em'>(</mo></mrow>"
            + "<mfrac><mi>a</mi><mi>b</mi></mfrac>"
            + "<mrow data-mjx-texclass='CLOSE'><mo minsize='1.2em' maxsize='1.2em'>)</mo></mrow>");
        Assert.Equal(1, Groups(omml));
        Assert.False(HasLiteralBracket(omml));
    }

    [Fact]
    public void BareBracketsInsideAnExplicitLeftRightBodyAreGroupedToo()
    {
        var omml = Omml(
            "<mrow data-mjx-texclass='INNER'><mo data-mjx-texclass='OPEN'>(</mo>"
            + "<mi>f</mi><mo>(</mo><mi>x</mi><mo>)</mo><mo>+</mo><mn>1</mn>"
            + "<mo data-mjx-texclass='CLOSE'>)</mo></mrow>");
        Assert.Equal(2, Groups(omml));
        Assert.False(HasLiteralBracket(omml));
    }

    [Fact]
    public void MismatchedIntervalBracketsStayLiteral()
    {
        var omml = Omml("<mo>[</mo><mn>0</mn><mo>,</mo><mn>1</mn><mo>)</mo>");
        Assert.Equal(0, Groups(omml));
    }

    [Fact]
    public void AnOpeningBracketWithoutAClosingOneStaysLiteral()
    {
        var omml = Omml("<mo>(</mo><mi>a</mi><mo>+</mo><mi>b</mi>");
        Assert.Equal(0, Groups(omml));
    }

    [Fact]
    public void EmptyBracketsStayLiteral()
    {
        var omml = Omml("<mi>f</mi><mo>(</mo><mo>)</mo>");
        Assert.Equal(0, Groups(omml));
    }

    [Fact]
    public void RelationClassBracketsAreNotTouched()
    {
        var source = $"<math xmlns='{Ns}'><mrow data-mjx-texclass='REL'><mo>(</mo></mrow>"
            + "<mi>x</mi><mrow data-mjx-texclass='REL'><mo>)</mo></mrow></math>";
        Assert.Equal(source, WordOmmlConverter.NormalizePlainBracketPairs(source));
    }

    [Fact]
    public void MathJaxBraceAndFloorBracketsAreGroupedDespiteTheirFenceFlag()
    {
        // MathJax writes fence="false" for \{, \lfloor and \langle.
        var omml = Omml(
            "<mo fence='false' stretchy='false'>{</mo><mi>x</mi><mo fence='false' stretchy='false'>}</mo>"
            + "<mo fence='false' stretchy='false'>&#x230A;</mo><mi>y</mi><mo fence='false' stretchy='false'>&#x230B;</mo>");
        Assert.Equal(2, Groups(omml));
    }

    [Fact]
    public void MathTypeSourcesKeepTheirLiteralBrackets()
    {
        var source = $"<math xmlns='{Ns}'><mo fence='false' stretchy='false'>(</mo>"
            + "<mi data-mtef-mtcode='120'>x</mi><mo fence='false' stretchy='false'>)</mo></math>";
        Assert.Equal(source, WordOmmlConverter.NormalizePlainBracketPairs(source));
    }

    [Fact]
    public void WordLiteralBracketsKeepTheirFixedSize()
    {
        var source = $"<math xmlns='{Ns}'><mo stretchy='false' mathvariant='normal'>(</mo><mi>x</mi>"
            + "<mo stretchy='false' mathvariant='normal'>)</mo></math>";
        Assert.Equal(source, WordOmmlConverter.NormalizePlainBracketPairs(source));
    }

    [Fact]
    public void VerticalBarsAreNotPaired()
    {
        var source = $"<math xmlns='{Ns}'><mo>|</mo><mi>x</mi><mo>|</mo></math>";
        Assert.Equal(source, WordOmmlConverter.NormalizePlainBracketPairs(source));
    }
}
