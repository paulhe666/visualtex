using System.Linq;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// MathType -> OMML of integrals and sums: the whole operand (with its dx) is the n-ary body.
// Before, the MTEF reader left the operand as loose tokens after the operator, so Word's
// MathML import (MML2OMML) wrote an empty body (a dotted box) and the integrand after it.
public sealed class MathTypeBigOperatorOperandTests
{
    private const string P = "http://www.w3.org/1998/Math/MathML";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private static XElement OmmlOfMathTypeReadBack(string mathMl)
    {
        var native = MathTypeMtefCodec.CreateEquationNative(mathMl, false);
        var read = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        return XElement.Parse(WordOmmlConverter.TransformMathMlToOmml(read));
    }

    private static string Text(XElement? element) =>
        element is null ? string.Empty : string.Concat(element.Descendants(M + "t").Select(t => t.Value));

    [Fact]
    public void IntegralOperandAndDifferentialAreTheNaryBody()
    {
        var omml = OmmlOfMathTypeReadBack(
            $"<math xmlns=\"{P}\"><mrow><msubsup><mo>∫</mo><mn>0</mn><mfrac><mi>π</mi><mn>2</mn></mfrac></msubsup>"
            + "<mrow><mo>(</mo><msup><mi>e</mi><mi>x</mi></msup><mo>+</mo><mn>1</mn><mo>)</mo><mi>d</mi><mi>x</mi></mrow></mrow>"
            + "<mo>≤</mo><mi>a</mi></math>");
        var nary = Assert.Single(omml.Descendants(M + "nary"));
        var body = nary.Element(M + "e");
        Assert.EndsWith("dx", Text(body));
        Assert.Contains("e", Text(body));
        Assert.Equal("≤a", string.Concat(nary.ElementsAfterSelf().Select(Text)));
    }

    [Fact]
    public void SumOperandIsTheNaryBody()
    {
        var omml = OmmlOfMathTypeReadBack(
            $"<math xmlns=\"{P}\"><mrow><munderover><mo>∑</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></munderover>"
            + "<mrow><msub><mi>a</mi><mi>i</mi></msub><msub><mi>b</mi><mi>i</mi></msub></mrow></mrow></math>");
        var nary = Assert.Single(omml.Descendants(M + "nary"));
        Assert.Equal("aibi", Text(nary.Element(M + "e")));
        Assert.Empty(nary.ElementsAfterSelf());
    }

    [Fact]
    public void EmptyOperandSlotTakesWhatFollowsUpToTheRelation()
    {
        // What MathType writes when the integrand was typed after the template.
        var omml = XElement.Parse(WordOmmlConverter.TransformMathMlToOmml(
            $"<math xmlns=\"{P}\"><mrow><msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup></mrow>"
            + "<mo>(</mo><mi>x</mi><mo>+</mo><mn>1</mn><mo>)</mo><mi>d</mi><mi>x</mi><mo>≤</mo><mi>a</mi></math>"));
        var nary = Assert.Single(omml.Descendants(M + "nary"));
        Assert.Equal("x+1dx", Text(nary.Element(M + "e")));
        Assert.Single(nary.Element(M + "e")!.Elements(M + "d"));
        Assert.Equal("≤a", string.Concat(nary.ElementsAfterSelf().Select(Text)));
    }

    [Fact]
    public void EmptySumSlotStopsAtTheNextTerm()
    {
        var omml = XElement.Parse(WordOmmlConverter.TransformMathMlToOmml(
            $"<math xmlns=\"{P}\"><mrow><munder><mo>∑</mo><mi>i</mi></munder></mrow>"
            + "<msub><mi>a</mi><mi>i</mi></msub><mi>x</mi><mo>+</mo><mi>c</mi></math>"));
        var nary = Assert.Single(omml.Descendants(M + "nary"));
        Assert.Equal("aix", Text(nary.Element(M + "e")));
        Assert.Equal("+c", string.Concat(nary.ElementsAfterSelf().Select(Text)));
    }

    [Fact]
    public void AccentOverABigOperatorKeepsItsMark()
    {
        var omml = XElement.Parse(WordOmmlConverter.TransformMathMlToOmml(
            $"<math xmlns=\"{P}\"><mover><mrow><mo>∮</mo></mrow><mo>~</mo></mover><mi>x</mi></math>"));
        Assert.DoesNotContain("~", Text(omml.Descendants(M + "nary").FirstOrDefault()?.Element(M + "e")));
    }
}
