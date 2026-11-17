using System.Linq;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// LaTeX -> MathML -> OMML of integrals and sums, as the shared export (apps/shared/visualtex-app/src/export/operandMathMl.ts)
// writes the MathML: the operand of a big operator is wrapped in an mrow (the rule of tex-layout/src/operand-scope.ts: an
// integral's up to a sign after its differential, a sum's up to a top-level sign), and the writer makes that row the body of
// its native n-ary object. Before, MathJax's loose items made the single element after the operator the body (the f of
// ∫ f(x) dx) and left the rest of the integrand outside the integral.
public sealed class NaryGroupedOperandTests
{
    private const string P = "http://www.w3.org/1998/Math/MathML";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private static XElement Omml(string body) =>
        XElement.Parse(WordOmmlConverter.TransformMathMlToOmml($"<math xmlns=\"{P}\" display=\"block\">{body}</math>"));

    private static string Text(XElement? element) =>
        element is null ? string.Empty : string.Concat(element.Descendants(M + "t").Select(t => t.Value)).Replace(" ", string.Empty);

    [Fact]
    public void IntegrandThroughItsDifferentialIsTheBodyAndThePlusStaysOutside()
    {
        // \int_a^b f(x)\,dx + c
        var omml = Omml(
            "<msubsup><mo>∫</mo><mi>a</mi><mi>b</mi></msubsup>"
            + "<mrow><mi>f</mi><mo stretchy=\"false\">(</mo><mi>x</mi><mo stretchy=\"false\">)</mo>"
            + "<mstyle scriptlevel=\"0\"><mspace width=\"0.167em\"></mspace></mstyle><mi>d</mi><mi>x</mi></mrow>"
            + "<mo>+</mo><mi>c</mi>");
        var nary = Assert.Single(omml.Descendants(M + "nary"));
        var body = nary.Element(M + "e");
        // (the parentheses are the attributes of Word's delimiter object, not text: f, (x), d, x)
        Assert.Equal("fxdx", Text(body));
        Assert.Single(body!.Elements(M + "d"));
        Assert.Equal("+c", string.Concat(nary.ElementsAfterSelf().Select(Text)));
    }

    [Fact]
    public void SumOperandIsTheBodyAndTheTopLevelPlusStaysOutside()
    {
        // \sum_{i=1}^n a_i + b
        var omml = Omml(
            "<munderover><mo>∑</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></munderover>"
            + "<mrow><msub><mi>a</mi><mi>i</mi></msub></mrow><mo>+</mo><mi>b</mi>");
        var nary = Assert.Single(omml.Descendants(M + "nary"));
        Assert.Equal("ai", Text(nary.Element(M + "e")));
        Assert.Equal("+b", string.Concat(nary.ElementsAfterSelf().Select(Text)));
    }

    [Fact]
    public void AnInnerOperatorAndItsOperandAreTheOuterBody()
    {
        // \sum_i \sum_j a_{ij} + b
        var omml = Omml(
            "<munder><mo>∑</mo><mi>i</mi></munder>"
            + "<mrow><munder><mo>∑</mo><mi>j</mi></munder><mrow><msub><mi>a</mi><mrow><mi>i</mi><mi>j</mi></mrow></msub></mrow></mrow>"
            + "<mo>+</mo><mi>b</mi>");
        var outer = omml.Descendants(M + "nary").First();
        var inner = Assert.Single(outer.Element(M + "e")!.Descendants(M + "nary"));
        Assert.Equal("aij", Text(inner.Element(M + "e")));
        Assert.Equal("+b", string.Concat(outer.ElementsAfterSelf().Select(Text)));
    }
}
