using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

// MathML of editor commands (as MathJax writes it) that used to fail the
// MathType writer's round-trip check or crash MathType's renderer.
public sealed class MathTypeEditorCommandCoverageTests
{
    private static string RoundTrip(string mathMl) =>
        MathTypeMtefCodec.ReadEquationNativeMathMl(
            MathTypeMtefCodec.CreateEquationNativeAtFontSize(mathMl, false, 12).EquationNative);

    private static void AssertRoundTrip(string mathMl) =>
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(mathMl), MathTypeMtefCodec.SemanticSignature(RoundTrip(mathMl)));

    [Theory]
    // \check: the caron is a modifier letter, still an accent
    [InlineData("<math><mrow><mover><mi>x</mi><mo stretchy=\"false\">&#x2C7;</mo></mover></mrow></math>")]
    // \overparen: MathType stores the arc as U+2322
    [InlineData("<math><mover><mi>x</mi><mo>&#x23DC;</mo></mover></math>")]
    // \liminf / \limsup: one function run per word
    [InlineData("<math><mo data-mjx-texclass=\"OP\" movablelimits=\"true\">lim&#x2006;inf</mo></math>")]
    [InlineData("<math><munder><mo data-mjx-texclass=\"OP\" movablelimits=\"true\">lim&#x2006;sup</mo><mrow><mi>n</mi></mrow></munder><msub><mi>x</mi><mi>n</mi></msub></math>")]
    // \sum^{n}: an upper limit only
    [InlineData("<math><mover><mo data-mjx-texclass=\"OP\">&#x2211;</mo><mi>n</mi></mover><msub><mi>a</mi><mi>i</mi></msub></math>")]
    // clockwise contour integrals, bare and with limits
    [InlineData("<math><mo data-mjx-texclass=\"OP\">&#x2232;</mo></math>")]
    [InlineData("<math><msubsup><mo data-mjx-texclass=\"OP\">&#x2233;</mo><mi>a</mi><mi>b</mi></msubsup><mi>f</mi></math>")]
    // one-sided floor and ceiling (their one-sided templates crash MathType's renderer)
    [InlineData("<math><mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">&#x230A;</mo><mfrac><mi>a</mi><mi>b</mi></mfrac><mo data-mjx-texclass=\"CLOSE\"></mo></mrow></math>")]
    [InlineData("<math><mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\"></mo><mfrac><mi>a</mi><mi>b</mi></mfrac><mo data-mjx-texclass=\"CLOSE\">&#x2309;</mo></mrow></math>")]
    // phantoms are room, not content
    [InlineData("<math><mphantom><mi>x</mi></mphantom><mo>+</mo><mi>y</mi></math>")]
    [InlineData("<math><mpadded width=\"0\"><mphantom><mi>x</mi></mphantom></mpadded><mi>y</mi></math>")]
    public void EditorCommandRoundTrips(string mathMl) => AssertRoundTrip(mathMl);

    [Fact]
    public void PhantomContentIsNeverWrittenAsVisibleCharacters()
    {
        var back = XDocument.Parse(RoundTrip("<math><mphantom><mi>q</mi><mi>q</mi></mphantom><mi>y</mi></math>"));
        Assert.DoesNotContain(back.Descendants(), element => element.Value == "q" && !element.HasElements);
        Assert.Contains(back.Descendants(), element => element.Name.LocalName == "mspace");
    }

    [Fact]
    public void ZeroWidthPhantomWritesNothing()
    {
        var back = XDocument.Parse(RoundTrip("<math><mpadded width=\"0\"><mphantom><mi>q</mi></mphantom></mpadded><mi>y</mi></math>"));
        Assert.DoesNotContain(back.Descendants(), element => element.Name.LocalName is "mspace" || element.Value == "q");
    }

    [Fact]
    public void MultiWordFunctionNameIsNotText()
    {
        var back = XDocument.Parse(RoundTrip("<math><mo movablelimits=\"true\">lim&#x2006;inf</mo></math>"));
        Assert.DoesNotContain(back.Descendants(), element => element.Name.LocalName == "mtext");
    }
}
