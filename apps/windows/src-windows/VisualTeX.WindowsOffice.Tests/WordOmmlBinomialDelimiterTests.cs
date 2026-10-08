using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlBinomialDelimiterTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string Ns = "http://www.w3.org/1998/Math/MathML";

    // MathJax output for \binom{n}{k}.
    private static string Binomial(string open = "(", string close = ")") =>
        $"<math xmlns='{Ns}' display='block'><mrow data-mjx-texclass='ORD'>"
        + $"<mrow data-mjx-texclass='OPEN'><mo minsize='2.047em' maxsize='2.047em'>{open}</mo></mrow>"
        + "<mfrac linethickness='0'><mi>n</mi><mi>k</mi></mfrac>"
        + $"<mrow data-mjx-texclass='CLOSE'><mo minsize='2.047em' maxsize='2.047em'>{close}</mo></mrow>"
        + "</mrow></math>";

    [Fact]
    public void MathJaxBinomialBecomesAStretchyDelimiterAroundABarlessFraction()
    {
        var omml = XDocument.Parse(WordOmmlConverter.TransformMathMlToOmml(Binomial()));
        var delimiter = Assert.Single(omml.Descendants(M + "d"));
        var fraction = Assert.Single(delimiter.Descendants(M + "f"));
        Assert.Equal("noBar", (string?)fraction.Element(M + "fPr")?.Element(M + "type")?.Attribute(M + "val"));
        // No loose bracket characters may remain beside the m:d.
        Assert.DoesNotContain(
            omml.Descendants(M + "t"),
            text => text.Value is "(" or ")");
    }

    [Fact]
    public void BinomialWithOtherDelimitersKeepsThem()
    {
        var omml = XDocument.Parse(WordOmmlConverter.TransformMathMlToOmml(Binomial("[", "]")));
        var properties = Assert.Single(omml.Descendants(M + "dPr"));
        Assert.Equal("[", (string?)properties.Element(M + "begChr")?.Attribute(M + "val"));
        Assert.Equal("]", (string?)properties.Element(M + "endChr")?.Attribute(M + "val"));
    }

    [Fact]
    public void RuledFractionBetweenFixedDelimitersIsNotTreatedAsABinomial()
    {
        var source = Binomial().Replace(" linethickness='0'", string.Empty);
        var normalized = WordOmmlConverter.NormalizeMathJaxBinomialDelimiters(source);
        Assert.Equal(source, normalized);
    }

    [Fact]
    public void HandSizedDelimitersAroundAnythingElseAreLeftAlone()
    {
        var source = $"<math xmlns='{Ns}'><mrow data-mjx-texclass='ORD'>"
            + "<mrow data-mjx-texclass='OPEN'><mo minsize='1.2em' maxsize='1.2em'>(</mo></mrow>"
            + "<mi>x</mi>"
            + "<mrow data-mjx-texclass='CLOSE'><mo minsize='1.2em' maxsize='1.2em'>)</mo></mrow>"
            + "</mrow></math>";
        Assert.Equal(source, WordOmmlConverter.NormalizeMathJaxBinomialDelimiters(source));
    }
}
