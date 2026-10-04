using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;
using VisualTeX.WordVsto;
using Xunit;
using Xunit.Abstractions;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlUprightStyleAuthorityTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace P = "http://www.w3.org/1998/Math/MathML";
    private readonly ITestOutputHelper output;
    public WordOmmlUprightStyleAuthorityTests(ITestOutputHelper output) => this.output = output;

    private static string Source(string body) => $"<math xmlns=\"{P}\">{body}</math>";
    private static XElement[] Runs(string omml) => XDocument.Parse(omml).Descendants(M + "r").ToArray();
    private static string Text(XElement run) => string.Concat(run.Elements(M + "t").Select(t => t.Value));
    private static string? Style(XElement run) => (string?)run.Element(M + "rPr")?.Element(M + "sty")?.Attribute(M + "val");
    private static bool Upright(XElement run) => Style(run) == "p" || run.Element(M + "rPr")?.Element(M + "nor") is not null;

    // Keep the installed stylesheet as an independent baseline. The product's
    // GetTransform now applies its own in-memory structural style correction.
    // This does not modify Office or create a Word host.
    private static string StockForward(string mathMl)
    {
        var transform = new XslCompiledTransform();
        transform.Load(WordOmmlConverter.ResolveTransformPath(),
            new XsltSettings(enableDocumentFunction: false, enableScript: false), null);
        using var input = XmlReader.Create(new StringReader(mathMl), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, IgnoreWhitespace = true });
        using var result = new StringWriter();
        var settings = transform.OutputSettings!.Clone();
        settings.OmitXmlDeclaration = true;
        using (var writer = XmlWriter.Create(result, settings)) transform.Transform(input, writer);
        return WordOmmlConverter.ExtractSingleOMath(result.ToString());
    }

    [Theory]
    [InlineData("<mi>sin</mi>", "sin")]
    [InlineData("<mi>lim</mi>", "lim")]
    [InlineData("<mo data-mjx-texclass=\"OP\">sin</mo>", "sin")]
    [InlineData("<mo data-mjx-texclass=\"OP\">lim</mo>", "lim")]
    [InlineData("<mi mathvariant=\"normal\">ABC</mi>", "ABC")]
    [InlineData("<mi mathvariant=\"normal\">d</mi>", "d")]
    [InlineData("<mo mathvariant=\"normal\" data-mjx-texclass=\"OP\">sin</mo>", "sin")]
    [InlineData("<mo mathvariant=\"normal\" data-mjx-texclass=\"OP\">lim</mo>", "lim")]
    [InlineData("<mi mathvariant=\"normal\">sin</mi><mo>&#x2061;</mo><mi>x</mi>", "sin")]
    public void StockDefaultUprightEvidence(string body, string text)
    {
        var raw = StockForward(Source(body));
        output.WriteLine("STOCK " + body + " => " + raw);
        var run = Assert.Single(Runs(raw), r => Text(r) == text);
        Assert.Equal(!body.StartsWith("<mo data-mjx", StringComparison.Ordinal), Upright(run));
        var actual = WordOmmlConverter.TransformMathMlToOmml(Source(body));
        Assert.True(Upright(Assert.Single(Runs(actual), r => Text(r) == text)), actual);
        output.WriteLine("PROJECTED " + body + " => " + actual);
        if (body.Contains("2061", StringComparison.Ordinal))
            Assert.Single(XDocument.Parse(actual).Descendants(M + "func"));
    }

    [Theory]
    [InlineData("italic", "")]
    [InlineData("bold-italic", "bi")]
    public void StockExplicitVariantPrecedesMulticharDefault(string variant, string expectedStyle)
    {
        var raw = StockForward(Source($"<mi mathvariant=\"{variant}\">ABC</mi>"));
        output.WriteLine("STOCK " + variant + " => " + raw);
        var run = Assert.Single(Runs(raw));
        Assert.Equal(expectedStyle, Style(run) ?? "");
        Assert.False(Upright(run));
    }

    [Theory]
    [InlineData("italic", "i")]
    [InlineData("bold-italic", "bi")]
    public void ExplicitVariantSurvivesUprightNormalization(string variant, string expectedStyle)
    {
        var xml = Source($"<mi mathvariant=\"{variant}\">ABC</mi>");
        var raw = StockForward(xml);
        var actual = WordOmmlConverter.TransformMathMlToOmml(xml);
        output.WriteLine("NORMALIZED " + variant + " STOCK=" + raw + " ACTUAL=" + actual);
        var run = Assert.Single(Runs(actual));
        Assert.Equal(expectedStyle, Style(run) ?? "");
        Assert.False(Upright(run));
    }

    [Theory]
    [InlineData("italic", "i")]
    [InlineData("bold-italic", "bi")]
    public void NearestInheritedVariantSurvivesUprightNormalization(string variant, string expectedStyle)
    {
        var xml = Source($"<mstyle mathvariant=\"normal\"><mstyle mathvariant=\"{variant}\"><mi>ABC</mi></mstyle></mstyle>");
        var raw = StockForward(xml);
        var actual = WordOmmlConverter.TransformMathMlToOmml(xml);
        output.WriteLine("INHERITED " + variant + " STOCK=" + raw + " ACTUAL=" + actual);
        var run = Assert.Single(Runs(actual));
        Assert.Equal(expectedStyle, Style(run) ?? "");
        Assert.False(Upright(run));
    }

    [Theory]
    [InlineData("italic", "i")]
    [InlineData("bold-italic", "bi")]
    public void SameWordNormalAndStyledOccurrencesDoNotContaminateEachOther(string variant, string expectedStyle)
    {
        var xml = Source($"<mi mathvariant=\"normal\">ABC</mi><mo>+</mo><mi mathvariant=\"{variant}\">ABC</mi>");
        var raw = StockForward(xml);
        var actual = WordOmmlConverter.TransformMathMlToOmml(xml);
        output.WriteLine("MIXED " + variant + " STOCK=" + raw + " ACTUAL=" + actual);
        Assert.Equal("ABC+ABC", string.Concat(Runs(actual).Select(Text)));
        var runs = Runs(actual).Where(r => Text(r).Contains("ABC", StringComparison.Ordinal)).ToArray();
        Assert.Equal(2, runs.Length);
        Assert.Equal("p", Style(runs[0]));
        Assert.Equal(expectedStyle, Style(runs[1]) ?? "");
        Assert.False(Upright(runs[1]));
    }

    [Fact]
    public void SourceProjectionMaterializesNearestVariantAndTokenDefaults()
    {
        var source = Source("<mstyle mathvariant=\"normal\"><mstyle mathvariant=\"bold-italic\"><mi>ABC</mi><mo>+</mo><mn>2</mn><mi mathvariant=\"normal\">d</mi></mstyle></mstyle><mi>x</mi><mo>+</mo><mn>3</mn><mi>sin</mi>");
        var projected = XDocument.Parse(WordOmmlConverter.NormalizeMathMlTokenStyleAuthority(source));
        var tokens = projected.Descendants().Where(e => e.Name.LocalName is "mi" or "mo" or "mn").ToArray();
        Assert.Equal(new[] { "bold-italic", "bold-italic", "bold-italic", "normal", "italic", "normal", "normal", "normal" }, tokens.Select(t => (string?)t.Attribute("mathvariant")));
    }

    [Fact]
    public void SameWordProseNormalAndItalicRemainSeparate()
    {
        var actual = WordOmmlConverter.TransformMathMlToOmml(Source("<mtext>ABC</mtext><mo>+</mo><mi mathvariant=\"normal\">ABC</mi><mo>+</mo><mi mathvariant=\"italic\">ABC</mi>"));
        Assert.Equal("ABC+ABC+ABC", string.Concat(Runs(actual).Select(Text)));
        var runs = Runs(actual).Where(r => Text(r).Contains("ABC", StringComparison.Ordinal)).ToArray();
        Assert.Equal(3, runs.Length);
        Assert.NotNull(runs[0].Element(M + "rPr")?.Element(M + "nor"));
        Assert.Equal("p", Style(runs[1]));
        Assert.Null(runs[1].Element(M + "rPr")?.Element(M + "nor"));
        Assert.Equal("i", Style(runs[2]));
        Assert.Null(runs[2].Element(M + "rPr")?.Element(M + "nor"));
    }

    [Theory]
    [InlineData("italic", "i")]
    [InlineData("bold-italic", "bi")]
    public void NamedOperatorDefaultDoesNotOverrideSameWordExplicitVariant(string variant, string expectedStyle)
    {
        var actual = WordOmmlConverter.TransformMathMlToOmml(Source($"<mo data-mjx-texclass=\"OP\">lim</mo><mo>+</mo><mo mathvariant=\"{variant}\" data-mjx-texclass=\"OP\">lim</mo>"));
        var runs = Runs(actual).Where(r => Text(r).Contains("lim", StringComparison.Ordinal)).ToArray();
        Assert.Equal(2, runs.Length);
        Assert.Equal("p", Style(runs[0]));
        Assert.Equal(expectedStyle, Style(runs[1]) ?? "");
    }

    [Fact]
    public void ExplicitItalicOperatorStyleRemainsDistinctFromPlainOperatorDefault()
    {
        var xml = Source("<mo mathvariant=\"italic\">∥</mo><mo>+</mo><mo>∥</mo>");
        var projected = XDocument.Parse(WordOmmlConverter.NormalizeMathMlTokenStyleAuthority(xml));
        Assert.Equal(new[] { "italic", "normal", "normal" }, projected.Descendants(P + "mo").Select(t => (string?)t.Attribute("mathvariant")));
        var actual = WordOmmlConverter.TransformMathMlToOmml(xml);
        output.WriteLine("F02 FORWARD PROBE " + actual);
        output.WriteLine("F02 REVERSE PROBE " + WordOmmlConverter.TransformOmmlToMathMl(actual, false));
        Assert.Equal(new[] { "∥", "+∥" }, Runs(actual).Select(Text));
        Assert.Equal(new[] { "i", "p" }, Runs(actual).Select(Style));
        var reverse = XDocument.Parse(WordOmmlConverter.TransformOmmlToMathMl(actual, false));
        Assert.Equal(new[] { "italic", "normal", "normal" },
            reverse.Descendants(P + "mo").Select(token => (string?)token.Attribute("mathvariant")));
    }
}
