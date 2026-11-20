using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;
using Xunit.Abstractions;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlScriptBaseRoleRecoveryTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace P = "http://www.w3.org/1998/Math/MathML";
    private readonly ITestOutputHelper _output;

    public WordOmmlScriptBaseRoleRecoveryTests(ITestOutputHelper output) => _output = output;

    [Theory]
    [InlineData("C01434", false)]
    [InlineData("C01435", true)]
    public void CorpusBareBracketPairWithAPowerBecomesOneNativeGroup(string id, bool display)
    {
        var source = Math("<mo stretchy=\"false\">(</mo><mi>x</mi><mo>+</mo><mi>y</mi>"
            + "<msup><mo stretchy=\"false\">)</mo><mn>2</mn></msup>", display);
        var actual = RoundTrip(source, display);
        ReportCorpus(id, source, actual);
        var script = Assert.Single(actual.Descendants(P + "msup"));
        var slots = script.Elements().ToArray();
        Assert.Equal(2, slots.Length);
        // The base is the whole bracketed group, not a lone closing character.
        Assert.Equal("x+y", slots[0].Value);
        Assert.Equal("2", slots[1].Value);
    }

    [Theory]
    [InlineData("C01470", false)]
    [InlineData("C01471", true)]
    public void CorpusScriptedBarsKeepBothOperatorRoles(string id, bool display)
    {
        var source = Math("<mo data-mjx-texclass=\"OPEN\" fence=\"false\" stretchy=\"false\">|</mo><mi>x</mi>"
            + "<msubsup><mo data-mjx-texclass=\"CLOSE\" fence=\"false\" stretchy=\"false\">|</mo><mi>i</mi><mn>2</mn></msubsup>"
            + "<mo>+</mo><mo data-mjx-texclass=\"OPEN\" fence=\"false\" stretchy=\"false\">&#x2016;</mo><mi>y</mi>"
            + "<msubsup><mo data-mjx-texclass=\"CLOSE\" fence=\"false\" stretchy=\"false\">&#x2016;</mo><mi>j</mi><mn>3</mn></msubsup>", display);
        var actual = RoundTrip(source, display);
        ReportCorpus(id, source, actual);
        var scripts = actual.Descendants(P + "msubsup").ToArray();
        Assert.Equal(2, scripts.Length);
        AssertToken(scripts[0].Elements().First(), "mo", "|");
        AssertToken(scripts[1].Elements().First(), "mo", "‖");
        Assert.Equal(new[] { "|", "i", "2" }, scripts[0].Elements().Select(e => e.Value));
        Assert.Equal(new[] { "‖", "j", "3" }, scripts[1].Elements().Select(e => e.Value));
    }

    [Theory]
    [InlineData(")", "mo")]
    [InlineData("|", "mo")]
    [InlineData("‖", "mo")]
    [InlineData("+", "mo")]
    [InlineData("23", "mn")]
    [InlineData("sin", "mi")]
    [InlineData("abc", "mi")]
    [InlineData("x", "mi")]
    public void NativePlainRunPreservesProjectedTokenRole(string text, string role)
    {
        var actual = Reverse(Sup(Run(text, "<m:sty m:val=\"p\"/>"), Run("2")));
        _output.WriteLine(actual.ToString(SaveOptions.DisableFormatting));
        var script = Assert.Single(actual.Descendants(P + "msup"));
        AssertToken(script.Elements().First(), role, text);
        Assert.Equal("2", script.Elements().Last().Value);
        if (role == "mi")
            Assert.Equal("normal", (string?)script.Elements().First().DescendantsAndSelf(P + "mi").Single().Attribute("mathvariant"));
    }

    [Theory]
    [InlineData("sin")]
    [InlineData("units")]
    [InlineData("a+b")]
    [InlineData("+")]
    [InlineData("‖")]
    public void NativeNormalTextScriptBaseRemainsProse(string text)
    {
        var actual = Reverse(Sup(Run(text, "<m:nor/>"), Run("2")));
        _output.WriteLine(actual.ToString(SaveOptions.DisableFormatting));
        AssertToken(Assert.Single(actual.Descendants(P + "msup")).Elements().First(), "mtext", text);
    }

    [Fact]
    public void RepeatedTextKeepsItsOriginalStyleAndRoleOwnership()
    {
        var actual = Reverse(Sup(Run("x", "<m:sty m:val=\"i\"/>"), Run("2"))
            + Sup(Run("x", "<m:nor/>"), Run("3"))
            + Sup(Run("x", "<m:sty m:val=\"p\"/>"), Run("4")));
        var scripts = actual.Descendants(P + "msup").ToArray();
        Assert.Equal(3, scripts.Length);
        AssertToken(scripts[0].Elements().First(), "mi", "x");
        Assert.Equal("italic", MathMlTokenStyleContract.ResolveEffectiveVariant(
            scripts[0].Elements().First().DescendantsAndSelf(P + "mi").Single()));
        AssertToken(scripts[1].Elements().First(), "mtext", "x");
        AssertToken(scripts[2].Elements().First(), "mi", "x");
        Assert.Equal("normal", MathMlTokenStyleContract.ResolveEffectiveVariant(
            scripts[2].Elements().First().DescendantsAndSelf(P + "mi").Single()));
        Assert.Equal(new[] { "2", "3", "4" }, scripts.Select(s => s.Elements().Last().Value));
    }

    [Fact]
    public void NormalTextWithPlainStyleRemainsProse()
    {
        var actual = Reverse(Sup(Run("+", "<m:nor/><m:sty m:val=\"p\"/>"), Run("2")));
        AssertToken(Assert.Single(actual.Descendants(P + "msup")).Elements().First(), "mtext", "+");
    }

    [Fact]
    public void MixedNativeRunsKeepTheirSeparateRoles()
    {
        var actual = Reverse(Sup(Run("x") + Run("+", "<m:sty m:val=\"p\"/>")
            + Run("units", "<m:nor/>"), Run("2")));
        var script = Assert.Single(actual.Descendants(P + "msup"));
        var tokens = script.Elements().First().DescendantsAndSelf().Where(IsToken).ToArray();
        Assert.Equal(new[] { "mi", "mo", "mtext" }, tokens.Select(t => t.Name.LocalName));
        Assert.Equal(new[] { "x", "+", "units" }, tokens.Select(t => t.Value));
        Assert.Equal("2", script.Elements().Last().Value);
    }

    [Fact]
    public void NestedScriptsKeepTheirOwnBaseSlots()
    {
        var actual = Reverse(Sup(Sup(Run(")", "<m:sty m:val=\"p\"/>"), Run("2")), Run("3")));
        var outer = actual.Descendants(P + "msup").First();
        var inner = Assert.Single(outer.Elements().First().DescendantsAndSelf(P + "msup"));
        AssertToken(inner.Elements().First(), "mo", ")");
        Assert.Equal("2", inner.Elements().Last().Value);
        Assert.Equal("3", outer.Elements().Last().Value);
    }

    [Fact]
    public void StructuredBaseWithOnePlainDescendantRunIsNotFlattened()
    {
        var radical = "<m:rad><m:radPr><m:degHide m:val=\"1\"/></m:radPr><m:deg/><m:e>"
            + Run("23", "<m:sty m:val=\"p\"/>") + "</m:e></m:rad>";
        var actual = Reverse(Sup(radical, Run("2")));
        var script = Assert.Single(actual.Descendants(P + "msup"));
        Assert.Single(script.Elements().First().DescendantsAndSelf(P + "msqrt"));
        Assert.Equal("23", script.Elements().First().Value);
        Assert.Equal("2", script.Elements().Last().Value);
    }

    private void ReportCorpus(string id, string source, XDocument actual)
    {
        _output.WriteLine(id + ": " + actual);
        _output.WriteLine("Source signature: " + MathTypeMtefCodec.SemanticSignature(source));
        _output.WriteLine("Output signature: " + MathTypeMtefCodec.SemanticSignature(actual.ToString()));
    }

    private static bool IsToken(XElement e) => e.Name.Namespace == P
        && e.Name.LocalName is "mi" or "mo" or "mn" or "mtext";

    private static void AssertToken(XElement slot, string role, string text)
    {
        var token = Assert.Single(slot.DescendantsAndSelf(), IsToken);
        Assert.Equal(P + role, token.Name);
        Assert.Equal(text, token.Value);
    }

    private static string Math(string body, bool display) =>
        new XElement(P + "math", new XAttribute("display", display ? "block" : "inline"),
            XElement.Parse("<root xmlns=\"" + P + "\">" + body + "</root>").Nodes())
        .ToString(SaveOptions.DisableFormatting);

    private static XDocument RoundTrip(string math, bool display) =>
        XDocument.Parse(WordOmmlConverter.TransformOmmlToMathMl(
            WordOmmlConverter.TransformMathMlToOmml(math), display));

    private static string Run(string text, string properties = "") =>
        "<m:r>" + (properties.Length == 0 ? "" : "<m:rPr>" + properties + "</m:rPr>")
        + "<m:t>" + new XText(text).ToString() + "</m:t></m:r>";

    private static string Sup(string baseSlot, string exponent) =>
        "<m:sSup><m:e>" + baseSlot + "</m:e><m:sup>" + exponent + "</m:sup></m:sSup>";

    private static XDocument Reverse(string body) =>
        XDocument.Parse(WordOmmlConverter.TransformOmmlToMathMl(
            "<m:oMath xmlns:m=\"" + M + "\">" + body + "</m:oMath>", false));
}
