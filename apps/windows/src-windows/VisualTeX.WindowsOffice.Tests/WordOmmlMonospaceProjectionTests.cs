using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;
using Xunit.Abstractions;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlMonospaceProjectionTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace P = "http://www.w3.org/1998/Math/MathML";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private readonly ITestOutputHelper _output;

    public WordOmmlMonospaceProjectionTests(ITestOutputHelper output) => _output = output;

    [Theory]
    [InlineData("mi", "x")]
    [InlineData("mi", "AZ")]
    [InlineData("mi", "az")]
    [InlineData("mn", "09")]
    public void ExplicitMonospaceEmitsMathematicalScriptAndRoundTrips(string role, string text)
    {
        var source = Math("<" + role + " mathvariant=\"monospace\">" + text + "</" + role + ">");
        var native = Forward(source);
        _output.WriteLine("Native: " + native);
        var run = Assert.Single(native.Descendants(M + "r"));
        AssertMonospaceRun(run);
        var actual = Reverse(native.ToString(SaveOptions.DisableFormatting));
        _output.WriteLine("Reverse: " + actual);
        AssertStyledTokens(actual.Root!, role, text, "monospace");
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(actual.ToString()));
        Assert.Contains("\\mathtt{", MathMlToLatexConverter.Convert(actual.ToString()), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("x", "mi")]
    [InlineData("AZ", "mi")]
    [InlineData("az", "mi")]
    [InlineData("09", "mn")]
    public void NativeMonospaceScriptRecoversItsAlphabet(string text, string role)
    {
        var actual = Reverse(Native(Run(text, "monospace", "p")));
        _output.WriteLine(actual.ToString(SaveOptions.DisableFormatting));
        AssertStyledTokens(actual.Root!, role, text, "monospace");
        Assert.Contains("\\mathtt{", MathMlToLatexConverter.Convert(actual.ToString()), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("p", "normal")]
    [InlineData("i", "italic")]
    [InlineData("b", "bold")]
    [InlineData("bi", "bold-italic")]
    public void OrdinaryNativeDigitsKeepTheirExistingStyle(string style, string variant)
    {
        var actual = Reverse(Native(Run("09", null, style)));
        AssertStyledTokens(actual.Root!, "mn", "09", variant);
    }

    [Theory]
    [InlineData("double-struck", "p", "double-struck")]
    [InlineData("sans-serif", "p", "sans-serif")]
    [InlineData("sans-serif", "b", "bold-sans-serif")]
    public void OtherRepresentableNativeDigitAlphabetsStayDistinct(string script, string style, string variant)
    {
        var actual = Reverse(Native(Run("09", script, style)));
        _output.WriteLine(actual.ToString(SaveOptions.DisableFormatting));
        AssertStyledTokens(actual.Root!, "mn", "09", variant);
    }

    [Fact]
    public void ExplicitRomanDigitStyleKeepsOrdinaryDefault()
    {
        var actual = Reverse(Native(Run("09", "roman", "p")));
        var number = Assert.Single(actual.Descendants(P + "mn"));
        Assert.Null(number.Attribute("mathvariant"));
        Assert.Equal("09", number.Value);
    }

    [Fact]
    public void InheritedVariantAndLocalOverridesKeepSeparateRuns()
    {
        var native = Forward(Math("<mstyle mathvariant=\"monospace\"><mi>A</mi><mn>0</mn>"
            + "<mi mathvariant=\"normal\">b</mi><mi mathvariant=\"italic\">c</mi></mstyle>"));
        var runs = native.Descendants(M + "r").ToArray();
        Assert.Equal(4, runs.Length);
        AssertMonospaceRun(runs[0]);
        AssertMonospaceRun(runs[1]);
        Assert.Null(runs[2].Element(M + "rPr")?.Element(M + "scr"));
        Assert.Equal("p", (string?)runs[2].Element(M + "rPr")?.Element(M + "sty")?.Attribute(M + "val"));
        Assert.Null(runs[3].Element(M + "rPr")?.Element(M + "scr"));
        Assert.Equal("i", (string?)runs[3].Element(M + "rPr")?.Element(M + "sty")?.Attribute(M + "val"));
    }

    [Fact]
    public void MixedIdentifierNumberAndProseRolesStayDistinct()
    {
        var native = Forward(Math("<mi>x</mi><mi mathvariant=\"monospace\">x</mi>"
            + "<mn mathvariant=\"monospace\">09</mn><mtext>units</mtext>"));
        var runs = native.Descendants(M + "r").ToArray();
        Assert.Equal(4, runs.Length);
        Assert.Null(runs[0].Element(M + "rPr")?.Element(M + "scr"));
        AssertMonospaceRun(runs[1]);
        AssertMonospaceRun(runs[2]);
        Assert.NotNull(runs[3].Element(M + "rPr")?.Element(M + "nor"));
        var actual = Reverse(native.ToString());
        var tokens = actual.Descendants().Where(IsToken).ToArray();
        Assert.Equal(new[] { "mi", "mi", "mn", "mtext" }, tokens.Select(e => e.Name.LocalName));
        Assert.Equal(new[] { "x", "x", "09", "units" }, tokens.Select(e => e.Value));
    }

    [Fact]
    public void MonospaceScriptBaseAndExponentKeepSlotOwnership()
    {
        var source = Math("<msubsup><mi mathvariant=\"monospace\">A</mi><mi>i</mi>"
            + "<mn mathvariant=\"monospace\">09</mn></msubsup>");
        var native = Forward(source);
        var actual = Reverse(native.ToString());
        var slots = Assert.Single(actual.Descendants(P + "msubsup")).Elements().ToArray();
        Assert.Equal(3, slots.Length);
        AssertStyledTokens(slots[0], "mi", "A", "monospace");
        AssertStyledTokens(slots[1], "mi", "i", "italic");
        AssertStyledTokens(slots[2], "mn", "09", "monospace");
    }

    [Theory]
    [InlineData("α")]
    [InlineData("+")]
    public void OutsideAsciiAlphabetCoverageIsNotClaimedFromVariantAlone(string text)
    {
        var projected = XDocument.Parse(WordOmmlConverter.NormalizeMathMlTokenStyleAuthority(
            Math("<mi mathvariant=\"monospace\">" + text + "</mi>")));
        var token = Assert.Single(projected.Descendants(P + "mi"));
        Assert.Equal("monospace", (string?)token.Attribute("mathvariant"));
        Assert.Equal(text, token.Value);
        _output.WriteLine("Source variant retained; native monospace glyph coverage is unsupported/unverified for " + text
            + ". This test does not assert successful native alphabet rendering or native round-trip fidelity.");
    }

    private static void AssertMonospaceRun(XElement run)
    {
        Assert.Equal("monospace", (string?)run.Element(M + "rPr")?.Element(M + "scr")?.Attribute(M + "val"));
        Assert.Equal("p", (string?)run.Element(M + "rPr")?.Element(M + "sty")?.Attribute(M + "val"));
        Assert.Null(run.Element(M + "rPr")?.Element(M + "nor"));
        Assert.Empty(run.Descendants(W + "rFonts"));
    }

    private static void AssertStyledTokens(XElement root, string role, string text, string variant)
    {
        var tokens = root.DescendantsAndSelf().Where(IsToken).ToArray();
        Assert.NotEmpty(tokens);
        Assert.All(tokens, token =>
        {
            Assert.Equal(P + role, token.Name);
            Assert.Equal(variant, MathMlTokenStyleContract.ResolveEffectiveVariant(token));
        });
        Assert.Equal(text, string.Concat(tokens.Select(e => e.Value)));
    }

    private static bool IsToken(XElement e) => e.Name.Namespace == P
        && e.Name.LocalName is "mi" or "mo" or "mn" or "mtext";

    private static string Math(string body) => "<math xmlns=\"" + P + "\">" + body + "</math>";
    private static string Native(string body) => "<m:oMath xmlns:m=\"" + M + "\">" + body + "</m:oMath>";
    private static string Run(string text, string? script, string style) =>
        "<m:r><m:rPr>" + (script is null ? "" : "<m:scr m:val=\"" + script + "\"/>")
        + "<m:sty m:val=\"" + style + "\"/></m:rPr><m:t>" + text + "</m:t></m:r>";
    private static XDocument Forward(string math) => XDocument.Parse(WordOmmlConverter.TransformMathMlToOmml(math));
    private static XDocument Reverse(string native) => XDocument.Parse(WordOmmlConverter.TransformOmmlToMathMl(native, false));
}
