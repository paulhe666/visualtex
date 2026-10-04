using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlExplicitSpaceTests
{
    private static readonly XNamespace P = "http://www.w3.org/1998/Math/MathML";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private static string Source(string body) => $"<math xmlns=\"{P}\">{body}</math>";
    private static XDocument Forward(string body) =>
        XDocument.Parse(WordOmmlConverter.TransformMathMlToOmml(Source(body)));
    private static XDocument Reverse(XDocument omml) =>
        XDocument.Parse(WordOmmlConverter.TransformOmmlToMathMl(
            omml.ToString(SaveOptions.DisableFormatting), false));
    private static string NativeText(XContainer omml) =>
        string.Concat(omml.Descendants(M + "t").Select(text => text.Value));
    private static XDocument Native(string text, bool prose = false) =>
        new(new XElement(M + "oMath", new XElement(M + "r",
            prose ? new XElement(M + "rPr", new XElement(M + "nor")) : null,
            new XElement(M + "t", text))));

    [Theory]
    [InlineData("1em", "1em", "\u2003")]
    [InlineData("+001.000em", "1em", "\u2003")]
    [InlineData(" 01.000 em ", "1em", "\u2003")]
    [InlineData("2em", "2em", "\u2003\u2003")]
    [InlineData("+002.000em", "2em", "\u2003\u2003")]
    public void ExactSourceSpellingsUseTheSameBidirectionalContract(
        string width, string canonical, string nativeText)
    {
        var capability = MathSpaceContract.DescribeOmmlSpace(width);
        Assert.True(capability.CanProject);
        Assert.Empty(capability.AuditLimits);
        Assert.Equal(nativeText, capability.Text);
        Assert.Equal(canonical, MathSpaceContract.FindOmmlSpaceText(nativeText)!.MathMlWidth);
        var forward = Forward($"<mi>a</mi><mspace width=\"{width}\"/><mi>b</mi>");
        Assert.Equal("a" + nativeText + "b", NativeText(forward));
        Assert.DoesNotContain(forward.Descendants().Attributes(),
            a => a.Name.LocalName.StartsWith("data-visualtex", StringComparison.Ordinal));
        var reverse = Reverse(forward);
        var space = Assert.Single(reverse.Descendants(P + "mspace"));
        Assert.Equal(canonical, (string?)space.Attribute("width"));
        Assert.Equal(new[] { "mi", "mspace", "mi" },
            reverse.Root!.Elements().Select(element => element.Name.LocalName));
        Assert.Equal(new[] { "a", "b" }, reverse.Descendants(P + "mi").Select(token => token.Value));
    }

    [Theory]
    [InlineData("1.00000000000000000000000000001em")]
    [InlineData("0.99999999999999999999999999999em")]
    [InlineData("2.00000000000000000000000000001em")]
    [InlineData("1.99999999999999999999999999999em")]
    [InlineData("0.167em")]
    [InlineData("1.25em")]
    [InlineData("12pt")]
    [InlineData("-1em")]
    [InlineData("thinmathspace")]
    [InlineData("12%")]
    public void UnmappedWidthsRemainExplicitLimitsAndNeverRound(string width)
    {
        var source = Source($"<mi>a</mi><mspace width=\"{width}\"/><mi>b</mi>");
        var projection = WordOmmlConverter.ProjectMathMlSpacesForOmml(source);
        var capability = Assert.Single(projection.Capabilities);
        Assert.False(capability.CanProject);
        Assert.False(capability.WidthSupported);
        Assert.Null(capability.Text);
        Assert.Contains(capability.AuditLimits, limit => limit.Contains("width"));
        Assert.Equal(width, XDocument.Parse(projection.MathMl)
            .Descendants(P + "mspace").Single().Attribute("width")!.Value);
        // Existing stock omission remains an audit limit, not an exact mapping.
        Assert.Equal("ab", NativeText(Forward($"<mi>a</mi><mspace width=\"{width}\"/><mi>b</mi>")));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("0em")]
    [InlineData("-0.00pt")]
    [InlineData("+000.0px")]
    public void ZeroWidthEmitsNoCharacterAndNoEmptyRun(string? width)
    {
        var capability = MathSpaceContract.DescribeOmmlSpace(width);
        Assert.True(capability.CanProject);
        Assert.Equal("", capability.Text);
        var attribute = width is null ? "" : $" width=\"{width}\"";
        var forward = Forward($"<mi>a</mi><mspace{attribute}/><mi>b</mi>");
        Assert.Equal("ab", NativeText(forward));
        Assert.All(forward.Descendants(M + "r"), run => Assert.NotEmpty(run.Value));
        Assert.Empty(Reverse(forward).Descendants(P + "mspace"));
    }

    [Theory]
    [InlineData("height", "1em")]
    [InlineData("depth", "2pt")]
    [InlineData("linebreak", "newline")]
    [InlineData("linebreak", "nobreak")]
    public void NonWidthDimensionsHaveIndependentAuditLimits(string name, string value)
    {
        var projection = WordOmmlConverter.ProjectMathMlSpacesForOmml(
            Source($"<mspace width=\"1em\" {name}=\"{value}\"/>"));
        var capability = Assert.Single(projection.Capabilities);
        Assert.True(capability.WidthSupported);
        Assert.False(capability.CanProject);
        Assert.Equal(name != "height", capability.HeightSupported);
        Assert.Equal(name != "depth", capability.DepthSupported);
        Assert.Equal(name != "linebreak", capability.LinebreakSupported);
        Assert.Contains(capability.AuditLimits, limit => limit.Contains(name));
        Assert.Equal("", NativeText(Forward($"<mspace width=\"1em\" {name}=\"{value}\"/>")));
    }

    [Fact]
    public void ExplicitZeroExtentsAndDefaultLinebreakPermitWidthProjection()
    {
        var projection = WordOmmlConverter.ProjectMathMlSpacesForOmml(
            Source("<mspace width=\"1em\" height=\"-0.00pt\" depth=\"0px\" linebreak=\"auto\"/>"));
        Assert.True(Assert.Single(projection.Capabilities).CanProject);
        Assert.Equal("\u2003", NativeText(Forward(
            "<mspace width=\"1em\" height=\"-0.00pt\" depth=\"0px\" linebreak=\"auto\"/>")));
    }

    [Theory]
    [InlineData("\u2003", "1em")]
    [InlineData("\u2003\u2003", "2em")]
    public void CoalescedNativeMathTextDecodesAtItsActualSourcePosition(string gap, string width)
    {
        var reverse = Reverse(Native("a" + gap + "b"));
        Assert.Equal(new[] { "mi", "mspace", "mi" },
            reverse.Root!.Elements().Select(element => element.Name.LocalName));
        Assert.Equal(width, Assert.Single(reverse.Descendants(P + "mspace")).Attribute("width")!.Value);
    }

    [Theory]
    [InlineData("\u2003")]
    [InlineData("\u2003\u2003")]
    [InlineData("a\u2003b")]
    [InlineData("word\u2003words")]
    public void OrdinaryTextSpacesAreNeverPromotedToMathSpacing(string text)
    {
        var reverse = Reverse(Native(text, prose: true));
        Assert.Empty(reverse.Descendants(P + "mspace"));
        Assert.Equal(text, Assert.Single(reverse.Descendants(P + "mtext")).Value);
        var roundtrip = Reverse(Forward("<mtext>" + text + "</mtext>"));
        Assert.Empty(roundtrip.Descendants(P + "mspace"));
        Assert.Equal(text, Assert.Single(roundtrip.Descendants(P + "mtext")).Value);
    }

    [Theory]
    [InlineData("\u00A0")]
    [InlineData("\u2002")]
    [InlineData("\u2009")]
    [InlineData("\u205F")]
    public void UncharacterizedNativeSpacesAreNotReinterpreted(string text)
    {
        Assert.Null(MathSpaceContract.FindOmmlSpaceText(text));
        var reverse = Reverse(Native("a" + text + "b"));
        Assert.Empty(reverse.Descendants(P + "mspace"));
        Assert.Equal("a" + text + "b", reverse.Root!.Value);
    }

    [Fact]
    public void ExplicitGapsPreserveSurroundingStylesAndScriptStructure()
    {
        var forward = Forward("<mi mathvariant=\"normal\">a</mi><mspace width=\"1em\"/>"
            + "<msub><mi mathvariant=\"bold\">x</mi><mrow><mi mathvariant=\"italic\">b</mi>"
            + "<mspace width=\"2em\"/><mi mathvariant=\"normal\">c</mi></mrow></msub>");
        var script = Assert.Single(forward.Descendants(M + "sSub"));
        Assert.Equal("b\u2003\u2003c", NativeText(script.Element(M + "sub")!));
        var runs = forward.Descendants(M + "r")
            .Where(run => run.Element(M + "t")!.Value != "\u2003"
                && run.Element(M + "t")!.Value != "\u2003\u2003").ToArray();
        Assert.Equal(new[] { "a", "x", "b", "c" }, runs.Select(run => run.Element(M + "t")!.Value));
        Assert.Equal(new[] { "p", "b", "i", "p" },
            runs.Select(run => (string?)run.Element(M + "rPr")?.Element(M + "sty")?.Attribute(M + "val")));
        var reverse = Reverse(forward);
        Assert.Equal(new[] { "1em", "2em" },
            reverse.Descendants(P + "mspace").Select(space => space.Attribute("width")!.Value));
        var reverseScript = Assert.Single(reverse.Descendants(P + "msub"));
        Assert.Equal("2em", Assert.Single(reverseScript.Elements().Last()
            .DescendantsAndSelf(P + "mspace")).Attribute("width")!.Value);
    }

    [Fact]
    public void NamespaceAliasesAndUntrustedProjectionAttributesCannotBypassTheContract()
    {
        var source = $"<p:math xmlns:p=\"{P}\"><p:mi>a</p:mi>"
            + "<p:mspace width=\"1em\" data-visualtex-omml-space=\"wrong\"/><p:mi>b</p:mi>"
            + "<p:mspace width=\"1.1em\" data-visualtex-omml-space=\"wrong\"/><p:mi>c</p:mi></p:math>";
        var projection = WordOmmlConverter.ProjectMathMlSpacesForOmml(source);
        Assert.True(projection.Capabilities[0].CanProject);
        Assert.False(projection.Capabilities[1].CanProject);
        var forward = XDocument.Parse(WordOmmlConverter.TransformMathMlToOmml(source));
        Assert.Equal("a\u2003bc", NativeText(forward));
        Assert.DoesNotContain("wrong", forward.ToString());
        Assert.DoesNotContain(forward.Descendants().Attributes(),
            a => a.Name.LocalName.StartsWith("data-visualtex", StringComparison.Ordinal));
    }

    [Fact]
    public void AdjacentSourceGapsAndStructureBoundariesKeepTheirPositions()
    {
        var forward = Forward("<mi>a</mi><mspace width=\"1em\"/><mspace width=\"1em\"/>"
            + "<mfrac><mi>x</mi><mi>y</mi></mfrac><mspace width=\"1em\"/><mi>b</mi>");
        var reverse = Reverse(forward);
        Assert.Equal(new[] { "mi", "mspace", "mfrac", "mspace", "mi" },
            reverse.Root!.Elements().Select(element => element.Name.LocalName));
        Assert.Equal(new[] { "2em", "1em" },
            reverse.Descendants(P + "mspace").Select(space => space.Attribute("width")!.Value));
    }
}
