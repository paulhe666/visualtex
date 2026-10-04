using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlExplicitFenceStructureTests
{
    private static readonly XNamespace P = "http://www.w3.org/1998/Math/MathML";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private static string Source(string body) => $"<math xmlns=\"{P}\">{body}</math>";
    private static string Row(string body, string open = "⟮", string close = "⟯",
        string openAttributes = "", string closeAttributes = "", string rowAttributes = "") =>
        $"<mrow data-mjx-texclass=\"INNER\" {rowAttributes}>"
        + $"<mo data-mjx-texclass=\"OPEN\" {openAttributes}>{open}</mo>"
        + body + $"<mo data-mjx-texclass=\"CLOSE\" {closeAttributes}>{close}</mo></mrow>";
    private static XDocument Normalize(string source) =>
        XDocument.Parse(WordOmmlConverter.NormalizeFencedMathMlTables(source));
    private static XDocument Forward(string source) =>
        XDocument.Parse(WordOmmlConverter.TransformMathMlToOmml(source));
    private static string NativeText(XContainer root) =>
        string.Concat(root.Descendants(M + "t").Select(text => text.Value));
    private static string Delimiter(XElement delimiter, string name, string defaultValue) =>
        (string?)delimiter.Element(M + "dPr")?.Element(M + name)?.Attribute(M + "val")
        ?? defaultValue;
    private const string Table =
        "<mtable columnalign=\"center center\"><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr>"
        + "<mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable>";

    [Fact]
    public void OriginalF09SourcePreservesBothNativeFenceRolesAndEmptyClose()
    {
        // Exact source shape from representative20/manifest.json, F09.
        var source = $"<math xmlns=\"{P}\" display=\"block\">"
            + "<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">&#x27EE;</mo>"
            + "<mi>x</mi><mo data-mjx-texclass=\"CLOSE\">&#x27EF;</mo></mrow><mo>+</mo>"
            + "<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\" fence=\"true\""
            + " stretchy=\"true\" symmetric=\"true\">&#x23D0;</mo><mi>y</mi>"
            + "<mo data-mjx-texclass=\"CLOSE\" fence=\"true\" stretchy=\"true\" symmetric=\"true\"></mo>"
            + "</mrow></math>";
        var normalized = Normalize(source);
        var fences = normalized.Descendants(P + "mfenced").ToArray();
        Assert.Equal(2, fences.Length);
        Assert.Equal(new[] { "⟮", "⏐" }, fences.Select(f => (string?)f.Attribute("open")));
        Assert.Equal(new[] { "⟯", "" }, fences.Select(f => (string?)f.Attribute("close")));
        var forward = Forward(source);
        var delimiters = forward.Descendants(M + "d").ToArray();
        Assert.Equal(2, delimiters.Length);
        Assert.Equal(new[] { "⟮", "⏐" }, delimiters.Select(d => Delimiter(d, "begChr", "(")));
        Assert.Equal(new[] { "⟯", "" }, delimiters.Select(d => Delimiter(d, "endChr", ")")));
        Assert.Equal("x+y", NativeText(forward));
        var reverse = XDocument.Parse(WordOmmlConverter.TransformOmmlToMathMl(forward.ToString(), false));
        Assert.Equal(new[] { "⟮", "⏐" }, reverse.Descendants(P + "mfenced").Select(f => (string?)f.Attribute("open")));
        Assert.Equal(new[] { "⟯", "" }, reverse.Descendants(P + "mfenced").Select(f => (string?)f.Attribute("close")));
    }

    [Theory]
    [InlineData("⟦", "")]
    [InlineData("", "⟧")]
    [InlineData("⦃", "⦄")]
    public void ExplicitStructureSupportsDelimitersOutsideLegacyTableCharacterList(string open, string close)
    {
        var source = Source(Row("<mi>x</mi>", open, close));
        var fence = Assert.Single(Normalize(source).Descendants(P + "mfenced"));
        Assert.Equal(open, (string?)fence.Attribute("open"));
        Assert.Equal(close, (string?)fence.Attribute("close"));
        var native = Assert.Single(Forward(source).Descendants(M + "d"));
        Assert.Equal(open, Delimiter(native, "begChr", "("));
        Assert.Equal(close, Delimiter(native, "endChr", ")"));
    }

    [Fact]
    public void AllBodyChildrenRemainOneOperandWithoutInventedSeparators()
    {
        var source = Source(Row("<mi>a</mi><!-- body boundary --><mo>+</mo><mi>b</mi><mo>,</mo><mi>c</mi>"));
        var normalized = Normalize(source);
        var fence = Assert.Single(normalized.Descendants(P + "mfenced"));
        Assert.Equal("", (string?)fence.Attribute("separators"));
        var operand = Assert.Single(fence.Elements());
        Assert.Equal(P + "mrow", operand.Name);
        Assert.Equal(new[] { "mi", "mo", "mi", "mo", "mi" }, operand.Elements().Select(e => e.Name.LocalName));
        Assert.Single(operand.Nodes().OfType<XComment>());
        var forward = Forward(source);
        var delimiter = Assert.Single(forward.Descendants(M + "d"));
        Assert.Single(delimiter.Elements(M + "e"));
        Assert.Equal("a+b,c", NativeText(delimiter));
        Assert.Equal(normalized.ToString(), Normalize(normalized.ToString()).ToString());
    }

    [Fact]
    public void NestedRowsAndScriptAttachmentArePreservedInsideOut()
    {
        var source = Source("<msup>" + Row("<mi>a</mi><mo>+</mo>" + Row("<mi>b</mi>", "[", "]"))
            + "<mn>2</mn></msup>");
        var normalized = Normalize(source);
        Assert.Equal(2, normalized.Descendants(P + "mfenced").Count());
        var script = Assert.Single(Forward(source).Descendants(M + "sSup"));
        var outer = Assert.Single(script.Element(M + "e")!.Elements(M + "d"));
        Assert.Single(outer.Descendants(M + "d"));
        Assert.Equal("a+b", NativeText(outer));
        Assert.Equal("2", NativeText(script.Element(M + "sup")!));
    }

    [Fact]
    public void RowAttributesInheritanceAndAppliedFunctionBoundaryStayInPlace()
    {
        var source = Source("<mstyle mathvariant=\"bold\"><mi>f</mi><mo>&#x2061;</mo>"
            + Row("<mi>x</mi><mo>+</mo><mi mathvariant=\"normal\">y</mi>",
                rowAttributes: "id=\"argument\" mathcolor=\"blue\" scriptlevel=\"1\"")
            + "</mstyle>");
        var before = XDocument.Parse(source);
        var normalized = Normalize(source);
        var row = normalized.Descendants(P + "mrow").Single(r => (string?)r.Attribute("id") == "argument");
        Assert.Equal(before.Descendants(P + "mrow").Single().Attributes().Select(a => a.ToString()),
            row.Attributes().Select(a => a.ToString()));
        Assert.Equal(new[] { "mi", "mo", "mrow" }, normalized.Root!.Element(P + "mstyle")!
            .Elements().Select(e => e.Name.LocalName));
        Assert.Equal("bold", (string?)row.Parent!.Attribute("mathvariant"));
        Assert.Equal("normal", (string?)row.Descendants(P + "mi").Last().Attribute("mathvariant"));
        var forward = Forward(source);
        var function = Assert.Single(forward.Descendants(M + "func"));
        Assert.Single(function.Element(M + "e")!.Descendants(M + "d"));
        Assert.Equal("f", NativeText(function.Element(M + "fName")!));
        Assert.Equal("x+y", NativeText(function.Element(M + "e")!));
    }

    [Theory]
    [InlineData("fence=\"false\"")]
    [InlineData("stretchy=\"false\"")]
    [InlineData("fence=\"FALSE\" stretchy=\"true\"")]
    public void ExplicitFixedFlagsAreNeverPromotedEvenAroundTables(string fixedAttributes)
    {
        foreach (var body in new[] { "<mi>x</mi>", Table })
        foreach (var fixedLeft in new[] { true, false })
        {
            var source = Source(Row(body, "(", ")", fixedLeft ? fixedAttributes : "",
                fixedLeft ? "" : fixedAttributes));
            Assert.Empty(Normalize(source).Descendants(P + "mfenced"));
            Assert.Empty(Forward(source).Descendants(M + "d"));
        }
    }

    [Theory]
    [InlineData("mathvariant=\"bold\"")]
    [InlineData("mathsize=\"18pt\"")]
    [InlineData("minsize=\"2em\" maxsize=\"3em\"")]
    [InlineData("mathcolor=\"red\"")]
    [InlineData("style=\"font-size: 20px\"")]
    [InlineData("data-mjx-alternate=\"1\"")]
    [InlineData("symmetric=\"false\"")]
    public void UnsupportedDelimiterLocalPresentationRemainsExplicitInsteadOfBeingDiscarded(string localAttributes)
    {
        foreach (var body in new[] { "<mi>x</mi>", Table })
        foreach (var localOnLeft in new[] { true, false })
        {
            var source = Source(Row(body, openAttributes: localOnLeft ? localAttributes : "",
                closeAttributes: localOnLeft ? "" : localAttributes));
            var normalized = Normalize(source);
            Assert.Empty(normalized.Descendants(P + "mfenced"));
            Assert.True(XNode.DeepEquals(XDocument.Parse(source), normalized));
        }
    }

    [Fact]
    public void DistinctPerDelimiterStylesAndSizesRemainOnTheirOriginalTokens()
    {
        var source = Source(Row("<mi>x</mi>", openAttributes: "mathvariant=\"bold\" mathsize=\"18pt\"",
            closeAttributes: "mathvariant=\"italic\" mathsize=\"10pt\""));
        var normalized = Normalize(source);
        Assert.Empty(normalized.Descendants(P + "mfenced"));
        Assert.True(XNode.DeepEquals(XDocument.Parse(source), normalized));
        Assert.Equal(new[] { "bold", "italic" },
            normalized.Descendants(P + "mo").Select(mo => (string?)mo.Attribute("mathvariant")));
        Assert.Equal(new[] { "18pt", "10pt" },
            normalized.Descendants(P + "mo").Select(mo => (string?)mo.Attribute("mathsize")));
    }

    [Theory]
    [InlineData("")]
    [InlineData(" data-mjx-texclass=\"ORD\"")]
    public void UnmarkedOrOrdinaryRowsDoNotInferRolesFromCharacters(string attributes)
    {
        var source = Source($"<mrow{attributes}><mo data-mjx-texclass=\"OPEN\">(</mo><mi>x</mi>"
            + "<mo data-mjx-texclass=\"CLOSE\">)</mo></mrow>");
        Assert.Empty(Normalize(source).Descendants(P + "mfenced"));
    }

    [Theory]
    [InlineData("OPEN", "CLOSE")]
    [InlineData("CLOSE", "OPEN")]
    public void UnmarkedOperatorsAndWrongExplicitRolesDoNotBecomeFences(string first, string last)
    {
        var body = "<mrow data-mjx-texclass=\"INNER\"><mo>(</mo><mi>x</mi><mo>)</mo></mrow>";
        Assert.Empty(Normalize(Source(body)).Descendants(P + "mfenced"));
        if (first == "CLOSE")
        {
            body = $"<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"{first}\">(</mo><mi>x</mi>"
                + $"<mo data-mjx-texclass=\"{last}\">)</mo></mrow>";
            Assert.Empty(Normalize(Source(body)).Descendants(P + "mfenced"));
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ExistingTableRepresentationStillHasDirectTableAndNativeMatrix(bool explicitRoles)
    {
        var source = explicitRoles ? Source(Row(Table, "|", "|"))
            : Source("<mrow><mo>|</mo>" + Table + "<mo>|</mo></mrow>");
        var fence = Assert.Single(Normalize(source).Descendants(P + "mfenced"));
        Assert.Equal(P + "mtable", Assert.Single(fence.Elements()).Name);
        var delimiter = Assert.Single(Forward(source).Descendants(M + "d"));
        var matrix = Assert.Single(delimiter.Descendants(M + "m"));
        Assert.Equal(2, matrix.Elements(M + "mr").Count());
        Assert.All(matrix.Elements(M + "mr"), r => Assert.Equal(2, r.Elements(M + "e").Count()));
    }

    [Fact]
    public void NamespaceAliasesUseExpandedNamesAndForeignRowsAreUntouched()
    {
        var source = $"<p:math xmlns:p=\"{P}\" xmlns:q=\"urn:other\">"
            + "<p:mrow data-mjx-texclass=\"INNER\"><p:mo data-mjx-texclass=\"OPEN\">⟮</p:mo>"
            + "<p:mi>x</p:mi><p:mo data-mjx-texclass=\"CLOSE\">⟯</p:mo></p:mrow>"
            + "<q:mrow data-mjx-texclass=\"INNER\"><q:mo data-mjx-texclass=\"OPEN\">(</q:mo>"
            + "<q:mi>y</q:mi><q:mo data-mjx-texclass=\"CLOSE\">)</q:mo></q:mrow></p:math>";
        var normalized = Normalize(source);
        Assert.Single(normalized.Descendants(P + "mfenced"));
        Assert.True(XNode.DeepEquals(XDocument.Parse(source).Root!.Elements().Last(),
            normalized.Root!.Elements().Last()));
    }
}
