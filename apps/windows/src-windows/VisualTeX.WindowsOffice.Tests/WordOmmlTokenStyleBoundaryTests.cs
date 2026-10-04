using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlTokenStyleBoundaryTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace P = "http://www.w3.org/1998/Math/MathML";

    private static string Source(string body) => $"<math xmlns=\"{P}\">{body}</math>";
    private static XDocument Forward(string body) =>
        XDocument.Parse(WordOmmlConverter.TransformMathMlToOmml(Source(body)));
    private static XElement[] Runs(XContainer document) => document.Descendants(M + "r").ToArray();
    private static string Text(XElement run) => string.Concat(run.Elements(M + "t").Select(t => t.Value));
    private static string? Style(XElement run) =>
        (string?)run.Element(M + "rPr")?.Element(M + "sty")?.Attribute(M + "val");

    [Theory]
    [InlineData("∥")]
    [InlineData("∦")]
    [InlineData("⋅")]
    [InlineData("×")]
    [InlineData("+")]
    public void AdjacentIdenticalOperatorGlyphsKeepIndependentEffectiveStyles(string glyph)
    {
        var document = Forward($"<mo>{glyph}</mo><mo mathvariant=\"italic\">{glyph}</mo>"
            + $"<mo>{glyph}</mo><mo mathvariant=\"bold\">{glyph}</mo>"
            + $"<mo mathvariant=\"bold-italic\">{glyph}</mo>");
        var runs = Runs(document);
        Assert.Equal(new[] { glyph, glyph, glyph, glyph, glyph }, runs.Select(Text));
        Assert.Equal(new[] { "p", "i", "p", "b", "bi" }, runs.Select(Style));
        Assert.DoesNotContain(document.Descendants().Attributes(),
            attribute => attribute.Name.LocalName.StartsWith("data-visualtex", StringComparison.Ordinal));

        var reverse = XDocument.Parse(WordOmmlConverter.TransformOmmlToMathMl(
            document.ToString(SaveOptions.DisableFormatting), false));
        Assert.Equal(new[] { "normal", "italic", "normal", "bold", "bold-italic" },
            reverse.Descendants(P + "mo").Select(token => (string?)token.Attribute("mathvariant")));
    }

    [Fact]
    public void SharedTokenContractMaterializesDefaultsAndNearestDeclarationBeforeGrouping()
    {
        var source = Source("<mi>x</mi><mo>+</mo><mn>2</mn><mtext>ABC</mtext><ms>ABC</ms>"
            + "<mstyle mathvariant=\"normal\"><mstyle mathvariant=\"italic\">"
            + "<mi>ABC</mi><mo>∥</mo><mn>3</mn><mo mathvariant=\"bold-italic\">∦</mo>"
            + "</mstyle><mo>∥</mo></mstyle>");
        var projected = XDocument.Parse(WordOmmlConverter.NormalizeMathMlTokenStyleAuthority(source));
        var tokens = projected.Descendants().Where(token =>
            token.Name.Namespace == P && token.Name.LocalName is "mi" or "mo" or "mn" or "mtext" or "ms").ToArray();
        Assert.Equal(new[] { "italic", "normal", "normal", "normal", "normal",
            "italic", "italic", "italic", "bold-italic", "normal" },
            tokens.Select(token => (string?)token.Attribute("mathvariant")));
        foreach (var token in tokens)
            Assert.Equal(MathMlTokenStyleContract.ResolveEffectiveVariant(token),
                (string?)token.Attribute("mathvariant"));
    }

    [Fact]
    public void NestedVariantsKeepExplicitNormalOperatorDistinctFromInheritedItalic()
    {
        var document = Forward("<mstyle mathvariant=\"normal\"><mo>∥</mo>"
            + "<mstyle mathvariant=\"italic\"><mo>∥</mo><mo mathvariant=\"normal\">∥</mo>"
            + "<mo mathvariant=\"bold-italic\">∦</mo></mstyle><mo>∦</mo></mstyle>");
        Assert.Equal(new[] { "∥", "∥", "∥", "∦", "∦" }, Runs(document).Select(Text));
        Assert.Equal(new[] { "p", "i", "p", "bi", "p" }, Runs(document).Select(Style));
    }

    [Fact]
    public void DefaultIdentifiersOperatorsAndDigitsUseTheirOwnNativeStyle()
    {
        var document = Forward("<mi>a</mi><mo>∥</mo><mi>b</mi><mn>2</mn>");
        Assert.Equal(new[] { "a", "∥", "b", "2" }, Runs(document).Select(Text));
        Assert.Equal(new[] { "i", "p", "i", "p" }, Runs(document).Select(Style));
    }

    [Fact]
    public void SingleCharacterFunctionNameKeepsTheSharedIdentifierDefault()
    {
        var document = Forward("<mi>f</mi><mo>&#x2061;</mo><mi>x</mi>");
        var function = Assert.Single(document.Descendants(M + "func"));
        var name = Assert.Single(function.Element(M + "fName")!.Descendants(M + "r"));
        Assert.Equal("f", Text(name));
        Assert.Equal("i", Style(name));
        Assert.Equal("x", string.Concat(function.Element(M + "e")!.Descendants(M + "t").Select(t => t.Value)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FunctionMarkerStartsAfterItsNameAndParenthesizedArgumentKeepsItsTail(bool namedOperator)
    {
        var name = namedOperator ? "<mo data-mjx-texclass=\"OP\">sin</mo>" : "<mi>sin</mi>";
        var document = Forward("<mo>+</mo>" + name + "<mo>&#x2061;</mo><mo>(</mo>"
            + "<mi>x</mi><mo>+</mo><mi>y</mi><mo>)</mo><mo>+</mo><mi>z</mi>");
        var function = Assert.Single(document.Descendants(M + "func"));
        var nameRun = Assert.Single(function.Element(M + "fName")!.Descendants(M + "r"));
        Assert.Equal("sin", Text(nameRun));
        Assert.Equal("p", Style(nameRun));
        Assert.Equal("(x+y)", string.Concat(function.Element(M + "e")!.Descendants(M + "t").Select(t => t.Value)));
        Assert.Equal(new[] { "+", "+", "z" }, document.Root!.Elements(M + "r").Select(Text));
        Assert.DoesNotContain(document.Descendants(M + "t"), text => text.Value.Contains('\u2061'));
    }

    [Fact]
    public void NativeThreeChildFunctionRowRemainsOneFunction()
    {
        var document = Forward("<mrow><mi>sin</mi><mo>&#x2061;</mo><mi>x</mi></mrow>");
        var function = Assert.Single(document.Descendants(M + "func"));
        Assert.Equal("sin", function.Element(M + "fName")!.Value);
        Assert.Equal("x", function.Element(M + "e")!.Value);
    }

    [Fact]
    public void EqualNormalStyleWithoutApplyFunctionDoesNotCreateAFunction()
    {
        var document = Forward("<mi>sin</mi><mo>+</mo><mi mathvariant=\"normal\">ABC</mi>");
        Assert.Empty(document.Descendants(M + "func"));
        Assert.Equal(new[] { "sin", "+", "ABC" }, Runs(document).Select(Text));
        Assert.Equal(new[] { "p", "p", "p" }, Runs(document).Select(Style));
    }

    [Fact]
    public void NamedOperatorsDoNotAbsorbAdjacentEqualStylePunctuation()
    {
        var document = Forward("<mo data-mjx-texclass=\"OP\">lim</mo><mo>+</mo>"
            + "<mo data-mjx-texclass=\"OP\" mathvariant=\"italic\">lim</mo><mo>+</mo>"
            + "<mo data-mjx-texclass=\"OP\" mathvariant=\"bold-italic\">lim</mo>");
        Assert.Equal(new[] { "lim", "+", "lim", "+", "lim" }, Runs(document).Select(Text));
        Assert.Equal(new[] { "p", "p", "i", "p", "bi" }, Runs(document).Select(Style));
    }

    [Fact]
    public void SameWordInProseNormalMathAndItalicMathKeepsItsOwnRole()
    {
        var document = Forward("<mtext>ABC</mtext><mo>+</mo><mi mathvariant=\"normal\">ABC</mi>"
            + "<mo>+</mo><mi mathvariant=\"italic\">ABC</mi>");
        var words = Runs(document).Where(run => Text(run) == "ABC").ToArray();
        Assert.Equal(3, words.Length);
        Assert.NotNull(words[0].Element(M + "rPr")?.Element(M + "nor"));
        Assert.Null(words[0].Element(M + "rPr")?.Element(M + "sty"));
        Assert.Null(words[1].Element(M + "rPr")?.Element(M + "nor"));
        Assert.Equal("p", Style(words[1]));
        Assert.Null(words[2].Element(M + "rPr")?.Element(M + "nor"));
        Assert.Equal("i", Style(words[2]));
    }

    [Theory]
    [InlineData("a", "normal", "a")]
    [InlineData("ABC words", "normal", "ABC\u00A0words")]
    [InlineData("x + 2 words", "normal", "x\u00A0+\u00A02\u00A0words")]
    [InlineData("a", "italic", "a")]
    [InlineData("ABC words", "bold", "ABC\u00A0words")]
    [InlineData("x + 2 words", "bold-italic", "x\u00A0+\u00A02\u00A0words")]
    public void ProseTokensKeepNormalTextRoleThroughBothBoundaries(
        string text, string variant, string expectedReverseText)
    {
        var document = Forward($"<mtext mathvariant=\"{variant}\">{text}</mtext>");
        var run = Assert.Single(Runs(document));
        Assert.NotNull(run.Element(M + "rPr")?.Element(M + "nor"));
        Assert.Equal(text, Text(run));
        var reverse = XDocument.Parse(WordOmmlConverter.TransformOmmlToMathMl(
            document.ToString(SaveOptions.DisableFormatting), false));
        // Stock OMML2MML projects prose spaces as NBSP. Keep both boundary
        // expectations explicit; this role test does not normalize whitespace.
        Assert.Equal(expectedReverseText, Assert.Single(reverse.Descendants(P + "mtext")).Value);
        Assert.Empty(reverse.Descendants(P + "mi"));
        Assert.Empty(reverse.Descendants(P + "mo"));
        Assert.Empty(reverse.Descendants(P + "mn"));
    }

    [Fact]
    public void StyleProjectionLeavesFractionAndScriptOwnershipIntact()
    {
        var document = Forward("<mfrac><mrow><mo mathvariant=\"italic\">∥</mo><mi>x</mi></mrow>"
            + "<msup><mi>y</mi><mn>2</mn></msup></mfrac>");
        var fraction = Assert.Single(document.Descendants(M + "f"));
        var script = Assert.Single(fraction.Descendants(M + "sSup"));
        Assert.Equal(new[] { "∥", "x" }, fraction.Element(M + "num")!.Descendants(M + "r").Select(Text));
        Assert.All(fraction.Element(M + "num")!.Descendants(M + "r"),
            run => Assert.Equal("i", Style(run)));
        Assert.Equal("i", Style(Assert.Single(script.Element(M + "e")!.Descendants(M + "r"))));
        Assert.Equal("p", Style(Assert.Single(script.Element(M + "sup")!.Descendants(M + "r"))));
    }

    [Fact]
    public void RepresentativeItalicParallelOperatorsReachReverseProjectionWithExplicitStyle()
    {
        var document = Forward("<mrow><mi mathvariant=\"italic\">a</mi><mo mathvariant=\"italic\">∥</mo>"
            + "<mi mathvariant=\"italic\">b</mi><mo mathvariant=\"italic\">∦</mo>"
            + "<mi mathvariant=\"italic\">a</mi><mo mathvariant=\"italic\">∥</mo>"
            + "<mi mathvariant=\"italic\">b</mi><mo mathvariant=\"italic\">∦</mo></mrow>");
        var operators = Runs(document).Where(run => Text(run) is "∥" or "∦").ToArray();
        Assert.Equal(4, operators.Length);
        Assert.All(operators, run => Assert.Equal("i", Style(run)));
        var reverse = XDocument.Parse(WordOmmlConverter.TransformOmmlToMathMl(
            document.ToString(SaveOptions.DisableFormatting), false));
        Assert.Equal(new[] { "∥", "∦", "∥", "∦" }, reverse.Descendants(P + "mo").Select(token => token.Value));
        Assert.All(reverse.Descendants(P + "mo"),
            token => Assert.Equal("italic", (string?)token.Attribute("mathvariant")));
    }
}
