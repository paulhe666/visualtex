using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlRevisionNumberingTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Theory]
    [InlineData("", false)]
    [InlineData("", true)]
    [InlineData("ins", false)]
    [InlineData("ins", true)]
    [InlineData("moveTo", false)]
    [InlineData("moveTo", true)]
    public void CurrentRevisionNumberIsRecognizedAndStrippedWithoutDuplicating(string revision, bool merged)
    {
        var source = Fixture(revision, merged);
        var text = source.ToString(SaveOptions.DisableFormatting);
        Assert.True(WordOmmlConverter.HasWordNativeEquationNumberHost(text));
        Assert.True(WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(text, null));
        foreach (var semantic in new[] {
            WordOmmlConverter.StripWordNativeEquationNumberHost(text),
            WordOmmlConverter.StripManagedVisualTeXNativeEquationNumber(text) })
        {
            var body = XElement.Parse(semantic);
            Assert.Empty(body.Descendants(M + "eqArr"));
            Assert.Empty(body.Descendants(W + "fldChar"));
            Assert.Equal(merged ? "xy" : "x", string.Concat(body.Descendants(M + "t").Select(t => t.Value)));
            Assert.DoesNotContain("#", semantic);
        }
    }

    [Theory]
    [InlineData("", false)]
    [InlineData("", true)]
    [InlineData("ins", false)]
    [InlineData("ins", true)]
    [InlineData("moveTo", false)]
    [InlineData("moveTo", true)]
    public void BodyReplacementKeepsNumberRevisionPropertiesAndFieldsVerbatim(string revision, bool merged)
    {
        var source = Fixture(revision, merged);
        var suffix = source.Descendants(M + "d").Single();
        var separatorProperties = source.Element(M + "eqArr")!.Element(M + "e")!
            .Elements(M + "r").Last().Element(W + "rPr")!;
        var replacement = WordOmmlConverter.ReplaceWordNativeEquationNumberHostBody(
            source.ToString(SaveOptions.DisableFormatting),
            new XElement(M + "oMath", new XElement(M + "r", new XElement(M + "t", "z"))).ToString());
        var after = XElement.Parse(replacement);
        Assert.Single(after.Descendants(M + "eqArr"));
        Assert.True(XNode.DeepEquals(WithoutNamespaceDeclarations(suffix),
            WithoutNamespaceDeclarations(after.Descendants(M + "d").Single())));
        var separator = after.Element(M + "eqArr")!.Element(M + "e")!.Elements(M + "r").Last();
        Assert.Equal("#", string.Concat(separator.Descendants(M + "t").Select(t => t.Value)));
        Assert.True(XNode.DeepEquals(WithoutNamespaceDeclarations(separatorProperties),
            WithoutNamespaceDeclarations(separator.Element(W + "rPr")!)));
        if (revision.Length != 0)
            Assert.Equal("number-author", (string?)separator.Element(W + revision)?.Attribute(W + "author"));
        Assert.True(WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(replacement, null));
    }

    [Theory]
    [InlineData("del")]
    [InlineData("moveFrom")]
    public void DeletedOrMovedFromSeparatorIsNotCurrentNumbering(string revision)
    {
        var source = Fixture(revision, true).ToString();
        Assert.False(WordOmmlConverter.HasWordNativeEquationNumberHost(source));
        Assert.False(WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(source, null));
    }

    [Theory]
    [InlineData("del")]
    [InlineData("moveFrom")]
    public void DeletedSequenceCodeCannotClaimAnActiveLiteralNumber(string revision)
    {
        var source = Fixture("", false);
        var instruction = source.Descendants(M + "t").Single(t => t.Value.Contains("SEQ "));
        instruction.ReplaceWith(new XElement(W + revision, new XAttribute(W + "id", "80"), new XElement(instruction)));
        Assert.True(WordOmmlConverter.HasWordNativeEquationNumberHost(source.ToString()));
        Assert.False(WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(source.ToString(), null));
    }

    [Fact]
    public void SplitTrackedSeparatorUsesOnlyCurrentTextAndPreservesBothAuthors()
    {
        var source = Fixture("ins", true);
        var separator = source.Element(M + "eqArr")!.Element(M + "e")!.Elements(M + "r").Last();
        separator.Element(W + "ins")!.ReplaceNodes(new XElement(M + "t", "y"));
        separator.Add(new XElement(W + "del", new XAttribute(W + "id", "90"), new XElement(M + "t", "old")));
        separator.Add(new XElement(W + "ins", new XAttribute(W + "id", "91"),
            new XAttribute(W + "author", "separator-author"), new XElement(M + "t", "#")));
        var semantic = XElement.Parse(WordOmmlConverter.StripWordNativeEquationNumberHost(source.ToString()));
        Assert.DoesNotContain(semantic.Descendants(M + "t"), t => t.Value == "#");
        Assert.Equal("number-author", (string?)semantic.Descendants(W + "ins").First().Attribute(W + "author"));
        var after = XElement.Parse(WordOmmlConverter.ReplaceWordNativeEquationNumberHostBody(source.ToString(),
            new XElement(M + "oMath", new XElement(M + "r", new XElement(M + "t", "z"))).ToString()));
        var hash = after.Descendants(M + "t").Single(t => t.Value == "#");
        Assert.Equal("separator-author", (string?)hash.Parent!.Attribute(W + "author"));
        Assert.Equal("old", after.Descendants(W + "del").Single().Value);
    }

    [Theory]
    [InlineData("2", 1, ".")]
    [InlineData("2.1", 2, "-")]
    [InlineData("IV", 1, ".")]
    public void LiteralHeadingIsNativeTextAndOneLiveSequence(string prefix, int level, string separator)
    {
        var semantic = new XElement(M + "oMath", new XElement(M + "r", new XElement(M + "t", "x"))).ToString();
        var xml = WordOmmlConverter.BuildWordNativeNumberedOmml(semantic, "UserEquation", null,
            level, separator, literalHeadingPrefix: prefix);
        var equation = XElement.Parse(xml);
        var instructions = equation.Descendants(W + "instrText").Select(e => e.Value).ToArray();
        Assert.Single(instructions);
        Assert.Contains("SEQ UserEquation", instructions[0]);
        Assert.Contains("\\s " + level, instructions[0]);
        Assert.DoesNotContain("STYLEREF", xml);
        Assert.Contains(equation.Descendants(M + "t"), t => t.Value == prefix);
        Assert.True(WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(xml, null));
        Assert.True(WordOmmlConverter.HasWordNativeEquationNumberHost(xml));
        Assert.DoesNotContain("UserEquation", WordOmmlConverter.StripWordNativeEquationNumberHost(xml));
        Assert.Empty(equation.Descendants(W + "bookmarkStart"));
    }

    [Fact]
    public void AutomaticHeadingKeepsStyleRefAndSequenceAndQuotedIdentifier()
    {
        var semantic = new XElement(M + "oMath", new XElement(M + "r", new XElement(M + "t", "x"))).ToString();
        var xml = WordOmmlConverter.BuildWordNativeNumberedOmml(semantic, "User Equation", null, 2, ".");
        var instructions = XElement.Parse(xml).Descendants(W + "instrText").Select(e => e.Value).ToArray();
        Assert.Equal(2, instructions.Length);
        Assert.Contains("STYLEREF 2", instructions[0]);
        Assert.Contains("SEQ \"User Equation\"", instructions[1]);
    }

    [Theory]
    [InlineData(0, "2")]
    [InlineData(1, "")]
    [InlineData(1, "2\r3")]
    [InlineData(1, "2\u00133")]
    public void InvalidLiteralPrefixCannotEnterANativeNumber(int level, string prefix)
    {
        var semantic = new XElement(M + "oMath", new XElement(M + "r", new XElement(M + "t", "x"))).ToString();
        Assert.Throws<ArgumentException>(() => WordOmmlConverter.BuildWordNativeNumberedOmml(
            semantic, "Equation", null, level, ".", literalHeadingPrefix: prefix));
    }

    private static XElement WithoutNamespaceDeclarations(XElement element)
    {
        var copy = new XElement(element);
        copy.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
        return copy;
    }

    private static XElement Fixture(string revision, bool merged)
    {
        XElement Content(XElement child) => revision.Length == 0 ? child :
            new XElement(W + revision, new XAttribute(W + "id", "23"),
                new XAttribute(W + "author", "number-author"),
                new XAttribute(W + "date", "2026-09-29T04:56:00Z"), child);
        XElement Run(XElement child) => new(M + "r",
            new XElement(W + "rPr", new XElement(W + "color", new XAttribute(W + "val", "445566"))), Content(child));
        var number = new XElement(M + "d", new XElement(M + "e",
            Run(new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"))),
            Run(new XElement(M + "t", " SEQ AnyUserSequence \\* ARABIC ")),
            Run(new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))),
            Run(new XElement(M + "t", "3")),
            Run(new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end")))));
        return new XElement(M + "oMath", new XElement(M + "eqArr", new XElement(M + "e",
            new XElement(M + "r", new XElement(M + "t", "x")),
            Run(new XElement(M + "t", merged ? "y#" : "#")), number)));
    }
}
