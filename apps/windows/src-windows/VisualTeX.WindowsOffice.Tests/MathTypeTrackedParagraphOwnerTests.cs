using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeTrackedParagraphOwnerTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";
    private static readonly XNamespace Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
    private const string Locator = "VTMT_TX_0123456789abcdef0123456789abcdef";
    private static XElement Run(string id, string progId = "Equation.DSMT4") => new(W + "r",
        new XElement(W + "object", new XElement(O + "OLEObject",
            new XAttribute("ProgID", progId), new XAttribute("ObjectID", id))));
    private static XElement P(string alignment, bool deletedMark, params object[] content) => new(W + "p",
        new XElement(W + "pPr", new XElement(W + "jc", new XAttribute(W + "val", alignment)),
            deletedMark ? new XElement(W + "rPr", new XElement(W + "del", new XAttribute(W + "author", "Old author"))) : null), content);
    private static string Package(params XElement[] content) => new XDocument(new XElement(Pkg + "package",
        new XElement(Pkg + "part", new XAttribute(Pkg + "name", "/word/document.xml"),
            new XElement(Pkg + "xmlData", new XElement(W + "document", new XElement(W + "body", content)))))).ToString();
    private static XElement Prepare(string owner, int ordinal, int count) => MathTypeLocalFragment.Body(XDocument.Parse(
        MathTypeLocalFragment.Prepare(Package(new XElement(W + "p", Run("new-export-id"))), owner, Locator, ordinal, count)));

    [Theory]
    [InlineData("")]
    [InlineData("ins")]
    [InlineData("moveTo")]
    public void OmmlReplacementChangesOnlyTheSelectedObjectAtom(string wrapper)
    {
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var atom = Run("source");
        XElement content = wrapper.Length == 0 ? atom : new XElement(W + wrapper,
            new XAttribute(W + "id", "14"), new XAttribute(W + "author", "Original author"), atom);
        var original = Package(new XElement(W + "p", content));
        var owner = Package(P("center", false, content, new XElement(W + "r", new XElement(W + "t", "outside"))));
        var omml = new XElement(m + "oMath", new XElement(m + "r", new XElement(m + "t", "x+1"))).ToString();
        var result = MathTypeLocalFragment.Body(XDocument.Parse(MathTypeLocalFragment.PrepareOmmlReplacement(
            original, owner, Locator, omml, 1, 1)));
        var equation = Assert.Single(result.Descendants(m + "oMath"));
        Assert.Empty(result.Descendants(O + "OLEObject"));
        Assert.Equal("center", (string?)result.Element(W + "p")?.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val"));
        Assert.DoesNotContain("outside", result.Value);
        if (wrapper.Length != 0)
            Assert.Equal("Original author", (string?)equation.Parent?.Attribute(W + "author"));
        Assert.Single(result.Descendants(W + "bookmarkStart"));
    }

    [Fact]
    public void OmmlReplacementDoesNotReimportDeletedPredecessorOrNeighboringOle()
    {
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var source = Package(new XElement(W + "p", Run("exact-export")));
        var owner = Package(P("left", true, new XElement(W + "del", Run("deleted", "Equation.3"))),
            P("both", false, Run("selected-export"), Run("neighbor")));
        var result = MathTypeLocalFragment.Body(XDocument.Parse(MathTypeLocalFragment.PrepareOmmlReplacement(
            source, owner, Locator, new XElement(m + "oMath", new XElement(m + "r", new XElement(m + "t", "z"))).ToString(), 2, 3)));
        Assert.Single(result.Descendants(m + "oMath"));
        Assert.Empty(result.Descendants(O + "OLEObject"));
        Assert.Empty(result.Descendants(W + "del"));
        Assert.Equal("both", (string?)result.Element(W + "p")?.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val"));
    }

    [Fact]
    public void OmmlReplacementRejectsMixedSourceRunRatherThanDroppingUserText()
    {
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var run = Run("source");
        run.Add(new XElement(W + "t", "user text"));
        var source = Package(P("left", false, run));
        Assert.Throws<InvalidDataException>(() => MathTypeLocalFragment.PrepareOmmlReplacement(
            source, source, Locator, new XElement(m + "oMath", new XElement(m + "r", new XElement(m + "t", "x"))).ToString(), 1, 1));
    }

    [Fact]
    public void OmmlReplacementRejectsDeletedTargetAndNonEquationPayload()
    {
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var current = Package(P("left", false, Run("current")));
        var deleted = Package(P("left", false, new XElement(W + "del", Run("deleted"))));
        var omml = new XElement(m + "oMath", new XElement(m + "r", new XElement(m + "t", "x"))).ToString();
        Assert.Throws<InvalidOperationException>(() => MathTypeLocalFragment.PrepareOmmlReplacement(deleted, deleted, Locator, omml, 1, 1));
        Assert.Throws<InvalidDataException>(() => MathTypeLocalFragment.PrepareOmmlReplacement(current, current, Locator, "<not-math/>", 1, 1));
    }

    [Fact]
    public void DeletedPreviousParagraphDoesNotBecomeTheTargetOwner()
    {
        var owner = Package(P("center", true, new XElement(W + "del", Run("old-export-id", "Equation.3"))),
            P("both", false, Run("different-export-id"), new XElement(W + "r", new XElement(W + "t", "remaining prose"))));
        var result = Prepare(owner, 2, 2);
        Assert.Equal("both", (string?)result.Element(W + "p")?.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val"));
        Assert.Equal("new-export-id", (string?)Assert.Single(result.Descendants(O + "OLEObject")).Attribute("ObjectID"));
        Assert.Empty(result.Descendants(W + "del"));
        Assert.DoesNotContain("remaining prose", result.Value);
    }

    [Fact]
    public void CurrentObjectBeforeADeletedParagraphMarkUsesTheSurvivingMarkProperties()
    {
        var owner = Package(P("left", true, Run("a")), P("right", false, Run("b")));
        var result = Prepare(owner, 1, 2);
        Assert.Equal("right", (string?)result.Element(W + "p")?.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val"));
        Assert.Empty(result.Descendants(W + "del"));
    }

    [Fact]
    public void NoOrdinalCannotGuessBetweenDuplicateNativeObjects()
    {
        var owner = Package(P("left", false, Run("same")), P("right", false, Run("same")));
        Assert.Throws<InvalidDataException>(() => Prepare(owner, 0, 2));
        Assert.Throws<InvalidDataException>(() => Prepare(owner, 2, 3));
    }

    [Fact]
    public void AnOrdinalCannotPromoteADeletedOleToCurrentContent()
    {
        var owner = Package(P("center", true, new XElement(W + "del", Run("a"))), P("both", false, Run("b")));
        Assert.Throws<InvalidDataException>(() => Prepare(owner, 1, 2));
    }

    [Fact]
    public void DeletedParagraphMarkCannotBorrowFormattingFromAnotherCell()
    {
        var owner = Package(new XElement(W + "tbl", new XElement(W + "tr",
            new XElement(W + "tc", P("left", true, Run("a"))),
            new XElement(W + "tc", P("right", false, Run("b"))))));
        Assert.Throws<InvalidDataException>(() => Prepare(owner, 1, 2));
    }

    [Fact]
    public void ACellTrackedJoinKeepsOnlyOneParagraphInItsImport()
    {
        var owner = Package(new XElement(W + "tbl", new XElement(W + "tr",
            new XElement(W + "tc", new XElement(W + "p")),
            new XElement(W + "tc", P("left", true, Run("a")), P("center", false, Run("b"))))), new XElement(W + "p"));
        var result = Prepare(owner, 1, 2);
        Assert.Equal(W + "p", Assert.Single(result.Elements()).Name);
        Assert.Equal("center", (string?)result.Element(W + "p")?.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val"));
        Assert.Empty(result.Descendants(W + "tbl"));
    }
}
