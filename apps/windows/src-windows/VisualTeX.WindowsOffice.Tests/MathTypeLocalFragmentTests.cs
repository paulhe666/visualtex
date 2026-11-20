using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeLocalFragmentTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";
    private static readonly XNamespace Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
    private const string Locator = "VTMT_TX_0123456789abcdef0123456789abcdef";

    [Fact]
    public void ObjectImportUsesTheActualOwnerParagraphPropertiesWithoutTableOrSectionScaffolding()
    {
        var objectRun = ObjectRun();
        var source = Package(new XElement(W + "tbl", new XElement(W + "tr", new XElement(W + "tc",
            new XElement(W + "p", objectRun)))), new XElement(W + "sectPr"));
        var properties = new XElement(W + "pPr",
            new XElement(W + "jc", new XAttribute(W + "val", "center")),
            new XElement(W + "ind", new XAttribute(W + "firstLineChars", "200")),
            new XElement(W + "tabs", new XElement(W + "tab", new XAttribute(W + "val", "right"), new XAttribute(W + "pos", "8306"))),
            new XElement(W + "pPrChange", new XAttribute(W + "author", "Original author"), new XElement(W + "pPr")));
        var owner = Package(new XElement(W + "p", properties, new XElement(W + "r", new XElement(W + "t", "outside")), ObjectRun()));
        var output = XDocument.Parse(MathTypeLocalFragment.Prepare(source, owner, Locator));
        var body = MathTypeLocalFragment.Body(output);
        var paragraph = Assert.Single(body.Elements());
        Assert.Equal(W + "p", paragraph.Name);
        Assert.True(XNode.DeepEquals(WithoutPrefixes(properties), WithoutPrefixes(paragraph.Element(W + "pPr")!)),
            "Expected: " + properties + " Actual: " + paragraph.Element(W + "pPr"));
        Assert.Single(paragraph.Descendants(O + "OLEObject"));
        Assert.Empty(body.Descendants(W + "tbl"));
        Assert.Empty(body.Descendants(W + "sectPr"));
        Assert.DoesNotContain("outside", body.Value);
        Assert.Equal("AAEC", output.Root!.Elements(Pkg + "part").Last().Element(Pkg + "binaryData")!.Value);
    }

    [Fact]
    public void CellOwnerPropertiesAreResolvedWithoutAdoptingEmptyTransportParagraphs()
    {
        var source = Package(new XElement(W + "p", ObjectRun()));
        var properties = new XElement(W + "pPr", new XElement(W + "jc", new XAttribute(W + "val", "center")));
        var owner = Package(new XElement(W + "tbl", new XElement(W + "tr",
            new XElement(W + "tc", new XElement(W + "p")),
            new XElement(W + "tc", new XElement(W + "p", properties, ObjectRun())))),
            new XElement(W + "p"), new XElement(W + "sectPr"));
        var result = MathTypeLocalFragment.Body(XDocument.Parse(MathTypeLocalFragment.Prepare(source, owner, Locator)));
        Assert.Equal(W + "p", Assert.Single(result.Elements()).Name);
        Assert.True(XNode.DeepEquals(WithoutPrefixes(properties), WithoutPrefixes(result.Elements().Single().Element(W + "pPr")!)));
        Assert.Single(result.Descendants(O + "OLEObject"));
    }

    [Fact]
    public void RevisionProvenanceAndExistingLocalBookmarksAreNotStripped()
    {
        var ins = new XElement(W + "ins", new XAttribute(W + "id", "5"),
            new XAttribute(W + "author", "Earlier author"), new XAttribute(W + "date", "2004-01-02T03:04:05Z"), ObjectRun());
        var paragraph = new XElement(W + "p",
            new XElement(W + "bookmarkStart", new XAttribute(W + "id", "0"), new XAttribute(W + "name", "_RefKeep")),
            ins, new XElement(W + "bookmarkEnd", new XAttribute(W + "id", "0")));
        var source = Package(paragraph);
        var output = MathTypeLocalFragment.Body(XDocument.Parse(MathTypeLocalFragment.Prepare(source, source, Locator)));
        Assert.True(XNode.DeepEquals(WithoutPrefixes(ins), WithoutPrefixes(Assert.Single(output.Descendants(W + "ins")))),
            "Expected: " + ins + " Actual: " + output);
        var names = MathTypeLocalFragment.BookmarkNames(Package(output.Elements().Single()));
        Assert.Contains("_RefKeep", names);
        Assert.Contains(Locator, names);
        var starts = output.Descendants(W + "bookmarkStart").ToArray();
        Assert.Equal(2, starts.Select(e => (string)e.Attribute(W + "id")!).Distinct().Count());
    }

    [Theory]
    [InlineData("del")]
    [InlineData("moveFrom")]
    public void ADeletedObjectCannotBeReintroducedAsTheCurrentFormula(string kind)
    {
        var source = Package(new XElement(W + "p", new XElement(W + kind, ObjectRun())));
        Assert.Throws<InvalidOperationException>(() => MathTypeLocalFragment.Prepare(source, source, Locator));
    }

    [Fact]
    public void RecoveryComparesExactLiveGeometryInsteadOfTransientWordPreviewBounds()
    {
        var source = RecoveryPackage("width:2in;height:30.85pt", "2860");
        var restored = RecoveryPackage("width:143pt;height:32pt", "2860");
        var expected = MathTypeLocalFragment.RecoveryEvidence(source, "0,26,1,Equation.DSMT4,143,32", 1);
        Assert.Equal(expected, MathTypeLocalFragment.RecoveryEvidence(restored, "0,26,1,Equation.DSMT4,143,32", 1));
        // A real resize, even below a point, must still fail the exact comparison.
        Assert.NotEqual(expected, MathTypeLocalFragment.RecoveryEvidence(restored, "0,26,1,Equation.DSMT4,143.05,32", 1));
        // Intrinsic geometry and other style properties remain authoritative XML.
        Assert.NotEqual(expected, MathTypeLocalFragment.RecoveryEvidence(
            RecoveryPackage("width:143pt;height:32pt", "2861"), "0,26,1,Equation.DSMT4,143,32", 1));
        Assert.NotEqual(expected, MathTypeLocalFragment.RecoveryEvidence(
            RecoveryPackage("width:143pt;height:32pt;rotation:5", "2860"), "0,26,1,Equation.DSMT4,143,32", 1));
        Assert.Contains("width:2in;height:30.85pt", source); // input document is not rewritten
    }

    [Fact]
    public void RecoveryDoesNotHideAMissingOrDuplicatedObject()
    {
        Assert.Throws<InvalidDataException>(() => MathTypeLocalFragment.RecoveryEvidence(
            RecoveryPackage("width:143pt;height:32pt", "2860"), "0,26,1,Equation.DSMT4,143,32", 0));
    }

    private static string RecoveryPackage(string style, string originalWidth)
    {
        XNamespace v = "urn:schemas-microsoft-com:vml";
        return Package(new XElement(W + "p", new XElement(W + "r", new XElement(W + "object",
            new XAttribute(W + "dxaOrig", originalWidth), new XAttribute(W + "dyaOrig", "640"),
            new XElement(v + "shape", new XAttribute("id", "shape1"), new XAttribute("style", style)),
            new XElement(O + "OLEObject", new XAttribute("Type", "Embed"),
                new XAttribute("ProgID", "Equation.DSMT4"), new XAttribute("ShapeID", "shape1"))))));
    }

    // XML namespace aliases are serialization choices, not native formatting.
    // Keep every namespace-qualified element/attribute and value in comparison.
    private static XElement WithoutPrefixes(XElement element)
    {
        var copy = new XElement(element);
        copy.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
        return copy;
    }

    private static XElement ObjectRun() => new(W + "r", new XElement(W + "object",
        new XAttribute(W + "dxaOrig", "200"), new XElement(O + "OLEObject",
            new XAttribute("ProgID", "Equation.DSMT4"), new XAttribute("ObjectID", "original"))));

    private static string Package(params XElement[] body) => new XDocument(new XElement(Pkg + "package",
        new XElement(Pkg + "part", new XAttribute(Pkg + "name", "/word/document.xml"),
            new XElement(Pkg + "xmlData", new XElement(W + "document", new XElement(W + "body", body.Select(e => new XElement(e)))))),
        new XElement(Pkg + "part", new XAttribute(Pkg + "name", "/word/embeddings/object.bin"),
            new XElement(Pkg + "binaryData", "AAEC")))).ToString(SaveOptions.DisableFormatting);
}
