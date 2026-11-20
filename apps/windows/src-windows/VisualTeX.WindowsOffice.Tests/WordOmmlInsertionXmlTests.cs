using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlInsertionXmlTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace M =
        "http://schemas.openxmlformats.org/officeDocument/2006/math";

    [Fact]
    public void TrackedInsertionKeepsWordInsWrapperAroundOmml()
    {
        const string marker = "VTXmarker";
        var paragraph = XElement.Parse(
            $"<w:p xmlns:w=\"{W}\" xmlns:m=\"{M}\">"
            + "<w:r><w:t>before</w:t></w:r>"
            + "<w:ins w:id=\"2\" w:author=\"tester\">"
            + "<w:r><w:rPr><w:b/></w:rPr><w:t>left VTXmarker right</w:t></w:r>"
            + "</w:ins>"
            + "</w:p>");
        var markerNode = Assert.Single(
            paragraph.Descendants(W + "t"),
            node => node.Value.Contains(marker, StringComparison.Ordinal));
        var equation = XElement.Parse(
            $"<m:oMath xmlns:m=\"{M}\"><m:r><m:t>x</m:t></m:r></m:oMath>");

        WordOmmlInsertionXml.ReplaceOrdinaryMarkerRun(
            markerNode,
            marker,
            equation,
            paragraph);

        var insertion = Assert.Single(paragraph.Elements(W + "ins"));
        Assert.Equal("2", (string?)insertion.Attribute(W + "id"));
        Assert.Equal("tester", (string?)insertion.Attribute(W + "author"));
        var children = insertion.Elements().ToArray();
        Assert.Equal(3, children.Length);
        Assert.Equal(W + "r", children[0].Name);
        Assert.Equal("left ", children[0].Element(W + "t")?.Value);
        Assert.Equal("preserve", (string?)children[0].Element(W + "t")?.Attribute(XNamespace.Xml + "space"));
        Assert.Equal(M + "oMath", children[1].Name);
        Assert.Equal("x", children[1].Value);
        Assert.Equal(W + "r", children[2].Name);
        Assert.Equal(" right", children[2].Element(W + "t")?.Value);
        Assert.NotNull(children[0].Element(W + "rPr")?.Element(W + "b"));
        Assert.NotNull(children[2].Element(W + "rPr")?.Element(W + "b"));
        Assert.DoesNotContain(marker, paragraph.Value, StringComparison.Ordinal);
    }

    [Fact]
    public void DirectParagraphRunStillUsesSameLocalSplice()
    {
        const string marker = "VTXmarker";
        var paragraph = XElement.Parse(
            $"<w:p xmlns:w=\"{W}\" xmlns:m=\"{M}\">"
            + "<w:r><w:t>VTXmarker</w:t></w:r>"
            + "</w:p>");
        var markerNode = Assert.Single(paragraph.Descendants(W + "t"));
        var equation = XElement.Parse(
            $"<m:oMath xmlns:m=\"{M}\"><m:r><m:t>y</m:t></m:r></m:oMath>");

        WordOmmlInsertionXml.ReplaceOrdinaryMarkerRun(
            markerNode,
            marker,
            equation,
            paragraph);

        Assert.Single(paragraph.Elements(M + "oMath"));
        Assert.Equal("y", paragraph.Element(M + "oMath")?.Value);
        Assert.Empty(paragraph.Elements(W + "r"));
    }

    [Fact]
    public void MarkerFromDifferentParagraphIsRejected()
    {
        const string marker = "VTXmarker";
        var owner = XElement.Parse(
            $"<w:p xmlns:w=\"{W}\"><w:ins><w:r><w:t>{marker}</w:t></w:r></w:ins></w:p>");
        var other = XElement.Parse($"<w:p xmlns:w=\"{W}\"/>");
        var markerNode = Assert.Single(owner.Descendants(W + "t"));
        var equation = XElement.Parse(
            $"<m:oMath xmlns:m=\"{M}\"><m:r><m:t>z</m:t></m:r></m:oMath>");

        Assert.Throws<InvalidDataException>(() =>
            WordOmmlInsertionXml.ReplaceOrdinaryMarkerRun(
                markerNode,
                marker,
                equation,
                other));
    }

    [Fact]
    public void ParagraphPropertiesKeepStyleNumberingAndLayoutButNoRevisionsOrSection()
    {
        var export =
            $"<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"><pkg:part pkg:name=\"/word/document.xml\"><pkg:xmlData>"
            + $"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:pPr>"
            + "<w:pStyle w:val=\"4\"/><w:numPr><w:ilvl w:val=\"3\"/><w:numId w:val=\"1\"/></w:numPr>"
            + "<w:spacing w:before=\"120\"/><w:ind w:left=\"420\"/>"
            + "<w:rPr><w:ins w:id=\"1\" w:author=\"a\"/><w:b/></w:rPr>"
            + "<w:sectPr/><w:pPrChange w:id=\"2\" w:author=\"a\"><w:pPr/></w:pPrChange>"
            + "</w:pPr></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
        var properties = XElement.Parse(WordOmmlInsertionXml.ParagraphProperties(export)!);
        Assert.Equal(W + "pPr", properties.Name);
        Assert.Equal(new[] { "pStyle", "numPr", "spacing", "ind", "rPr" }, properties.Elements().Select(e => e.Name.LocalName));
        Assert.Equal(new[] { "b" }, properties.Element(W + "rPr")!.Elements().Select(e => e.Name.LocalName));
    }

    [Fact]
    public void ParagraphStyleDefinitionIsTheStyleThePropertiesName()
    {
        var export =
            "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">"
            + $"<pkg:part pkg:name=\"/word/document.xml\"><pkg:xmlData><w:document xmlns:w=\"{W}\"><w:body><w:p><w:pPr>"
            + "<w:pStyle w:val=\"a3\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>"
            + "</w:pPr></w:p></w:body></w:document></pkg:xmlData></pkg:part>"
            + $"<pkg:part pkg:name=\"/word/styles.xml\"><pkg:xmlData><w:styles xmlns:w=\"{W}\">"
            + "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"a\"><w:name w:val=\"Normal\"/></w:style>"
            + "<w:style w:type=\"paragraph\" w:styleId=\"a3\"><w:name w:val=\"List Paragraph\"/><w:basedOn w:val=\"a\"/>"
            + "<w:pPr><w:ind w:firstLineChars=\"200\" w:firstLine=\"420\"/></w:pPr></w:style>"
            + "</w:styles></pkg:xmlData></pkg:part></pkg:package>";

        var style = XElement.Parse(WordOmmlInsertionXml.ParagraphStyleDefinition(export)!);

        Assert.Equal("a3", (string?)style.Attribute(W + "styleId"));
        Assert.Equal("List Paragraph", (string?)style.Element(W + "name")!.Attribute(W + "val"));
        Assert.Null(WordOmmlInsertionXml.ParagraphStyleDefinition(
            export.Replace("<w:pStyle w:val=\"a3\"/>", string.Empty)));
    }

    [Fact]
    public void SingleEquationPackageCarriesTheParagraphStyleInAStylesPart()
    {
        var entry = new WordOmmlBatchSources.BatchEntry(
            "VisualTeXDirect0123456789abcdef", "fingerprint",
            $"<m:oMath xmlns:m=\"{M}\"><m:r><m:t>x</m:t></m:r></m:oMath>", 31902);
        var style = $"<w:style xmlns:w=\"{W}\" w:type=\"paragraph\" w:styleId=\"a3\"><w:name w:val=\"List Paragraph\"/></w:style>";
        var properties = $"<w:pPr xmlns:w=\"{W}\"><w:pStyle w:val=\"a3\"/></w:pPr>";
        XNamespace package = "http://schemas.microsoft.com/office/2006/xmlPackage";
        XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";

        var withStyle = XDocument.Parse(WordOmmlDocuments.CreateSingleEquationFlatOpc(
            entry, "Cambria Math", properties, paragraphStyleDefinition: style));
        var stylesPart = Assert.Single(withStyle.Descendants(package + "part"),
            part => (string?)part.Attribute(package + "name") == "/word/styles.xml");
        Assert.Equal("a3", (string?)stylesPart.Descendants(W + "style").Single().Attribute(W + "styleId"));
        Assert.Contains(withStyle.Descendants(relationships + "Relationship"),
            relationship => ((string?)relationship.Attribute("Type"))!.EndsWith("/styles", StringComparison.Ordinal)
                && (string?)relationship.Attribute("Target") == "styles.xml");

        var plain = XDocument.Parse(WordOmmlDocuments.CreateSingleEquationFlatOpc(entry, "Cambria Math", properties));
        Assert.DoesNotContain(plain.Descendants(package + "part"),
            part => (string?)part.Attribute(package + "name") == "/word/styles.xml");
    }

    [Fact]
    public void OtherContentOfTheMarkerRunIsKeptOnItsSide()
    {
        const string marker = "VTXmarker";
        var paragraph = XElement.Parse(
            $"<w:p xmlns:w=\"{W}\" xmlns:m=\"{M}\">"
            + "<w:r><w:rPr><w:i/></w:rPr><w:t>Text</w:t><w:tab/><w:t>a VTXmarker b</w:t><w:tab/><w:br/><w:t>end</w:t></w:r>"
            + "</w:p>");
        var markerNode = Assert.Single(
            paragraph.Descendants(W + "t"),
            node => node.Value.Contains(marker, StringComparison.Ordinal));
        var equation = XElement.Parse(
            $"<m:oMath xmlns:m=\"{M}\"><m:r><m:t>x</m:t></m:r></m:oMath>");

        WordOmmlInsertionXml.ReplaceOrdinaryMarkerRun(markerNode, marker, equation, paragraph);

        var children = paragraph.Elements().ToArray();
        Assert.Equal(3, children.Length);
        Assert.Equal(new[] { "rPr", "t", "tab", "t" }, children[0].Elements().Select(e => e.Name.LocalName));
        Assert.Equal("Texta ", string.Concat(children[0].Elements(W + "t").Select(e => e.Value)));
        Assert.Equal(M + "oMath", children[1].Name);
        Assert.Equal(new[] { "rPr", "t", "tab", "br", "t" }, children[2].Elements().Select(e => e.Name.LocalName));
        Assert.Equal(" bend", string.Concat(children[2].Elements(W + "t").Select(e => e.Value)));
    }
}
