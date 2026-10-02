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
}
