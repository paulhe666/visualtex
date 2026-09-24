using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordDocumentXmlEmptyExportTests
{
    private static XElement Body(string contents) => XElement.Parse(
        "<w:body xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>" + contents + "</w:body>");

    [Fact]
    public void ZeroFormulaCountsCannotAuthorizeAnEmptyLaTexSourceCheckpoint()
    {
        Assert.True(WordDocumentXml.IsUnexpectedlyEmptyBody(Body("<w:p/>"), "$x=1$\r"));
        Assert.True(WordDocumentXml.IsUnexpectedlyEmptyBody(Body("<w:p/><w:sectPr/>"), "正文与公式\r"));
    }

    [Fact]
    public void APristineBlankDocumentIsNotMistakenForATruncatedExport()
    {
        Assert.False(WordDocumentXml.IsUnexpectedlyEmptyBody(Body("<w:p/>"), "\r"));
        Assert.False(WordDocumentXml.IsUnexpectedlyEmptyBody(Body("<w:p/>"), "\r\t \u200C"));
    }

    [Fact]
    public void PresentTextDoesNotTriggerTheEmptyBodyRetry()
    {
        Assert.False(WordDocumentXml.IsUnexpectedlyEmptyBody(
            Body("<w:p><w:r><w:t>正文</w:t></w:r></w:p>"), "正文\r"));
    }
}
