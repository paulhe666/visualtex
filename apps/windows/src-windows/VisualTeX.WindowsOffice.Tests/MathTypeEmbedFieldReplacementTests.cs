using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeEmbedFieldReplacementTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private static XElement Paragraph(string instruction) => XElement.Parse(
        $"<w:p xmlns:w=\"{W}\" xmlns:m=\"{M}\">"
        + "<w:bookmarkStart w:id=\"0\" w:name=\"VTMT_TX_1\"/>"
        + "<w:r><w:rPr><w:position w:val=\"-6\"/></w:rPr><w:fldChar w:fldCharType=\"begin\"/></w:r>"
        + $"<w:r><w:rPr><w:position w:val=\"-6\"/></w:rPr><w:instrText xml:space=\"preserve\">{instruction}</w:instrText></w:r>"
        + "<w:r><w:rPr><w:position w:val=\"-6\"/></w:rPr><w:fldChar w:fldCharType=\"separate\"/></w:r>"
        + "<m:oMath><m:r><m:t>a</m:t></m:r></m:oMath>"
        + "<w:r><w:rPr><w:position w:val=\"-6\"/></w:rPr><w:fldChar w:fldCharType=\"end\"/></w:r>"
        + "<w:bookmarkEnd w:id=\"0\"/>"
        + "</w:p>");

    [Fact]
    public void TheEmbedFieldOfTheReplacedObjectIsRemoved()
    {
        // WPS exports a MathType object as the result of its EMBED field.
        var paragraph = Paragraph(" EMBED Equation.DSMT4 ");
        var equation = paragraph.Element(M + "oMath")!;
        Assert.True(MathTypeLocalFragment.RemoveEnclosingEmbedField(equation));
        Assert.Empty(paragraph.Descendants(W + "fldChar"));
        Assert.Empty(paragraph.Descendants(W + "instrText"));
        Assert.Equal(new[] { "bookmarkStart", "oMath", "bookmarkEnd" },
            paragraph.Elements().Select(e => e.Name.LocalName));
    }

    [Fact]
    public void AnotherFieldAroundTheEquationIsKept()
    {
        var paragraph = Paragraph(" MACROBUTTON MTPlaceRef ");
        var equation = paragraph.Element(M + "oMath")!;
        Assert.False(MathTypeLocalFragment.RemoveEnclosingEmbedField(equation));
        Assert.Equal(4, paragraph.Descendants(W + "fldChar").Count() + paragraph.Descendants(W + "instrText").Count());
    }

    [Fact]
    public void AnEquationWithoutAFieldIsUnchanged()
    {
        var paragraph = XElement.Parse($"<w:p xmlns:w=\"{W}\" xmlns:m=\"{M}\"><w:r><w:t>x</w:t></w:r><m:oMath/></w:p>");
        Assert.False(MathTypeLocalFragment.RemoveEnclosingEmbedField(paragraph.Element(M + "oMath")!));
        Assert.Equal(2, paragraph.Elements().Count());
    }
}
