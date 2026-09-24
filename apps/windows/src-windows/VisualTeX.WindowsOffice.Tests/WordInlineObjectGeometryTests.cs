using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordInlineObjectGeometryTests
{
    private static string Evidence(string width, float actualWidth = 110.9f, string otherStyle = "", string progId = "Equation.DSMT4")
    {
        var body = XElement.Parse("<w:body xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
            + " xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:v=\"urn:schemas-microsoft-com:vml\">"
            + "<w:p><w:r><w:object><v:shape id=\"s1\" style=\"width:" + width + "pt;height:68pt" + otherStyle
            + "\"/><o:OLEObject Type=\"Embed\" ProgID=\"" + progId + "\" ShapeID=\"s1\"/></w:object></w:r></w:p></w:body>");
        WordInlineObjectGeometry.BindToEvidence(body, new[] {
            new WordInlineObjectGeometry.Item { Start=2, End=28, Type=1, ProgId="Equation.DSMT4", Width=actualWidth, Height=68.05f } });
        return body.ToString(SaveOptions.DisableFormatting);
    }

    [Fact]
    public void RoundedPreviewAndActualWordBoundsRequireTheSameLiveGeometry()
        => Assert.Equal(Evidence("111"), Evidence("110.9"));

    [Fact]
    public void SubpointResizeIsNotHiddenByIntegerPreviewBounds()
        => Assert.NotEqual(Evidence("111"), Evidence("111", actualWidth:110.95f));

    [Theory]
    [InlineData("110.8")]
    [InlineData("112")]
    public void OtherXmlDimensionsRemainDistinct(string changed)
        => Assert.NotEqual(Evidence("111"), Evidence(changed));

    [Fact]
    public void OtherLayoutPropertiesRemainInTheEvidence()
        => Assert.NotEqual(Evidence("111"), Evidence("111", otherStyle:";rotation:5"));

    [Fact]
    public void DifferentPhysicalObjectTypeCannotClaimGeometry()
        => Assert.Throws<InvalidDataException>(() => Evidence("111", progId:"VisualTeX.Formula.1"));

    [Fact]
    public void UnmatchedLiveOleObjectCannotBeIgnored()
    {
        var body = XElement.Parse("<w:body xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>");
        Assert.Throws<InvalidDataException>(() => WordInlineObjectGeometry.BindToEvidence(body,
            new[] {new WordInlineObjectGeometry.Item {Type=1, ProgId="Equation.DSMT4",Width=10,Height=10}}));
    }
}
