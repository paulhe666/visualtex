using System.Xml.Linq;
using VisualTeX.WindowsOffice.VstoShared;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WpsRibbonImagesTests
{
    private const string Ribbon =
        "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><ribbon><tabs><tab id=\"VisualTeX\" label=\"VisualTeX\">"
        + "<group id=\"g\" label=\"g\">"
        + "<button id=\"VisualTeX.WordVsto.Inline\" label=\"OLE\" size=\"large\" tag=\"oleInline\" getImage=\"GetRibbonImage\" onAction=\"OnInsertInline\" />"
        + "<button id=\"VisualTeX.WordVsto.FontSizeIncrease\" label=\"+\" imageMso=\"FontSizeIncrease\" onAction=\"OnIncrease\" />"
        + "<button id=\"VisualTeX.WordVsto.InsertReference\" label=\"ref\" imageMso=\"HyperlinkInsert\" onAction=\"OnRef\" />"
        + "<button id=\"Other\" label=\"x\" getImage=\"GetRibbonImage\" />"
        + "</group></tab></tabs></ribbon></customUI>";

    private static XElement Control(XDocument document, string id) =>
        document.Descendants().Single(element => (string?)element.Attribute("id") == id);

    [Fact]
    public void ControlsGetIconsWpsKnows()
    {
        var document = XDocument.Parse(WpsRibbonImages.UseWpsIcons(Ribbon));
        var inline = Control(document, "VisualTeX.WordVsto.Inline");
        Assert.Null(inline.Attribute("getImage"));
        Assert.Equal("EquationScriptGallery", (string?)inline.Attribute("imageMso"));
        Assert.Equal("OnInsertInline", (string?)inline.Attribute("onAction"));
        // An Office id WPS lacks is replaced too; one it knows is kept.
        Assert.Equal("GrowFont", (string?)Control(document, "VisualTeX.WordVsto.FontSizeIncrease").Attribute("imageMso"));
        Assert.Equal("HyperlinkInsert", (string?)Control(document, "VisualTeX.WordVsto.InsertReference").Attribute("imageMso"));
        // Controls without a WPS icon keep their own picture.
        Assert.Equal("GetRibbonImage", (string?)Control(document, "Other").Attribute("getImage"));
    }

    [Fact]
    public void InvalidXmlIsReturnedUnchanged() =>
        Assert.Equal("<customUI", WpsRibbonImages.UseWpsIcons("<customUI"));

    [Fact]
    public void OutsideWpsTheRibbonIsUnchanged() =>
        Assert.Same(Ribbon, WpsRibbonImages.ForHost(Ribbon));
}
