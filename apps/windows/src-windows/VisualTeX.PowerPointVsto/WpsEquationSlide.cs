namespace VisualTeX.PowerPointVsto;

/// <summary>
/// A one-slide presentation whose only text box holds the paragraphs of a
/// native equation package (PowerPointMathClipboard.Paragraphs). WPS
/// Presentation pastes nothing from PowerPoint's equation clipboard format, but
/// Slides.InsertFromFile imports a14:m equations from a presentation file as
/// native equations, and its own copy of that text pastes them anywhere.
/// </summary>
internal static class WpsEquationSlide
{
    private const string P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string A14 = "http://schemas.microsoft.com/office/drawing/2010/main";
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string Declaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    private const string ColorMap = "<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>";
    private const string EmptyTree = "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>";

    /// <summary>The .pptx bytes (a stored ZIP) holding <paramref name="paragraphs"/>.</summary>
    internal static byte[] Build(string paragraphs) =>
        PowerPointMathClipboard.BuildStoredZip(Entries(paragraphs));

    internal static List<KeyValuePair<string, string>> Entries(string paragraphs)
    {
        const string slideType = "application/vnd.openxmlformats-officedocument.presentationml";
        string Relationships(params (string Id, string Type, string Target)[] items) =>
            $"{Declaration}<Relationships xmlns=\"{Rel}\">"
            + string.Concat(items.Select(item => $"<Relationship Id=\"{item.Id}\" Type=\"{R}/{item.Type}\" Target=\"{item.Target}\"/>"))
            + "</Relationships>";
        // The math shape is written the way PowerPoint saves one: under an a14
        // choice. The box only carries the text; WPS takes its own layout.
        var slide =
            $"{Declaration}<p:sld xmlns:a=\"{A}\" xmlns:r=\"{R}\" xmlns:p=\"{P}\" xmlns:a14=\"{A14}\" xmlns:m=\"{M}\"><p:cSld><p:spTree>{EmptyTree}"
            + $"<mc:AlternateContent xmlns:mc=\"{Mc}\"><mc:Choice Requires=\"a14\"><p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"VisualTeX Equation\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>"
            + "<p:spPr><a:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"9144000\" cy=\"914400\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>"
            + $"<p:txBody><a:bodyPr wrap=\"none\"><a:spAutoFit/></a:bodyPr><a:lstStyle/>{paragraphs}</p:txBody></p:sp></mc:Choice></mc:AlternateContent>"
            + "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>";
        return new List<KeyValuePair<string, string>>
        {
            new("[Content_Types].xml",
                $"{Declaration}<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + $"<Override PartName=\"/ppt/presentation.xml\" ContentType=\"{slideType}.presentation.main+xml\"/>"
                + $"<Override PartName=\"/ppt/slides/slide1.xml\" ContentType=\"{slideType}.slide+xml\"/>"
                + $"<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"{slideType}.slideLayout+xml\"/>"
                + $"<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"{slideType}.slideMaster+xml\"/>"
                + "<Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/></Types>"),
            new("_rels/.rels", Relationships(("rId1", "officeDocument", "ppt/presentation.xml"))),
            new("ppt/presentation.xml",
                $"{Declaration}<p:presentation xmlns:a=\"{A}\" xmlns:r=\"{R}\" xmlns:p=\"{P}\"><p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id=\"256\" r:id=\"rId2\"/></p:sldIdLst><p:sldSz cx=\"12192000\" cy=\"6858000\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/></p:presentation>"),
            new("ppt/_rels/presentation.xml.rels", Relationships(
                ("rId1", "slideMaster", "slideMasters/slideMaster1.xml"),
                ("rId2", "slide", "slides/slide1.xml"),
                ("rId3", "theme", "theme/theme1.xml"))),
            new("ppt/slideMasters/slideMaster1.xml",
                $"{Declaration}<p:sldMaster xmlns:a=\"{A}\" xmlns:r=\"{R}\" xmlns:p=\"{P}\"><p:cSld><p:spTree>{EmptyTree}</p:spTree></p:cSld>{ColorMap}<p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rId1\"/></p:sldLayoutIdLst></p:sldMaster>"),
            new("ppt/slideMasters/_rels/slideMaster1.xml.rels", Relationships(
                ("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"),
                ("rId2", "theme", "../theme/theme1.xml"))),
            new("ppt/slideLayouts/slideLayout1.xml",
                $"{Declaration}<p:sldLayout xmlns:a=\"{A}\" xmlns:r=\"{R}\" xmlns:p=\"{P}\" type=\"blank\"><p:cSld name=\"Blank\"><p:spTree>{EmptyTree}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"),
            new("ppt/slideLayouts/_rels/slideLayout1.xml.rels", Relationships(("rId1", "slideMaster", "../slideMasters/slideMaster1.xml"))),
            new("ppt/slides/slide1.xml", slide),
            new("ppt/slides/_rels/slide1.xml.rels", Relationships(("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"))),
            new("ppt/theme/theme1.xml", Declaration + PowerPointMathClipboard.PresentationTheme()),
        };
    }
}
