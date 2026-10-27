using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using VisualTeX.PowerPointVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// How WPS Presentation spells native equations in its text (measured in WPS 12.1 with
// tools/wpp-model-probe.ps1): structure start U+001A, each argument closed by U+001B, the
// math characters as WPS saves them in m:t.
public sealed class WpsMathTextTests
{
    private const string P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string A14 = "http://schemas.microsoft.com/office/drawing/2010/main";
    private static readonly string S = ((char)0x1A).ToString();
    private static readonly string E = ((char)0x1B).ToString();
    private static readonly string X = char.ConvertFromUtf32(0x1D465); // math italic x
    private static readonly string Y = char.ConvertFromUtf32(0x1D466); // math italic y
    private const string Rpr = "<a:rPr sz=\"2400\"><a:latin typeface=\"Cambria Math\"/></a:rPr>";

    private static string R(string text) => $"<m:r>{Rpr}<m:t>{text}</m:t></m:r>";
    private static string Ctrl(string name) => $"<m:{name}Pr><m:ctrlPr>{Rpr}</m:ctrlPr></m:{name}Pr>";
    private static XElement Zone(string math) =>
        XElement.Parse($"<a14:m xmlns:a14=\"{A14}\" xmlns:a=\"{A}\"><m:oMath xmlns:m=\"{M}\">{math}</m:oMath></a14:m>");

    [Fact]
    public void SpellsRunsAsSaved() =>
        Assert.Equal(X + "=1", WpsMathText.Spell(Zone(R(X) + R("=") + R("1"))));

    [Fact]
    public void SpellsEveryStructureWithItsArgumentsInXmlOrder()
    {
        // x^2: start, base, end of base, superscript, end
        Assert.Equal(S + X + E + "2" + E,
            WpsMathText.Spell(Zone($"<m:sSup>{Ctrl("sSup")}<m:e>{R(X)}</m:e><m:sup>{R("2")}</m:sup></m:sSup>")));
        // a hidden radical degree still is an (empty) argument
        Assert.Equal(S + E + X + E,
            WpsMathText.Spell(Zone($"<m:rad><m:radPr><m:degHide m:val=\"on\"/></m:radPr><m:deg/><m:e>{R(X)}</m:e></m:rad>")));
        // n-ary: sub, sup, base
        Assert.Equal(S + "i=1" + E + "n" + E + X + E,
            WpsMathText.Spell(Zone($"<m:nary><m:naryPr><m:chr m:val=\"S\"/></m:naryPr><m:sub>{R("i=1")}</m:sub><m:sup>{R("n")}</m:sup><m:e>{R(X)}</m:e></m:nary>")));
        // a 2x2 matrix: one argument per cell, rows add nothing
        Assert.Equal(S + "a" + E + "b" + E + "c" + E + "d" + E,
            WpsMathText.Spell(Zone($"<m:m><m:mr><m:e>{R("a")}</m:e><m:e>{R("b")}</m:e></m:mr><m:mr><m:e>{R("c")}</m:e><m:e>{R("d")}</m:e></m:mr></m:m>")));
        // delimiters and accents: their characters live in properties, not in the text
        Assert.Equal(S + X + E + Y + E,
            WpsMathText.Spell(Zone($"<m:d><m:dPr><m:begChr m:val=\"[\"/></m:dPr><m:e>{R(X)}</m:e><m:e>{R(Y)}</m:e></m:d>")));
        Assert.Equal(S + X + E,
            WpsMathText.Spell(Zone($"<m:acc><m:accPr><m:chr m:val=\"^\"/></m:accPr><m:e>{R(X)}</m:e></m:acc>")));
        // a function: name, then argument; nested fraction inside a pre-script
        Assert.Equal(S + "sin" + E + X + E,
            WpsMathText.Spell(Zone($"<m:func><m:fName>{R("sin")}</m:fName><m:e>{R(X)}</m:e></m:func>")));
        Assert.Equal(S + "1" + E + "2" + E + S + X + E + Y + E + E,
            WpsMathText.Spell(Zone($"<m:sPre><m:sub>{R("1")}</m:sub><m:sup>{R("2")}</m:sup><m:e><m:f><m:num>{R(X)}</m:num><m:den>{R(Y)}</m:den></m:f></m:e></m:sPre>")));
    }

    private static XElement Shape(string paragraphs) =>
        XElement.Parse($"<p:sp xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:a14=\"{A14}\" xmlns:m=\"{M}\"><p:txBody><a:bodyPr/>{paragraphs}</p:txBody></p:sp>");

    private static string Math(string math) => $"<a14:m><m:oMath>{math}</m:oMath></a14:m>";

    [Fact]
    public void LocatesEquationsBetweenTextRunsBreaksAndParagraphs()
    {
        var shape = Shape(
            $"<a:p><a:r><a:t>Let </a:t></a:r>{Math(R(X))}<a:r><a:t> and </a:t></a:r>{Math($"<m:sSup><m:e>{R(Y)}</m:e><m:sup>{R("2")}</m:sup></m:sSup>")}</a:p>"
            + "<a:p><a:r><a:t>line1</a:t></a:r><a:br/><a:r><a:t>line2</a:t></a:r></a:p>"
            + $"<a:p><a:pPr algn=\"ctr\"/><a14:m><m:oMathPara><m:oMathParaPr><m:jc m:val=\"centerGroup\"/></m:oMathParaPr><m:oMath>{R("1")}</m:oMath></m:oMathPara></a14:m></a:p>");
        var live = "Let " + X + " and " + S + Y + E + "2" + E + "\rline1\vline2\r1";
        var zones = WpsMathText.Locate(shape, live);
        Assert.NotNull(zones);
        Assert.Equal(new[] { (5, 2), (12, 6), (31, 1) }, zones!.ToArray());
    }

    [Fact]
    public void SeparatesAdjacentEquations()
    {
        var shape = Shape($"<a:p>{Math(R(X))}{Math(R(Y))}</a:p>");
        Assert.Equal(new[] { (1, 2), (3, 2) }, WpsMathText.Locate(shape, X + Y)!.ToArray());
    }

    // WPS saves the break before a display equation as a CR inside the preceding run;
    // serialized XML reads it back as a line feed.
    [Fact]
    public void ReadsTheCrWpsKeepsInARunBeforeADisplayEquation()
    {
        var shape = Shape($"<a:p><a:r><a:t>Before.&#13;</a:t></a:r>{Math(R(X))}</a:p>");
        var reparsed = XElement.Parse(shape.ToString(SaveOptions.DisableFormatting));
        Assert.Equal(new[] { (9, 2) }, WpsMathText.Locate(reparsed, "Before.\r" + X)!.ToArray());
    }

    [Fact]
    public void TrustsNoPositionWhenTheTextDiffers()
    {
        var shape = Shape($"<a:p><a:r><a:t>Let </a:t></a:r>{Math(R(X))}</a:p>");
        Assert.Null(WpsMathText.Locate(shape, "Let x"));
        Assert.Null(WpsMathText.Locate(shape, "Let " + X + " more"));
    }

    [Fact]
    public void FindsTheOneEquationAtACaretOrSelection()
    {
        var zones = new[] { (5, 2), (13, 6) };
        Assert.Equal(0, WpsMathText.ZoneAt(zones, 5, 0));
        Assert.Equal(0, WpsMathText.ZoneAt(zones, 6, 0));
        Assert.Equal(-1, WpsMathText.ZoneAt(zones, 7, 0));
        Assert.Equal(1, WpsMathText.ZoneAt(zones, 13, 6));
        Assert.Equal(-1, WpsMathText.ZoneAt(zones, 4, 12));
    }

    [Fact]
    public void FiltersTextWithoutMath()
    {
        Assert.False(WpsMathText.MayContainMath("plain text 2+3"));
        Assert.True(WpsMathText.MayContainMath("Let " + X));
        Assert.True(WpsMathText.MayContainMath(S + "a" + E));
    }

    [Fact]
    public void EquationSlideIsAOneSlidePresentationWithTheParagraphs()
    {
        var paragraphs = PowerPointMathClipboard.Paragraphs(
            PowerPointMathClipboard.Entries($"<m:oMath xmlns:m=\"{M}\"><m:r><m:t>x</m:t></m:r></m:oMath>", 20, display: true));
        Assert.StartsWith("<a:p ", paragraphs);
        Assert.Single(XElement.Parse(paragraphs).Descendants(XName.Get("m", A14)));
        using var archive = new ZipArchive(new MemoryStream(WpsEquationSlide.Build(paragraphs)), ZipArchiveMode.Read);
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
            return XDocument.Parse(reader.ReadToEnd());
        });
        Assert.Equal(XName.Get("theme", A), parts["ppt/theme/theme1.xml"].Root!.Name);
        var slide = parts["ppt/slides/slide1.xml"];
        var box = slide.Descendants(XName.Get("sp", P)).Single();
        Assert.Equal("a14", (string?)box.Parent!.Attribute("Requires"));
        var paragraph = box.Descendants(XName.Get("p", A)).Single();
        Assert.Equal("ctr", (string?)paragraph.Element(XName.Get("pPr", A))!.Attribute("algn"));
        Assert.Equal("x", paragraph.Descendants(XName.Get("t", M)).Single().Value);
    }

    [Fact]
    public void RunsCarryTheTextLanguageWhenGiven()
    {
        var omml = $"<m:oMath xmlns:m=\"{M}\"><m:r><m:t>a</m:t></m:r><m:r><m:t>&gt;</m:t></m:r></m:oMath>";
        var drawing = PowerPointMathClipboard.Entries(omml, 18, display: false, "en-US")
            .Single(entry => entry.Key == "clipboard/drawings/drawing1.xml").Value;
        var runs = XElement.Parse(drawing).Descendants(XName.Get("r", M))
            .Select(run => run.Element(XName.Get("rPr", A))!).ToArray();
        Assert.Equal(2, runs.Length);
        Assert.All(runs, properties =>
        {
            Assert.Equal("en-US", (string?)properties.Attribute("lang"));
            Assert.Equal("en-US", (string?)properties.Attribute("altLang"));
        });
        Assert.DoesNotContain("lang=", PowerPointMathClipboard.Entries(omml, 18, display: false)
            .Single(entry => entry.Key == "clipboard/drawings/drawing1.xml").Value);
        Assert.Throws<InvalidDataException>(() => PowerPointMathClipboard.Entries(omml, 18, display: false, "en\"US"));
    }
}
