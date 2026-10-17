using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using VisualTeX.PowerPointVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class PowerPointMathClipboardTests
{
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string A14 = "http://schemas.microsoft.com/office/drawing/2010/main";

    // A run with math style, Word run formatting, an on/off property and text
    // that needs escaping.
    private static readonly string Sample =
        $"<m:oMath xmlns:m=\"{M}\" xmlns:w=\"{W}\"><m:rad><m:radPr><m:degHide m:val=\"1\"/></m:radPr><m:deg/>"
        + "<m:e><m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/>"
        + "<w:color w:val=\"FF0000\"/></w:rPr><m:t>a&lt;b&amp;c</m:t></m:r></m:e></m:rad></m:oMath>";

    private static string Drawing(IEnumerable<KeyValuePair<string, string>> entries) =>
        entries.Single(entry => entry.Key == "clipboard/drawings/drawing1.xml").Value;

    // VisualTeX for macOS packs the same OMML into the same bytes
    // (src-tauri office/powerpoint_omml.rs, test windows_package_parity). Change
    // both or neither.
    [Fact]
    public void PackageMatchesTheMacPackage()
    {
        var package = PowerPointMathClipboard.Build(Sample, 24, display: true);
        using var sha = SHA256.Create();
        Assert.Equal(
            "abc0b2da908c1622ccc8c236064955abe2647ddb56c881801837d252fb66f1ec",
            BitConverter.ToString(sha.ComputeHash(package)).Replace("-", "").ToLowerInvariant());
    }

    [Fact]
    public void PackageIsAStoredZipOfTheTextDrawing()
    {
        var package = PowerPointMathClipboard.Build(Sample, 24, display: true);
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        Assert.Equal(
            new[] { "[Content_Types].xml", "_rels/.rels", "clipboard/drawings/drawing1.xml",
                "clipboard/drawings/_rels/drawing1.xml.rels", "clipboard/theme/theme1.xml" },
            archive.Entries.Select(entry => entry.FullName).ToArray());
        string drawing;
        using (var reader = new StreamReader(archive.GetEntry("clipboard/drawings/drawing1.xml")!.Open(), Encoding.UTF8))
            drawing = reader.ReadToEnd();
        Assert.Contains("<a:pPr algn=\"ctr\"><a:buNone/></a:pPr><a14:m><m:oMathPara>", drawing);
        Assert.Contains("<m:degHide m:val=\"on\"></m:degHide>", drawing);
        Assert.Contains("<m:rPr><m:sty m:val=\"p\"></m:sty></m:rPr><a:rPr sz=\"2400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>"
            + "<a:latin typeface=\"Cambria Math\"/></a:rPr><m:t>a&lt;b&amp;c</m:t>", drawing);
        Assert.Contains("<a:endParaRPr sz=\"2400\"/>", drawing);
        Assert.DoesNotContain("xmlns:w", drawing);
    }

    [Fact]
    public void InlinePackageHasNoParagraphFormatting() =>
        Assert.DoesNotContain("<a:pPr", Drawing(PowerPointMathClipboard.Entries(Sample, 24, display: false)));

    [Theory]
    [InlineData("1", "on")]
    [InlineData("true", "on")]
    [InlineData("0", "off")]
    [InlineData("false", "off")]
    [InlineData("on", "on")]
    public void OnOffPropertiesUsePowerPointsCanonicalValues(string value, string expected)
    {
        var omml = $"<m:oMath xmlns:m=\"{M}\"><m:nary><m:naryPr><m:supHide m:val=\"{value}\"/></m:naryPr><m:sub/><m:sup/>"
            + "<m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary></m:oMath>";
        Assert.Contains($"<m:supHide m:val=\"{expected}\">", Drawing(PowerPointMathClipboard.Entries(omml, 18, display: false)));
    }

    [Fact]
    public void EditKeepsTheZoneAroundTheNewEquation()
    {
        var zone = $"<a14:m xmlns:a14=\"{A14}\"><m:oMathPara xmlns:m=\"{M}\"><m:oMathParaPr><m:jc m:val=\"left\"/></m:oMathParaPr>"
            + "<m:oMath><m:r><m:t>old</m:t></m:r></m:oMath></m:oMathPara></a14:m>";
        var drawing = Drawing(PowerPointMathClipboard.EditEntries(Sample, 18, zone));
        var math = XElement.Parse(drawing).Descendants(XName.Get("m", A14)).Single();
        Assert.Equal("left", (string?)math.Descendants(XName.Get("jc", M)).Single().Attribute(XName.Get("val", M)));
        Assert.Equal("a<b&c", string.Concat(math.Descendants(XName.Get("t", M)).Select(text => text.Value)));
        Assert.DoesNotContain("old", drawing);
        Assert.Contains("sz=\"1800\"", drawing);
        Assert.DoesNotContain("<a:pPr", drawing);
    }

    [Fact]
    public void EditRejectsAZoneItCannotReplaceExactly()
    {
        var two = $"<a14:m xmlns:a14=\"{A14}\"><m:oMathPara xmlns:m=\"{M}\"><m:oMath><m:r><m:t>a</m:t></m:r></m:oMath>"
            + "<m:oMath><m:r><m:t>b</m:t></m:r></m:oMath></m:oMathPara></a14:m>";
        Assert.Throws<InvalidDataException>(() => PowerPointMathClipboard.EditEntries(Sample, 18, two));
        Assert.Throws<InvalidDataException>(() => PowerPointMathClipboard.EditEntries(Sample, 18, $"<m:oMath xmlns:m=\"{M}\"/>"));
        Assert.Throws<InvalidDataException>(() => PowerPointMathClipboard.EditEntries(Sample, 18, "<a14:m"));
    }

    [Fact]
    public void RejectsSizesPowerPointCannotTake()
    {
        Assert.Throws<InvalidDataException>(() => PowerPointMathClipboard.Build(Sample, 0.5, display: false));
        Assert.Throws<InvalidDataException>(() => PowerPointMathClipboard.Build(Sample, double.NaN, display: false));
    }
}
