using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeRecoveryOwnershipTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static string Live(string firstClass) =>
        $"<geometry><object start=\"1\" end=\"23\" type=\"1\" progId=\"{firstClass}\" width=\"64\" height=\"19.5\" />"
        + "<object start=\"28\" end=\"54\" type=\"1\" progId=\"Equation.DSMT4\" width=\"32.5\" height=\"14\" /></geometry>";

    private static XDocument Package(string firstClass = "Equation.3", bool stale = true)
    {
        XElement Object(string id, string cls, string style, string originalWidth, string originalHeight) =>
            new(W + "r", new XElement(W + "rPr", new XElement(W + "position", new XAttribute(W + "val", "-10"))),
                new XElement(W + "object", new XAttribute(W + "dxaOrig", originalWidth), new XAttribute(W + "dyaOrig", originalHeight),
                    new XElement(V + "shape", new XAttribute("id", id), new XAttribute("style", style),
                        new XElement(V + "imagedata", new XAttribute(R + "id", id + "Preview"))),
                    new XElement(O + "OLEObject", new XAttribute("Type", "Embed"), new XAttribute("ProgID", cls),
                        new XAttribute("ShapeID", id), new XAttribute(R + "id", id + "Payload"))));
        var body = new XElement(W + "body",
            new XElement(W + "p", new XElement(W + "pPr", new XElement(W + "rPr",
                new XElement(W + "del", new XAttribute(W + "id", "1"), new XAttribute(W + "author", "Original author")))),
                new XElement(W + "del", new XAttribute(W + "id", "2"), new XAttribute(W + "author", "Original author"),
                    new XAttribute(W + "date", "2026-01-02T03:04:05Z"),
                    Object("legacy", firstClass, stale ? "width:61.7pt;height:20.55pt" : "width:64pt;height:19.5pt", "1283", "394"))),
            new XElement(W + "p", new XElement(W + "pPr", new XElement(W + "jc", new XAttribute(W + "val", "both"))),
                new XElement(W + "r", new XElement(W + "t", "prefix")),
                new XElement(W + "bookmarkStart", new XAttribute(W + "id", "7"), new XAttribute(W + "name", "_Keep")),
                Object("current", "Equation.DSMT4", "width:32.5pt;height:14pt", "660", "279"),
                new XElement(W + "bookmarkEnd", new XAttribute(W + "id", "7")),
                new XElement(W + "r", new XElement(W + "t", "suffix"))));
        return new XDocument(new XElement(Pkg + "package",
            new XElement(Pkg + "part", new XAttribute(Pkg + "name", "/word/document.xml"),
                new XElement(Pkg + "xmlData", new XElement(W + "document", body))),
            new XElement(Pkg + "part", new XAttribute(Pkg + "name", "/word/embeddings/legacy.bin"), new XElement(Pkg + "binaryData", "AAEC")),
            new XElement(Pkg + "part", new XAttribute(Pkg + "name", "/word/media/legacy.wmf"), new XElement(Pkg + "binaryData", "AwQF"))));
    }

    private static string Evidence(XDocument package, string? live = null) =>
        MathTypeLocalFragment.RecoveryEvidence(package.ToString(SaveOptions.DisableFormatting), live ?? Live("Equation.3"), 2);

    [Theory]
    [InlineData("Equation.3")]
    [InlineData("Equation.DSMT4")]
    [InlineData("Visio.Drawing.11")]
    [InlineData("Other.Server")]
    public void WordOwnedPreviewCacheIsIndependentOfEveryLocalOleServerClass(string firstClass)
    {
        var original = Package(firstClass);
        var untouched = original.ToString();
        Assert.Equal(Evidence(original, Live(firstClass)), Evidence(Package(firstClass, stale: false), Live(firstClass)));
        Assert.Equal(untouched, original.ToString()); // Evidence normalization never changes the source tree.
    }

    [Theory]
    [InlineData("width=\"64\"", "width=\"64.05\"")]
    [InlineData("height=\"19.5\"", "height=\"19.55\"")]
    [InlineData("start=\"1\"", "start=\"2\"")]
    [InlineData("end=\"54\"", "end=\"55\"")]
    [InlineData("type=\"1\"", "type=\"2\"")]
    [InlineData("Equation.3", "Equation.DSMT4")]
    public void RealHostGeometrySpanAndClassChangesRemainDetectable(string from, string to)
    {
        Assert.NotEqual(Evidence(Package()), Evidence(Package(stale: false), Live("Equation.3").Replace(from, to)));
    }

    [Theory]
    [InlineData("payload")]
    [InlineData("preview")]
    [InlineData("intrinsicWidth")]
    [InlineData("intrinsicHeight")]
    [InlineData("position")]
    [InlineData("rotation")]
    [InlineData("author")]
    [InlineData("date")]
    [InlineData("paragraphMark")]
    [InlineData("bookmark")]
    [InlineData("text")]
    public void HistoricalContentAndProvenanceAreNotNormalizedAway(string change)
    {
        var changed = Package(stale: false);
        switch (change)
        {
            case "payload": changed.Descendants(Pkg + "binaryData").First().Value = "AQEC"; break;
            case "preview": changed.Descendants(Pkg + "binaryData").Last().Value = "BAQF"; break;
            case "intrinsicWidth": changed.Descendants(W + "object").First().SetAttributeValue(W + "dxaOrig", "1284"); break;
            case "intrinsicHeight": changed.Descendants(W + "object").First().SetAttributeValue(W + "dyaOrig", "395"); break;
            case "position": changed.Descendants(W + "position").First().SetAttributeValue(W + "val", "-9"); break;
            case "rotation": changed.Descendants(V + "shape").First().Attribute("style")!.Value += ";rotation:5"; break;
            case "author": changed.Descendants(W + "del").Last().SetAttributeValue(W + "author", "Changed author"); break;
            case "date": changed.Descendants(W + "del").Last().SetAttributeValue(W + "date", "2026-01-03T03:04:05Z"); break;
            case "paragraphMark": changed.Descendants(W + "del").First().Remove(); break;
            case "bookmark": changed.Descendants(W + "bookmarkStart").Single().SetAttributeValue(W + "name", "_Different"); break;
            case "text": changed.Descendants(W + "t").Last().Value = "changed suffix"; break;
            default: throw new ArgumentOutOfRangeException(nameof(change));
        }
        Assert.NotEqual(Evidence(Package()), Evidence(changed));
    }

    [Fact]
    public void MissingOwnerOrAmbiguousGeometryFailsClosed()
    {
        var missing = Package();
        missing.Descendants(O + "OLEObject").First().Remove();
        Assert.Throws<InvalidDataException>(() => Evidence(missing));
        var duplicated = Package();
        duplicated.Descendants(V + "shape").First().SetAttributeValue("style", "width:64pt;height:19.5pt;width:64pt");
        Assert.Throws<InvalidDataException>(() => Evidence(duplicated));
    }
}
