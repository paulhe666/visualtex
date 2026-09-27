using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.PowerPoint;
using VisualTeX.WordVsto;

namespace VisualTeX.PowerPointVsto;

// Read the embedded storage without activating MathType or taking the clipboard.
// Slide IDs and shape IDs are scoped explicitly; shape names are not unique
// across slides and must never be used to locate a third-party equation.
internal static class PowerPointMathTypePackage
{
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";

    internal sealed class Equation
    {
        internal byte[] CompoundFile { get; set; } = Array.Empty<byte>();
        internal float WidthPt { get; set; }
        internal float HeightPt { get; set; }
    }

    internal static Equation Read(Presentation presentation, int slideId, int shapeId)
    {
        var path = Path.Combine(Path.GetTempPath(), $"visualtex-ppt-mathtype-{Guid.NewGuid():N}.pptx");
        try
        {
            presentation.SaveCopyAs(path, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            using var file = File.OpenRead(path);
            using var archive = new ZipArchive(file, ZipArchiveMode.Read);
            return Read(archive, slideId, shapeId);
        }
        finally { try { File.Delete(path); } catch { } }
    }

    internal static Equation Read(ZipArchive archive, int slideId, int shapeId)
    {
        const string presentationPart = "ppt/presentation.xml";
        var slide = ReadXml(archive, presentationPart).Descendants(P + "sldId")
            .SingleOrDefault(element => (int?)element.Attribute("id") == slideId)
            ?? throw new InvalidDataException("The MathType slide no longer exists.");
        var slidePart = ResolveRelationship(archive, presentationPart, (string?)slide.Attribute(R + "id"));
        var shape = ReadXml(archive, slidePart).Descendants(P + "cNvPr")
            .Where(element => (int?)element.Attribute("id") == shapeId)
            .Select(element => element.Ancestors(P + "graphicFrame").FirstOrDefault())
            .FirstOrDefault(element => element is not null)
            ?? throw new InvalidDataException("The selected shape is not an embedded MathType object.");
        var ole = shape.Descendants(P + "oleObj").FirstOrDefault()
            ?? throw new InvalidDataException("The MathType object has no OLE storage.");
        if (ole.Element(P + "link") is not null)
            throw new InvalidOperationException("Linked MathType objects cannot be edited as embedded equations.");
        var storagePart = ResolveRelationship(archive, slidePart, (string?)ole.Attribute(R + "id"));
        var entry = archive.GetEntry(storagePart)
            ?? throw new InvalidDataException("The embedded MathType storage is missing.");
        if (entry.Length <= 0 || entry.Length > MathTypeOleStorage.MaximumCompoundFileBytes)
            throw new InvalidDataException("The embedded MathType storage has an invalid size.");
        using var stream = entry.Open();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        var compound = buffer.ToArray();
        if (!MathTypeOleStorage.LooksLikeMathTypeCompoundFile(compound))
            throw new InvalidDataException("The selected object is not a supported MathType equation.");
        return new Equation
        {
            CompoundFile = compound,
            WidthPt = (float)((double?)ole.Attribute("imgW") ?? 0) / 12700f,
            HeightPt = (float)((double?)ole.Attribute("imgH") ?? 0) / 12700f,
        };
    }

    private static string ResolveRelationship(ZipArchive archive, string owner, string? id)
    {
        var slash = owner.LastIndexOf('/');
        var rels = owner.Substring(0, slash + 1) + "_rels/" + owner.Substring(slash + 1) + ".rels";
        var relationship = ReadXml(archive, rels).Descendants(Rel + "Relationship")
            .SingleOrDefault(element => (string?)element.Attribute("Id") == id)
            ?? throw new InvalidDataException("The MathType package relationship is missing.");
        if (string.Equals((string?)relationship.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("External MathType relationships are not embedded equations.");
        var target = (string?)relationship.Attribute("Target")
            ?? throw new InvalidDataException("The MathType relationship has no target.");
        var uri = new Uri(new Uri("https://package.invalid/" + owner), target);
        if (uri.Host != "package.invalid") throw new InvalidDataException("Invalid package target.");
        return Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/'));
    }

    private static XDocument ReadXml(ZipArchive archive, string part)
    {
        var entry = archive.GetEntry(part) ?? throw new InvalidDataException($"Missing PPTX part: {part}");
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = 32 * 1024 * 1024,
        });
        return XDocument.Load(reader);
    }
}
