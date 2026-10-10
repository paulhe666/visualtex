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
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string OleObjectRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
    private const string ImageRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    internal const string StagingPlaceholderName = "VisualTeX.MathType.Staging";

    internal sealed class Equation
    {
        internal byte[] CompoundFile { get; set; } = Array.Empty<byte>();
        internal float WidthPt { get; set; }
        internal float HeightPt { get; set; }
    }

    // Saving a copy is cheap even for large decks (116 MB of pictures: ~0.1 s):
    // PowerPoint copies the stored parts. It also leaves the clipboard alone.
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

    // Replace the staging placeholder on the package's first slide with an
    // embedded Equation.DSMT4 object and its preview picture. This is the same
    // markup PowerPoint saves for a MathType equation; loading it needs no OLE
    // server, so it works whether or not MathType is installed.
    internal static void WriteEquation(
        string packagePath,
        byte[] compoundFile,
        byte[] previewEmf,
        float naturalWidthPt,
        float naturalHeightPt)
    {
        if (!MathTypeOleStorage.LooksLikeMathTypeCompoundFile(compoundFile))
            throw new InvalidDataException("The MathType storage to embed is not valid.");
        if (previewEmf.Length < 88 || BitConverter.ToUInt32(previewEmf, 0) != 1 || BitConverter.ToUInt32(previewEmf, 40) != 0x464D4520)
            throw new InvalidDataException("The MathType preview is not an enhanced metafile.");
        using var file = File.Open(packagePath, FileMode.Open, FileAccess.ReadWrite);
        using var archive = new ZipArchive(file, ZipArchiveMode.Update);
        const string presentationPart = "ppt/presentation.xml";
        var firstSlide = ReadXml(archive, presentationPart).Descendants(P + "sldId").FirstOrDefault()
            ?? throw new InvalidDataException("The staging presentation has no slide.");
        var slidePart = ResolveRelationship(archive, presentationPart, (string?)firstSlide.Attribute(R + "id"));
        var slash = slidePart.LastIndexOf('/');
        var relsPart = slidePart.Substring(0, slash + 1) + "_rels/" + slidePart.Substring(slash + 1) + ".rels";

        var slide = ReadXml(archive, slidePart);
        var placeholder = slide.Descendants(P + "sp").SingleOrDefault(shape =>
                (string?)shape.Element(P + "nvSpPr")?.Element(P + "cNvPr")?.Attribute("name") == StagingPlaceholderName)
            ?? throw new InvalidDataException("The staging placeholder is missing.");
        var id = (string?)placeholder.Element(P + "nvSpPr")!.Element(P + "cNvPr")!.Attribute("id")
            ?? throw new InvalidDataException("The staging placeholder has no shape id.");
        var transform = placeholder.Element(P + "spPr")?.Element(A + "xfrm")
            ?? throw new InvalidDataException("The staging placeholder has no geometry.");

        var relationships = ReadXml(archive, relsPart);
        var usedIds = new HashSet<string>(relationships.Root!.Elements(Rel + "Relationship")
            .Select(item => (string?)item.Attribute("Id") ?? string.Empty), StringComparer.Ordinal);
        string NewId(string stem) { var n = 1; while (usedIds.Contains(stem + n)) n++; usedIds.Add(stem + n); return stem + n; }
        string NewPart(string folder, string stem, string extension)
        {
            var n = 1;
            while (archive.GetEntry($"ppt/{folder}/{stem}{n}.{extension}") is not null) n++;
            return $"ppt/{folder}/{stem}{n}.{extension}";
        }
        var storagePart = NewPart("embeddings", "oleObjectVisualTeX", "bin");
        var previewPart = NewPart("media", "imageVisualTeX", "emf");
        var storageId = NewId("rIdVisualTeXOle");
        var previewId = NewId("rIdVisualTeXPreview");
        relationships.Root.Add(
            new XElement(Rel + "Relationship", new XAttribute("Id", storageId),
                new XAttribute("Type", OleObjectRelationship), new XAttribute("Target", "../" + storagePart.Substring(4))),
            new XElement(Rel + "Relationship", new XAttribute("Id", previewId),
                new XAttribute("Type", ImageRelationship), new XAttribute("Target", "../" + previewPart.Substring(4))));

        string Emu(float points) => Math.Max(1L, (long)Math.Round(points * 12700d)).ToString(System.Globalization.CultureInfo.InvariantCulture);
        placeholder.ReplaceWith(new XElement(P + "graphicFrame",
            new XElement(P + "nvGraphicFramePr",
                new XElement(P + "cNvPr", new XAttribute("id", id), new XAttribute("name", "Equation")),
                new XElement(P + "cNvGraphicFramePr", new XElement(A + "graphicFrameLocks", new XAttribute("noChangeAspect", "1"))),
                new XElement(P + "nvPr")),
            new XElement(P + "xfrm", transform.Elements()),
            new XElement(A + "graphic",
                new XElement(A + "graphicData", new XAttribute("uri", "http://schemas.openxmlformats.org/presentationml/2006/ole"),
                    new XElement(P + "oleObj",
                        new XAttribute("name", "Equation"), new XAttribute(R + "id", storageId),
                        new XAttribute("imgW", Emu(naturalWidthPt)), new XAttribute("imgH", Emu(naturalHeightPt)),
                        new XAttribute("progId", MathTypeOleInterop.CanonicalProgId),
                        new XElement(P + "embed"),
                        new XElement(P + "pic",
                            new XElement(P + "nvPicPr",
                                new XElement(P + "cNvPr", new XAttribute("id", "0"), new XAttribute("name", "")),
                                new XElement(P + "cNvPicPr"), new XElement(P + "nvPr")),
                            new XElement(P + "blipFill",
                                new XElement(A + "blip", new XAttribute(R + "embed", previewId)),
                                new XElement(A + "stretch", new XElement(A + "fillRect"))),
                            new XElement(P + "spPr",
                                new XElement(transform),
                                new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst")))))))));

        var types = ReadXml(archive, "[Content_Types].xml");
        void EnsureDefault(string extension, string contentType)
        {
            if (types.Root!.Elements(ContentTypes + "Default").Any(item =>
                    string.Equals((string?)item.Attribute("Extension"), extension, StringComparison.OrdinalIgnoreCase)))
                return;
            types.Root.AddFirst(new XElement(ContentTypes + "Default",
                new XAttribute("Extension", extension), new XAttribute("ContentType", contentType)));
        }
        EnsureDefault("bin", "application/vnd.openxmlformats-officedocument.oleObject");
        EnsureDefault("emf", "image/x-emf");

        WriteXml(archive, slidePart, slide);
        WriteXml(archive, relsPart, relationships);
        WriteXml(archive, "[Content_Types].xml", types);
        WriteBytes(archive, storagePart, compoundFile);
        WriteBytes(archive, previewPart, previewEmf);
    }

    private static void WriteXml(ZipArchive archive, string part, XDocument document)
    {
        using var buffer = new MemoryStream();
        using (var writer = XmlWriter.Create(buffer, new XmlWriterSettings { Encoding = new System.Text.UTF8Encoding(false) }))
            document.Save(writer);
        WriteBytes(archive, part, buffer.ToArray());
    }

    private static void WriteBytes(ZipArchive archive, string part, byte[] bytes)
    {
        archive.GetEntry(part)?.Delete();
        using var stream = archive.CreateEntry(part).Open();
        stream.Write(bytes, 0, bytes.Length);
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
