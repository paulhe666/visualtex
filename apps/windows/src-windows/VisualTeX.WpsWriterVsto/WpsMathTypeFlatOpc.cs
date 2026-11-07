using System.Globalization;
using System.IO;
using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WpsWriterVsto;

/// <summary>
/// Builds the Flat OPC fragment WPS/Word uses to materialize a standalone
/// MathType Equation.DSMT4 object. The fragment carries the genuine compound
/// file plus a placeable WMF preview, so it never needs a registered MathType
/// server to load.
/// </summary>
internal static class WpsMathTypeFlatOpc
{
    private static readonly XNamespace PackageNamespace =
        "http://schemas.microsoft.com/office/2006/xmlPackage";
    private static readonly XNamespace RelationshipNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace OfficeRelationshipNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace OfficeNamespace =
        "urn:schemas-microsoft-com:office:office";
    private static readonly XNamespace VmlNamespace =
        "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace WordNamespace =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal enum NumberPlacement
    {
        None,
        Right,
        Left,
    }

    internal readonly struct DisplayTabs
    {
        internal DisplayTabs(float centerPt, float rightPt)
        {
            CenterPt = centerPt;
            RightPt = rightPt;
        }

        internal float CenterPt { get; }
        internal float RightPt { get; }
    }

    internal static string Build(
        byte[] compoundFile,
        byte[] previewWmf,
        float widthPt,
        float heightPt,
        bool display,
        string? numberFormatId,
        string? bookmarkName,
        DisplayTabs? displayTabs = null)
    {
        if (compoundFile is null || compoundFile.Length == 0)
            throw new InvalidDataException("MathType Flat OPC requires a compound file.");
        if (previewWmf is null || previewWmf.Length <= 22)
            throw new InvalidDataException("MathType Flat OPC requires a placeable WMF preview.");
        if (!(widthPt > 0) || !(heightPt > 0))
            throw new InvalidDataException(
                $"Invalid MathType preview size {widthPt}x{heightPt} pt.");

        previewWmf = EnsurePlaceableWmfWindowRecords(previewWmf);

        var segments = display ? CreateNumberSegments(numberFormatId) : Array.Empty<NumberSegmentSpec>();
        var placement = segments.Length == 0
            ? NumberPlacement.None
            : NumberPlacement.Right;
        if (placement != NumberPlacement.None && !display)
            throw new InvalidDataException(
                "Equation numbering is valid only for display formulas.");

        var oleProgId = MathTypeOleInterop.CanonicalProgId;
        try
        {
            var identity = MathTypeOleStorage.ReadCompoundFileIdentity(compoundFile);
            if (!string.IsNullOrWhiteSpace(identity.ProgId))
                oleProgId = identity.ProgId;
        }
        catch
        {
            // Identity probing is best effort; the canonical ProgID is valid.
        }

        var instanceId = Guid.NewGuid();
        var instanceValue = BitConverter.ToUInt32(instanceId.ToByteArray(), 0);
        var shapeId = "_x0000_i" + (1025u + (instanceValue % 2_000_000_000u))
            .ToString(CultureInfo.InvariantCulture);
        var bookmarkId = (1u + (instanceValue % 2_000_000_000u))
            .ToString(CultureInfo.InvariantCulture);
        var objectId = "_" + unchecked((uint)instanceId.GetHashCode())
            .ToString(CultureInfo.InvariantCulture);
        var width = widthPt.ToString("0.###", CultureInfo.InvariantCulture);
        var height = heightPt.ToString("0.###", CultureInfo.InvariantCulture);

        var rootRelationships = new XElement(
            RelationshipNamespace + "Relationships",
            new XElement(
                RelationshipNamespace + "Relationship",
                new XAttribute("Id", "rId1"),
                new XAttribute(
                    "Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
                new XAttribute("Target", "word/document.xml")));

        var documentRelationships = new XElement(
            RelationshipNamespace + "Relationships",
            new XElement(
                RelationshipNamespace + "Relationship",
                new XAttribute("Id", "rId1"),
                new XAttribute(
                    "Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
                new XAttribute("Target", "media/image1.wmf")),
            new XElement(
                RelationshipNamespace + "Relationship",
                new XAttribute("Id", "rId2"),
                new XAttribute(
                    "Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"),
                new XAttribute("Target", "embeddings/oleObject1.bin")));

        var shapeType = new XElement(
            VmlNamespace + "shapetype",
            new XAttribute("id", "_x0000_t75"),
            new XAttribute("coordsize", "21600,21600"),
            new XAttribute(OfficeNamespace + "spt", "75"),
            new XAttribute(OfficeNamespace + "preferrelative", "t"),
            new XAttribute("path", "m@4@5l@4@11@9@11@9@5xe"),
            new XAttribute("filled", "f"),
            new XAttribute("stroked", "f"),
            new XElement(VmlNamespace + "stroke", new XAttribute("joinstyle", "miter")),
            new XElement(
                VmlNamespace + "formulas",
                new XElement(VmlNamespace + "f", new XAttribute("eqn", "if lineDrawn pixelLineWidth 0")),
                new XElement(VmlNamespace + "f", new XAttribute("eqn", "sum @0 1 0")),
                new XElement(VmlNamespace + "f", new XAttribute("eqn", "sum 0 0 @1")),
                new XElement(VmlNamespace + "f", new XAttribute("eqn", "prod @2 1 2")),
                new XElement(VmlNamespace + "f", new XAttribute("eqn", "prod @3 21600 pixelWidth")),
                new XElement(VmlNamespace + "f", new XAttribute("eqn", "prod @3 21600 pixelHeight")),
                new XElement(VmlNamespace + "f", new XAttribute("eqn", "sum @0 0 1")),
                new XElement(VmlNamespace + "f", new XAttribute("eqn", "prod @6 1 2")),
                new XElement(VmlNamespace + "f", new XAttribute("eqn", "prod @7 21600 pixelWidth")),
                new XElement(VmlNamespace + "f", new XAttribute("eqn", "sum @8 21600 0")),
                new XElement(VmlNamespace + "f", new XAttribute("eqn", "prod @7 21600 pixelHeight")),
                new XElement(VmlNamespace + "f", new XAttribute("eqn", "sum @10 21600 0"))),
            new XElement(
                VmlNamespace + "path",
                new XAttribute(OfficeNamespace + "extrusionok", "f"),
                new XAttribute("gradientshapeok", "t"),
                new XAttribute(OfficeNamespace + "connecttype", "rect")),
            new XElement(
                OfficeNamespace + "lock",
                new XAttribute(VmlNamespace + "ext", "edit"),
                new XAttribute("aspectratio", "t")));

        var shape = new XElement(
            VmlNamespace + "shape",
            new XAttribute("id", shapeId),
            new XAttribute("type", "#_x0000_t75"),
            new XAttribute("style", $"width:{width}pt;height:{height}pt"),
            new XAttribute(OfficeNamespace + "ole", string.Empty),
            new XElement(
                VmlNamespace + "imagedata",
                new XAttribute(OfficeRelationshipNamespace + "id", "rId1"),
                new XAttribute(OfficeNamespace + "title", string.Empty)));

        var oleObject = new XElement(
            OfficeNamespace + "OLEObject",
            new XAttribute("Type", "Embed"),
            new XAttribute("ProgID", oleProgId),
            new XAttribute("ShapeID", shapeId),
            new XAttribute("DrawAspect", "Content"),
            new XAttribute("ObjectID", objectId),
            new XAttribute(OfficeRelationshipNamespace + "id", "rId2"));

        var wordObject = new XElement(
            WordNamespace + "object",
            new XAttribute(WordNamespace + "dxaOrig", Twips(widthPt)),
            new XAttribute(WordNamespace + "dyaOrig", Twips(heightPt)),
            shapeType,
            shape,
            oleObject);

        var paragraph = new XElement(WordNamespace + "p");

        // A centered display equation needs a center tab at the middle of the
        // text column and a right tab at its right edge. Paragraph centering
        // plus a single right tab overflows the column and wraps the number.
        if (display && displayTabs is { } tabs)
        {
            paragraph.Add(new XElement(
                WordNamespace + "pPr",
                new XElement(
                    WordNamespace + "tabs",
                    new XElement(
                        WordNamespace + "tab",
                        new XAttribute(WordNamespace + "val", "center"),
                        new XAttribute(WordNamespace + "pos", Twips(tabs.CenterPt))),
                    new XElement(
                        WordNamespace + "tab",
                        new XAttribute(WordNamespace + "val", "right"),
                        new XAttribute(WordNamespace + "pos", Twips(tabs.RightPt)))),
                new XElement(
                    WordNamespace + "jc",
                    new XAttribute(WordNamespace + "val", "left"))));
        }

        if (display && placement != NumberPlacement.Left)
            paragraph.Add(TabRun());
        else if (placement == NumberPlacement.Left)
            AddNumber(paragraph, segments);

        if (placement == NumberPlacement.Left)
            paragraph.Add(TabRun());

        if (!string.IsNullOrWhiteSpace(bookmarkName))
        {
            paragraph.Add(new XElement(
                WordNamespace + "bookmarkStart",
                new XAttribute(WordNamespace + "id", bookmarkId),
                new XAttribute(WordNamespace + "name", bookmarkName!)));
        }

        paragraph.Add(new XElement(WordNamespace + "r", wordObject));

        if (!string.IsNullOrWhiteSpace(bookmarkName))
        {
            paragraph.Add(new XElement(
                WordNamespace + "bookmarkEnd",
                new XAttribute(WordNamespace + "id", bookmarkId)));
        }

        if (placement == NumberPlacement.Right)
        {
            paragraph.Add(TabRun());
            AddNumber(paragraph, segments);
        }

        var wordDocument = new XElement(
            WordNamespace + "document",
            new XAttribute(XNamespace.Xmlns + "w", WordNamespace),
            new XAttribute(XNamespace.Xmlns + "r", OfficeRelationshipNamespace),
            new XAttribute(XNamespace.Xmlns + "o", OfficeNamespace),
            new XAttribute(XNamespace.Xmlns + "v", VmlNamespace),
            new XElement(
                WordNamespace + "body",
                paragraph,
                new XElement(
                    WordNamespace + "sectPr",
                    new XElement(
                        WordNamespace + "pgSz",
                        new XAttribute(WordNamespace + "w", "12240"),
                        new XAttribute(WordNamespace + "h", "15840")),
                    new XElement(
                        WordNamespace + "pgMar",
                        new XAttribute(WordNamespace + "top", "1440"),
                        new XAttribute(WordNamespace + "right", "1440"),
                        new XAttribute(WordNamespace + "bottom", "1440"),
                        new XAttribute(WordNamespace + "left", "1440")))));

        static XElement XmlPart(string name, string contentType, XElement xml) =>
            new(
                PackageNamespace + "part",
                new XAttribute(PackageNamespace + "name", name),
                new XAttribute(PackageNamespace + "contentType", contentType),
                new XElement(PackageNamespace + "xmlData", xml));

        static XElement BinaryPart(string name, string contentType, byte[] bytes) =>
            new(
                PackageNamespace + "part",
                new XAttribute(PackageNamespace + "name", name),
                new XAttribute(PackageNamespace + "contentType", contentType),
                new XElement(PackageNamespace + "binaryData", Convert.ToBase64String(bytes)));

        return new XDocument(
            new XElement(
                PackageNamespace + "package",
                new XAttribute(XNamespace.Xmlns + "pkg", PackageNamespace),
                XmlPart(
                    "/_rels/.rels",
                    "application/vnd.openxmlformats-package.relationships+xml",
                    rootRelationships),
                XmlPart(
                    "/word/document.xml",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                    wordDocument),
                XmlPart(
                    "/word/_rels/document.xml.rels",
                    "application/vnd.openxmlformats-package.relationships+xml",
                    documentRelationships),
                BinaryPart("/word/media/image1.wmf", "image/x-wmf", previewWmf),
                BinaryPart(
                    "/word/embeddings/oleObject1.bin",
                    "application/vnd.openxmlformats-officedocument.oleObject",
                    compoundFile)))
            .ToString(SaveOptions.DisableFormatting);
    }

    internal static bool IsKnownNumberFormat(string? formatId) =>
        formatId is "off" or "continuous" or "heading1-dot" or "heading1-dash"
            or "heading2-dot" or "heading2-dash";

    private readonly struct NumberSegmentSpec
    {
        internal NumberSegmentSpec(bool isField, string value)
        {
            IsField = isField;
            Value = value;
        }

        internal bool IsField { get; }
        internal string Value { get; }
    }

    private static NumberSegmentSpec[] CreateNumberSegments(string? formatId)
    {
        var format = string.IsNullOrWhiteSpace(formatId) ? "off" : formatId!;
        if (format == "off") return Array.Empty<NumberSegmentSpec>();

        var headingLevel = format.StartsWith("heading2", StringComparison.OrdinalIgnoreCase)
            ? 2
            : format.StartsWith("heading1", StringComparison.OrdinalIgnoreCase)
                ? 1
                : 0;
        var separator = format.EndsWith("-dash", StringComparison.OrdinalIgnoreCase)
            ? "-"
            : ".";

        var segments = new List<NumberSegmentSpec>
        {
            new(false, "("),
        };
        if (headingLevel >= 1)
        {
            segments.Add(new(true, "STYLEREF 1 \\s \\* MERGEFORMAT"));
            segments.Add(new(false, headingLevel == 1 ? separator : "."));
        }
        if (headingLevel >= 2)
        {
            segments.Add(new(true, "STYLEREF 2 \\s \\* MERGEFORMAT"));
            segments.Add(new(false, separator));
        }
        segments.Add(new(true, "SEQ VisualTeXEquation \\* ARABIC"));
        segments.Add(new(false, ")"));
        return segments.ToArray();
    }

    private static void AddNumber(XElement paragraph, NumberSegmentSpec[] segments)
    {
        foreach (var segment in segments)
        {
            if (!segment.IsField)
            {
                paragraph.Add(new XElement(
                    WordNamespace + "r",
                    new XElement(
                        WordNamespace + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        segment.Value)));
                continue;
            }

            paragraph.Add(new XElement(
                WordNamespace + "r",
                new XElement(
                    WordNamespace + "fldChar",
                    new XAttribute(WordNamespace + "fldCharType", "begin"))));
            paragraph.Add(new XElement(
                WordNamespace + "r",
                new XElement(
                    WordNamespace + "instrText",
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    " " + segment.Value + " ")));
            paragraph.Add(new XElement(
                WordNamespace + "r",
                new XElement(
                    WordNamespace + "fldChar",
                    new XAttribute(WordNamespace + "fldCharType", "separate"))));
            paragraph.Add(new XElement(
                WordNamespace + "r",
                new XElement(WordNamespace + "t", "1")));
            paragraph.Add(new XElement(
                WordNamespace + "r",
                new XElement(
                    WordNamespace + "fldChar",
                    new XAttribute(WordNamespace + "fldCharType", "end"))));
        }
    }

    private static XElement TabRun() =>
        new(WordNamespace + "r", new XElement(WordNamespace + "tab"));

    private static string Twips(float points) =>
        checked((int)Math.Round(points * 20d)).ToString(CultureInfo.InvariantCulture);

    // WPS/Word's on-screen OLE renderer accepts a plain placeable WMF, but the
    // PDF/print pipeline needs explicit SETWINDOWORG + SETWINDOWEXT records.
    // The MathPage renderer can omit them, which yields a blank equation in PDF.
    internal static byte[] EnsurePlaceableWmfWindowRecords(byte[] previewWmf)
    {
        const uint aldusPlaceableKey = 0x9ac6cdd7;
        const ushort metaSetWindowOrg = 0x020b;
        const ushort metaSetWindowExt = 0x020c;
        const int firstRecordOffset = 40;

        if (previewWmf is null
            || previewWmf.Length < firstRecordOffset + 6
            || BitConverter.ToUInt32(previewWmf, 0) != aldusPlaceableKey)
            throw new InvalidDataException("MathType preview is not a valid placeable WMF.");

        var declaredWords = BitConverter.ToUInt32(previewWmf, 28);
        if (declaredWords > int.MaxValue / 2
            || checked((int)declaredWords * 2) != previewWmf.Length - 22)
            throw new InvalidDataException("MathType WMF header length does not match its payload.");

        var hasWindowOrigin = false;
        var hasWindowExtent = false;
        var insertionOffset = firstRecordOffset;
        var leadingSetup = true;
        for (var offset = firstRecordOffset; offset + 6 <= previewWmf.Length;)
        {
            var words = BitConverter.ToUInt32(previewWmf, offset);
            if (words < 3 || words > int.MaxValue / 2)
                throw new InvalidDataException("MathType WMF contains an invalid record length.");
            var size = checked((int)words * 2);
            if (offset + size > previewWmf.Length)
                throw new InvalidDataException("MathType WMF contains a truncated record.");
            var function = BitConverter.ToUInt16(previewWmf, offset + 4);
            hasWindowOrigin |= function == metaSetWindowOrg;
            hasWindowExtent |= function == metaSetWindowExt;
            if (leadingSetup && function is 0x0102 or 0x0201 or 0x012e)
                insertionOffset = offset + size;
            else
                leadingSetup = false;
            offset += size;
            if (function == 0) break;
        }

        if (hasWindowOrigin && hasWindowExtent) return previewWmf;
        if (hasWindowOrigin || hasWindowExtent)
            throw new InvalidDataException(
                "MathType WMF contains only one of SETWINDOWORG/SETWINDOWEXT.");

        var left = BitConverter.ToInt16(previewWmf, 6);
        var top = BitConverter.ToInt16(previewWmf, 8);
        var right = BitConverter.ToInt16(previewWmf, 10);
        var bottom = BitConverter.ToInt16(previewWmf, 12);
        var width = checked((short)(right - left));
        var height = checked((short)(bottom - top));
        if (width <= 0 || height <= 0)
            throw new InvalidDataException("MathType placeable WMF has invalid logical bounds.");

        var result = new byte[checked(previewWmf.Length + 20)];
        Buffer.BlockCopy(previewWmf, 0, result, 0, insertionOffset);
        WriteWmfRecord(result, insertionOffset, metaSetWindowOrg, top, left);
        WriteWmfRecord(result, insertionOffset + 10, metaSetWindowExt, height, width);
        Buffer.BlockCopy(
            previewWmf,
            insertionOffset,
            result,
            insertionOffset + 20,
            previewWmf.Length - insertionOffset);
        BitConverter.GetBytes(checked(declaredWords + 10u)).CopyTo(result, 28);
        return result;
    }

    private static void WriteWmfRecord(
        byte[] buffer,
        int offset,
        ushort function,
        short y,
        short x)
    {
        BitConverter.GetBytes(5u).CopyTo(buffer, offset);
        BitConverter.GetBytes(function).CopyTo(buffer, offset + 4);
        BitConverter.GetBytes(y).CopyTo(buffer, offset + 6);
        BitConverter.GetBytes(x).CopyTo(buffer, offset + 8);
    }
}
