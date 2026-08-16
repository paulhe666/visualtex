using System.Runtime.InteropServices;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

/// <summary>
/// Reads and rewrites a traditional MathType OLE directly through Word's Flat OPC
/// representation. This path preserves the embedded Equation.DSMT4 Compound File
/// and Word's companion WMF preview without requiring a registered MathType OLE
/// server on the current machine.
/// </summary>
internal static class MathTypeWordOpenXml
{
    private const uint AldusPlaceableKey = 0x9AC6CDD7;
    private const ushort PlaceableInch = 2304;
    private const int MmAnisotropic = 8;

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

    internal sealed class Fragment
    {
        internal string WordOpenXml { get; set; } = string.Empty;
        internal string ProgId { get; set; } = string.Empty;
        internal byte[] CompoundFile { get; set; } = Array.Empty<byte>();
        internal byte[] PreviewWmf { get; set; } = Array.Empty<byte>();
        internal float WidthPt { get; set; }
        internal float HeightPt { get; set; }
    }

    internal sealed class NumberTemplate
    {
        internal List<NumberSegment> Segments { get; } = new();
    }

    internal readonly struct NumberSegment
    {
        private NumberSegment(bool isField, string value)
        {
            IsField = isField;
            Value = value ?? string.Empty;
        }

        internal bool IsField { get; }
        internal string Value { get; }

        internal static NumberSegment Text(string value) => new(false, value);
        internal static NumberSegment Field(string instruction) => new(true, instruction);
    }

    internal static NumberTemplate CreateDefaultNumberTemplate()
    {
        var template = new NumberTemplate();
        template.Segments.Add(NumberSegment.Text(" \\* MERGEFORMAT "));
        template.Segments.Add(NumberSegment.Field(" SEQ MTEqn \\h \\* MERGEFORMAT "));
        template.Segments.Add(NumberSegment.Text("("));
        template.Segments.Add(NumberSegment.Field(" SEQ MTSec \\c \\* Arabic \\* MERGEFORMAT "));
        template.Segments.Add(NumberSegment.Text("."));
        template.Segments.Add(NumberSegment.Field(" SEQ MTEqn \\c \\* Arabic \\* MERGEFORMAT "));
        template.Segments.Add(NumberSegment.Text(")"));
        return template;
    }

    internal static Fragment Read(Word.InlineShape shape)
    {
        Word.Range? range = null;
        try
        {
            range = shape.Range;
            return Read(range.WordOpenXML);
        }
        finally { Release(range); }
    }

    internal static Fragment Read(string wordOpenXml)
    {
        if (string.IsNullOrWhiteSpace(wordOpenXml))
            throw new InvalidDataException("Word returned empty Flat OPC for the MathType equation.");

        var document = XDocument.Parse(wordOpenXml, LoadOptions.PreserveWhitespace);
        var references = ResolveObjectReferences(document);
        var compound = ReadBinaryPart(document, references.EmbeddingPartName);
        if (!MathTypeOleStorage.LooksLikeMathTypeCompoundFile(compound))
            throw new InvalidDataException(
                "The Flat OPC embedded object is not a MathType Equation.DSMT4 Compound File.");
        var preview = ReadBinaryPart(document, references.PreviewPartName);
        if (preview.Length == 0)
            throw new InvalidDataException("The MathType Flat OPC preview image is empty.");

        var (widthPt, heightPt) = ReadShapeSize(document);
        return new Fragment
        {
            WordOpenXml = wordOpenXml,
            ProgId = references.ProgId,
            CompoundFile = compound,
            PreviewWmf = preview,
            WidthPt = widthPt,
            HeightPt = heightPt,
        };
    }

    internal static string Rewrite(
        string sourceWordOpenXml,
        byte[] rewrittenCompoundFile,
        string emfPath,
        float widthPt,
        float heightPt)
    {
        if (string.IsNullOrWhiteSpace(emfPath) || !File.Exists(emfPath))
            throw new FileNotFoundException(
                "MathType Flat OPC rewrite requires an existing VisualTeX EMF preview.",
                emfPath);
        var previewWmf = ConvertEnhancedMetafileToPlaceableWmf(
            emfPath,
            widthPt,
            heightPt);
        return RewriteWithPlaceableWmf(
            sourceWordOpenXml,
            rewrittenCompoundFile,
            previewWmf,
            widthPt,
            heightPt);
    }

    internal static string CreateWithPlaceableWmf(
        byte[] compoundFile,
        byte[] previewWmf,
        float widthPt,
        float heightPt,
        bool display = false,
        NumberTemplate? numberTemplate = null)
    {
        if (!MathTypeOleStorage.LooksLikeMathTypeCompoundFile(compoundFile))
            throw new InvalidDataException(
                "MathType Flat OPC creation requires a valid Equation.DSMT4 Compound File.");
        if (previewWmf is null || previewWmf.Length <= 22)
            throw new InvalidDataException("MathType Flat OPC creation requires a valid WMF preview.");
        if (!(widthPt > 0) || !(heightPt > 0))
            throw new InvalidDataException(
                $"Invalid MathType preview size {widthPt}x{heightPt} pt.");

        var imageRelationshipId = "rId1";
        var oleRelationshipId = "rId2";
        var shapeId = "_x0000_i1025";
        var objectId = "_" + unchecked((uint)Guid.NewGuid().GetHashCode()).ToString(
            System.Globalization.CultureInfo.InvariantCulture);
        var width = widthPt.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
        var height = heightPt.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);

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
                new XAttribute("Id", imageRelationshipId),
                new XAttribute(
                    "Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
                new XAttribute("Target", "media/image1.wmf")),
            new XElement(
                RelationshipNamespace + "Relationship",
                new XAttribute("Id", oleRelationshipId),
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
                new XAttribute(OfficeRelationshipNamespace + "id", imageRelationshipId),
                new XAttribute(OfficeNamespace + "title", string.Empty)));
        var oleObject = new XElement(
            OfficeNamespace + "OLEObject",
            new XAttribute("Type", "Embed"),
            new XAttribute("ProgID", "Equation.DSMT4"),
            new XAttribute("ShapeID", shapeId),
            new XAttribute("DrawAspect", "Content"),
            new XAttribute("ObjectID", objectId),
            new XAttribute(OfficeRelationshipNamespace + "id", oleRelationshipId));
        var wordObject = new XElement(
            WordNamespace + "object",
            new XAttribute(
                WordNamespace + "dxaOrig",
                Math.Max(1, (int)Math.Round(widthPt * 20d))),
            new XAttribute(
                WordNamespace + "dyaOrig",
                Math.Max(1, (int)Math.Round(heightPt * 20d))),
            shapeType,
            shape,
            oleObject);
        var paragraph = new XElement(WordNamespace + "p");
        if (display)
            paragraph.Add(new XElement(
                WordNamespace + "r",
                new XElement(WordNamespace + "tab")));
        paragraph.Add(new XElement(WordNamespace + "r", wordObject));
        if (numberTemplate is not null)
        {
            if (!display)
                throw new InvalidDataException(
                    "MathType equation numbering is valid only for display equations.");
            paragraph.Add(new XElement(
                WordNamespace + "r",
                new XElement(WordNamespace + "tab")));
            foreach (var node in BuildMathTypePlaceRef(numberTemplate))
                paragraph.Add(node);
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

        XElement XmlPart(string name, string contentType, XElement xml) =>
            new(
                PackageNamespace + "part",
                new XAttribute(PackageNamespace + "name", name),
                new XAttribute(PackageNamespace + "contentType", contentType),
                new XElement(PackageNamespace + "xmlData", xml));
        XElement BinaryPart(string name, string contentType, byte[] bytes) =>
            new(
                PackageNamespace + "part",
                new XAttribute(PackageNamespace + "name", name),
                new XAttribute(PackageNamespace + "contentType", contentType),
                new XElement(PackageNamespace + "binaryData", Convert.ToBase64String(bytes)));

        var package = new XDocument(
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
                    compoundFile)));
        var xmlText = package.ToString(SaveOptions.DisableFormatting);
        var validation = Read(xmlText);
        if (!string.Equals(validation.ProgId, "Equation.DSMT4", StringComparison.OrdinalIgnoreCase)
            || !validation.CompoundFile.SequenceEqual(compoundFile)
            || !validation.PreviewWmf.SequenceEqual(previewWmf))
            throw new InvalidDataException(
                "VisualTeX's standalone MathType Flat OPC failed self-validation.");
        return xmlText;
    }

    internal static string CreateDefaultSectionBreakFlatOpc(string displayLabel)
    {
        if (string.IsNullOrWhiteSpace(displayLabel))
            throw new InvalidDataException("MathType chapter/section break label is empty.");

        var rootRelationships = new XElement(
            RelationshipNamespace + "Relationships",
            new XElement(
                RelationshipNamespace + "Relationship",
                new XAttribute("Id", "rId1"),
                new XAttribute(
                    "Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
                new XAttribute("Target", "word/document.xml")));
        var documentRelationships = new XElement(RelationshipNamespace + "Relationships");
        var paragraph = new XElement(WordNamespace + "p");
        foreach (var run in BuildMathTypeSectionBreak(displayLabel))
            paragraph.Add(run);
        var wordDocument = new XElement(
            WordNamespace + "document",
            new XAttribute(XNamespace.Xmlns + "w", WordNamespace),
            new XElement(
                WordNamespace + "body",
                paragraph,
                new XElement(WordNamespace + "sectPr")));

        XElement XmlPart(string name, string contentType, XElement xml) =>
            new(
                PackageNamespace + "part",
                new XAttribute(PackageNamespace + "name", name),
                new XAttribute(PackageNamespace + "contentType", contentType),
                new XElement(PackageNamespace + "xmlData", xml));
        var package = new XDocument(
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
                    documentRelationships)));
        return package.ToString(SaveOptions.DisableFormatting);
    }

    private static IEnumerable<XElement> BuildMathTypeSectionBreak(string displayLabel)
    {
        yield return FieldCharRun("begin");
        yield return InstructionRun($" MACROBUTTON MTEditEquationSection2 {displayLabel}");
        foreach (var run in BuildSimpleComplexField(" SEQ MTEqn \\r \\h \\* MERGEFORMAT "))
            yield return run;
        foreach (var run in BuildSimpleComplexField(" SEQ MTSec \\r 1 \\h \\* MERGEFORMAT "))
            yield return run;
        foreach (var run in BuildSimpleComplexField(" SEQ MTChap \\r 1 \\h \\* MERGEFORMAT "))
            yield return run;
        yield return FieldCharRun("separate");
        yield return FieldCharRun("end");
    }

    private static IEnumerable<XElement> BuildMathTypePlaceRef(NumberTemplate template)
    {
        if (template.Segments.Count == 0)
            throw new InvalidDataException("MathType MTPlaceRef numbering template is empty.");

        yield return FieldCharRun("begin");
        yield return InstructionRun(" MACROBUTTON MTPlaceRef");
        foreach (var segment in template.Segments)
        {
            if (!segment.IsField)
            {
                if (!string.IsNullOrEmpty(segment.Value))
                    yield return InstructionRun(segment.Value);
                continue;
            }
            foreach (var nestedRun in BuildSimpleComplexField(segment.Value))
                yield return nestedRun;
        }
        yield return FieldCharRun("separate");
        yield return FieldCharRun("end");
    }

    private static IEnumerable<XElement> BuildSimpleComplexField(string instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            throw new InvalidDataException("MathType nested Word field instruction is empty.");
        yield return FieldCharRun("begin");
        yield return InstructionRun(instruction);
        yield return FieldCharRun("separate");
        yield return FieldCharRun("end");
    }

    private static XElement FieldCharRun(string type) =>
        new(
            WordNamespace + "r",
            new XElement(
                WordNamespace + "fldChar",
                new XAttribute(WordNamespace + "fldCharType", type)));

    private static XElement InstructionRun(string text) =>
        new(
            WordNamespace + "r",
            new XElement(
                WordNamespace + "instrText",
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                text));

    internal static string RewriteWithPlaceableWmf(
        string sourceWordOpenXml,
        byte[] rewrittenCompoundFile,
        byte[] previewWmf,
        float widthPt,
        float heightPt)
    {
        if (!MathTypeOleStorage.LooksLikeMathTypeCompoundFile(rewrittenCompoundFile))
            throw new InvalidDataException(
                "MathType Flat OPC rewrite requires a valid Equation.DSMT4 Compound File.");
        if (previewWmf is null || previewWmf.Length <= 22)
            throw new InvalidDataException("MathType Flat OPC rewrite requires a valid WMF preview.");
        if (!(widthPt > 0) || !(heightPt > 0))
            throw new InvalidDataException(
                $"Invalid MathType preview size {widthPt}x{heightPt} pt.");

        var document = XDocument.Parse(sourceWordOpenXml, LoadOptions.PreserveWhitespace);
        var references = ResolveObjectReferences(document);
        WriteBinaryPart(document, references.EmbeddingPartName, rewrittenCompoundFile);
        var previewPart = FindPart(document, references.PreviewPartName);
        previewPart.SetAttributeValue(PackageNamespace + "contentType", "image/x-wmf");
        WriteBinaryPart(document, references.PreviewPartName, previewWmf);
        UpdateShapeSize(document, widthPt, heightPt);

        var rewrittenXml = document.ToString(SaveOptions.DisableFormatting);
        var validation = Read(rewrittenXml);
        if (!string.Equals(
                validation.ProgId,
                references.ProgId,
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException(
                "MathType Flat OPC rewrite unexpectedly changed the OLE ProgID.");
        if (!validation.CompoundFile.SequenceEqual(rewrittenCompoundFile))
            throw new InvalidDataException(
                "MathType Flat OPC rewrite did not preserve the rewritten Compound File bytes.");
        return rewrittenXml;
    }

    internal static string ConvertPlaceableWmfToEnhancedMetafile(
        string wmfPath,
        float widthPt,
        float heightPt,
        string outputDirectory)
    {
        if (string.IsNullOrWhiteSpace(wmfPath) || !File.Exists(wmfPath))
            throw new FileNotFoundException("WMF preview is unavailable.", wmfPath);
        if (!(widthPt > 0) || !(heightPt > 0))
            throw new InvalidDataException(
                $"Invalid MathType native preview size {widthPt}x{heightPt} pt.");
        if (string.IsNullOrWhiteSpace(outputDirectory))
            outputDirectory = Path.GetTempPath();
        Directory.CreateDirectory(outputDirectory);

        var placeableWmf = File.ReadAllBytes(wmfPath);
        if (placeableWmf.Length <= 22
            || BitConverter.ToUInt32(placeableWmf, 0) != AldusPlaceableKey)
            throw new InvalidDataException(
                "MathType native preview is not an Aldus placeable WMF.");
        var rawWmf = new byte[placeableWmf.Length - 22];
        Buffer.BlockCopy(placeableWmf, 22, rawWmf, 0, rawWmf.Length);

        var metafilePicture = new MetafilePictureNative
        {
            MappingMode = MmAnisotropic,
            XExt = checked((int)Math.Max(
                1,
                Math.Round(widthPt * 2540d / 72d))),
            YExt = checked((int)Math.Max(
                1,
                Math.Round(heightPt * 2540d / 72d))),
            Metafile = IntPtr.Zero,
        };
        var referenceDc = GetDC(IntPtr.Zero);
        IntPtr enhancedMetafile = IntPtr.Zero;
        try
        {
            enhancedMetafile = SetWinMetaFileBits(
                checked((uint)rawWmf.Length),
                rawWmf,
                referenceDc,
                ref metafilePicture);
            if (enhancedMetafile == IntPtr.Zero)
                throw new InvalidDataException(
                    "Windows could not convert the MathType native WMF to an enhanced metafile.");
            var byteCount = GetEnhMetaFileBits(enhancedMetafile, 0, null);
            if (byteCount == 0 || byteCount > 64 * 1024 * 1024)
                throw new InvalidDataException(
                    $"Windows returned an invalid MathType enhanced-metafile size {byteCount}.");
            var bytes = new byte[byteCount];
            if (GetEnhMetaFileBits(enhancedMetafile, byteCount, bytes) != byteCount)
                throw new InvalidDataException(
                    "Windows returned an incomplete MathType enhanced metafile.");
            var path = Path.Combine(
                outputDirectory,
                $"mathtype-native-presentation-{Guid.NewGuid():N}.emf");
            File.WriteAllBytes(path, bytes);
            return path;
        }
        finally
        {
            if (enhancedMetafile != IntPtr.Zero) DeleteEnhMetaFile(enhancedMetafile);
            if (referenceDc != IntPtr.Zero) ReleaseDC(IntPtr.Zero, referenceDc);
        }
    }

    internal static byte[] ConvertEnhancedMetafileToPlaceableWmf(
        string emfPath,
        float widthPt,
        float heightPt)
    {
        if (string.IsNullOrWhiteSpace(emfPath) || !File.Exists(emfPath))
            throw new FileNotFoundException("EMF preview is unavailable.", emfPath);
        if (!(widthPt > 0) || !(heightPt > 0))
            throw new InvalidDataException(
                $"Invalid WMF preview size {widthPt}x{heightPt} pt.");

        var enhancedMetafile = GetEnhMetaFileW(emfPath);
        if (enhancedMetafile == IntPtr.Zero)
            throw new InvalidDataException($"Windows could not open EMF preview '{emfPath}'.");
        var referenceDc = GetDC(IntPtr.Zero);
        try
        {
            var rawLength = GetWinMetaFileBits(
                enhancedMetafile,
                0,
                null,
                MmAnisotropic,
                referenceDc);
            if (rawLength == 0 || rawLength > 64 * 1024 * 1024)
                throw new InvalidDataException(
                    $"Windows could not convert the EMF preview to WMF; size={rawLength}.");
            var rawWmf = new byte[rawLength];
            var written = GetWinMetaFileBits(
                enhancedMetafile,
                rawLength,
                rawWmf,
                MmAnisotropic,
                referenceDc);
            if (written != rawLength)
                throw new InvalidDataException(
                    $"Windows returned an incomplete WMF preview: {written}/{rawLength} bytes.");

            var right = checked((short)Math.Max(
                1,
                Math.Min(short.MaxValue, (int)Math.Round(widthPt * PlaceableInch / 72d))));
            var bottom = checked((short)Math.Max(
                1,
                Math.Min(short.MaxValue, (int)Math.Round(heightPt * PlaceableInch / 72d))));
            var placeable = BuildPlaceableHeader(right, bottom);
            var result = new byte[placeable.Length + rawWmf.Length];
            Buffer.BlockCopy(placeable, 0, result, 0, placeable.Length);
            Buffer.BlockCopy(rawWmf, 0, result, placeable.Length, rawWmf.Length);
            return result;
        }
        finally
        {
            if (referenceDc != IntPtr.Zero) ReleaseDC(IntPtr.Zero, referenceDc);
            DeleteEnhMetaFile(enhancedMetafile);
        }
    }

    private static byte[] BuildPlaceableHeader(short right, short bottom)
    {
        var bytes = new byte[22];
        WriteUInt32(bytes, 0, AldusPlaceableKey);
        WriteUInt16(bytes, 4, 0);
        WriteInt16(bytes, 6, 0);
        WriteInt16(bytes, 8, 0);
        WriteInt16(bytes, 10, right);
        WriteInt16(bytes, 12, bottom);
        WriteUInt16(bytes, 14, PlaceableInch);
        WriteUInt32(bytes, 16, 0);
        ushort checksum = 0;
        for (var offset = 0; offset < 20; offset += 2)
            checksum ^= BitConverter.ToUInt16(bytes, offset);
        WriteUInt16(bytes, 20, checksum);
        return bytes;
    }

    private static ObjectReferences ResolveObjectReferences(XDocument document)
    {
        var oleObject = document.Descendants(OfficeNamespace + "OLEObject").SingleOrDefault()
            ?? throw new InvalidDataException("Flat OPC does not contain exactly one OLEObject.");
        var progId = (string?)oleObject.Attribute("ProgID") ?? string.Empty;
        if (string.IsNullOrWhiteSpace(progId)
            || !(string.Equals(progId, "Equation", StringComparison.OrdinalIgnoreCase)
                || progId.StartsWith("Equation.", StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException(
                $"Flat OPC OLE object has unexpected ProgID '{progId}'.");
        var oleRelationshipId = (string?)oleObject.Attribute(OfficeRelationshipNamespace + "id")
            ?? throw new InvalidDataException("Flat OPC OLEObject has no relationship id.");

        var imageData = document.Descendants(VmlNamespace + "imagedata").SingleOrDefault()
            ?? throw new InvalidDataException("Flat OPC MathType object has no VML preview image.");
        var imageRelationshipId = (string?)imageData.Attribute(OfficeRelationshipNamespace + "id")
            ?? throw new InvalidDataException("Flat OPC preview image has no relationship id.");

        var relationshipsPart = FindPart(document, "/word/_rels/document.xml.rels");
        var relationships = relationshipsPart.Descendants(RelationshipNamespace + "Relationship").ToArray();
        var oleTarget = relationships
            .Single(relationship => string.Equals(
                (string?)relationship.Attribute("Id"),
                oleRelationshipId,
                StringComparison.Ordinal))
            .Attribute("Target")?.Value
            ?? throw new InvalidDataException("Flat OPC OLE relationship has no target.");
        var imageTarget = relationships
            .Single(relationship => string.Equals(
                (string?)relationship.Attribute("Id"),
                imageRelationshipId,
                StringComparison.Ordinal))
            .Attribute("Target")?.Value
            ?? throw new InvalidDataException("Flat OPC image relationship has no target.");

        return new ObjectReferences
        {
            ProgId = progId,
            EmbeddingPartName = ResolveWordPartName(oleTarget),
            PreviewPartName = ResolveWordPartName(imageTarget),
        };
    }

    private static string ResolveWordPartName(string target)
    {
        var normalized = target.Replace('\\', '/').TrimStart('/');
        if (normalized.StartsWith("word/", StringComparison.OrdinalIgnoreCase))
            return "/" + normalized;
        return "/word/" + normalized;
    }

    private static XElement FindPart(XDocument document, string partName)
    {
        var matches = document
            .Descendants(PackageNamespace + "part")
            .Where(part => string.Equals(
                (string?)part.Attribute(PackageNamespace + "name"),
                partName,
                StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (matches.Length != 1)
            throw new InvalidDataException(
                $"Expected exactly one Flat OPC part '{partName}', found {matches.Length}.");
        return matches[0];
    }

    private static byte[] ReadBinaryPart(XDocument document, string partName)
    {
        var part = FindPart(document, partName);
        var binary = part.Element(PackageNamespace + "binaryData")
            ?? throw new InvalidDataException($"Flat OPC part '{partName}' is not binary.");
        var encoded = new string(binary.Value.Where(character => !char.IsWhiteSpace(character)).ToArray());
        try { return Convert.FromBase64String(encoded); }
        catch (FormatException error)
        {
            throw new InvalidDataException(
                $"Flat OPC part '{partName}' contains invalid base64 data.",
                error);
        }
    }

    private static void WriteBinaryPart(
        XDocument document,
        string partName,
        byte[] bytes)
    {
        var part = FindPart(document, partName);
        var binary = part.Element(PackageNamespace + "binaryData")
            ?? throw new InvalidDataException($"Flat OPC part '{partName}' is not binary.");
        binary.Value = Convert.ToBase64String(bytes);
    }

    private static (float WidthPt, float HeightPt) ReadShapeSize(XDocument document)
    {
        var shape = document.Descendants(VmlNamespace + "shape").SingleOrDefault(element =>
            string.Equals(
                (string?)element.Attribute(OfficeNamespace + "ole"),
                string.Empty,
                StringComparison.Ordinal)
            || element.Attribute(OfficeNamespace + "ole") is not null)
            ?? document.Descendants(VmlNamespace + "shape").LastOrDefault()
            ?? throw new InvalidDataException("Flat OPC MathType object has no VML shape.");
        var style = (string?)shape.Attribute("style") ?? string.Empty;
        return (
            ReadStylePoints(style, "width"),
            ReadStylePoints(style, "height"));
    }

    private static float ReadStylePoints(string style, string property)
    {
        foreach (var segment in style.Split(';'))
        {
            var parts = segment.Split(new[] { ':' }, 2);
            if (parts.Length != 2
                || !string.Equals(parts[0].Trim(), property, StringComparison.OrdinalIgnoreCase))
                continue;
            var value = parts[1].Trim();
            if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
                value = value.Substring(0, value.Length - 2);
            if (float.TryParse(
                    value,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var points)
                && points > 0)
                return points;
        }
        throw new InvalidDataException(
            $"Flat OPC MathType VML shape has no valid {property} in '{style}'.");
    }

    private static void UpdateShapeSize(XDocument document, float widthPt, float heightPt)
    {
        var shape = document.Descendants(VmlNamespace + "shape").SingleOrDefault(element =>
            element.Attribute(OfficeNamespace + "ole") is not null)
            ?? document.Descendants(VmlNamespace + "shape").LastOrDefault()
            ?? throw new InvalidDataException("Flat OPC MathType object has no VML shape.");
        var style = (string?)shape.Attribute("style") ?? string.Empty;
        style = ReplaceStylePoints(style, "width", widthPt);
        style = ReplaceStylePoints(style, "height", heightPt);
        shape.SetAttributeValue("style", style);

        var wordObject = document.Descendants(WordNamespace + "object").SingleOrDefault();
        if (wordObject is not null)
        {
            wordObject.SetAttributeValue(
                WordNamespace + "dxaOrig",
                Math.Max(1, (int)Math.Round(widthPt * 20d)));
            wordObject.SetAttributeValue(
                WordNamespace + "dyaOrig",
                Math.Max(1, (int)Math.Round(heightPt * 20d)));
        }
    }

    private static string ReplaceStylePoints(string style, string property, float valuePt)
    {
        var replacement = valuePt.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + "pt";
        var segments = style.Split(';').ToList();
        var replaced = false;
        for (var index = 0; index < segments.Count; index++)
        {
            var parts = segments[index].Split(new[] { ':' }, 2);
            if (parts.Length != 2
                || !string.Equals(parts[0].Trim(), property, StringComparison.OrdinalIgnoreCase))
                continue;
            segments[index] = parts[0] + ":" + replacement;
            replaced = true;
            break;
        }
        if (!replaced) segments.Add(property + ":" + replacement);
        return string.Join(";", segments);
    }

    private static void WriteUInt16(byte[] bytes, int offset, ushort value)
    {
        var encoded = BitConverter.GetBytes(value);
        Buffer.BlockCopy(encoded, 0, bytes, offset, encoded.Length);
    }

    private static void WriteInt16(byte[] bytes, int offset, short value) =>
        WriteUInt16(bytes, offset, unchecked((ushort)value));

    private static void WriteUInt32(byte[] bytes, int offset, uint value)
    {
        var encoded = BitConverter.GetBytes(value);
        Buffer.BlockCopy(encoded, 0, bytes, offset, encoded.Length);
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.FinalReleaseComObject(value); }
        catch { }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MetafilePictureNative
    {
        internal int MappingMode;
        internal int XExt;
        internal int YExt;
        internal IntPtr Metafile;
    }

    private sealed class ObjectReferences
    {
        internal string ProgId { get; set; } = string.Empty;
        internal string EmbeddingPartName { get; set; } = string.Empty;
        internal string PreviewPartName { get; set; } = string.Empty;
    }

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr GetEnhMetaFileW(string fileName);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeleteEnhMetaFile(IntPtr enhancedMetafile);

    [DllImport("gdi32.dll")]
    private static extern IntPtr SetWinMetaFileBits(
        uint bufferSize,
        byte[] data,
        IntPtr referenceDc,
        ref MetafilePictureNative metafilePicture);

    [DllImport("gdi32.dll")]
    private static extern uint GetEnhMetaFileBits(
        IntPtr enhancedMetafile,
        uint bufferSize,
        [Out] byte[]? data);

    [DllImport("gdi32.dll")]
    private static extern uint GetWinMetaFileBits(
        IntPtr enhancedMetafile,
        uint bufferSize,
        [Out] byte[]? data,
        int mapMode,
        IntPtr referenceDc);

    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr window);

    [DllImport("user32.dll")]
    private static extern int ReleaseDC(IntPtr window, IntPtr dc);
}
