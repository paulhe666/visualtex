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
    private const int MaximumFlatOpcCharacters = 192 * 1024 * 1024;

    private static readonly XNamespace PackageNamespace =
        "http://schemas.microsoft.com/office/2006/xmlPackage";
    private static readonly XNamespace RelationshipNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace OfficeRelationshipNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace OfficeNamespace =
        "urn:schemas-microsoft-com:office:office";
    private const string EquationBookmarkPrefix = "ZEqnNum";
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

    internal sealed class OleSnapshot
    {
        internal string ProgId { get; set; } = string.Empty;
        internal byte[] CompoundFile { get; set; } = Array.Empty<byte>();
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

    /// <summary>
    /// Maps VisualTeX's five document numbering presets onto MathType's own
    /// MTChap/MTSec/MTEqn field model. The resulting MTPlaceRef is still a native
    /// MathType number: MathType's Update/Insert Reference commands continue to
    /// understand it and VisualTeX does not introduce a parallel sequence.
    /// </summary>
    internal static NumberTemplate CreateVisualTeXNumberTemplate(string? formatId)
    {
        var format = EquationNumberFormat.Resolve(formatId);
        var template = new NumberTemplate();
        template.Segments.Add(NumberSegment.Text(" \\* MERGEFORMAT "));
        template.Segments.Add(NumberSegment.Field(" SEQ MTEqn \\h \\* MERGEFORMAT "));
        template.Segments.Add(NumberSegment.Text("("));
        if (format.HeadingLevel >= 1)
        {
            template.Segments.Add(NumberSegment.Field(
                " SEQ MTChap \\c \\* Arabic \\* MERGEFORMAT "));
            // Keep VisualTeX's document-wide preset identical across OMML,
            // VisualTeX OLE and MathType-native fields. At heading level 2 the
            // chapter/section delimiter remains '.', while the configured
            // separator is used immediately before MTEqn (for example 1.1-1).
            template.Segments.Add(NumberSegment.Text(
                format.HeadingLevel == 1 ? format.Separator : "."));
        }
        if (format.HeadingLevel >= 2)
        {
            template.Segments.Add(NumberSegment.Field(
                " SEQ MTSec \\c \\* Arabic \\* MERGEFORMAT "));
            template.Segments.Add(NumberSegment.Text(format.Separator));
        }
        template.Segments.Add(NumberSegment.Field(
            " SEQ MTEqn \\c \\* Arabic \\* MERGEFORMAT "));
        template.Segments.Add(NumberSegment.Text(")"));
        return template;
    }

    /// <summary>
    /// Rewrites one exact MTPlaceRef field-range Flat OPC fragment using the
    /// topology emitted by MathType's own Word commands. The caller deliberately
    /// supplies only the outer field span, so tabs, the Equation.DSMT4 object,
    /// paragraph marks and MTEditEquationSection2 state remain outside this
    /// transaction and cannot be moved by the rewrite.
    /// </summary>
    internal static string RewriteMathTypePlaceRefFieldFlatOpc(
        string sourceFlatOpc,
        NumberTemplate numberingTemplate)
    {
        if (string.IsNullOrWhiteSpace(sourceFlatOpc))
            throw new ArgumentException(
                "MathType MTPlaceRef source Flat OPC is empty.",
                nameof(sourceFlatOpc));
        if (numberingTemplate is null || numberingTemplate.Segments.Count == 0)
            throw new InvalidDataException(
                "MathType MTPlaceRef numbering template is empty.");
        EnsureFlatOpcSize(sourceFlatOpc, "MathType MTPlaceRef field package");

        var package = XDocument.Parse(
            sourceFlatOpc,
            LoadOptions.PreserveWhitespace);
        var documentPart = FindPart(package, "/word/document.xml");
        var xmlData = documentPart.Element(PackageNamespace + "xmlData")
            ?? throw new InvalidDataException(
                "MathType MTPlaceRef Flat OPC document part has no xmlData.");
        var wordDocument = xmlData.Elements().SingleOrDefault()
            ?? throw new InvalidDataException(
                "MathType MTPlaceRef Flat OPC document part is empty.");

        var ownerParagraphs = wordDocument
            .Descendants(WordNamespace + "p")
            .Where(paragraph => string.Concat(
                    paragraph
                        .Descendants(WordNamespace + "instrText")
                        .Select(node => node.Value))
                .IndexOf(
                    "MACROBUTTON MTPlaceRef",
                    StringComparison.OrdinalIgnoreCase) >= 0)
            .ToList();
        if (ownerParagraphs.Count != 1)
            throw new InvalidDataException(
                $"MathType MTPlaceRef field package must contain exactly one owner paragraph; found {ownerParagraphs.Count}.");

        var owner = ownerParagraphs[0];
        if (owner.Descendants(WordNamespace + "object").Any()
            || owner.Descendants(WordNamespace + "tab").Any())
            throw new InvalidDataException(
                "MathType MTPlaceRef field package unexpectedly includes an OLE object or tab outside the field boundary.");

        var directChildren = owner.Elements().ToList();
        var unsupportedNodes = directChildren
            .Where(node => node.Name != WordNamespace + "pPr"
                && node.Name != WordNamespace + "r"
                && node.Name != WordNamespace + "bookmarkStart"
                && node.Name != WordNamespace + "bookmarkEnd")
            .Select(node => node.Name.LocalName)
            .Distinct(StringComparer.Ordinal)
            .ToList();
        if (unsupportedNodes.Count > 0)
            throw new InvalidDataException(
                "MathType MTPlaceRef field package contains unsupported inline nodes: "
                + string.Join(", ", unsupportedNodes));

        var bookmarkStarts = directChildren
            .Where(node => node.Name == WordNamespace + "bookmarkStart")
            .ToList();
        var nonMathTypeBookmarks = bookmarkStarts
            .Where(node => !(node.Attribute(WordNamespace + "name")?.Value
                    .StartsWith(
                        EquationBookmarkPrefix,
                        StringComparison.OrdinalIgnoreCase)
                ?? false))
            .Select(node => node.Attribute(WordNamespace + "name")?.Value
                ?? "<unnamed>")
            .ToList();
        if (nonMathTypeBookmarks.Count > 0)
            throw new InvalidDataException(
                "MathType MTPlaceRef field package contains non-MathType bookmarks: "
                + string.Join(", ", nonMathTypeBookmarks));

        var bookmarkIds = bookmarkStarts
            .Select(node => node.Attribute(WordNamespace + "id")?.Value
                ?? throw new InvalidDataException(
                    "MathType ZEqnNum bookmark has no w:id."))
            .ToList();
        if (bookmarkIds.Count != bookmarkIds.Distinct(StringComparer.Ordinal).Count())
            throw new InvalidDataException(
                "MathType MTPlaceRef field contains duplicate bookmark ids.");

        var bookmarkEndsById = directChildren
            .Where(node => node.Name == WordNamespace + "bookmarkEnd")
            .GroupBy(
                node => node.Attribute(WordNamespace + "id")?.Value
                    ?? string.Empty,
                StringComparer.Ordinal)
            .ToDictionary(
                group => group.Key,
                group => group.ToList(),
                StringComparer.Ordinal);
        foreach (var bookmarkId in bookmarkIds)
        {
            if (!bookmarkEndsById.TryGetValue(bookmarkId, out var ends)
                || ends.Count != 1)
                throw new InvalidDataException(
                    $"MathType ZEqnNum bookmark id '{bookmarkId}' does not have exactly one matching end.");
        }
        var unmatchedBookmarkEnds = bookmarkEndsById.Keys
            .Where(id => !bookmarkIds.Contains(id, StringComparer.Ordinal))
            .ToList();
        if (unmatchedBookmarkEnds.Count > 0)
            throw new InvalidDataException(
                "MathType MTPlaceRef field contains unmatched bookmark ends: "
                + string.Join(", ", unmatchedBookmarkEnds));

        var paragraphProperties = directChildren
            .Where(node => node.Name == WordNamespace + "pPr")
            .Select(node => new XElement(node))
            .ToList();
        var bookmarkStartClones = bookmarkStarts
            .Select(node => new XElement(node))
            .ToList();
        var bookmarkEndClones = bookmarkIds
            .AsEnumerable()
            .Reverse()
            .Select(id => new XElement(bookmarkEndsById[id][0]))
            .ToList();

        var rebuiltField = BuildMathTypePlaceRef(
                numberingTemplate,
                bookmarkStartClones,
                bookmarkEndClones)
            .ToList();
        owner.RemoveNodes();
        owner.Add(paragraphProperties);
        owner.Add(rebuiltField);
        return package.ToString(SaveOptions.DisableFormatting);
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

    internal static IReadOnlyList<OleSnapshot> ReadOleSnapshots(Word.Range range)
    {
        if (range is null) throw new ArgumentNullException(nameof(range));
        return ReadOleSnapshots(range.WordOpenXML);
    }

    internal static IReadOnlyList<OleSnapshot> ReadOleSnapshots(string wordOpenXml)
    {
        if (string.IsNullOrWhiteSpace(wordOpenXml))
            return Array.Empty<OleSnapshot>();
        EnsureFlatOpcSize(wordOpenXml, "Word OLE snapshot package");

        var package = XDocument.Parse(wordOpenXml, LoadOptions.PreserveWhitespace);
        var documentPart = FindPart(package, "/word/document.xml");
        var relationshipsPart = FindPart(package, "/word/_rels/document.xml.rels");
        var relationships = relationshipsPart
            .Descendants(RelationshipNamespace + "Relationship")
            .Where(item => item.Attribute("Id") is not null)
            .ToDictionary(
                item => item.Attribute("Id")!.Value,
                item => item,
                StringComparer.Ordinal);
        var parts = package
            .Descendants(PackageNamespace + "part")
            .Where(part => part.Attribute(PackageNamespace + "name") is not null)
            .GroupBy(
                part => part.Attribute(PackageNamespace + "name")!.Value,
                StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                group => group.Key,
                group => group.First(),
                StringComparer.OrdinalIgnoreCase);

        var result = new List<OleSnapshot>();
        foreach (var oleObject in documentPart.Descendants(OfficeNamespace + "OLEObject"))
        {
            var snapshot = new OleSnapshot
            {
                ProgId = (string?)oleObject.Attribute("ProgID") ?? string.Empty,
            };
            var relationshipId =
                (string?)oleObject.Attribute(OfficeRelationshipNamespace + "id");
            if (!string.IsNullOrWhiteSpace(relationshipId)
                && relationships.TryGetValue(relationshipId!, out var relationship)
                && !string.Equals(
                    (string?)relationship.Attribute("TargetMode"),
                    "External",
                    StringComparison.OrdinalIgnoreCase))
            {
                var target = (string?)relationship.Attribute("Target");
                if (!string.IsNullOrWhiteSpace(target))
                {
                    var partName = ResolveWordPartName(target!);
                    if (parts.TryGetValue(partName, out var part))
                    {
                        var binary = part.Element(PackageNamespace + "binaryData");
                        if (binary is not null)
                        {
                            var encoded = new string(
                                binary.Value
                                    .Where(character => !char.IsWhiteSpace(character))
                                    .ToArray());
                            try
                            {
                                snapshot.CompoundFile = DecodeCompoundFileBase64(
                                    encoded,
                                    "Flat OPC OLE snapshot");
                            }
                            catch (InvalidDataException)
                            {
                                snapshot.CompoundFile = Array.Empty<byte>();
                            }
                        }
                    }
                }
            }
            result.Add(snapshot);
        }
        return result;
    }

    internal static Fragment Read(string wordOpenXml)
    {
        if (string.IsNullOrWhiteSpace(wordOpenXml))
            throw new InvalidDataException("Word returned empty Flat OPC for the MathType equation.");
        EnsureFlatOpcSize(wordOpenXml, "MathType Flat OPC package");

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
        NumberTemplate? numberTemplate = null,
        string mathTypeNumberPosition = "right")
    {
        if (!MathTypeOleStorage.LooksLikeMathTypeCompoundFile(compoundFile))
            throw new InvalidDataException(
                "MathType Flat OPC creation requires a valid Equation.DSMT4 Compound File.");
        if (previewWmf is null || previewWmf.Length <= 22)
            throw new InvalidDataException("MathType Flat OPC creation requires a valid WMF preview.");
        if (!(widthPt > 0) || !(heightPt > 0))
            throw new InvalidDataException(
                $"Invalid MathType preview size {widthPt}x{heightPt} pt.");
        var storageIdentity = MathTypeOleStorage.ReadCompoundFileIdentity(compoundFile);
        var oleProgId = string.IsNullOrWhiteSpace(storageIdentity.ProgId)
            ? MathTypeOleInterop.CanonicalProgId
            : storageIdentity.ProgId;

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
                new XAttribute(OfficeRelationshipNamespace + "id", imageRelationshipId),
                new XAttribute(OfficeNamespace + "title", string.Empty)));
        var oleObject = new XElement(
            OfficeNamespace + "OLEObject",
            new XAttribute("Type", "Embed"),
            new XAttribute("ProgID", oleProgId),
            new XAttribute("ShapeID", shapeId),
            new XAttribute("DrawAspect", "Content"),
            new XAttribute("ObjectID", objectId),
            new XAttribute(OfficeRelationshipNamespace + "id", oleRelationshipId));
        var wordObject = new XElement(
            WordNamespace + "object",
            new XAttribute(
                WordNamespace + "dxaOrig",
                MathTypeOriginalTwips(widthPt)),
            new XAttribute(
                WordNamespace + "dyaOrig",
                MathTypeOriginalTwips(heightPt)),
            shapeType,
            shape,
            oleObject);
        var paragraph = new XElement(WordNamespace + "p");
        if (numberTemplate is not null)
        {
            if (!display)
                throw new InvalidDataException(
                    "MathType equation numbering is valid only for display equations.");
            if (!string.Equals(mathTypeNumberPosition, "left", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(mathTypeNumberPosition, "right", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "MathType equation number position must be left or right.");
        }

        var numberOnLeft = numberTemplate is not null
            && string.Equals(mathTypeNumberPosition, "left", StringComparison.OrdinalIgnoreCase);
        if (numberOnLeft)
        {
            foreach (var node in BuildMathTypePlaceRef(numberTemplate!))
                paragraph.Add(node);
            paragraph.Add(new XElement(
                WordNamespace + "r",
                new XElement(WordNamespace + "tab")));
        }
        else if (display)
        {
            paragraph.Add(new XElement(
                WordNamespace + "r",
                new XElement(WordNamespace + "tab")));
        }

        paragraph.Add(new XElement(WordNamespace + "r", wordObject));
        if (numberTemplate is not null && !numberOnLeft)
        {
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
        if (!string.Equals(validation.ProgId, oleProgId, StringComparison.OrdinalIgnoreCase)
            || !validation.CompoundFile.SequenceEqual(compoundFile)
            || !validation.PreviewWmf.SequenceEqual(previewWmf))
            throw new InvalidDataException(
                "VisualTeX's standalone MathType Flat OPC failed self-validation.");
        return xmlText;
    }

    internal static string CreateDefaultSectionBreakFlatOpc(string displayLabel) =>
        CreateSectionBreakFlatOpc(displayLabel, chapter: 1, section: 1);

    internal static string CreateSectionBreakFlatOpc(
        string displayLabel,
        int chapter,
        int section)
    {
        if (string.IsNullOrWhiteSpace(displayLabel))
            throw new InvalidDataException("MathType chapter/section break label is empty.");
        if (chapter < 0 || section < 0)
            throw new ArgumentOutOfRangeException(
                nameof(chapter),
                "MathType chapter/section state cannot be negative.");

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
        foreach (var run in BuildMathTypeSectionBreak(displayLabel, chapter, section))
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

    private static IEnumerable<XElement> BuildMathTypeSectionBreak(
        string displayLabel,
        int chapter,
        int section)
    {
        yield return FieldCharRun("begin");
        yield return InstructionRun($" MACROBUTTON MTEditEquationSection2 {displayLabel}");
        foreach (var run in BuildSimpleComplexField(" SEQ MTEqn \\r \\h \\* MERGEFORMAT "))
            yield return run;
        foreach (var run in BuildSimpleComplexField(
                     $" SEQ MTSec \\r {section} \\h \\* MERGEFORMAT "))
            yield return run;
        foreach (var run in BuildSimpleComplexField(
                     $" SEQ MTChap \\r {chapter} \\h \\* MERGEFORMAT "))
            yield return run;
        yield return FieldCharRun("separate");
        yield return FieldCharRun("end");
    }

    private static IEnumerable<XElement> BuildMathTypePlaceRef(
        NumberTemplate template) =>
        BuildMathTypePlaceRef(
            template,
            Array.Empty<XElement>(),
            Array.Empty<XElement>());

    private static IEnumerable<XElement> BuildMathTypePlaceRef(
        NumberTemplate template,
        IReadOnlyList<XElement> bookmarkStarts,
        IReadOnlyList<XElement> bookmarkEnds)
    {
        if (template.Segments.Count == 0)
            throw new InvalidDataException("MathType MTPlaceRef numbering template is empty.");
        if (bookmarkStarts.Count != bookmarkEnds.Count)
            throw new InvalidDataException(
                "MathType MTPlaceRef bookmark start/end counts differ.");

        var nodes = new List<XElement>
        {
            FieldCharRun("begin"),
            // Native MathType keeps MERGEFORMAT in the outer instruction and has
            // no outer field separator/result. Everything through the closing
            // parenthesis is part of MTPlaceRef.Code.
            InstructionRun(" MACROBUTTON MTPlaceRef \\* MERGEFORMAT "),
        };
        var hiddenIncrementCount = 0;
        var visibleSequenceCount = 0;
        var bookmarksInserted = false;

        void InsertBookmarkStarts()
        {
            if (bookmarksInserted) return;
            nodes.AddRange(bookmarkStarts.Select(node => new XElement(node)));
            bookmarksInserted = true;
        }

        foreach (var segment in template.Segments)
        {
            if (!segment.IsField)
            {
                // Legacy in-memory templates keep the outer MERGEFORMAT switch as
                // their first literal. The exact native XML builder owns it above.
                if (IsOuterMergeFormatSegment(segment.Value)
                    || string.IsNullOrEmpty(segment.Value))
                    continue;
                if (hiddenIncrementCount > 0) InsertBookmarkStarts();
                nodes.Add(InstructionRun(segment.Value));
                continue;
            }

            var hidden = IsHiddenMathTypeSequenceInstruction(segment.Value);
            if (!hidden) InsertBookmarkStarts();
            nodes.AddRange(BuildNativeMathTypeSequenceField(
                segment.Value,
                hidden));
            if (hidden) hiddenIncrementCount++;
            else visibleSequenceCount++;
        }

        if (hiddenIncrementCount != 1)
            throw new InvalidDataException(
                $"MathType MTPlaceRef must own exactly one hidden MTEqn increment; found {hiddenIncrementCount}.");
        if (visibleSequenceCount == 0)
            throw new InvalidDataException(
                "MathType MTPlaceRef has no visible sequence field.");
        InsertBookmarkStarts();
        nodes.AddRange(bookmarkEnds.Select(node => new XElement(node)));
        nodes.Add(FieldCharRun("end"));
        return nodes;
    }

    private static IEnumerable<XElement> BuildNativeMathTypeSequenceField(
        string instruction,
        bool hidden)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            throw new InvalidDataException(
                "MathType nested Word field instruction is empty.");
        yield return FieldCharRun("begin");
        yield return InstructionRun(instruction);
        if (!hidden)
        {
            yield return FieldCharRun("separate");
            // The installed MathType add-in serializes visible SEQ results as
            // instrText with noProof. Word refreshes this placeholder immediately
            // from the live MTChap/MTSec/MTEqn state.
            yield return new XElement(
                WordNamespace + "r",
                new XElement(
                    WordNamespace + "rPr",
                    new XElement(WordNamespace + "noProof")),
                new XElement(WordNamespace + "instrText", "0"));
        }
        yield return FieldCharRun("end");
    }

    private static bool IsHiddenMathTypeSequenceInstruction(string instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction)) return false;
        var normalized = instruction
            .Replace('\t', ' ')
            .Replace('\r', ' ')
            .Replace('\n', ' ')
            .Trim();
        return normalized.StartsWith(
                "SEQ MTEqn ",
                StringComparison.OrdinalIgnoreCase)
            && normalized.IndexOf("\\h", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static bool IsOuterMergeFormatSegment(string? value) =>
        string.Equals(
            value?.Trim(),
            "\\* MERGEFORMAT",
            StringComparison.OrdinalIgnoreCase);

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

            // Genuine MathType WMFs use whole-point picture bounds (the same
            // dimensions later written to Word's dxaOrig/dyaOrig).  Preserve the
            // visual VML size separately, but make the embedded WMF's original
            // coordinate extent match MathType's native object format.
            var originalWidthPt = MathTypeOriginalPoints(widthPt);
            var originalHeightPt = MathTypeOriginalPoints(heightPt);
            var right = checked((short)Math.Max(
                1,
                Math.Min(short.MaxValue, (int)Math.Round(originalWidthPt * PlaceableInch / 72d))));
            var bottom = checked((short)Math.Max(
                1,
                Math.Min(short.MaxValue, (int)Math.Round(originalHeightPt * PlaceableInch / 72d))));
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
        if (string.IsNullOrWhiteSpace(progId))
            throw new InvalidDataException(
                "Flat OPC OLE object has no ProgID.");
        // Alternate MathType releases/registrations are validated by the embedded
        // CFB identity and MTEF payload in Read(), not by an Equation.* name prefix.
        // This keeps offline documents portable while still rejecting non-MathType
        // equation servers before any rewrite is accepted.
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

    private static void EnsureFlatOpcSize(string wordOpenXml, string context)
    {
        if (wordOpenXml.Length > MaximumFlatOpcCharacters)
            throw new InvalidDataException(
                $"{context} exceeds the supported safety limit of {MaximumFlatOpcCharacters} characters.");
    }

    private static byte[] DecodeCompoundFileBase64(
        string encoded,
        string context)
    {
        var maximumEncodedLength =
            ((long)MathTypeOleStorage.MaximumCompoundFileBytes + 2L) / 3L * 4L + 4096L;
        if (encoded.Length > maximumEncodedLength)
            throw new InvalidDataException(
                $"{context} exceeds the supported MathType OLE safety limit of {MathTypeOleStorage.MaximumCompoundFileBytes} bytes.");
        byte[] compoundFile;
        try { compoundFile = Convert.FromBase64String(encoded); }
        catch (FormatException error)
        {
            throw new InvalidDataException(
                $"{context} contains invalid base64 data.",
                error);
        }
        if (compoundFile.Length > MathTypeOleStorage.MaximumCompoundFileBytes)
            throw new InvalidDataException(
                $"{context} exceeds the supported MathType OLE safety limit of {MathTypeOleStorage.MaximumCompoundFileBytes} bytes.");
        return compoundFile;
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
            var pointsPerUnit = 1d;
            var units = new[]
            {
                (Suffix: "pt", Points: 1d),
                (Suffix: "in", Points: 72d),
                (Suffix: "pc", Points: 12d),
                (Suffix: "cm", Points: 72d / 2.54d),
                (Suffix: "mm", Points: 72d / 25.4d),
                // VML/CSS pixels use the Office/CSS 96 dpi convention.
                (Suffix: "px", Points: 72d / 96d),
            };
            foreach (var unit in units)
            {
                if (!value.EndsWith(unit.Suffix, StringComparison.OrdinalIgnoreCase))
                    continue;
                value = value.Substring(0, value.Length - unit.Suffix.Length).Trim();
                pointsPerUnit = unit.Points;
                break;
            }

            if (double.TryParse(
                    value,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var numericValue)
                && numericValue > 0)
            {
                var points = numericValue * pointsPerUnit;
                if (points > 0 && points <= float.MaxValue)
                    return (float)points;
            }
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
                MathTypeOriginalTwips(widthPt));
            wordObject.SetAttributeValue(
                WordNamespace + "dyaOrig",
                MathTypeOriginalTwips(heightPt));
        }
    }

    private static int MathTypeOriginalPoints(float valuePt) =>
        Math.Max(1, (int)Math.Round(valuePt, MidpointRounding.AwayFromZero));

    private static int MathTypeOriginalTwips(float valuePt) =>
        checked(MathTypeOriginalPoints(valuePt) * 20);

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
