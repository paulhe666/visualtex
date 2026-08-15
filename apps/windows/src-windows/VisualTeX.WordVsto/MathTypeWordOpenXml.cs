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
