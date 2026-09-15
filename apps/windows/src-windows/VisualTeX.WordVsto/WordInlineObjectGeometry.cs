using System.Globalization;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

// Read-only companion to OOXML recovery evidence. Word can serialize a newly
// inserted MathType object's rounded preview bounds until Undo/save materializes
// its actual twip dimensions. Compare those actual dimensions, not two encodings
// of the same geometry. The resulting XML is only hashed; never written to Word.
internal static class WordInlineObjectGeometry
{
    internal sealed class Item
    {
        internal int Start { get; set; }
        internal int End { get; set; }
        internal int Type { get; set; }
        internal string ProgId { get; set; } = string.Empty;
        internal float Width { get; set; }
        internal float Height { get; set; }
    }

    // An OLE selection can expose a transient presentation extent through COM.
    // Use the stored Word layout when that reading disagrees with the last host
    // checkpoint; otherwise opening an editor can turn a 12 pt formula into
    // 10.5 pt without a user resize. This is target-family geometry, not a
    // document/formula-specific baseline correction.
    internal static (float Width, float Height) ReadStableVisualTeXSize(
        InlineShape shape, FormulaMetadata metadata)
    {
        var width = shape.Width;
        var height = shape.Height;
        if (!string.Equals(metadata.DisplayMode, "inline", StringComparison.OrdinalIgnoreCase))
            return (width, height);
        if (metadata.WordInlineOleWidthPt is > 0 && metadata.WordInlineOleHeightPt is > 0
            && Math.Abs(width - metadata.WordInlineOleWidthPt.Value) <= 0.25
            && Math.Abs(height - metadata.WordInlineOleHeightPt.Value) <= 0.25)
            return ((float)metadata.WordInlineOleWidthPt.Value, (float)metadata.WordInlineOleHeightPt.Value);
        Range? range = null;
        try
        {
            range = shape.Range;
            if (TryReadSerializedVisualTeXSize(range.WordOpenXML, metadata.FormulaId,
                    out var serializedWidth, out var serializedHeight))
                return (serializedWidth, serializedHeight);
        }
        catch (COMException) { }
        catch (System.Xml.XmlException) { }
        finally { Release(range); }
        return (width, height);
    }

    internal static bool TryReadSerializedVisualTeXSize(
        string xml, string formulaId, out float width, out float height)
    {
        width = height = 0;
        if (string.IsNullOrWhiteSpace(xml) || !Guid.TryParse(formulaId, out _)) return false;
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace o = "urn:schemas-microsoft-com:office:office";
        XNamespace v = "urn:schemas-microsoft-com:vml";
        var document = XDocument.Parse(xml, LoadOptions.None);
        var found = false;
        foreach (var ole in document.Descendants(o + "OLEObject"))
        {
            if (ole.Parent?.Name != w + "object"
                || !string.Equals((string?)ole.Attribute("ProgID"), FormulaOleContract.ProgId,
                    StringComparison.OrdinalIgnoreCase)) continue;
            var shapeId = (string?)ole.Attribute("ShapeID");
            if (shapeId is null) continue;
            var shapes = ole.Parent.Elements(v + "shape")
                .Where(e => (string?)e.Attribute("id") == shapeId).ToArray();
            if (shapes.Length != 1) return false;
            var shape = shapes[0];
            var cached = FormulaMetadataCodec.Decode((string?)shape.Attribute("alt"))
                ?? FormulaMetadataCodec.Decode((string?)shape.Attribute("title"));
            if (!string.Equals(cached?.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase)) continue;
            // Duplicate identity is not a license to choose the nearest object.
            if (found) { width = height = 0; return false; }
            float? parsedWidth = null, parsedHeight = null;
            foreach (var declaration in ((string?)shape.Attribute("style") ?? string.Empty).Split(';'))
            {
                var colon = declaration.IndexOf(':');
                if (colon <= 0) continue;
                var key = declaration.Substring(0, colon).Trim();
                if (key != "width" && key != "height") continue;
                var value = declaration.Substring(colon + 1).Trim();
                if (!value.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                    || !float.TryParse(value.Substring(0, value.Length - 2), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var points)
                    || points <= 0 || float.IsNaN(points) || float.IsInfinity(points)) return false;
                if (key == "width") { if (parsedWidth.HasValue) return false; parsedWidth = points; }
                else { if (parsedHeight.HasValue) return false; parsedHeight = points; }
            }
            if (!parsedWidth.HasValue || !parsedHeight.HasValue) return false;
            width = parsedWidth.Value; height = parsedHeight.Value; found = true;
        }
        return found;
    }

    internal static IReadOnlyList<Item> Capture(Range scope)
    {
        InlineShapes? shapes = null;
        var result = new List<Item>();
        try
        {
            shapes = scope.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                OLEFormat? ole = null;
                try
                {
                    shape = shapes[index];
                    range = shape.Range;
                    if (range.StoryType != scope.StoryType || range.Start < scope.Start || range.End > scope.End)
                        throw new InvalidDataException("An inline object is outside its recovery scope.");
                    var type = shape.Type;
                    var progId = string.Empty;
                    if (type is WdInlineShapeType.wdInlineShapeEmbeddedOLEObject or WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                    {
                        ole = shape.OLEFormat;
                        progId = ole.ProgID;
                    }
                    result.Add(new Item { Start = range.Start - scope.Start, End = range.End - scope.Start,
                        Type = (int)type, ProgId = progId, Width = shape.Width, Height = shape.Height });
                }
                finally { Release(ole); Release(range); Release(shape); }
            }
            return result;
        }
        finally { Release(shapes); }
    }

    internal static void BindToEvidence(XElement body, IReadOnlyList<Item> actual)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace o = "urn:schemas-microsoft-com:office:office";
        XNamespace v = "urn:schemas-microsoft-com:vml";
        var xmlObjects = body.Descendants(o + "OLEObject")
            .Where(e => e.Parent?.Name == w + "object" && !e.Ancestors(w + "txbxContent").Any()).ToArray();
        var liveObjects = actual.Where(e => e.Type is 1 or 2).ToArray();
        if (xmlObjects.Length != liveObjects.Length)
            throw new InvalidDataException("COM and XML disagree on inline OLE recovery ownership.");

        double GeometryScore(XElement xmlObject, Item live)
        {
            var shapeId = (string?)xmlObject.Attribute("ShapeID");
            var shape = shapeId is null
                ? null
                : xmlObject.Parent?.Elements(v + "shape")
                    .FirstOrDefault(e => (string?)e.Attribute("id") == shapeId);
            var style = (string?)shape?.Attribute("style");
            if (string.IsNullOrWhiteSpace(style)) return double.MaxValue;
            double? width = null;
            double? height = null;
            foreach (var declaration in style.Split(';'))
            {
                var separator = declaration.IndexOf(':');
                if (separator <= 0) continue;
                var name = declaration.Substring(0, separator).Trim();
                var value = declaration.Substring(separator + 1).Trim();
                if (!value.EndsWith("pt", StringComparison.Ordinal)
                    || !double.TryParse(
                        value.Substring(0, value.Length - 2),
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        out var points))
                    continue;
                if (name == "width") width = points;
                else if (name == "height") height = points;
            }
            if (!width.HasValue || !height.HasValue) return double.MaxValue;
            return Math.Abs(width.Value - live.Width)
                + Math.Abs(height.Value - live.Height);
        }

        // Word's COM InlineShapes collection and Flat OPC can legally enumerate
        // different OLE families in a different cross-type order (especially in
        // compatibility-mode documents with legacy Equation.DSMT4 objects). The
        // old positional zip rejected such documents even when both views exposed
        // exactly the same physical OLE set. Bind by object identity signature and,
        // for repeated ProgIDs, by the closest serialized geometry instead.
        var unmatchedXmlObjects = xmlObjects.ToList();
        foreach (var live in liveObjects)
        {
            var expectedType = live.Type == 1 ? "Embed" : "Link";
            var candidates = unmatchedXmlObjects
                .Where(xmlObject =>
                    string.Equals(
                        (string?)xmlObject.Attribute("ProgID"),
                        live.ProgId,
                        StringComparison.OrdinalIgnoreCase)
                    && string.Equals(
                        (string?)xmlObject.Attribute("Type"),
                        expectedType,
                        StringComparison.Ordinal))
                .OrderBy(xmlObject => GeometryScore(xmlObject, live))
                .ToArray();
            if (candidates.Length == 0)
                throw new InvalidDataException(
                    "COM and XML disagree on the inline OLE type or ownership.");
            var xmlObject = candidates[0];
            unmatchedXmlObjects.Remove(xmlObject);
            if (!live.ProgId.StartsWith("Equation.DSMT", StringComparison.OrdinalIgnoreCase)) continue;
            var shapeId = (string?)xmlObject.Attribute("ShapeID");
            var shapes = xmlObject.Parent!.Elements(v + "shape")
                .Where(e => (string?)e.Attribute("id") == shapeId).ToArray();
            if (shapeId is null || shapes.Length != 1)
                throw new InvalidDataException("MathType recovery geometry has no unique XML shape owner.");
            var style = shapes[0].Attribute("style")
                ?? throw new InvalidDataException("MathType recovery geometry is missing its XML bounds.");
            var declarations = style.Value.Split(';');
            foreach (var dimension in new[] { (Name: "width", Value: live.Width), (Name: "height", Value: live.Height) })
            {
                var matches = Enumerable.Range(0, declarations.Length)
                    .Where(i => declarations[i].Split(':')[0].Trim() == dimension.Name).ToArray();
                if (matches.Length != 1)
                    throw new InvalidDataException("MathType XML bounds are ambiguous.");
                var declaration = declarations[matches[0]];
                var value = declaration.Substring(declaration.IndexOf(':') + 1).Trim();
                if (!value.EndsWith("pt", StringComparison.Ordinal)
                    || !double.TryParse(value.Substring(0, value.Length - 2), NumberStyles.Float, CultureInfo.InvariantCulture, out var points))
                    continue; // Other encodings retain their strict XML comparison.
                var livePoints = (double)dimension.Value;
                if (livePoints <= 0 || double.IsNaN(livePoints) || double.IsInfinity(livePoints))
                    throw new InvalidDataException("Word returned invalid inline object dimensions.");
                // Only the observed integer-point preview encoding, or the actual
                // COM point value, is equivalent. Arbitrary nearby values remain
                // in the strict XML signature. Actual dimensions below are always
                // compared independently, including subpoint size changes.
                if (Math.Abs(points - livePoints) <= 0.0001 || points == Math.Round(livePoints))
                    declarations[matches[0]] = dimension.Name + ":" + dimension.Value.ToString("R", CultureInfo.InvariantCulture) + "pt";
            }
            style.Value = string.Join(";", declarations);
        }
        body.SetAttributeValue("recoveryInlineGeometry", string.Join("|", actual.Select(e =>
            string.Join(",", e.Start, e.End, e.Type, e.ProgId,
                e.Width.ToString("R", CultureInfo.InvariantCulture), e.Height.ToString("R", CultureInfo.InvariantCulture)))));
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
