using System.Globalization;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
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
        for (var index = 0; index < xmlObjects.Length; index++)
        {
            var xmlObject = xmlObjects[index];
            var live = liveObjects[index];
            if (!string.Equals((string?)xmlObject.Attribute("ProgID"), live.ProgId, StringComparison.OrdinalIgnoreCase)
                || (string?)xmlObject.Attribute("Type") != (live.Type == 1 ? "Embed" : "Link"))
                throw new InvalidDataException("COM and XML disagree on the inline OLE type and order.");
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
