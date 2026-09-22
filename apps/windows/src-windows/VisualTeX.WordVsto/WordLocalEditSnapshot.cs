using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace VisualTeX.WordVsto;

// Canonical Flat-OPC signature utility used only to compare Word recovery
// evidence. It has no formula identity, metadata, lookup, or rollback authority.
internal static class WordLocalEditSnapshot
{
    internal static string Signature(string xml, IReadOnlyList<WordInlineObjectGeometry.Item>? geometry = null)
    {
        using var hash = SHA256.Create();
        return Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(NormalizedBody(xml, geometry))));
    }

    internal static string NormalizedBody(string xml, IReadOnlyList<WordInlineObjectGeometry.Item>? geometry = null)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var document = XDocument.Parse(xml);
        var body = document.Descendants(w + "body").Single();
        if (geometry is not null) WordInlineObjectGeometry.BindToEvidence(body, geometry);
        XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace v = "urn:schemas-microsoft-com:vml";
        XNamespace o = "urn:schemas-microsoft-com:office:office";
        var parts = document.Descendants(pkg + "part")
            .ToDictionary(p => (string)p.Attribute(pkg + "name")!, p => p, StringComparer.Ordinal);
        if (parts.TryGetValue("/word/_rels/document.xml.rels", out var relationships))
        {
            var targets = relationships.Descendants(rel + "Relationship")
                .ToDictionary(e => (string)e.Attribute("Id")!, e => e, StringComparer.Ordinal);
            foreach (var reference in body.DescendantsAndSelf().Attributes()
                         .Where(a => a.Name.Namespace == r).ToArray())
            {
                if (!targets.TryGetValue(reference.Value, out var relationship))
                    throw new InvalidDataException("A Word content relationship cannot be resolved for recovery verification.");
                var target = (string)relationship.Attribute("Target")!;
                var mode = (string?)relationship.Attribute("TargetMode");
                var type = (string?)relationship.Attribute("Type");
                if (mode == "External") reference.Value = type + ":external:" + target;
                else
                {
                    var path = new Uri(new Uri("http://word/word/document.xml"), target).AbsolutePath;
                    if (!parts.TryGetValue(path, out var part))
                        throw new InvalidDataException("A Word content payload is missing from recovery evidence.");
                    var binary = part.Element(pkg + "binaryData");
                    var xmlPayload = part.Element(pkg + "xmlData");
                    if (binary is null && xmlPayload is null)
                        throw new InvalidDataException("A Word recovery relationship has neither binary nor XML payload.");
                    var normalizedPayload = xmlPayload is null ? string.Empty
                        : type is not null && (type.EndsWith("/header", StringComparison.Ordinal)
                            || type.EndsWith("/footer", StringComparison.Ordinal))
                            ? WordRecoveryXmlEquivalence.NormalizeHeaderFooterPayload(xmlPayload)
                            : xmlPayload.ToString(SaveOptions.DisableFormatting);
                    var bytes = binary is not null ? Convert.FromBase64String(binary.Value)
                        : Encoding.UTF8.GetBytes(normalizedPayload);
                    using var payloadHash = SHA256.Create();
                    reference.Value = type + ":sha256:" + Convert.ToBase64String(payloadHash.ComputeHash(bytes));
                }
            }
        }
        // Native Undo can allocate new VML/Object IDs and relationship numbers.
        // Bind the shape ID by document order and its references by exact payload
        // hashes above; keep geometry, object type, metadata and embedded bytes.
        var shapeIds = body.Descendants(v + "shape").Select((shape, index) => (shape, index))
            .ToDictionary(item => (string)item.shape.Attribute("id")!, item => "shape:" + item.index);
        foreach (var shape in body.Descendants(v + "shape"))
            shape.SetAttributeValue("id", shapeIds[(string)shape.Attribute("id")!]);
        foreach (var ole in body.Descendants(o + "OLEObject"))
        {
            var oldShape = (string?)ole.Attribute("ShapeID");
            if (oldShape is not null && shapeIds.TryGetValue(oldShape, out var stableShape))
                ole.SetAttributeValue("ShapeID", stableShape);
            ole.Attribute("ObjectID")?.Remove();
        }
        // Word's revision stamps and numeric bookmark IDs are serialization
        // details. Keep bookmark names, spans, field instructions/results, math,
        // run/paragraph formatting, and table/cell structure in the comparison.
        foreach (var attribute in body.DescendantsAndSelf().Attributes()
                     .Where(a => a.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)
                         || a.Name.LocalName == "paraId" || a.Name.LocalName == "textId").ToArray())
            attribute.Remove();
        var names = body.Descendants(w + "bookmarkStart")
            .ToDictionary(e => (string)e.Attribute(w + "id")!, e => (string)e.Attribute(w + "name")!);
        foreach (var bookmark in body.Descendants().Where(e => e.Name == w + "bookmarkStart" || e.Name == w + "bookmarkEnd"))
        {
            var id = bookmark.Attribute(w + "id");
            if (id is not null && names.TryGetValue(id.Value, out var name)) id.Value = name;
        }
        // Undo may reorder starts (or ends) at exactly the same XML position.
        // Canonicalize only consecutive markers of the same kind: names, spans,
        // enclosing paragraphs/runs and every intervening content node remain.
        foreach (var parent in body.DescendantsAndSelf().ToArray())
        {
            var siblings = parent.Nodes().ToArray();
            for (var index = 0; index < siblings.Length; index++)
            {
                if (siblings[index] is not XElement first
                    || (first.Name != w + "bookmarkStart" && first.Name != w + "bookmarkEnd")) continue;
                var group = new List<XElement> { first };
                while (index + 1 < siblings.Length && siblings[index + 1] is XElement next && next.Name == first.Name)
                { group.Add(next); index++; }
                if (group.Count < 2) continue;
                first.AddBeforeSelf(group.OrderBy(e => (string?)e.Attribute(w + "id"), StringComparer.Ordinal)
                    .Select(e => new XElement(e)).ToArray());
                foreach (var marker in group) marker.Remove();
            }
        }
        return body.ToString(SaveOptions.DisableFormatting);
    }
}
