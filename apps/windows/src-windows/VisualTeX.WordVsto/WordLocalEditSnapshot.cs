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
        var body = WordDocumentXml.GetRequiredMainDocumentBody(document);
        if (geometry is not null) WordInlineObjectGeometry.BindToEvidence(body, geometry);
        XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace v = "urn:schemas-microsoft-com:vml";
        XNamespace o = "urn:schemas-microsoft-com:office:office";
        // This normalizer is total: both sides of a comparison go through the
        // same code, so unusual packages (third-party writers, WPS, duplicate
        // ids) get a deterministic placeholder instead of refusing the edit.
        var parts = document.Descendants(pkg + "part")
            .Where(p => p.Attribute(pkg + "name") is not null)
            .GroupBy(p => (string)p.Attribute(pkg + "name")!, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal);
        if (parts.TryGetValue("/word/_rels/document.xml.rels", out var relationships))
        {
            var targets = relationships.Descendants(rel + "Relationship")
                .Where(e => e.Attribute("Id") is not null)
                .GroupBy(e => (string)e.Attribute("Id")!, StringComparer.Ordinal)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal);
            foreach (var reference in body.DescendantsAndSelf().Attributes()
                         .Where(a => a.Name.Namespace == r).ToArray())
            {
                if (!targets.TryGetValue(reference.Value, out var relationship))
                {
                    reference.Value = "unresolved";
                    continue;
                }
                var target = (string?)relationship.Attribute("Target") ?? string.Empty;
                var mode = (string?)relationship.Attribute("TargetMode");
                var type = (string?)relationship.Attribute("Type");
                if (mode == "External") reference.Value = type + ":external:" + target;
                else
                {
                    var path = Uri.TryCreate(new Uri("http://word/word/document.xml"), target, out var resolved)
                        ? resolved.AbsolutePath : target;
                    if (!parts.TryGetValue(path, out var part))
                    {
                        reference.Value = type + ":missing:" + target;
                        continue;
                    }
                    var binary = part.Element(pkg + "binaryData");
                    var xmlPayload = part.Element(pkg + "xmlData");
                    var normalizedPayload = xmlPayload is null ? string.Empty
                        : type is not null && (type.EndsWith("/header", StringComparison.Ordinal)
                            || type.EndsWith("/footer", StringComparison.Ordinal))
                            ? WordRecoveryXmlEquivalence.NormalizeHeaderFooterPayload(xmlPayload)
                            : xmlPayload.ToString(SaveOptions.DisableFormatting);
                    var bytes = binary is not null ? Convert.FromBase64String(binary.Value)
                        : Encoding.UTF8.GetBytes(normalizedPayload);
                    // An embedded OLE object is compared by its streams, not by
                    // its container layout (WPS rewrites it after an undo).
                    if (binary is not null
                        && OleCompoundFileDigest.TryCompute(bytes) is { } oleContent)
                    {
                        reference.Value = type + ":ole-content:" + oleContent;
                        continue;
                    }
                    using var payloadHash = SHA256.Create();
                    reference.Value = type + ":sha256:" + Convert.ToBase64String(payloadHash.ComputeHash(bytes));
                }
            }
        }
        // Native Undo can allocate new VML/Object IDs and relationship numbers.
        // Bind the shape ID by document order and its references by exact payload
        // hashes above; keep geometry, object type, metadata and embedded bytes.
        var shapes = body.Descendants(v + "shape").ToArray();
        var shapeIds = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var index = 0; index < shapes.Length; index++)
        {
            var shape = shapes[index];
            var id = (string?)shape.Attribute("id");
            var canonical = "shape:" + index;
            // Word can reallocate o:spid independently of v:shape/@id during
            // native Undo. Both names identify this SAME physical shape. Bind
            // their complete alias graph before normalizing either declaration
            // or an OLE reference; never erase a real owner/payload mismatch.
            foreach (var alias in new[] { id, (string?)shape.Attribute(o + "spid") })
            {
                // A pasted shape can repeat another's id; the first owner
                // keeps the alias and both documents bind it the same way.
                if (string.IsNullOrEmpty(alias) || shapeIds.ContainsKey(alias!)) continue;
                shapeIds[alias!] = canonical;
            }
        }
        for (var index = 0; index < shapes.Length; index++)
        {
            shapes[index].SetAttributeValue("id", "shape:" + index);
            if (shapes[index].Attribute(o + "spid") is not null)
                shapes[index].SetAttributeValue(o + "spid", "shape:" + index);
        }
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
            .Where(e => e.Attribute(w + "id") is not null)
            .GroupBy(e => (string)e.Attribute(w + "id")!, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => (string?)g.First().Attribute(w + "name") ?? string.Empty, StringComparer.Ordinal);
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
