using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

// A verification checkpoint for a local Word undo transaction. It never restores
// a document by rebuilding its XML, and captures only the affected paragraph/row.
internal sealed class WordLocalEditSnapshot
{
    private readonly int start;
    private readonly int end;
    private readonly int documentEnd;
    private readonly string bodySignature;
    private readonly string mathFont;
    private readonly WordUndoHistorySnapshot undoHistory;
    private readonly WordBookmarkRecoverySnapshot bookmarkRecovery;

    internal WordLocalEditSnapshot(Document document, Range formulaRange, string formulaId)
    {
        var watch = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF") == "1"
            ? System.Diagnostics.Stopwatch.StartNew() : null;
        long checkpoint = 0;
        void Trace(string stage)
        {
            if (watch is null) return;
            var elapsed = watch.ElapsedMilliseconds;
            WordDoubleClickHook.TraceMessage($"local-checkpoint-perf stage={stage} deltaMs={elapsed - checkpoint} totalMs={elapsed}");
            checkpoint = elapsed;
        }
        Range? scope = null;
        Range? content = null;
        Tables? tables = null;
        Table? table = null;
        Rows? rows = null;
        Row? row = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        try
        {
            tables = formulaRange.Tables;
            if (tables.Count == 1)
            {
                table = tables[1];
                if (WordEquationNumbering.TryGetManagedNumberTableRowIndex(table, formulaRange, 2, out var rowIndex))
                {
                    rows = table.Rows;
                    row = rows[rowIndex];
                    scope = row.Range;
                }
            }
            if (scope is null)
            {
                // An ordinary user table is not a managed numbering host. Its
                // containing paragraph is still a valid local edit checkpoint.
                paragraphs = formulaRange.Paragraphs;
                paragraph = paragraphs[1];
                scope = paragraph.Range;
            }
            start = scope.Start;
            end = scope.End;
            content = document.Content;
            documentEnd = content.End;
            mathFont = document.OMathFontName;
            Trace("scope");
            var originalXml = scope.WordOpenXML;
            Trace("xml");
            var geometry = WordInlineObjectGeometry.Capture(scope);
            bodySignature = Signature(originalXml, geometry);
            Trace("signature");
            bookmarkRecovery = new WordBookmarkRecoverySnapshot(document, scope,
                NormalizedBody(originalXml, geometry), WordBookmarkRecoverySnapshot.NamesForFormula(formulaId));
            Trace("bookmarks");
            undoHistory = new WordUndoHistorySnapshot(document);
            Trace("undo-history");
        }
        finally
        {
            Release(scope);
            Release(content);
            Release(paragraph);
            Release(paragraphs);
            Release(row);
            Release(rows);
            Release(table);
            Release(tables);
        }
    }

    internal void VerifyRestored(Document document)
    {
        if (!Matches(document))
            throw new InvalidDataException("Word undo did not restore the formula row, numbering, bookmarks and formatting.");
    }

    internal bool Matches(Document document)
        => string.Equals(document.OMathFontName, mathFont, StringComparison.Ordinal)
            && ContentMatches(document);

    private bool ContentMatches(Document document)
    {
        Range? restored = null;
        Range? content = null;
        try
        {
            content = document.Content;
            if (content.End != documentEnd)
                return false;
            restored = document.Range(start, end);
            return string.Equals(Signature(restored.WordOpenXML, WordInlineObjectGeometry.Capture(restored)), bodySignature, StringComparison.Ordinal);
        }
        finally { Release(restored); Release(content); }
    }

    // The caller must end its owned custom undo record first. Custom XML and
    // document math-font preferences are restored explicitly; neither may justify
    // undoing a preceding user action when this edit made no content change.
    internal void RestoreOmml(Document document, string formulaId, FormulaMetadata? metadata)
    {
        undoHistory.UndoChanges(document);
        if (!ContentMatches(document))
        {
            Range? restoredScope = null;
            Range? content = null;
            try
            {
                content = document.Content;
                if (content.End != documentEnd)
                    throw new InvalidDataException("Word undo has not restored the original document extent.");
                restoredScope = document.Range(start, end);
                bookmarkRecovery.Restore(document,
                    NormalizedBody(restoredScope.WordOpenXML, WordInlineObjectGeometry.Capture(restoredScope)));
            }
            finally { Release(restoredScope); Release(content); }
        }
        if (!string.Equals(document.OMathFontName, mathFont, StringComparison.Ordinal))
            document.OMathFontName = mathFont;
        WordOmmlFormulaStore.InvalidateDocumentCache(document);
        if (metadata is not null) WordOmmlFormulaStore.Save(document, metadata);
        else WordOmmlFormulaStore.Delete(document, formulaId);
        VerifyRestored(document);
        var restored = WordOmmlFormulaStore.TryRead(document, formulaId);
        if ((restored is null) != (metadata is null)
            || (restored is not null && metadata is not null
                && !string.Equals(FormulaMetadataCodec.Encode(restored), FormulaMetadataCodec.Encode(metadata), StringComparison.Ordinal)))
            throw new InvalidDataException("Formula metadata was not restored after the failed edit.");
    }

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
                    var bytes = binary is not null ? Convert.FromBase64String(binary.Value)
                        : Encoding.UTF8.GetBytes(part.Element(pkg + "xmlData")!.ToString(SaveOptions.DisableFormatting));
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

    private static void Release(object? value)
    {
        if (value is not null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            System.Runtime.InteropServices.Marshal.ReleaseComObject(value);
    }
}
