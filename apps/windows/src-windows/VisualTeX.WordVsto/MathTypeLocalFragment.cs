using System.Xml.Linq;

namespace VisualTeX.WordVsto;

/// <summary>
/// The serialization envelope for replacing one existing native OLE. A clipped
/// Word object export omits the owning paragraph properties and boundary
/// bookmarks; it can also carry table/section scaffolding that is not the object.
/// Keep the object's inline content and parts, supply its real paragraph pPr,
/// and delimit the imported object with a transaction-only locator.
/// </summary>
internal static class MathTypeLocalFragment
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
    private static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";

    internal static XElement Body(XDocument package)
    {
        var part = package.Root?.Elements(Pkg + "part")
            .SingleOrDefault(p => (string?)p.Attribute(Pkg + "name") == "/word/document.xml");
        return part?.Element(Pkg + "xmlData")?.Element(W + "document")?.Element(W + "body")
            ?? throw new InvalidDataException("The local Word package has no main document body.");
    }

    internal static string Prepare(string objectXml, string ownerXml, string locatorName,
        int sourceOleOrdinal = 0, int expectedOwnerOleCount = -1)
    {
        if (string.IsNullOrWhiteSpace(locatorName) || locatorName.Length > 40
            || !char.IsLetter(locatorName[0])
            || locatorName.Any(c => !char.IsLetterOrDigit(c) && c != '_'))
            throw new ArgumentException("Invalid local OLE locator name.", nameof(locatorName));

        var package = XDocument.Parse(objectXml, LoadOptions.PreserveWhitespace);
        var body = Body(package);
        var ole = body.Descendants(O + "OLEObject").Single();
        if (ole.Ancestors().Any(e => e.Name == W + "del" || e.Name == W + "moveFrom"))
            throw new InvalidOperationException("A deleted or moved-from MathType revision cannot be edited as the current equation.");
        var paragraph = new XElement(ole.Ancestors(W + "p").First());
        var ownerBody = Body(XDocument.Parse(ownerXml, LoadOptions.PreserveWhitespace));
        // A live Word paragraph can contain several serialized w:p nodes when
        // paragraph marks are tracked deletions. Object/Shape IDs are allocated
        // afresh for each export and cannot correlate the two packages. Use the
        // exact source's ordinal in the already-captured LOCAL physical OLE set.
        var ownerObjects = InlineObjects(ownerBody);
        if (expectedOwnerOleCount >= 0 && ownerObjects.Length != expectedOwnerOleCount)
            throw new InvalidDataException("The local owner export disagrees with its captured physical OLE set.");
        if (sourceOleOrdinal == 0 && ownerObjects.Length == 1) sourceOleOrdinal = 1;
        if (sourceOleOrdinal < 1 || sourceOleOrdinal > ownerObjects.Length)
            throw new InvalidDataException("The source OLE has no unambiguous ordinal in its local owner export.");
        var ownerObject = ownerObjects[sourceOleOrdinal - 1];
        if (ownerObject.Ancestors().Any(e => e.Name == W + "del" || e.Name == W + "moveFrom")
            || !string.Equals((string?)ownerObject.Attribute("ProgID"), (string?)ole.Attribute("ProgID"), StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("The selected local OLE owner is deleted or belongs to another object class.");
        var owner = ownerObject.Ancestors(W + "p").First();
        var cell = owner.Ancestors(W + "tc").FirstOrDefault();
        var localParagraphs = ownerBody.Descendants(W + "p")
            .Where(p => !p.Ancestors(W + "p").Any()
                && ReferenceEquals(p.Ancestors(W + "tc").FirstOrDefault(), cell)).ToArray();
        var ownerIndex = Array.IndexOf(localParagraphs, owner);
        // The surviving paragraph mark owns the formatting of a tracked join.
        // Never import a deleted paragraph mark into a new one-object fragment.
        while (owner.Element(W + "pPr")?.Element(W + "rPr")?.Elements()
                   .Any(e => e.Name == W + "del" || e.Name == W + "moveFrom") == true)
        {
            if (++ownerIndex >= localParagraphs.Length)
                throw new InvalidDataException("A tracked paragraph join has no surviving owner mark in its local export.");
            owner = localParagraphs[ownerIndex];
        }
        paragraph.Elements(W + "pPr").Remove();
        if (owner.Element(W + "pPr") is XElement properties)
            paragraph.AddFirst(new XElement(properties));

        var ids = new HashSet<string>(paragraph.Descendants(W + "bookmarkStart")
            .Select(e => (string?)e.Attribute(W + "id") ?? string.Empty), StringComparer.Ordinal);
        var id = Enumerable.Range(0, ids.Count + 1)
            .Select(i => i.ToString(System.Globalization.CultureInfo.InvariantCulture))
            .First(i => !ids.Contains(i));
        var start = new XElement(W + "bookmarkStart",
            new XAttribute(W + "id", id), new XAttribute(W + "name", locatorName));
        if (paragraph.Element(W + "pPr") is XElement pPr) pPr.AddAfterSelf(start);
        else paragraph.AddFirst(start);
        paragraph.Add(new XElement(W + "bookmarkEnd", new XAttribute(W + "id", id)));
        // No table wrapper or synthetic section properties belong to an OLE edit.
        body.ReplaceNodes(paragraph);
        return package.ToString(SaveOptions.DisableFormatting);
    }

    // A cross-family conversion uses the same exact-object envelope as editing,
    // not an export of the containing live paragraph. That paragraph can include
    // deleted predecessor objects which must never be re-imported as new content.
    // Replace the one object atom while retaining its Word-owned revision wrapper
    // and the surviving paragraph mark's properties.
    internal static string PrepareOmmlReplacement(string objectXml, string ownerXml,
        string locatorName, string omml, int sourceOleOrdinal, int expectedOwnerOleCount,
        string? retiredIdentityBookmark = null)
    {
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var equation = XElement.Parse(omml, LoadOptions.PreserveWhitespace);
        if (equation.Name != m + "oMath")
            throw new InvalidDataException("An OLE replacement requires exactly one prepared native OMath.");
        var package = XDocument.Parse(Prepare(objectXml, ownerXml, locatorName,
            sourceOleOrdinal, expectedOwnerOleCount), LoadOptions.PreserveWhitespace);
        var body = Body(package);
        var atom = body.Descendants(W + "object").Single();
        var run = atom.Parent;
        if (run is null || run.Name != W + "r"
            || run.Elements().Any(e => e.Name != W + "rPr" && !ReferenceEquals(e, atom)))
            throw new InvalidDataException("The selected OLE does not occupy one exact Word object run.");
        // OLE preview baseline/size belongs to the removed object, not to native
        // mathematics. The prepared OMML carries its own glyph formatting.
        var inserted = new XElement(equation);
        run.ReplaceWith(inserted);
        // WPS exports an OLE object as the result of its EMBED field. That field
        // belongs to the removed object: left around the equation, WPS imports
        // a display equation after a manual line break and the written equation
        // no longer matches its own range.
        RemoveEnclosingEmbedField(inserted);
        if (!string.IsNullOrEmpty(retiredIdentityBookmark))
        {
            var identities = body.Descendants(W + "bookmarkStart").Where(e =>
                string.Equals((string?)e.Attribute(W + "name"), retiredIdentityBookmark, StringComparison.OrdinalIgnoreCase)).ToArray();
            var ids = new HashSet<string>(identities.Select(e => (string?)e.Attribute(W + "id") ?? string.Empty), StringComparer.Ordinal);
            body.Descendants(W + "bookmarkEnd").Where(e => ids.Contains((string?)e.Attribute(W + "id") ?? string.Empty)).Remove();
            identities.Remove();
        }
        return package.ToString(SaveOptions.DisableFormatting);
    }

    // A Word paragraph export can carry stale VML preview extents for ANY of its
    // inline OLE hosts, including an untouched/deleted legacy object beside the
    // edited MathType equation. Undo may refresh those caches for the whole local
    // scope. Word owns the placed geometry independently of the OLE server class.
    // Compare the exact live ledger (spans, types, ProgIDs, dimensions) separately;
    // retain every payload, intrinsic extent, run/pPr and revision in the XML.
    // Normalize only VML width/height, not the object or its history. These bytes
    // are recovery evidence only and must NEVER be imported into Word.
    internal static string RecoveryEvidence(string xml, string exactLiveGeometry, int liveOleCount)
    {
        if (string.IsNullOrEmpty(exactLiveGeometry))
            throw new ArgumentException("Exact local Word geometry is required.", nameof(exactLiveGeometry));
        var package = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        var body = Body(package);
        var objects = NormalizeRecoveryPreviewGeometry(body);
        if (objects != liveOleCount)
            throw new InvalidDataException("Local live and serialized OLE recovery ownership disagree.");
        body.SetAttributeValue("mathTypeRecoveryLiveGeometry", exactLiveGeometry);
        return package.ToString(SaveOptions.DisableFormatting);
    }

    // A NEW insertion owns only its insertion boundary, not every pre-existing
    // OLE elsewhere in a potentially huge paragraph. Preserve the complete
    // serialized paragraph/revision/bookmark/payload evidence while normalizing
    // only Word-owned VML preview width/height. Exact live geometry is verified
    // separately for a bounded range around the insertion point.
    internal static string InsertionRecoveryEvidence(string xml)
    {
        var package = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        var body = Body(package);
        NormalizeRecoveryPreviewGeometry(body);
        body.SetAttributeValue("mathTypeInsertionRecovery", "serialized");
        return package.ToString(SaveOptions.DisableFormatting);
    }

    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";

    // The OLE objects Word lists in Range.InlineShapes: floating objects
    // (position:absolute VML) and text-box content are Shapes, not inline, so
    // counting them would misalign the live ordinal with the serialized one.
    internal static XElement[] InlineObjects(XElement body) => body.Descendants(O + "OLEObject")
        .Where(o => o.Parent?.Name == W + "object" && !o.Ancestors(W + "txbxContent").Any()
            && !IsFloating(o))
        .ToArray();

    private static bool IsFloating(XElement ole)
    {
        var id = (string?)ole.Attribute("ShapeID");
        var shape = ole.Parent!.Elements(V + "shape").FirstOrDefault(e => (string?)e.Attribute("id") == id);
        var style = (string?)shape?.Attribute("style") ?? string.Empty;
        return style.Split(';').Any(d => d.Split(':')[0].Trim() == "position"
            && d.Substring(d.IndexOf(':') + 1).Trim() == "absolute");
    }

    private static int NormalizeRecoveryPreviewGeometry(XElement body)
    {
        var objects = InlineObjects(body);
        foreach (var ole in objects)
        {
            // Normalization is evidence hygiene only. A preview shape Word
            // wrote without a unique id or size stays as serialized; both
            // sides of the comparison go through the same code.
            var id = (string?)ole.Attribute("ShapeID");
            var shapes = ole.Parent!.Elements(V + "shape").Where(e => (string?)e.Attribute("id") == id).ToArray();
            if (shapes.Length != 1 || shapes[0].Attribute("style") is not XAttribute style) continue;
            var declarations = style.Value.Split(';');
            foreach (var dimension in new[] { "width", "height" })
            {
                var matches = Enumerable.Range(0, declarations.Length)
                    .Where(i => declarations[i].Split(':')[0].Trim() == dimension).ToArray();
                if (matches.Length == 1) declarations[matches[0]] = dimension + ":word-host";
            }
            style.Value = string.Join(";", declarations);
        }
        return objects.Length;
    }

    // Removes the field runs around `content` when they form exactly one EMBED
    // field whose result is `content`: begin, instruction run(s), separate before
    // it and end after it (bookmarks may sit between). Anything else is left.
    internal static bool RemoveEnclosingEmbedField(XElement content)
    {
        static bool IsMarker(XElement element) =>
            element.Name == W + "bookmarkStart" || element.Name == W + "bookmarkEnd" || element.Name == W + "proofErr";
        static XElement? Single(XElement run) =>
            run.Name != W + "r" ? null
                : run.Elements().Where(e => e.Name != W + "rPr").ToArray() is { Length: 1 } items ? items[0] : null;
        static bool IsFieldChar(XElement run, string kind) =>
            Single(run) is { } item && item.Name == W + "fldChar"
            && (string?)item.Attribute(W + "fldCharType") == kind;

        var before = content.ElementsBeforeSelf().Reverse().Where(e => !IsMarker(e)).ToList();
        var after = content.ElementsAfterSelf().Where(e => !IsMarker(e)).ToList();
        if (before.Count < 3 || after.Count < 1 || !IsFieldChar(before[0], "separate") || !IsFieldChar(after[0], "end"))
            return false;
        var instruction = new List<XElement>();
        var index = 1;
        while (index < before.Count && Single(before[index]) is { } item && item.Name == W + "instrText")
            instruction.Add(before[index++]);
        if (instruction.Count == 0 || index >= before.Count || !IsFieldChar(before[index], "begin"))
            return false;
        var code = string.Concat(instruction.AsEnumerable().Reverse().Select(run => Single(run)!.Value)).Trim();
        if (!code.StartsWith("EMBED ", StringComparison.OrdinalIgnoreCase))
            return false;
        before[0].Remove();
        foreach (var run in instruction) run.Remove();
        before[index].Remove();
        after[0].Remove();
        return true;
    }

    internal static string[] BookmarkNames(string ownerXml) =>
        Body(XDocument.Parse(ownerXml)).Descendants(W + "bookmarkStart")
            .Select(e => (string?)e.Attribute(W + "name"))
            .Where(name => !string.IsNullOrEmpty(name)).Select(name => name!)
            .Distinct(StringComparer.Ordinal).ToArray();
}
