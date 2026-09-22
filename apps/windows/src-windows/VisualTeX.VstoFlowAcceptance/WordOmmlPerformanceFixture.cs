using System.IO.Compression;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // TEST FIXTURES ONLY. Work on a fresh package copy before Word opens it.
    // Preserve the first N independent generated 1x3 OMML hosts and their prose,
    // styles, relationships and exact managed metadata. Never rewrite a source.
    private static void CreateEmptyWordPerformanceSeed(string output)
    {
        if (File.Exists(output)) throw new InvalidDataException("Do not overwrite a Word startup seed.");
        using var package = ZipFile.Open(output, ZipArchiveMode.Create);
        void Part(string name, string xml)
        {
            using var writer = new StreamWriter(package.CreateEntry(name).Open(), new System.Text.UTF8Encoding(false));
            writer.Write(xml);
        }
        Part("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
        Part("_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
        Part("word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p/><w:sectPr/></w:body></w:document>");
    }

    private static void PromoteOpenDocumentToManagedUnnumberedOmml(Word.Document document)
    {
        Word.Bookmarks? bookmarks = null;
        Word.OMaths? maths = null;
        var ranges = new List<Word.Range>();
        var createdIds = new List<string>();
        try
        {
            bookmarks = document.Bookmarks;
            var staleNames = new List<string>();
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Word.Bookmark? bookmark = null;
                try
                {
                    bookmark = bookmarks[index];
                    var name = bookmark.Name ?? string.Empty;
                    if (name.StartsWith(
                            WordOmmlFormulaStore.BookmarkPrefix,
                            StringComparison.OrdinalIgnoreCase))
                        staleNames.Add(name);
                }
                finally { Release(bookmark); }
            }
            foreach (var name in staleNames)
            {
                Word.Bookmark? bookmark = null;
                try
                {
                    if (!bookmarks.Exists(name)) continue;
                    bookmark = bookmarks[name];
                    bookmark.Delete();
                }
                finally { Release(bookmark); }
            }

            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                Word.OMath? math = null;
                Word.Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range.Duplicate;
                    ranges.Add(range);
                    range = null;
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
        }
        finally
        {
            Release(maths);
            Release(bookmarks);
        }

        foreach (var range in ranges)
        {
            Word.Bookmark? identity = null;
            try
            {
                var metadata = WordOmmlNativeSource.CreateForNative(
                    document,
                    range);
                if (metadata.Numbered)
                    throw new InvalidDataException(
                        "Performance fixture promotion requires unnumbered OMML.");
                identity = WordOmmlFormulaStore.Wrap(
                    document,
                    range,
                    metadata,
                    replaceExisting: true);
                WordOmmlFormulaStore.Save(document, metadata);
                createdIds.Add(metadata.FormulaId);
            }
            finally
            {
                Release(identity);
                Release(range);
            }
        }

        foreach (var formulaId in createdIds)
        {
            if (WordOmmlFormulaStore.TryRead(document, formulaId) is null)
                throw new InvalidDataException(
                    "Managed OMML performance fixture lost stored metadata.");
            Word.Bookmark? identity = null;
            try
            {
                identity = WordOmmlFormulaStore.FindByFormulaId(
                    document,
                    formulaId)
                    ?? throw new InvalidDataException(
                        "Managed OMML performance fixture lost its identity bookmark.");
            }
            finally { Release(identity); }
        }
        Console.WriteLine(
            $"[PERF FIXTURE] managedUnnumberedOmml={createdIds.Count}; metadata={createdIds.Count}");
    }

    private static void RunPrepareManagedUnnumberedOmmlPerformanceFixtureAcceptance(
        string artifactRoot)
    {
        if (AttachActiveWord)
            throw new InvalidOperationException(
                "Managed OMML performance fixture preparation must never attach active user Word.");
        Directory.CreateDirectory(artifactRoot);
        var source = Path.GetFullPath(Path.Combine(
            "docs",
            "remediation-3d207d7",
            "evidence",
            "baseline-body100-s0-omml-audit.xml"));
        if (!File.Exists(source))
            throw new FileNotFoundException(
                "The 100-formula Flat OPC performance source is missing.",
                source);
        var output = Path.Combine(
            artifactRoot,
            "managed-unnumbered-100.docx");
        MaterializeFlatOpcOmmlPerformanceFixture(
            source,
            output,
            expectedCount: 100);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Bookmarks? bookmarks = null;
        try
        {
            application = CreateFreshAcceptanceAutomationWord(artifactRoot);
            document = application.Documents.Open(
                output,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            PromoteOpenDocumentToManagedUnnumberedOmml(document);
            document.Save();

            AssertEqual(100, document.OMaths.Count,
                "Managed unnumbered performance fixture changed the OMath count.");
            AssertEqual(0, document.Tables.Count,
                "Managed unnumbered performance fixture unexpectedly contains tables.");
            AssertEqual(0, document.Fields.Count,
                "Managed unnumbered performance fixture unexpectedly contains fields.");

            bookmarks = document.Bookmarks;
            var managedIds = new List<string>();
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Word.Bookmark? bookmark = null;
                try
                {
                    bookmark = bookmarks[index];
                    var name = bookmark.Name ?? string.Empty;
                    if (!name.StartsWith(
                            WordOmmlFormulaStore.BookmarkPrefix,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    var suffix = name.Substring(
                        WordOmmlFormulaStore.BookmarkPrefix.Length);
                    if (Guid.TryParseExact(
                            suffix,
                            "N",
                            out var parsed))
                        managedIds.Add(parsed.ToString("D"));
                }
                finally { Release(bookmark); }
            }
            AssertEqual(100, managedIds.Count,
                "Managed unnumbered performance fixture did not retain 100 VTOMML identities.");
            foreach (var formulaId in managedIds)
            {
                var metadata = WordOmmlFormulaStore.TryRead(
                    document,
                    formulaId)
                    ?? throw new InvalidDataException(
                        "Managed unnumbered performance fixture lost metadata for "
                        + formulaId);
                AssertTrue(!metadata.Numbered,
                    "Managed unnumbered performance fixture contains a numbered formula.");
            }
            Console.WriteLine(
                $"[MANAGED UNNUMBERED PERF FIXTURE] omml={document.OMaths.Count}; "
                + $"identities={managedIds.Count}; tables={document.Tables.Count}; "
                + $"fields={document.Fields.Count}; output={output}");
        }
        finally
        {
            Release(bookmarks);
            if (document is not null)
            {
                try
                {
                    document.Close(
                        Word.WdSaveOptions.wdSaveChanges);
                }
                catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void MaterializeFlatOpcOmmlPerformanceFixture(string source, string output, int expectedCount)
    {
        if (File.Exists(output)) throw new InvalidDataException("Do not overwrite a Flat OPC performance fixture.");
        var sourceXml = XDocument.Load(source, LoadOptions.PreserveWhitespace);
        XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        XNamespace contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        var packageRoot = sourceXml.Root;
        if (packageRoot?.Name != pkg + "package")
            throw new InvalidDataException("Performance snapshot is not a Flat OPC package.");
        var parts = packageRoot.Elements(pkg + "part").Select(part => new
        {
            Element = part,
            Name = (string?)part.Attribute(pkg + "name") ?? "",
            ContentType = (string?)part.Attribute(pkg + "contentType") ?? "",
            Compression = (string?)part.Attribute(pkg + "compression") ?? "",
        }).ToArray();
        if (parts.Length == 0 || parts.Any(part => string.IsNullOrWhiteSpace(part.Name)
                || !part.Name.StartsWith("/", StringComparison.Ordinal)
                || string.IsNullOrWhiteSpace(part.ContentType)))
            throw new InvalidDataException("Flat OPC snapshot contains an invalid package part.");

        using (var archive = ZipFile.Open(output, ZipArchiveMode.Create))
        {
            var types = new XDocument(
                new XElement(contentTypes + "Types",
                    new XElement(contentTypes + "Default",
                        new XAttribute("Extension", "rels"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                    new XElement(contentTypes + "Default",
                        new XAttribute("Extension", "xml"),
                        new XAttribute("ContentType", "application/xml"))));
            foreach (var part in parts)
            {
                if (string.Equals(part.Name, "/[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
                    continue;
                var extension = Path.GetExtension(part.Name).TrimStart('.');
                var defaultType = string.Equals(extension, "rels", StringComparison.OrdinalIgnoreCase)
                    ? "application/vnd.openxmlformats-package.relationships+xml"
                    : string.Equals(extension, "xml", StringComparison.OrdinalIgnoreCase)
                        ? "application/xml"
                        : null;
                if (!string.Equals(defaultType, part.ContentType, StringComparison.OrdinalIgnoreCase))
                {
                    types.Root!.Add(new XElement(contentTypes + "Override",
                        new XAttribute("PartName", part.Name),
                        new XAttribute("ContentType", part.ContentType)));
                }

                var entry = archive.CreateEntry(
                    part.Name.TrimStart('/'),
                    string.Equals(part.Compression, "store", StringComparison.OrdinalIgnoreCase)
                        ? CompressionLevel.NoCompression
                        : CompressionLevel.Optimal);
                using var stream = entry.Open();
                var xmlData = part.Element.Element(pkg + "xmlData");
                var binaryData = part.Element.Element(pkg + "binaryData");
                if (xmlData is not null)
                {
                    var root = xmlData.Elements().SingleOrDefault()
                        ?? throw new InvalidDataException("Flat OPC XML part has no root element: " + part.Name);
                    using var writer = System.Xml.XmlWriter.Create(stream, new System.Xml.XmlWriterSettings
                    {
                        Encoding = new System.Text.UTF8Encoding(false),
                        OmitXmlDeclaration = false,
                        Indent = false,
                    });
                    new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(root))
                        .Save(writer);
                }
                else if (binaryData is not null)
                {
                    var bytes = Convert.FromBase64String(binaryData.Value);
                    stream.Write(bytes, 0, bytes.Length);
                }
                else
                {
                    throw new InvalidDataException("Flat OPC part has neither XML nor binary content: " + part.Name);
                }
            }
            using var typesWriter = new StreamWriter(
                archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal).Open(),
                new System.Text.UTF8Encoding(false));
            types.Save(typesWriter, SaveOptions.DisableFormatting);
        }

        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        using (var update = ZipFile.Open(output, ZipArchiveMode.Update))
        {
            var entry = update.GetEntry("word/document.xml")
                ?? throw new InvalidDataException("Materialized Flat OPC fixture has no Word document part.");
            XDocument doc;
            using (var stream = entry.Open()) doc = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            var bodyNode = doc.Root?.Element(w + "body")
                ?? throw new InvalidDataException("Materialized Flat OPC fixture has no document body.");
            var formulaOwners = bodyNode.Elements()
                .Where(element => element.Name != w + "sectPr" && element.Descendants(m + "oMath").Any())
                .ToArray();
            if (expectedCount <= 0 || expectedCount > formulaOwners.Length)
                throw new InvalidDataException($"Requested Flat OPC formula prefix is invalid: requested={expectedCount}; available={formulaOwners.Length}.");
            if (expectedCount < formulaOwners.Length)
            {
                var last = formulaOwners[expectedCount - 1];
                foreach (var element in last.ElementsAfterSelf().Where(element => element.Name != w + "sectPr").ToArray())
                    element.Remove();
                var sectPr = bodyNode.Element(w + "sectPr");
                if (sectPr is null)
                    bodyNode.Add(new XElement(w + "sectPr"));
                entry.Delete();
                using var replacement = update.CreateEntry("word/document.xml", CompressionLevel.Optimal).Open();
                doc.Save(replacement, SaveOptions.DisableFormatting);
            }
        }

        using var verification = ZipFile.OpenRead(output);
        var documentEntry = verification.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("Materialized Flat OPC fixture has no Word document part.");
        XDocument document;
        using (var stream = documentEntry.Open()) document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        var body = document.Root?.Element(w + "body")
            ?? throw new InvalidDataException("Materialized Flat OPC fixture has no document body.");
        var maths = body.Descendants(m + "oMath").Count();
        var identityBookmarks = body.Descendants(w + "bookmarkStart")
            .Count(bookmark => ((string?)bookmark.Attribute(w + "name") ?? "")
                .StartsWith("VTOMML_", StringComparison.OrdinalIgnoreCase));
        if (maths != expectedCount
            || identityBookmarks != expectedCount
            || body.Descendants(w + "tbl").Any()
            || body.Descendants(w + "fldSimple").Any()
            || body.Descendants(w + "instrText").Any())
            throw new InvalidDataException(
                $"Flat OPC fixture is not the expected unnumbered OMML corpus: maths={maths}; identities={identityBookmarks}; tables={body.Descendants(w + "tbl").Count()}; fields={body.Descendants(w + "fldSimple").Count() + body.Descendants(w + "instrText").Count()}.");
        Console.WriteLine($"[PERF FIXTURE] flatOpcUnnumberedOmml={maths}; identities={identityBookmarks}; sourceUnchanged={source}");
    }

    private static void CopyNumberedOmmlPerformanceFixture(string source, string output, int count)
    {
        if (string.Equals(Path.GetFullPath(source), Path.GetFullPath(output), StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("A performance fixture must be a new copy.");
        File.Copy(source, output, overwrite: false);
        if (count <= 0) return;
        using var package = ZipFile.Open(output, ZipArchiveMode.Update);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        XNamespace types = "http://schemas.openxmlformats.org/package/2006/content-types";
        XNamespace metadataNs = "urn:visualtex:word-omml:1";
        XDocument ReadXml(string path)
        {
            var entry = package.GetEntry(path) ?? throw new InvalidDataException("Missing fixture part: " + path);
            using var stream = entry.Open();
            return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        }
        void ReplaceXml(string path, XDocument xml)
        {
            package.GetEntry(path)?.Delete();
            using var stream = package.CreateEntry(path, CompressionLevel.Optimal).Open();
            xml.Save(stream, SaveOptions.DisableFormatting);
        }
        var document = ReadXml("word/document.xml");
        var body = document.Root?.Element(w + "body")
            ?? throw new InvalidDataException("Fixture has no document body.");
        var tables = body.Elements(w + "tbl").ToArray();
        if (count > tables.Length || tables.Any(table =>
                table.Elements(w + "tr").Count() != 1
                || table.Element(w + "tr")!.Elements(w + "tc").Count() != 3
                || table.Descendants(m + "oMath").Count() != 1
                || table.Descendants(w + "tbl").Any()))
            throw new InvalidDataException("Prefix fixtures require independent 1x3 single-OMath body tables.");
        // The full-size benchmark must be the exact historical package copy,
        // not a regenerated equivalent fixture.
        if (count == tables.Length) return;
        var last = tables[count - 1];
        var tail = last.ElementsAfterSelf().ToArray();
        foreach (var element in tail.Where(element => element.Name != w + "sectPr")) element.Remove();
        // Word requires a body paragraph after a terminal table. This test-only
        // paragraph contains no bookmark or field copied from a removed formula.
        last.AddAfterSelf(new XElement(w + "p"));
        var retainedNames = new HashSet<string>(body.Descendants(w + "bookmarkStart")
            .Select(bookmark => (string?)bookmark.Attribute(w + "name") ?? ""), StringComparer.OrdinalIgnoreCase);
        var deletedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var retainedMetadata = 0;
        foreach (var entry in package.Entries.Where(entry => entry.FullName.StartsWith("customXml/item", StringComparison.Ordinal)
                     && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)
                     && !entry.FullName.StartsWith("customXml/itemProps", StringComparison.Ordinal)).ToArray())
        {
            XDocument xml;
            using (var stream = entry.Open()) xml = XDocument.Load(stream);
            if (xml.Root?.Name != metadataNs + "formula") continue;
            var id = (string?)xml.Root.Attribute("formulaId")
                ?? throw new InvalidDataException("Formula metadata has no identity.");
            if (retainedNames.Contains(WordOmmlFormulaStore.BookmarkName(id)))
            {
                retainedMetadata++;
                continue;
            }
            var relPath = "customXml/_rels/" + Path.GetFileName(entry.FullName) + ".rels";
            if (package.GetEntry(relPath) is not null)
            {
                var rels = ReadXml(relPath);
                foreach (var rel in rels.Descendants(relationships + "Relationship"))
                {
                    var target = (string?)rel.Attribute("Target") ?? "";
                    if (target.StartsWith("itemProps", StringComparison.Ordinal) && !target.Contains("/"))
                        deletedParts.Add("customXml/" + target);
                    else throw new InvalidDataException("Unexpected custom XML relationship in prefix fixture.");
                }
                deletedParts.Add(relPath);
            }
            deletedParts.Add(entry.FullName);
        }
        var documentRels = ReadXml("word/_rels/document.xml.rels");
        foreach (var rel in documentRels.Descendants(relationships + "Relationship").ToArray())
        {
            var target = (string?)rel.Attribute("Target") ?? "";
            if (target.StartsWith("../", StringComparison.Ordinal) && deletedParts.Contains(target.Substring(3))) rel.Remove();
        }
        var contentTypes = ReadXml("[Content_Types].xml");
        foreach (var type in contentTypes.Descendants(types + "Override").ToArray())
        {
            var path = ((string?)type.Attribute("PartName") ?? "").TrimStart('/');
            if (deletedParts.Contains(path)) type.Remove();
        }
        if (body.Descendants(m + "oMath").Count() != count || retainedMetadata != count)
            throw new InvalidDataException("Prefix fixture did not preserve exactly one managed metadata part per equation.");
        foreach (var part in deletedParts) package.GetEntry(part)?.Delete();
        ReplaceXml("word/document.xml", document);
        ReplaceXml("word/_rels/document.xml.rels", documentRels);
        ReplaceXml("[Content_Types].xml", contentTypes);
        Console.WriteLine($"[PERF FIXTURE] independentNumberedHosts={count}; metadata={retainedMetadata}; sourceUnchanged={source}");
    }
}
