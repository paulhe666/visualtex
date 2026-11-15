using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordOmmlConverter;
using static VisualTeX.WordVsto.WordOmmlBatchSources;

namespace VisualTeX.WordVsto;

/// <summary>
/// The documents that carry OMML into Word: the document XML, its math font, temporary docx and Flat OPC packages,
/// and the WPS inline sentinel.
/// </summary>
internal static class WordOmmlDocuments
{
    // WPS turns an equation that is the only content of an imported paragraph into a
    // display oMathPara and keeps it display after merging it into a text paragraph.
    // An inline import there carries this trailing word joiner, removed right after.
    private const string WpsInlineImportSentinel = "\u2060";

    internal static string BuildDocumentXml(
        string omml,
        bool includeLeadingTab = false,
        bool forceInline = false)
    {
        var equation = ExtractSingleOMath(omml);
        var contextPrefix = forceInline ? "<w:r><w:t>L</w:t></w:r>" : string.Empty;
        var contextSuffix = forceInline ? "<w:r><w:t>R</w:t></w:r>" : string.Empty;
        var insertPrefix = includeLeadingTab ? "<w:r><w:tab/></w:r>" : string.Empty;
        var bookmarkedFormula =
            $"<w:bookmarkStart w:id=\"0\" w:name=\"{FormulaBookmarkName}\"/>"
            + insertPrefix
            + equation
            + "<w:bookmarkEnd w:id=\"0\"/>";
        return $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{WordNamespace}\" xmlns:m=\"{MathNamespace}\">"
            + $"<w:body><w:p>{contextPrefix}{bookmarkedFormula}{contextSuffix}</w:p><w:sectPr/></w:body></w:document>";
    }

    internal static void PrewarmOmmlReadPipeline()
    {
        _ = OmmlMathMlProjection.Project(
            "<m:oMath xmlns:m=\"" + MathNamespace + "\"><m:r><m:t>x</m:t></m:r></m:oMath>");
        _ = MathMlToLatexConverter.Convert(
            @"<math xmlns=""http://www.w3.org/1998/Math/MathML"" display=""inline""><mi>x</mi></math>");
    }

    // WPS reports the math font of its documents as "default" and writes that into runs
    // (Word reads it back from WPS-saved files): a placeholder, not a font. Written as a run
    // font it makes WPS draw upright letters (\ker, \mathrm{d}) as unrelated symbols.
    internal static string NormalizeMathFontName(string? mathFontName) =>
        string.IsNullOrWhiteSpace(mathFontName)
        || string.Equals(mathFontName!.Trim(), "default", StringComparison.OrdinalIgnoreCase)
            ? "Cambria Math"
            : mathFontName.Trim();

    internal static void ApplyDocumentMathFont(
        Document document,
        string mathFontName)
    {
        var normalized = NormalizeMathFontName(mathFontName);
        string current;
        try { current = document.OMathFontName ?? string.Empty; }
        catch (COMException error)
        {
            throw new InvalidOperationException(
                "Word could not read the temporary document's Office Math font.",
                error);
        }
        if (string.Equals(current, normalized, StringComparison.OrdinalIgnoreCase))
            return;
        try { document.OMathFontName = normalized; }
        catch (COMException error)
        {
            throw new InvalidOperationException(
                $"Word could not apply '{normalized}' to the temporary OMML document.",
                error);
        }
        var applied = document.OMathFontName ?? string.Empty;
        if (!string.Equals(applied, normalized, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                $"Word rejected '{normalized}' as the temporary OMML document's math font.");
    }

    private static void WriteMinimalDocxScaffold(
        ZipArchive archive,
        string mathFontName)
    {
        var normalized = NormalizeMathFontName(mathFontName);
        WriteEntry(
            archive,
            "[Content_Types].xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
            + "<Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>"
            + "</Types>");
        WriteEntry(
            archive,
            "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
            + "</Relationships>");
        WriteEntry(
            archive,
            "word/_rels/document.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rIdSettings\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/>"
            + "</Relationships>");

        XNamespace word = WordNamespace;
        XNamespace math = MathNamespace;
        var settings = new XElement(
            word + "settings",
            new XAttribute(XNamespace.Xmlns + "w", WordNamespace),
            new XAttribute(XNamespace.Xmlns + "m", MathNamespace),
            new XElement(
                math + "mathPr",
                new XElement(
                    math + "mathFont",
                    new XAttribute(math + "val", normalized))));
        WriteEntry(
            archive,
            "word/settings.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + settings.ToString(SaveOptions.DisableFormatting));
    }

    internal static string ApplyExplicitTransferMathFont(
        string omml,
        string mathFontName)
    {
        if (string.IsNullOrWhiteSpace(omml)) return omml;
        var normalizedMathFontName = NormalizeMathFontName(mathFontName);
        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;

        static void SetRunFonts(
            XElement wordRunProperties,
            XNamespace wordNamespace,
            string fontName)
        {
            var fonts = wordRunProperties.Element(wordNamespace + "rFonts");
            if (fonts is null)
            {
                fonts = new XElement(wordNamespace + "rFonts");
                wordRunProperties.AddFirst(fonts);
            }
            fonts.SetAttributeValue(wordNamespace + "ascii", fontName);
            fonts.SetAttributeValue(wordNamespace + "hAnsi", fontName);
        }

        foreach (var run in document.Descendants(math + "r"))
        {
            // m:nor is deliberate ordinary text inside math (for example \mathrm).
            // Keep its body/text font untouched; only native mathematical glyph
            // runs need the document's selected OpenType MATH font forced during
            // cross-document FormattedText transfer.
            var mathRunProperties = run.Element(math + "rPr");
            if (mathRunProperties?.Element(math + "nor") is not null)
                continue;
            var wordRunProperties = run.Element(word + "rPr");
            if (wordRunProperties is null)
            {
                wordRunProperties = new XElement(word + "rPr");
                if (mathRunProperties is not null)
                    mathRunProperties.AddAfterSelf(wordRunProperties);
                else
                    run.AddFirst(wordRunProperties);
            }
            SetRunFonts(wordRunProperties, word, normalizedMathFontName);
        }

        foreach (var controlProperties in document.Descendants(math + "ctrlPr"))
        {
            var wordRunProperties = controlProperties.Element(word + "rPr");
            if (wordRunProperties is null)
            {
                wordRunProperties = new XElement(word + "rPr");
                controlProperties.Add(wordRunProperties);
            }
            SetRunFonts(wordRunProperties, word, normalizedMathFontName);
        }

        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    internal static string CreateTemporaryDocumentDocx(
        string documentXml,
        string? mathFontName = null)
    {
        var path = Path.Combine(
            Path.GetTempPath(),
            $"visualtex-omml-document-{Guid.NewGuid():N}.docx");
        using var stream = new FileStream(
            path,
            FileMode.CreateNew,
            FileAccess.ReadWrite,
            FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: false);
        WriteMinimalDocxScaffold(
            archive,
            NormalizeMathFontName(mathFontName));
        WriteEntry(archive, "word/document.xml", documentXml);
        return path;
    }

    internal static string CreateTemporaryDisplayGroupDocx(
        IReadOnlyList<BatchEntry> entries,
        string mathFontName)
    {
        var body = new StringBuilder();
        for (var index = 0; index < entries.Count; index++)
        {
            body.Append("<w:p>")
                .Append(ApplyExplicitTransferMathFont(
                    ExtractSingleOMath(entries[index].Omml),
                    mathFontName))
                .Append("</w:p>");
        }
        // The opened source document supplies one terminal ordinary paragraph.
        // ReplaceDisplayParagraphGroup excludes only that final paragraph mark and
        // copies the complete preceding multi-paragraph topology.
        body.Append("<w:p/>");
        var documentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{WordNamespace}\" xmlns:m=\"{MathNamespace}\">"
            + $"<w:body>{body}<w:sectPr/></w:body></w:document>";
        return CreateTemporaryDocumentDocx(documentXml, mathFontName);
    }

    internal static string CreateTemporaryAdjacentInlineGroupDocx(
        IReadOnlyList<BatchEntry> entries,
        string mathFontName)
    {
        var body = new StringBuilder("<w:p>");
        for (var index = 0; index < entries.Count; index++)
        {
            var entry = entries[index];
            body.Append(ApplyExplicitTransferMathFont(
                ExtractSingleOMath(entry.Omml),
                mathFontName));
            if (index + 1 < entries.Count)
            {
                // Keep an explicit ordinary Word run between sibling OMath nodes.
                // It is hidden and contains a zero-width non-joiner, so it has no
                // visible layout footprint while still being structurally outside
                // both math zones in the source OpenXML.
                body.Append("<w:r><w:rPr><w:vanish/></w:rPr><w:t>‌</w:t></w:r>");
            }
        }
        body.Append("</w:p>");
        var documentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{WordNamespace}\" xmlns:m=\"{MathNamespace}\">"
            + $"<w:body>{body}<w:sectPr/></w:body></w:document>";
        return CreateTemporaryDocumentDocx(documentXml, mathFontName);
    }

    // Puts the WPS inline import separator right after the last equation of the main
    // body's last paragraph (at its end when it has none).
    internal static string AppendWpsInlineSentinel(string flatOpc)
    {
        XNamespace word = WordNamespace;
        XNamespace math = MathNamespace;
        var document = XDocument.Parse(flatOpc, LoadOptions.PreserveWhitespace);
        var paragraph = WordDocumentXml.GetRequiredMainDocumentBody(document)
            .Elements(word + "p")
            .LastOrDefault()
            ?? throw new InvalidDataException(
                "The OMML insertion package has no body paragraph.");
        var sentinel = new XElement(
            word + "r",
            new XElement(word + "t", WpsInlineImportSentinel));
        var equation = paragraph
            .Elements()
            .LastOrDefault(element => element.Name == math + "oMath" || element.Name == math + "oMathPara");
        if (equation is null) paragraph.Add(sentinel);
        else equation.AddAfterSelf(sentinel);
        return (document.Declaration?.ToString() ?? string.Empty)
            + document.ToString(SaveOptions.DisableFormatting);
    }

    internal static void RemoveWpsInlineSentinel(OMath math)
    {
        Range? equation = null;
        Range? after = null;
        try
        {
            equation = math.Range;
            after = equation.Duplicate;
            after.SetRange(equation.End, equation.End + 1);
            if (!string.Equals(after.Text, WpsInlineImportSentinel, StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "WPS did not keep the inline OMML import separator after the equation.");
            after.Delete();
        }
        finally
        {
            Release(after);
            Release(equation);
        }
    }

    internal static string CreateSingleEquationFlatOpc(
        BatchEntry entry,
        string mathFontName,
        string? paragraphMarkRunProperties = null,
        bool trailingWpsSentinel = false)
    {
        XNamespace package =
            "http://schemas.microsoft.com/office/2006/xmlPackage";
        XNamespace relationships =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        XNamespace word = WordNamespace;
        XNamespace math = MathNamespace;

        var equation = XElement.Parse(
            ApplyExplicitTransferMathFont(
                ExtractSingleOMath(entry.Omml),
                mathFontName),
            LoadOptions.PreserveWhitespace);
        var document = new XElement(
            word + "document",
            new XAttribute(XNamespace.Xmlns + "w", word),
            new XAttribute(XNamespace.Xmlns + "m", math),
            new XElement(
                word + "body",
                new XElement(
                    word + "p",
                    // Where Word keeps this package's paragraph mark (an import
                    // into an empty paragraph), it keeps the package's paragraph
                    // properties: they are the target paragraph's own (style,
                    // list numbering, indentation, spacing, mark formatting).
                    paragraphMarkRunProperties is null
                        ? null
                        : XElement.Parse(paragraphMarkRunProperties),
                    new XElement(
                        word + "bookmarkStart",
                        new XAttribute(word + "id", entry.BookmarkId),
                        new XAttribute(word + "name", entry.BookmarkName)),
                    equation,
                    new XElement(
                        word + "bookmarkEnd",
                        new XAttribute(word + "id", entry.BookmarkId)),
                    trailingWpsSentinel
                        ? new XElement(
                            word + "r",
                            new XElement(word + "t", WpsInlineImportSentinel))
                        : null),
                new XElement(word + "sectPr")));
        var rootRelationships = new XElement(
            relationships + "Relationships",
            new XElement(
                relationships + "Relationship",
                new XAttribute("Id", "rId1"),
                new XAttribute(
                    "Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
                new XAttribute("Target", "word/document.xml")));
        // Word 2019/2021 rejects a synthetic settings.xml relationship unless
        // the complete companion style/theme/font package is also present.
        // Range.WordOpenXML proves that a document part with an empty relationship
        // part is sufficient for OMath InsertXML, so keep the single-formula
        // transport package deliberately minimal.
        var documentRelationships = new XElement(
            relationships + "Relationships");

        XElement XmlPart(
            string name,
            string contentType,
            XElement content) =>
            new(
                package + "part",
                new XAttribute(package + "name", name),
                new XAttribute(package + "contentType", contentType),
                new XElement(package + "xmlData", content));

        var flatOpc = new XElement(
            package + "package",
            new XAttribute(XNamespace.Xmlns + "pkg", package),
            XmlPart(
                "/_rels/.rels",
                "application/vnd.openxmlformats-package.relationships+xml",
                rootRelationships),
            XmlPart(
                "/word/document.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                document),
            XmlPart(
                "/word/_rels/document.xml.rels",
                "application/vnd.openxmlformats-package.relationships+xml",
                documentRelationships));
        return "<?xml version=\"1.0\" standalone=\"yes\"?>"
            + flatOpc.ToString(SaveOptions.DisableFormatting);
    }

    internal static string CreateTemporaryBatchDocx(
        IReadOnlyList<BatchEntry> entries,
        string mathFontName)
    {
        var path = Path.Combine(
            Path.GetTempPath(),
            $"visualtex-omml-batch-{Guid.NewGuid():N}.docx");
        using var stream = new FileStream(
            path,
            FileMode.CreateNew,
            FileAccess.ReadWrite,
            FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: false);
        WriteMinimalDocxScaffold(archive, mathFontName);

        var body = new StringBuilder();
        foreach (var entry in entries)
        {
            var equation = ApplyExplicitTransferMathFont(
                ExtractSingleOMath(entry.Omml),
                mathFontName);
            body.Append("<w:p><w:r><w:t>L</w:t></w:r>")
                .Append("<w:bookmarkStart w:id=\"")
                .Append(entry.BookmarkId)
                .Append("\" w:name=\"")
                .Append(entry.BookmarkName)
                .Append("\"/>")
                .Append(equation)
                .Append("<w:bookmarkEnd w:id=\"")
                .Append(entry.BookmarkId)
                .Append("\"/>")
                .Append("<w:r><w:t>R</w:t></w:r></w:p>");
        }
        // Keep one explicit pristine Normal paragraph at the end. Besides making
        // the DOCX conventional, this gives MathType→OMML conversion a clean
        // paragraph mark that can replace the source OLE paragraph's hidden live
        // layout state without opening another scratch document.
        body.Append("<w:p/>");
        WriteEntry(
            archive,
            "word/document.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{WordNamespace}\" xmlns:m=\"{MathNamespace}\">"
            + $"<w:body>{body}<w:sectPr/></w:body></w:document>");
        return path;
    }

    internal static string CreateTemporaryDocx(
        string omml,
        bool includeLeadingTab,
        bool forceInline,
        string mathFontName)
    {
        var path = Path.Combine(
            Path.GetTempPath(),
            $"visualtex-omml-{Guid.NewGuid():N}.docx");
        using var stream = new FileStream(
            path,
            FileMode.CreateNew,
            FileAccess.ReadWrite,
            FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: false);
        WriteMinimalDocxScaffold(archive, mathFontName);
        WriteEntry(
            archive,
            "word/document.xml",
            BuildDocumentXml(
                ApplyExplicitTransferMathFont(omml, mathFontName),
                includeLeadingTab,
                forceInline));
        return path;
    }

    internal static void WriteEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(content);
    }
}
