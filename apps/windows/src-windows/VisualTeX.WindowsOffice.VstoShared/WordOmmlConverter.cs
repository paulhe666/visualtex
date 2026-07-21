using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static class WordOmmlConverter
{
    private const string MathNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string WordNamespace =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly object TransformLock = new();
    private static XslCompiledTransform? _mathMlToOmml;
    private static XslCompiledTransform? _ommlToMathMl;

    internal static Range Insert(
        Application application,
        Document targetDocument,
        Range insertionRange,
        string mathMl,
        bool display,
        bool includeLeadingTab = false,
        bool replaceTarget = false)
    {
        var omml = TransformMathMlToOmml(mathMl);
        var tempPath = CreateTemporaryDocx(
            omml,
            includeLeadingTab: display && includeLeadingTab,
            forceInline: false);
        Document? sourceDocument = null;
        OMaths? sourceMaths = null;
        OMath? sourceMath = null;
        Range? sourceRange = null;
        OMath? insertedMath = null;
        Range? result = null;
        try
        {
            sourceDocument = application.Documents.Open(
                FileName: tempPath,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            sourceMaths = sourceDocument.OMaths;
            if (sourceMaths.Count != 1)
                throw new InvalidDataException(
                    "The temporary OMML document did not contain exactly one equation.");
            sourceMath = sourceMaths[1];
            if (display && includeLeadingTab)
            {
                var paragraph = sourceMath.Range.Paragraphs[1];
                try
                {
                    sourceRange = paragraph.Range.Duplicate;
                    sourceRange.End = Math.Max(sourceRange.Start, sourceRange.End - 1);
                }
                finally { Release(paragraph); }
            }
            else
            {
                sourceRange = sourceMath.Range;
            }

            if (!replaceTarget)
                insertionRange.Collapse(WdCollapseDirection.wdCollapseStart);
            var insertionStart = insertionRange.Start;
            insertionRange.FormattedText = sourceRange.FormattedText;
            insertedMath = FindMathAtPosition(
                    targetDocument,
                    insertionStart)
                ?? throw new InvalidOperationException(
                    "Word did not materialize the inserted OMML equation.");
            insertedMath.Type = display
                ? WdOMathType.wdOMathDisplay
                : WdOMathType.wdOMathInline;
            insertedMath.BuildUp();
            result = insertedMath.Range.Duplicate;
            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(result);
            Release(insertedMath);
            Release(sourceRange);
            Release(sourceMath);
            Release(sourceMaths);
            if (sourceDocument is not null)
            {
                try { sourceDocument.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(sourceDocument);
            try { File.Delete(tempPath); } catch { }
        }
    }

    internal static string TransformMathMlToOmml(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException("VisualTeX did not provide MathML for the Word OMML formula.");
        mathMl = NormalizeNaryArguments(mathMl);
        var display = IsBlockMathMl(mathMl);
        var transform = GetTransform();
        var inputSettings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            MaxCharactersInDocument = 4_000_000,
        };
        var outputSettings = transform.OutputSettings.Clone();
        outputSettings.OmitXmlDeclaration = true;
        outputSettings.Encoding = new UTF8Encoding(false);
        using var sourceText = new StringReader(mathMl);
        using var source = XmlReader.Create(sourceText, inputSettings);
        using var outputText = new StringWriter();
        using (var output = XmlWriter.Create(outputText, outputSettings))
            transform.Transform(source, output);
        var transformed = outputText.ToString();
        return NormalizeDisplayNaryOmml(ExtractSingleOMath(transformed), display);
    }

    internal static string NormalizeNaryArguments(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)) return mathMl;
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = false,
            MaxCharactersInDocument = 4_000_000,
        };
        using var text = new StringReader(mathMl);
        using var reader = XmlReader.Create(text, settings);
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        XNamespace mathMlNamespace = "http://www.w3.org/1998/Math/MathML";
        var limitNames = new HashSet<XName>
        {
            mathMlNamespace + "munder",
            mathMlNamespace + "mover",
            mathMlNamespace + "munderover",
            mathMlNamespace + "msub",
            mathMlNamespace + "msup",
            mathMlNamespace + "msubsup",
        };
        const string naryCharacters = "∑∏∐∫∬∭∮∯∰⋀⋁⋂⋃";
        var display = string.Equals(
            document.Root?.Attribute("display")?.Value,
            "block",
            StringComparison.OrdinalIgnoreCase);
        if (display)
        {
            foreach (var op in document
                         .Descendants(mathMlNamespace + "mo")
                         .Where(element =>
                             !string.IsNullOrEmpty(element.Value)
                             && element.Value.All(character => naryCharacters.IndexOf(character) >= 0)
                             && (element.Parent is null || !limitNames.Contains(element.Parent.Name)))
                         .ToList())
            {
                var argument = op.ElementsAfterSelf().FirstOrDefault();
                var syntheticLimit = new XElement(
                    mathMlNamespace + "msub",
                    new XElement(op),
                    new XElement(mathMlNamespace + "mrow"));
                op.ReplaceWith(syntheticLimit);
                if (argument is null)
                {
                    syntheticLimit.AddAfterSelf(
                        new XElement(
                            mathMlNamespace + "mrow",
                            new XElement(mathMlNamespace + "mspace", new XAttribute("width", "0em"))));
                }
                else if (argument.Name != mathMlNamespace + "mrow"
                         && argument.Name != mathMlNamespace + "mstyle")
                {
                    argument.ReplaceWith(new XElement(mathMlNamespace + "mrow", argument));
                }
            }
        }

        foreach (var limit in document.Descendants().Where(element => limitNames.Contains(element.Name)).ToList())
        {
            var op = limit.Elements().FirstOrDefault();
            if (op?.Name != mathMlNamespace + "mo"
                || string.IsNullOrEmpty(op.Value)
                || op.Value.Any(character => naryCharacters.IndexOf(character) < 0))
                continue;
            var argument = limit.ElementsAfterSelf().FirstOrDefault();
            if (argument is null
                || argument.Name == mathMlNamespace + "mrow"
                || argument.Name == mathMlNamespace + "mstyle")
                continue;

            // Office's MML2OMML.XSL recognizes an n-ary operand only when the
            // immediately following sibling is mrow or mstyle. MathJax emits a
            // valid flat Presentation MathML sequence (for example
            // munderover + mi), which Office otherwise converts to <m:e/> and
            // Word displays as a dotted placeholder box.
            argument.ReplaceWith(new XElement(mathMlNamespace + "mrow", argument));
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }

    private static bool IsBlockMathMl(string mathMl)
    {
        try
        {
            using var text = new StringReader(mathMl);
            using var reader = XmlReader.Create(text, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                MaxCharactersInDocument = 4_000_000,
            });
            var document = XDocument.Load(reader, LoadOptions.None);
            return string.Equals(
                document.Root?.Attribute("display")?.Value,
                "block",
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    internal static string NormalizeDisplayNaryOmml(string omml, bool display)
    {
        if (!display) return omml;
        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        foreach (var nary in document.Descendants(math + "nary"))
        {
            var properties = nary.Element(math + "naryPr");
            if (properties is null)
            {
                properties = new XElement(math + "naryPr");
                nary.AddFirst(properties);
            }
            var grow = properties.Element(math + "grow");
            if (grow is null)
            {
                grow = new XElement(math + "grow");
                properties.Add(grow);
            }
            grow.SetAttributeValue(math + "val", "1");

            SetNaryLimitVisibility(
                properties,
                math + "subHide",
                !HasNaryLimitContent(nary.Element(math + "sub")));
            SetNaryLimitVisibility(
                properties,
                math + "supHide",
                !HasNaryLimitContent(nary.Element(math + "sup")));
        }
        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    private static bool HasNaryLimitContent(XElement? limit)
    {
        if (limit is null) return false;
        return limit
            .DescendantsAndSelf()
            .Where(element => element.Name.LocalName == "t")
            .Any(element => !string.IsNullOrWhiteSpace(element.Value));
    }

    private static void SetNaryLimitVisibility(
        XElement properties,
        XName propertyName,
        bool hidden)
    {
        var property = properties.Element(propertyName);
        if (!hidden)
        {
            property?.Remove();
            return;
        }
        if (property is null)
        {
            property = new XElement(propertyName);
            properties.Add(property);
        }
        property.SetAttributeValue(XName.Get("val", MathNamespace), "1");
    }

    internal static string TransformOmmlToMathMl(string wordOpenXml, bool display)
    {
        var omml = ExtractSingleOMath(wordOpenXml);
        var transform = GetOmmlToMathMlTransform();
        var inputSettings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            MaxCharactersInDocument = 4_000_000,
        };
        var outputSettings = transform.OutputSettings.Clone();
        outputSettings.OmitXmlDeclaration = true;
        outputSettings.Encoding = new UTF8Encoding(false);
        using var sourceText = new StringReader(omml);
        using var source = XmlReader.Create(sourceText, inputSettings);
        using var outputText = new StringWriter();
        using (var output = XmlWriter.Create(outputText, outputSettings))
            transform.Transform(source, output);
        var transformed = outputText.ToString();
        using var transformedText = new StringReader(transformed);
        using var transformedReader = XmlReader.Create(transformedText, inputSettings);
        var document = XDocument.Load(transformedReader, LoadOptions.None);
        var root = document.Root?.Name.LocalName == "math"
            ? document.Root
            : document.Descendants().FirstOrDefault(element => element.Name.LocalName == "math");
        if (root is null)
            throw new InvalidDataException("Office OMML conversion did not produce a MathML math node.");
        root.SetAttributeValue("display", display ? "block" : "inline");
        return root.ToString(SaveOptions.DisableFormatting);
    }

    internal static string ComputeOmmlFingerprint(string wordOpenXml)
    {
        var normalized = ExtractSingleOMath(wordOpenXml);
        using var hash = SHA256.Create();
        var bytes = hash.ComputeHash(Encoding.UTF8.GetBytes(normalized));
        return string.Concat(bytes.Select(value => value.ToString("x2")));
    }

    internal static string ExtractSingleOMath(string omml)
    {
        if (string.IsNullOrWhiteSpace(omml))
            throw new InvalidDataException("Office produced an empty OMML transformation.");
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            MaxCharactersInDocument = 4_000_000,
        };
        using var text = new StringReader(omml);
        using var reader = XmlReader.Create(text, settings);
        var document = XDocument.Load(reader, LoadOptions.None);
        XNamespace math = MathNamespace;
        var equation = document.Root?.Name == math + "oMath"
            ? document.Root
            : document.Descendants(math + "oMath").FirstOrDefault();
        if (equation is null)
            throw new InvalidDataException("Office MathML conversion did not produce an m:oMath node.");
        return equation.ToString(SaveOptions.DisableFormatting);
    }

    internal static string BuildDocumentXml(
        string omml,
        bool includeLeadingTab = false,
        bool forceInline = false)
    {
        var equation = ExtractSingleOMath(omml);
        var prefix = includeLeadingTab ? "<w:r><w:tab/></w:r>" : string.Empty;
        if (forceInline)
        {
            // Surround the source equation with ordinary runs so Word opens it
            // as an inline OMath. Only sourceMath.Range is copied later; these
            // sentinels never enter the target document.
            prefix += "<w:r><w:t>L</w:t></w:r>";
        }
        var suffix = forceInline ? "<w:r><w:t>R</w:t></w:r>" : string.Empty;
        return $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{WordNamespace}\" xmlns:m=\"{MathNamespace}\">"
            + $"<w:body><w:p>{prefix}{equation}{suffix}</w:p><w:sectPr/></w:body></w:document>";
    }

    internal static string ResolveTransformPath() => ResolveTransformPath("MML2OMML.XSL");

    internal static string ResolveReverseTransformPath() => ResolveTransformPath("OMML2MML.XSL");

    private static string ResolveTransformPath(string fileName)
    {
        var candidates = new List<string>();
        AddCandidateRoot(candidates, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), fileName);
        AddCandidateRoot(candidates, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), fileName);
        AddCandidateRoot(candidates, Environment.GetEnvironmentVariable("ProgramW6432"), fileName);
        AddCandidateRoot(candidates, AppContext.BaseDirectory, fileName);
        foreach (var candidate in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (File.Exists(candidate)) return candidate;
        }
        throw new FileNotFoundException(
            $"Unable to locate Office {fileName}. Repair Microsoft Word or reinstall the Office integration.");
    }

    private static void AddCandidateRoot(List<string> candidates, string? root, string fileName)
    {
        if (string.IsNullOrWhiteSpace(root)) return;
        candidates.Add(Path.Combine(root, "Microsoft Office", "root", "Office16", fileName));
        candidates.Add(Path.Combine(root, "Office16", fileName));
        candidates.Add(Path.Combine(root, fileName));
    }

    private static XslCompiledTransform GetTransform()
    {
        lock (TransformLock)
        {
            if (_mathMlToOmml is not null) return _mathMlToOmml;
            _mathMlToOmml = LoadTransform(ResolveTransformPath());
            return _mathMlToOmml;
        }
    }

    private static XslCompiledTransform GetOmmlToMathMlTransform()
    {
        lock (TransformLock)
        {
            if (_ommlToMathMl is not null) return _ommlToMathMl;
            _ommlToMathMl = LoadTransform(ResolveReverseTransformPath());
            return _ommlToMathMl;
        }
    }

    private static XslCompiledTransform LoadTransform(string path)
    {
        var transform = new XslCompiledTransform(enableDebug: false);
        transform.Load(
            path,
            new XsltSettings(enableDocumentFunction: false, enableScript: false),
            null);
        return transform;
    }

    private static string CreateTemporaryDocx(
        string omml,
        bool includeLeadingTab,
        bool forceInline)
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
        WriteEntry(
            archive,
            "[Content_Types].xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
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
            "word/document.xml",
            BuildDocumentXml(omml, includeLeadingTab, forceInline));
        return path;
    }

    private static void WriteEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(content);
    }

    private static OMath? FindMathAtPosition(Document document, int position)
    {
        OMaths? maths = null;
        OMath? best = null;
        var bestDistance = int.MaxValue;
        try
        {
            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    if (range.Start < position) continue;
                    var distance = range.Start - position;
                    if (distance > 16 || distance >= bestDistance) continue;
                    Release(best);
                    best = math;
                    math = null;
                    bestDistance = distance;
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
            return best;
        }
        finally { Release(maths); }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
