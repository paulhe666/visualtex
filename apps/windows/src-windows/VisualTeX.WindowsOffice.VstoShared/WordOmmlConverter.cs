using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
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
    private const string NaryCharacters =
        "∑∏∐∫∬∭∮∯∰∱∲∳⨋⨌⨍⨎⨏⨐⨑⨒⨓⨔⨕⨖⨗⨘⨙⨚⨛⨜⋀⋁⋂⋃";
    private const string ExtendedIntegralCharacters =
        "∯∰∱∲∳⨋⨌⨍⨎⨏⨐⨑⨒⨓⨔⨕⨖⨗⨘⨙⨚⨛⨜";
    private const string FormulaBookmarkName = "VisualTeXFormula";
    private const string InlineScratchPlaceholder = "\uE001";
    private static readonly object TransformLock = new();
    private static readonly object InlineScratchLock = new();
    private static Document? _inlineScratchDocument;
    private static XslCompiledTransform? _mathMlToOmml;
    private static XslCompiledTransform? _ommlToMathMl;

    internal sealed class BatchSource : IDisposable
    {
        private Document? _document;
        private readonly string _path;
        private readonly IReadOnlyDictionary<string, BatchEntry> _entries;

        internal BatchSource(
            Document document,
            string path,
            IReadOnlyDictionary<string, BatchEntry> entries)
        {
            _document = document;
            _path = path;
            _entries = entries;
        }

        internal Range Insert(
            Document targetDocument,
            Range insertionRange,
            string formulaId,
            bool display,
            out string sourceFingerprint,
            bool replaceTarget)
        {
            var sourceDocument = _document
                ?? throw new ObjectDisposedException(nameof(BatchSource));
            if (!_entries.TryGetValue(formulaId, out var entry))
                throw new InvalidDataException(
                    $"The OMML batch source does not contain formula {formulaId}.");

            Bookmarks? bookmarks = null;
            Bookmark? bookmark = null;
            Range? sourceRange = null;
            Range? formattedSource = null;
            Range? target = null;
            OMath? insertedMath = null;
            Range? result = null;
            try
            {
                bookmarks = sourceDocument.Bookmarks;
                if (!bookmarks.Exists(entry.BookmarkName))
                    throw new InvalidDataException(
                        $"The OMML batch bookmark {entry.BookmarkName} is missing.");
                bookmark = bookmarks[entry.BookmarkName];
                sourceRange = bookmark.Range;
                formattedSource = sourceRange.FormattedText;
                target = insertionRange.Duplicate;
                if (!replaceTarget)
                    target.Collapse(WdCollapseDirection.wdCollapseStart);
                var insertionStart = target.Start;
                target.FormattedText = formattedSource;
                insertedMath = FindMathAtPosition(
                        targetDocument,
                        insertionStart,
                        target.End)
                    ?? throw new InvalidOperationException(
                        "Word did not materialize the batch OMML equation.");
                var targetType = display
                    ? WdOMathType.wdOMathDisplay
                    : WdOMathType.wdOMathInline;
                if (insertedMath.Type != targetType)
                    insertedMath.Type = targetType;
                result = insertedMath.Range.Duplicate;
                sourceFingerprint = entry.SourceFingerprint;
                var returned = result;
                result = null;
                return returned;
            }
            finally
            {
                Release(result);
                Release(insertedMath);
                Release(target);
                Release(formattedSource);
                Release(sourceRange);
                Release(bookmark);
                Release(bookmarks);
            }
        }

        public void Dispose()
        {
            var document = _document;
            _document = null;
            if (document is not null)
            {
                try { document.Saved = true; } catch { }
                try { document.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
                Release(document);
            }
            try { File.Delete(_path); } catch { }
        }
    }

    internal sealed class WholeDocumentSource : IDisposable
    {
        private Document? _document;
        private readonly string _path;

        internal WholeDocumentSource(Document document, string path)
        {
            _document = document;
            _path = path;
        }

        internal Range Insert(Document targetDocument, Range insertionRange)
        {
            var sourceDocument = _document
                ?? throw new ObjectDisposedException(nameof(WholeDocumentSource));
            Range? sourceRange = null;
            Range? formattedSource = null;
            Range? target = null;
            Range? result = null;
            try
            {
                sourceRange = sourceDocument.Content.Duplicate;
                // Exclude the source document's final paragraph mark/section
                // boundary while retaining the explicit VisualTeX end marker.
                if (sourceRange.End > sourceRange.Start)
                    sourceRange.End--;
                formattedSource = sourceRange.FormattedText;
                target = insertionRange.Duplicate;
                target.Collapse(WdCollapseDirection.wdCollapseStart);
                var insertionStart = target.Start;
                target.FormattedText = formattedSource;
                result = targetDocument.Range(insertionStart, target.End);
                var returned = result;
                result = null;
                return returned;
            }
            finally
            {
                Release(result);
                Release(target);
                Release(formattedSource);
                Release(sourceRange);
            }
        }

        public void Dispose()
        {
            var document = _document;
            _document = null;
            if (document is not null)
            {
                try { document.Saved = true; } catch { }
                try { document.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
                Release(document);
            }
            try { File.Delete(_path); } catch { }
        }
    }

    internal static WholeDocumentSource CreateWholeDocumentSource(
        Application application,
        string documentXml)
    {
        if (application is null) throw new ArgumentNullException(nameof(application));
        if (string.IsNullOrWhiteSpace(documentXml))
            throw new InvalidDataException("The bulk OMML document XML is empty.");
        var path = CreateTemporaryDocumentDocx(documentXml);
        Document? document = null;
        try
        {
            document = application.Documents.Open(
                FileName: path,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            var source = new WholeDocumentSource(document, path);
            document = null;
            return source;
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
                Release(document);
                try { File.Delete(path); } catch { }
            }
        }
    }

    internal sealed class BatchEntry
    {
        internal BatchEntry(
            string bookmarkName,
            string sourceFingerprint,
            string omml,
            int bookmarkId)
        {
            BookmarkName = bookmarkName;
            SourceFingerprint = sourceFingerprint;
            Omml = omml;
            BookmarkId = bookmarkId;
        }

        internal string BookmarkName { get; }
        internal string SourceFingerprint { get; }
        internal string Omml { get; }
        internal int BookmarkId { get; }
    }

    internal static BatchSource CreateBatchSource(
        Application application,
        IReadOnlyList<(string FormulaId, string MathMl)> formulas)
    {
        if (application is null) throw new ArgumentNullException(nameof(application));
        if (formulas is null) throw new ArgumentNullException(nameof(formulas));
        if (formulas.Count == 0)
            throw new ArgumentOutOfRangeException(
                nameof(formulas),
                "At least one OMML formula is required for a batch source.");

        var entries = new Dictionary<string, BatchEntry>(
            formulas.Count,
            StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < formulas.Count; index++)
        {
            var formula = formulas[index];
            if (string.IsNullOrWhiteSpace(formula.FormulaId))
                throw new InvalidDataException("The OMML batch formula id is missing.");
            var omml = TransformMathMlToOmml(formula.MathMl);
            entries.Add(
                formula.FormulaId,
                new BatchEntry(
                    $"VisualTeXBatch{index:D4}",
                    ComputeOmmlFingerprint(omml),
                    omml,
                    index));
        }

        var path = CreateTemporaryBatchDocx(entries.Values.ToList());
        Document? document = null;
        try
        {
            document = application.Documents.Open(
                FileName: path,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            var source = new BatchSource(document, path, entries);
            document = null;
            return source;
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
                Release(document);
                try { File.Delete(path); } catch { }
            }
        }
    }

    internal static Range Insert(
        Application application,
        Document targetDocument,
        Range insertionRange,
        string mathMl,
        bool display,
        out string sourceFingerprint,
        bool includeLeadingTab = false,
        bool replaceTarget = false)
    {
        var omml = TransformMathMlToOmml(mathMl);
        sourceFingerprint = ComputeOmmlFingerprint(omml);
        var tempPath = CreateTemporaryDocx(
            omml,
            includeLeadingTab: display && includeLeadingTab,
            forceInline: !display);
        // Import the bookmarked formula range directly from the DOCX. This keeps
        // the native OMML fidelity of the old path without opening and closing a
        // second hidden Word document for every edit.
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_FAST_OMML_INSERT"),
                "0",
                StringComparison.Ordinal))
        {
            try
            {
                var imported = !display && replaceTarget
                    ? InsertBookmarkedFileThroughScratchDocument(
                        application,
                        targetDocument,
                        insertionRange,
                        tempPath)
                    : InsertBookmarkedFile(
                        targetDocument,
                        insertionRange,
                        tempPath,
                        display,
                        replaceTarget);
                try { File.Delete(tempPath); } catch { }
                return imported;
            }
            catch (Exception error)
            {
                if (string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                    Console.WriteLine(
                        $"    [perf] WordOmmlConverter InsertFile fallback: {error.Message}");
                // Fall through to the proven formatted-text transfer path.
            }
        }
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
                    insertionStart,
                    insertionRange.End)
                ?? throw new InvalidOperationException(
                    "Word did not materialize the inserted OMML equation.");
            insertedMath.Type = display
                ? WdOMathType.wdOMathDisplay
                : WdOMathType.wdOMathInline;
            // sourceRange already comes from a fully built native OMML tree.
            // Re-running BuildUp asks Word to parse professional OMML as if it
            // were linear equation text and can introduce dotted placeholder
            // slots, especially around matrices and nested scripts.
            result = insertedMath.Range.Duplicate;
            RemoveImportedFormulaBookmark(targetDocument, result);
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

    private static Range InsertBookmarkedFileThroughScratchDocument(
        Application application,
        Document targetDocument,
        Range insertionRange,
        string filePath)
    {
        lock (InlineScratchLock)
            return InsertBookmarkedFileThroughScratchDocumentCore(
                application,
                targetDocument,
                insertionRange,
                filePath);
    }

    private static Range InsertBookmarkedFileThroughScratchDocumentCore(
        Application application,
        Document targetDocument,
        Range insertionRange,
        string filePath)
    {
        Document? scratchDocument = null;
        Range? scratchInsertion = null;
        OMaths? scratchMaths = null;
        OMath? scratchMath = null;
        Range? scratchFormula = null;
        Range? target = null;
        OMath? insertedMath = null;
        Range? result = null;
        try
        {
            scratchDocument = GetOrCreateInlineScratchDocument(application);
            scratchInsertion = scratchDocument.Content;
            var scratchStart = scratchInsertion.Start;
            scratchInsertion.Text = "L" + InlineScratchPlaceholder + "R";
            Release(scratchInsertion);
            scratchInsertion = scratchDocument.Range(
                scratchStart + 1,
                scratchStart + 1 + InlineScratchPlaceholder.Length);
            // The scratch document itself supplies stable ordinary L/R context.
            // Import only the bookmarked equation, avoiding a full DOCX insertion
            // while still forcing Word to materialize a genuine inline OMath.
            scratchInsertion.InsertFile(
                FileName: filePath,
                Range: FormulaBookmarkName,
                ConfirmConversions: false,
                Link: false,
                Attachment: false);
            scratchMaths = scratchDocument.OMaths;
            if (scratchMaths.Count != 1)
                throw new InvalidDataException(
                    "The OMML scratch document did not contain exactly one equation.");
            scratchMath = scratchMaths[1];
            scratchFormula = scratchMath.Range;

            target = insertionRange.Duplicate;
            if (target.Start >= target.End)
                throw new InvalidOperationException(
                    "Inline OMML replacement requires a non-collapsed placeholder range.");
            var insertionStart = target.Start;
            target.FormattedText = scratchFormula.FormattedText;
            insertedMath = FindMathAtPosition(
                    targetDocument,
                    insertionStart,
                    target.End)
                ?? throw new InvalidOperationException(
                    "Word did not materialize the inline OMML replacement.");
            if (insertedMath.Type != WdOMathType.wdOMathInline)
                insertedMath.Type = WdOMathType.wdOMathInline;
            result = insertedMath.Range.Duplicate;
            RemoveImportedFormulaBookmark(targetDocument, result);
            try { scratchDocument.Saved = true; } catch { }
            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(result);
            Release(insertedMath);
            Release(target);
            Release(scratchFormula);
            Release(scratchMath);
            Release(scratchMaths);
            Release(scratchInsertion);
            // The hidden scratch document is intentionally cached for the life
            // of this Word process. Reusing it removes Documents.Add/Close from
            // every inline edit; it is marked Saved so Word can quit silently.
            try { targetDocument.Activate(); } catch { }
        }
    }

    private static Document GetOrCreateInlineScratchDocument(Application application)
    {
        if (_inlineScratchDocument is not null)
        {
            Application? scratchApplication = null;
            try
            {
                scratchApplication = _inlineScratchDocument.Application;
                if (IsSameComObject(scratchApplication, application))
                    return _inlineScratchDocument;
            }
            catch
            {
                // The previous Word process has exited or invalidated its RCW.
            }
            finally { Release(scratchApplication); }

            try
            {
                _inlineScratchDocument.Saved = true;
                _inlineScratchDocument.Close(WdSaveOptions.wdDoNotSaveChanges);
            }
            catch { }
            Release(_inlineScratchDocument);
            _inlineScratchDocument = null;
        }

        _inlineScratchDocument = application.Documents.Add(Visible: false);
        try { _inlineScratchDocument.Saved = true; } catch { }
        return _inlineScratchDocument;
    }

    private static Range InsertBookmarkedFile(
        Document targetDocument,
        Range insertionRange,
        string filePath,
        bool display,
        bool replaceTarget)
    {
        Range? target = null;
        OMath? insertedMath = null;
        Range? result = null;
        try
        {
            target = insertionRange.Duplicate;
            var insertionStart = target.Start;
            if (!replaceTarget)
                target.Collapse(WdCollapseDirection.wdCollapseStart);
            // InsertFile replaces a non-collapsed Range in place. Do not delete
            // the placeholder first: collapsing at its former boundary lets Word
            // move the imported OMath behind the adjacent typing sentinel.
            target.InsertFile(
                FileName: filePath,
                Range: FormulaBookmarkName,
                ConfirmConversions: false,
                Link: false,
                Attachment: false);
            insertedMath = FindMathAtPosition(
                    targetDocument,
                    insertionStart,
                    target.End)
                ?? throw new InvalidOperationException(
                    "Word did not materialize the bookmarked OMML equation.");
            var targetType = display
                ? WdOMathType.wdOMathDisplay
                : WdOMathType.wdOMathInline;
            if (insertedMath.Type != targetType)
                insertedMath.Type = targetType;
            // The imported DOCX already contains professional OMML. Calling
            // BuildUp again reparses an already-built tree and adds substantial
            // latency without changing the equation structure.
            result = insertedMath.Range.Duplicate;
            RemoveImportedFormulaBookmark(targetDocument, result);
            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(result);
            Release(insertedMath);
            Release(target);
        }
    }

    private static void RemoveImportedFormulaBookmark(
        Document targetDocument,
        Range insertedRange)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        try
        {
            bookmarks = targetDocument.Bookmarks;
            if (!bookmarks.Exists(FormulaBookmarkName)) return;
            bookmark = bookmarks[FormulaBookmarkName];
            bookmarkRange = bookmark.Range;
            var overlapsInsertedEquation = bookmarkRange.Start <= insertedRange.End
                && bookmarkRange.End >= insertedRange.Start;
            if (overlapsInsertedEquation) bookmark.Delete();
        }
        catch
        {
            // This bookmark exists only to locate the formula inside the temporary
            // DOCX. The native equation remains valid if cleanup is rejected by
            // an older Word build.
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static Range InsertXmlDirect(
        Document targetDocument,
        Range insertionRange,
        string omml,
        bool display,
        bool replaceTarget)
    {
        Range? target = null;
        Range? content = null;
        Range? probe = null;
        OMaths? maths = null;
        OMath? selected = null;
        Range? selectedRange = null;
        try
        {
            target = insertionRange.Duplicate;
            if (!replaceTarget)
                target.Collapse(WdCollapseDirection.wdCollapseStart);
            var insertionStart = target.Start;
            target.InsertXML(omml);

            content = targetDocument.Content;
            object probeStart = Math.Max(content.Start, insertionStart - 1);
            object probeEnd = Math.Min(
                content.End,
                Math.Max(target.End + 2, insertionStart + 256));
            probe = targetDocument.Range(ref probeStart, ref probeEnd);
            maths = probe.OMaths;
            var bestDistance = int.MaxValue;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? candidate = null;
                Range? candidateRange = null;
                try
                {
                    candidate = maths[index];
                    candidateRange = candidate.Range;
                    var distance = Math.Abs(candidateRange.Start - insertionStart);
                    if (distance >= bestDistance) continue;
                    Release(selectedRange);
                    Release(selected);
                    selected = candidate;
                    candidate = null;
                    selectedRange = candidateRange;
                    candidateRange = null;
                    bestDistance = distance;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }
            if (selected is null || selectedRange is null)
                throw new InvalidOperationException(
                    "Word did not materialize the directly inserted OMML equation.");
            selected.Type = display
                ? WdOMathType.wdOMathDisplay
                : WdOMathType.wdOMathInline;
            selected.BuildUp();
            var result = selected.Range.Duplicate;
            return result;
        }
        finally
        {
            Release(selectedRange);
            Release(selected);
            Release(maths);
            Release(probe);
            Release(content);
            Release(target);
        }
    }

    internal static string TransformMathMlToOmml(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException("VisualTeX did not provide MathML for the Word OMML formula.");
        mathMl = RemoveVisualTeXBoundaryArtifactsFromMathMl(mathMl);
        ValidateMathMlForOmml(mathMl);
        mathMl = NormalizeFencedMathMlTables(mathMl);
        mathMl = NormalizeNestedEmptyBaseScripts(mathMl);
        mathMl = NormalizeMathMlAccents(mathMl);
        mathMl = NormalizeNaryArguments(mathMl);
        var placeholderResult = ReplaceExtendedIntegralsWithOfficePlaceholders(mathMl);
        mathMl = placeholderResult.MathMl;
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
        var outputSettings = transform.OutputSettings?.Clone() ?? new XmlWriterSettings();
        outputSettings.OmitXmlDeclaration = true;
        outputSettings.Encoding = new UTF8Encoding(false);
        using var sourceText = new StringReader(mathMl);
        using var source = XmlReader.Create(sourceText, inputSettings);
        using var outputText = new StringWriter();
        using (var output = XmlWriter.Create(outputText, outputSettings))
            transform.Transform(source, output);
        var transformed = outputText.ToString();
        var omml = ExtractSingleOMath(transformed);
        omml = RestoreExtendedIntegralCharacters(omml, placeholderResult.NaryCharacters);
        omml = NormalizeExplicitUprightRuns(omml, mathMl);
        omml = NormalizeExplicitTableColumnAlignment(omml, mathMl);
        omml = NormalizeOmmlPlaceholderVisibility(omml);
        omml = NormalizeDisplayNaryOmml(omml, display);
        ValidateOmmlResult(omml, mathMl);
        return omml;
    }

    private static string RemoveVisualTeXBoundaryArtifactsFromMathMl(string mathMl)
    {
        static bool IsBoundaryArtifact(char character) =>
            character is '\u200B' or '\u200C' or '\u2060' or '\uFEFF';

        if (mathMl.IndexOf('\u200B') < 0
            && mathMl.IndexOf('\u200C') < 0
            && mathMl.IndexOf('\u2060') < 0
            && mathMl.IndexOf('\uFEFF') < 0
            && mathMl.IndexOf("200B", StringComparison.OrdinalIgnoreCase) < 0
            && mathMl.IndexOf("200C", StringComparison.OrdinalIgnoreCase) < 0
            && mathMl.IndexOf("2060", StringComparison.OrdinalIgnoreCase) < 0
            && mathMl.IndexOf("FEFF", StringComparison.OrdinalIgnoreCase) < 0)
            return mathMl;

        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        var affectedAncestors = new HashSet<XElement>();
        foreach (var textNode in document.DescendantNodes().OfType<XText>().ToList())
        {
            var original = textNode.Value;
            if (!original.Any(IsBoundaryArtifact)) continue;
            foreach (var ancestor in textNode.Ancestors())
                affectedAncestors.Add(ancestor);
            textNode.Value = new string(
                original.Where(character => !IsBoundaryArtifact(character)).ToArray());
        }

        foreach (var element in document.Descendants().Reverse().ToList())
        {
            if (!affectedAncestors.Contains(element)) continue;
            if (element.Elements().Any()) continue;
            if (!string.IsNullOrWhiteSpace(element.Value)) continue;
            if (element.Parent is null) continue;
            element.Remove();
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }

    internal static void ValidateMathMlForOmml(string mathMl)
    {
        // The common path contains semantic mi/mo/mn nodes only. Avoid an
        // additional XML parse for every Word edit unless MathJax emitted one
        // of the permissive fallback/error constructs that can hide a command.
        if (mathMl.IndexOf("<mtext", StringComparison.OrdinalIgnoreCase) < 0
            && mathMl.IndexOf("<merror", StringComparison.OrdinalIgnoreCase) < 0
            && mathMl.IndexOf("mathcolor", StringComparison.OrdinalIgnoreCase) < 0)
            return;
        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var unresolved = document
            .Descendants(presentationMath + "mtext")
            .FirstOrDefault(element =>
                Regex.IsMatch(element.Value, @"\\[A-Za-z@]+")
                || string.Equals(
                    element.Attribute("mathcolor")?.Value,
                    "red",
                    StringComparison.OrdinalIgnoreCase));
        if (unresolved is not null)
        {
            throw new InvalidDataException(
                "MathML contains an unresolved LaTeX command and cannot be inserted as OMML: "
                + unresolved.Value.Trim());
        }
        if (document.Descendants(presentationMath + "merror").Any())
            throw new InvalidDataException(
                "MathML contains an error node and cannot be inserted as OMML.");
    }

    internal static string NormalizeFencedMathMlTables(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)
            || mathMl.IndexOf("<mtable", StringComparison.OrdinalIgnoreCase) < 0)
            return mathMl;
        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var fallbackFenceCharacters = new HashSet<string>(StringComparer.Ordinal)
        {
            "(", ")", "[", "]", "{", "}", "|", "‖", "⌈", "⌉", "⌊", "⌋",
            "⟨", "⟩", "/", "\\", "↑", "↓", "↕", "⇑", "⇓", "⇕",
        };

        bool IsFenceOperator(XElement element, string expectedTexClass)
        {
            if (element.Name != presentationMath + "mo") return false;
            var texClass = element.Attribute("data-mjx-texclass")?.Value;
            if (string.Equals(texClass, expectedTexClass, StringComparison.OrdinalIgnoreCase))
                return true;
            if (string.Equals(element.Attribute("fence")?.Value, "true", StringComparison.OrdinalIgnoreCase)
                || string.Equals(element.Attribute("stretchy")?.Value, "true", StringComparison.OrdinalIgnoreCase))
                return true;
            return fallbackFenceCharacters.Contains(element.Value);
        }

        foreach (var row in document.Descendants(presentationMath + "mrow").ToList())
        {
            var children = row.Elements().ToArray();
            if (children.Length != 3
                || children[1].Name != presentationMath + "mtable"
                || !IsFenceOperator(children[0], "OPEN")
                || !IsFenceOperator(children[2], "CLOSE"))
                continue;
            var open = children[0].Value;
            var close = children[2].Value;
            if (open.Length == 0 && close.Length == 0)
                continue;

            // MathJax serializes matrices and other fenced multi-line structures
            // as OPEN mo + mtable + CLOSE mo. For one-sided delimiters (notably
            // cases and \left...\right.) the invisible side is an empty stretchy
            // mo. Office's MML2OMML transform treats the visible mo as an ordinary
            // glyph unless the structure is normalized to mfenced first, which
            // makes Word emit one native m:d whose delimiter grows with the table.
            row.ReplaceWith(
                new XElement(
                    presentationMath + "mfenced",
                    new XAttribute("open", open),
                    new XAttribute("close", close),
                    new XAttribute("separators", string.Empty),
                    new XElement(children[1])));
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }

    private static bool IsMatrixLikeMathMlTable(XElement table, XNamespace presentationMath)
    {
        var parent = table.Parent;
        if (parent?.Name == presentationMath + "mfenced")
        {
            var fencedRows = table.Elements(presentationMath + "mtr").ToArray();
            if (fencedRows.Length == 0) return false;
            var fencedColumns = fencedRows[0].Elements(presentationMath + "mtd").Count();
            // Office may represent a one-column fenced stack as an equation
            // array rather than m:m. The surrounding m:d is still correct and
            // stretchable, so matrix-dimension validation is only meaningful
            // for real multi-column tables.
            return fencedColumns > 1
                && fencedRows.All(row =>
                    row.Elements(presentationMath + "mtd").Count() == fencedColumns);
        }

        // MathJax also uses mtable for aligned equations, cases and substack.
        // Those structures legitimately become m:eqArr or limit constructs in
        // OMML, so only validate table shapes that carry matrix-like alignment.
        if (string.Equals(
                table.Attribute("displaystyle")?.Value,
                "false",
                StringComparison.OrdinalIgnoreCase)
            || table.Attribute("scriptlevel") is not null)
            return false;
        var alignments = (table.Attribute("columnalign")?.Value ?? string.Empty)
            .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (alignments.Any(value =>
                string.Equals(value, "left", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "right", StringComparison.OrdinalIgnoreCase)))
            return false;
        var rows = table.Elements(presentationMath + "mtr").ToArray();
        if (rows.Length == 0) return false;
        var columns = rows[0].Elements(presentationMath + "mtd").Count();
        // A one-column mtable is the normal MathJax representation of
        // \substack and must become an OMML limit stack rather than m:m.
        return columns > 1
            && rows.All(row =>
                row.Elements(presentationMath + "mtd").Count() == columns);
    }

    internal static void ValidateOmmlResult(string omml, string mathMl)
    {
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        XNamespace officeMath = MathNamespace;

        // Backslashes do not occur in normal OMML output. Only pay the XML
        // traversal cost when one is actually present, then restrict the check
        // to visible m:t runs to avoid interpreting XML metadata as formula text.
        if (omml.IndexOf('\\') >= 0)
        {
            var commandTarget = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
            var literalCommand = commandTarget
                .Descendants(officeMath + "t")
                .FirstOrDefault(text => Regex.IsMatch(text.Value, @"\\[A-Za-z@]+"));
            if (literalCommand is not null)
                throw new InvalidDataException(
                    "OMML contains unresolved LaTeX command text: " + literalCommand.Value.Trim());
        }

        // Matrix integrity is the only structural validation that needs both
        // trees. Most formulas in normal editing contain no table at all, so
        // keep that path allocation-free to preserve the 100-formula latency.
        if (mathMl.IndexOf("<mtable", StringComparison.OrdinalIgnoreCase) < 0)
            return;
        var source = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        var sourceTables = source
            .Descendants(presentationMath + "mtable")
            .Where(table => IsMatrixLikeMathMlTable(table, presentationMath))
            .ToArray();
        if (sourceTables.Length == 0) return;
        var target = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        ValidateNoVisibleEmptyOmmlSlots(target);
        var targetMatrices = target.Descendants(officeMath + "m").ToArray();
        foreach (var matrix in targetMatrices)
        {
            var placeholderVisibility = matrix
                .Element(officeMath + "mPr")?
                .Element(officeMath + "plcHide")?
                .Attribute(officeMath + "val")?
                .Value;
            if (!string.Equals(placeholderVisibility, "1", StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Office matrix OMML did not hide dotted placeholder slots.");
        }
        var matchedMatrices = new HashSet<XElement>();
        for (var tableIndex = 0; tableIndex < sourceTables.Length; tableIndex++)
        {
            var sourceRows = sourceTables[tableIndex]
                .Elements(presentationMath + "mtr")
                .ToArray();
            var sourceColumnCounts = sourceRows
                .Select(row => row.Elements(presentationMath + "mtd").Count())
                .ToArray();
            var targetMatrix = targetMatrices.FirstOrDefault(candidate =>
            {
                if (matchedMatrices.Contains(candidate)) return false;
                var candidateRows = candidate.Elements(officeMath + "mr").ToArray();
                return candidateRows.Length == sourceRows.Length
                    && candidateRows.Select(row => row.Elements(officeMath + "e").Count())
                        .SequenceEqual(sourceColumnCounts);
            });
            if (targetMatrix is null)
                throw new InvalidDataException(
                    $"Office did not preserve matrix {tableIndex + 1} dimensions "
                    + $"({sourceRows.Length}x{string.Join("/", sourceColumnCounts)})." );
            matchedMatrices.Add(targetMatrix);
            var targetRows = targetMatrix.Elements(officeMath + "mr").ToArray();
            for (var rowIndex = 0; rowIndex < sourceRows.Length; rowIndex++)
            {
                var sourceCells = sourceRows[rowIndex]
                    .Elements(presentationMath + "mtd")
                    .ToArray();
                var targetCells = targetRows[rowIndex]
                    .Elements(officeMath + "e")
                    .ToArray();
                for (var cellIndex = 0; cellIndex < sourceCells.Length; cellIndex++)
                {
                    var sourceHasVisibleContent = sourceCells[cellIndex]
                        .DescendantsAndSelf()
                        .Any(element =>
                            element.Name.Namespace == presentationMath
                            && element.Name.LocalName is "mi" or "mn" or "mo" or "mtext"
                            && !string.IsNullOrWhiteSpace(element.Value));
                    if (!sourceHasVisibleContent) continue;
                    var targetHasVisibleContent = targetCells[cellIndex]
                        .Descendants(officeMath + "t")
                        .Any(text => !string.IsNullOrWhiteSpace(text.Value));
                    if (!targetHasVisibleContent)
                        throw new InvalidDataException(
                            $"Office produced an empty matrix slot at matrix {tableIndex + 1}, "
                            + $"row {rowIndex + 1}, column {cellIndex + 1}.");
                }
            }
        }
    }

    internal static string NormalizeExplicitUprightRuns(string omml, string mathMl)
    {
        if (string.IsNullOrWhiteSpace(omml) || string.IsNullOrWhiteSpace(mathMl))
            return omml;

        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var mathMlDocument = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        var uprightTokens = mathMlDocument
            .Descendants()
            .Where(element =>
            {
                if (element.Name.Namespace != presentationMath) return false;
                var variant = element.Attribute("mathvariant")?.Value ?? string.Empty;
                return variant.IndexOf("normal", StringComparison.OrdinalIgnoreCase) >= 0
                    || variant.IndexOf("upright", StringComparison.OrdinalIgnoreCase) >= 0;
            })
            .Select(element => element.Value)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToHashSet(StringComparer.Ordinal);
        if (uprightTokens.Count == 0) return omml;

        var ommlDocument = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;
        foreach (var run in ommlDocument.Descendants(math + "r"))
        {
            var text = string.Concat(run.Elements(math + "t").Select(element => element.Value));
            if (!uprightTokens.Contains(text)) continue;

            var properties = run.Element(math + "rPr");
            var plainStyle = properties?.Element(math + "sty");
            if (plainStyle?.Attribute(math + "val")?.Value != "p") continue;

            plainStyle.Remove();
            if (properties!.Element(math + "nor") is null)
                properties.AddFirst(new XElement(math + "nor"));

            // Upright identifiers such as e, i and d are mathematical tokens,
            // not prose. Word otherwise spell-checks the normal-style run and
            // paints a red proofing underline below an otherwise correct OMML
            // equation. Preserve the math style while explicitly disabling
            // proofing for this run.
            var wordProperties = run.Element(word + "rPr");
            if (wordProperties is null)
            {
                wordProperties = new XElement(word + "rPr");
                var mathProperties = run.Element(math + "rPr");
                if (mathProperties is not null) mathProperties.AddAfterSelf(wordProperties);
                else run.AddFirst(wordProperties);
            }
            if (wordProperties.Element(word + "noProof") is null)
                wordProperties.Add(new XElement(word + "noProof"));
        }

        return ommlDocument.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    internal static string NormalizeMathMlAccents(string mathMl)
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
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";

        var accentCharacters = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["^"] = "\u0302",
            ["~"] = "\u0303",
            ["→"] = "\u20D7",
            ["←"] = "\u20D6",
            ["↔"] = "\u20E1",
            ["¯"] = "\u0305",
            ["‾"] = "\u0305",
            ["―"] = "\u0305",
            ["ˉ"] = "\u0305",
            ["˙"] = "\u0307",
            ["¨"] = "\u0308",
            ["ˇ"] = "\u030C",
            ["˘"] = "\u0306",
            ["´"] = "\u0301",
            ["`"] = "\u0300",
            ["˚"] = "\u030A",
        };

        foreach (var mover in document.Descendants(presentationMath + "mover").ToList())
        {
            var children = mover.Elements().ToArray();
            if (children.Length != 2) continue;
            var mark = children[1];
            if (mark.Name != presentationMath + "mo") continue;
            if (string.Equals(
                    mark.Attribute("accent")?.Value,
                    "false",
                    StringComparison.OrdinalIgnoreCase)
                || string.Equals(
                    mover.Attribute("accent")?.Value,
                    "false",
                    StringComparison.OrdinalIgnoreCase))
                continue;

            var sourceCharacter = mark.Value;
            if (!accentCharacters.TryGetValue(sourceCharacter, out var combiningCharacter))
                continue;

            // Office's MML2OMML.XSL only creates a native m:acc node when the
            // MathML mover is explicitly marked as an accent and the mark is a
            // combining accent character. MathJax emits spacing characters
            // such as ^, ˙ and → without accent=true, which Office otherwise
            // converts into m:limUpp or replacement glyphs/placeholder boxes.
            mover.SetAttributeValue("accent", "true");
            mark.SetAttributeValue("accent", "true");
            mark.SetAttributeValue("stretchy", null);
            mark.SetAttributeValue("data-mjx-pseudoscript", null);
            mark.Value = combiningCharacter;
        }

        return document.ToString(SaveOptions.DisableFormatting);
    }

    internal static string NormalizeNestedEmptyBaseScripts(string mathMl)
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
        var simpleScriptNames = new HashSet<XName>
        {
            mathMlNamespace + "msub",
            mathMlNamespace + "msup",
        };
        var allScriptNames = new HashSet<XName>(simpleScriptNames)
        {
            mathMlNamespace + "msubsup",
        };
        var transparentWrappers = new HashSet<XName>
        {
            mathMlNamespace + "mrow",
            mathMlNamespace + "mstyle",
            mathMlNamespace + "mpadded",
            mathMlNamespace + "mphantom",
            mathMlNamespace + "semantics",
        };

        bool IsEmptyMathNode(XElement element)
        {
            if (element.Name == mathMlNamespace + "mspace") return false;
            if (element.Nodes().OfType<XText>().Any(node => !string.IsNullOrWhiteSpace(node.Value)))
                return false;
            var children = element.Elements().ToArray();
            return children.Length == 0 || children.All(IsEmptyMathNode);
        }

        bool IsOnlyContentOfOuterScriptArgument(XElement candidate)
        {
            XElement current = candidate;
            while (current.Parent is XElement parent)
            {
                if (transparentWrappers.Contains(parent.Name))
                {
                    if (parent.Elements().Any(sibling =>
                            sibling != current && !IsEmptyMathNode(sibling)))
                        return false;
                    if (parent.Nodes().OfType<XText>().Any(node =>
                            !string.IsNullOrWhiteSpace(node.Value)))
                        return false;
                    current = parent;
                    continue;
                }
                if (!allScriptNames.Contains(parent.Name)) return false;
                var children = parent.Elements().ToList();
                var position = children.IndexOf(current);
                return position >= 1;
            }
            return false;
        }

        foreach (var script in document
                     .Descendants()
                     .Where(element => simpleScriptNames.Contains(element.Name))
                     .Reverse()
                     .ToList())
        {
            var children = script.Elements().ToArray();
            if (children.Length < 2
                || !IsEmptyMathNode(children[0])
                || !IsOnlyContentOfOuterScriptArgument(script))
                continue;

            // MathJax represents sources such as f_{_{\\mathrm H}} as an
            // outer subscript whose argument contains another subscript with
            // an empty base. Office faithfully renders that empty base as a
            // dotted equation placeholder. Inside an existing script slot the
            // extra empty-base level carries no useful layout information, so
            // replace it with its visible script argument. Standalone empty-
            // base scripts are intentionally preserved for prescript/tensor
            // notation.
            script.ReplaceWith(new XElement(children[1]));
        }
        return document.ToString(SaveOptions.DisableFormatting);
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
        foreach (var op in document
                     .Descendants(mathMlNamespace + "mo")
                     .Where(element =>
                         !string.IsNullOrEmpty(element.Value)
                         && element.Value.All(character => NaryCharacters.IndexOf(character) >= 0)
                         && (element.Parent is null || !limitNames.Contains(element.Parent.Name)))
                     .ToList())
        {
            // Office only creates a native m:nary for a bare operator when it
            // is carried by a limit structure. Apply that structure to inline
            // formulas too; later normalization hides the synthetic empty
            // limit, while only display equations receive m:grow=1.
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

        foreach (var limit in document.Descendants().Where(element => limitNames.Contains(element.Name)).ToList())
        {
            var op = limit.Elements().FirstOrDefault();
            if (op?.Name != mathMlNamespace + "mo"
                || string.IsNullOrEmpty(op.Value)
                || op.Value.Any(character => NaryCharacters.IndexOf(character) < 0))
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

    internal static (string MathMl, IReadOnlyList<string> NaryCharacters)
        ReplaceExtendedIntegralsWithOfficePlaceholders(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            return (mathMl, Array.Empty<string>());

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
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var operators = document
            .Descendants(presentationMath + "mo")
            .Where(element =>
                element.Value.Length == 1
                && NaryCharacters.IndexOf(element.Value[0]) >= 0)
            .ToArray();
        var characters = operators.Select(element => element.Value).ToArray();

        foreach (var op in operators)
        {
            if (ExtendedIntegralCharacters.IndexOf(op.Value[0]) >= 0)
            {
                // Office's MML2OMML transform knows how to attach limits and
                // the following operand to a standard integral. Use it only as
                // a structural placeholder; the exact extended character is
                // restored in OMML immediately after the transform.
                op.Value = "∫";
            }
        }

        return (document.ToString(SaveOptions.DisableFormatting), characters);
    }

    internal static string RestoreExtendedIntegralCharacters(
        string omml,
        IReadOnlyList<string> sourceNaryCharacters)
    {
        if (sourceNaryCharacters.Count == 0
            || !sourceNaryCharacters.Any(character =>
                character.Length == 1
                && ExtendedIntegralCharacters.IndexOf(character[0]) >= 0))
            return omml;

        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        var naries = document.Descendants(math + "nary").ToArray();
        if (naries.Length != sourceNaryCharacters.Count)
        {
            throw new InvalidDataException(
                "Office changed the number of n-ary operators while converting extended integrals. "
                + $"MathML={sourceNaryCharacters.Count}; OMML={naries.Length}.");
        }

        for (var index = 0; index < sourceNaryCharacters.Count; index++)
        {
            var sourceCharacter = sourceNaryCharacters[index];
            if (sourceCharacter.Length != 1
                || ExtendedIntegralCharacters.IndexOf(sourceCharacter[0]) < 0)
                continue;

            var properties = naries[index].Element(math + "naryPr");
            if (properties is null)
            {
                properties = new XElement(math + "naryPr");
                naries[index].AddFirst(properties);
            }
            var character = properties.Element(math + "chr");
            if (character is null)
            {
                character = new XElement(math + "chr");
                properties.AddFirst(character);
            }
            character.SetAttributeValue(math + "val", sourceCharacter);
        }

        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
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

    internal static string NormalizeExplicitTableColumnAlignment(
        string omml,
        string mathMl)
    {
        if (string.IsNullOrWhiteSpace(omml) || string.IsNullOrWhiteSpace(mathMl))
            return omml;

        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        XNamespace officeMath = MathNamespace;
        var sourceDocument = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        var targetDocument = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var sourceTables = sourceDocument
            .Descendants(presentationMath + "mtable")
            .Where(table => !string.IsNullOrWhiteSpace(table.Attribute("columnalign")?.Value))
            .Select(table =>
            {
                var rows = table.Elements(presentationMath + "mtr").ToArray();
                var columnCount = rows.Length == 0
                    ? 0
                    : rows.Max(row => row.Elements(presentationMath + "mtd").Count());
                var raw = (table.Attribute("columnalign")?.Value ?? string.Empty)
                    .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                if (columnCount <= 0 || raw.Length == 0) return null;
                var alignments = Enumerable.Range(0, columnCount)
                    .Select(index => NormalizeMathMlColumnAlignment(
                        raw[Math.Min(index, raw.Length - 1)]))
                    .ToArray();
                return new
                {
                    RowCount = rows.Length,
                    ColumnCount = columnCount,
                    Alignments = alignments,
                };
            })
            .Where(table => table is not null)
            .ToArray();
        if (sourceTables.Length == 0) return omml;

        var matrices = targetDocument.Descendants(officeMath + "m").ToList();
        var used = new HashSet<XElement>();
        foreach (var sourceTable in sourceTables)
        {
            var target = matrices.FirstOrDefault(matrix =>
            {
                if (used.Contains(matrix)) return false;
                var rows = matrix.Elements(officeMath + "mr").ToArray();
                if (rows.Length != sourceTable!.RowCount) return false;
                return rows.All(row =>
                    row.Elements(officeMath + "e").Count() == sourceTable.ColumnCount);
            });
            if (target is null) continue;
            used.Add(target);

            var properties = target.Element(officeMath + "mPr");
            if (properties is null)
            {
                properties = new XElement(officeMath + "mPr");
                target.AddFirst(properties);
            }
            var columns = new XElement(officeMath + "mcs");
            foreach (var alignment in sourceTable!.Alignments)
            {
                columns.Add(
                    new XElement(
                        officeMath + "mc",
                        new XElement(
                            officeMath + "mcPr",
                            new XElement(
                                officeMath + "count",
                                new XAttribute(officeMath + "val", "1")),
                            new XElement(
                                officeMath + "mcJc",
                                new XAttribute(officeMath + "val", alignment)))));
            }
            var existing = properties.Element(officeMath + "mcs");
            if (existing is null) properties.Add(columns);
            else existing.ReplaceWith(columns);
        }

        return targetDocument.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    internal static void ValidateNoVisibleEmptyOmmlSlots(XDocument document)
    {
        XNamespace math = MathNamespace;
        var structuralSlotNames = new HashSet<XName>
        {
            math + "e",
            math + "sub",
            math + "sup",
            math + "num",
            math + "den",
            math + "lim",
            math + "fName",
            math + "deg",
        };
        foreach (var slot in document.Descendants()
                     .Where(element => structuralSlotNames.Contains(element.Name)))
        {
            if (slot.Descendants(math + "t")
                .Any(text => !string.IsNullOrWhiteSpace(text.Value)))
                continue;

            var parent = slot.Parent;
            var hidden = slot.Name == math + "deg"
                && parent?.Name == math + "rad"
                && string.Equals(
                    parent.Element(math + "radPr")?
                        .Element(math + "degHide")?
                        .Attribute(math + "val")?
                        .Value,
                    "1",
                    StringComparison.Ordinal)
                || slot.Name == math + "sub"
                && parent?.Name == math + "nary"
                && string.Equals(
                    parent.Element(math + "naryPr")?
                        .Element(math + "subHide")?
                        .Attribute(math + "val")?
                        .Value,
                    "1",
                    StringComparison.Ordinal)
                || slot.Name == math + "sup"
                && parent?.Name == math + "nary"
                && string.Equals(
                    parent.Element(math + "naryPr")?
                        .Element(math + "supHide")?
                        .Attribute(math + "val")?
                        .Value,
                    "1",
                    StringComparison.Ordinal);
            if (hidden) continue;

            throw new InvalidDataException(
                $"Office OMML contains a visible empty {slot.Name.LocalName} slot "
                + $"inside {parent?.Name.LocalName ?? "unknown"}.");
        }
    }

    internal static string NormalizeOmmlPlaceholderVisibility(string omml)
    {
        if (string.IsNullOrWhiteSpace(omml)
            || omml.IndexOf("<m:m", StringComparison.Ordinal) < 0)
            return omml;
        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        foreach (var matrix in document.Descendants(math + "m"))
        {
            var properties = matrix.Element(math + "mPr");
            if (properties is null)
            {
                properties = new XElement(math + "mPr");
                matrix.AddFirst(properties);
            }
            var placeholderVisibility = properties.Element(math + "plcHide");
            if (placeholderVisibility is null)
            {
                placeholderVisibility = new XElement(math + "plcHide");
                properties.AddFirst(placeholderVisibility);
            }
            placeholderVisibility.SetAttributeValue(math + "val", "1");
        }
        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    private static string NormalizeMathMlColumnAlignment(string value)
    {
        return value.Trim().ToLowerInvariant() switch
        {
            "left" => "left",
            "right" => "right",
            _ => "center",
        };
    }

    internal static string NormalizeDisplayNaryOmml(string omml, bool display)
    {
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

            // Only display equations should force a growing integral/sum. Empty
            // limits, however, must be hidden for both inline and display OMML.
            // The previous early return skipped inline formulas entirely, so a
            // bare inline integral could expose Word's dotted sub/sup slots.
            var grow = properties.Element(math + "grow");
            if (display)
            {
                if (grow is null)
                {
                    grow = new XElement(math + "grow");
                    properties.Add(grow);
                }
                grow.SetAttributeValue(math + "val", "1");
            }
            else
            {
                // Office's MML2OMML transform may add grow=1 even for inline
                // n-ary operators. Remove it so the operator follows the Word
                // paragraph font size instead of taking display proportions.
                grow?.Remove();
            }

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

    internal static void ValidateMaterializedOmml(string wordOpenXml)
    {
        var equation = ExtractSingleOMath(wordOpenXml);
        var document = XDocument.Parse(equation, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;

        foreach (var matrix in document.Descendants(math + "m"))
        {
            var hidePlaceholders = IsMathBooleanEnabled(
                matrix.Element(math + "mPr"),
                math + "plcHide");
            foreach (var cell in matrix
                         .Elements(math + "mr")
                         .SelectMany(row => row.Elements(math + "e")))
            {
                if (HasVisibleMathContent(cell) || hidePlaceholders) continue;
                throw new InvalidDataException(
                    "Word materialized an unhidden empty OMML matrix/alignment cell.");
            }
        }

        foreach (var radical in document.Descendants(math + "rad"))
        {
            var properties = radical.Element(math + "radPr");
            var degree = radical.Element(math + "deg");
            if (degree is not null
                && !HasVisibleMathContent(degree)
                && !IsMathBooleanEnabled(properties, math + "degHide"))
                throw new InvalidDataException(
                    "Word materialized an unhidden empty OMML radical degree slot.");
            RequireVisibleMathContent(
                radical.Element(math + "e"),
                "radical body");
        }

        foreach (var nary in document.Descendants(math + "nary"))
        {
            var properties = nary.Element(math + "naryPr");
            var sub = nary.Element(math + "sub");
            var sup = nary.Element(math + "sup");
            if (sub is not null
                && !HasVisibleMathContent(sub)
                && !IsMathBooleanEnabled(properties, math + "subHide"))
                throw new InvalidDataException(
                    "Word materialized an unhidden empty OMML n-ary subscript slot.");
            if (sup is not null
                && !HasVisibleMathContent(sup)
                && !IsMathBooleanEnabled(properties, math + "supHide"))
                throw new InvalidDataException(
                    "Word materialized an unhidden empty OMML n-ary superscript slot.");
            RequireVisibleMathContent(
                nary.Element(math + "e"),
                "n-ary operand");
        }

        foreach (var script in document.Descendants().Where(element =>
                     element.Name == math + "sSub"
                     || element.Name == math + "sSup"
                     || element.Name == math + "sSubSup"))
        {
            RequireVisibleMathContent(script.Element(math + "e"), "script base");
            if (script.Name == math + "sSub" || script.Name == math + "sSubSup")
                RequireVisibleMathContent(script.Element(math + "sub"), "subscript");
            if (script.Name == math + "sSup" || script.Name == math + "sSubSup")
                RequireVisibleMathContent(script.Element(math + "sup"), "superscript");
        }

        foreach (var fraction in document.Descendants(math + "f"))
        {
            RequireVisibleMathContent(fraction.Element(math + "num"), "fraction numerator");
            RequireVisibleMathContent(fraction.Element(math + "den"), "fraction denominator");
        }
    }

    private static bool HasVisibleMathContent(XElement? slot)
    {
        if (slot is null) return false;
        return slot
            .DescendantsAndSelf()
            .Where(element => element.Name.LocalName == "t")
            .Any(element => !string.IsNullOrWhiteSpace(element.Value));
    }

    private static bool IsMathBooleanEnabled(XElement? properties, XName propertyName)
    {
        var property = properties?.Element(propertyName);
        if (property is null) return false;
        var value = property.Attribute(XName.Get("val", MathNamespace))?.Value;
        return string.IsNullOrWhiteSpace(value)
            || value == "1"
            || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
    }

    private static void RequireVisibleMathContent(XElement? slot, string description)
    {
        if (HasVisibleMathContent(slot)) return;
        throw new InvalidDataException(
            $"Word materialized an empty OMML {description} slot.");
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
        var outputSettings = transform.OutputSettings?.Clone() ?? new XmlWriterSettings();
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

        // MML2OMML's reverse stylesheet may serialize the result as <mml:math>.
        // That is XML-equivalent MathML, but VisualTeX's existing MathType bridge
        // intentionally accepts the canonical <math xmlns="..."> contract used by
        // the renderer. Rebuild the presentation tree with one default MathML
        // namespace so OMML -> MathType can use the exact same validated path.
        var canonicalRoot = CanonicalizeMathMlElement(root);
        RestoreOmmlAccentSemantics(omml, canonicalRoot);
        canonicalRoot.SetAttributeValue("display", display ? "block" : "inline");
        return canonicalRoot.ToString(SaveOptions.DisableFormatting);
    }

    private static void RestoreOmmlAccentSemantics(string omml, XElement mathMlRoot)
    {
        XNamespace officeMath = MathNamespace;
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        var source = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var sourceAccents = source.Descendants(officeMath + "acc").ToArray();
        if (sourceAccents.Length == 0) return;

        // Office's bundled OMML->MathML stylesheet still contains legacy symbol
        // mappings for m:acc. In current Word documents that can turn a perfectly
        // valid Unicode combining accent (for example U+0305/U+20D7/U+0307) into
        // '-', '?' or mojibake. The OMML itself is the authoritative source, so
        // restore the accent mark after the XSL transform before MathType/LaTeX
        // consumers see it.
        var targetAccents = mathMlRoot
            .DescendantsAndSelf(mathMl + "mover")
            .Where(element => string.Equals(
                (string?)element.Attribute("accent"),
                "true",
                StringComparison.OrdinalIgnoreCase))
            .ToArray();
        var count = Math.Min(sourceAccents.Length, targetAccents.Length);
        for (var index = 0; index < count; index++)
        {
            var sourceAccent = sourceAccents[index];
            var officeCharacter = sourceAccent
                .Element(officeMath + "accPr")?
                .Element(officeMath + "chr")?
                .Attribute(officeMath + "val")?
                .Value;
            var canonicalCharacter = CanonicalMathMlAccentCharacter(officeCharacter);
            var targetAccent = targetAccents[index];
            var children = targetAccent.Elements().ToArray();
            if (children.Length < 2) continue;
            children[1].ReplaceWith(
                new XElement(
                    mathMl + "mo",
                    new XAttribute("accent", "true"),
                    canonicalCharacter));
        }
    }

    private static string CanonicalMathMlAccentCharacter(string? officeCharacter) =>
        officeCharacter switch
        {
            null or "" => "^", // OMML's omitted m:chr means the default hat.
            "\u0302" => "^",
            "\u0303" => "~",
            "\u20D7" => "→",
            "\u20D6" => "←",
            "\u20E1" => "↔",
            "\u0305" => "¯",
            "\u0307" => "˙",
            "\u0308" => "¨",
            "\u030C" => "ˇ",
            "\u0306" => "˘",
            "\u0301" => "´",
            "\u0300" => "`",
            "\u030A" => "˚",
            _ => officeCharacter,
        };

    private static XElement CanonicalizeMathMlElement(XElement source)
    {
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        var result = new XElement(mathMl + source.Name.LocalName);
        foreach (var attribute in source.Attributes())
        {
            if (attribute.IsNamespaceDeclaration) continue;
            result.SetAttributeValue(attribute.Name, attribute.Value);
        }
        foreach (var node in source.Nodes())
        {
            switch (node)
            {
                case XElement child:
                    result.Add(CanonicalizeMathMlElement(child));
                    break;
                case XCData cdata:
                    result.Add(new XCData(cdata.Value));
                    break;
                case XText text:
                    result.Add(new XText(text.Value));
                    break;
            }
        }
        return result;
    }

    internal static string ComputeOmmlFingerprint(string wordOpenXml)
    {
        var normalized = ExtractSingleOMath(wordOpenXml);
        var document = XDocument.Parse(normalized, LoadOptions.PreserveWhitespace);
        XNamespace word = WordNamespace;
        XNamespace math = MathNamespace;

        // Word stores the visible math size and other proof/font state in
        // ordinary run properties. These are presentation state, not formula
        // content, and Word may add them while importing the same OMML.
        document.Descendants(word + "rPr").Remove();
        document.Descendants(math + "ctrlPr").Remove();
        document.Descendants(word + "bookmarkStart").Remove();
        document.Descendants(word + "bookmarkEnd").Remove();
        NormalizeMathRunGrouping(document, math);

        // Word stores the visible math size in ordinary run properties. Font
        // size is presentation state, not formula content: changing 14 pt to
        // 18 pt must not force an OMML -> MathML -> LaTeX source refresh.
        document
            .Descendants()
            .Where(element => element.Name == word + "sz" || element.Name == word + "szCs")
            .Remove();

        normalized = document.Root?.ToString(SaveOptions.DisableFormatting) ?? normalized;
        using var hash = SHA256.Create();
        var bytes = hash.ComputeHash(Encoding.UTF8.GetBytes(normalized));
        return string.Concat(bytes.Select(value => value.ToString("x2")));
    }

    private static void NormalizeMathRunGrouping(XDocument document, XNamespace math)
    {
        foreach (var run in document.Descendants(math + "r").ToArray())
        {
            var texts = run.Elements(math + "t").ToArray();
            if (texts.Length > 1)
            {
                var combined = string.Concat(texts.Select(text => text.Value));
                texts[0].Value = combined;
                foreach (var extra in texts.Skip(1)) extra.Remove();
            }
            var properties = run.Element(math + "rPr");
            if (properties is not null && !properties.HasElements && !properties.HasAttributes)
                properties.Remove();
        }

        foreach (var parent in document.Root?.DescendantsAndSelf().ToArray()
                     ?? Array.Empty<XElement>())
        {
            XElement? previousRun = null;
            string? previousKey = null;
            foreach (var child in parent.Elements().ToArray())
            {
                if (child.Name != math + "r")
                {
                    previousRun = null;
                    previousKey = null;
                    continue;
                }
                var key = child.Element(math + "rPr")?.ToString(SaveOptions.DisableFormatting)
                    ?? string.Empty;
                if (previousRun is not null
                    && string.Equals(previousKey, key, StringComparison.Ordinal))
                {
                    var previousText = previousRun.Element(math + "t");
                    var currentText = child.Element(math + "t");
                    if (previousText is not null && currentText is not null)
                    {
                        previousText.Value += currentText.Value;
                        child.Remove();
                        continue;
                    }
                }
                previousRun = child;
                previousKey = key;
            }
        }
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

    private static string CreateTemporaryDocumentDocx(string documentXml)
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
        WriteEntry(archive, "word/document.xml", documentXml);
        return path;
    }

    private static string CreateTemporaryBatchDocx(
        IReadOnlyList<BatchEntry> entries)
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

        var body = new StringBuilder();
        foreach (var entry in entries)
        {
            var equation = ExtractSingleOMath(entry.Omml);
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
        WriteEntry(
            archive,
            "word/document.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{WordNamespace}\" xmlns:m=\"{MathNamespace}\">"
            + $"<w:body>{body}<w:sectPr/></w:body></w:document>");
        return path;
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

    private static OMath? FindMathAtPosition(
        Document document,
        int position,
        int preferredEnd)
    {
        Range? content = null;
        Range? probe = null;
        OMaths? maths = null;
        OMath? best = null;
        var bestSpan = -1;
        var bestDistance = int.MaxValue;
        try
        {
            content = document.Content;
            object probeStart = Math.Max(content.Start, position - 1);
            object probeEnd = Math.Min(
                content.End,
                Math.Max(preferredEnd + 2, position + 1024));
            probe = document.Range(ref probeStart, ref probeEnd);
            maths = probe.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    // Word exposes nested matrix rows as OMath objects in a local
                    // range. Select the largest equation that overlaps the exact
                    // insertion target so the bookmark wraps the top-level math,
                    // not an inner row such as e_x,e_y,e_z.
                    if (range.End <= position || range.Start > preferredEnd + 1) continue;
                    var span = range.End - range.Start;
                    var distance = Math.Abs(range.Start - position);
                    if (distance > 16
                        || span < bestSpan
                        || (span == bestSpan && distance >= bestDistance))
                        continue;
                    Release(best);
                    best = math;
                    math = null;
                    bestSpan = span;
                    bestDistance = distance;
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
            if (best is not null) return best;

            // Defensive fallback for an unusual Word range expansion. Normal
            // insertions resolve through the local probe above, avoiding an
            // O(n) enumeration of every equation in a 100-formula document.
            Release(maths);
            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    if (range.End <= position || range.Start > preferredEnd + 1) continue;
                    var span = range.End - range.Start;
                    var distance = Math.Abs(range.Start - position);
                    if (distance > 16
                        || span < bestSpan
                        || (span == bestSpan && distance >= bestDistance))
                        continue;
                    Release(best);
                    best = math;
                    math = null;
                    bestSpan = span;
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
        finally
        {
            Release(maths);
            Release(probe);
            Release(content);
        }
    }

    private static bool IsSameComObject(object left, object right)
    {
        IntPtr leftIdentity = IntPtr.Zero;
        IntPtr rightIdentity = IntPtr.Zero;
        try
        {
            leftIdentity = Marshal.GetIUnknownForObject(left);
            rightIdentity = Marshal.GetIUnknownForObject(right);
            return leftIdentity == rightIdentity;
        }
        catch { return false; }
        finally
        {
            if (rightIdentity != IntPtr.Zero) Marshal.Release(rightIdentity);
            if (leftIdentity != IntPtr.Zero) Marshal.Release(leftIdentity);
        }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
