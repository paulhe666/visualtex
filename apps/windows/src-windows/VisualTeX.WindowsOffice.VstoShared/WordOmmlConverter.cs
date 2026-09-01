using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using VisualTeX.WindowsOffice.Contracts;
using Application = Microsoft.Office.Interop.Word.Application;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static class WordOmmlConverter
{
    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr windowHandle);

    private const int MaximumFormulaXmlCharacters = 16 * 1024 * 1024;
    private const int MaximumFormulaXmlDepth = 256;
    private const int MaximumFormulaXmlElements = 250_000;

    private static XmlReader CreateSafeXmlReader(string xml, string kind)
    {
        if (xml.Length > MaximumFormulaXmlCharacters)
            throw new InvalidDataException(
                $"{kind} exceeds the supported safety limit of {MaximumFormulaXmlCharacters} characters.");
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = MaximumFormulaXmlCharacters,
            IgnoreWhitespace = false,
            CloseInput = true,
        };

        var elements = 0;
        using (var preflight = XmlReader.Create(new StringReader(xml), settings))
        {
            while (preflight.Read())
            {
                if (preflight.NodeType != XmlNodeType.Element) continue;
                if (preflight.Depth >= MaximumFormulaXmlDepth)
                    throw new InvalidDataException(
                        $"{kind} nesting exceeds the supported safety limit of {MaximumFormulaXmlDepth} levels.");
                if (++elements > MaximumFormulaXmlElements)
                    throw new InvalidDataException(
                        $"{kind} contains more than the supported safety limit of {MaximumFormulaXmlElements} elements.");
            }
        }
        return XmlReader.Create(new StringReader(xml), settings);
    }

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
    private const string VisualTeXEquationNumberPlaceholderPrefix = "981730";
    private const string VisualTeXNativeNumberBookmarkPrefix = "VTEqNum_";
    private const string VisualTeXEquationSequenceName = "VisualTeXEquation";
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
        private readonly string _mathFontName;

        internal BatchSource(
            Document document,
            string path,
            IReadOnlyDictionary<string, BatchEntry> entries,
            string mathFontName)
        {
            _document = document;
            _path = path;
            _entries = entries;
            _mathFontName = mathFontName;
        }

        internal string GetSourceFingerprint(string formulaId)
        {
            if (!_entries.TryGetValue(formulaId, out var entry))
                throw new InvalidDataException(
                    $"The OMML batch source does not contain formula {formulaId}.");
            return entry.SourceFingerprint;
        }

        internal IReadOnlyList<Range> InsertAdjacentInlineGroup(
            Application application,
            Document targetDocument,
            Range targetRange,
            IReadOnlyList<string> formulaIds)
        {
            if (formulaIds is null || formulaIds.Count < 2)
                throw new ArgumentOutOfRangeException(
                    nameof(formulaIds),
                    "An adjacent OMML group requires at least two formulas.");
            var groupEntries = new List<BatchEntry>(formulaIds.Count);
            foreach (var formulaId in formulaIds)
            {
                if (!_entries.TryGetValue(formulaId, out var entry))
                    throw new InvalidDataException(
                        $"The OMML batch source does not contain formula {formulaId}.");
                groupEntries.Add(entry);
            }

            var path = CreateTemporaryAdjacentInlineGroupDocx(
                groupEntries,
                _mathFontName);
            Document? sourceDocument = null;
            Range? sourceRange = null;
            Range? target = null;
            OMaths? maths = null;
            OMath? math = null;
            Range? mathRange = null;
            var results = new List<Range>(formulaIds.Count);
            try
            {
                sourceDocument = application.Documents.Open(
                    FileName: path,
                    ConfirmConversions: false,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    Visible: false,
                    OpenAndRepair: false);
                sourceRange = sourceDocument.Content.Duplicate;
                if (sourceRange.End > sourceRange.Start)
                    sourceRange.End--;
                target = targetRange.Duplicate;
                var insertionStart = target.Start;
                target.FormattedText = sourceRange.FormattedText;
                var insertionEnd = target.End;

                maths = targetDocument.OMaths;
                var candidates = new List<(int Start, int End)>();
                for (var index = 1; index <= maths.Count; index++)
                {
                    Release(mathRange); mathRange = null;
                    Release(math); math = maths[index];
                    mathRange = math.Range;
                    if (mathRange.Start < insertionStart || mathRange.End > insertionEnd)
                        continue;
                    candidates.Add((mathRange.Start, mathRange.End));
                }
                candidates.Sort((left, right) => left.Start.CompareTo(right.Start));
                if (candidates.Count != formulaIds.Count)
                    throw new InvalidOperationException(
                        $"Word materialized {candidates.Count} OMath objects for an adjacent group of {formulaIds.Count} formulas.");
                foreach (var candidate in candidates)
                    results.Add(targetDocument.Range(candidate.Start, candidate.End));
                return results;
            }
            catch
            {
                foreach (var result in results) Release(result);
                throw;
            }
            finally
            {
                Release(mathRange);
                Release(math);
                Release(maths);
                Release(target);
                Release(sourceRange);
                if (sourceDocument is not null)
                {
                    try { sourceDocument.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
                }
                Release(sourceDocument);
                try { File.Delete(path); } catch { }
            }
        }

        internal IReadOnlyList<Range> ReplaceDisplayParagraphGroup(
            Application application,
            Document targetDocument,
            Range targetRange,
            IReadOnlyList<string> formulaIds)
        {
            if (formulaIds is null || formulaIds.Count == 0)
                throw new ArgumentOutOfRangeException(
                    nameof(formulaIds),
                    "A display OMML group requires at least one formula.");
            var entries = new List<BatchEntry>(formulaIds.Count);
            foreach (var formulaId in formulaIds)
            {
                if (!_entries.TryGetValue(formulaId, out var entry))
                    throw new InvalidDataException(
                        $"The OMML batch source does not contain formula {formulaId}.");
                entries.Add(entry);
            }

            var path = CreateTemporaryDisplayGroupDocx(entries, _mathFontName);
            Document? sourceDocument = null;
            Range? sourceRange = null;
            Range? formattedSource = null;
            Range? target = null;
            Range? insertedRange = null;
            OMaths? sourceMaths = null;
            OMaths? maths = null;
            OMath? math = null;
            Range? mathRange = null;
            var results = new List<Range>(formulaIds.Count);
            try
            {
                sourceDocument = application.Documents.Open(
                    FileName: path,
                    ConfirmConversions: false,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    Visible: false,
                    OpenAndRepair: false);
                sourceRange = sourceDocument.Content.Duplicate;
                // Preserve every formula paragraph mark while excluding only the
                // source document's final empty paragraph mark/section boundary.
                // A bookmark whose endpoints sit inside paragraphs does not carry
                // those boundaries reliably through cross-document FormattedText;
                // Word then merges an adjacent display group into one OMath.
                if (sourceRange.End > sourceRange.Start)
                    sourceRange.End--;
                sourceMaths = sourceRange.OMaths;
                if (sourceMaths.Count != formulaIds.Count)
                    throw new InvalidDataException(
                        $"The temporary display group contains {sourceMaths.Count}/{formulaIds.Count} OMath objects.");
                formattedSource = sourceRange.FormattedText;

                target = targetRange.Duplicate;
                var insertionStart = target.Start;
                // Replace the entire contiguous MathType owner range in one
                // FormattedText assignment. This still makes Word tear down every
                // Equation.DSMT4/MTPlaceRef tree as one transaction, while unlike
                // Range.InsertFile the live target Range expands to the complete
                // multi-paragraph payload rather than only the first OMath.
                target.FormattedText = formattedSource;
                try { targetDocument.Activate(); } catch { }
                var insertionEnd = target.End;
                insertedRange = targetDocument.Range(insertionStart, insertionEnd);
                maths = insertedRange.OMaths;
                if (maths.Count != formulaIds.Count)
                    throw new InvalidOperationException(
                        $"Word materialized {maths.Count} display OMath objects for a group of {formulaIds.Count} formulas.");
                for (var index = 1; index <= maths.Count; index++)
                {
                    Release(mathRange); mathRange = null;
                    Release(math); math = maths[index];
                    if (math.Type != WdOMathType.wdOMathDisplay)
                        math.Type = WdOMathType.wdOMathDisplay;
                    mathRange = math.Range.Duplicate;
                    if (mathRange.Start < insertionStart
                        || mathRange.End > insertionEnd)
                        throw new InvalidOperationException(
                            "A grouped display OMath escaped the atomic replacement range.");
                    results.Add(targetDocument.Range(mathRange.Start, mathRange.End));
                }
                return results;
            }
            catch
            {
                foreach (var result in results) Release(result);
                throw;
            }
            finally
            {
                Release(mathRange);
                Release(math);
                Release(maths);
                Release(sourceMaths);
                Release(insertedRange);
                Release(target);
                Release(formattedSource);
                Release(sourceRange);
                if (sourceDocument is not null)
                {
                    try { sourceDocument.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
                }
                Release(sourceDocument);
                try { File.Delete(path); } catch { }
            }
        }

        internal Range ReplaceTargetParagraphAtomicallyWithCleanParagraph(
            Document targetDocument,
            Range targetParagraphRange)
        {
            var sourceDocument = _document
                ?? throw new ObjectDisposedException(nameof(BatchSource));
            Paragraphs? sourceParagraphs = null;
            Paragraph? sourceParagraph = null;
            Range? sourceRange = null;
            Range? target = null;
            Range? result = null;
            try
            {
                sourceParagraphs = sourceDocument.Paragraphs;
                if (sourceParagraphs.Count == 0)
                    throw new InvalidDataException(
                        "The OMML batch source has no clean terminal paragraph.");
                sourceParagraph = sourceParagraphs[sourceParagraphs.Count];
                sourceRange = sourceParagraph.Range.Duplicate;
                if (!string.Equals(sourceRange.Text, "\r", StringComparison.Ordinal)
                    || sourceRange.OMaths.Count != 0
                    || sourceRange.InlineShapes.Count != 0
                    || sourceRange.Fields.Count != 0
                    || sourceRange.Tables.Count != 0)
                    throw new InvalidDataException(
                        "The OMML batch source terminal paragraph is not structurally empty.");

                target = targetParagraphRange.Duplicate;
                var start = target.Start;
                if (target.Paragraphs.Count != 1
                    || target.End <= target.Start
                    || !string.Equals(
                        target.Text?.Substring(Math.Max(0, target.Text.Length - 1)),
                        "\r",
                        StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "The MathType source is not one complete Word paragraph.");

                // Replace the complete, prevalidated MathType display paragraph in
                // one Word operation. This lets Word tear down the Equation.DSMT4
                // OLE and the outer MTPlaceRef/nested sequence tree as one owner,
                // instead of exposing any partially deleted field hierarchy.
                target.FormattedText = sourceRange.FormattedText;
                result = targetDocument.Range(start, start);
                var returned = result;
                result = null;
                return returned;
            }
            finally
            {
                Release(result);
                Release(target);
                Release(sourceRange);
                Release(sourceParagraph);
                Release(sourceParagraphs);
            }
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
                formattedSource = sourceRange.FormattedText
                    ?? throw new InvalidDataException(
                        $"The OMML batch bookmark {entry.BookmarkName} has no formatted source range.");
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
        string documentXml,
        string? mathFontName = null)
    {
        if (application is null) throw new ArgumentNullException(nameof(application));
        if (string.IsNullOrWhiteSpace(documentXml))
            throw new InvalidDataException("The bulk OMML document XML is empty.");
        var path = CreateTemporaryDocumentDocx(
            documentXml,
            NormalizeMathFontName(mathFontName));
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
        IReadOnlyList<(string FormulaId, string MathMl)> formulas,
        Func<string, string, string>? transformOmml = null,
        string? mathFontName = null)
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
            if (transformOmml is not null)
                omml = transformOmml(formula.FormulaId, omml);
            entries.Add(
                formula.FormulaId,
                new BatchEntry(
                    $"VisualTeXBatch{index:D4}",
                    ComputeOmmlFingerprint(omml),
                    omml,
                    index));
        }

        var normalizedMathFontName = NormalizeMathFontName(mathFontName);
        var path = CreateTemporaryBatchDocx(
            entries.Values.ToList(),
            normalizedMathFontName);
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
            var source = new BatchSource(
                document,
                path,
                entries,
                normalizedMathFontName);
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
        bool replaceTarget = false,
        Func<string, string>? transformOmml = null,
        string? mathFontName = null)
    {
        var omml = TransformMathMlToOmml(mathMl);
        if (transformOmml is not null)
            omml = transformOmml(omml);
        sourceFingerprint = ComputeOmmlFingerprint(omml);
        var normalizedMathFontName = NormalizeMathFontName(mathFontName);
        var tempPath = CreateTemporaryDocx(
            omml,
            includeLeadingTab: display && includeLeadingTab,
            forceInline: !display,
            mathFontName: normalizedMathFontName);
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
                var directInlineReplacement = !display
                    && replaceTarget
                    && !string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_DIRECT_INLINE_OMML_INSERT"),
                        "0",
                        StringComparison.Ordinal);
                var imported = !display && replaceTarget && !directInlineReplacement
                    ? InsertBookmarkedFileThroughScratchDocument(
                        application,
                        targetDocument,
                        insertionRange,
                        tempPath,
                        normalizedMathFontName)
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

    internal static float? MeasurePreparedDisplayHeightPoints(
        Application application,
        Document targetDocument,
        string omml,
        string? mathFontName = null)
    {
        if (application is null) throw new ArgumentNullException(nameof(application));
        if (targetDocument is null) throw new ArgumentNullException(nameof(targetDocument));
        if (string.IsNullOrWhiteSpace(omml)) return null;

        var normalizedMathFontName = NormalizeMathFontName(mathFontName);
        string semanticOmml;
        try
        {
            semanticOmml = ApplyExplicitTransferMathFont(
                ExtractSingleOMath(omml),
                normalizedMathFontName);
        }
        catch
        {
            return null;
        }
        Range? content = null;
        Range? separatorInsertion = null;
        Range? formulaInsertion = null;
        Range? measuredRange = null;
        Range? cleanupRange = null;
        Window? window = null;
        Microsoft.Office.Interop.Word.View? view = null;
        Zoom? zoom = null;
        var scratchBoundary = -1;
        var measurementStage = "prepare-target-scratch";
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        int? previousVerticalScroll = null;
        int? previousHorizontalScroll = null;
        try
        {
            targetDocument.Activate();
            try
            {
                previousScreenUpdating = application.ScreenUpdating;
                application.ScreenUpdating = false;
                screenUpdatingSuspended = true;
            }
            catch { }
            content = targetDocument.Content;
            scratchBoundary = Math.Max(content.Start, content.End - 1);
            Release(content);
            content = null;

            // Create one ordinary terminal paragraph without touching the current
            // selection. The caller runs this inside its existing Word undo record
            // with ScreenUpdating disabled. Everything from scratchBoundary to the
            // final document mark is removed in finally, so no temporary source,
            // bookmark or paragraph survives the measurement.
            separatorInsertion = targetDocument.Range(
                scratchBoundary,
                scratchBoundary);
            separatorInsertion.InsertBefore("\r");
            formulaInsertion = targetDocument.Range(
                scratchBoundary + 1,
                scratchBoundary + 1);
            measurementStage = "insert-target-scratch";
            measuredRange = ReplaceWithPreparedOmml(
                application,
                targetDocument,
                formulaInsertion,
                semanticOmml,
                display: true,
                mathFontName: normalizedMathFontName);
            measurementStage = "repaginate-target-scratch";
            try { targetDocument.Repaginate(); } catch { }
            window = targetDocument.ActiveWindow;
            try { previousVerticalScroll = window.VerticalPercentScrolled; } catch { }
            try { previousHorizontalScroll = window.HorizontalPercentScrolled; } catch { }
            object scrollStart = true;
            try { window.ScrollIntoView(measuredRange, ref scrollStart); } catch { }
            measurementStage = "get-point-target-scratch";
            window.GetPoint(
                out _,
                out _,
                out _,
                out var heightPixels,
                measuredRange);
            view = window.View;
            zoom = view.Zoom;
            var zoomPercentage = zoom.Percentage;
            var dpi = 96u;
            try
            {
                var detected = GetDpiForWindow(new IntPtr(window.Hwnd));
                if (detected > 0) dpi = detected;
            }
            catch (EntryPointNotFoundException) { }
            if (heightPixels <= 0 || zoomPercentage <= 0 || dpi == 0)
                return null;
            var heightPoints = heightPixels
                * 72f
                * 100f
                / dpi
                / zoomPercentage;
            return heightPoints > 0f
                && !float.IsNaN(heightPoints)
                && !float.IsInfinity(heightPoints)
                    ? heightPoints
                    : null;
        }
        catch (Exception error)
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"[PREPARED OMML HEIGHT FAILED] stage={measurementStage} type={error.GetType().Name} hresult=0x{error.HResult:X8} message={error.Message}");
            return null;
        }
        finally
        {
            Release(zoom);
            Release(view);
            Release(measuredRange);
            Release(formulaInsertion);
            Release(separatorInsertion);
            if (scratchBoundary >= 0)
            {
                try
                {
                    content = targetDocument.Content;
                    var cleanupEnd = Math.Max(
                        scratchBoundary,
                        content.End - 1);
                    cleanupRange = targetDocument.Range(
                        scratchBoundary,
                        cleanupEnd);
                    cleanupRange.Delete();
                }
                catch { }
            }
            Release(cleanupRange);
            Release(content);
            try { targetDocument.Activate(); } catch { }
            if (window is not null)
            {
                try
                {
                    if (previousHorizontalScroll.HasValue)
                        window.HorizontalPercentScrolled = previousHorizontalScroll.Value;
                }
                catch { }
                try
                {
                    if (previousVerticalScroll.HasValue)
                        window.VerticalPercentScrolled = previousVerticalScroll.Value;
                }
                catch { }
            }
            Release(window);
            if (screenUpdatingSuspended)
            {
                try { application.ScreenUpdating = previousScreenUpdating; } catch { }
            }
        }
    }

    internal static Range ReplaceWithPreparedOmml(
        Application application,
        Document targetDocument,
        Range targetRange,
        string omml,
        bool display,
        string? mathFontName = null)
    {
        if (application is null) throw new ArgumentNullException(nameof(application));
        if (targetDocument is null) throw new ArgumentNullException(nameof(targetDocument));
        if (targetRange is null) throw new ArgumentNullException(nameof(targetRange));
        if (string.IsNullOrWhiteSpace(omml))
            throw new InvalidDataException("The prepared OMML replacement is empty.");

        var normalizedMathFontName = NormalizeMathFontName(mathFontName);
        var tempPath = CreateTemporaryDocx(
            omml,
            includeLeadingTab: false,
            forceInline: !display,
            mathFontName: normalizedMathFontName);
        Document? sourceDocument = null;
        OMaths? sourceMaths = null;
        OMath? sourceMath = null;
        Range? sourceRange = null;
        Range? target = null;
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
                    "The prepared OMML replacement document did not contain exactly one equation.");
            sourceMath = sourceMaths[1];
            sourceRange = sourceMath.Range;

            target = targetRange.Duplicate;
            var insertionStart = target.Start;
            // Word can replace one professional OMath with another through
            // FormattedText without linearizing either equation. This preserves
            // radicals, fractions, matrices and the display-math separators while
            // avoiding the placeholder/BuildUp corruption seen in older builds.
            target.FormattedText = sourceRange.FormattedText;
            insertedMath = FindMathAtPosition(
                    targetDocument,
                    insertionStart,
                    target.End)
                ?? throw new InvalidOperationException(
                    "Word did not materialize the prepared OMML replacement.");
            var targetType = display
                ? WdOMathType.wdOMathDisplay
                : WdOMathType.wdOMathInline;
            if (insertedMath.Type != targetType)
                insertedMath.Type = targetType;
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
            Release(sourceRange);
            Release(sourceMath);
            Release(sourceMaths);
            if (sourceDocument is not null)
            {
                try { sourceDocument.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(sourceDocument);
            // Opening the temporary prepared-OMML document can make it Word's
            // ActiveDocument even when Visible=false. Once that document closes,
            // some Word builds expose no ActiveDocument until the caller manually
            // activates the real document; the next editor Apply then fails before
            // it can resolve its source. Restore the supplied target explicitly.
            try { targetDocument.Activate(); } catch { }
            try { File.Delete(tempPath); } catch { }
        }
    }

    private static Range InsertBookmarkedFileThroughScratchDocument(
        Application application,
        Document targetDocument,
        Range insertionRange,
        string filePath,
        string mathFontName)
    {
        lock (InlineScratchLock)
            return InsertBookmarkedFileThroughScratchDocumentCore(
                application,
                targetDocument,
                insertionRange,
                filePath,
                mathFontName);
    }

    private static Range InsertBookmarkedFileThroughScratchDocumentCore(
        Application application,
        Document targetDocument,
        Range insertionRange,
        string filePath,
        string mathFontName)
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
            scratchDocument = GetOrCreateInlineScratchDocument(
                application,
                mathFontName);
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

    private static Document GetOrCreateInlineScratchDocument(
        Application application,
        string mathFontName)
    {
        if (_inlineScratchDocument is not null)
        {
            Application? scratchApplication = null;
            try
            {
                scratchApplication = _inlineScratchDocument.Application;
                if (IsSameComObject(scratchApplication, application))
                {
                    ApplyDocumentMathFont(_inlineScratchDocument, mathFontName);
                    return _inlineScratchDocument;
                }
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
        ApplyDocumentMathFont(_inlineScratchDocument, mathFontName);
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
        mathMl = NormalizeMathTypeBinomialPiles(mathMl);
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
        omml = NormalizeAppliedFunctionStructures(omml, mathMl);
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

    internal static string NormalizeMathTypeBinomialPiles(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)
            || mathMl.IndexOf("data-mtef-pile", StringComparison.OrdinalIgnoreCase) < 0)
            return mathMl;

        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        foreach (var table in document.Descendants(presentationMath + "mtable").ToList())
        {
            if (!string.Equals(
                    table.Attribute("data-mtef-pile")?.Value,
                    "true",
                    StringComparison.OrdinalIgnoreCase))
                continue;

            var rows = table.Elements(presentationMath + "mtr").ToArray();
            if (rows.Length != 2) continue;
            var cells = rows
                .Select(row => row.Elements(presentationMath + "mtd").ToArray())
                .ToArray();
            if (cells.Any(row => row.Length != 1)) continue;

            // MTEF uses a two-row, one-column PILE inside parentheses for its
            // native binomial template. Leaving that PILE as <mtable> makes the
            // Office XSLT materialize a 2x1 matrix, which is visually similar but
            // semantically different from Word's native no-bar fraction. The
            // private decoder marker distinguishes this from a genuine MATRIX,
            // so explicit column matrices must not enter this path.
            XElement CopyCellAsRow(XElement cell) =>
                new(
                    presentationMath + "mrow",
                    new XElement(cell).Nodes());

            table.ReplaceWith(
                new XElement(
                    presentationMath + "mfrac",
                    new XAttribute("linethickness", "0"),
                    CopyCellAsRow(cells[0][0]),
                    CopyCellAsRow(cells[1][0])));
        }
        return document.ToString(SaveOptions.DisableFormatting);
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

        static string CanonicalToken(string value) =>
            new(value
                .Where(character =>
                    !char.IsWhiteSpace(character)
                    && character is not '\u200B' and not '\u200C' and not '\u2060' and not '\uFEFF')
                .ToArray());

        static bool ContainsLetter(string value) => value.Any(char.IsLetter);

        // m:nor is Office Math's "Normal Text" switch. It is correct for genuine
        // prose produced by MathML mtext / LaTeX \\text{...}, but it must not be
        // used merely to make a mathematical identifier upright. Office's stock
        // MML2OMML transform unfortunately emits m:nor for several named functions
        // (sin/log/exp/...) and older VisualTeX code also converted explicit
        // mathvariant=normal identifiers (d/e/i from \\mathrm) to m:nor. Both cases
        // bypass the document m:mathFont and make the token look like body text.
        // Collect only source nodes that carry mathematical, not prose, semantics
        // and normalize those target runs back to plain/upright Office Math.
        var uprightTokens = new HashSet<string>(StringComparer.Ordinal);
        var uprightWords = new HashSet<string>(StringComparer.Ordinal);
        foreach (var element in mathMlDocument.Descendants())
        {
            if (element.Name.Namespace != presentationMath
                || element.Name == presentationMath + "mtext")
                continue;

            var text = element.Value;
            if (string.IsNullOrWhiteSpace(text) || !ContainsLetter(text))
                continue;

            var variant = element.Attribute("mathvariant")?.Value ?? string.Empty;
            var explicitlyUpright =
                variant.IndexOf("normal", StringComparison.OrdinalIgnoreCase) >= 0
                || variant.IndexOf("upright", StringComparison.OrdinalIgnoreCase) >= 0;
            var namedIdentifier = element.Name == presentationMath + "mi"
                && CanonicalToken(text).Length > 1;
            var namedOperator = element.Name == presentationMath + "mo"
                && string.Equals(
                    element.Attribute("data-mjx-texclass")?.Value,
                    "OP",
                    StringComparison.OrdinalIgnoreCase);
            var appliedFunction = element.Name == presentationMath + "mi"
                && element.ElementsAfterSelf()
                    .FirstOrDefault(candidate => candidate.Name.Namespace == presentationMath)
                    is XElement next
                && next.Name == presentationMath + "mo"
                && string.Equals(next.Value, "\u2061", StringComparison.Ordinal);

            if (!explicitlyUpright && !namedIdentifier && !namedOperator && !appliedFunction)
                continue;

            var canonical = CanonicalToken(text);
            if (canonical.Length > 0) uprightTokens.Add(canonical);
            foreach (Match match in Regex.Matches(text, @"\p{L}+"))
                uprightWords.Add(match.Value);
        }
        if (uprightTokens.Count == 0) return omml;

        var ommlDocument = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;

        void MakeMathUpright(XElement run)
        {
            var properties = run.Element(math + "rPr");
            if (properties is null)
            {
                properties = new XElement(math + "rPr");
                run.AddFirst(properties);
            }
            properties.Element(math + "nor")?.Remove();
            var plainStyle = properties.Element(math + "sty");
            if (plainStyle is null)
            {
                plainStyle = new XElement(math + "sty");
                properties.AddFirst(plainStyle);
            }
            plainStyle.SetAttributeValue(math + "val", "p");

            // These are still mathematical tokens, but spell checking them as
            // prose is distracting. w:noProof does not change glyph selection, so
            // keep it while the actual glyphs continue to come from m:mathFont.
            var wordProperties = run.Element(word + "rPr");
            if (wordProperties is null)
            {
                wordProperties = new XElement(word + "rPr");
                properties.AddAfterSelf(wordProperties);
            }
            if (wordProperties.Element(word + "noProof") is null)
                wordProperties.Add(new XElement(word + "noProof"));
        }

        foreach (var run in ommlDocument.Descendants(math + "r").ToList())
        {
            var text = string.Concat(run.Elements(math + "t").Select(element => element.Value));
            if (string.IsNullOrEmpty(text)) continue;
            if (uprightTokens.Contains(CanonicalToken(text)))
            {
                MakeMathUpright(run);
                continue;
            }

            // Office sometimes coalesces limit-style operators into a surrounding
            // ordinary run (for example "...(x)+lim(x)+max(x)..."). In that case
            // changing the whole run to plain style would incorrectly upright x and
            // other variables. Split only the named operator words and preserve all
            // original characters/properties on the surrounding segments.
            var matches = Regex.Matches(text, @"\p{L}+")
                .Cast<Match>()
                .Where(match => uprightWords.Contains(match.Value))
                .ToArray();
            if (matches.Length == 0) continue;

            var replacements = new List<XElement>();
            var cursor = 0;
            void AddSegment(string segment, bool upright)
            {
                if (segment.Length == 0) return;
                var replacement = new XElement(run);
                replacement.Elements(math + "t").Remove();
                var textElement = new XElement(math + "t", segment);
                if (char.IsWhiteSpace(segment[0])
                    || char.IsWhiteSpace(segment[segment.Length - 1]))
                    textElement.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                replacement.Add(textElement);
                if (upright) MakeMathUpright(replacement);
                replacements.Add(replacement);
            }

            foreach (var match in matches)
            {
                AddSegment(text.Substring(cursor, match.Index - cursor), upright: false);
                AddSegment(match.Value, upright: true);
                cursor = match.Index + match.Length;
            }
            AddSegment(text.Substring(cursor), upright: false);
            run.ReplaceWith(replacements);
        }

        return ommlDocument.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    internal static string NormalizeAppliedFunctionStructures(string omml, string mathMl)
    {
        if (string.IsNullOrWhiteSpace(omml)
            || string.IsNullOrWhiteSpace(mathMl)
            || (mathMl.IndexOf('\u2061') < 0
                && mathMl.IndexOf("2061", StringComparison.OrdinalIgnoreCase) < 0))
            return omml;

        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        XNamespace math = MathNamespace;
        var source = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);

        var applications = new List<(string FunctionToken, string? Open, string? Close, int ArgumentTextElements)>();
        foreach (var marker in source
                     .Descendants(presentationMath + "mo")
                     .Where(element => string.Equals(element.Value, "\u2061", StringComparison.Ordinal)))
        {
            var parent = marker.Parent;
            if (parent is null) continue;
            var siblings = parent.Elements().ToArray();
            var markerIndex = Array.IndexOf(siblings, marker);
            if (markerIndex <= 0 || markerIndex + 1 >= siblings.Length) continue;

            var functionSource = siblings[markerIndex - 1];
            var functionToken = functionSource
                .DescendantsAndSelf()
                .Where(element =>
                    element.Name == presentationMath + "mi"
                    || element.Name == presentationMath + "mo")
                .Select(element => Regex.Match(element.Value, @"\p{L}+"))
                .FirstOrDefault(match => match.Success)?
                .Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(functionToken)) continue;

            var argumentSource = siblings[markerIndex + 1];
            string? open = null;
            string? close = null;
            if (argumentSource.Name == presentationMath + "mo")
            {
                (open, close) = argumentSource.Value switch
                {
                    "(" => ("(", ")"),
                    "[" => ("[", "]"),
                    "{" => ("{", "}"),
                    _ => (null, null),
                };
            }

            var simpleArgument = argumentSource.Name == presentationMath + "mi"
                || argumentSource.Name == presentationMath + "mn"
                || argumentSource.Name == presentationMath + "mo"
                || argumentSource.Name == presentationMath + "mtext";
            var argumentTextElements = simpleArgument && open is null
                ? StringInfo.ParseCombiningCharacters(argumentSource.Value).Length
                : 0;
            applications.Add((functionToken, open, close, argumentTextElements));
        }
        if (applications.Count == 0) return omml;

        var target = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);

        static string RunText(XElement run, XNamespace mathNamespace) =>
            string.Concat(run.Elements(mathNamespace + "t").Select(text => text.Value));

        static XElement CloneRunWithText(
            XElement sourceRun,
            XNamespace mathNamespace,
            string text)
        {
            var clone = new XElement(sourceRun);
            clone.Elements(mathNamespace + "t").Remove();
            var textElement = new XElement(mathNamespace + "t", text);
            if (text.Length > 0
                && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])))
                textElement.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            clone.Add(textElement);
            return clone;
        }

        static int TextElementEnd(string value, int count)
        {
            if (string.IsNullOrEmpty(value) || count <= 0) return 0;
            var starts = StringInfo.ParseCombiningCharacters(value);
            if (starts.Length <= count) return value.Length;
            return starts[count];
        }

        static int FindMatchingDelimiterEnd(
            string value,
            string open,
            string close,
            ref int depth,
            ref bool started)
        {
            for (var index = 0; index < value.Length; index++)
            {
                var character = value[index].ToString();
                if (string.Equals(character, open, StringComparison.Ordinal))
                {
                    depth++;
                    started = true;
                }
                else if (started && string.Equals(character, close, StringComparison.Ordinal))
                {
                    depth--;
                    if (depth == 0) return index + 1;
                }
            }
            return -1;
        }

        foreach (var application in applications)
        {
            var markerText = target
                .Descendants(math + "t")
                .FirstOrDefault(text => text.Value.IndexOf('\u2061') >= 0);
            var markerRun = markerText?.Parent;
            if (markerRun?.Name != math + "r") continue;
            var parent = markerRun.Parent;
            if (parent is null) continue;
            var functionElement = markerRun.ElementsBeforeSelf().LastOrDefault();
            if (functionElement is null) continue;
            var functionVisibleText = string.Concat(
                functionElement.Descendants(math + "t").Select(text => text.Value));
            if (functionVisibleText.IndexOf(
                    application.FunctionToken,
                    StringComparison.OrdinalIgnoreCase) < 0)
                continue;

            var markerRunText = RunText(markerRun, math);
            var markerIndex = markerRunText.IndexOf('\u2061');
            if (markerIndex < 0) continue;
            // NormalizeExplicitUprightRuns deliberately split the function name
            // into its own m:r, so ApplyFunction must begin the following run.
            // Anything else is an unfamiliar Office-XSL shape and is left intact.
            if (markerIndex != 0) continue;
            var afterMarker = markerRunText.Substring(1);
            var argumentElements = new List<XElement>();
            var consumed = new List<XElement> { markerRun };
            XElement? tail = null;
            var completed = false;

            if (application.Open is not null && application.Close is not null)
            {
                var depth = 0;
                var started = false;
                var end = FindMatchingDelimiterEnd(
                    afterMarker,
                    application.Open,
                    application.Close,
                    ref depth,
                    ref started);
                if (end >= 0)
                {
                    argumentElements.Add(CloneRunWithText(
                        markerRun,
                        math,
                        afterMarker.Substring(0, end)));
                    if (end < afterMarker.Length)
                        tail = CloneRunWithText(markerRun, math, afterMarker.Substring(end));
                    completed = true;
                }
                else if (started)
                {
                    if (afterMarker.Length > 0)
                        argumentElements.Add(CloneRunWithText(markerRun, math, afterMarker));
                    foreach (var sibling in markerRun.ElementsAfterSelf().ToArray())
                    {
                        if (sibling.Name == math + "r")
                        {
                            var siblingText = RunText(sibling, math);
                            end = FindMatchingDelimiterEnd(
                                siblingText,
                                application.Open,
                                application.Close,
                                ref depth,
                                ref started);
                            consumed.Add(sibling);
                            if (end >= 0)
                            {
                                if (end > 0)
                                    argumentElements.Add(CloneRunWithText(
                                        sibling,
                                        math,
                                        siblingText.Substring(0, end)));
                                if (end < siblingText.Length)
                                    tail = CloneRunWithText(sibling, math, siblingText.Substring(end));
                                completed = true;
                                break;
                            }
                            argumentElements.Add(new XElement(sibling));
                            continue;
                        }

                        var visible = string.Concat(
                            sibling.Descendants(math + "t").Select(text => text.Value));
                        _ = FindMatchingDelimiterEnd(
                            visible,
                            application.Open,
                            application.Close,
                            ref depth,
                            ref started);
                        consumed.Add(sibling);
                        argumentElements.Add(new XElement(sibling));
                        if (started && depth == 0)
                        {
                            completed = true;
                            break;
                        }
                    }
                }
            }
            else
            {
                var sourceTextElements = Math.Max(0, application.ArgumentTextElements);
                if (afterMarker.Length > 0)
                {
                    var take = sourceTextElements > 0
                        ? TextElementEnd(afterMarker, sourceTextElements)
                        : TextElementEnd(afterMarker, 1);
                    if (take > 0)
                    {
                        argumentElements.Add(CloneRunWithText(
                            markerRun,
                            math,
                            afterMarker.Substring(0, take)));
                        if (take < afterMarker.Length)
                            tail = CloneRunWithText(markerRun, math, afterMarker.Substring(take));
                        completed = true;
                    }
                }
                else
                {
                    var sibling = markerRun.ElementsAfterSelf().FirstOrDefault();
                    if (sibling is not null)
                    {
                        consumed.Add(sibling);
                        if (sibling.Name == math + "r" && sourceTextElements > 0)
                        {
                            var siblingText = RunText(sibling, math);
                            var take = TextElementEnd(siblingText, sourceTextElements);
                            if (take > 0)
                            {
                                argumentElements.Add(CloneRunWithText(
                                    sibling,
                                    math,
                                    siblingText.Substring(0, take)));
                                if (take < siblingText.Length)
                                    tail = CloneRunWithText(sibling, math, siblingText.Substring(take));
                                completed = true;
                            }
                        }
                        else
                        {
                            argumentElements.Add(new XElement(sibling));
                            completed = true;
                        }
                    }
                }
            }

            if (!completed || argumentElements.Count == 0) continue;

            var function = new XElement(
                math + "func",
                new XElement(math + "fName", new XElement(functionElement)),
                new XElement(math + "e", argumentElements));
            functionElement.ReplaceWith(function);
            foreach (var element in consumed)
                element.Remove();
            if (tail is not null) function.AddAfterSelf(tail);
        }

        return target.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
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
            var hidden = slot.Name == math + "e"
                && parent?.Name == math + "mr"
                && parent.Parent?.Name == math + "m"
                && IsMathBooleanEnabled(
                    parent.Parent.Element(math + "mPr"),
                    math + "plcHide")
                || slot.Name == math + "deg"
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
        var omml = StripVisualTeXNativeEquationNumber(wordOpenXml);
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
        RestoreMergedNumericPunctuationTokens(canonicalRoot);
        RestoreMathSymbolTextTokens(canonicalRoot);
        RestoreOmmlNoBarFractionSemantics(omml, canonicalRoot);
        RestoreOmmlAccentSemantics(omml, canonicalRoot);
        RestoreOmmlLimitBaseSemantics(omml, canonicalRoot);
        RestoreOmmlScriptBaseRunSemantics(omml, canonicalRoot);
        RestoreOmmlFunctionApplicationRuns(omml, canonicalRoot);
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

    private static void RestoreOmmlNoBarFractionSemantics(string omml, XElement mathMlRoot)
    {
        XNamespace officeMath = MathNamespace;
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        var source = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var sourceFractions = source.Descendants(officeMath + "f").ToArray();
        var targetFractions = mathMlRoot.DescendantsAndSelf(mathMl + "mfrac").ToArray();
        if (sourceFractions.Length == 0 || sourceFractions.Length != targetFractions.Length)
            return;

        for (var index = 0; index < sourceFractions.Length; index++)
        {
            var type = sourceFractions[index]
                .Element(officeMath + "fPr")?
                .Element(officeMath + "type")?
                .Attribute(officeMath + "val")?
                .Value;
            if (!string.Equals(type, "noBar", StringComparison.OrdinalIgnoreCase)) continue;
            targetFractions[index].SetAttributeValue("linethickness", "0");
        }
    }

    private static void RestoreMergedNumericPunctuationTokens(XElement mathMlRoot)
    {
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        foreach (var number in mathMlRoot.DescendantsAndSelf(mathMl + "mn").ToArray())
        {
            var value = number.Value;
            if (string.IsNullOrWhiteSpace(value)
                || (value.IndexOf(',') < 0 && value.IndexOf(';') < 0))
                continue;
            if (!value.All(character => char.IsDigit(character)
                    || character == ','
                    || character == ';'))
                continue;

            var tokens = new List<XElement>();
            var start = 0;
            for (var index = 0; index <= value.Length; index++)
            {
                var atEnd = index == value.Length;
                var punctuation = !atEnd && (value[index] == ',' || value[index] == ';');
                if (!atEnd && !punctuation) continue;
                if (index > start)
                    tokens.Add(new XElement(mathMl + "mn", value.Substring(start, index - start)));
                if (punctuation)
                    tokens.Add(new XElement(mathMl + "mo", value[index].ToString()));
                start = index + 1;
            }
            if (tokens.Count <= 1) continue;

            if (number.Parent?.Name == mathMl + "mrow")
                number.ReplaceWith(tokens.Cast<object>().ToArray());
            else
                number.ReplaceWith(new XElement(mathMl + "mrow", tokens));
        }
    }

    private static void RestoreMathSymbolTextTokens(XElement mathMlRoot)
    {
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        foreach (var text in mathMlRoot.DescendantsAndSelf(mathMl + "mtext").ToArray())
        {
            var value = text.Value;
            if (string.IsNullOrEmpty(value) || !ContainsOnlyMathSymbols(value)) continue;
            text.Name = mathMl + "mo";
        }
    }

    private static bool ContainsOnlyMathSymbols(string value)
    {
        var sawSymbol = false;
        for (var index = 0; index < value.Length;)
        {
            var category = CharUnicodeInfo.GetUnicodeCategory(value, index);
            var scalarLength = char.IsHighSurrogate(value[index])
                && index + 1 < value.Length
                && char.IsLowSurrogate(value[index + 1])
                ? 2
                : 1;
            if (category != UnicodeCategory.MathSymbol) return false;
            sawSymbol = true;
            index += scalarLength;
        }
        return sawSymbol;
    }

    private static void RestoreOmmlLimitBaseSemantics(string omml, XElement mathMlRoot)
    {
        XNamespace officeMath = MathNamespace;
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        var source = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var limits = source
            .Descendants()
            .Where(element => element.Name == officeMath + "limLow"
                || element.Name == officeMath + "limUpp")
            .ToArray();
        if (limits.Length == 0) return;

        var usedTargets = new HashSet<XElement>();
        foreach (var limit in limits)
        {
            var baseSlot = limit.Element(officeMath + "e");
            if (baseSlot is null) continue;
            var visibleRuns = baseSlot
                .Descendants(officeMath + "r")
                .Where(run => run.Elements(officeMath + "t")
                    .Any(text => !string.IsNullOrWhiteSpace(text.Value)))
                .ToArray();
            if (visibleRuns.Length == 0) continue;

            var baseText = string.Concat(
                visibleRuns.SelectMany(run => run.Elements(officeMath + "t"))
                    .Select(text => text.Value));
            if (string.IsNullOrWhiteSpace(baseText)) continue;

            // Word's BuildUp/save materialization can discard m:nor from the base
            // run of m:limLow/m:limUpp even though the visible operator is still a
            // native limit operator. Preserve explicit m:nor whenever it survives;
            // when it does not, only infer upright grouping for TeX's standard
            // named operators. This restores lim/max/min/sup/inf/limsup/... without
            // ever joining an arbitrary italic variable sequence such as l*i*m.
            var explicitlyNormal = visibleRuns.All(run =>
                run.Element(officeMath + "rPr")?
                    .Element(officeMath + "nor") is not null);
            if (!explicitlyNormal && !IsStandardTexOperatorName(baseText))
                continue;

            var lower = limit.Name == officeMath + "limLow";
            var targetNames = lower
                ? new HashSet<XName> { mathMl + "msub", mathMl + "munder" }
                : new HashSet<XName> { mathMl + "msup", mathMl + "mover" };
            var target = mathMlRoot
                .DescendantsAndSelf()
                .Where(element => targetNames.Contains(element.Name)
                    && !usedTargets.Contains(element))
                .FirstOrDefault(element =>
                {
                    var children = element.Elements().ToArray();
                    return children.Length >= 2
                        && string.Equals(
                            FlattenMathMlTokenText(children[0]),
                            baseText,
                            StringComparison.Ordinal);
                });
            if (target is null) continue;

            var targetChildren = target.Elements().ToArray();
            targetChildren[0].ReplaceWith(
                new XElement(
                    mathMl + "mi",
                    new XAttribute("mathvariant", "normal"),
                    baseText));
            usedTargets.Add(target);
        }
    }

    private static void RestoreOmmlScriptBaseRunSemantics(string omml, XElement mathMlRoot)
    {
        XNamespace officeMath = MathNamespace;
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        var source = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var scripts = source
            .Descendants()
            .Where(element => element.Name == officeMath + "sSub"
                || element.Name == officeMath + "sSup"
                || element.Name == officeMath + "sSubSup")
            .ToArray();
        if (scripts.Length == 0) return;

        var usedTargets = new HashSet<XElement>();
        foreach (var script in scripts)
        {
            var baseSlot = script.Element(officeMath + "e");
            if (baseSlot is null) continue;
            var visibleRuns = baseSlot
                .Descendants(officeMath + "r")
                .Where(run => run.Elements(officeMath + "t")
                    .Any(text => !string.IsNullOrWhiteSpace(text.Value)))
                .ToArray();
            if (visibleRuns.Length != 1) continue;
            var run = visibleRuns[0];
            var properties = run.Element(officeMath + "rPr");
            var style = properties?
                .Element(officeMath + "sty")?
                .Attribute(officeMath + "val")?
                .Value;
            var explicitNormal = properties?.Element(officeMath + "nor") is not null
                || string.Equals(style, "p", StringComparison.OrdinalIgnoreCase);
            if (!explicitNormal) continue;

            var baseText = string.Concat(
                run.Elements(officeMath + "t").Select(text => text.Value));
            if (string.IsNullOrWhiteSpace(baseText)) continue;

            XName targetName;
            if (script.Name == officeMath + "sSub") targetName = mathMl + "msub";
            else if (script.Name == officeMath + "sSup") targetName = mathMl + "msup";
            else targetName = mathMl + "msubsup";

            var target = mathMlRoot
                .DescendantsAndSelf(targetName)
                .Where(element => !usedTargets.Contains(element))
                .FirstOrDefault(element =>
                {
                    var children = element.Elements().ToArray();
                    return children.Length >= 2
                        && string.Equals(
                            FlattenMathMlTokenText(children[0]),
                            baseText,
                            StringComparison.Ordinal);
                });
            if (target is null) continue;

            var targetChildren = target.Elements().ToArray();
            targetChildren[0].ReplaceWith(
                new XElement(
                    mathMl + "mi",
                    new XAttribute("mathvariant", "normal"),
                    baseText));
            usedTargets.Add(target);
        }
    }

    private static void RestoreOmmlFunctionApplicationRuns(string omml, XElement mathMlRoot)
    {
        XNamespace officeMath = MathNamespace;
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        var source = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var sourceRuns = source.Descendants(officeMath + "r").ToArray();
        for (var index = 0; index + 1 < sourceRuns.Length; index++)
        {
            var run = sourceRuns[index];
            var properties = run.Element(officeMath + "rPr");
            var style = properties?
                .Element(officeMath + "sty")?
                .Attribute(officeMath + "val")?
                .Value;
            var explicitNormal = properties?.Element(officeMath + "nor") is not null
                || string.Equals(style, "p", StringComparison.OrdinalIgnoreCase);
            if (!explicitNormal) continue;

            var functionName = string.Concat(
                run.Elements(officeMath + "t").Select(text => text.Value));
            if (string.IsNullOrWhiteSpace(functionName)) continue;

            var nextRun = sourceRuns[index + 1];
            var nextText = string.Concat(
                nextRun.Elements(officeMath + "t").Select(text => text.Value));
            if (nextText.Length == 0 || nextText[0] != '\u2061') continue;

            CollapseMathMlUprightSequence(mathMlRoot, mathMl, functionName);
        }
    }

    private static bool CollapseMathMlUprightSequence(
        XElement root,
        XNamespace mathMl,
        string expectedText)
    {
        foreach (var parent in root.DescendantsAndSelf().ToArray())
        {
            var children = parent.Elements().ToArray();
            for (var start = 0; start < children.Length; start++)
            {
                var matched = new List<XElement>();
                var builder = new StringBuilder();
                for (var cursor = start; cursor < children.Length; cursor++)
                {
                    var child = children[cursor];
                    if (child.Name != mathMl + "mi") break;
                    var variant = child.Attribute("mathvariant")?.Value ?? string.Empty;
                    if (variant.IndexOf("normal", StringComparison.OrdinalIgnoreCase) < 0
                        && variant.IndexOf("upright", StringComparison.OrdinalIgnoreCase) < 0)
                        break;
                    matched.Add(child);
                    builder.Append(child.Value);
                    var candidate = builder.ToString();
                    if (!expectedText.StartsWith(candidate, StringComparison.Ordinal)) break;
                    if (!string.Equals(candidate, expectedText, StringComparison.Ordinal)) continue;

                    var replacement = new XElement(
                        mathMl + "mi",
                        new XAttribute("mathvariant", "normal"),
                        expectedText);
                    matched[0].AddBeforeSelf(replacement);
                    foreach (var element in matched) element.Remove();
                    return true;
                }
            }
        }
        return false;
    }

    private static bool IsStandardTexOperatorName(string value)
    {
        // Plain TeX/LaTeX operator names whose conventional math presentation is
        // upright. Keep this intentionally finite; arbitrary alphabetic runs must
        // remain variables even when Word happens to place them in a limit object.
        return value is "arccos" or "arcsin" or "arctan" or "arg"
            or "cos" or "cosh" or "cot" or "coth" or "csc"
            or "deg" or "det" or "dim" or "exp" or "gcd" or "hom"
            or "inf" or "ker" or "lg" or "lim" or "liminf" or "limsup"
            or "ln" or "log" or "max" or "min" or "Pr" or "sec"
            or "sin" or "sinh" or "sup" or "tan" or "tanh";
    }

    private static string FlattenMathMlTokenText(XElement element)
    {
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        if (element.Name.Namespace != mathMl) return string.Empty;
        if (element.Name.LocalName is "mi" or "mn" or "mo" or "mtext")
            return element.Value;
        if (element.Name.LocalName is not ("mrow" or "mstyle" or "mpadded" or "semantics"))
            return string.Empty;
        var builder = new StringBuilder();
        foreach (var child in element.Elements())
        {
            var value = FlattenMathMlTokenText(child);
            if (value.Length == 0 && child.HasElements) return string.Empty;
            builder.Append(value);
        }
        return builder.ToString();
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
        var normalized = StripVisualTeXNativeEquationNumber(wordOpenXml);
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

    internal static string BuildImmutableHashSequenceNumberedOmml(
        string semanticOmml,
        string sequenceName,
        string numberBookmarkName,
        string visibleBookmarkName,
        string captionBookmarkName,
        string prefix,
        int restartHeadingLevel,
        string initialSequenceResult)
    {
        if (string.IsNullOrWhiteSpace(semanticOmml))
            throw new ArgumentException(
                "The semantic OMML payload must not be empty.",
                nameof(semanticOmml));
        if (!Regex.IsMatch(
                sequenceName ?? string.Empty,
                @"^[A-Za-z][A-Za-z0-9_]{0,39}$",
                RegexOptions.CultureInvariant))
            throw new ArgumentException(
                "The Word SEQ identifier is invalid.",
                nameof(sequenceName));
        if (restartHeadingLevel < 0 || restartHeadingLevel > 9)
            throw new ArgumentOutOfRangeException(
                nameof(restartHeadingLevel),
                "A Word heading reset level must be between 0 and 9.");
        if (string.IsNullOrWhiteSpace(initialSequenceResult)
            || initialSequenceResult.Any(character =>
                character is '\r' or '\n' or '\u0013' or '\u0014' or '\u0015'))
            throw new ArgumentException(
                "The initial SEQ result is invalid.",
                nameof(initialSequenceResult));
        prefix ??= string.Empty;
        if (prefix.Any(character =>
                character is '\r' or '\n' or '\u0013' or '\u0014' or '\u0015'))
            throw new ArgumentException(
                "The equation-number prefix contains an invalid Word field-control character.",
                nameof(prefix));

        var normalizedFormulaId = ValidateManagedEquationBookmarkName(
            numberBookmarkName,
            VisualTeXNativeNumberBookmarkPrefix,
            nameof(numberBookmarkName));
        if (!string.Equals(
                normalizedFormulaId,
                ValidateManagedEquationBookmarkName(
                    visibleBookmarkName,
                    "VTEq_",
                    nameof(visibleBookmarkName)),
                StringComparison.OrdinalIgnoreCase)
            || !string.Equals(
                normalizedFormulaId,
                ValidateManagedEquationBookmarkName(
                    captionBookmarkName,
                    "VTEqCap_",
                    nameof(captionBookmarkName)),
                StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(
                "All VisualTeX equation-number bookmarks must carry the same FormulaId.");

        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;
        var equation = XElement.Parse(
            ExtractSingleOMath(semanticOmml),
            LoadOptions.PreserveWhitespace);
        var formulaNodes = equation
            .Elements()
            .Select(element => new XElement(element))
            .Cast<object>()
            .ToList();
        if (formulaNodes.Count == 0)
            throw new InvalidDataException(
                "A numbered OMML display equation must contain a nonempty mathematical body.");

        var representativeRunProperties = equation
            .Descendants(word + "rPr")
            .FirstOrDefault(properties => properties.Element(word + "sz") is not null)
            ?? equation.Descendants(word + "rPr").FirstOrDefault();
        XElement WrapperRunProperties(bool noProof = false, bool italic = false)
        {
            var properties = new XElement(word + "rPr");
            foreach (var name in new[] { word + "rFonts", word + "sz", word + "szCs" })
            {
                var property = representativeRunProperties?.Element(name);
                if (property is not null)
                    properties.Add(new XElement(property));
            }
            if (italic)
                properties.Add(new XElement(word + "i"));
            if (noProof)
                properties.Add(new XElement(word + "noProof"));
            return properties;
        }
        XElement WrapperControlProperties() =>
            new(
                math + "ctrlPr",
                // Match Word's own professional #({SEQ ...}) serialization. The
                // control run is italic, while the visible numeric result is not.
                // Marking field-control runs with m:nor makes Word treat the hidden
                // instruction as ordinary mathematical content and expands the
                // equation-number slot leftward instead of keeping it at the right
                // margin.
                WrapperRunProperties(italic: true));
        XElement FieldBoundaryRun(XElement content) =>
            new(
                math + "r",
                WrapperRunProperties(italic: true),
                content);
        XElement FieldInstructionRun(XElement content) =>
            new(
                math + "r",
                WrapperRunProperties(),
                content);
        XElement FieldResultRun(XElement content) =>
            new(
                math + "r",
                WrapperRunProperties(noProof: true),
                content);
        XElement TextRun(string text) =>
            new(
                math + "r",
                WrapperRunProperties(),
                new XElement(math + "t", text));
        XElement HashSeparatorRun() =>
            // Keep the hash itself Word-canonical. Word is allowed to merge this run
            // into the preceding simple formula run during normalization; the native
            // number parser therefore recognizes both a standalone '#' token and a
            // formula-tail run whose final character is '#'. Adding m:nor here makes
            // Word treat the separator as ordinary math and destroys right-margin
            // #() label geometry.
            TextRun("#");

        var sequenceSwitch = restartHeadingLevel > 0
            ? $" \\s {restartHeadingLevel}"
            : string.Empty;
        var sequenceInstruction =
            // Word itself adds MERGEFORMAT when a SEQ field is inserted with
            // PreserveFormatting=true before professional BuildUp. Keep the same
            // canonical field instruction so the #() number remains a right-margin
            // label after save/reopen instead of behaving like ordinary math text.
            $" SEQ {sequenceName}{sequenceSwitch} \\* ARABIC \\* MERGEFORMAT ";
        var fieldElements = new List<object>
        {
            // Keep the number alias outermost: VTEqNum must cover the complete
            // chapter prefix plus the live SEQ result read by ordinary body REF.
            new XElement(
                word + "bookmarkStart",
                new XAttribute(word + "id", "31801"),
                new XAttribute(word + "name", numberBookmarkName)),
            new XElement(
                word + "bookmarkStart",
                new XAttribute(word + "id", "31802"),
                new XAttribute(word + "name", visibleBookmarkName)),
            new XElement(
                word + "bookmarkStart",
                new XAttribute(word + "id", "31803"),
                new XAttribute(word + "name", captionBookmarkName)),
        };
        if (!string.IsNullOrEmpty(prefix))
            fieldElements.Add(TextRun(prefix));
        fieldElements.Add(FieldBoundaryRun(new XElement(
            word + "fldChar",
            // Word's own Ctrl+F9/BuildUp #(SEQ) field is not imported as dirty.
            // Setting w:dirty=true on an XML-inserted field makes interactive Word
            // show the "fields may refer to other files" update-confirmation dialog
            // even in a brand-new document. VisualTeX explicitly updates this SEQ
            // after insertion, so the dirty flag is both unnecessary and harmful.
            new XAttribute(word + "fldCharType", "begin"))));
        fieldElements.Add(FieldInstructionRun(new XElement(
            word + "instrText",
            new XAttribute(XNamespace.Xml + "space", "preserve"),
            sequenceInstruction)));
        fieldElements.Add(FieldBoundaryRun(new XElement(
            word + "fldChar",
            new XAttribute(word + "fldCharType", "separate"))));
        fieldElements.Add(FieldResultRun(new XElement(math + "t", initialSequenceResult)));
        fieldElements.Add(FieldBoundaryRun(new XElement(
            word + "fldChar",
            new XAttribute(word + "fldCharType", "end"))));
        fieldElements.Add(new XElement(
            word + "bookmarkEnd",
            new XAttribute(word + "id", "31803")));
        fieldElements.Add(new XElement(
            word + "bookmarkEnd",
            new XAttribute(word + "id", "31802")));
        fieldElements.Add(new XElement(
            word + "bookmarkEnd",
            new XAttribute(word + "id", "31801")));

        var delimiter = new XElement(
            math + "d",
            new XElement(
                math + "dPr",
                WrapperControlProperties()),
            new XElement(math + "e", fieldElements));
        var equationBody = new XElement(math + "e", formulaNodes);
        equationBody.Add(
            HashSeparatorRun(),
            delimiter);
        return new XElement(
                math + "oMath",
                new XElement(
                    math + "eqArr",
                    new XElement(
                        math + "eqArrPr",
                        new XElement(
                            math + "maxDist",
                            new XAttribute(math + "val", "1")),
                        WrapperControlProperties()),
                    equationBody))
            .ToString(SaveOptions.DisableFormatting);
    }

    private static string ValidateManagedEquationBookmarkName(
        string? bookmarkName,
        string expectedPrefix,
        string parameterName)
    {
        var value = bookmarkName ?? string.Empty;
        if (!value.StartsWith(expectedPrefix, StringComparison.Ordinal)
            || value.Length != expectedPrefix.Length + 32)
            throw new ArgumentException(
                $"The VisualTeX bookmark must use the {expectedPrefix}<32-hex-FormulaId> form.",
                parameterName);
        var identifier = value.Substring(expectedPrefix.Length);
        if (!Guid.TryParseExact(identifier, "N", out var parsed))
            throw new ArgumentException(
                $"The VisualTeX bookmark must use the {expectedPrefix}<32-hex-FormulaId> form.",
                parameterName);
        return parsed.ToString("N");
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

    internal static bool HasVisualTeXNativeEquationNumber(string wordOpenXml)
    {
        var equation = XElement.Parse(
            ExtractSingleOMath(wordOpenXml),
            LoadOptions.PreserveWhitespace);
        return TryResolveVisualTeXNativeEquationNumber(
            equation,
            out _,
            out _,
            out _);
    }

    internal static bool HasVisualTeXDirectSequenceEquationNumber(
        string wordOpenXml,
        string? formulaId = null)
    {
        var equation = XElement.Parse(
            ExtractSingleOMath(wordOpenXml),
            LoadOptions.PreserveWhitespace);
        return TryResolveVisualTeXDirectSequenceEquationNumber(
            equation,
            formulaId);
    }

    internal static string StripVisualTeXNativeEquationNumber(string wordOpenXml) =>
        StripVisualTeXNativeEquationNumberCore(
            wordOpenXml,
            allowUnboundDirectSequence: false);

    internal static string StripVisualTeXNativeEquationNumberForManagedRepair(
        string wordOpenXml) =>
        StripVisualTeXNativeEquationNumberCore(
            wordOpenXml,
            allowUnboundDirectSequence: true);

    // Compatibility name used by the managed-numbering repair and its XML
    // regression tests. Both entry points deliberately require the caller to have
    // already proven VisualTeX ownership through metadata/FormulaId before an
    // otherwise unbound direct SEQ wrapper is stripped.
    internal static string StripManagedVisualTeXNativeEquationNumber(
        string wordOpenXml) =>
        StripVisualTeXNativeEquationNumberForManagedRepair(wordOpenXml);

    private static string StripVisualTeXNativeEquationNumberCore(
        string wordOpenXml,
        bool allowUnboundDirectSequence)
    {
        var equation = XElement.Parse(
            ExtractSingleOMath(wordOpenXml),
            LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        var removed = 0;
        while (TryResolveVisualTeXNativeEquationNumber(
                   equation,
                   out var body,
                   out var separatorIndex,
                   out _,
                   allowUnboundDirectSequence)
               && body is not null
               && separatorIndex >= 0)
        {
            var bodyElements = body.Elements().ToArray();
            var formulaElements = bodyElements
                .Take(separatorIndex)
                .Select(element => new XElement(element))
                .ToList();
            var separator = bodyElements[separatorIndex];
            var separatorText = string.Concat(
                separator.Elements(math + "t").Select(text => text.Value));
            if (separatorText.Length > 1
                && separatorText.EndsWith("#", StringComparison.Ordinal))
            {
                var mergedTail = new XElement(separator);
                var tailTexts = mergedTail.Elements(math + "t").ToArray();
                if (tailTexts.Length > 0)
                {
                    var lastText = tailTexts[tailTexts.Length - 1];
                    if (lastText.Value.EndsWith("#", StringComparison.Ordinal))
                        lastText.Value = lastText.Value.Substring(0, lastText.Value.Length - 1);
                    if (tailTexts.Any(text => !string.IsNullOrEmpty(text.Value)))
                        formulaElements.Add(mergedTail);
                }
            }
            if (formulaElements.Count == 0)
                throw new InvalidDataException(
                    "The generated VisualTeX equation-number wrapper contains no formula body.");
            equation = new XElement(math + "oMath", formulaElements);
            removed++;
            if (removed > 8)
                throw new InvalidDataException(
                    "The VisualTeX native equation-number wrapper is recursively malformed.");
        }
        return equation.ToString(SaveOptions.DisableFormatting);
    }

    private static bool TryResolveVisualTeXNativeEquationNumber(
        XElement equation,
        out XElement? body,
        out int separatorIndex,
        out XElement? numberDelimiter,
        bool allowUnboundDirectSequence = false)
    {
        body = null;
        separatorIndex = -1;
        numberDelimiter = null;
        XNamespace math = MathNamespace;
        var equationArray = equation.Elements(math + "eqArr").SingleOrDefault();
        if (equationArray is null) return false;
        var entries = equationArray.Elements(math + "e").ToArray();
        if (entries.Length != 1) return false;
        var candidateBody = entries[0];
        var children = candidateBody.Elements().ToArray();
        for (var index = 0; index + 1 < children.Length; index++)
        {
            var candidateSeparator = children[index];
            if (candidateSeparator.Name != math + "r") continue;
            var separatorText = string.Concat(
                candidateSeparator.Elements(math + "t").Select(text => text.Value));
            if (!separatorText.EndsWith("#", StringComparison.Ordinal)) continue;
            var candidateNumber = children[index + 1];
            if (candidateNumber.Name != math + "d") continue;
            var numberText = string.Concat(
                candidateNumber.Descendants(math + "t").Select(text => text.Value));
            var generatedPlaceholder = numberText.IndexOf(
                VisualTeXEquationNumberPlaceholderPrefix,
                StringComparison.Ordinal) >= 0;
            XNamespace word = WordNamespace;
            var fieldCode = string.Concat(
                candidateNumber
                    .Descendants(word + "instrText")
                    .Select(text => text.Value));
            var generatedReference = numberText.IndexOf(
                    "REF " + VisualTeXNativeNumberBookmarkPrefix,
                    StringComparison.OrdinalIgnoreCase) >= 0
                || fieldCode.IndexOf(
                    "REF " + VisualTeXNativeNumberBookmarkPrefix,
                    StringComparison.OrdinalIgnoreCase) >= 0;
            var generatedDirectSequence =
                IsVisualTeXDirectSequenceDelimiter(candidateNumber, formulaId: null)
                && (allowUnboundDirectSequence
                    || candidateNumber
                        .Descendants(word + "bookmarkStart")
                        .Any(element =>
                            ((string?)element.Attribute(word + "name") ?? string.Empty)
                                .StartsWith(
                                    VisualTeXNativeNumberBookmarkPrefix,
                                    StringComparison.OrdinalIgnoreCase)));
            if (!generatedPlaceholder && !generatedReference && !generatedDirectSequence)
                continue;
            body = candidateBody;
            separatorIndex = index;
            numberDelimiter = candidateNumber;
            return true;
        }
        return false;
    }

    private static bool TryResolveVisualTeXDirectSequenceEquationNumber(
        XElement equation,
        string? formulaId)
    {
        XNamespace math = MathNamespace;
        var equationArray = equation.Elements(math + "eqArr").SingleOrDefault();
        if (equationArray is null) return false;
        var entries = equationArray.Elements(math + "e").ToArray();
        if (entries.Length != 1) return false;
        var children = entries[0].Elements().ToArray();
        for (var index = 0; index + 1 < children.Length; index++)
        {
            var separator = children[index];
            if (separator.Name != math + "r") continue;
            var separatorText = string.Concat(
                separator.Elements(math + "t").Select(text => text.Value));
            if (!separatorText.EndsWith("#", StringComparison.Ordinal)) continue;
            var delimiter = children[index + 1];
            if (delimiter.Name != math + "d") continue;
            if (IsVisualTeXDirectSequenceDelimiter(delimiter, formulaId))
                return true;
        }
        return false;
    }

    private static bool IsVisualTeXDirectSequenceDelimiter(
        XElement delimiter,
        string? formulaId)
    {
        XNamespace word = WordNamespace;
        XNamespace math = MathNamespace;
        // Word accepts w:instrText when the prepared OMath is imported, but its
        // native math normalizer commonly serializes that field instruction back as
        // m:r/m:t while COM still exposes the same live Field.Code.Text. Recognize
        // both spellings; never rewrite either representation in place.
        var fieldCode = string.Concat(
            delimiter.Descendants(word + "instrText").Select(text => text.Value))
            + string.Concat(
                delimiter.Descendants(math + "t").Select(text => text.Value));
        if (!Regex.IsMatch(
                fieldCode,
                @"\bSEQ\s+(?:""VisualTeXEquation""|VisualTeXEquation)\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            return false;
        if (Regex.IsMatch(
                fieldCode,
                @"\bREF\s+VTEqNum_",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            return false;

        // OMath.Range.WordOpenXML is not a stable bookmark boundary API: after a
        // paragraph is inserted before an existing equation, Word can omit a
        // bookmarkStart that shares the OMath start while the bookmark remains
        // healthy in Document.Bookmarks and the full document XML. A null FormulaId
        // therefore means structure-only recognition. Callers that own a FormulaId
        // validate its bookmark through COM range containment separately.
        if (string.IsNullOrWhiteSpace(formulaId))
            return true;
        if (!Guid.TryParse(formulaId, out var formulaGuid)) return false;
        var expectedBookmark =
            VisualTeXNativeNumberBookmarkPrefix + formulaGuid.ToString("N");
        return delimiter
            .Descendants(word + "bookmarkStart")
            .Any(element => string.Equals(
                (string?)element.Attribute(word + "name") ?? string.Empty,
                expectedBookmark,
                StringComparison.OrdinalIgnoreCase));
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
        var overrideRoot = Environment.GetEnvironmentVariable(
            "VISUALTEX_OFFICE_MATH_XSL_ROOT");
        AddCandidateRoot(candidates, overrideRoot, fileName);

        // App Paths is the authoritative Office location for MSI, Click-to-Run,
        // per-user and alternate-bit installations. Derive the stylesheet folder
        // from every visible 32/64-bit registry view before trying conventional
        // Program Files roots.
        foreach (var wordPath in ReadRegisteredWordPaths())
        {
            var directory = Path.GetDirectoryName(wordPath);
            if (!string.IsNullOrWhiteSpace(directory))
                candidates.Add(Path.Combine(directory!, fileName));
        }

        AddCandidateRoot(candidates, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), fileName);
        AddCandidateRoot(candidates, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), fileName);
        AddCandidateRoot(candidates, Environment.GetEnvironmentVariable("ProgramW6432"), fileName);
        AddCandidateRoot(candidates, Environment.GetEnvironmentVariable("CommonProgramFiles"), fileName);
        AddCandidateRoot(candidates, Environment.GetEnvironmentVariable("CommonProgramFiles(x86)"), fileName);
        AddCandidateRoot(candidates, AppContext.BaseDirectory, fileName);
        foreach (var candidate in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (File.Exists(candidate)) return candidate;
        }
        throw new FileNotFoundException(
            $"Unable to locate Office {fileName}. Searched the registered Word installation and Office 2007-365 roots. Repair Microsoft Word or reinstall the Office integration.");
    }

    private static void AddCandidateRoot(List<string> candidates, string? root, string fileName)
    {
        if (string.IsNullOrWhiteSpace(root)) return;
        var expandedRoot = Environment.ExpandEnvironmentVariables(
            root!.Trim().Trim('"'));
        if (File.Exists(expandedRoot))
        {
            candidates.Add(expandedRoot);
            return;
        }

        candidates.Add(Path.Combine(expandedRoot, fileName));
        foreach (var version in new[] { "Office16", "Office15", "Office14", "Office12" })
        {
            candidates.Add(Path.Combine(
                expandedRoot,
                "Microsoft Office",
                "root",
                version,
                fileName));
            candidates.Add(Path.Combine(
                expandedRoot,
                "Microsoft Office",
                version,
                fileName));
            candidates.Add(Path.Combine(expandedRoot, version, fileName));
            candidates.Add(Path.Combine(
                expandedRoot,
                "Microsoft Shared",
                version.ToUpperInvariant(),
                fileName));
        }
    }

    private static IEnumerable<string> ReadRegisteredWordPaths()
    {
        var views = Environment.Is64BitOperatingSystem
            ? new[] { RegistryView.Registry64, RegistryView.Registry32 }
            : new[] { RegistryView.Registry32 };
        foreach (var view in views)
        {
            foreach (var hive in new[]
                     {
                         RegistryHive.CurrentUser,
                         RegistryHive.LocalMachine,
                     })
            {
                RegistryKey? baseKey = null;
                RegistryKey? appPathKey = null;
                try
                {
                    baseKey = RegistryKey.OpenBaseKey(hive, view);
                    appPathKey = baseKey.OpenSubKey(
                        "Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE");
                    var value = appPathKey?.GetValue(null) as string;
                    if (string.IsNullOrWhiteSpace(value)) continue;
                    var path = NormalizeRegisteredExecutablePath(value!);
                    if (!string.IsNullOrWhiteSpace(path)) yield return path;
                }
                finally
                {
                    appPathKey?.Dispose();
                    baseKey?.Dispose();
                }
            }
        }
    }

    private static string NormalizeRegisteredExecutablePath(string value)
    {
        var source = Environment.ExpandEnvironmentVariables(value.Trim());
        if (source.StartsWith("\"", StringComparison.Ordinal))
        {
            var closing = source.IndexOf('"', 1);
            return closing > 1
                ? source.Substring(1, closing - 1)
                : source.Trim('"');
        }
        var exeEnd = source.IndexOf(".exe", StringComparison.OrdinalIgnoreCase);
        return (exeEnd >= 0
                ? source.Substring(0, exeEnd + 4)
                : source.Split(new[] { ' ', '\t' }, 2)[0])
            .Trim()
            .Trim('"');
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

    private static string NormalizeMathFontName(string? mathFontName) =>
        string.IsNullOrWhiteSpace(mathFontName)
            ? "Cambria Math"
            : mathFontName.Trim();

    private static void ApplyDocumentMathFont(
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

    private static string ApplyExplicitTransferMathFont(
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

    private static string CreateTemporaryDocumentDocx(
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

    private static string CreateTemporaryDisplayGroupDocx(
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

    private static string CreateTemporaryAdjacentInlineGroupDocx(
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

    private static string CreateTemporaryBatchDocx(
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

    private static string CreateTemporaryDocx(
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
        const int rpcCallRejected = unchecked((int)0x80010001);
        const int rpcServerCallRetryLater = unchecked((int)0x8001010A);
        const int maximumAttempts = 40;
        for (var attempt = 0; ; attempt++)
        {
            try
            {
                return FindMathAtPositionCore(document, position, preferredEnd);
            }
            catch (COMException error)
                when ((error.HResult == rpcCallRejected
                        || error.HResult == rpcServerCallRetryLater)
                    && attempt < maximumAttempts - 1)
            {
                // InsertFile/FormattedText can return just before Word has committed
                // the new native equation tree. Retrying the read-only locator is
                // safer than replaying the insertion and cannot duplicate content.
                System.Threading.Thread.Sleep(50);
            }
        }
    }

    private static OMath? FindMathAtPositionCore(
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
