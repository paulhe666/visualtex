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
    private static bool _inlineScratchPrimed;
    private static XslCompiledTransform? _mathMlToOmml;
    private static XslCompiledTransform? _ommlToMathMl;

    internal sealed class BatchSource : IDisposable
    {
        private Document? _document;
        private readonly bool _fileBacked;
        private readonly string _path;
        private readonly IReadOnlyDictionary<string, BatchEntry> _entries;
        private readonly string _mathFontName;
        private readonly Dictionary<string, (Document Document, string Path)> _preparedGroups = new(StringComparer.Ordinal);
        private bool _groupInsertionStarted;

        private static string GroupKey(IReadOnlyList<string> ids, bool display) =>
            (display ? "display:" : "inline:") + string.Join(",", ids.Select(id => id.ToUpperInvariant()));

        internal void PrepareGroup(Application application, IReadOnlyList<string> formulaIds, bool display)
        {
            if (_document is null) throw new ObjectDisposedException(nameof(BatchSource));
            if (_groupInsertionStarted) throw new InvalidOperationException("OMML group sources must be prepared before insertion starts.");
            if (formulaIds is null || formulaIds.Count < (display ? 1 : 2)
                || formulaIds.Distinct(StringComparer.OrdinalIgnoreCase).Count() != formulaIds.Count)
                throw new InvalidDataException("OMML group preparation requires independent formula identities.");
            var key = GroupKey(formulaIds, display);
            if (_preparedGroups.ContainsKey(key)) return;
            var entries = formulaIds.Select(id => _entries.TryGetValue(id, out var entry) ? entry
                : throw new InvalidDataException($"OMML group source {id} is unavailable.")).ToList();
            var path = display ? CreateTemporaryDisplayGroupDocx(entries, _mathFontName)
                : CreateTemporaryAdjacentInlineGroupDocx(entries, _mathFontName);
            Document? source = null;
            try
            {
                source = application.Documents.Open(FileName: path, ConfirmConversions: false,
                    ReadOnly: true, AddToRecentFiles: false, Visible: false, OpenAndRepair: false);
                _preparedGroups.Add(key, (source, path));
                source = null;
            }
            finally
            {
                if (source is not null) { try { source.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { } Release(source); }
                if (!_preparedGroups.ContainsKey(key)) { try { File.Delete(path); } catch { } }
            }
        }

        private Document ReadPreparedGroup(IReadOnlyList<string> formulaIds, bool display)
        {
            if (_document is null) throw new ObjectDisposedException(nameof(BatchSource));
            if (!_preparedGroups.TryGetValue(GroupKey(formulaIds, display), out var group))
                throw new InvalidOperationException("The OMML group was not prepared before the target edit transaction.");
            _groupInsertionStarted = true;
            return group.Document;
        }

        internal BatchSource(
            Document? document,
            string path,
            IReadOnlyDictionary<string, BatchEntry> entries,
            string mathFontName,
            bool fileBacked = false)
        {
            if (document is null && !fileBacked)
                throw new ArgumentNullException(nameof(document));
            _document = document;
            _fileBacked = fileBacked;
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

        internal string GetSourceOmml(string formulaId)
        {
            if (!_entries.TryGetValue(formulaId, out var entry))
                throw new InvalidDataException($"The OMML batch source does not contain formula {formulaId}.");
            return entry.Omml;
        }

        internal IReadOnlyList<Range> InsertAdjacentInlineGroup(
            Application application,
            Document targetDocument,
            Range targetRange,
            IReadOnlyList<string> formulaIds,
            Action<string, Range> retainInsertedMath)
        {
            if (retainInsertedMath is null) throw new ArgumentNullException(nameof(retainInsertedMath));
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

            var sourceDocument = ReadPreparedGroup(formulaIds, display: false);
            Range? sourceRange = null;
            Range? target = null;
            OMaths? maths = null;
            OMath? math = null;
            Range? mathRange = null;
            var results = new List<Range>(formulaIds.Count);
            try
            {
                sourceRange = sourceDocument.Content.Duplicate;
                if (sourceRange.End > sourceRange.Start)
                    sourceRange.End--;
                target = targetRange.Duplicate;
                var insertionStart = target.Start;
                target.FormattedText = sourceRange.FormattedText;
                var insertionEnd = target.End;

                var insertedRange = targetDocument.Range(insertionStart, insertionEnd);
                try { maths = insertedRange.OMaths; }
                finally { Release(insertedRange); }
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
                for (var index = 0; index < candidates.Count; index++)
                {
                    var candidate = candidates[index];
                    var result = targetDocument.Range(candidate.Start, candidate.End);
                    results.Add(result);
                    retainInsertedMath(formulaIds[index], result);
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
                Release(target);
                Release(sourceRange);
                // Only BatchSource owns this prepared document; closing it here
                // would break the target document's ongoing Undo transaction.
            }
        }

        internal IReadOnlyList<Range> ReplaceDisplayParagraphGroup(
            Application application,
            Document targetDocument,
            Range targetRange,
            IReadOnlyList<string> formulaIds,
            Action<string, Range> retainInsertedMath)
        {
            if (retainInsertedMath is null) throw new ArgumentNullException(nameof(retainInsertedMath));
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

            var bodyFormatting = new List<WordCharacterFormatting>(formulaIds.Count);
            Paragraphs? targetParagraphs = null;
            try
            {
                targetParagraphs = targetRange.Paragraphs;
                if (targetParagraphs.Count != formulaIds.Count)
                    throw new InvalidDataException("The display group does not have one body paragraph per formula.");
                for (var index = 1; index <= targetParagraphs.Count; index++)
                {
                    Paragraph? paragraph = null;
                    Range? paragraphRange = null;
                    try
                    {
                        paragraph = targetParagraphs[index];
                        paragraphRange = paragraph.Range;
                        bodyFormatting.Add(WordCharacterFormatting.CaptureParagraphMark(paragraphRange));
                    }
                    finally { Release(paragraphRange); Release(paragraph); }
                }
            }
            finally { Release(targetParagraphs); }

            var sourceDocument = ReadPreparedGroup(formulaIds, display: true);
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
                    bodyFormatting[index - 1].ApplyToParagraphMark(mathRange);
                    var result = targetDocument.Range(mathRange.Start, mathRange.End);
                    results.Add(result);
                    retainInsertedMath(formulaIds[index - 1], result);
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
                // The batch owns all source documents until mutation has ended.
            }
        }

        internal Range ReplaceTargetParagraphAtomicallyWithCleanParagraph(
            Document targetDocument,
            Range targetParagraphRange)
        {
            if (_document is null && !_fileBacked)
                throw new ObjectDisposedException(nameof(BatchSource));
            Range? target = null;
            Range? editable = null;
            Range? emptyParagraph = null;
            InlineShapes? emptyShapes = null;
            OMaths? emptyMaths = null;
            Fields? emptyFields = null;
            Tables? emptyTables = null;
            Range? result = null;
            try
            {
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

                // Keep the ORIGINAL paragraph mark. Assigning FormattedText from a
                // hidden document to the complete paragraph also replaces its final
                // ¶. Word 2021 can then re-home unrelated bookmarks and the boundary
                // of a following user table onto the replacement paragraph. The
                // caller has already proved that this MathType display paragraph
                // contains no ordinary user text, so clear only its editable body
                // [Start, End-1). This removes Equation.DSMT4, MTPlaceRef and nested
                // sequence fields as one Word text replacement while preserving the
                // document's structural paragraph boundary and all later anchors.
                var bodyFormatting = WordCharacterFormatting.CaptureParagraphMark(target);
                editable = targetDocument.Range(start, target.End - 1);
                editable.Text = string.Empty;

                emptyParagraph = targetDocument.Range(start, start + 1);
                emptyShapes = emptyParagraph.InlineShapes;
                emptyMaths = emptyParagraph.OMaths;
                emptyFields = emptyParagraph.Fields;
                emptyTables = emptyParagraph.Tables;
                if (emptyParagraph.Text != "\r"
                    || emptyShapes.Count != 0
                    || emptyMaths.Count != 0
                    || emptyFields.Count != 0
                    || emptyTables.Count != 0)
                    throw new InvalidDataException(
                        "The MathType body replacement did not leave one structurally empty body paragraph.");
                bodyFormatting.ApplyToParagraphMark(emptyParagraph);

                result = targetDocument.Range(start, start);
                var returned = result;
                result = null;
                return returned;
            }
            finally
            {
                Release(result);
                Release(emptyTables);
                Release(emptyFields);
                Release(emptyMaths);
                Release(emptyShapes);
                Release(emptyParagraph);
                Release(editable);
                Release(target);
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
            var sourceDocument = _document;
            if (sourceDocument is null && !_fileBacked)
                throw new ObjectDisposedException(nameof(BatchSource));
            if (!_entries.TryGetValue(formulaId, out var entry))
                throw new InvalidDataException(
                    $"The OMML batch source does not contain formula {formulaId}.");

            Bookmarks? bookmarks = null;
            Bookmark? bookmark = null;
            Range? sourceRange = null;
            Range? formattedSource = null;
            Range? target = null;
            OMath? insertedMath = null;
            Bookmarks? targetBookmarks = null;
            Bookmark? copiedBatchBookmark = null;
            Range? copiedBatchBookmarkRange = null;
            Range? result = null;
            try
            {
                target = insertionRange.Duplicate;
                if (!replaceTarget)
                    target.Collapse(WdCollapseDirection.wdCollapseStart);
                var insertionStart = target.Start;

                if (_fileBacked)
                {
                    // The source OMath was already materialized and normalized in
                    // a separate hidden Word automation instance. Do NOT use
                    // Range.InsertFile here: Word 2019/2021 internally opens that
                    // DOCX in the target Application and briefly makes it
                    // ActiveDocument even though the API call never exposes a
                    // visible source window. Insert the already-normalized OMath
                    // fragment directly into the target range instead. This keeps
                    // the user's Application on one ActiveDocument for the entire
                    // insertion and removes the repaint/top-flash at its source.
                    var insertionXml = entry.InsertionWordOpenXml
                        ?? entry.Omml;
                    target.InsertXML(insertionXml);
                }
                else
                {
                    if (sourceDocument is null)
                        throw new ObjectDisposedException(nameof(BatchSource));
                    bookmarks = sourceDocument.Bookmarks;
                    var sourceBookmarkIndex = entry.BookmarkId + 1;
                    if (sourceBookmarkIndex < 1 || sourceBookmarkIndex > bookmarks.Count)
                        throw new InvalidDataException(
                            $"The OMML batch bookmark {entry.BookmarkName} is missing.");
                    bookmark = bookmarks[sourceBookmarkIndex];
                    if (!string.Equals(
                            bookmark.Name,
                            entry.BookmarkName,
                            StringComparison.Ordinal))
                        throw new InvalidDataException(
                            $"The OMML batch bookmark order changed at {entry.BookmarkName}.");
                    sourceRange = bookmark.Range;
                    formattedSource = sourceRange.FormattedText
                        ?? throw new InvalidDataException(
                            $"The OMML batch bookmark {entry.BookmarkName} has no formatted source range.");
                    target.FormattedText = formattedSource;
                }

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

                targetBookmarks = target.Bookmarks;
                for (var targetBookmarkIndex = 1;
                     targetBookmarkIndex <= targetBookmarks.Count;
                     targetBookmarkIndex++)
                {
                    Release(copiedBatchBookmark);
                    copiedBatchBookmark = targetBookmarks[targetBookmarkIndex];
                    if (!string.Equals(
                            copiedBatchBookmark.Name,
                            entry.BookmarkName,
                            StringComparison.Ordinal))
                        continue;
                    copiedBatchBookmarkRange = copiedBatchBookmark.Range.Duplicate;
                    var allowedStart = _fileBacked
                        ? Math.Max(0, result.Start - 1)
                        : insertionStart;
                    var allowedEnd = _fileBacked
                        ? result.End + 1
                        : target.End;
                    if (copiedBatchBookmarkRange.StoryType != result.StoryType
                        || copiedBatchBookmarkRange.Start < allowedStart
                        || copiedBatchBookmarkRange.End > allowedEnd)
                        throw new InvalidOperationException(
                            $"The copied OMML transport bookmark {entry.BookmarkName} escaped the materialized equation.");
                    copiedBatchBookmark.Delete();
                    break;
                }

                sourceFingerprint = entry.SourceFingerprint;
                var returned = result;
                result = null;
                return returned;
            }
            finally
            {
                Release(result);
                Release(copiedBatchBookmarkRange);
                Release(copiedBatchBookmark);
                Release(targetBookmarks);
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
            foreach (var group in _preparedGroups.Values)
            {
                try { group.Document.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
                Release(group.Document);
                try { File.Delete(group.Path); } catch { }
            }
            _preparedGroups.Clear();
            try { File.Delete(_path); } catch { }
        }
    }

    internal sealed class WholeDocumentSource : IDisposable
    {
        private Document? _document;
        private readonly string _path;
        private readonly IReadOnlyDictionary<string, string> _transportBookmarks;

        internal WholeDocumentSource(
            Document document,
            string path,
            IReadOnlyDictionary<string, string>? transportBookmarks = null)
        {
            _document = document;
            _path = path;
            _transportBookmarks = transportBookmarks
                ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
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

        internal IReadOnlyDictionary<string, Range> ReplaceBodyAndResolveTransportBookmarks(
            Document targetDocument)
        {
            var sourceDocument = _document
                ?? throw new ObjectDisposedException(nameof(WholeDocumentSource));
            if (_transportBookmarks.Count == 0)
                throw new InvalidOperationException(
                    "The whole-document source has no formula transport identities.");

            Range? sourceRange = null;
            Range? formattedSource = null;
            Range? targetRange = null;
            Bookmarks? bookmarks = null;
            Bookmark? bookmark = null;
            Range? bookmarkRange = null;
            OMaths? maths = null;
            OMath? math = null;
            Range? mathRange = null;
            var resolved = new Dictionary<string, Range>(
                _transportBookmarks.Count,
                StringComparer.OrdinalIgnoreCase);
            try
            {
                sourceRange = sourceDocument.Content.Duplicate;
                targetRange = targetDocument.Content.Duplicate;
                // Preserve the target document's terminal paragraph/section
                // boundary.  The cloned source has the same body topology, so one
                // FormattedText assignment replaces every Equation.DSMT4 run while
                // leaving headers, footers and section ownership on the real file.
                if (sourceRange.End > sourceRange.Start) sourceRange.End--;
                if (targetRange.End > targetRange.Start) targetRange.End--;
                formattedSource = sourceRange.FormattedText;
                targetRange.FormattedText = formattedSource;
                try { targetDocument.Activate(); } catch { }
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-whole-inline-omml-body-replaced sourceMaths={ReadOmmlCount(sourceDocument)} targetMaths={ReadOmmlCount(targetDocument)} transports={_transportBookmarks.Count}");

                bookmarks = targetDocument.Bookmarks;
                for (var index = bookmarks.Count; index >= 1; index--)
                {
                    Release(bookmark); bookmark = null;
                    Release(bookmarkRange); bookmarkRange = null;
                    Release(maths); maths = null;
                    Release(math); math = null;
                    Release(mathRange); mathRange = null;

                    bookmark = bookmarks[index];
                    if (!_transportBookmarks.TryGetValue(
                            bookmark.Name,
                            out var formulaId))
                        continue;
                    if (resolved.ContainsKey(formulaId))
                        throw new InvalidDataException(
                            $"The whole-document OMML transport identity {formulaId} is duplicated.");
                    bookmarkRange = bookmark.Range;
                    maths = bookmarkRange.OMaths;
                    if (maths.Count != 1)
                        throw new InvalidDataException(
                            $"The whole-document OMML transport identity {formulaId} owns {maths.Count} equations.");
                    math = maths[1];
                    if (math.Type != WdOMathType.wdOMathInline)
                        math.Type = WdOMathType.wdOMathInline;
                    mathRange = math.Range.Duplicate;
                    if (mathRange.StoryType != bookmarkRange.StoryType
                        || mathRange.Start < bookmarkRange.Start
                        || mathRange.End > bookmarkRange.End)
                        throw new InvalidDataException(
                            $"The whole-document OMML transport identity {formulaId} escaped its bookmark.");
                    resolved.Add(
                        formulaId,
                        targetDocument.Range(mathRange.Start, mathRange.End));
                    bookmark.Delete();
                }
                if (resolved.Count != _transportBookmarks.Count)
                    throw new InvalidDataException(
                        $"Word retained {resolved.Count}/{_transportBookmarks.Count} whole-document OMML transport identities.");
                return resolved;
            }
            catch
            {
                foreach (var range in resolved.Values) Release(range);
                throw;
            }
            finally
            {
                Release(mathRange);
                Release(math);
                Release(maths);
                Release(bookmarkRange);
                Release(bookmark);
                Release(bookmarks);
                Release(targetRange);
                Release(formattedSource);
                Release(sourceRange);
            }
        }

        private static int ReadOmmlCount(Document document)
        {
            OMaths? maths = null;
            try
            {
                maths = document.OMaths;
                return maths.Count;
            }
            finally { Release(maths); }
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

    internal static WholeDocumentSource CreateMathTypeInlineWholeDocumentSource(
        Application application,
        Document sourceDocument,
        IReadOnlyList<(string FormulaId, string Omml)> formulas,
        string? mathFontName = null)
    {
        if (application is null) throw new ArgumentNullException(nameof(application));
        if (sourceDocument is null) throw new ArgumentNullException(nameof(sourceDocument));
        if (formulas is null || formulas.Count == 0)
            throw new ArgumentOutOfRangeException(
                nameof(formulas),
                "At least one inline MathType formula is required.");

        var normalizedMathFontName = NormalizeMathFontName(mathFontName);
        var path = Path.Combine(
            Path.GetTempPath(),
            $"visualtex-omml-whole-{Guid.NewGuid():N}.docx");
        Document? scratch = null;
        Range? sourceRange = null;
        Range? scratchRange = null;
        Document? materialized = null;
        var ownershipTransferred = false;
        var transportBookmarks = new Dictionary<string, string>(
            formulas.Count,
            StringComparer.OrdinalIgnoreCase);
        try
        {
            scratch = application.Documents.Add(Visible: false);
            sourceRange = sourceDocument.Content.Duplicate;
            scratchRange = scratch.Content.Duplicate;
            // Clone the live (possibly unsaved) main story once.  This transfers
            // styles, images, fields and all non-formula content through Word's
            // native model before the package copy is edited offline. Preserve the
            // scratch document's one terminal paragraph mark instead of importing
            // a second one from the live document.
            if (sourceRange.End > sourceRange.Start) sourceRange.End--;
            if (scratchRange.End > scratchRange.Start) scratchRange.End--;
            scratchRange.FormattedText = sourceRange.FormattedText;
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-whole-inline-omml-cloned sourceMaths={ReadDocumentOmmlCount(sourceDocument)} scratchMaths={ReadDocumentOmmlCount(scratch)} formulas={formulas.Count}");
            ApplyDocumentMathFont(scratch, normalizedMathFontName);
            scratch.SaveAs2(
                FileName: path,
                FileFormat: WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            scratch.Saved = true;
            scratch.Close(WdSaveOptions.wdDoNotSaveChanges);
            Release(scratch); scratch = null;
            Release(scratchRange); scratchRange = null;
            Release(sourceRange); sourceRange = null;

            ReplaceMathTypeRunsInWholeDocumentPackage(
                path,
                formulas,
                normalizedMathFontName,
                transportBookmarks);

            materialized = application.Documents.Open(
                FileName: path,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-whole-inline-omml-materialized maths={ReadDocumentOmmlCount(materialized)} formulas={formulas.Count}");
            var result = new WholeDocumentSource(
                materialized,
                path,
                transportBookmarks);
            materialized = null;
            ownershipTransferred = true;
            return result;
        }
        finally
        {
            Release(scratchRange);
            Release(sourceRange);
            if (scratch is not null)
            {
                try { scratch.Saved = true; } catch { }
                try { scratch.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
                Release(scratch);
            }
            if (materialized is not null)
            {
                try { materialized.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
                Release(materialized);
            }
            if (!ownershipTransferred)
                try { File.Delete(path); } catch { }
        }
    }

    private static int ReadDocumentOmmlCount(Document document)
    {
        OMaths? maths = null;
        try
        {
            maths = document.OMaths;
            return maths.Count;
        }
        finally { Release(maths); }
    }

    private static void ReplaceMathTypeRunsInWholeDocumentPackage(
        string path,
        IReadOnlyList<(string FormulaId, string Omml)> formulas,
        string mathFontName,
        IDictionary<string, string> transportBookmarks)
    {
        using var stream = new FileStream(
            path,
            FileMode.Open,
            FileAccess.ReadWrite,
            FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: false);
        var entry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("The cloned Word document has no main document XML.");
        string documentXml;
        using (var entryStream = entry.Open())
        using (var reader = new StreamReader(entryStream, Encoding.UTF8, true))
            documentXml = reader.ReadToEnd();

        var package = XDocument.Parse(documentXml, LoadOptions.PreserveWhitespace);
        XNamespace word = WordNamespace;
        XNamespace math = MathNamespace;
        XNamespace office = "urn:schemas-microsoft-com:office:office";
        var mathTypeObjects = package
            .Descendants(office + "OLEObject")
            .Where(element => IsMathTypePackageProgId(
                (string?)element.Attribute("ProgID")))
            .ToArray();
        if (mathTypeObjects.Length != formulas.Count)
            throw new InvalidDataException(
                $"The cloned document contains {mathTypeObjects.Length}/{formulas.Count} MathType objects.");

        var bookmarkId = package
            .Descendants(word + "bookmarkStart")
            .Select(element => int.TryParse(
                    (string?)element.Attribute(word + "id"),
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out var parsed)
                ? parsed
                : 0)
            .DefaultIfEmpty(0)
            .Max();
        var transportBookmarkIds = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < formulas.Count; index++)
        {
            var formula = formulas[index];
            if (!Guid.TryParse(formula.FormulaId, out var parsedFormulaId))
                throw new InvalidDataException(
                    $"The whole-document OMML formula id '{formula.FormulaId}' is invalid.");
            var transportName = "VTW_" + parsedFormulaId.ToString("N");
            if (transportBookmarks.ContainsKey(transportName))
                throw new InvalidDataException(
                    $"The whole-document OMML formula id '{formula.FormulaId}' is duplicated.");

            var oleObject = mathTypeObjects[index];
            var run = oleObject.Ancestors(word + "r").FirstOrDefault()
                ?? throw new InvalidDataException(
                    "A MathType object in the cloned document has no Word run.");
            if (run.Parent?.Name != word + "p"
                || run.Elements().Count() != 1
                || run.Element(word + "object") is null
                || run.Descendants(office + "OLEObject")
                    .Count(element => IsMathTypePackageProgId(
                        (string?)element.Attribute("ProgID"))) != 1)
                throw new InvalidDataException(
                    "The whole-document fast path requires one standalone inline MathType object per Word run.");

            var equation = XElement.Parse(
                ApplyExplicitTransferMathFont(
                    ExtractSingleOMath(formula.Omml),
                    mathFontName),
                LoadOptions.PreserveWhitespace);
            if (equation.Name != math + "oMath")
                throw new InvalidDataException(
                    "The whole-document replacement payload is not one inline OMath.");
            var nextBookmarkId = checked(++bookmarkId);
            transportBookmarkIds.Add(
                nextBookmarkId.ToString(CultureInfo.InvariantCulture));
            run.AddBeforeSelf(
                new XElement(
                    word + "bookmarkStart",
                    new XAttribute(word + "id", nextBookmarkId),
                    new XAttribute(word + "name", transportName)),
                equation,
                new XElement(
                    word + "bookmarkEnd",
                    new XAttribute(word + "id", nextBookmarkId)));
            run.Remove();
            transportBookmarks.Add(transportName, formula.FormulaId);
        }

        // Word merges directly adjacent inline m:oMath siblings into one physical
        // OMath when this package is opened.  Keep the same invisible ordinary-run
        // boundary used by the mature adjacent-group importer so each source OLE
        // retains one independent formula identity without changing visible text.
        foreach (var paragraph in package.Descendants(word + "p").ToArray())
        {
            var children = paragraph.Elements().ToArray();
            for (var index = 0; index + 1 < children.Length; index++)
            {
                var left = children[index];
                var right = children[index + 1];
                if (left.Name != word + "bookmarkEnd"
                    || right.Name != word + "bookmarkStart"
                    || !transportBookmarkIds.Contains(
                        (string?)left.Attribute(word + "id") ?? string.Empty)
                    || !transportBookmarks.ContainsKey(
                        (string?)right.Attribute(word + "name") ?? string.Empty))
                    continue;
                left.AddAfterSelf(
                    new XElement(
                        word + "r",
                        new XElement(word + "rPr", new XElement(word + "vanish")),
                        new XElement(word + "t", "\u200C")));
            }
        }

        entry.Delete();
        WriteEntry(
            archive,
            "word/document.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + package.Root!.ToString(SaveOptions.DisableFormatting));
    }

    private static bool IsMathTypePackageProgId(string? progId) =>
        !string.IsNullOrWhiteSpace(progId)
        && (progId.StartsWith("Equation.DSMT4", StringComparison.OrdinalIgnoreCase)
            || string.Equals(progId, "DSEquations", StringComparison.OrdinalIgnoreCase)
            || string.Equals(progId, "Equation", StringComparison.OrdinalIgnoreCase));

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
        // Complete Range.WordOpenXML captured from the Word-normalized OMath in
        // an isolated automation instance. Single interactive inserts can replay
        // this directly into the user's target Range without opening/importing
        // the temporary DOCX in that target Application.
        internal string? InsertionWordOpenXml { get; set; }
    }

    internal static BatchSource CreateBatchSource(
        Application application,
        IReadOnlyList<(string FormulaId, string MathMl)> formulas,
        Func<string, string, string>? transformOmml = null,
        string? mathFontName = null,
        bool normalizeMaterializedSource = false)
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
        if (normalizeMaterializedSource && formulas.Count == 1)
        {
            // Single interactive insertion does not need a temporary DOCX at all.
            // InsertXML accepts a minimal Flat OPC package as long as it contains
            // the document part plus an empty document-relationships part. Word
            // itself performs the final OMath materialization in the target range.
            // This avoids both target-Application document activation and the
            // ~2s cost of launching a second WINWORD process for every formula.
            foreach (var entry in entries.Values)
                entry.InsertionWordOpenXml =
                    CreateSingleEquationFlatOpc(entry, normalizedMathFontName);
            return new BatchSource(
                document: null,
                path: string.Empty,
                entries,
                normalizedMathFontName,
                fileBacked: true);
        }

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
            if (normalizeMaterializedSource)
                entries = ReadMaterializedBatchEntries(document, entries);
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

    private static Dictionary<string, BatchEntry> ReadMaterializedBatchEntriesIsolated(
        string path,
        IReadOnlyDictionary<string, BatchEntry> prepared)
    {
        Application? isolatedApplication = null;
        Document? isolatedDocument = null;
        try
        {
            isolatedApplication = new Application
            {
                Visible = false,
                DisplayAlerts = WdAlertLevel.wdAlertsNone,
                ScreenUpdating = false,
            };
            isolatedDocument = isolatedApplication.Documents.Open(
                FileName: path,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            return ReadMaterializedBatchEntries(isolatedDocument, prepared);
        }
        finally
        {
            if (isolatedDocument is not null)
            {
                try { isolatedDocument.Saved = true; } catch { }
                try
                {
                    isolatedDocument.Close(
                        WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(isolatedDocument);
            if (isolatedApplication is not null)
            {
                try
                {
                    isolatedApplication.Quit(
                        WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(isolatedApplication);
        }
    }

    private static Dictionary<string, BatchEntry> ReadMaterializedBatchEntries(
        Document document, IReadOnlyDictionary<string, BatchEntry> prepared)
    {
        // Export the read-only source before the target edit's Undo starts. Its
        // actual Word-normalized tree is what FormattedText will transfer; validate
        // it against the requested content instead of guessing Word's defaults.
        var package = XDocument.Parse(WordDocumentXml.Read(document));
        XNamespace w = WordNamespace;
        XNamespace m = MathNamespace;
        var body = package.Descendants(w + "body").Single();
        if (body.Descendants(m + "oMath").Count() != prepared.Count)
            throw new InvalidDataException("The normalized OMML source inventory is incomplete.");
        var anchors = body.Descendants(w + "bookmarkStart")
            .ToDictionary(e => (string?)e.Attribute(w + "name") ?? string.Empty, StringComparer.Ordinal);
        var result = new Dictionary<string, BatchEntry>(StringComparer.OrdinalIgnoreCase);
        var claimed = new HashSet<XElement>();
        foreach (var pair in prepared)
        {
            var entry = pair.Value;
            if (!anchors.TryGetValue(entry.BookmarkName, out var anchor))
                throw new InvalidDataException("The normalized OMML source lost its transport identity.");
            var paragraph = anchor.Ancestors(w + "p").FirstOrDefault()
                ?? throw new InvalidDataException("The normalized source anchor has no paragraph.");
            var equation = paragraph.Descendants(m + "oMath").Single();
            if (!claimed.Add(equation))
                throw new InvalidDataException("Two source identities own the same normalized OMath.");
            var actual = equation.ToString(SaveOptions.DisableFormatting);
            var fingerprint = ComputeVerifiedMaterializedOmmlFingerprint(entry.Omml, actual);

            Bookmarks? documentBookmarks = null;
            Bookmark? materializedBookmark = null;
            Range? materializedBookmarkRange = null;
            OMaths? materializedMaths = null;
            OMath? materializedMath = null;
            Range? materializedMathRange = null;
            string insertionWordOpenXml;
            try
            {
                documentBookmarks = document.Bookmarks;
                if (!documentBookmarks.Exists(entry.BookmarkName))
                    throw new InvalidDataException(
                        $"The normalized OMML source lost COM bookmark {entry.BookmarkName}.");
                materializedBookmark = documentBookmarks[entry.BookmarkName];
                materializedBookmarkRange = materializedBookmark.Range;
                materializedMaths = materializedBookmarkRange.OMaths;
                if (materializedMaths.Count != 1)
                    throw new InvalidDataException(
                        $"The normalized OMML source bookmark {entry.BookmarkName} owns {materializedMaths.Count} equations.");
                materializedMath = materializedMaths[1];
                materializedMathRange = materializedMath.Range.Duplicate;
                insertionWordOpenXml = materializedMathRange.WordOpenXML
                    ?? throw new InvalidDataException(
                        $"The normalized OMML source {entry.BookmarkName} did not export WordOpenXML.");
                // Preserve Word's Flat OPC byte-for-byte at this layer. Parsing
                // and reserializing Range.WordOpenXML before InsertXML can make
                // Word 2021 reject an otherwise valid package. The transport
                // bookmark is removed from the live target immediately after the
                // OMath is materialized below.
            }
            finally
            {
                Release(materializedMathRange);
                Release(materializedMath);
                Release(materializedMaths);
                Release(materializedBookmarkRange);
                Release(materializedBookmark);
                Release(documentBookmarks);
            }

            var materializedEntry = new BatchEntry(
                entry.BookmarkName,
                fingerprint,
                actual,
                entry.BookmarkId)
            {
                InsertionWordOpenXml = insertionWordOpenXml,
            };
            result.Add(pair.Key, materializedEntry);
        }
        return result;
    }

    private static string RemoveBatchTransportBookmarkFromWordOpenXml(
        string wordOpenXml,
        string bookmarkName)
    {
        if (string.IsNullOrWhiteSpace(wordOpenXml)
            || string.IsNullOrWhiteSpace(bookmarkName))
            return wordOpenXml;

        var package = XDocument.Parse(
            wordOpenXml,
            LoadOptions.PreserveWhitespace);
        XNamespace w = WordNamespace;
        var starts = package
            .Descendants(w + "bookmarkStart")
            .Where(element => string.Equals(
                (string?)element.Attribute(w + "name"),
                bookmarkName,
                StringComparison.Ordinal))
            .ToArray();
        if (starts.Length == 0) return wordOpenXml;

        var ids = starts
            .Select(element => (string?)element.Attribute(w + "id"))
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .ToHashSet(StringComparer.Ordinal);
        foreach (var start in starts)
            start.Remove();
        foreach (var end in package
                     .Descendants(w + "bookmarkEnd")
                     .Where(element => ids.Contains(
                         (string?)element.Attribute(w + "id") ?? string.Empty))
                     .ToArray())
            end.Remove();
        return package.ToString(SaveOptions.DisableFormatting);
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
                var scratchInlineReplacement = !display
                    && replaceTarget
                    && ShouldUseScratchDocumentForInlineReplacement(
                        targetDocument,
                        insertionRange);
                Range imported;
                if (scratchInlineReplacement)
                {
                    imported = InsertBookmarkedFileThroughScratchDocument(
                        application,
                        targetDocument,
                        insertionRange,
                        tempPath,
                        normalizedMathFontName,
                        omml,
                        out var materializedFingerprint);
                    sourceFingerprint = materializedFingerprint;
                }
                else
                {
                    imported = InsertBookmarkedFile(
                        targetDocument,
                        insertionRange,
                        tempPath,
                        display,
                        replaceTarget);
                }
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

    private static bool ShouldUseScratchDocumentForInlineReplacement(
        Document targetDocument,
        Range insertionRange)
    {
        var configured = Environment.GetEnvironmentVariable(
            "VISUALTEX_DIRECT_INLINE_OMML_INSERT");
        if (string.Equals(configured, "0", StringComparison.Ordinal))
            return true;
        if (string.Equals(configured, "1", StringComparison.Ordinal))
            return false;

        Range? content = null;
        try
        {
            content = targetDocument.Content;
            // InsertFile is fastest for a small local document, but Word walks
            // and repaginates the complete tail when importing near the front of
            // a formula-heavy document. Materialize the OMML in the cached tiny
            // scratch document and transfer its already-built FormattedText when
            // the target has a substantial tail. This keeps the same professional
            // OMath tree while making the edit cost local instead of proportional
            // to every later equation.
            return content.End - insertionRange.End >= 256;
        }
        catch
        {
            return false;
        }
        finally { Release(content); }
    }

    internal static bool TryPrimeInlineScratchDocumentForLargeTail(
        Application application,
        Document targetDocument,
        Range equationRange,
        string capturedWordOpenXml,
        string? mathFontName)
    {
        if (!ShouldUseScratchDocumentForInlineReplacement(
                targetDocument,
                equationRange))
            return false;

        lock (InlineScratchLock)
        {
            Range? scratchContent = null;
            Range? scratchTarget = null;
            OMaths? maths = null;
            OMath? scratchMath = null;
            Range? scratchFormula = null;
            Document? transferWarmDocument = null;
            Range? transferWarmContent = null;
            Range? transferWarmTarget = null;
            OMaths? transferWarmMaths = null;
            var tempPath = string.Empty;
            try
            {
                var scratch = GetOrCreateInlineScratchDocument(
                    application,
                    NormalizeMathFontName(mathFontName));
                if (_inlineScratchPrimed) return true;

                var omml = ExtractSingleOMath(capturedWordOpenXml);
                tempPath = CreateTemporaryDocx(
                    omml,
                    includeLeadingTab: false,
                    forceInline: true,
                    mathFontName: NormalizeMathFontName(mathFontName));
                scratchContent = scratch.Content;
                var scratchStart = scratchContent.Start;
                scratchContent.Text = "L" + InlineScratchPlaceholder + "R";
                Release(scratchContent);
                scratchContent = null;
                scratchTarget = scratch.Range(
                    scratchStart + 1,
                    scratchStart + 1 + InlineScratchPlaceholder.Length);
                scratchTarget.InsertFile(
                    FileName: tempPath,
                    Range: FormulaBookmarkName,
                    ConfirmConversions: false,
                    Link: false,
                    Attachment: false);
                maths = scratch.OMaths;
                if (maths.Count != 1)
                    throw new InvalidDataException(
                        "The warmed OMML scratch document did not contain exactly one equation.");
                scratchMath = maths[1];
                scratchFormula = scratchMath.Range;

                // The first cross-document FormattedText assignment in a Word
                // process pays an additional OMath/style initialization cost.
                // Pay it while the editor is opening (where the health/session
                // work already has latency headroom), not after the user presses
                // Apply. This transfer is isolated to disposable hidden documents
                // and never mutates the user's target.
                transferWarmDocument = application.Documents.Add(Visible: false);
                ApplyDocumentMathFont(
                    transferWarmDocument,
                    NormalizeMathFontName(mathFontName));
                transferWarmContent = transferWarmDocument.Content;
                var transferStart = transferWarmContent.Start;
                transferWarmContent.Text = "L" + InlineScratchPlaceholder + "R";
                Release(transferWarmContent);
                transferWarmContent = null;
                transferWarmTarget = transferWarmDocument.Range(
                    transferStart + 1,
                    transferStart + 1 + InlineScratchPlaceholder.Length);
                transferWarmTarget.FormattedText = scratchFormula.FormattedText;
                transferWarmMaths = transferWarmDocument.OMaths;
                if (transferWarmMaths.Count != 1)
                    throw new InvalidDataException(
                        "The warmed OMML FormattedText transfer did not contain exactly one equation.");
                _inlineScratchPrimed = true;
                try { scratch.Saved = true; } catch { }
                try { transferWarmDocument.Saved = true; } catch { }
                return true;
            }
            catch
            {
                _inlineScratchPrimed = false;
                return false;
            }
            finally
            {
                Release(transferWarmMaths);
                Release(transferWarmTarget);
                Release(transferWarmContent);
                Release(scratchFormula);
                Release(scratchMath);
                Release(maths);
                Release(scratchTarget);
                Release(scratchContent);
                if (transferWarmDocument is not null)
                {
                    try { transferWarmDocument.Saved = true; } catch { }
                    try
                    {
                        transferWarmDocument.Close(
                            WdSaveOptions.wdDoNotSaveChanges);
                    }
                    catch { }
                }
                Release(transferWarmDocument);
                try { targetDocument.Activate(); } catch { }
                if (!string.IsNullOrWhiteSpace(tempPath))
                {
                    try { File.Delete(tempPath); } catch { }
                }
            }
        }
    }

    private static Range ReplaceExistingOmmlNodesInParagraph(
        Document targetDocument,
        Range targetRange,
        string omml,
        bool display,
        string? mathFontName)
    {
        XNamespace math =
            MathNamespace;

        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        OMaths? paragraphMaths = null;
        OMath? paragraphMath = null;
        Range? exact = null;
        OMath? insertedMath = null;
        Range? result = null;
        try
        {
            paragraphs =
                targetRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "The adjacent inline OMath merge spans more than one Word paragraph.");
            paragraph =
                paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            paragraphMaths =
                paragraphRange.OMaths;

            var targetOrdinals =
                new List<int>();
            for (var index = 1;
                 index <= paragraphMaths.Count;
                 index++)
            {
                Release(exact);
                exact = null;
                Release(paragraphMath);
                paragraphMath =
                    paragraphMaths[index];
                exact =
                    paragraphMath.Range.Duplicate;
                if (exact.Start >= targetRange.Start
                    && exact.End <= targetRange.End)
                {
                    targetOrdinals.Add(
                        index - 1);
                }
            }

            if (targetOrdinals.Count == 0)
                throw new InvalidDataException(
                    "The requested OMath replacement range contains no complete Word OMath.");
            for (var index = 1;
                 index < targetOrdinals.Count;
                 index++)
            {
                if (targetOrdinals[index]
                    != targetOrdinals[index - 1] + 1)
                    throw new InvalidDataException(
                        "The requested OMath replacement range contains non-contiguous Word OMath hosts.");
            }

            var paragraphXml =
                XDocument.Parse(
                    paragraphRange.WordOpenXML,
                    LoadOptions.PreserveWhitespace);
            var xmlMaths =
                paragraphXml
                    .Descendants(
                        math + "oMath")
                    .Where(element =>
                        !element.Ancestors(
                            math + "oMath")
                            .Any())
                    .ToArray();
            if (xmlMaths.Length !=
                paragraphMaths.Count)
                throw new InvalidDataException(
                    "Word paragraph OMath COM/XML inventories disagree during native inline merge.");

            var replacement =
                XElement.Parse(
                    ApplyExplicitTransferMathFont(
                        ExtractSingleOMath(
                            omml),
                        NormalizeMathFontName(
                            mathFontName)),
                    LoadOptions.PreserveWhitespace);

            var firstOrdinal =
                targetOrdinals[0];
            xmlMaths[firstOrdinal]
                .ReplaceWith(
                    replacement);
            for (var index =
                     targetOrdinals.Count - 1;
                 index >= 1;
                 index--)
            {
                xmlMaths[targetOrdinals[index]]
                    .Remove();
            }

            var insertionStart =
                targetRange.Start;
            paragraphRange.InsertXML(
                paragraphXml.ToString(
                    SaveOptions.DisableFormatting));

            insertedMath =
                FindMathAtPosition(
                    targetDocument,
                    insertionStart,
                    ResolveCurrentParagraphEnd(
                        targetDocument,
                        insertionStart))
                ?? throw new InvalidOperationException(
                    "Word did not materialize the paragraph-local OMath node replacement.");

            var targetType =
                display
                    ? WdOMathType.wdOMathDisplay
                    : WdOMathType.wdOMathInline;
            if (insertedMath.Type !=
                targetType)
                insertedMath.Type =
                    targetType;
            insertedMath.BuildUp();

            result =
                insertedMath.Range.Duplicate;
            var returned =
                result;
            result = null;
            return returned;
        }
        finally
        {
            Release(result);
            Release(insertedMath);
            Release(exact);
            Release(paragraphMath);
            Release(paragraphMaths);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static Range ReplaceWithPreparedOmmlInParagraph(
        Document targetDocument,
        Range targetRange,
        string omml,
        bool display,
        string? mathFontName = null)
    {
        if (targetDocument is null)
            throw new ArgumentNullException(nameof(targetDocument));
        if (targetRange is null)
            throw new ArgumentNullException(nameof(targetRange));
        if (string.IsNullOrWhiteSpace(omml))
            throw new InvalidDataException(
                "The prepared paragraph-local OMML replacement is empty.");

        OMaths? existingTargetMaths = null;
        OMath? existingTargetMath = null;
        Range? existingTargetExact = null;
        try
        {
            existingTargetMaths =
                targetRange.OMaths;
            if (existingTargetMaths.Count == 1)
            {
                existingTargetMath =
                    existingTargetMaths[1];
                existingTargetExact =
                    existingTargetMath.Range.Duplicate;
                if (existingTargetExact.StoryType ==
                        targetRange.StoryType
                    && existingTargetExact.Start ==
                        targetRange.Start
                    && existingTargetExact.End ==
                        targetRange.End)
                {
                    return ReplaceExistingOmmlNodesInParagraph(
                        targetDocument,
                        targetRange,
                        omml,
                        display,
                        mathFontName);
                }
            }
        }
        finally
        {
            Release(existingTargetExact);
            Release(existingTargetMath);
            Release(existingTargetMaths);
        }

        XNamespace word =
            WordNamespace;
        XNamespace math =
            MathNamespace;
        var markerText =
            "VTX"
            + Guid.NewGuid().ToString("N");
        var insertionStart =
            targetRange.Start;

        Range? markerRange = null;
        Range? followingEmptyParagraphGuard = null;
        string? followingEmptyParagraphGuardText = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        OMath? insertedMath = null;
        Range? result = null;
        try
        {
            if (display)
            {
                followingEmptyParagraphGuard =
                    TryCreateFollowingEmptyParagraphGuard(
                        targetDocument,
                        targetRange,
                        out followingEmptyParagraphGuardText);
            }

            markerRange =
                targetRange.Duplicate;
            markerRange.Text =
                markerText;
            markerRange.SetRange(
                insertionStart,
                insertionStart + markerText.Length);

            paragraphs =
                markerRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "The OMML insertion marker does not belong to exactly one Word paragraph.");
            paragraph =
                paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;

            var paragraphXml =
                XDocument.Parse(
                    paragraphRange.WordOpenXML,
                    LoadOptions.PreserveWhitespace);
            var ordinaryMarkerNodes =
                paragraphXml
                    .Descendants(
                        word + "t")
                    .Where(node =>
                        node.Value.IndexOf(
                            markerText,
                            StringComparison.Ordinal)
                        >= 0)
                    .ToArray();
            var mathMarkerNodes =
                paragraphXml
                    .Descendants(
                        math + "t")
                    .Where(node =>
                        node.Value.IndexOf(
                            markerText,
                            StringComparison.Ordinal)
                        >= 0)
                    .ToArray();
            if (ordinaryMarkerNodes.Length
                    + mathMarkerNodes.Length
                != 1)
                throw new InvalidDataException(
                    $"Word paragraph XML contains {ordinaryMarkerNodes.Length} ordinary and {mathMarkerNodes.Length} mathematical copies of the OMML insertion marker.");

            var equation =
                XElement.Parse(
                    ApplyExplicitTransferMathFont(
                        ExtractSingleOMath(
                            omml),
                        NormalizeMathFontName(
                            mathFontName)),
                    LoadOptions.PreserveWhitespace);

            if (ordinaryMarkerNodes.Length == 1)
            {
                var markerNode =
                    ordinaryMarkerNodes[0];
                var sourceRun =
                    markerNode.Parent;
                if (sourceRun is null
                    || sourceRun.Name !=
                        word + "r")
                    throw new InvalidDataException(
                        "The OMML insertion marker is not inside one ordinary Word run.");
                if (sourceRun.Parent is null
                    || sourceRun.Parent.Name !=
                        word + "p")
                    throw new InvalidDataException(
                        "The OMML insertion marker is not a direct child run of its Word paragraph.");

                var markerIndex =
                    markerNode.Value.IndexOf(
                        markerText,
                        StringComparison.Ordinal);
                var prefix =
                    markerNode.Value.Substring(
                        0,
                        markerIndex);
                var suffix =
                    markerNode.Value.Substring(
                        markerIndex + markerText.Length);

                XElement CreateSiblingRun(
                    string text)
                {
                    var run =
                        new XElement(
                            word + "r",
                            sourceRun.Attributes()
                                .Select(attribute =>
                                    new XAttribute(
                                        attribute)));
                    var runProperties =
                        sourceRun.Element(
                            word + "rPr");
                    if (runProperties is not null)
                        run.Add(
                            new XElement(
                                runProperties));
                    var textElement =
                        new XElement(
                            word + "t",
                            text);
                    if (text.Length > 0
                        && (char.IsWhiteSpace(
                                text[0])
                            || char.IsWhiteSpace(
                                text[text.Length - 1])))
                    {
                        textElement.SetAttributeValue(
                            XNamespace.Xml + "space",
                            "preserve");
                    }
                    run.Add(
                        textElement);
                    return run;
                }

                var replacements =
                    new List<object>();
                if (!string.IsNullOrEmpty(
                        prefix))
                    replacements.Add(
                        CreateSiblingRun(
                            prefix));
                replacements.Add(
                    equation);
                if (!string.IsNullOrEmpty(
                        suffix))
                    replacements.Add(
                        CreateSiblingRun(
                            suffix));
                sourceRun.ReplaceWith(
                    replacements.ToArray());
            }
            else
            {
                // Word may absorb a collapsed insertion directly beside an
                // inline OMath into that OMath. In that native case the marker is
                // an m:t node, not ordinary w:t. Splice the new equation's
                // children into the existing m:oMath exactly where Word placed
                // the marker, preserving Word's single merged OMath topology.
                var markerNode =
                    mathMarkerNodes[0];
                var sourceMathRun =
                    markerNode.Parent;
                if (sourceMathRun is null
                    || sourceMathRun.Name !=
                        math + "r"
                    || sourceMathRun.Parent is null
                    || sourceMathRun.Parent.Name !=
                        math + "oMath")
                    throw new InvalidDataException(
                        "The absorbed OMML insertion marker is not one direct math run inside an inline OMath.");

                var markerIndex =
                    markerNode.Value.IndexOf(
                        markerText,
                        StringComparison.Ordinal);
                var prefix =
                    markerNode.Value.Substring(
                        0,
                        markerIndex);
                var suffix =
                    markerNode.Value.Substring(
                        markerIndex + markerText.Length);

                XElement CreateMathSiblingRun(
                    string text)
                {
                    var run =
                        new XElement(
                            math + "r");
                    var runProperties =
                        sourceMathRun.Element(
                            math + "rPr");
                    if (runProperties is not null)
                        run.Add(
                            new XElement(
                                runProperties));
                    run.Add(
                        new XElement(
                            math + "t",
                            text));
                    return run;
                }

                var replacements =
                    new List<object>();
                if (!string.IsNullOrEmpty(
                        prefix))
                    replacements.Add(
                        CreateMathSiblingRun(
                            prefix));
                foreach (var child in
                         equation.Nodes())
                {
                    if (child is XElement element)
                        replacements.Add(
                            new XElement(
                                element));
                    else
                        replacements.Add(
                            child.ToString());
                }
                if (!string.IsNullOrEmpty(
                        suffix))
                    replacements.Add(
                        CreateMathSiblingRun(
                            suffix));

                sourceMathRun.ReplaceWith(
                    replacements.ToArray());
            }

            paragraphRange.InsertXML(
                paragraphXml.ToString(
                    SaveOptions.DisableFormatting));

            insertedMath =
                FindMathAtPosition(
                    targetDocument,
                    insertionStart,
                    ResolveCurrentParagraphEnd(
                        targetDocument,
                        insertionStart))
                ?? throw new InvalidOperationException(
                    "Word did not materialize the paragraph-local OMML replacement.");

            var targetType =
                display
                    ? WdOMathType.wdOMathDisplay
                    : WdOMathType.wdOMathInline;
            if (insertedMath.Type !=
                targetType)
                insertedMath.Type =
                    targetType;
            insertedMath.BuildUp();

            if (followingEmptyParagraphGuard is not null)
            {
                RemoveFollowingEmptyParagraphGuard(
                    targetDocument,
                    insertedMath,
                    followingEmptyParagraphGuardText
                    ?? throw new InvalidDataException(
                        "The OMML paragraph-boundary guard lost its expected text."));
                Release(followingEmptyParagraphGuard);
                followingEmptyParagraphGuard = null;
                followingEmptyParagraphGuardText = null;
            }

            result =
                insertedMath.Range.Duplicate;
            var returned =
                result;
            result = null;
            return returned;
        }
        finally
        {
            if (followingEmptyParagraphGuard is not null
                && !string.IsNullOrEmpty(
                    followingEmptyParagraphGuardText))
            {
                try
                {
                    if (string.Equals(
                            followingEmptyParagraphGuard.Text,
                            followingEmptyParagraphGuardText,
                            StringComparison.Ordinal))
                        followingEmptyParagraphGuard.Text = string.Empty;
                }
                catch { }
            }
            Release(followingEmptyParagraphGuard);
            Release(result);
            Release(insertedMath);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(markerRange);
        }
    }

    private static Range? TryCreateFollowingEmptyParagraphGuard(
        Document document,
        Range targetRange,
        out string? guardText)
    {
        guardText = null;
        Paragraphs? targetParagraphs = null;
        Paragraph? targetParagraph = null;
        Range? targetParagraphRange = null;
        Range? nextProbe = null;
        Paragraphs? nextParagraphs = null;
        Paragraph? nextParagraph = null;
        Range? nextRange = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Frames? frames = null;
        Tables? tables = null;
        Range? insertion = null;
        Range? guardRange = null;
        try
        {
            targetParagraphs = targetRange.Paragraphs;
            if (targetParagraphs.Count != 1)
                return null;
            targetParagraph = targetParagraphs[1];
            targetParagraphRange =
                targetParagraph.Range.Duplicate;
            if (targetParagraphRange.End >=
                document.Content.End)
                return null;

            nextProbe = document.Range(
                targetParagraphRange.End,
                Math.Min(
                    document.Content.End,
                    targetParagraphRange.End + 1));
            if (Convert.ToBoolean(
                    nextProbe.get_Information(
                        WdInformation.wdWithInTable)))
                return null;
            nextParagraphs = nextProbe.Paragraphs;
            if (nextParagraphs.Count != 1)
                return null;
            nextParagraph = nextParagraphs[1];
            nextRange = nextParagraph.Range.Duplicate;
            if (nextRange.Start !=
                targetParagraphRange.End)
                return null;

            tables = nextRange.Tables;
            shapes = nextRange.InlineShapes;
            maths = nextRange.OMaths;
            fields = nextRange.Fields;
            frames = nextRange.Frames;
            if (tables.Count != 0
                || shapes.Count != 0
                || maths.Count != 0
                || fields.Count != 0
                || frames.Count != 0)
                return null;

            var text =
                nextRange.Text
                ?? string.Empty;
            foreach (var character in text)
            {
                if (character is '\r' or '\n'
                    or '\t' or '\v' or '\a'
                    || char.IsWhiteSpace(character))
                    continue;
                return null;
            }

            var token =
                "VTOB"
                + Guid.NewGuid().ToString("N");
            insertion = document.Range(
                nextRange.Start,
                nextRange.Start);
            insertion.Text = token;
            guardRange = document.Range(
                nextRange.Start,
                nextRange.Start + token.Length);
            if (!string.Equals(
                    guardRange.Text,
                    token,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Word did not retain the temporary OMML paragraph-boundary guard.");

            guardText = token;
            var result = guardRange;
            guardRange = null;
            return result;
        }
        finally
        {
            Release(guardRange);
            Release(insertion);
            Release(tables);
            Release(frames);
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(nextRange);
            Release(nextParagraph);
            Release(nextParagraphs);
            Release(nextProbe);
            Release(targetParagraphRange);
            Release(targetParagraph);
            Release(targetParagraphs);
        }
    }

    private static void RemoveFollowingEmptyParagraphGuard(
        Document document,
        OMath insertedMath,
        string expectedText)
    {
        Range? mathRange = null;
        Paragraphs? mathParagraphs = null;
        Paragraph? mathParagraph = null;
        Range? mathParagraphRange = null;
        Range? nextProbe = null;
        Paragraphs? nextParagraphs = null;
        Paragraph? nextParagraph = null;
        Range? nextRange = null;
        Range? guard = null;
        try
        {
            mathRange = insertedMath.Range.Duplicate;
            mathParagraphs = mathRange.Paragraphs;
            if (mathParagraphs.Count != 1)
                throw new InvalidDataException(
                    "The inserted display OMML no longer belongs to exactly one Word paragraph while restoring its boundary.");
            mathParagraph = mathParagraphs[1];
            mathParagraphRange =
                mathParagraph.Range.Duplicate;
            if (mathParagraphRange.End >=
                document.Content.End)
                throw new InvalidDataException(
                    "The following OMML paragraph-boundary guard disappeared at the end of the document.");

            nextProbe = document.Range(
                mathParagraphRange.End,
                Math.Min(
                    document.Content.End,
                    mathParagraphRange.End + 1));
            nextParagraphs = nextProbe.Paragraphs;
            if (nextParagraphs.Count != 1)
                throw new InvalidDataException(
                    "The preserved paragraph after display OMML has ambiguous Word paragraph ownership.");
            nextParagraph = nextParagraphs[1];
            nextRange = nextParagraph.Range.Duplicate;
            if (nextRange.Start !=
                mathParagraphRange.End)
                throw new InvalidDataException(
                    "Word moved the preserved paragraph away from the display OMML boundary.");
            var nextText = nextRange.Text
                ?? string.Empty;
            if (!nextText.StartsWith(
                    expectedText,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Word moved or changed the temporary OMML paragraph-boundary guard while replacing the equation paragraph.");

            guard = document.Range(
                nextRange.Start,
                nextRange.Start + expectedText.Length);
            if (!string.Equals(
                    guard.Text,
                    expectedText,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The OMML paragraph-boundary guard cannot be resolved at the preserved paragraph start.");
            guard.Text = string.Empty;
        }
        finally
        {
            Release(guard);
            Release(nextRange);
            Release(nextParagraph);
            Release(nextParagraphs);
            Release(nextProbe);
            Release(mathParagraphRange);
            Release(mathParagraph);
            Release(mathParagraphs);
            Release(mathRange);
        }
    }

    internal static Range ReplaceWithPreparedOmmlDirect(
        Document targetDocument,
        Range targetRange,
        string omml,
        bool display,
        string? mathFontName = null)
    {
        if (targetDocument is null)
            throw new ArgumentNullException(nameof(targetDocument));
        if (targetRange is null)
            throw new ArgumentNullException(nameof(targetRange));
        if (string.IsNullOrWhiteSpace(omml))
            throw new InvalidDataException(
                "The prepared direct OMML replacement is empty.");

        var bookmarkName =
            "VisualTeXDirect"
            + Guid.NewGuid().ToString("N").Substring(0, 16);
        var entry = new BatchEntry(
            bookmarkName,
            ComputeOmmlFingerprint(omml),
            omml,
            31902);
        var flatOpc =
            CreateSingleEquationFlatOpc(
                entry,
                NormalizeMathFontName(mathFontName));

        Range? target = null;
        Range? content = null;
        Range? rightBefore = null;
        Range? transportBreak = null;
        Range? rightAfterBreak = null;
        OMath? insertedMath = null;
        Range? result = null;
        Bookmarks? bookmarks = null;
        Bookmark? imported = null;
        Range? importedRange = null;
        try
        {
            target = targetRange.Duplicate;
            var insertionStart = target.Start;

            content = targetDocument.Content;
            var rightBoundaryLength = Math.Min(
                32,
                Math.Max(
                    0,
                    content.End - target.End));
            var rightBoundaryBefore = string.Empty;
            if (rightBoundaryLength > 0)
            {
                rightBefore = targetDocument.Range(
                    target.End,
                    target.End + rightBoundaryLength);
                rightBoundaryBefore =
                    rightBefore.Text
                    ?? string.Empty;
            }
            var hasImmediateRightUserText =
                rightBoundaryBefore.Length > 0
                && rightBoundaryBefore[0] != '\r'
                && rightBoundaryBefore[0] != '\u0007';

            target.InsertXML(flatOpc);

            insertedMath = FindMathAtPosition(
                    targetDocument,
                    insertionStart,
                    target.End)
                ?? throw new InvalidOperationException(
                    "Word did not materialize the direct OMML replacement.");
            var targetType = display
                ? WdOMathType.wdOMathDisplay
                : WdOMathType.wdOMathInline;
            if (insertedMath.Type != targetType)
                insertedMath.Type = targetType;

            result = insertedMath.Range.Duplicate;

            if (hasImmediateRightUserText)
            {
                Release(content);
                content = targetDocument.Content;
                if (result.End < content.End)
                {
                    transportBreak = targetDocument.Range(
                        result.End,
                        Math.Min(
                            content.End,
                            result.End + 1));
                    if (string.Equals(
                            transportBreak.Text,
                            "\r",
                            StringComparison.Ordinal))
                    {
                        var comparableLength = Math.Min(
                            rightBoundaryBefore.Length,
                            Math.Max(
                                0,
                                content.End - (result.End + 1)));
                        if (comparableLength > 0)
                        {
                            rightAfterBreak = targetDocument.Range(
                                result.End + 1,
                                result.End + 1 + comparableLength);
                            var afterText =
                                rightAfterBreak.Text
                                ?? string.Empty;
                            var expected =
                                rightBoundaryBefore.Substring(
                                    0,
                                    comparableLength);
                            if (string.Equals(
                                    afterText,
                                    expected,
                                    StringComparison.Ordinal))
                            {
                                transportBreak.Delete();

                                Release(result);
                                result = null;
                                Release(insertedMath);
                                insertedMath = null;
                                Release(content);
                                content = targetDocument.Content;
                                insertedMath = FindMathAtPosition(
                                        targetDocument,
                                        insertionStart,
                                        Math.Min(
                                            content.End,
                                            insertionStart + 256))
                                    ?? throw new InvalidOperationException(
                                        "Word lost the direct OMML replacement after transport paragraph cleanup.");
                                result =
                                    insertedMath.Range.Duplicate;
                            }
                        }
                    }
                }
            }

            bookmarks = targetDocument.Bookmarks;
            if (bookmarks.Exists(bookmarkName))
            {
                imported = bookmarks[bookmarkName];
                importedRange = imported.Range.Duplicate;
                if (importedRange.StoryType == result.StoryType
                    && importedRange.Start <= result.End
                    && importedRange.End >= result.Start)
                    imported.Delete();
            }

            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(importedRange);
            Release(imported);
            Release(bookmarks);
            Release(result);
            Release(insertedMath);
            Release(rightAfterBreak);
            Release(transportBreak);
            Release(rightBefore);
            Release(content);
            Release(target);
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
        string mathFontName,
        string preparedOmml,
        out string materializedFingerprint)
    {
        lock (InlineScratchLock)
            return InsertBookmarkedFileThroughScratchDocumentCore(
                application,
                targetDocument,
                insertionRange,
                filePath,
                mathFontName,
                preparedOmml,
                out materializedFingerprint);
    }

    private static Range InsertBookmarkedFileThroughScratchDocumentCore(
        Application application,
        Document targetDocument,
        Range insertionRange,
        string filePath,
        string mathFontName,
        string preparedOmml,
        out string materializedFingerprint)
    {
        Document? scratchDocument = null;
        Range? scratchInsertion = null;
        OMaths? scratchMaths = null;
        OMath? scratchMath = null;
        Range? scratchFormula = null;
        Range? target = null;
        OMath? insertedMath = null;
        Range? result = null;
        materializedFingerprint = string.Empty;
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
            materializedFingerprint = ComputeVerifiedMaterializedOmmlFingerprint(
                preparedOmml,
                scratchFormula.WordOpenXML);

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
            _inlineScratchPrimed = false;
        }

        _inlineScratchDocument = application.Documents.Add(Visible: false);
        _inlineScratchPrimed = false;
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

    internal static Range InsertXmlDirect(
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
        mathMl = NormalizeMathMlPrescripts(mathMl);
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
        omml = NormalizeOptionalPrescriptSlots(omml);
        omml = NormalizeDisplayNaryOmml(omml, display);
        omml = NormalizeGeneratedOmmlArgumentOrder(omml);
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
        // Presentation MathML is already parsed mathematical content. Literal
        // mtext (including paths, command documentation and colored text) must
        // not be interpreted as LaTeX a second time. The TeX producer reports
        // unknown commands at parse time; this boundary rejects real error nodes.
        // Use expanded XML names, so namespace aliases cannot bypass validation.
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            MaxCharactersInDocument = 4_000_000,
        };
        using var text = new StringReader(mathMl);
        using var reader = XmlReader.Create(text, settings);
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        if (document.Root?.Name != presentationMath + "math")
            throw new InvalidDataException("The OMML source must have a Presentation MathML math root.");
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
        // Conservative fast gate; the parsed expanded name below is authoritative.
        // A namespace alias such as p:mtable must enter the same normalization.
        if (string.IsNullOrWhiteSpace(mathMl)
            || mathMl.IndexOf("mtable", StringComparison.Ordinal) < 0)
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

        // The generated and Word-materialized trees obey the SAME argument
        // contract. Detect malformed structures before document mutation, not
        // only when a matrix happens to trigger a stricter validator.
        var target = XDocument.Parse(ExtractSingleOMath(omml), LoadOptions.PreserveWhitespace);
        ValidateNoVisibleEmptyOmmlSlots(target);
        var source = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        var sourceTables = source
            .Descendants(presentationMath + "mtable")
            .Where(table => IsMatrixLikeMathMlTable(table, presentationMath))
            .ToArray();
        if (sourceTables.Length == 0) return;
        var targetMatrices = target.Descendants(officeMath + "m").ToArray();
        foreach (var matrix in targetMatrices)
        {
            if (!IsMathBooleanEnabled(matrix.Element(officeMath + "mPr"), officeMath + "plcHide"))
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
        var conditionallyUprightSingleTokens =
            new HashSet<string>(StringComparer.Ordinal);
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
            if (canonical.Length == 1)
                conditionallyUprightSingleTokens.Add(canonical);
            else if (canonical.Length > 1)
                uprightTokens.Add(canonical);
            foreach (Match match in Regex.Matches(text, @"\p{L}+"))
            {
                // A single upright source letter is not a document-wide style
                // declaration. OMML reverse conversion can represent "sin" as
                // three normal mi tokens; adding "n" here used to upright every
                // later variable n in the same formula. Multi-letter names remain
                // safe for the coalesced-run matcher below.
                if (match.Value.Length > 1)
                    uprightWords.Add(match.Value);
            }
        }
        if (uprightTokens.Count == 0
            && conditionallyUprightSingleTokens.Count == 0)
            return omml;

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

        bool HasExistingUprightMathStyle(XElement run)
        {
            var properties = run.Element(math + "rPr");
            if (properties is null) return false;
            if (properties.Element(math + "nor") is not null) return true;
            return string.Equals(
                properties.Element(math + "sty")?.Attribute(math + "val")?.Value,
                "p",
                StringComparison.OrdinalIgnoreCase);
        }

        foreach (var run in ommlDocument.Descendants(math + "r").ToList())
        {
            var text = string.Concat(run.Elements(math + "t").Select(element => element.Value));
            if (string.IsNullOrEmpty(text)) continue;
            var canonicalRunText = CanonicalToken(text);
            var isConditionallyUprightRun =
                canonicalRunText.Length > 0
                && HasExistingUprightMathStyle(run)
                && canonicalRunText.All(character =>
                    conditionallyUprightSingleTokens.Contains(character.ToString()));
            if (uprightTokens.Contains(canonicalRunText)
                || isConditionallyUprightRun)
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

    internal static string NormalizeMathMlPrescripts(string mathMl)
    {
        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        XNamespace p = "http://www.w3.org/1998/Math/MathML";
        var scriptNames = new HashSet<XName> { p + "msub", p + "msup", p + "msubsup" };
        // Only these containers infer a sequence of expressions. In a fraction,
        // root, script, etc. the next child is ANOTHER ARGUMENT, not the next atom.
        var sequenceNames = new HashSet<XName>
        {
            p + "math", p + "mrow", p + "mstyle", p + "mtd", p + "msqrt",
            p + "mphantom", p + "mpadded", p + "menclose",
        };
        var baseNames = new HashSet<XName>
        {
            p + "mi", p + "mn", p + "mrow", p + "mstyle", p + "mfrac",
            p + "msqrt", p + "mroot", p + "msub", p + "msup", p + "msubsup",
            p + "mmultiscripts", p + "mover", p + "munder", p + "munderover", p + "mfenced",
        };
        bool IsEmptyBase(XElement element) =>
            (element.Name == p + "mrow" || element.Name == p + "mi"
                || element.Name == p + "mn" || element.Name == p + "mtext")
            && element.Value.Length == 0
            && element.Elements().All(IsEmptyBase);

        foreach (var script in document.Descendants()
                     .Where(element => scriptNames.Contains(element.Name)).Reverse().ToArray())
        {
            if (script.Parent is null || !sequenceNames.Contains(script.Parent.Name)) continue;
            var arguments = script.Elements().ToArray();
            if (arguments.Length != (script.Name == p + "msubsup" ? 3 : 2)
                || !IsEmptyBase(arguments[0])) continue;
            var basis = script.ElementsAfterSelf().FirstOrDefault();
            if (basis is null || !baseNames.Contains(basis.Name)
                || string.IsNullOrWhiteSpace(basis.Value)) continue;
            var sub = script.Name == p + "msup" ? new XElement(p + "none") : new XElement(arguments[1]);
            var sup = script.Name == p + "msub" ? new XElement(p + "none")
                : new XElement(arguments[script.Name == p + "msubsup" ? 2 : 1]);
            if (string.IsNullOrWhiteSpace(sub.Value) && string.IsNullOrWhiteSpace(sup.Value)) continue;

            // {}^a X and {}_b^a X explicitly encode left-side scripts. Convert
            // their source structure to native MathML prescripts so Office emits
            // sPre, rather than permitting a dotted, empty sSup/sSub base. No
            // filler glyphs or changes to the source LaTeX/adjacent operators.
            var native = new XElement(p + "mmultiscripts",
                script.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
                    .Select(attribute => new XAttribute(attribute)),
                new XElement(basis), new XElement(p + "mprescripts"), sub, sup);
            script.ReplaceWith(native);
            basis.Remove();
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
            var hasArgument = op.ElementsAfterSelf().Any();
            var syntheticLimit = new XElement(
                mathMlNamespace + "msub",
                new XElement(op),
                new XElement(mathMlNamespace + "mrow"));
            op.ReplaceWith(syntheticLimit);
            if (!hasArgument)
            {
                syntheticLimit.AddAfterSelf(
                    new XElement(
                        mathMlNamespace + "mrow",
                        new XElement(mathMlNamespace + "mspace", new XAttribute("width", "0em"))));
            }
        }

        bool IsNaryLimit(XElement element)
        {
            if (!limitNames.Contains(element.Name)) return false;
            var op = element.Elements().FirstOrDefault();
            return op?.Name == mathMlNamespace + "mo"
                && !string.IsNullOrEmpty(op.Value)
                && op.Value.All(character => NaryCharacters.IndexOf(character) >= 0);
        }

        // Presentation MathML expresses consecutive operators as a flat sequence:
        // sum_n sum_m <body>. Office's stylesheet consumes only the immediately
        // following mrow as an n-ary operand. Wrapping just the second sum therefore
        // creates an inner m:nary with an empty <m:e/> and leaves <body> outside it.
        // Fold every consecutive n-ary chain from right to left so the innermost
        // operator owns the real body and each outer operator owns that complete
        // nested expression.
        foreach (var parent in document
                     .Descendants()
                     .Where(element => element.Elements().Any())
                     .ToList())
        {
            var children = parent.Elements().ToArray();
            var chains = new List<(int Start, int End, XElement? Operand)>();
            for (var start = 0; start < children.Length; start++)
            {
                if (!IsNaryLimit(children[start])) continue;
                var end = start;
                while (end + 1 < children.Length
                       && IsNaryLimit(children[end + 1]))
                    end++;
                if (end > start)
                {
                    chains.Add((
                        start,
                        end,
                        end + 1 < children.Length ? children[end + 1] : null));
                }
                start = end;
            }

            foreach (var chain in chains.OrderByDescending(item => item.Start))
            {
                XElement operand;
                if (chain.Operand is null)
                {
                    operand = new XElement(
                        mathMlNamespace + "mrow",
                        new XElement(
                            mathMlNamespace + "mspace",
                            new XAttribute("width", "0em")));
                    children[chain.End].AddAfterSelf(operand);
                }
                else if (chain.Operand.Name == mathMlNamespace + "mrow"
                         || chain.Operand.Name == mathMlNamespace + "mstyle")
                {
                    operand = chain.Operand;
                }
                else
                {
                    operand = new XElement(mathMlNamespace + "mrow");
                    chain.Operand.ReplaceWith(operand);
                    operand.Add(chain.Operand);
                }

                for (var index = chain.End; index > chain.Start; index--)
                {
                    var inner = children[index];
                    var outer = children[index - 1];
                    inner.Remove();
                    operand.Remove();
                    var nestedOperand = new XElement(
                        mathMlNamespace + "mrow",
                        inner,
                        operand);
                    outer.AddAfterSelf(nestedOperand);
                    operand = nestedOperand;
                }
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

    internal static string NormalizeOptionalPrescriptSlots(string omml)
    {
        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        foreach (var script in document.Descendants(math + "sPre"))
        {
            if (!HasVisibleMathContent(script.Element(math + "e"))) continue;
            foreach (var name in new[] { "sub", "sup" })
            {
                var slot = script.Element(math + name);
                var other = script.Element(math + (name == "sub" ? "sup" : "sub"));
                if (slot is null || !HasVisibleMathContent(other) || HasVisibleMathContent(slot)) continue;
                // Preserve unfamiliar controls/extensions rather than replacing them.
                if (slot.Descendants().Any(element => element.Name != math + "r" && element.Name != math + "t"
                    && element.Name != math + "rPr")) continue;
                // CT_SPre has a pair of slots but no subHide/supHide. Represent the
                // explicitly absent half as an empty native phantom with zero extent.
                // There is no filler character, duplicated operand, font change or
                // visible placeholder in this structural representation of <none/>.
                slot.ReplaceNodes(new XElement(math + "phant",
                    new XElement(math + "phantPr",
                        new XElement(math + "show", new XAttribute(math + "val", "0")),
                        new XElement(math + "zeroWid", new XAttribute(math + "val", "1")),
                        new XElement(math + "zeroAsc", new XAttribute(math + "val", "1")),
                        new XElement(math + "zeroDesc", new XAttribute(math + "val", "1"))),
                    new XElement(math + "e")));
            }
        }
        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    private static bool IsHiddenPhantom(XElement element)
    {
        XNamespace math = MathNamespace;
        var show = element.Element(math + "phantPr")?.Element(math + "show");
        return element.Name == math + "phant" && show is not null && !ReadMathBooleanValue(show);
    }

    private static string[] RequiredOmmlArgumentNames(string ownerName) => ownerName switch
    {
        "f" => new[] { "num", "den" },
        "rad" => new[] { "deg", "e" },
        "nary" => new[] { "sub", "sup", "e" },
        "sSub" => new[] { "e", "sub" },
        "sSup" => new[] { "e", "sup" },
        "sSubSup" => new[] { "e", "sub", "sup" },
        "sPre" => new[] { "sub", "sup", "e" },
        _ => Array.Empty<string>(),
    };

    internal static string NormalizeGeneratedOmmlArgumentOrder(string omml)
    {
        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        foreach (var owner in document.Descendants().Where(element => element.Name.Namespace == math).Reverse().ToArray())
        {
            var roles = RequiredOmmlArgumentNames(owner.Name.LocalName);
            if (roles.Length == 0) continue;
            var propertyName = math + (owner.Name.LocalName + "Pr");
            var children = owner.Elements().ToArray();
            // Normalize the producer's known argument grammar, never reorder
            // operands within an argument or guess through extensions/duplicates.
            if (roles.Any(role => children.Count(child => child.Name == math + role) != 1)
                || children.Count(child => child.Name == propertyName) > 1
                || children.Any(child => child.Name != propertyName && !roles.Any(role => child.Name == math + role))
                || owner.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
                continue;
            var ordered = children.Where(child => child.Name == propertyName)
                .Concat(roles.Select(role => children.Single(child => child.Name == math + role))).ToArray();
            if (children.SequenceEqual(ordered)) continue;
            // Office MML2OMML emits sPre as e/sub/sup, but CT_SPre requires
            // sub/sup/e. Word repairs this on import. Produce the proper grammar
            // before capture/hash rather than relaxing the materialization check.
            owner.ReplaceNodes(ordered.Select(child => new XElement(child)));
        }
        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
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
        foreach (var owner in document.Descendants().Where(element => element.Name.Namespace == math))
        {
            // An absent required argument is not an intentionally empty argument.
            // Keep this in the shared contract rather than a second post-insert list.
            var required = RequiredOmmlArgumentNames(owner.Name.LocalName);
            foreach (var name in required)
                if (owner.Elements(math + name).Count() != 1)
                    throw new InvalidDataException(
                        $"Office OMML requires exactly one {name} argument inside {owner.Name.LocalName}.");
        }
        foreach (var slot in document.Descendants()
                     .Where(element => structuralSlotNames.Contains(element.Name)))
        {
            if (HasVisibleMathContent(slot)) continue;
            // A phantom explicitly suppresses its inner editor placeholder as well
            // as glyphs. Required-argument cardinality is still checked above.
            if (slot.Ancestors().Any(IsHiddenPhantom)) continue;

            var parent = slot.Parent;
            var hidden = slot.Name == math + "e"
                && parent?.Name == math + "mr"
                && parent.Parent?.Name == math + "m"
                && IsMathBooleanEnabled(
                    parent.Parent.Element(math + "mPr"),
                    math + "plcHide")
                || slot.Name == math + "deg"
                && parent?.Name == math + "rad"
                && IsMathBooleanEnabled(parent.Element(math + "radPr"), math + "degHide")
                || slot.Name == math + "sub"
                && parent?.Name == math + "nary"
                && IsMathBooleanEnabled(parent.Element(math + "naryPr"), math + "subHide")
                || slot.Name == math + "sup"
                && parent?.Name == math + "nary"
                && IsMathBooleanEnabled(parent.Element(math + "naryPr"), math + "supHide")
                // Native prescripts carry a pair of arguments; either half may
                // be absent in the source. The base and the other half must exist.
                || (slot.Name == math + "sub" || slot.Name == math + "sup")
                && parent?.Name == math + "sPre"
                && HasVisibleMathContent(parent.Element(math + "e"))
                && HasVisibleMathContent(parent.Element(slot.Name == math + "sub" ? math + "sup" : math + "sub"));
            if (hidden) continue;

            throw new InvalidDataException(
                $"Office OMML contains a visible empty {slot.Name.LocalName} slot "
                + $"inside {parent?.Name.LocalName ?? "unknown"}.");
        }
    }

    internal static string NormalizeOmmlPlaceholderVisibility(string omml)
    {
        if (string.IsNullOrWhiteSpace(omml)) return omml;
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
        ValidateNoVisibleEmptyOmmlSlots(document);
    }

    private static bool HasVisibleMathContent(XElement? slot)
    {
        if (slot is null) return false;
        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;
        return slot.DescendantsAndSelf()
            .Any(element => (element.Name == math + "t" || element.Name == word + "t")
                && element.Value.Length > 0);
    }

    private static bool IsMathBooleanEnabled(XElement? properties, XName propertyName)
    {
        var property = properties?.Element(propertyName);
        if (property is null) return false;
        return ReadMathBooleanValue(property);
    }

    private static bool ReadMathBooleanValue(XElement property)
    {
        // Element omission is handled by the caller. Only an omitted ATTRIBUTE
        // means true here; empty/unknown values must never erase content as false.
        var value = property.Attribute(XName.Get("val", MathNamespace))?.Value;
        return value switch
        {
            null or "1" or "true" or "on" => true,
            "0" or "false" or "off" => false,
            _ => throw new InvalidDataException(
                $"Invalid OMML boolean value '{value}' for {property.Name.LocalName}."),
        };
    }

    internal static string TransformOmmlToMathMl(string wordOpenXml, bool display)
    {
        using var operationMetric = VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("WordOmmlConverter.TransformOmmlToMathMl");
        var omml = StripVisualTeXNativeEquationNumber(wordOpenXml);
        // Remove only the empty zero-extent representation of a missing native
        // prescript half before reverse conversion; real phantom content is retained.
        var reverseDocument = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace reverseMath = MathNamespace;
        foreach (var script in reverseDocument.Descendants(reverseMath + "sPre"))
            foreach (var slot in script.Elements().Where(element => element.Name == reverseMath + "sub" || element.Name == reverseMath + "sup"))
            {
                var slotChildren = slot.Elements().Take(2).ToArray();
                var phantom = slotChildren.Length == 1 ? slotChildren[0] : null;
                if (phantom is null || !IsHiddenPhantom(phantom) || HasVisibleMathContent(phantom)) continue;
                var properties = phantom.Element(reverseMath + "phantPr");
                if (new[] { "zeroAsc", "zeroDesc", "zeroWid" }.All(name => IsMathBooleanEnabled(properties, reverseMath + name)))
                    slot.RemoveNodes();
            }
        omml = reverseDocument.Root!.ToString(SaveOptions.DisableFormatting);
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
        RestoreWordCanonicalAsteriskOperators(canonicalRoot);
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

    private static void RestoreWordCanonicalAsteriskOperators(
        XElement mathMlRoot)
    {
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        foreach (var token in mathMlRoot
                     .DescendantsAndSelf()
                     .Where(element =>
                         element.Name == mathMl + "mi"
                         || element.Name == mathMl + "mo")
                     .Where(element =>
                         !element.HasElements
                         && string.Equals(element.Value, "*", StringComparison.Ordinal))
                     .ToList())
        {
            // Word canonicalizes U+2217 ASTERISK OPERATOR to ASCII '*' when a
            // professional equation is materialized. Its reverse stylesheet then
            // misclassifies the punctuation character as <mi>, losing both the
            // operator token type and the original \ast semantics. In mathematical
            // token positions, canonicalize that Word spelling back to the Unicode
            // operator. Deliberate text-mode '*' remains mtext and is untouched.
            token.Name = mathMl + "mo";
            token.Value = RestoreWordAsteriskSpelling(token.Value);
        }
    }

    private static string RestoreWordAsteriskSpelling(string mathematicalText)
        => mathematicalText.Replace('*', '∗');

    private static void NormalizeLegacyMathAlphabetSpelling(
        XDocument document,
        XNamespace math)
    {
        // Word may materialize legacy Unicode mathematical letter symbols as a
        // base Latin letter plus m:scr (for example ℜ -> R + fraktur), while the
        // converter can emit the equivalent precomposed Unicode symbol. Canonicalize
        // only the standard one-to-one legacy spellings so a real script/style
        // change still produces a different content signature.
        var spellings = new Dictionary<char, (char Base, string Script)>
        {
            ['ℭ'] = ('C', "fraktur"),
            ['ℌ'] = ('H', "fraktur"),
            ['ℑ'] = ('I', "fraktur"),
            ['ℜ'] = ('R', "fraktur"),
            ['ℨ'] = ('Z', "fraktur"),
            ['ℬ'] = ('B', "script"),
            ['ℰ'] = ('E', "script"),
            ['ℱ'] = ('F', "script"),
            ['ℋ'] = ('H', "script"),
            ['ℐ'] = ('I', "script"),
            ['ℒ'] = ('L', "script"),
            ['ℳ'] = ('M', "script"),
            ['ℛ'] = ('R', "script"),
            ['ℯ'] = ('e', "script"),
            ['ℊ'] = ('g', "script"),
            ['ℴ'] = ('o', "script"),
            ['ℂ'] = ('C', "double-struck"),
            ['ℍ'] = ('H', "double-struck"),
            ['ℕ'] = ('N', "double-struck"),
            ['ℙ'] = ('P', "double-struck"),
            ['ℚ'] = ('Q', "double-struck"),
            ['ℝ'] = ('R', "double-struck"),
            ['ℤ'] = ('Z', "double-struck"),
        };

        foreach (var run in document.Descendants(math + "r").ToArray())
        {
            var text = run.Element(math + "t");
            if (text is null || text.Value.Length != 1
                || !spellings.TryGetValue(text.Value[0], out var spelling))
                continue;

            var properties = run.Element(math + "rPr");
            var script = properties?.Element(math + "scr");
            var currentScript = (string?)script?.Attribute(math + "val");
            if (script is not null
                && !string.Equals(currentScript, spelling.Script, StringComparison.Ordinal))
                continue;

            if (properties is null)
            {
                properties = new XElement(math + "rPr");
                run.AddFirst(properties);
            }
            if (script is null)
            {
                script = new XElement(
                    math + "scr",
                    new XAttribute(math + "val", spelling.Script));
                properties.AddFirst(script);
            }
            else if (script.Attribute(math + "val") is null)
            {
                script.SetAttributeValue(math + "val", spelling.Script);
            }
            text.Value = spelling.Base.ToString();
        }
    }

    private static void NormalizeMathScriptRunBoundaries(
        XDocument document,
        XNamespace math)
    {
        foreach (var run in document.Descendants(math + "r").ToArray())
        {
            var properties = run.Element(math + "rPr");
            if (properties?.Element(math + "scr") is null) continue;
            var texts = run.Elements(math + "t").ToArray();
            if (texts.Length != 1
                || run.Elements().Any(element => element.Name != math + "rPr" && element.Name != math + "t"))
                continue;

            var value = texts[0].Value;
            var segments = new List<(string Text, bool ScriptSensitive)>();
            var builder = new StringBuilder();
            bool? currentSensitive = null;
            for (var index = 0; index < value.Length;)
            {
                var category = CharUnicodeInfo.GetUnicodeCategory(value, index);
                var scalarLength = char.IsHighSurrogate(value[index])
                    && index + 1 < value.Length
                    && char.IsLowSurrogate(value[index + 1])
                    ? 2
                    : 1;
                var sensitive = category is UnicodeCategory.UppercaseLetter
                    or UnicodeCategory.LowercaseLetter
                    or UnicodeCategory.TitlecaseLetter
                    or UnicodeCategory.ModifierLetter
                    or UnicodeCategory.OtherLetter
                    or UnicodeCategory.DecimalDigitNumber
                    or UnicodeCategory.LetterNumber;
                if (currentSensitive.HasValue && currentSensitive.Value != sensitive)
                {
                    segments.Add((builder.ToString(), currentSensitive.Value));
                    builder.Clear();
                }
                currentSensitive = sensitive;
                builder.Append(value, index, scalarLength);
                index += scalarLength;
            }
            if (builder.Length > 0 && currentSensitive.HasValue)
                segments.Add((builder.ToString(), currentSensitive.Value));
            if (segments.Count == 1)
            {
                if (!segments[0].ScriptSensitive)
                {
                    // A math-alphabet script has no defined effect on a
                    // symbol-only run. Word removes it while materializing
                    // operators such as U+2201 COMPLEMENT, so compare the
                    // prepared and native trees through that same canonical
                    // spelling. Script/fraktur/double-struck on letters and
                    // digits remains semantic and is deliberately retained.
                    properties.Element(math + "scr")?.Remove();
                    if (!properties.HasElements
                        && !properties.HasAttributes
                        && string.IsNullOrWhiteSpace(properties.Value))
                        properties.Remove();
                }
                continue;
            }

            foreach (var segment in segments)
            {
                XElement? segmentProperties = new XElement(properties);
                if (!segment.ScriptSensitive)
                {
                    segmentProperties.Element(math + "scr")?.Remove();
                    if (!segmentProperties.HasElements
                        && !segmentProperties.HasAttributes
                        && string.IsNullOrWhiteSpace(segmentProperties.Value))
                        segmentProperties = null;
                }
                var segmentText = new XElement(math + "t", segment.Text);
                foreach (var attribute in texts[0].Attributes())
                    segmentText.SetAttributeValue(attribute.Name, attribute.Value);
                var replacement = new XElement(math + "r");
                if (segmentProperties is not null) replacement.Add(segmentProperties);
                replacement.Add(segmentText);
                run.AddBeforeSelf(replacement);
            }
            run.Remove();
        }
    }

    private static string NormalizeImportedMathTextSpelling(string mathematicalText) =>
        RestoreWordAsteriskSpelling(
            mathematicalText
                .Replace('\u2212', '-')
                .Replace("\u2032", "'")
                .Replace("\u2033", "''")
                .Replace("\u2034", "'''")
                .Replace("\u2057", "''''"));

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

    // v2 uses the SAME mathematical-content model as import verification. XML
    // namespace aliases and equivalent default properties are not content edits.
    // The prefix is part of the persisted value; never silently reinterpret v1.
    internal static string ComputeOmmlFingerprint(string wordOpenXml) =>
        OmmlFingerprintFormat.CanonicalPrefix + ComputeImportedOmmlContentSignature(wordOpenXml);

    internal static string ComputeOmmlFingerprintForExpectedVersion(
        string wordOpenXml, string expectedFingerprint) =>
        OmmlFingerprintFormat.GetVersion(expectedFingerprint) == 1
            ? ComputeLegacyOmmlFingerprint(wordOpenXml)
            : ComputeOmmlFingerprint(wordOpenXml);

    internal static bool MatchesStoredOmmlFingerprint(string wordOpenXml, string? expectedFingerprint) =>
        !string.IsNullOrWhiteSpace(expectedFingerprint)
        && string.Equals(ComputeOmmlFingerprintForExpectedVersion(wordOpenXml, expectedFingerprint!),
            expectedFingerprint, StringComparison.OrdinalIgnoreCase);

    internal static string UpgradeVerifiedOmmlFingerprint(string wordOpenXml, string storedFingerprint)
    {
        if (!MatchesStoredOmmlFingerprint(wordOpenXml, storedFingerprint))
            throw new InvalidDataException("The captured equation does not match its stored OMML fingerprint.");
        return ComputeOmmlFingerprint(wordOpenXml);
    }

    // Frozen v1 reader for existing documents. Do not add new normalizations here:
    // a legacy hash cannot be upgraded without the captured equation it describes.
    internal static string ComputeLegacyOmmlFingerprint(string wordOpenXml)
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

    internal static string ComputeVerifiedMaterializedOmmlFingerprint(
        string preparedWordOpenXml,
        string materializedWordOpenXml)
    {
        if (!string.Equals(ComputeImportedOmmlContentSignature(preparedWordOpenXml),
                ComputeImportedOmmlContentSignature(materializedWordOpenXml), StringComparison.Ordinal))
        {
            // Optional pure evidence: preserve both sides before the caller's
            // native Undo. Never relax content equality or let diagnostics block recovery.
            if (Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_RECOVERY_XML") == "1")
            {
                try
                {
                    var trace = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
                    if (!string.IsNullOrWhiteSpace(trace))
                    {
                        var path = trace + ".omml-mismatch-" + Guid.NewGuid().ToString("N");
                        File.WriteAllText(path + "-prepared.xml", preparedWordOpenXml);
                        File.WriteAllText(path + "-actual.xml", materializedWordOpenXml);
                        WordDoubleClickHook.TraceMessage("omml-materialization-mismatch evidence=" + path);
                    }
                }
                catch { }
            }
            throw new InvalidDataException("The materialized Word equation differs from its prepared formula content.");
        }
        // Word's final native representation is the durable identity. The
        // prepared XML is only evidence of the intended content, never a proxy
        // for what Word actually stored after importing and laying out the row.
        return ComputeOmmlFingerprint(materializedWordOpenXml);
    }

    // Shared content model for first-import verification and v2 persistence.
    // Changes that alter persisted v2 digests need a version/reader review;
    // legacy identity remains isolated in ComputeLegacyOmmlFingerprint.
    // Namespace spelling, empty property containers, on/off
    // lexical forms and Word's default integral glyph are equivalent encodings,
    // not a reason to bind an identity to different mathematical content.
    internal static string ComputeImportedOmmlContentSignature(string wordOpenXml)
    {
        var document = XDocument.Parse(StripVisualTeXNativeEquationNumber(wordOpenXml), LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;
        document.Descendants(word + "rPr").Remove();
        document.Descendants(math + "ctrlPr").Remove();
        document.Descendants(word + "bookmarkStart").Remove();
        document.Descendants(word + "bookmarkEnd").Remove();
        NormalizeImportedMathOnOffProperties(document, math);
        // The Office MathML stylesheet emits a legacy scrLvl=0 hint for
        // display-style fraction arguments (e.g. \\dfrac). Word 2021 drops this
        // hint when importing the same numerator/denominator. Match only this
        // observed neutral encoding; preserve nonzero hints, valid argSz
        // properties, and every operand/structural node for strict comparison.
        foreach (var hint in document.Descendants(math + "scrLvl").Where(element =>
                     element.Parent?.Name == math + "argPr"
                     && (string?)element.Attribute(math + "val") == "0"
                     && !element.HasElements
                     && string.IsNullOrWhiteSpace(element.Value)
                     && element.Attributes().All(attribute =>
                         attribute.IsNamespaceDeclaration || attribute.Name == math + "val")).ToArray())
            hint.Remove();
        foreach (var nary in document.Descendants(math + "nary"))
        {
            var properties = nary.Element(math + "naryPr");
            if (properties is null) { properties = new XElement(math + "naryPr"); nary.AddFirst(properties); }
            if (properties.Element(math + "chr") is null)
                properties.AddFirst(new XElement(math + "chr", new XAttribute(math + "val", "∫")));
        }
        // ISO/IEC 29500 m:fPr/m:type defaults to bar both when the element
        // is absent and when val is absent. Word omits this default on import.
        // Keep every nondefault type: a stack or skewed fraction is not bar.
        foreach (var fraction in document.Descendants(math + "f"))
        {
            var properties = fraction.Element(math + "fPr");
            if (properties is null) { properties = new XElement(math + "fPr"); fraction.AddFirst(properties); }
            var type = properties.Element(math + "type");
            if (type is null) properties.AddFirst(new XElement(math + "type", new XAttribute(math + "val", "bar")));
            else if (type.Attribute(math + "val") is null) type.SetAttributeValue(math + "val", "bar");
        }
        // ISO/IEC 29500 m:accPr/m:chr defaults to the combining circumflex.
        // Word commonly removes an explicit U+0302 after importing \hat{x}; the
        // omission and explicit value are the same mathematical accent. Preserve
        // every nondefault accent character.
        foreach (var accent in document.Descendants(math + "acc"))
        {
            var properties = accent.Element(math + "accPr");
            if (properties is null)
            {
                properties = new XElement(math + "accPr");
                accent.AddFirst(properties);
            }
            var character = properties.Element(math + "chr");
            if (character is null)
                properties.AddFirst(new XElement(
                    math + "chr",
                    new XAttribute(math + "val", "\u0302")));
            else if (character.Attribute(math + "val") is null)
                character.SetAttributeValue(math + "val", "\u0302");
        }
        // ISO/IEC 29500 m:mPr/m:baseJc defaults to center. Word removes the
        // explicit center during matrix import; keep nondefault top/bot intact.
        // This is a representational equivalence, not permission to change a
        // matrix's cells, row order, alignment or any mathematical contents.
        foreach (var matrix in document.Descendants(math + "m"))
        {
            var justification = matrix.Element(math + "mPr")?.Element(math + "baseJc");
            if (justification is not null
                && ((string?)justification.Attribute(math + "val") is null or "center"))
                justification.Remove();
        }
        NormalizeImportedMatrixColumnGroups(document, math);
        // A separator is rendered only between distinct delimiter arguments.
        // A single argument (including a multirow equation array) has none;
        // Word drops sepChr there. Beginning/end delimiters are always retained.
        foreach (var delimiter in document.Descendants(math + "d"))
            if (delimiter.Elements(math + "e").Count() == 1)
                delimiter.Element(math + "dPr")?.Elements(math + "sepChr").Remove();
        NormalizeLegacyMathAlphabetSpelling(document, math);
        NormalizeMathScriptRunBoundaries(document, math);
        foreach (var text in document.Descendants(math + "t"))
        {
            var properties = text.Parent?.Element(math + "rPr");
            var literal = properties?.Elements().Any(e => (e.Name == math + "nor" || e.Name == math + "lit")
                && (string?)e.Attribute(math + "val") == "1") == true;
            if (!literal) text.Value = NormalizeImportedMathTextSpelling(text.Value);
        }
        NormalizeImportedDefaultProperties(document, math);
        NormalizeImportedPlainStyleDefaults(document, math);
        // The MathML stylesheet can leave a zero-length normal-text run between
        // operands. Word removes that empty run and joins its neighbours. Remove
        // only truly empty text/format-only runs in the import comparison: a real
        // space, alignment/break control, object, or empty argument container must
        // remain observable in both import verification and v2 identity.
        foreach (var run in document.Descendants(math + "r").Where(run =>
                     !run.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)
                     && run.Elements(math + "t").Any()
                     && run.Elements(math + "t").All(text => !text.HasElements && text.Value.Length == 0)
                     && run.Elements().All(child => child.Name == math + "t"
                         || (child.Name == math + "rPr"
                             && !child.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)
                             && child.Elements().All(property => property.Name == math + "nor"
                                 || property.Name == math + "lit" || property.Name == math + "sty"
                                 || property.Name == math + "scr")))).ToArray())
            run.Remove();
        foreach (var property in document.Descendants().Where(e => e.Name.Namespace == math
                     && e.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal)
                     && !e.HasElements && !e.Attributes().Any(a => !a.IsNamespaceDeclaration)
                     && string.IsNullOrWhiteSpace(e.Value)).ToArray())
            property.Remove();
        NormalizeImportedMathRunGrouping(document, math);
        var canonical = SerializeCanonicalOmmlElement(
            document.Root ?? throw new InvalidDataException("The imported OMML has no equation root."));
        using var hash = SHA256.Create();
        return string.Concat(hash.ComputeHash(Encoding.UTF8.GetBytes(canonical))
            .Select(value => value.ToString("x2")));
    }

    private static string SerializeCanonicalOmmlElement(XElement root)
    {
        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;
        var canonical = new StringBuilder();
        void AppendToken(string token) { canonical.Append(token.Length).Append(':').Append(token); }
        void AppendElement(XElement element)
        {
            canonical.Append('('); AppendToken(element.Name.ToString());
            foreach (var attribute in element.Attributes().Where(a => !a.IsNamespaceDeclaration)
                         .OrderBy(a => a.Name.ToString(), StringComparer.Ordinal))
            { canonical.Append('@'); AppendToken(attribute.Name.ToString()); AppendToken(attribute.Value); }
            foreach (var node in element.Nodes())
            {
                if (node is XElement child) AppendElement(child);
                else if (node is XText text && (element.Name == math + "t"
                    || element.Name == word + "t" || element.Name == word + "instrText"
                    || element.Name == word + "delText" || element.Name == word + "delInstrText"
                    || !string.IsNullOrWhiteSpace(text.Value)))
                { canonical.Append('#'); AppendToken(text.Value); }
            }
            canonical.Append(')');
        }
        AppendElement(root);
        return canonical.ToString();
    }

    private static void NormalizeImportedMathRunGrouping(XDocument document, XNamespace math)
    {
        // This is intentionally separate from the frozen v1 grouping routine.
        // m:r can contain Word controls, field code and annotated text. Only
        // plain, unannotated mathematical text with known style properties can
        // be concatenated without erasing a control or an extension payload.
        bool PlainRun(XElement run)
        {
            if (run.Name != math + "r" || run.Attributes().Any(a => !a.IsNamespaceDeclaration)
                || run.Nodes().OfType<XText>().Any(t => !string.IsNullOrWhiteSpace(t.Value))
                || run.Elements(math + "rPr").Count() > 1) return false;
            var texts = run.Elements(math + "t").ToArray();
            if (texts.Length == 0 || texts.Any(t => t.HasElements
                || t.Attributes().Any(a => !a.IsNamespaceDeclaration))) return false;
            if (run.Elements().Any(e => e.Name != math + "t" && e.Name != math + "rPr")) return false;
            var properties = run.Element(math + "rPr");
            return properties is null || (!properties.Attributes().Any(a => !a.IsNamespaceDeclaration)
                && !properties.Nodes().OfType<XText>().Any(t => !string.IsNullOrWhiteSpace(t.Value))
                && properties.Elements().All(p =>
                    (p.Name == math + "sty" || p.Name == math + "scr"
                        || p.Name == math + "nor" || p.Name == math + "lit")
                    && !p.HasElements && string.IsNullOrWhiteSpace(p.Value)
                    && p.Attributes().All(a => a.IsNamespaceDeclaration || a.Name == math + "val")));
        }

        foreach (var run in document.Descendants(math + "r").Where(PlainRun).ToArray())
        {
            var texts = run.Elements(math + "t").ToArray();
            texts[0].Value = string.Concat(texts.Select(t => t.Value));
            foreach (var extra in texts.Skip(1)) extra.Remove();
        }
        foreach (var parent in document.Root?.DescendantsAndSelf().ToArray() ?? Array.Empty<XElement>())
        {
            XElement? previous = null;
            string? previousKey = null;
            foreach (var run in parent.Elements().ToArray())
            {
                if (!PlainRun(run)) { previous = null; previousKey = null; continue; }
                var properties = run.Element(math + "rPr");
                var key = properties is null ? string.Empty : SerializeCanonicalOmmlElement(properties);
                if (previous is not null && string.Equals(previousKey, key, StringComparison.Ordinal))
                {
                    previous.Element(math + "t")!.Value += run.Element(math + "t")!.Value;
                    run.Remove();
                }
                else { previous = run; previousKey = key; }
            }
        }
    }

    private static void NormalizeImportedMatrixColumnGroups(
        XDocument document,
        XNamespace math)
    {
        foreach (var matrix in document.Descendants(math + "m"))
        {
            var columns = matrix.Element(math + "mPr")?.Element(math + "mcs");
            if (columns is null) continue;
            var groups = columns.Elements(math + "mc").ToArray();
            if (groups.Length == 0) continue;

            var expanded = new List<XElement>();
            var safe = true;
            foreach (var group in groups)
            {
                if (group.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)
                    || group.Elements().Count() != 1
                    || group.Element(math + "mcPr") is not XElement properties
                    || group.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
                {
                    safe = false;
                    break;
                }

                if (properties.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)
                    || properties.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
                {
                    safe = false;
                    break;
                }

                var propertyElements = properties.Elements().ToArray();
                if (propertyElements.Length != 2
                    || propertyElements.Count(element => element.Name == math + "count") != 1
                    || propertyElements.Count(element => element.Name == math + "mcJc") != 1)
                {
                    safe = false;
                    break;
                }

                var count = properties.Element(math + "count");
                var justification = properties.Element(math + "mcJc");
                var countValue = (string?)count?.Attribute(math + "val");
                var justificationValue = (string?)justification?.Attribute(math + "val");
                if (count is null
                    || justification is null
                    || count.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != math + "val")
                    || justification.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != math + "val")
                    || count.HasElements
                    || justification.HasElements
                    || !string.IsNullOrWhiteSpace(count.Value)
                    || !string.IsNullOrWhiteSpace(justification.Value)
                    || !int.TryParse(countValue, NumberStyles.None, CultureInfo.InvariantCulture, out var repetitions)
                    || repetitions <= 0
                    || repetitions > 256
                    || string.IsNullOrWhiteSpace(justificationValue))
                {
                    safe = false;
                    break;
                }

                for (var index = 0; index < repetitions; index++)
                {
                    expanded.Add(
                        new XElement(
                            math + "mc",
                            new XElement(
                                math + "mcPr",
                                new XElement(
                                    math + "count",
                                    new XAttribute(math + "val", "1")),
                                new XElement(
                                    math + "mcJc",
                                    new XAttribute(math + "val", justificationValue)))));
                }
            }

            if (safe)
                columns.ReplaceNodes(expanded);
        }
    }

    private static void NormalizeImportedMathOnOffProperties(
        XDocument document,
        XNamespace math)
    {
        // OfficeMath CT_OnOff uses two independent defaults:
        //   1) when an element is present without m:val, that val means true;
        //   2) when the element itself is absent, the property's semantic default
        //      depends on the property (and for m:grow, on its parent object).
        // Word routinely removes properties that equal their semantic default while
        // materializing OMML. Canonicalize those representation-only omissions, but
        // keep every nondefault value so a real mathematical/layout change remains
        // a different content signature.
        var onOffNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "degHide", "subHide", "supHide", "grow", "nor", "lit", "aln", "diff", "noBreak",
            "opEmu", "transp", "zeroAsc", "zeroDesc", "zeroWid", "show", "plcHide",
            "hideTop", "hideBot", "hideLeft", "hideRight", "strikeH", "strikeV",
            "strikeBLTR", "strikeTLBR",
        };
        var defaultFalseNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "degHide", "subHide", "supHide", "nor", "lit", "aln", "diff", "noBreak",
            "opEmu", "transp", "zeroAsc", "zeroDesc", "zeroWid", "plcHide",
            "hideTop", "hideBot", "hideLeft", "hideRight", "strikeH", "strikeV",
            "strikeBLTR", "strikeTLBR",
        };

        foreach (var property in document.Descendants()
                     .Where(element => element.Name.Namespace == math
                         && onOffNames.Contains(element.Name.LocalName))
                     .ToArray())
        {
            var normalizedValue = ReadMathBooleanValue(property);
            property.SetAttributeValue(math + "val", normalizedValue ? "1" : "0");

            bool? semanticDefault = null;
            if (defaultFalseNames.Contains(property.Name.LocalName))
            {
                semanticDefault = false;
            }
            else if (property.Name == math + "show")
            {
                // Phantom content is shown when m:show is omitted.
                semanticDefault = true;
            }
            else if (property.Name == math + "grow")
            {
                // ISO/IEC 29500 gives m:grow a context-sensitive default:
                // delimiter objects grow by default, n-ary operators do not.
                if (property.Parent?.Name == math + "dPr")
                    semanticDefault = true;
                else if (property.Parent?.Name == math + "naryPr")
                    semanticDefault = false;
            }

            if (!semanticDefault.HasValue) continue;
            if (normalizedValue == semanticDefault.Value)
                property.Remove();
        }
    }

    private static void NormalizeImportedDefaultProperties(XDocument document, XNamespace math)
    {
        void NormalizeScalar(XElement? property, string elementDefault, string attributeDefault)
        {
            if (property is null || property.HasElements || !string.IsNullOrWhiteSpace(property.Value)
                || property.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != math + "val"))
                return;
            var value = (string?)property.Attribute(math + "val") ?? attributeDefault;
            if (value == elementDefault) property.Remove();
            else property.SetAttributeValue(math + "val", value);
        }
        // ISO 29500 distinguishes omission of the element from omission of val:
        // begChr/endChr absent = parentheses; present without val = NO delimiter.
        // sty has the same italic default in both cases. Preserve every explicit
        // nondefault value and unknown extension attribute.
        foreach (var delimiter in document.Descendants(math + "d"))
        {
            var properties = delimiter.Element(math + "dPr");
            NormalizeScalar(properties?.Element(math + "begChr"), "(", string.Empty);
            NormalizeScalar(properties?.Element(math + "endChr"), ")", string.Empty);
        }
        foreach (var run in document.Descendants(math + "r"))
            NormalizeScalar(run.Element(math + "rPr")?.Element(math + "sty"), "i", "i");
        foreach (var text in document.Descendants(math + "t"))
        {
            var space = text.Attribute(XNamespace.Xml + "space");
            // A preservation annotation on unpadded text changes no characters.
            // Keep real leading/trailing whitespace and every character untouched.
            if (space?.Value is "preserve" or "default"
                && text.Value.Length > 0
                && text.Value.Trim(' ', '\t', '\r', '\n') == text.Value)
                space.Remove();
        }
    }

    private static void NormalizeImportedPlainStyleDefaults(
        XDocument document,
        XNamespace math)
    {
        foreach (var run in document.Descendants(math + "r").ToArray())
        {
            var properties = run.Element(math + "rPr");
            var style = properties?.Element(math + "sty");
            if (style is null
                || !string.Equals(
                    (string?)style.Attribute(math + "val"),
                    "p",
                    StringComparison.Ordinal))
                continue;

            var text = string.Concat(run.Elements(math + "t").Select(item => item.Value));
            if (string.IsNullOrEmpty(text) || ContainsUnicodeLetter(text))
                continue;

            // Word commonly splits a mixed math run such as "E=m" into
            // E / = / m and adds an explicit plain style to the operator. It also
            // writes m:sty="p" on ordinary digits such as a superscript 2. For
            // characters with no Unicode letter semantics this is only an explicit
            // spelling of their normal upright math appearance, not a change of
            // formula content. Keep plain styling on any run containing a letter:
            // \mathrm{x}, upright Greek, script alphabets, etc. remain semantic.
            style.Remove();
            if (properties is not null
                && !properties.HasElements
                && !properties.HasAttributes)
                properties.Remove();
        }
    }

    private static bool ContainsUnicodeLetter(string value)
    {
        for (var index = 0; index < value.Length;)
        {
            var category = CharUnicodeInfo.GetUnicodeCategory(value, index);
            if (category is UnicodeCategory.UppercaseLetter
                or UnicodeCategory.LowercaseLetter
                or UnicodeCategory.TitlecaseLetter
                or UnicodeCategory.ModifierLetter
                or UnicodeCategory.OtherLetter
                or UnicodeCategory.LetterNumber)
                return true;
            index += char.IsSurrogatePair(value, index) ? 2 : 1;
        }
        return false;
    }

    // Frozen v1 serialization behavior. New content comparisons must use the
    // control-preserving NormalizeImportedMathRunGrouping instead.
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
        sequenceName = (sequenceName ?? string.Empty).Trim();
        if (sequenceName.Length == 0
            || sequenceName.Length > 80
            || sequenceName.Any(character =>
                character is '\r' or '\n' or '\u0013' or '\u0014' or '\u0015' or '"'))
            throw new ArgumentException(
                "The Word SEQ identifier is invalid.",
                nameof(sequenceName));
        var sequenceToken = Regex.IsMatch(
                sequenceName,
                @"^[A-Za-z][A-Za-z0-9_]{0,39}$",
                RegexOptions.CultureInvariant)
            ? sequenceName
            : """ + sequenceName + """;
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

    internal static string BuildWordNativeNumberedOmml(
        string semanticOmml,
        string sequenceName,
        string? numberBookmarkName,
        int restartHeadingLevel,
        string headingSeparator,
        string initialSequenceResult = "1")
    {
        if (string.IsNullOrWhiteSpace(semanticOmml))
            throw new ArgumentException(
                "The semantic OMML payload must not be empty.",
                nameof(semanticOmml));
        if (string.IsNullOrWhiteSpace(sequenceName)
            || sequenceName.Length > 80
            || sequenceName.Any(character =>
                character is '\r' or '\n' or '\u0013' or '\u0014' or '\u0015'))
            throw new ArgumentException(
                "The Word SEQ identifier is invalid.",
                nameof(sequenceName));
        if (sequenceName.IndexOf('"') >= 0
            || sequenceName.IndexOf('\\') >= 0)
            throw new ArgumentException(
                "The Word SEQ identifier contains field-syntax characters.",
                nameof(sequenceName));
        var sequenceToken =
            sequenceName.Any(char.IsWhiteSpace)
                ? """ + sequenceName + """
                : sequenceName;
        if (restartHeadingLevel < 0 || restartHeadingLevel > 9)
            throw new ArgumentOutOfRangeException(
                nameof(restartHeadingLevel));
        if (string.IsNullOrWhiteSpace(initialSequenceResult)
            || initialSequenceResult.Any(character =>
                character is '\r' or '\n' or '\u0013' or '\u0014' or '\u0015'))
            throw new ArgumentException(
                "The initial SEQ result is invalid.",
                nameof(initialSequenceResult));

        if (!string.IsNullOrWhiteSpace(numberBookmarkName))
        {
            _ = ValidateManagedEquationBookmarkName(
                numberBookmarkName,
                VisualTeXNativeNumberBookmarkPrefix,
                nameof(numberBookmarkName));
        }

        headingSeparator ??= string.Empty;
        if (headingSeparator.Any(character =>
                character is '\r' or '\n' or '\u0013' or '\u0014' or '\u0015'))
            throw new ArgumentException(
                "The heading separator contains an invalid Word field-control character.",
                nameof(headingSeparator));

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
                "A numbered Word equation requires a nonempty mathematical body.");

        var representativeRunProperties = equation
            .Descendants(word + "rPr")
            .FirstOrDefault(properties => properties.Element(word + "sz") is not null)
            ?? equation.Descendants(word + "rPr").FirstOrDefault();

        XElement RunProperties(bool noProof = false, bool italic = false)
        {
            var properties = new XElement(word + "rPr");
            foreach (var name in new[]
                     {
                         word + "rFonts",
                         word + "sz",
                         word + "szCs",
                     })
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

        XElement ControlProperties() =>
            new(
                math + "ctrlPr",
                RunProperties(italic: true));

        XElement BoundaryRun(XElement content) =>
            new(
                math + "r",
                RunProperties(italic: true),
                content);

        XElement InstructionRun(string instruction) =>
            new(
                math + "r",
                RunProperties(),
                new XElement(
                    word + "instrText",
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    instruction));

        XElement ResultRun(string result) =>
            new(
                math + "r",
                RunProperties(noProof: true),
                new XElement(math + "t", result));

        XElement TextRun(string text) =>
            new(
                math + "r",
                RunProperties(),
                new XElement(math + "t", text));

        IEnumerable<object> NativeField(string instruction, string result)
        {
            yield return BoundaryRun(
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "begin")));
            yield return InstructionRun(instruction);
            yield return BoundaryRun(
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "separate")));
            yield return ResultRun(result);
            yield return BoundaryRun(
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "end")));
        }

        var numberNodes = new List<object>();
        if (!string.IsNullOrWhiteSpace(numberBookmarkName))
        {
            numberNodes.Add(
                new XElement(
                    word + "bookmarkStart",
                    new XAttribute(word + "id", "31901"),
                    new XAttribute(word + "name", numberBookmarkName)));
        }

        if (restartHeadingLevel > 0)
        {
            // Keep the chapter/section component native as well. STYLEREF reads
            // the nearest numbered Heading level; SEQ \s resets the equation
            // counter at that same level. VisualTeX does not cache either value.
            numberNodes.AddRange(
                NativeField(
                    $" STYLEREF {restartHeadingLevel} \\s \\* MERGEFORMAT ",
                    "1"));
            if (!string.IsNullOrEmpty(headingSeparator))
                numberNodes.Add(TextRun(headingSeparator));
        }

        var sequenceSwitch = restartHeadingLevel > 0
            ? $" \\s {restartHeadingLevel}"
            : string.Empty;
        numberNodes.AddRange(
            NativeField(
                $" SEQ {sequenceName}{sequenceSwitch} \\* ARABIC \\* MERGEFORMAT ",
                initialSequenceResult));
        if (!string.IsNullOrWhiteSpace(numberBookmarkName))
        {
            numberNodes.Add(
                new XElement(
                    word + "bookmarkEnd",
                    new XAttribute(word + "id", "31901")));
        }

        var delimiter = new XElement(
            math + "d",
            new XElement(
                math + "dPr",
                ControlProperties()),
            new XElement(math + "e", numberNodes));

        var equationBody = new XElement(
            math + "e",
            formulaNodes);
        equationBody.Add(
            TextRun("#"),
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
                        ControlProperties()),
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

    internal static string CombineInlineOmml(
        params string[] equations)
    {
        if (equations is null || equations.Length == 0)
            throw new ArgumentException(
                "At least one OMML equation is required.",
                nameof(equations));

        XNamespace math = MathNamespace;
        var combined = new XElement(math + "oMath");
        foreach (var source in equations)
        {
            var equation = XElement.Parse(
                ExtractSingleOMath(source),
                LoadOptions.PreserveWhitespace);
            foreach (var child in equation.Elements())
                combined.Add(new XElement(child));
        }

        if (!combined.Elements().Any())
            throw new InvalidDataException(
                "The combined inline OMML equation has no mathematical content.");

        return combined.ToString(
            SaveOptions.DisableFormatting);
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
        // This API promises one complete equation, not the first equation in a
        // larger range. Counting also detects an illegally nested oMath root.
        var equations = document.Descendants(math + "oMath").Take(2).ToArray();
        if (equations.Length == 0)
            throw new InvalidDataException("Office MathML conversion did not produce an m:oMath node.");
        if (equations.Length != 1)
            throw new InvalidDataException("Expected exactly one OMML equation; the input contains multiple equations.");
        return equations[0].ToString(SaveOptions.DisableFormatting);
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
                IsVisualTeXDirectSequenceDelimiter(
                    candidateNumber,
                    formulaId: null);
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
                @"\bSEQ\s+(?:""[^""]+""|[^\s\\]+)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            return false;
        if (Regex.IsMatch(
                fieldCode,
                @"\bREF\s+VTEqNum_",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            return false;

        // Native Word equation numbering is structural. FormulaId/bookmark
        // ownership is intentionally irrelevant.
        return true;
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

    private static string CreateSingleEquationFlatOpc(
        BatchEntry entry,
        string mathFontName)
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
                    new XElement(
                        word + "bookmarkStart",
                        new XAttribute(word + "id", entry.BookmarkId),
                        new XAttribute(word + "name", entry.BookmarkName)),
                    equation,
                    new XElement(
                        word + "bookmarkEnd",
                        new XAttribute(word + "id", entry.BookmarkId))),
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

    internal static bool ShouldPreferInsertedMathCandidate(
        int distance,
        int span,
        int bestDistance,
        int bestSpan)
    {
        if (distance > 16) return false;
        if (distance < bestDistance) return true;
        if (distance > bestDistance) return false;
        return span > bestSpan;
    }

    private static int ResolveCurrentParagraphEnd(
        Document document,
        int position)
    {
        Range? content = null;
        Range? point = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            content =
                document.Content;
            var safePosition =
                Math.Max(
                    content.Start,
                    Math.Min(
                        position,
                        Math.Max(
                            content.Start,
                            content.End - 1)));
            point =
                document.Range(
                    safePosition,
                    safePosition);
            paragraphs =
                point.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "The rewritten OMML insertion point no longer belongs to exactly one Word paragraph.");

            paragraph =
                paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            return paragraphRange.End;
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(point);
            Release(content);
        }
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
                    // range. The physical equation nearest the insertion point must
                    // win first; only candidates at the same start distance use the
                    // larger span to prefer the top-level matrix over an inner row.
                    // Prioritizing span before distance can select the next, longer
                    // display equation in an adjacent paragraph.
                    if (range.End <= position || range.Start > preferredEnd + 1) continue;
                    var span = range.End - range.Start;
                    var distance = Math.Abs(range.Start - position);
                    if (!ShouldPreferInsertedMathCandidate(
                            distance,
                            span,
                            bestDistance,
                            bestSpan))
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
                    if (!ShouldPreferInsertedMathCandidate(
                            distance,
                            span,
                            bestDistance,
                            bestSpan))
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
