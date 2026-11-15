using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using Application = Microsoft.Office.Interop.Word.Application;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordOmmlConverter;
using static VisualTeX.WordVsto.WordOmmlDocuments;
using static VisualTeX.WordVsto.WordOmmlFingerprints;
using static VisualTeX.WordVsto.WordOmmlInsertion;

namespace VisualTeX.WordVsto;

/// <summary>
/// OMML for many formulas at once: batch and whole-document sources materialized through one package, and MathType
/// objects replaced in a whole document.
/// </summary>
internal static class WordOmmlBatchSources
{
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
                if (source is not null) { try { source.Close(WdSaveOptions.wdDoNotSaveChanges); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); } Release(source); }
                if (!_preparedGroups.ContainsKey(key)) { try { File.Delete(path); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); } }
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
                try { targetDocument.Activate(); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
                var wpsInlineSentinel = false;

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
                    wpsInlineSentinel = OfficeHostEnvironment.IsWps
                        && !display
                        && entry.InsertionWordOpenXml is not null;
                    WpsXmlInsertion.Insert(
                        target,
                        wpsInlineSentinel
                            ? AppendWpsInlineSentinel(insertionXml)
                            : insertionXml);
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
                if (wpsInlineSentinel)
                {
                    RemoveWpsInlineSentinel(insertedMath);
                    Release(insertedMath);
                    insertedMath = null;
                    insertedMath = FindMathAtPosition(
                            targetDocument,
                            insertionStart,
                            Math.Max(insertionStart, target.End))
                        ?? throw new InvalidOperationException(
                            "WPS lost the batch OMML equation after removing its import separator.");
                }
                var targetType = display
                    ? WdOMathType.wdOMathDisplay
                    : WdOMathType.wdOMathInline;
                // WPS keeps an equation alone in its paragraph display (see
                // ReplaceWithPreparedOmmlDirect).
                if (insertedMath.Type != targetType
                    && (display || !OfficeHostEnvironment.IsWps))
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
                try { document.Saved = true; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                try { document.Close(WdSaveOptions.wdDoNotSaveChanges); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                Release(document);
            }
            foreach (var group in _preparedGroups.Values)
            {
                try { group.Document.Close(WdSaveOptions.wdDoNotSaveChanges); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                Release(group.Document);
                try { File.Delete(group.Path); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            }
            _preparedGroups.Clear();
            try { File.Delete(_path); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
                try { targetDocument.Activate(); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
                try { document.Saved = true; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                try { document.Close(WdSaveOptions.wdDoNotSaveChanges); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                Release(document);
            }
            try { File.Delete(_path); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
                try { document.Close(WdSaveOptions.wdDoNotSaveChanges); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                Release(document);
                try { File.Delete(path); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
                try { scratch.Saved = true; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                try { scratch.Close(WdSaveOptions.wdDoNotSaveChanges); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                Release(scratch);
            }
            if (materialized is not null)
            {
                try { materialized.Close(WdSaveOptions.wdDoNotSaveChanges); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                Release(materialized);
            }
            if (!ownershipTransferred)
                try { File.Delete(path); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
                try { document.Close(WdSaveOptions.wdDoNotSaveChanges); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                Release(document);
                try { File.Delete(path); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
                try { isolatedDocument.Saved = true; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                try
                {
                    isolatedDocument.Close(
                        WdSaveOptions.wdDoNotSaveChanges);
                }
                catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            }
            Release(isolatedDocument);
            if (isolatedApplication is not null)
            {
                try
                {
                    isolatedApplication.Quit(
                        WdSaveOptions.wdDoNotSaveChanges);
                }
                catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
        var body = WordDocumentXml.GetRequiredMainDocumentBody(package);
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
            var fingerprint = ComputeOmmlFingerprint(actual);

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
}
