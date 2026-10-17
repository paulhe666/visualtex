using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// An ordinary native edit keeps its paragraph/numbering topology. All source
/// serialization and semantic-number composition are preflight; Apply owns only
/// the exact OMath substitution. Native numbering is not regenerated on an edit.
/// </summary>
internal sealed class WordNativeOmmlEdit : IDisposable
{
    private readonly Document _document;
    private readonly WordFormulaHostDescriptor _source;
    private readonly WordFormulaHostWriteRequest _request;
    private readonly string _omml;
    private readonly bool _refreshNumberInsideApply;
    private string? _markRunProperties;
    private MathTypeLocalEditScope? _recovery;
    private readonly List<(string Name, bool Full, bool AtEnd)> _boundaryBookmarks = new();
    private readonly List<(string Name, int Start, int End, int Length, string Text)> _numberBookmarks = new();
    private bool _applied;

    internal static bool Supports(WordFormulaHostDescriptor source, WordFormulaHostWriteRequest request) =>
        source.Kind == WordFormulaHostKind.Omml && request.Kind == WordFormulaHostKind.Omml
        && source.Numbering.Numbered == request.Numbered
        && string.Equals(source.DisplayMode, request.DisplayMode, StringComparison.OrdinalIgnoreCase);

    private WordNativeOmmlEdit(Document document, WordFormulaHostDescriptor source,
        WordFormulaHostWriteRequest request, string omml,
        bool refreshNumberInsideApply, MathTypeLocalEditScope recovery)
    {
        _document = document; _source = source; _request = request;
        _omml = omml;
        _refreshNumberInsideApply = refreshNumberInsideApply;
        _recovery = recovery;
    }

    internal static WordNativeOmmlEdit Prepare(Document document,
        WordFormulaHostDescriptor source, WordFormulaHostWriteRequest request)
    {
        if (!Supports(source, request))
            throw new InvalidOperationException("A native edit plan requires unchanged display/number topology.");
        return PrepareCore(
            document,
            source,
            request,
            preparedTopologyOmml: null,
            allowNumberAttachment: false);
    }

    /// <summary>
    /// Prepares the one supported topology-changing native edit: an existing
    /// unnumbered display OMath becomes a numbered display OMath. The complete
    /// #(SEQ/STYLEREF) target was already composed by the document-wide native
    /// numbering adoption plan. Keeping this in the same exact native edit
    /// transport as ordinary OMML edits is essential: Word's paragraph import
    /// path can terminate CustomUndoRecord, while direct exact OMath replacement
    /// is the already-proven one-step-Undo path.
    /// </summary>
    internal static WordNativeOmmlEdit PrepareNumberAttachment(
        Document document,
        WordFormulaHostDescriptor source,
        WordFormulaHostWriteRequest request,
        string preparedOmml)
    {
        if (source.Kind != WordFormulaHostKind.Omml
            || request.Kind != WordFormulaHostKind.Omml
            || source.Numbering.Numbered
            || !request.Numbered
            || !source.Display
            || !string.Equals(
                request.DisplayMode,
                "block",
                StringComparison.OrdinalIgnoreCase)
            || string.IsNullOrWhiteSpace(preparedOmml))
            throw new InvalidOperationException(
                "A native number-attachment edit requires one unnumbered display OMath and one prepared numbered display target.");

        return PrepareCore(
            document,
            source,
            request,
            preparedOmml,
            allowNumberAttachment: true);
    }

    private static WordNativeOmmlEdit PrepareCore(
        Document document,
        WordFormulaHostDescriptor source,
        WordFormulaHostWriteRequest request,
        string? preparedTopologyOmml,
        bool allowNumberAttachment)
    {
        Range? exact = null;
        Range? number = null;
        Fields? fields = null;
        Bookmarks? localBookmarks = null;
        Bookmarks? bookmarks = null;
        MathTypeLocalEditScope? recovery = null;
        WordNativeOmmlEdit? plan = null;
        try
        {
            exact = WordFormulaHostSemanticReader.CreateRange(document, source.Range);
            recovery = MathTypeLocalEditScope.Capture(document, exact);
            var xml = exact.WordOpenXML;
            var semantic = WordOmmlHostWriter.ApplySemanticOmmlTypography(
                WordOmmlConverter.TransformMathMlToOmml(request.MathMl!), request.FontSizePoints);
            var hasNumberHost = WordOmmlConverter.HasWordNativeEquationNumberHost(xml);
            fields = exact.Fields;
            var preserve = source.Display && hasNumberHost
                && (request.Numbered || fields.Count == 0);
            if (request.Numbered && !preserve && !allowNumberAttachment)
                throw new InvalidDataException("The captured native number topology no longer exists.");
            if (allowNumberAttachment && hasNumberHost)
                throw new InvalidDataException(
                    "The source acquired a native number before the prepared number-attachment edit ran.");

            var prepared = allowNumberAttachment
                ? preparedTopologyOmml
                    ?? throw new InvalidDataException(
                        "The native number-attachment edit lost its prepared target.")
                : preserve
                    ? WordOmmlConverter.ReplaceWordNativeEquationNumberHostBody(xml, semantic)
                    : semantic;
            plan = new WordNativeOmmlEdit(
                document,
                source,
                request,
                prepared,
                refreshNumberInsideApply: request.Numbered && !allowNumberAttachment,
                recovery);
            recovery = null;
            plan._markRunProperties = WordOmmlInsertionXml.ParagraphMarkRunProperties(plan._recovery!.ParagraphXml);

            // A clipped equation export can omit existing bookmarks at its outer
            // seam. Number bookmarks also need explicit ownership: under Track
            // Changes Word can otherwise retain their names on the deleted old
            // equation rather than on the inserted current number.
            if (preserve) number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
            var names = new HashSet<string>(MathTypeLocalFragment.BookmarkNames(plan._recovery!.ParagraphXml),
                StringComparer.Ordinal);
            localBookmarks = exact.Bookmarks;
            for (var index = 1; index <= localBookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                try { bookmark = localBookmarks[index]; names.Add(bookmark.Name); }
                finally { Release(bookmark); }
            }
            bookmarks = document.Bookmarks;
            foreach (var name in names)
            {
                Bookmark? bookmark = null;
                Range? span = null;
                try
                {
                    if (!bookmarks.Exists(name)) continue;
                    bookmark = bookmarks[name]; span = bookmark.Range;
                    if (!span.InStory(exact)) continue;
                    var full = span.Start == exact.Start && span.End == exact.End;
                    var emptyBoundary = span.Start == span.End
                        && (span.Start == exact.Start || span.Start == exact.End);
                    if (full || emptyBoundary)
                        plan._boundaryBookmarks.Add((name, full, span.Start == exact.End));
                    else if (number is not null && span.Start >= number.Start && span.End <= number.End)
                        plan._numberBookmarks.Add((name, span.Start - number.Start, span.End - number.Start,
                            number.End - number.Start, span.Text ?? string.Empty));
                }
                finally { Release(span); Release(bookmark); }
            }
            var result = plan; plan = null;
            return result;
        }
        finally
        {
            plan?.Dispose(); recovery?.Dispose();
            Release(bookmarks); Release(localBookmarks); Release(fields); Release(number); Release(exact);
        }
    }

    internal WordFormulaHostDescriptor Apply()
    {
        if (_applied || _recovery is null)
            throw new InvalidOperationException("The native edit plan is no longer available.");
        Range? exact = null;
        Range? replacement = null;
        OMaths? maths = null;
        OMath? native = null;
        Range? actual = null;
        try
        {
            exact = WordFormulaHostSemanticReader.CreateRange(_document, _source.Range);
            maths = exact.OMaths;
            if (maths.Count != 1) throw new InvalidDataException("The native edit lost its exact source OMath.");
            native = maths[1]; actual = native.Range;
            if (actual.Start != exact.Start || actual.End != exact.End || !actual.InStory(exact))
                throw new InvalidDataException("The native edit source no longer has its captured boundary.");
            _applied = true;
            BindBoundaryBookmarks(exact, detach: true);
            BindNumberBookmarks(exact, detach: true);
            var written = WordOmmlHostWriter.InsertPreparedNative(
                _document, exact, _request, _omml,
                paragraphMarkRunProperties: _markRunProperties).Host;
            replacement = WordFormulaHostSemanticReader.CreateRange(_document, written.Range);
            BindBoundaryBookmarks(replacement, detach: false);
            BindNumberBookmarks(replacement, detach: false);
            if (_refreshNumberInsideApply)
                written = WordNativeOmmlNumbering.RefreshCanonicalNativeFields(_document, written);
            return WordFormulaMutationValidator.ValidateInsertedHost(_document, _request, written);
        }
        finally { Release(actual); Release(native); Release(maths); Release(replacement); Release(exact); }
    }

    private void BindBoundaryBookmarks(Range exact, bool detach)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = _document.Bookmarks;
            foreach (var item in _boundaryBookmarks)
            {
                Bookmark? bookmark = null;
                Range? span = null;
                try
                {
                    if (detach)
                    {
                        if (!bookmarks.Exists(item.Name))
                            throw new InvalidDataException("A preflight native boundary bookmark disappeared.");
                        bookmark = bookmarks[item.Name]; bookmark.Delete();
                    }
                    else
                    {
                        span = exact.Duplicate;
                        if (!item.Full) span.Collapse(item.AtEnd
                            ? WdCollapseDirection.wdCollapseEnd : WdCollapseDirection.wdCollapseStart);
                        bookmark = bookmarks.Add(item.Name, span);
                    }
                }
                finally { Release(span); Release(bookmark); }
            }
        }
        finally { Release(bookmarks); }
    }

    private void BindNumberBookmarks(Range exact, bool detach)
    {
        if (_numberBookmarks.Count == 0) return;
        Bookmarks? bookmarks = null;
        Range? number = null;
        try
        {
            bookmarks = _document.Bookmarks;
            if (!detach) number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
            foreach (var item in _numberBookmarks)
            {
                Bookmark? bookmark = null;
                Range? span = null;
                try
                {
                    if (detach)
                    {
                        if (!bookmarks.Exists(item.Name))
                            throw new InvalidDataException("A captured native number bookmark disappeared before editing.");
                        bookmark = bookmarks[item.Name]; bookmark.Delete();
                    }
                    else
                    {
                        var length = number!.End - number.Start;
                        var full = item.Start == 0 && item.End == item.Length;
                        var seam = item.Start == item.End && (item.Start == 0 || item.Start == item.Length);
                        if (!full && !seam && length != item.Length)
                            throw new InvalidDataException("An internal native number bookmark cannot be mapped to a changed number span.");
                        span = number.Duplicate;
                        if (seam) span.Collapse(item.Start == 0 ? WdCollapseDirection.wdCollapseStart : WdCollapseDirection.wdCollapseEnd);
                        else if (!full) span.SetRange(number.Start + item.Start, number.Start + item.End);
                        if (!string.Equals(span.Text ?? string.Empty, item.Text, StringComparison.Ordinal))
                            throw new InvalidDataException("A preserved native number bookmark changed its owned content during editing.");
                        bookmark = bookmarks.Add(item.Name, span);
                    }
                }
                finally { Release(span); Release(bookmark); }
            }
        }
        finally { Release(number); Release(bookmarks); }
    }

    internal void VerifyRollback() => _recovery!.VerifyRollback();
    public void Dispose() { _recovery?.Dispose(); _recovery = null; }
    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
