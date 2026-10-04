using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// A representation change replaces one physical OLE, not a live paragraph's
/// current/deleted history. All WordOpenXML preparation precedes the transaction:
/// that Word getter can close its application-wide Custom UndoRecord.
/// </summary>
internal static class WordOleToOmmlReplacement
{
    internal sealed class Prepared : IDisposable
    {
        private readonly Document _document;
        private Range? _source;
        private MathTypeLocalEditScope? _owner;
        private readonly string _fragment;
        private readonly bool _display;
        private readonly string? _retiredIdentityBookmark;
        private WordMathTypeHostAdapter.NumberedScaffold? _numberScaffold;
        private bool _applied;

        internal Prepared(Document document, Range source, MathTypeLocalEditScope owner,
            string fragment, bool display, string? retiredIdentityBookmark,
            WordMathTypeHostAdapter.NumberedScaffold? numberScaffold)
        {
            _document = document;
            _source = source;
            _owner = owner;
            _fragment = fragment;
            _display = display;
            _retiredIdentityBookmark = retiredIdentityBookmark;
            _numberScaffold = numberScaffold;
        }

        internal Range Apply(Range selected)
        {
            if (_applied || _owner is null || _source is null)
                throw new InvalidOperationException("The exact OLE replacement plan is no longer available.");
            if (!selected.InStory(_source) || selected.Start != _source.Start || selected.End != _source.End)
                throw new InvalidOperationException("The exact OLE source moved outside its prepared replacement plan.");

            InlineShapes? sources = null;
            InlineShape? source = null;
            Range? exact = null;
            Bookmarks? bookmarks = null;
            Bookmark? locator = null;
            Range? component = null;
            OMaths? maths = null;
            OMath? math = null;
            var trackingPaused = false;
            try
            {
                WordFormulaMutationTransaction.TraceBoundary(_document.Application, "exact-ole-apply-enter");
                sources = selected.InlineShapes;
                if (sources.Count != 1)
                    throw new InvalidDataException("The prepared replacement no longer owns one source OLE.");
                source = sources[1];
                exact = source.Range;
                if (source.Type is not (WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                        or WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                    || exact.Start != selected.Start || exact.End != selected.End)
                    throw new InvalidDataException("The prepared OLE source no longer has its exact physical boundary.");

                // Earlier replacements in this same paragraph can change its
                // right-hand text and remaining OLE count. Capture those LIVE
                // witnesses now, without exporting XML or rereading native OLE.
                if (_document.TrackRevisions)
                {
                    _document.TrackRevisions = false;
                    trackingPaused = true;
                }
                _applied = true;
                WordFormulaMutationTransaction.TraceBoundary(_document.Application, "exact-ole-before-scaffold");
                _numberScaffold?.Remove();
                WordFormulaMutationTransaction.TraceBoundary(_document.Application, "exact-ole-after-scaffold");
                _owner.CaptureReplacementBoundary(selected);
                WordFormulaMutationTransaction.TraceBoundary(_document.Application, "exact-ole-before-import");
                selected.InsertXML(_fragment);
                WordFormulaMutationTransaction.TraceBoundary(_document.Application, "exact-ole-after-import");
                bookmarks = _document.Bookmarks;
                if (!bookmarks.Exists(_owner.LocatorName))
                    throw new InvalidOperationException("Word lost the exact OLE replacement boundary.");
                locator = bookmarks[_owner.LocatorName];
                component = locator.Range;
                maths = component.OMaths;
                if (maths.Count != 1)
                    throw new InvalidOperationException("The OLE replacement did not materialize exactly one native OMath.");
                math = maths[1];
                // As in ordinary native OMML insertion, a lone paragraph/cell
                // formula may be promoted to display by Word. Forcing it back
                // to inline makes Word add a typing space and another paragraph.
                // Explicit display is authoritative; native inline promotion is
                // not a reason to manufacture synthetic surrounding content.
                if (_display && math.Type != WdOMathType.wdOMathDisplay)
                    math.Type = WdOMathType.wdOMathDisplay;
                // InsertXML already materialized the prepared OMML tree.
                // BuildUp reparses literal runs and can replace fixed/styled
                // delimiters with different structures; preserve the input tree.
                Release(component); component = locator.Range;
                _owner.VerifyReplacementBoundary(component);
                WordFormulaMutationTransaction.TraceBoundary(_document.Application, "exact-ole-after-boundary-check");
                _owner.RestoreContainedBookmarks(component, _retiredIdentityBookmark);
                _numberScaffold?.RestoreNumberBookmarks(_document, component);
                _owner.RemoveLocator();
                WordFormulaMutationTransaction.TraceBoundary(_document.Application, "exact-ole-apply-complete");
                return math.Range.Duplicate;
            }
            finally
            {
                // The prepared fragment carries its source revision envelope.
                // This temporary setting prevents tracking the serialization as
                // a second delete/insert, and is restored before native rollback.
                try { if (trackingPaused) _document.TrackRevisions = true; }
                finally
                {
                    Release(math); Release(maths); Release(component); Release(locator);
                    Release(bookmarks); Release(exact); Release(source); Release(sources);
                }
            }
        }

        internal void VerifyRollback() => _owner?.VerifyRollback();

        public void Dispose()
        {
            _numberScaffold?.Dispose(); _numberScaffold = null;
            _owner?.Dispose(); _owner = null;
            Release(_source); _source = null;
        }
    }

    internal static Prepared? Prepare(Document document, Range selected, string omml, bool display,
        string? retiredIdentityBookmark = null, bool replaceMathTypeNumber = false)
    {
        if (selected.Start == selected.End) return null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? source = null;
        MathTypeLocalEditScope? owner = null;
        WordMathTypeHostAdapter.NumberedScaffold? numberScaffold = null;
        try
        {
            shapes = selected.InlineShapes;
            if (shapes.Count != 1) return null;
            shape = shapes[1];
            if (shape.Type is not (WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                or WdInlineShapeType.wdInlineShapeLinkedOLEObject)) return null;
            source = shape.Range.Duplicate;
            if (source.Start != selected.Start || source.End != selected.End
                || !source.InStory(selected)) return null;

            using var metric = WordOperationMetrics.Measure("OmmlExactOleReplacement.Prepare");
            owner = MathTypeLocalEditScope.Capture(document, source);
            if (replaceMathTypeNumber)
                numberScaffold = WordMathTypeHostAdapter.PrepareNumberedScaffold(document, source, owner.ParagraphXml);
            var fragment = MathTypeLocalFragment.PrepareOmmlReplacement(
                source.WordOpenXML, owner.ParagraphXml, owner.LocatorName,
                WordOmmlConverter.ExtractSingleOMath(omml), owner.SourceOleOrdinal, owner.SourceOleCount,
                retiredIdentityBookmark);
            var prepared = new Prepared(document, source, owner, fragment, display, retiredIdentityBookmark, numberScaffold);
            numberScaffold = null;
            source = null;
            owner = null;
            return prepared;
        }
        finally
        {
            numberScaffold?.Dispose();
            owner?.Dispose();
            Release(source); Release(shape); Release(shapes);
        }
    }

    internal static bool TryReplace(Document document, Range selected, string omml,
        bool display, out Range replacement, Prepared? prepared = null)
    {
        replacement = null!;
        if (prepared is not null)
        {
            replacement = prepared.Apply(selected);
            return true;
        }
        using var local = Prepare(document, selected, omml, display);
        if (local is null) return false;
        replacement = local.Apply(selected);
        return true;
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
