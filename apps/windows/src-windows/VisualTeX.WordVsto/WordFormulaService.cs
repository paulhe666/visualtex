using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    private const string RangeReferencePrefix = "visualtex-word-vsto-range:";
    private const string InlineBaselineBookmarkPrefix = "VTBL_";
    // Ordinary spaces are used only while Word materializes an inline OMath and
    // are deleted immediately afterwards. Inline OLE formulas keep a real zero-
    // width non-joiner after the object so keyboard input inherits an ordinary
    // text run instead of the OLE object's negative baseline offset.
    private const string InlineMathGuard = " ";
    private const string InlineBaselineSentinel = " ";
    private const string InlineOleTypingAnchor = "\u200C";
    private const string LegacyInlineMathGuard = "\u200B";
    private const string LegacyInlineBaselineSentinel = "\u2060";
    private const string LegacyInlineNonbreakingBaselineSentinel = "\u00A0";
    private const string BulkInlineFormulaPlaceholder = "\uE000";
    private const float ParagraphBeforeOleDisplaySpaceAfterPoints = 0f;
    private readonly Application _application;

    internal sealed class WordViewState
    {
        internal int SelectionStart { get; set; }
        internal int SelectionEnd { get; set; }
        internal int? VerticalPercentScrolled { get; set; }
        internal int? HorizontalPercentScrolled { get; set; }
    }

    private sealed class ResolvedLatexRedrawTarget
    {
        internal WordLatexRedrawTarget Target { get; set; } = new();
        internal PreparedWordBulkFormula Formula { get; set; } = new();
        internal Range SourceRange { get; set; } = null!;
        internal int SourceStart { get; set; }
        internal int SourceEnd { get; set; }
        internal string ExpectedSource { get; set; } = string.Empty;
    }

    private sealed class LatexRedrawSourceContext
    {
        internal bool HasVisibleSurroundingText { get; set; }
        internal int FontContextRelativePosition { get; set; } = -1;
    }

    private sealed class FormulaToLatexTarget
    {
        internal FormulaMetadata Metadata { get; set; } = new();
        internal string ObjectMode { get; set; } = string.Empty;
        internal string LatexSource { get; set; } = string.Empty;
        internal int Start { get; set; }
        internal int End { get; set; }
        internal Range FormulaRange { get; set; } = null!;
        internal InlineShape? OleShape { get; set; }
        internal Bookmark? OmmlBookmark { get; set; }
    }

    private sealed class InlineFollowingTextVisibility
    {
        internal int CharacterCount { get; set; }
        internal int Hidden { get; set; }
    }

    private readonly struct NonProseHostRange
    {
        internal NonProseHostRange(int start, int end)
        {
            Start = start;
            End = end;
        }

        internal int Start { get; }
        internal int End { get; }

        internal bool Contains(int position) => position >= Start && position < End;
    }

    private sealed class MathTypeDisplayParagraphLayout
    {
        internal WdParagraphAlignment Alignment { get; set; }
        internal float LeftIndent { get; set; }
        internal float RightIndent { get; set; }
        internal float FirstLineIndent { get; set; }
        internal float SpaceBefore { get; set; }
        internal float SpaceAfter { get; set; }
        internal WdLineSpacing LineSpacingRule { get; set; }
        internal float LineSpacing { get; set; }
        internal int KeepTogether { get; set; }
        internal int KeepWithNext { get; set; }
        internal int WidowControl { get; set; }
        internal int PageBreakBefore { get; set; }
        internal List<(float Position, WdTabAlignment Alignment, WdTabLeader Leader)> SpecialTabStops { get; } = new();
    }

    public WordFormulaService(Application application)
    {
        _application = application;
    }

    private static void TraceAcceptancePerformance(
        string operation,
        string stage,
        Stopwatch stopwatch,
        ref long checkpoint)
    {
        var acceptanceTrace = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
            "1",
            StringComparison.Ordinal);
        var formatTrace = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF"),
            "1",
            StringComparison.Ordinal);
        var suppressAcceptanceConsoleTrace = string.Equals(
            Environment.GetEnvironmentVariable(
                "VISUALTEX_VSTO_SUPPRESS_ACCEPTANCE_PERF"),
            "1",
            StringComparison.Ordinal);
        if (!acceptanceTrace && !formatTrace)
            return;
        var elapsed = stopwatch.ElapsedMilliseconds;
        var message =
            $"[perf] {operation}.{stage}: +{elapsed - checkpoint}ms ({elapsed}ms total)";
        if (acceptanceTrace && !suppressAcceptanceConsoleTrace)
            Console.WriteLine("    " + message);
        if (formatTrace)
            WordDoubleClickHook.TraceMessage("format-conversion-subperf " + message);
        checkpoint = elapsed;
    }

    public OfficeSelection ReadSelection() => ReadSelection(null);

    public OfficeSelection ReadSelection(Selection? providedSelection)
    {
        Document? document = null;
        Selection? selection = null;
        Range? range = null;
        InlineShapes? inlineShapes = null;
        InlineShape? shape = null;
        Range? externalOleRange = null;
        Bookmark? ommlBookmark = null;
        Range? ommlEquationRange = null;
        var ownsSelection = providedSelection is null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            selection = providedSelection ?? _application.Selection;
            range = selection.Range;
            inlineShapes = range.InlineShapes;
            FormulaMetadata? metadata = null;
            string? objectMode = null;
            if (inlineShapes.Count == 1)
            {
                shape = inlineShapes[1];
                metadata = WordFormulaMetadataReader.TryRead(shape);
                if (metadata is not null)
                {
                    metadata = EnsureUniqueInlineFormulaIdentity(
                        document,
                        shape,
                        metadata);
                    metadata.FontSizePt = FormulaFontSize.InferOleFontSize(
                        shape.Width,
                        shape.Height,
                        metadata);
                    objectMode = WordFormulaMetadataReader.IsNativeOle(shape)
                        ? FormulaOleContract.NativeOleMode
                        : FormulaOleContract.CrossPlatformPictureMode;
                }
                else if (MathTypeOleInterop.IsMathTypeOle(shape))
                {
                    // MathType OLE stays third-party owned. Read its equation
                    // through MathType's IDataObject contract, but never stamp
                    // VisualTeX metadata/bookmarks into the embedded object.
                    metadata = MathTypeOleInterop.ReadMetadata(_application, shape);
                    objectMode = FormulaOleContract.MathTypeOleMode;
                    externalOleRange = shape.Range;
                }
            }
            if (metadata is null)
            {
                ommlBookmark = WordOmmlFormulaStore.FindAtRange(document, range);
                if (ommlBookmark is not null)
                {
                    metadata = WordOmmlFormulaStore.TryRead(document, ommlBookmark);
                    if (metadata is not null)
                    {
                        metadata = WordOmmlNativeSource.RefreshForVisualTeX(
                            document,
                            ommlBookmark,
                            metadata);
                        metadata.FontSizePt = ReadOmmlFontSize(ommlBookmark, metadata);
                        objectMode = FormulaOleContract.WordOmmlMode;
                        // A double-click usually supplies only a collapsed caret
                        // or a small subrange inside the OMath. Word clips an
                        // OMath.Range obtained from such a probe, so carrying the
                        // raw selection into replacement can splice a new formula
                        // into the middle of the old one. Persist the bookmark-
                        // resolved complete equation range as the edit hint.
                        ommlEquationRange = WordOmmlFormulaStore.GetEquationRange(
                            ommlBookmark);
                    }
                }
            }
            if (metadata is null)
            {
                ommlEquationRange = TryResolveNativeOmmlAtRange(document, range);
                if (ommlEquationRange is not null)
                {
                    metadata = WordOmmlNativeSource.CreateForNative(
                        document,
                        ommlEquationRange);
                    objectMode = FormulaOleContract.WordOmmlMode;
                    if (!document.ReadOnly)
                    {
                        try
                        {
                            ommlBookmark = WordOmmlFormulaStore.Wrap(
                                document,
                                ommlEquationRange,
                                metadata,
                                replaceExisting: false);
                            WordOmmlFormulaStore.Save(document, metadata);

                            if (metadata.Numbered
                                && string.Equals(
                                    metadata.DisplayMode,
                                    "block",
                                    StringComparison.OrdinalIgnoreCase))
                            {
                                // A pasted native #(SEQ) OMath keeps its direct SEQ
                                // field but Word cannot duplicate the source formula's
                                // bookmark names. Adopt it with a fresh FormulaId, then
                                // atomically rebuild only this copied OMath so its
                                // VTEqNum/VTEq/VTEqCap aliases belong to the copy and
                                // the source keeps its own identity and body REF target.
                                WordEquationNumbering.ReconcileFormula(
                                    document,
                                    ommlEquationRange,
                                    WordOmmlFormulaStore.EstimateHeightPoints(
                                        ommlEquationRange),
                                    metadata,
                                    numberingOrderMayHaveChanged: true,
                                    reuseExistingNumberedTableFormatting: false,
                                    knownNumberedTable: null);
                                Release(ommlEquationRange);
                                ommlEquationRange = null;
                                Release(ommlBookmark);
                                ommlBookmark = WordOmmlFormulaStore.FindByFormulaId(
                                    document,
                                    metadata.FormulaId)
                                    ?? throw new InvalidOperationException(
                                        "The adopted numbered OMML copy lost its new VTOMML identity.");
                                ommlEquationRange = WordOmmlFormulaStore.GetEquationRange(
                                    ommlBookmark);
                            }
                        }
                        catch
                        {
                            try { ommlBookmark?.Delete(); } catch { }
                            Release(ommlBookmark);
                            ommlBookmark = null;
                            try
                            {
                                WordOmmlFormulaStore.Delete(
                                    document,
                                    metadata.FormulaId);
                            }
                            catch { }
                            // The native equation is still readable even if Word
                            // refuses to persist the VisualTeX adoption metadata.
                        }
                    }
                }
            }
            return new OfficeSelection
            {
                Host = "word",
                DocumentId = DocumentIdentity(document),
                // OLE keeps the exact source range as a fast edit hint. OMML
                // must carry the complete bookmark-resolved equation range;
                // Word clips ranges obtained from a caret inside an OMath.
                ObjectId = RangeReference(ommlEquationRange ?? externalOleRange ?? range),
                ReadOnly = document.ReadOnly,
                FormulaId = metadata?.FormulaId,
                Metadata = metadata,
                ObjectMode = objectMode,
            };
        }
        finally
        {
            Release(ommlEquationRange);
            Release(ommlBookmark);
            Release(externalOleRange);
            Release(shape);
            Release(inlineShapes);
            Release(range);
            if (ownsSelection) Release(selection);
            Release(document);
        }
    }

    private static FormulaMetadata EnsureUniqueInlineFormulaIdentity(
        Document document,
        InlineShape shape,
        FormulaMetadata metadata)
    {
        if (document.ReadOnly || string.IsNullOrWhiteSpace(metadata.FormulaId))
            return metadata;

        Bookmarks? bookmarks = null;
        Bookmark? ownerBookmark = null;
        Range? ownerRange = null;
        Range? shapeRange = null;
        try
        {
            shapeRange = shape.Range;
            bookmarks = document.Bookmarks;
            var currentName = WordFormulaMetadataReader.IdentityBookmarkName(metadata.FormulaId);
            if (bookmarks.Exists(currentName))
            {
                ownerBookmark = bookmarks[currentName];
                ownerRange = ownerBookmark.Range;
                if (RangesIdentifySameInlineFormula(ownerRange, shapeRange))
                {
                    // A shape replacement can collapse a bookmark at the same
                    // insertion point. Refresh the anchor to cover the live OLE
                    // object so the identity remains durable across later edits.
                    if (ownerRange.Start != shapeRange.Start
                        || ownerRange.End != shapeRange.End)
                    {
                        ownerBookmark.Delete();
                        Release(ownerBookmark);
                        ownerBookmark = null;
                        var refreshedIdentity = bookmarks.Add(currentName, shapeRange);
                        Release(refreshedIdentity);
                    }
                    return metadata;
                }

                return RekeyCopiedInlineFormula(
                    document,
                    shape,
                    shapeRange,
                    bookmarks,
                    metadata);
            }

            // Existing numbered formulas already have durable VisualTeX
            // bookmarks. If those bookmarks belong to another table, this is a
            // copied formula even when the newer VTO_ identity bookmark did not
            // exist in an older document yet. Preserve the original equation's
            // FormulaId and re-key the copy.
            if (metadata.Numbered
                && WordEquationNumbering.HasCompleteFormulaNumberingArtifacts(
                    document,
                    metadata.FormulaId)
                && !WordEquationNumbering.FormulaRangeOwnsNumberingArtifacts(
                    document,
                    shapeRange,
                    metadata.FormulaId))
            {
                return RekeyCopiedInlineFormula(
                    document,
                    shape,
                    shapeRange,
                    bookmarks,
                    metadata);
            }

            var identityBookmark = bookmarks.Add(currentName, shapeRange);
            Release(identityBookmark);
            return metadata;
        }
        catch
        {
            // Identity repair must never make an otherwise readable formula
            // impossible to open. Targeted edit operations still carry the
            // exact source range as a secondary lookup hint.
            return metadata;
        }
        finally
        {
            Release(shapeRange);
            Release(ownerRange);
            Release(ownerBookmark);
            Release(bookmarks);
        }
    }

    private static void TraceNativeOleIdentityState(
        Document document,
        string formulaId,
        string stage)
    {
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            return;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        InlineShapes? bookmarkShapes = null;
        InlineShapes? allShapes = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = WordFormulaMetadataReader.IdentityBookmarkName(formulaId);
            var bookmarkExists = bookmarks.Exists(name);
            var bookmarkDescription = "missing";
            if (bookmarkExists)
            {
                bookmark = bookmarks[name];
                bookmarkRange = bookmark.Range;
                bookmarkShapes = bookmarkRange.InlineShapes;
                bookmarkDescription =
                    $"{bookmarkRange.Start}:{bookmarkRange.End}/shapes={bookmarkShapes.Count}";
            }
            allShapes = document.InlineShapes;
            var shapeDescriptions = new List<string>();
            for (var index = 1; index <= allShapes.Count; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                try
                {
                    shape = allShapes[index];
                    range = shape.Range;
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    shapeDescriptions.Add(
                        $"#{index}@{range.Start}:{range.End}/native={WordFormulaMetadataReader.IsNativeOle(shape)}/formulaId={metadata?.FormulaId ?? "<none>"}");
                }
                catch (Exception error)
                {
                    shapeDescriptions.Add($"#{index}/error={error.Message}");
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }
            Console.WriteLine(
                $"    [ole-identity] stage={stage} formulaId={formulaId} bookmark={bookmarkDescription} documentShapes={allShapes.Count} [{string.Join(";", shapeDescriptions)}]");
        }
        catch (Exception error)
        {
            Console.WriteLine(
                $"    [ole-identity] stage={stage} formulaId={formulaId} diagnosticError={error.Message}");
        }
        finally
        {
            Release(allShapes);
            Release(bookmarkShapes);
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static InlineShape? FindNativeOleByIdentityBookmark(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? candidate = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = WordFormulaMetadataReader.IdentityBookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return null;
            bookmark = bookmarks[name];
            range = bookmark.Range;
            shapes = range.InlineShapes;
            if (shapes.Count != 1) return null;
            candidate = shapes[1];
            if (!WordFormulaMetadataReader.IsNativeOle(candidate)) return null;
            var result = candidate;
            candidate = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(candidate);
            Release(shapes);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void BindOleIdentityBookmark(InlineShape shape, string formulaId)
    {
        Range? shapeRange = null;
        Document? document = null;
        Bookmarks? bookmarks = null;
        Bookmark? existing = null;
        Bookmark? bound = null;
        try
        {
            shapeRange = shape.Range;
            document = shapeRange.Document;
            if (document.ReadOnly) return;
            bookmarks = document.Bookmarks;
            var name = WordFormulaMetadataReader.IdentityBookmarkName(formulaId);
            if (bookmarks.Exists(name))
            {
                existing = bookmarks[name];
                existing.Delete();
                Release(existing);
                existing = null;
            }
            bound = bookmarks.Add(name, shapeRange);
        }
        catch
        {
            // Embedded metadata remains the source of truth. If Word rejects the
            // bookmark write, first read will retry the full duplicate-repair path.
        }
        finally
        {
            Release(bound);
            Release(existing);
            Release(bookmarks);
            Release(document);
            Release(shapeRange);
        }
    }

    private static FormulaMetadata RekeyCopiedInlineFormula(
        Document document,
        InlineShape shape,
        Range shapeRange,
        Bookmarks bookmarks,
        FormulaMetadata metadata)
    {
        var originalFormulaId = metadata.FormulaId;
        var newFormulaId = Guid.NewGuid().ToString("D");
        var rekeyed = WordFormulaMetadataReader.CloneWithFormulaId(
            metadata,
            newFormulaId);
        var identityBookmark = bookmarks.Add(
            WordFormulaMetadataReader.IdentityBookmarkName(newFormulaId),
            shapeRange);
        Release(identityBookmark);

        // A copied OLE object still contains the source object's embedded
        // FormulaId. The identity bookmark makes the current Word session see
        // the copy as independent, but relying on that bookmark alone is not
        // durable enough: some Word paste/save paths can drop or move bookmarks.
        // Persist the re-key into the embedded VisualTeX object as well, with the
        // shape cache as a fallback if a dormant pasted OLE cannot be activated.
        try { WordFormulaMetadataReader.Write(shape, rekeyed); }
        catch { WordFormulaMetadataReader.CacheMetadata(shape, rekeyed); }

        if (rekeyed.Numbered
            && string.Equals(rekeyed.DisplayMode, "block", StringComparison.Ordinal))
        {
            InlineShape? originalShape = null;
            Range? originalRange = null;
            try
            {
                // Word may move the copied table's visible-number bookmark away
                // from the original table when both carry the same FormulaId.
                // Repair only the two affected formulas instead of running a
                // document-wide Reconcile(), which used to freeze Word for
                // several seconds even in a six-formula document.
                var copyOwnsOriginalArtifacts =
                    WordEquationNumbering.FormulaRangeOwnsNumberingArtifacts(
                        document,
                        shapeRange,
                        originalFormulaId);
                if (copyOwnsOriginalArtifacts)
                {
                    WordEquationNumbering.RemoveFormulaNumberingArtifacts(
                        document,
                        originalFormulaId);
                    originalShape = FindByFormulaId(document, originalFormulaId);
                    if (originalShape is not null)
                    {
                        var originalMetadata = WordFormulaMetadataReader.TryRead(originalShape);
                        if (originalMetadata?.Numbered == true
                            && string.Equals(
                                originalMetadata.DisplayMode,
                                "block",
                                StringComparison.Ordinal))
                        {
                            originalRange = originalShape.Range;
                            WordEquationNumbering.ReconcileFormula(
                                document,
                                originalRange,
                                originalShape.Height,
                                originalMetadata,
                                numberingOrderMayHaveChanged: false);
                        }
                    }
                }

                WordEquationNumbering.ReconcileFormula(
                    document,
                    shapeRange,
                    shape.Height,
                    rekeyed,
                    numberingOrderMayHaveChanged: true,
                    reuseExistingNumberedTableFormatting: true);
            }
            catch
            {
                // Preserve the old repair safety net for genuinely malformed
                // legacy documents. Healthy copies stay on the local fast path.
                WordEquationNumbering.Reconcile(document);
            }
            finally
            {
                Release(originalRange);
                Release(originalShape);
            }
        }
        return rekeyed;
    }

    private static bool RangesIdentifySameInlineFormula(Range owner, Range candidate)
    {
        if (owner.Start == candidate.Start && owner.End == candidate.End)
            return true;
        if (owner.Start == owner.End && owner.Start == candidate.Start)
            return true;
        return owner.Start <= candidate.Start && owner.End >= candidate.End;
    }

    public OfficeSelection? ReadVisualTeXOmmlAtScreenPoint(
        int screenX,
        int screenY)
    {
        Document? document = null;
        Window? window = null;
        object? pointObject = null;
        Range? pointRange = null;
        Range? equationRange = null;
        Bookmark? bookmark = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            window = _application.ActiveWindow;
            pointObject = window.RangeFromPoint(screenX, screenY);
            pointRange = pointObject as Range;
            if (pointRange is null) return null;
            pointObject = null;

            equationRange = TryResolveNativeOmmlAtRange(document, pointRange);
            if (equationRange is null
                || !ScreenPointHitsRange(window, equationRange, screenX, screenY))
                return null;

            bookmark = WordOmmlFormulaStore.FindAtRange(document, equationRange);
            FormulaMetadata? metadata;
            if (bookmark is not null)
            {
                metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
                if (metadata is null) return null;
                metadata = WordOmmlNativeSource.RefreshForVisualTeX(
                    document,
                    bookmark,
                    metadata);
                metadata.FontSizePt = ReadOmmlFontSize(bookmark, metadata);
            }
            else
            {
                // MathType equations that have already been converted to Word
                // OMML are ordinary native OMath objects: there is no durable
                // MathType identity left for VisualTeX to key on. Adopt any
                // native OMML hit by the real mouse point so the very first
                // double-click can open VisualTeX instead of requiring a prior
                // Ribbon edit to create VTOMML metadata.
                metadata = WordOmmlNativeSource.CreateForNative(
                    document,
                    equationRange);
                if (!document.ReadOnly)
                {
                    try
                    {
                        bookmark = WordOmmlFormulaStore.Wrap(
                            document,
                            equationRange,
                            metadata,
                            replaceExisting: false);
                        WordOmmlFormulaStore.Save(document, metadata);
                    }
                    catch
                    {
                        try { bookmark?.Delete(); } catch { }
                        Release(bookmark);
                        bookmark = null;
                        try
                        {
                            WordOmmlFormulaStore.Delete(
                                document,
                                metadata.FormulaId);
                        }
                        catch { }
                        // The captured full equation range and Session metadata
                        // are enough for this edit even if Word temporarily
                        // rejects persistent adoption bookkeeping.
                    }
                }
            }
            return new OfficeSelection
            {
                Host = "word",
                DocumentId = DocumentIdentity(document),
                ObjectId = RangeReference(equationRange),
                ReadOnly = document.ReadOnly,
                FormulaId = metadata.FormulaId,
                Metadata = metadata,
                ObjectMode = FormulaOleContract.WordOmmlMode,
            };
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(bookmark);
            Release(equationRange);
            Release(pointRange);
            Release(pointObject);
            Release(window);
            Release(document);
        }
    }

    public bool IsFormulaAtScreenPoint(
        OfficeSelection? selected,
        int screenX,
        int screenY)
    {
        if (selected?.Metadata is null
            || string.IsNullOrWhiteSpace(selected.FormulaId))
            return false;

        Document? document = null;
        Window? window = null;
        InlineShape? shape = null;
        Bookmark? bookmark = null;
        Range? formulaRange = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null
                || !string.Equals(
                    DocumentIdentity(document),
                    selected.DocumentId,
                    StringComparison.Ordinal))
                return false;
            window = _application.ActiveWindow;

            if (string.Equals(
                    selected.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(
                    document,
                    selected.FormulaId!);
                if (bookmark is null) return false;
                formulaRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            }
            else if (string.Equals(
                         selected.ObjectMode,
                         FormulaOleContract.MathTypeOleMode,
                         StringComparison.Ordinal))
            {
                shape = FindMathTypeOleByRange(document, selected.ObjectId);
                if (shape is null) return false;
                formulaRange = shape.Range;
            }
            else
            {
                shape = FindByFormulaId(
                    document,
                    selected.FormulaId!,
                    selected.ObjectId);
                if (shape is null) return false;
                formulaRange = shape.Range;
            }

            return ScreenPointHitsRange(
                window,
                formulaRange,
                screenX,
                screenY);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(formulaRange);
            Release(bookmark);
            Release(shape);
            Release(window);
            Release(document);
        }
    }

    private static bool ScreenPointHitsRange(
        Window window,
        Range range,
        int screenX,
        int screenY)
    {
        try
        {
            window.GetPoint(
                out var left,
                out var top,
                out var width,
                out var height,
                range);
            return WordDoubleClickRouting.ScreenPointHitsFormulaRectangle(
                screenX,
                screenY,
                left,
                top,
                width,
                height);
        }
        catch
        {
            return false;
        }
    }

    private static Range? TryResolveNativeOmmlAtRange(
        Document document,
        Range selectionRange)
    {
        Range? probe = null;
        OMaths? maths = null;
        Range? best = null;
        var candidates = new List<(int Start, int End)>();
        try
        {
            probe = selectionRange.Duplicate;
            maths = probe.OMaths;
            if (maths.Count == 0)
            {
                Release(maths);
                maths = null;
                Release(probe);
                probe = null;
                Range? content = null;
                try
                {
                    content = document.Content;
                    var start = Math.Max(content.Start, selectionRange.Start - 1);
                    var end = Math.Min(content.End, selectionRange.End + 1);
                    probe = document.Range(start, end);
                }
                finally { Release(content); }
                maths = probe.OMaths;
            }

            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    var containsCaret = selectionRange.Start == selectionRange.End
                        && selectionRange.Start >= range.Start
                        && selectionRange.Start <= range.End;
                    var overlapsSelection = selectionRange.Start < selectionRange.End
                        && range.Start < selectionRange.End
                        && range.End > selectionRange.Start;
                    if (!containsCaret && !overlapsSelection) continue;
                    candidates.Add((range.Start, range.End));
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }

            if (candidates.Count == 0) return null;
            var roots = candidates
                .Where(candidate => !candidates.Any(other =>
                    other != candidate
                    && other.Start <= candidate.Start
                    && other.End >= candidate.End))
                .Distinct()
                .ToArray();
            if (roots.Length != 1) return null;
            best = document.Range(roots[0].Start, roots[0].End);
            var result = best;
            best = null;
            return result;
        }
        catch { return null; }
        finally
        {
            Release(best);
            Release(maths);
            Release(probe);
        }
    }

    public string ReadActiveDocumentId()
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            return DocumentIdentity(document);
        }
        finally { Release(document); }
    }

    private enum SelectedOleKind
    {
        None,
        VisualTeX,
        MathType,
    }

    public bool IsSelectedNativeOle() =>
        ReadSelectedOleKind() == SelectedOleKind.VisualTeX;

    public bool IsSelectedMathTypeOle() =>
        ReadSelectedOleKind() == SelectedOleKind.MathType;

    public bool OpenMathTypeNativeEditorAtRange(
        int targetStart,
        int targetEnd,
        int screenX,
        int screenY)
    {
        Document? document = null;
        Window? window = null;
        Range? content = null;
        Range? targetRange = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        OLEFormat? format = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null) return false;
            content = document.Content;
            var start = Math.Max(content.Start, Math.Min(targetStart, content.End));
            var end = Math.Max(start, Math.Min(targetEnd, content.End));
            if (end <= start) return false;

            targetRange = document.Range(start, end);
            shapes = targetRange.InlineShapes;
            if (shapes.Count != 1) return false;
            shape = shapes[1];
            if (!MathTypeOleInterop.IsMathTypeOle(shape)) return false;
            shapeRange = shape.Range;
            if (shapeRange.End <= start || shapeRange.Start >= end) return false;

            window = _application.ActiveWindow;
            if (!ScreenPointHitsRange(window, shapeRange, screenX, screenY))
                return false;

            format = shape.OLEFormat;
            object openVerb = (int)WdOLEVerb.wdOLEVerbOpen;
            format.DoVerb(ref openVerb);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(format);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(targetRange);
            Release(content);
            Release(window);
            Release(document);
        }
    }

    public bool IsSelectedInterceptableOle() =>
        ReadSelectedOleKind() is SelectedOleKind.VisualTeX or SelectedOleKind.MathType;

    private SelectedOleKind ReadSelectedOleKind()
    {
        Selection? selection = null;
        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        OLEFormat? format = null;
        try
        {
            selection = _application.Selection;
            range = selection.Range;
            shapes = range.InlineShapes;
            if (shapes.Count != 1) return SelectedOleKind.None;
            shape = shapes[1];
            if (shape.Type is not WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                and not WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                return SelectedOleKind.None;
            format = shape.OLEFormat;
            if (string.Equals(
                    format.ProgID,
                    FormulaOleContract.ProgId,
                    StringComparison.OrdinalIgnoreCase))
                return SelectedOleKind.VisualTeX;
            return MathTypeOleInterop.TryResolveCapabilities(format.ProgID, out _)
                ? SelectedOleKind.MathType
                : SelectedOleKind.None;
        }
        catch
        {
            return SelectedOleKind.None;
        }
        finally
        {
            Release(format);
            Release(shape);
            Release(shapes);
            Release(range);
            Release(selection);
        }
    }

    public void NormalizeTypingCaretAfterInlineFormula(Selection selection)
    {
        if (selection is null) return;
        if (WordEquationNumbering.ExpandCompactTrailingTypingParagraph(selection))
            return;
        Range? caret = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        try
        {
            caret = selection.Range;
            if (caret.Start != caret.End) return;
            paragraphs = caret.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            shapes = paragraphRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                if (TryMoveCaretPastInlineBoundary(selection, caret.Start, shape))
                    return;
            }
        }
        catch
        {
            // A selection-change repair must never interrupt ordinary Word use.
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(caret);
        }
    }

    private bool TryMoveCaretPastInlineBoundary(
        Selection selection,
        int caretPosition,
        InlineShape shape)
    {
        if (!WordFormulaMetadataReader.IsNativeOle(shape)) return false;
        var metadata = WordFormulaMetadataReader.TryRead(shape);
        if (metadata is null
            || !string.Equals(metadata.DisplayMode, "inline", StringComparison.Ordinal))
            return false;

        Range? formulaRange = null;
        Document? document = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? sentinel = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            formulaRange = shape.Range;
            if (caretPosition < formulaRange.End) return false;
            NormalizeFollowingInlineProseBaseline(formulaRange);
            document = formulaRange.Document;
            bookmarks = document.Bookmarks;
            var bookmarkName = InlineBaselineBookmarkName(metadata.FormulaId);
            if (bookmarks.Exists(bookmarkName))
            {
                bookmark = bookmarks[bookmarkName];
                sentinel = bookmark.Range;
                if (IsUsableInlineBaselineSentinel(sentinel, formulaRange)
                    && caretPosition <= sentinel.End)
                {
                    var boundaryPosition = NormalizeInlineBaselineBoundary(
                        document,
                        formulaRange,
                        metadata.FormulaId);
                    PositionSelectionAfterInlineTypingAnchor(
                        selection,
                        formulaRange,
                        boundaryPosition);
                    ApplyInlineTypingFormattingToSelection(
                        selection,
                        formulaRange);
                    return true;
                }
            }

            if (caretPosition != formulaRange.End) return false;
            var target = EnsureInlineBaselineSentinel(formulaRange, metadata.FormulaId);
            PositionSelectionAfterInlineTypingAnchor(
                selection,
                formulaRange,
                target);
            ApplyInlineTypingFormattingToSelection(
                selection,
                formulaRange);
            return true;
        }
        finally
        {
            Release(font);
            Release(sentinel);
            Release(bookmark);
            Release(bookmarks);
            Release(document);
            Release(formulaRange);
        }
    }

    private static void PositionSelectionAfterInlineTypingAnchor(
        Selection selection,
        Range formulaRange,
        int boundaryPosition)
    {
        if (selection.Start == boundaryPosition
            && selection.End == boundaryPosition)
            return;

        selection.SetRange(formulaRange.End, formulaRange.End);
        if (boundaryPosition > formulaRange.End)
        {
            _ = selection.MoveRight(
                WdUnits.wdCharacter,
                Math.Max(1, boundaryPosition - formulaRange.End),
                WdMovementType.wdMove);
        }
        if (selection.Start != boundaryPosition
            || selection.End != boundaryPosition)
            selection.SetRange(boundaryPosition, boundaryPosition);
    }

    private static void ApplyInlineTypingFormattingToSelection(
        Selection selection,
        Range formulaRange)
    {
        Range? source = null;
        Range? caret = null;
        Microsoft.Office.Interop.Word.Font? sourceFont = null;
        Microsoft.Office.Interop.Word.Font? selectionFont = null;
        try
        {
            caret = selection.Range;
            source = FindInlineTypingFormatSource(formulaRange, caret);
            if (source is not null)
            {
                sourceFont = source.Font;
                selectionFont = selection.Font;
                selectionFont.Name = sourceFont.Name;
                try { selectionFont.NameAscii = sourceFont.NameAscii; } catch { }
                try { selectionFont.NameFarEast = sourceFont.NameFarEast; } catch { }
                try { selectionFont.NameOther = sourceFont.NameOther; } catch { }
                selectionFont.Size = sourceFont.Size;
                selectionFont.Bold = sourceFont.Bold;
                selectionFont.Italic = sourceFont.Italic;
                try { selectionFont.Underline = sourceFont.Underline; } catch { }
                try { selectionFont.Color = sourceFont.Color; } catch { }
                var sourcePosition = sourceFont.Position;
                selectionFont.Position = sourcePosition == (int)WdConstants.wdUndefined
                    ? 0
                    : sourcePosition;
            }
            else
            {
                selectionFont = selection.Font;
                selectionFont.Position = 0;
            }
            selectionFont.Hidden = 0;
            selectionFont.Subscript = 0;
            selectionFont.Superscript = 0;
            try { selectionFont.Spacing = 0; } catch { }
            try { selectionFont.Scaling = 100; } catch { }
        }
        finally
        {
            Release(selectionFont);
            Release(sourceFont);
            Release(source);
            Release(caret);
        }
    }

    private static bool TryResolveWordFontSize(float value, out float fontSizePt)
    {
        fontSizePt = FormulaFontSize.DefaultPt;
        if (float.IsNaN(value)
            || float.IsInfinity(value)
            || value < FormulaFontSize.MinimumPt
            || value > FormulaFontSize.MaximumPt)
            return false;
        fontSizePt = FormulaFontSize.Normalize(value);
        return true;
    }

    public float ReadCurrentTypingFontSize()
    {
        Selection? selection = null;
        Range? selectionRange = null;
        Range? probeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            selection = _application.Selection;
            selectionRange = selection.Range;
            font = selection.Font;
            if (TryResolveWordFontSize(font.Size, out var selectedSize))
                return selectedSize;
            Release(font);
            font = null;

            paragraphs = selectionRange.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
            }

            // A collapsed insertion point can report Word's mixed-size sentinel.
            // Prefer the character immediately before the caret in the current
            // paragraph, then the next character, before falling back to the
            // paragraph run as a whole. This makes a new formula inherit the
            // actual surrounding body text instead of an arbitrary global size.
            if (selectionRange.Start > (paragraphRange?.Start ?? 0))
            {
                probeRange = selectionRange.Duplicate;
                probeRange.SetRange(selectionRange.Start - 1, selectionRange.Start);
                font = probeRange.Font;
                if (TryResolveWordFontSize(font.Size, out var previousSize))
                    return previousSize;
                Release(font);
                font = null;
                Release(probeRange);
                probeRange = null;
            }

            var paragraphEnd = Math.Max(
                paragraphRange?.Start ?? selectionRange.Start,
                (paragraphRange?.End ?? selectionRange.End) - 1);
            if (selectionRange.Start < paragraphEnd)
            {
                probeRange = selectionRange.Duplicate;
                probeRange.SetRange(selectionRange.Start, selectionRange.Start + 1);
                font = probeRange.Font;
                if (TryResolveWordFontSize(font.Size, out var nextSize))
                    return nextSize;
                Release(font);
                font = null;
                Release(probeRange);
                probeRange = null;
            }

            if (paragraphRange is not null)
            {
                if (paragraphRange.End > paragraphRange.Start)
                    paragraphRange.End -= 1;
                font = paragraphRange.Font;
                if (TryResolveWordFontSize(font.Size, out var paragraphSize))
                    return paragraphSize;
            }
            return FormulaFontSize.DefaultPt;
        }
        catch
        {
            return FormulaFontSize.DefaultPt;
        }
        finally
        {
            Release(font);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probeRange);
            Release(selectionRange);
            Release(selection);
        }
    }

    public float? GetSelectedFormulaFontSize()
    {
        // Ribbon getText/getEnabled callbacks run immediately after a Word
        // SelectionChange. They must be strictly read-only: calling ReadSelection()
        // here can adopt an unowned native OMath by adding a bookmark/CustomXML
        // part, which mutates the document while Word is entering its native
        // equation editor and can suppress the normal equation editing frame.
        Document? document = null;
        Selection? selection = null;
        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Bookmark? bookmark = null;
        Range? equationRange = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null) return null;
            selection = _application.Selection;
            range = selection.Range;

            shapes = range.InlineShapes;
            if (shapes.Count == 1)
            {
                shape = shapes[1];
                var metadata = WordFormulaMetadataReader.TryRead(shape);
                if (metadata is not null)
                    return FormulaFontSize.InferOleFontSize(
                        shape.Width,
                        shape.Height,
                        metadata);
            }

            bookmark = WordOmmlFormulaStore.FindAtRange(document, range);
            if (bookmark is not null)
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
                if (metadata is not null)
                {
                    equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                    return ReadFormulaFontSizeWithoutMutation(
                        equationRange,
                        metadata);
                }
            }

            equationRange = TryResolveNativeOmmlAtRange(document, range);
            return equationRange is null
                ? null
                : ReadFormulaFontSizeWithoutMutation(equationRange, metadata: null);
        }
        catch { return null; }
        finally
        {
            Release(equationRange);
            Release(bookmark);
            Release(shape);
            Release(shapes);
            Release(range);
            Release(selection);
            Release(document);
        }
    }

    private static float ReadFormulaFontSizeWithoutMutation(
        Range equationRange,
        FormulaMetadata? metadata)
    {
        // Numbered display OMML intentionally contains mixed native run sizes to
        // emulate Word display-style fraction arguments inside the OLE-compatible
        // tab host. The persisted VisualTeX metadata is therefore the semantic
        // size shown in the ribbon; Word may report its mixed-size sentinel here.
        if (metadata is not null && IsNumberedBlockOmml(metadata))
            return FormulaFontSize.ResolveSemanticFontSize(metadata);

        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = equationRange.Font;
            if (TryNormalizeDefinedWordFontSize(font.Size, out var size))
                return size;
        }
        catch { }
        finally { Release(font); }

        return metadata is null
            ? FormulaFontSize.DefaultPt
            : FormulaFontSize.ResolveSemanticFontSize(metadata);
    }

    public float SetSelectedFormulaFontSize(double requestedFontSizePt)
    {
        var selected = ReadSelection();
        if (selected.Metadata is null || string.IsNullOrWhiteSpace(selected.FormulaId))
            throw new InvalidOperationException("请先选择一个 VisualTeX 公式。");

        var target = FormulaFontSize.Normalize(requestedFontSizePt);
        Document? document = null;
        InlineShape? shape = null;
        Bookmark? bookmark = null;
        Range? equationRange = null;
        Table? numberedHeightTable = null;
        UndoRecord? undoRecord = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            undoRecord = BeginUndoRecord("VisualTeX Set Formula Font Size");
            var metadata = selected.Metadata;
            var sourceSemanticFontSize = FormulaFontSize.ResolveSemanticFontSize(metadata);

            if (string.Equals(
                    selected.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, selected.FormulaId!)
                    ?? throw new InvalidOperationException("The selected Word OMML formula no longer exists.");
                equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                metadata.FontSizePt = target;
                metadata.UpdatedWithVersion = "1.2.5";
                metadata.UpdatedAt = DateTimeOffset.UtcNow.ToString("O");
                RemoveInlineBaselineSentinel(document, metadata.FormulaId);
                // A numbered block OMML is genuine Word display math. Its separate
                // external REF host must not affect inline/block semantic detection.
                var alignInline = !IsNumberedBlockOmml(metadata)
                    && ShouldAlignInline(equationRange, metadata);
                if (alignInline) metadata.DisplayMode = "inline";
                if (IsNumberedBlockOmml(metadata))
                {
                    // Genuine Word display math can be resized in place. Replacing
                    // the OMath is unnecessary and can invalidate the separate
                    // floating REF Shape. Word's OpenType MATH engine derives all
                    // fraction, radical, matrix and large-operator display geometry
                    // from this one semantic run size.
                    ApplyOmmlTypography(equationRange, target, metadata);
                    var measuredDisplayHeight =
                        WordOmmlConverter.MeasurePreparedDisplayHeightPoints(
                            _application,
                            document,
                            equationRange.WordOpenXML ?? string.Empty,
                            document.OMathFontName);
                    if (measuredDisplayHeight.HasValue)
                    {
                        numberedHeightTable =
                            WordEquationNumbering.FindNumberedEquationTable(
                                document,
                                metadata.FormulaId);
                        if (numberedHeightTable is not null)
                            WordEquationNumbering.ApplyNativeOmmlTableMinimumDisplayHeight(
                                numberedHeightTable,
                                measuredDisplayHeight.Value);
                        Release(numberedHeightTable);
                        numberedHeightTable = null;
                    }
                    TryReconcileOmml(document, bookmark, equationRange, metadata);

                    Release(equationRange);
                    equationRange = null;
                    Release(bookmark);
                    bookmark = WordOmmlFormulaStore.FindByFormulaId(
                            document,
                            metadata.FormulaId)
                        ?? throw new InvalidOperationException(
                            "The numbered OMML formula bookmark was lost after its font-size update.");
                    equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                    // TryReconcileOmml already performs the one post-field Shape
                    // geometry pass. Avoid repaginating the same document again.
                    WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                        metadata,
                        equationRange);
                    if (!WordOmmlFormulaStore.IsCanonicalAnchor(bookmark, equationRange))
                    {
                        Release(bookmark);
                        bookmark = WordOmmlFormulaStore.Wrap(
                            document,
                            equationRange,
                            metadata,
                            replaceExisting: true);
                    }
                    WordOmmlFormulaStore.Save(document, metadata);
                    return target;
                }

                ApplyOmmlTypography(equationRange, target, metadata);
                WordOmmlNativeSource.StampFingerprint(metadata, equationRange);
                WordOmmlFormulaStore.Save(document, metadata);
                if (alignInline)
                    FinalizeInlineOmmlBoundary(
                        document,
                        equationRange,
                        metadata.FormulaId,
                        moveCaretOutsideMath: true);
                else
                    TryReconcileOmml(document, bookmark, equationRange, metadata);
                return target;
            }

            shape = FindByFormulaId(document, selected.FormulaId!)
                ?? throw new InvalidOperationException("The selected Word formula no longer exists.");
            var alignOleInline = ShouldAlignInline(shape, metadata);
            if (alignOleInline) metadata.DisplayMode = "inline";
            var existingFontPosition = ReadDefinedShapeFontPosition(shape);
            metadata.FontSizePt = target;
            var size = FormulaFontSize.OleSizeAt(metadata, target);
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            shape.Width = size.Width;
            shape.Height = size.Height;
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            StoreWordInlineOleSize(
                metadata,
                shape.Width,
                shape.Height,
                alignOleInline);
            if (WordFormulaMetadataReader.IsNativeOle(shape))
            {
                // Native VisualTeX OLE metadata is authoritative. Updating only
                // Word's Title/AlternativeText cache leaves the embedded object
                // with the old semantic size, so the next reopen can snap a 14 pt
                // formula back to the previous 13.5 pt value.
                WordFormulaMetadataReader.Write(shape, metadata);
            }
            else
            {
                var encoded = FormulaMetadataCodec.Encode(metadata);
                shape.Title = encoded;
                shape.AlternativeText = encoded;
            }
            if (alignOleInline)
            {
                ApplyInlineBaseline(
                    shape,
                    shape.Height,
                    (float)(metadata.RenderHeightPx ?? 0),
                    metadata.Baseline.HasValue ? (float?)metadata.Baseline.Value : null,
                    existingFontPosition,
                    sourceSemanticFontSize,
                    target);
                RestoreTypingBaselineAfter(shape, ensureTypingAnchor: true);
            }
            else
            {
                RemoveInlineBaselineSentinel(document, metadata.FormulaId);
                RemoveInlineOleTypingAnchorAfter(shape);
                ResetShapeFontPosition(shape);
                TryReconcileShape(document, shape, metadata);
            }
            return target;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(numberedHeightTable);
            Release(equationRange);
            Release(bookmark);
            Release(shape);
            Release(document);
        }
    }

    public string DeleteSelectedFormula()
    {
        var selected = ReadSelection();
        var formulaId = selected.FormulaId;
        if (string.IsNullOrWhiteSpace(formulaId))
            throw new InvalidOperationException("Please select one VisualTeX formula first.");
        var requiredFormulaId = formulaId!;

        Document? document = null;
        InlineShape? shape = null;
        Bookmark? ommlBookmark = null;
        Range? ommlRange = null;
        UndoRecord? undoRecord = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            undoRecord = BeginUndoRecord("VisualTeX Delete Formula");
            if (string.Equals(
                    selected.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                ommlBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    document,
                    requiredFormulaId)
                    ?? throw new InvalidOperationException(
                        "The selected Word OMML formula no longer exists.");
                ommlRange = WordOmmlFormulaStore.GetEquationRange(ommlBookmark);
                ommlBookmark.Delete();
                ommlRange.Delete();
                RemoveInlineBaselineSentinel(document, requiredFormulaId);
                WordOmmlFormulaStore.Delete(document, requiredFormulaId);
            }
            else
            {
                shape = FindByFormulaId(document, requiredFormulaId)
                    ?? throw new InvalidOperationException(
                        "The selected Word formula no longer exists.");
                RemoveInlineBaselineSentinel(document, requiredFormulaId);
                RemoveInlineOleTypingAnchorAfter(shape);
                shape.Delete();
            }
            WordEquationNumbering.TryReconcile(document);
            return requiredFormulaId;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(ommlRange);
            Release(ommlBookmark);
            Release(shape);
            Release(document);
        }
    }

    public int UpdateEquationNumbers()
    {
        Document? document = null;
        UndoRecord? undoRecord = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureEquationFieldResultsVisible(document);
            undoRecord = BeginUndoRecord("VisualTeX Update Equation Numbers");
            var visualTeXCount = WordEquationNumbering.UpdateEquationNumbers(document);
            var mathTypeCount = MathTypeEquationNumbering.UpdateEquationNumbers(document);
            return visualTeXCount + mathTypeCount;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(document);
        }
    }

    public string GetEquationNumberFormatId()
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            return WordEquationNumbering.GetEquationNumberFormatId(document);
        }
        finally { Release(document); }
    }

    public string GetEquationNumberFormatDisplayName()
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            return WordEquationNumbering.GetEquationNumberFormatDisplayName(document);
        }
        finally { Release(document); }
    }

    public int SetEquationNumberFormat(string formatId)
    {
        Document? document = null;
        UndoRecord? undoRecord = null;
        Selection? selection = null;
        Range? originalSelectionRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? trackingBookmark = null;
        Range? trackedSelectionRange = null;
        Range? documentContent = null;
        var originalSelectionStart = -1;
        var originalSelectionEnd = -1;
        var originalDocumentEnd = -1;
        var trackingBookmarkName = "VTFmt_" + Guid.NewGuid().ToString("N");
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureEquationFieldResultsVisible(document);

            // Rebuilding a native #(SEQ) OMath for a heading-format change can
            // move Word.Selection into the last replacement range. Preserve the
            // user's logical caret/selection with a temporary Word bookmark so
            // positions continue tracking when prefixes earlier in the document
            // grow or shrink. A raw integer position would point into the wrong
            // paragraph after 1 becomes 1.1.
            selection = _application.Selection;
            originalSelectionRange = selection.Range.Duplicate;
            originalSelectionStart = originalSelectionRange.Start;
            originalSelectionEnd = originalSelectionRange.End;
            documentContent = document.Content;
            originalDocumentEnd = documentContent.End;
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    number-format-selection capture={originalSelectionStart}:{originalSelectionEnd} documentEnd={originalDocumentEnd}");
            try
            {
                bookmarks = document.Bookmarks;
                trackingBookmark = bookmarks.Add(
                    trackingBookmarkName,
                    originalSelectionRange);
            }
            catch
            {
                // Protected/unusual stories may reject a temporary bookmark. The
                // live duplicate Range remains a best-effort tracking fallback.
            }

            undoRecord = BeginUndoRecord("VisualTeX Set Equation Number Format");
            var visualTeXCount = WordEquationNumbering.SetEquationNumberFormat(document, formatId);
            var mathTypeCount = MathTypeEquationNumbering.SetEquationNumberFormat(document, formatId);
            return visualTeXCount + mathTypeCount;
        }
        finally
        {
            try
            {
                if (document is not null && selection is not null)
                {
                    if (bookmarks is not null && bookmarks.Exists(trackingBookmarkName))
                    {
                        Release(trackingBookmark);
                        trackingBookmark = bookmarks[trackingBookmarkName];
                        trackedSelectionRange = trackingBookmark.Range.Duplicate;
                        selection.SetRange(
                            trackedSelectionRange.Start,
                            trackedSelectionRange.End);
                        trackingBookmark.Delete();
                    }
                    else if (originalSelectionRange is not null)
                    {
                        selection.SetRange(
                            originalSelectionRange.Start,
                            originalSelectionRange.End);
                    }
                    if (string.Equals(
                            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                            "1",
                            StringComparison.Ordinal))
                    {
                        Release(documentContent);
                        documentContent = document.Content;
                        Console.WriteLine(
                            $"    number-format-selection restored={selection.Start}:{selection.End} documentEnd={documentContent.End} original={originalSelectionStart}:{originalSelectionEnd}/{originalDocumentEnd} bookmarkAlive={(bookmarks is not null && bookmarks.Exists(trackingBookmarkName))}");
                    }
                }
            }
            catch
            {
                // Selection restoration is non-destructive and must not mask the
                // actual numbering result or failure.
                try { trackingBookmark?.Delete(); } catch { }
            }

            EndUndoRecord(undoRecord);
            Release(documentContent);
            Release(trackedSelectionRange);
            Release(trackingBookmark);
            Release(bookmarks);
            Release(originalSelectionRange);
            Release(selection);
            Release(undoRecord);
            Release(document);
        }
    }

    public string ExportSelectedOleAsPicture()
    {
        var selected = ReadSelection();
        var formulaId = selected.FormulaId;
        if (string.IsNullOrWhiteSpace(formulaId))
            throw new InvalidOperationException("Please select one VisualTeX formula first.");
        var requiredFormulaId = formulaId!;

        Document? document = null;
        InlineShape? oldShape = null;
        OLEFormat? format = null;
        object? oleObject = null;
        Range? oldRange = null;
        Range? insertion = null;
        InlineShape? replacement = null;
        UndoRecord? undoRecord = null;
        string? pngPath = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            undoRecord = BeginUndoRecord("VisualTeX Export OLE Formula As Picture");
            oldShape = FindByFormulaId(document, requiredFormulaId)
                ?? throw new InvalidOperationException("The selected Word formula no longer exists.");
            var metadata = WordFormulaMetadataReader.TryRead(oldShape)
                ?? throw new InvalidDataException("The selected formula metadata is invalid.");
            format = oldShape.OLEFormat;
            if (!string.Equals(
                    format.ProgID,
                    FormulaOleContract.ProgId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The selected formula is already a picture.");
            oleObject = WordOleObjectAccessor.GetRunningObject(format);
            pngPath = OlePngPreviewExtractor.MaterializePng(oleObject, requiredFormulaId);

            var oldWidth = oldShape.Width;
            var oldHeight = oldShape.Height;
            oldRange = oldShape.Range;
            insertion = oldRange.Duplicate;
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            object link = false;
            object save = true;
            object rangeObject = insertion;
            replacement = document.InlineShapes.AddPicture(
                pngPath,
                ref link,
                ref save,
                ref rangeObject);
            Configure(
                replacement,
                metadata,
                oldWidth,
                oldHeight,
                pngPath,
                (float)(metadata.RenderHeightPx ?? 0),
                metadata.Baseline.HasValue ? (float?)metadata.Baseline.Value : null,
                metadata.DisplayMode == "inline");
            oldShape.Delete();
            TryReconcileShape(document, replacement, metadata);
            return requiredFormulaId;
        }
        catch
        {
            TryDelete(replacement);
            throw;
        }
        finally
        {
            if (pngPath is not null)
            {
                try { File.Delete(pngPath); } catch { }
            }
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(replacement);
            Release(insertion);
            Release(oldRange);
            Release(oleObject);
            Release(format);
            Release(oldShape);
            Release(document);
        }
    }

    public OfficeObjectResult Insert(OfficeSessionDocument session, string imagePath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShape? shape = null;
        UndoRecord? undoRecord = null;
        try
        {
            undoRecord = BeginUndoRecord(
                session.DisplayMode == "inline"
                    ? "VisualTeX Insert Inline Formula"
                    : "VisualTeX Insert Display Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            selection = _application.Selection;
            insertion = ResolveSessionInsertionRange(document, session, selection);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            object link = false;
            object save = true;
            object rangeObject;
            if (session.DisplayMode == "inline")
            {
                rangeObject = insertion;
                shape = document.InlineShapes.AddPicture(
                    imagePath,
                    ref link,
                    ref save,
                    ref rangeObject);
            }
            else
            {
                CompactParagraphBeforeOleDisplayFormula(document, insertion);
                var displayInsertion = ResolveDisplayInsertionRange(document, insertion);
                Release(insertion);
                insertion = displayInsertion;
                rangeObject = insertion;
                shape = document.InlineShapes.AddPicture(
                    imagePath,
                    ref link,
                    ref save,
                    ref rangeObject);
            }
            Configure(
                shape,
                metadata,
                (session.ExportResult?.Width ?? 200) * 0.75f,
                (session.ExportResult?.Height ?? 60) * 0.75f,
                imagePath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            if (session.DisplayMode == "inline")
            {
                RestoreTypingBaselineAfter(shape);
            }
            else
            {
                TryReconcileShape(document, shape, metadata);
                Range? shapeRange = null;
                try
                {
                    shapeRange = shape.Range;
                    if (session.Numbered)
                    {
                        WordEquationNumbering.CleanupNumberedDisplayInsertionSpacing(
                            document,
                            metadata.FormulaId);
                        MoveSelectionAfterNumberedDisplayFormula(
                            document,
                            selection,
                            shapeRange,
                            metadata.FormulaId);
                    }
                    else
                    {
                        MoveSelectionAfterDisplayFormula(selection, shapeRange);
                    }
                }
                finally { Release(shapeRange); }
            }
            return Result(session, document);
        }
        catch
        {
            TryDelete(shape);
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(shape);
            Release(paragraphRange);
            Release(paragraph);
            Release(insertion);
            Release(selection);
            Release(document);
        }
    }

    public OfficeObjectResult InsertOle(
        OfficeSessionDocument session,
        string pngPath,
        string emfPath,
        bool deferNumberingLayout = false,
        bool numberingScaffoldOnly = false,
        bool preserveExistingDisplayParagraphBoundary = false)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShape? shape = null;
        UndoRecord? undoRecord = null;
        try
        {
            undoRecord = BeginUndoRecord(
                session.DisplayMode == "inline"
                    ? "VisualTeX Insert Native OLE Inline Formula"
                    : "VisualTeX Insert Native OLE Display Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            if (session.Numbered)
                EnsureEquationFieldResultsVisible(document);
            selection = _application.Selection;
            insertion = ResolveSessionInsertionRange(document, session, selection);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            if (session.DisplayMode == "inline")
            {
                shape = AddOleObject(document, insertion);
            }
            else
            {
                CompactParagraphBeforeOleDisplayFormula(document, insertion);
                // VisualTeX OLE display equations use the same native Word paragraph
                // geometry as MathType: a centered tab for the formula and a right tab
                // for the generated number. Native OMML now uses the same center/right-tab paragraph geometry while remaining a Word OMath.
                var displayInsertion = ResolveDisplayInsertionRange(
                    document,
                    insertion,
                    replaceAtExactInsertion: preserveExistingDisplayParagraphBoundary);
                Release(insertion);
                insertion = displayInsertion;
                shape = AddOleObject(document, insertion);
            }
            var initialWidth = (session.ExportResult?.Width ?? 200) * 0.75f;
            var initialHeight = (session.ExportResult?.Height ?? 60) * 0.75f;
            StoreWordInlineOleSize(
                metadata,
                initialWidth,
                initialHeight,
                session.DisplayMode == "inline");
            metadata.Validate();
            InitializeOle(shape, metadata, emfPath, pngPath);
            Configure(
                shape,
                metadata,
                initialWidth,
                initialHeight,
                pngPath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            if (session.DisplayMode == "inline")
            {
                // A newly inserted OLE is still inside Word's object-creation COM
                // transaction. Restore the immediate prose/caret baseline now, but
                // do not create or rewrite the persistent VTBL typing bookmark yet.
                // The performance baseline used this path successfully; forcing the
                // anchor here later made large sequential insert workloads fail in
                // Bookmark.Delete before measurement even began.
                RestoreTypingBaselineAfter(shape);
            }
            else
            {
                if (!deferNumberingLayout)
                {
                    TryReconcileShape(
                        document,
                        shape,
                        metadata,
                        reuseExistingNumberedTableFormatting: false,
                        knownNumberedTable: null,
                        numberingScaffoldOnly: numberingScaffoldOnly);
                }
                Range? shapeRange = null;
                try
                {
                    shapeRange = shape.Range;
                    if (session.Numbered && !deferNumberingLayout)
                    {
                        WordEquationNumbering.CleanupNumberedDisplayInsertionSpacing(
                            document,
                            metadata.FormulaId);
                        MoveSelectionAfterNumberedDisplayFormula(
                            document,
                            selection,
                            shapeRange,
                            metadata.FormulaId);
                    }
                    else if (preserveExistingDisplayParagraphBoundary)
                    {
                        selection.SetRange(shapeRange.End, shapeRange.End);
                    }
                    else
                    {
                        MoveSelectionAfterDisplayFormula(selection, shapeRange);
                    }
                }
                finally { Release(shapeRange); }
            }
            // Seed a durable identity bookmark before control returns to Word.
            // This is a known-new object, so bind it directly in O(1) rather
            // than running the duplicate-copy scan used for unknown pasted OLE.
            BindOleIdentityBookmark(shape, metadata.FormulaId);
            return Result(session, document);
        }
        catch
        {
            TryDelete(shape);
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(shape);
            Release(paragraphRange);
            Release(paragraph);
            Release(insertion);
            Release(selection);
            Release(document);
        }
    }

    internal string GetMathTypeNumberPositionPreference()
    {
        var saved = WordEquationNumbering.TryGetDefaultMathTypeNumberPosition();
        if (saved is not null) return saved;

        Document? document = null;
        try
        {
            // MathType stores its own document-wide default in
            // MTEqnNumsOnRight. VisualTeX may read that value once when no
            // VisualTeX preference exists, but must never write it while a Word
            // formula transaction is active: changing MathType's global document
            // state can make the MathType add-in re-enter Word layout/field code
            // while VisualTeX is still materializing an OLE object.
            document = _application.ActiveDocument;
            return document is null
                ? "right"
                : ReadMathTypeNumberPositionPreference(document);
        }
        catch { return "right"; }
        finally { Release(document); }
    }

    internal string GetMathTypeNumberPositionForRange(string? sourceObjectId)
    {
        Document? document = null;
        InlineShape? shape = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null) return "right";
            shape = FindMathTypeOleByRange(document, sourceObjectId);
            if (shape is not null
                && MathTypeOleInterop.TryReadDisplayNumberPosition(shape, out var position))
                return position;
            return GetMathTypeNumberPositionPreference();
        }
        catch { return "right"; }
        finally
        {
            Release(shape);
            Release(document);
        }
    }

    public OfficeObjectResult InsertMathTypeOle(
        OfficeSessionDocument session,
        string mathMl,
        string? emfPath,
        string? createdObjectBookmarkName = null,
        ResolvedEquationHeadingScope? preResolvedHeadingScope = null,
        ISet<int>? preparedHeadingScopeStarts = null,
        string? isolatedNativePreviewWmfPath = null,
        float isolatedNativePreviewWidthPt = 0,
        float isolatedNativePreviewHeightPt = 0,
        int isolatedNativePreviewWordPosition = 0,
        bool isolatedNativePreviewAttempted = false,
        bool reuseExistingInlineTypingBoundary = false,
        bool updateCreatedMathTypeNumberFields = false,
        bool preserveExistingDisplayParagraphBoundary = false)
    {
        if (string.IsNullOrWhiteSpace(mathMl)
            || !mathMl.TrimStart().StartsWith("<math", StringComparison.Ordinal))
            throw new InvalidDataException(
                "VisualTeX did not provide valid MathML for MathType OLE insertion.");
        var hasIsolatedNativePreview =
            !string.IsNullOrWhiteSpace(isolatedNativePreviewWmfPath)
            && File.Exists(isolatedNativePreviewWmfPath)
            && isolatedNativePreviewWidthPt > 0
            && isolatedNativePreviewHeightPt > 0;
        if (!hasIsolatedNativePreview
            && (string.IsNullOrWhiteSpace(emfPath) || !File.Exists(emfPath)))
            throw new FileNotFoundException(
                "VisualTeX did not provide a valid MathType native or vector preview for OLE insertion.",
                emfPath);

        var metadata = session.ToMetadata();
        metadata.Validate();
        if (string.IsNullOrWhiteSpace(metadata.Latex))
            throw new InvalidDataException(
                "VisualTeX did not provide LaTeX source for MathType OLE insertion.");

        var inline = string.Equals(
            session.DisplayMode,
            "inline",
            StringComparison.OrdinalIgnoreCase);
        var generated = MathTypeMtefCodec.CreateEquationNative(mathMl, inline);
        var compoundFile = MathTypeOleStorage.CreateStandaloneCompoundFile(generated);
        var generatedMathMl = MathTypeOleStorage.ReadMathMl(compoundFile);
        var expectedSignature = MathTypeMtefCodec.SemanticSignature(mathMl);
        if (!MathTypeMathMlRoundTripMatches(expectedSignature, generatedMathMl))
        {
            var detail = string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                ? $" expected='{expectedSignature}' generated='{MathTypeMtefCodec.SemanticSignature(generatedMathMl)}' generatedMathMl='{generatedMathMl}'"
                : string.Empty;
            throw new InvalidDataException(
                $"VisualTeX generated invalid standalone MathType MTEF for '{metadata.Latex}'.{detail}");
        }

        // The visible Word presentation must use MathType's own MTEF geometry.
        // Using the frontend/MathJax EMF here made otherwise-valid Equation Native
        // objects look subtly different from equations inserted by MathType itself
        // (notably relation/operator spacing) and could size the OLE host a few
        // pixels too narrowly, clipping the right-most italic glyph. Prefer the
        // MathPage native renderer whenever it is installed; keep the frontend EMF
        // only as a compatibility fallback for machines without MathPage.
        MathTypeNativePreviewRenderer.Result? nativePreview = null;
        byte[] previewWmf;
        float widthPt;
        float heightPt;
        int wordPosition;
        var renderRoot = !string.IsNullOrWhiteSpace(emfPath)
            ? Path.GetDirectoryName(emfPath) ?? Path.GetTempPath()
            : Path.GetTempPath();
        if (hasIsolatedNativePreview)
        {
            previewWmf = File.ReadAllBytes(isolatedNativePreviewWmfPath);
            widthPt = isolatedNativePreviewWidthPt;
            heightPt = isolatedNativePreviewHeightPt;
            wordPosition = isolatedNativePreviewWordPosition;
        }
        else if (!isolatedNativePreviewAttempted
            && MathTypeNativePreviewRenderer.TryRender(
                generated.Mtef,
                renderRoot,
                out var renderedNativePreview))
        {
            nativePreview = renderedNativePreview;
            previewWmf = File.ReadAllBytes(nativePreview.WmfPath);
            widthPt = nativePreview.WidthPt;
            heightPt = nativePreview.HeightPt;
            wordPosition = nativePreview.WordPosition;
        }
        else
        {
            widthPt = (float)Math.Max(1d, (session.ExportResult?.Width ?? 200d) * 0.75d);
            heightPt = (float)Math.Max(1d, (session.ExportResult?.Height ?? 60d) * 0.75d);
            var alignToWordTextBaseline = inline || session.Numbered;
            wordPosition = alignToWordTextBaseline
                ? CalculateMathTypeOleWordPosition(
                    heightPt,
                    session.ExportResult?.Height ?? 0f,
                    session.ExportResult?.Baseline)
                : 0;
            if (string.IsNullOrWhiteSpace(emfPath) || !File.Exists(emfPath))
                throw new FileNotFoundException(
                    "MathType native preview was unavailable and no fallback EMF exists.",
                    emfPath);
            previewWmf = MathTypeWordOpenXml.ConvertEnhancedMetafileToPlaceableWmf(
                emfPath!,
                widthPt,
                heightPt);
        }

        // Keep genuine MathType storage free of OlePres. Word owns the external
        // WMF presentation in the DOCX package, avoiding the blank-object
        // regression that prompted the visibility change in 6a43aec.

        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Range? displaySpacingAnchor = null;
        InlineShape? shape = null;
        Field? sourceNumberTemplateField = null;
        Bookmark? createdObjectBookmark = null;
        Range? createdObjectBookmarkRange = null;
        UndoRecord? undoRecord = null;
        var sourceParagraphCount = -1;
        var paragraphCountBeforeDisplayPreparation = -1;
        var insertionStart = -1;
        var createdSectionBreakCodeStart = -1;
        // Batch format conversion supplies a temporary object-identity bookmark.
        // Its caller performs one document-wide numbering reconciliation after all
        // replacements, so the per-item path can avoid repeated global scans.
        var useLocalConversionLookup =
            !string.IsNullOrWhiteSpace(createdObjectBookmarkName);
        var traceInsertPerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF"),
            "1",
            StringComparison.Ordinal);
        var insertPerfWatch = traceInsertPerformance ? Stopwatch.StartNew() : null;
        long insertPerfLastMs = 0;
        void TraceInsertPerf(string perfStage)
        {
            if (insertPerfWatch is null) return;
            var totalMs = insertPerfWatch.ElapsedMilliseconds;
            WordDoubleClickHook.TraceMessage(
                $"mathtype-insert-perf stage={perfStage} numbered={session.Numbered} inline={inline} deltaMs={totalMs - insertPerfLastMs} totalMs={totalMs}");
            insertPerfLastMs = totalMs;
        }
        var stage = "initialize";
        try
        {
            undoRecord = BeginUndoRecord(
                inline
                    ? "VisualTeX Insert MathType OLE Inline Formula"
                    : session.Numbered
                        ? "VisualTeX Insert MathType OLE Numbered Display Formula"
                        : "VisualTeX Insert MathType OLE Display Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            if (!inline && session.Numbered)
                EnsureEquationFieldResultsVisible(document);
            selection = _application.Selection;

            stage = "resolve-captured-insertion";
            insertion = ResolveSessionInsertionRange(document, session, selection);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            TraceInsertPerf("resolve-insertion");

            if (!inline && session.Numbered)
            {
                stage = "repair-incomplete-number-row";
                ClearIncompleteMathTypeNumberRowAtInsertion(document, insertion);
                TraceInsertPerf("clear-incomplete-number-row");
            }

            MathTypeWordOpenXml.NumberTemplate? numberTemplate = null;
            if (!inline && session.Numbered)
            {
                stage = "resolve-number-template";
                var documentNumberFormat = EquationNumberFormat.Resolve(
                    WordEquationNumbering.GetEquationNumberFormatId(document));
                if (useLocalConversionLookup)
                {
                    // For VisualTeX/OMML -> MathType batch conversion, every
                    // numbered target follows the document-selected number format.
                    // Avoid scanning document.Fields for the nearest MTPlaceRef on
                    // every formula; the batch performs one final numbering update.
                    numberTemplate = MathTypeWordOpenXml.CreateVisualTeXNumberTemplate(
                        documentNumberFormat.Id);
                }
                else
                {
                    sourceNumberTemplateField = FindNearestMathTypePlaceRefField(
                        document,
                        insertion.Start,
                        excludeStart: -1,
                        excludeEnd: -1);
                    if (sourceNumberTemplateField is not null
                        && TryReadReusableMathTypePlaceRefTemplate(
                            document,
                            sourceNumberTemplateField,
                            out var reusableNumberTemplate))
                    {
                        numberTemplate = reusableNumberTemplate;
                    }
                    else
                    {
                        // A malformed legacy/direct-insert MTPlaceRef can retain only
                        // punctuation such as "(.)" while its nested MTEqn/MTChap
                        // fields have escaped outside the MACROBUTTON tree. Never
                        // clone that corruption into every later equation. The
                        // document-selected VisualTeX format is authoritative when
                        // no structurally complete native MathType template exists.
                        numberTemplate = MathTypeWordOpenXml.CreateVisualTeXNumberTemplate(
                            documentNumberFormat.Id);
                    }
                }

                // Number formatting and chapter/section state are two separate
                // pieces of MathType's native model.  The old code stopped here
                // whenever *any* nearby MTPlaceRef existed, so subsequent direct
                // insertions/conversions cloned the field template but never
                // established MTChap/MTSec state for their own Word heading scope.
                // Always reconcile the native heading state when this MTPlaceRef
                // actually uses chapter/section sequences.
                if (documentNumberFormat.UsesHeading
                    && numberTemplate is not null
                    && MathTypeNumberTemplateUsesHeading(numberTemplate))
                {
                    var logicalInsertionStart = insertion.Start;
                    var scopeAlreadyPrepared = preResolvedHeadingScope is not null
                        && preResolvedHeadingScope.ScopeStart != int.MinValue
                        && preparedHeadingScopeStarts?.Contains(
                            preResolvedHeadingScope.ScopeStart) == true;
                    if (!scopeAlreadyPrepared
                        && preResolvedHeadingScope?.ScopeStart != int.MinValue)
                    {
                        var insertedSectionLength = EnsureMathTypeHeadingScopeState(
                            document,
                            logicalInsertionStart,
                            documentNumberFormat,
                            out createdSectionBreakCodeStart,
                            preResolvedHeadingScope);
                        if (preResolvedHeadingScope is not null)
                            preparedHeadingScopeStarts?.Add(
                                preResolvedHeadingScope.ScopeStart);
                        if (insertedSectionLength > 0)
                        {
                            var shiftedStart = Math.Min(
                                document.Content.End,
                                logicalInsertionStart + insertedSectionLength);
                            insertion.SetRange(shiftedStart, shiftedStart);
                        }
                    }
                }
                TraceInsertPerf("number-template-heading-state");
            }

            if (!inline)
            {
                stage = "prepare-display-row";
                paragraphCountBeforeDisplayPreparation = ReadDocumentParagraphCount(document);
                displaySpacingAnchor = insertion.Duplicate;
                var displayInsertion = ResolveStandaloneMathTypeDisplayInsertionRange(
                    document,
                    insertion,
                    replaceAtExactInsertion: preserveExistingDisplayParagraphBoundary);
                Release(insertion);
                insertion = displayInsertion;
                TraceInsertPerf("prepare-display-row");
            }

            insertionStart = insertion.Start;
            stage = "build-flat-opc";
            var wordOpenXml = MathTypeWordOpenXml.CreateWithPlaceableWmf(
                compoundFile,
                previewWmf,
                widthPt,
                heightPt,
                display: !inline,
                numberTemplate,
                session.MathTypeNumberPosition);
            TraceInsertPerf("build-flat-opc");

            // Resolve the inserted OLE from the exact mutation site for every path,
            // including ordinary interactive insertion. The previous interactive
            // safety check enumerated document.InlineShapes twice and read every
            // preceding shape.Range merely to predict the new ordinal. In a document
            // with N formulas that made each new insertion O(N), and a sequence of N
            // insertions O(N²). Word inserts one OLE character at the target range;
            // a bounded local probe plus the semantic CFB validation below is both
            // stricter and independent of total document size. Keep the global
            // nearest-position scan only as an exceptional compatibility fallback.
            sourceParagraphCount = ReadDocumentParagraphCount(document);
            TraceInsertPerf("pre-insert-bookkeeping");
            stage = "insert-flat-opc";
            insertion.InsertXML(wordOpenXml);
            TraceInsertPerf("insert-flat-opc");
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                && string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE"),
                    "after-flat-opc",
                    StringComparison.Ordinal))
                throw new COMException(
                    "Injected MathType Flat OPC failure for rollback acceptance.",
                    unchecked((int)0x8007000E));
            var insertedRangeEnd = Math.Min(
                document.Content.End,
                insertionStart + 8);
            var insertedRangeReference =
                $"{RangeReferencePrefix}{insertionStart}:{insertedRangeEnd}";
            shape = FindMathTypeOleByRange(
                    document,
                    insertedRangeReference,
                    allowGlobalFallback: false)
                ?? FindMathTypeOleInParagraphAtPosition(document, insertionStart)
                ?? FindMathTypeOleInLocalWindow(document, insertionStart)
                ?? FindMathTypeOleNearPosition(document, insertionStart);
            if (shape is null)
                throw new InvalidOperationException(
                    "Word inserted the MathType OLE data but VisualTeX could not resolve the new equation.");
            if (!MathTypeOleInterop.IsMathTypeOle(shape))
                throw new InvalidOperationException(
                    "Word did not materialize the standalone equation as Equation.DSMT4.");
            TraceInsertPerf("resolve-new-shape");
            stage = "validate-flat-opc-storage";
            // Interactive creation keeps the expensive post-insert Flat OPC
            // round-trip validation. Batch format conversion already validated
            // the exact standalone CFB before InsertXML and verifies the resulting
            // object class immediately above; serializing shape.Range.WordOpenXML
            // again for every converted equation costs several seconds per OLE in
            // large documents and turns a 50-formula conversion into minutes.
            if (!useLocalConversionLookup
                && MathTypeOleStorage.TryCaptureCompoundFileFromWordOpenXml(
                    shape,
                    out var materializedCompoundFile))
            {
                var materializedMathMl = MathTypeOleStorage.ReadMathMl(
                    materializedCompoundFile);
                if (!MathTypeMathMlRoundTripMatches(expectedSignature, materializedMathMl))
                    throw new InvalidDataException(
                        "Word materialized a different MathType equation than VisualTeX generated.");
            }

            // The Flat OPC already contains both VisualTeX's standalone MathType
            // CFB and its WMF/EMF presentation cache.  Do not round-trip the object
            // through Word PasteSpecial here: wdPasteOLEObject asks Windows to
            // instantiate the Equation.DSMT4 CLSID and therefore launches an
            // installed MathType OLE server.  Keeping this original InlineShape is
            // both sufficient for Word display and makes conversion fully offline.
            stage = "use-flat-opc-presentation";
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            shape.Width = widthPt;
            shape.Height = heightPt;
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            TraceInsertPerf("set-shape-geometry");

            if (inline && sourceParagraphCount >= 0)
                RepairMathTypeInsertXmlParagraphSplit(
                    document,
                    shape,
                    sourceParagraphCount);
            TraceInsertPerf("repair-inline-paragraph");

            // Word structural edits can move the InlineShape range. Re-resolve it
            // before the final baseline/numbering work without activating its OLE
            // server.
            stage = "refresh-native-shape";
            var refreshedRangeEnd = Math.Min(
                document.Content.End,
                insertionStart + 8);
            var refreshedShape = FindMathTypeOleByRange(
                    document,
                    $"{RangeReferencePrefix}{insertionStart}:{refreshedRangeEnd}",
                    allowGlobalFallback: false)
                ?? FindMathTypeOleInParagraphAtPosition(document, insertionStart)
                ?? FindMathTypeOleInLocalWindow(document, insertionStart)
                ?? FindMathTypeOleNearPosition(document, insertionStart);
            if (refreshedShape is null)
                throw new InvalidOperationException(
                    "Word retained the MathType OLE but VisualTeX could not refresh its live InlineShape handle.");
            Release(shape);
            shape = refreshedShape;
            TraceInsertPerf("refresh-native-shape");

            if (inline)
            {
                stage = "apply-native-baseline";
                SetInlineOleWordPosition(shape, wordPosition);
                TraceInsertPerf("set-native-baseline");
                if (!reuseExistingInlineTypingBoundary)
                {
                    if (useLocalConversionLookup)
                        RestoreTypingBaselineAfterMathTypeConversion(shape);
                    else
                        RestoreTypingBaselineAfter(shape);
                }
                TraceInsertPerf("restore-typing-baseline");
                var shapeRange = shape.Range;
                try { selection.SetRange(shapeRange.End, shapeRange.End); }
                finally { Release(shapeRange); }
            }
            else
            {
                stage = "configure-display-numbering";
                ConfigureNewMathTypeDisplayEquation(
                    document,
                    shape,
                    session.Numbered,
                    session.MathTypeNumberPosition,
                    updateNestedNumberFields:
                        !useLocalConversionLookup || updateCreatedMathTypeNumberFields);
                // MTDisplayEquation setup resets direct character formatting on
                // the OLE run. Apply the exported math baseline after numbering
                // and paragraph style are final so the adjacent number stays
                // vertically aligned with tall display formulas.
                stage = "apply-native-display-baseline";
                SetInlineOleWordPosition(shape, wordPosition);
                var shapeRange = shape.Range;
                try
                {
                    if (session.Numbered)
                        selection.SetRange(shapeRange.Start, shapeRange.End);
                    else if (preserveExistingDisplayParagraphBoundary)
                        selection.SetRange(shapeRange.End, shapeRange.End);
                    else
                        MoveSelectionAfterDisplayFormula(selection, shapeRange);
                }
                finally { Release(shapeRange); }
            }

            if (!inline && displaySpacingAnchor is not null)
            {
                stage = "finalize-display-spacing";
                CompactParagraphBeforeOleDisplayFormula(document, displaySpacingAnchor);
            }
            if (!string.IsNullOrWhiteSpace(createdObjectBookmarkName))
            {
                stage = "bind-created-object-identity";
                TryDeleteBookmark(document, createdObjectBookmarkName);
                createdObjectBookmarkRange = shape.Range;
                createdObjectBookmark = document.Bookmarks.Add(
                    createdObjectBookmarkName,
                    createdObjectBookmarkRange);
                TraceInsertPerf("bind-created-object-identity");
            }
            stage = "complete";
            TraceInsertPerf("complete");
            return Result(session, document);
        }
        catch (Exception error)
        {
            if (document is not null && !string.IsNullOrWhiteSpace(createdObjectBookmarkName))
                TryDeleteBookmark(document, createdObjectBookmarkName);
            if (!inline && document is not null && insertionStart >= 0)
            {
                RollbackStandaloneMathTypeDisplayInsertion(
                    document,
                    insertionStart,
                    paragraphCountBeforeDisplayPreparation);
            }
            else
            {
                TryDelete(shape);
            }
            if (createdSectionBreakCodeStart >= 0 && document is not null)
                RemoveMathTypeSectionBreakFieldAtCodeStart(
                    document,
                    createdSectionBreakCodeStart);
            var hresult = error is COMException
                ? $" HRESULT=0x{error.HResult:X8}."
                : string.Empty;
            throw new InvalidOperationException(
                $"MathType OLE insertion failed at stage '{stage}'.{hresult} {error.Message}",
                error);
        }
        finally
        {
            nativePreview?.Dispose();
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(createdObjectBookmarkRange);
            Release(createdObjectBookmark);
            Release(sourceNumberTemplateField);
            Release(shape);
            Release(displaySpacingAnchor);
            Release(insertion);
            Release(selection);
            Release(document);
        }
    }

    // Legacy server-backed writer kept temporarily for focused diagnostics only.
    // Production insertion/conversion must never activate the MathType UI.
    private OfficeObjectResult InsertMathTypeOleWithLiveServer(
        OfficeSessionDocument session,
        string mathMl,
        string emfPath,
        MathTypeWordOpenXml.NumberTemplate? explicitNumberTemplate = null,
        bool numberingSectionStatePrepared = false,
        int? expectedInlineShapeIndex = null,
        bool useNativePreview = true,
        string? isolatedNativePreviewWmfPath = null,
        float isolatedNativePreviewWidthPt = 0,
        float isolatedNativePreviewHeightPt = 0,
        int isolatedNativePreviewWordPosition = 0,
        string? createdObjectBookmarkName = null,
        bool replaceAtExactInsertion = false)
    {
        if (string.IsNullOrWhiteSpace(mathMl)
            || !mathMl.TrimStart().StartsWith("<math", StringComparison.Ordinal))
            throw new InvalidDataException(
                "VisualTeX did not provide valid MathML for MathType OLE insertion.");
        if (string.IsNullOrWhiteSpace(emfPath) || !File.Exists(emfPath))
            throw new FileNotFoundException(
                "VisualTeX did not provide a valid vector preview for MathType OLE insertion.",
                emfPath);

        var metadata = session.ToMetadata();
        metadata.Validate();
        if (string.IsNullOrWhiteSpace(metadata.Latex))
            throw new InvalidDataException(
                "VisualTeX did not provide LaTeX source for MathType OLE insertion.");

        var inline = string.Equals(
            session.DisplayMode,
            "inline",
            StringComparison.OrdinalIgnoreCase);
        var expectedSignature = MathTypeMtefCodec.SemanticSignature(mathMl);

        // The final Word object is now a genuine Equation.DSMT4 created by Word and
        // populated through MathType's official server API. The old standalone MTEF/
        // CFB writer is therefore no longer a correctness gate. Keep it only as an
        // optional native-preview optimization for direct insertion; unsupported
        // structures must fall back to the renderer/session geometry instead of
        // rejecting a formula that MathType itself can represent correctly.
        MathTypeMtefCodec.RewriteResult? generated = null;
        if (useNativePreview
            && (string.IsNullOrWhiteSpace(isolatedNativePreviewWmfPath)
                || !File.Exists(isolatedNativePreviewWmfPath)))
        {
            try { generated = MathTypeMtefCodec.CreateEquationNative(mathMl, inline); }
            catch (Exception previewPreparationError)
            {
                if (string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                {
                    Console.WriteLine(
                        "    [MathType preview] offline MTEF preview preparation unavailable; "
                        + "falling back to session geometry: "
                        + previewPreparationError.Message);
                }
            }
        }

        MathTypeNativePreviewRenderer.Result? nativePreview = null;
        float widthPt;
        float heightPt;
        var wordPosition = 0;
        var renderRoot = Path.GetDirectoryName(emfPath) ?? Path.GetTempPath();
        if (!string.IsNullOrWhiteSpace(isolatedNativePreviewWmfPath)
            && File.Exists(isolatedNativePreviewWmfPath)
            && isolatedNativePreviewWidthPt > 0
            && isolatedNativePreviewHeightPt > 0)
        {
            // Format conversion obtains this geometry from MathType's own MathPage
            // renderer in a sidecar process, so no synthetic OLE storage is needed.
            widthPt = isolatedNativePreviewWidthPt;
            heightPt = isolatedNativePreviewHeightPt;
            wordPosition = isolatedNativePreviewWordPosition;
        }
        else if (useNativePreview
            && generated is not null
            && MathTypeNativePreviewRenderer.TryRender(
                generated.Mtef,
                renderRoot,
                out var renderedNativePreview))
        {
            nativePreview = renderedNativePreview;
            widthPt = nativePreview.WidthPt;
            heightPt = nativePreview.HeightPt;
            wordPosition = nativePreview.WordPosition;
        }
        else
        {
            widthPt = (float)Math.Max(1d, (session.ExportResult?.Width ?? 200d) * 0.75d);
            heightPt = (float)Math.Max(1d, (session.ExportResult?.Height ?? 60d) * 0.75d);
        }

        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Range? displaySpacingAnchor = null;
        InlineShape? shape = null;
        Bookmark? createdObjectBookmark = null;
        Range? createdObjectBookmarkRange = null;
        Field? sourceNumberTemplateField = null;
        UndoRecord? undoRecord = null;
        var sourceParagraphCount = -1;
        var paragraphCountBeforeDisplayPreparation = -1;
        var insertionStart = -1;
        var createdSectionBreakCodeStart = -1;
        var stage = "initialize";
        try
        {
            undoRecord = BeginUndoRecord(
                inline
                    ? "VisualTeX Insert MathType OLE Inline Formula"
                    : session.Numbered
                        ? "VisualTeX Insert MathType OLE Numbered Display Formula"
                        : "VisualTeX Insert MathType OLE Display Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            selection = _application.Selection;

            // Resolve the captured Session insertion point before consulting any
            // MathType numbering state. The live Word Selection can move while the
            // external editor owns focus and must never decide which MTPlaceRef
            // template a create operation inherits.
            stage = "resolve-captured-insertion";
            insertion = ResolveSessionInsertionRange(document, session, selection);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);

            // Older failed builds could leave a number-only MTDisplayEquation row
            // at the exact captured caret. Clear only that structurally incomplete
            // row so a retry does not clone its orphan MTPlaceRef or increment the
            // native MathType sequence twice.
            if (!inline && session.Numbered)
            {
                stage = "repair-incomplete-number-row";
                ClearIncompleteMathTypeNumberRowAtInsertion(document, insertion);
            }

            MathTypeWordOpenXml.NumberTemplate? numberTemplate = explicitNumberTemplate;
            if (!inline && session.Numbered && numberTemplate is null)
            {
                stage = "resolve-number-template";
                sourceNumberTemplateField = FindNearestMathTypePlaceRefField(
                    document,
                    insertion.Start,
                    excludeStart: -1,
                    excludeEnd: -1);
                if (sourceNumberTemplateField is not null)
                    numberTemplate = ReadMathTypePlaceRefTemplate(
                        document,
                        sourceNumberTemplateField);
                else
                {
                    var documentNumberFormat = EquationNumberFormat.Resolve(
                        WordEquationNumbering.GetEquationNumberFormatId(document));
                    numberTemplate = MathTypeWordOpenXml.CreateVisualTeXNumberTemplate(
                        documentNumberFormat.Id);
                    if (!numberingSectionStatePrepared
                        && documentNumberFormat.UsesHeading)
                    {
                        var logicalInsertionStart = insertion.Start;
                        var insertedSectionLength = EnsureMathTypeHeadingScopeState(
                            document,
                            logicalInsertionStart,
                            documentNumberFormat,
                            out createdSectionBreakCodeStart);
                        if (insertedSectionLength > 0)
                        {
                            // The hidden native section state is inserted immediately
                            // after the real Word heading that owns this equation.
                            // Shift the captured logical formula position by the exact
                            // amount Word inserted before it.
                            var shiftedStart = Math.Min(
                                document.Content.End,
                                logicalInsertionStart + insertedSectionLength);
                            insertion.SetRange(shiftedStart, shiftedStart);
                        }
                    }
                }
            }

            if (!inline)
            {
                stage = "prepare-display-row";
                paragraphCountBeforeDisplayPreparation = ReadDocumentParagraphCount(document);
                displaySpacingAnchor = insertion.Duplicate;
                var displayInsertion = ResolveStandaloneMathTypeDisplayInsertionRange(
                    document,
                    insertion,
                    replaceAtExactInsertion);
                Release(insertion);
                insertion = displayInsertion;
            }

            insertionStart = insertion.Start;
            var sourceObjectCount = document.InlineShapes.Count;
            var expectedNewShapeIndex = expectedInlineShapeIndex is > 0
                ? expectedInlineShapeIndex.Value
                : ResolveInlineShapeInsertionIndex(
                    document,
                    insertionStart);
            sourceParagraphCount = ReadDocumentParagraphCount(document);

            // Create the final MathType client item directly. Do not insert a
            // synthetic Equation.DSMT4 Flat OPC object first: Word retains OLE host
            // state from that temporary object, and later conversions can invalidate
            // the live preview of earlier equations in the same batch.
            stage = "insert-independent-mathtype-ole";
            shape = document.InlineShapes.AddOLEObject(
                ClassType: "Equation.DSMT4",
                LinkToFile: false,
                DisplayAsIcon: false,
                Range: insertion);
            if (document.InlineShapes.Count != sourceObjectCount + 1)
                throw new InvalidOperationException(
                    "Independent MathType OLE insertion changed the expected Word OLE object count.");
            if (!MathTypeOleInterop.IsMathTypeOle(shape))
                throw new InvalidDataException(
                    "Independent MathType OLE insertion did not produce Equation.DSMT4.");
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                Range? createdRange = null;
                try
                {
                    createdRange = shape.Range;
                    string createdSignature;
                    try
                    {
                        createdSignature = MathTypeMtefCodec.SemanticSignature(
                            MathTypeOleStorage.ReadMathMl(shape));
                    }
                    catch (Exception createdReadError)
                    {
                        createdSignature = "<unreadable:" + createdReadError.Message + ">";
                    }
                    Console.WriteLine(
                        $"    [MathType just-created] expected-index={expectedNewShapeIndex}; range=[{createdRange.Start},{createdRange.End}]; signature='{createdSignature}'");
                }
                finally { Release(createdRange); }
            }

            if (!inline)
            {
                stage = "build-mathtype-field-scaffold";
                BuildIndependentMathTypeDisplayScaffold(
                    document,
                    shape,
                    session.Numbered,
                    session.MathTypeNumberPosition,
                    numberTemplate);
                // Inserting tabs/fields around the OLE changes Word ranges. Reacquire
                // the object by its stable document-order index before server write.
                Release(shape);
                shape = FindMathTypeOleAtIndex(document, expectedNewShapeIndex)
                    ?? throw new InvalidOperationException(
                        "Word lost the independent MathType OLE while building its display scaffold.");
            }

            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                && string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE"),
                    "after-flat-opc",
                    StringComparison.Ordinal))
                throw new COMException(
                    "Injected MathType insertion failure for rollback acceptance.",
                    unchecked((int)0x8007000E));

            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                Range? preWriteRange = null;
                try
                {
                    preWriteRange = shape.Range;
                    var preWriteMathMl = MathTypeOleStorage.ReadMathMl(shape);
                    Console.WriteLine(
                        $"    [MathType pre-write] expected-index={expectedNewShapeIndex}; range=[{preWriteRange.Start},{preWriteRange.End}]; "
                        + $"signature='{MathTypeMtefCodec.SemanticSignature(preWriteMathMl)}'; mathMl='{preWriteMathMl}'");
                }
                catch (Exception preWriteError)
                {
                    Console.WriteLine("    [MathType pre-write] diagnostic failed: " + preWriteError.Message);
                }
                finally { Release(preWriteRange); }
            }

            stage = "write-official-mathtype-mathml";
            MathTypeOleInterop.WriteMathMl(shape, mathMl);

            // IDataObject.SetData closes the running OLE object with NOSAVE after
            // MathType has updated the Word client item. Reacquire the final shape
            // before authoritative readback because the client item may be refreshed.
            Release(shape);
            shape = FindMathTypeOleAtIndex(document, expectedNewShapeIndex)
                ?? FindMathTypeOleNearPosition(document, insertionStart)
                ?? throw new InvalidOperationException(
                    "Word lost the MathType OLE after the official IDataObject MathML write.");
            var officialServerMathMl = MathTypeOleInterop.ReadMathMl(shape);
            var officialServerSignature = MathTypeMtefCodec.SemanticSignature(officialServerMathMl);
            if (!string.Equals(expectedSignature, officialServerSignature, StringComparison.Ordinal))
                throw new InvalidDataException(
                    string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal)
                        ? "MathType's live OLE server changed the converted formula before close. "
                            + $"expectedSignature='{expectedSignature}'; serverSignature='{officialServerSignature}'; "
                            + $"difference={DescribeSemanticSignatureDifference(expectedSignature, officialServerSignature)}; inputMathMl='{mathMl}'; serverMathMl='{officialServerMathMl}'"
                        : "MathType's live OLE server changed the converted formula before close.");

            // The readback also closes the running MathType OLE object. Reacquire
            // once more and verify the same saved Word client item can be read again.
            Release(shape);
            shape = FindMathTypeOleAtIndex(document, expectedNewShapeIndex)
                ?? FindMathTypeOleNearPosition(document, insertionStart)
                ?? throw new InvalidOperationException(
                    "Word lost the MathType OLE after the official MathML readback.");

            stage = "validate-official-mathtype-ole";
            var postCloseServerMathMl = MathTypeOleInterop.ReadMathMl(shape);
            var postCloseServerSignature = MathTypeMtefCodec.SemanticSignature(postCloseServerMathMl);
            if (!string.Equals(expectedSignature, postCloseServerSignature, StringComparison.Ordinal))
                throw new InvalidDataException(
                    string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal)
                        ? "MathType's official server changed the equation after close/reopen. "
                            + $"expectedSignature='{expectedSignature}'; beforeCloseSignature='{officialServerSignature}'; afterCloseSignature='{postCloseServerSignature}'; "
                            + $"difference={DescribeSemanticSignatureDifference(expectedSignature, postCloseServerSignature)}; afterCloseMathMl='{postCloseServerMathMl}'"
                        : "MathType changed the converted equation after its OLE object was closed and reopened.");

            // MTCloseOleObject can refresh the Word client item again. Always use a
            // fresh RCW after the authoritative post-close server readback.
            Release(shape);
            shape = FindMathTypeOleAtIndex(document, expectedNewShapeIndex)
                ?? FindMathTypeOleNearPosition(document, insertionStart)
                ?? throw new InvalidOperationException(
                    "Word lost the MathType OLE after the authoritative post-close readback.");

            // Keep the offline MTEF decoder as a diagnostic cross-check only. The
            // live MathType server is authoritative: MathType 7 can encode operator
            // glyphs with private typeface codes that an incomplete decoder may
            // temporarily misidentify even though the saved Equation.DSMT4 is valid.
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                try
                {
                    var nativeReadbackMathMl = MathTypeOleStorage.ReadMathMl(shape);
                    var nativeReadbackSignature = MathTypeMtefCodec.SemanticSignature(nativeReadbackMathMl);
                    if (!string.Equals(expectedSignature, nativeReadbackSignature, StringComparison.Ordinal))
                        Console.WriteLine(
                            "    [MathType decoder] official server readback is correct but offline Equation Native decoding differs: "
                            + $"difference={DescribeSemanticSignatureDifference(expectedSignature, nativeReadbackSignature)}; "
                            + $"nativeMathMl='{nativeReadbackMathMl}'");
                }
                catch (Exception decoderError)
                {
                    Console.WriteLine("    [MathType decoder] offline Equation Native diagnostic failed: " + decoderError.Message);
                }
            }

            if (!TryMeasureInlineOlePreview(shape).HasValue)
                throw new InvalidDataException(
                    "MathType's official OLE server returned a blank visible equation preview.");

            if (inline && sourceParagraphCount >= 0)
                RepairMathTypeInsertXmlParagraphSplit(
                    document,
                    shape,
                    sourceParagraphCount);

            stage = "apply-native-baseline";
            SetInlineOleWordPosition(shape, wordPosition);
            if (inline)
            {
                RestoreTypingBaselineAfter(shape);
                var shapeRange = shape.Range;
                try { selection.SetRange(shapeRange.End, shapeRange.End); }
                finally { Release(shapeRange); }
            }
            else
            {
                stage = "configure-display-numbering";
                ConfigureNewMathTypeDisplayEquation(
                    document,
                    shape,
                    session.Numbered,
                    session.MathTypeNumberPosition);
                var shapeRange = shape.Range;
                try
                {
                    if (session.Numbered)
                        selection.SetRange(shapeRange.Start, shapeRange.End);
                    else
                        MoveSelectionAfterDisplayFormula(selection, shapeRange);
                }
                finally { Release(shapeRange); }
            }

            if (!inline && displaySpacingAnchor is not null)
            {
                stage = "finalize-display-spacing";
                CompactParagraphBeforeOleDisplayFormula(document, displaySpacingAnchor);
            }

            if (!string.IsNullOrWhiteSpace(createdObjectBookmarkName))
            {
                stage = "bind-created-object-identity";
                TryDeleteBookmark(document, createdObjectBookmarkName!);
                createdObjectBookmarkRange = shape.Range;
                createdObjectBookmark = document.Bookmarks.Add(
                    createdObjectBookmarkName!,
                    createdObjectBookmarkRange);
            }
            stage = "complete";
            return Result(session, document);
        }
        catch (Exception error)
        {
            if (!inline && document is not null && insertionStart >= 0)
            {
                RollbackStandaloneMathTypeDisplayInsertion(
                    document,
                    insertionStart,
                    paragraphCountBeforeDisplayPreparation);
            }
            else
            {
                TryDelete(shape);
            }
            if (createdSectionBreakCodeStart >= 0 && document is not null)
                RemoveMathTypeSectionBreakFieldAtCodeStart(
                    document,
                    createdSectionBreakCodeStart);
            var hresult = error is COMException
                ? $" HRESULT=0x{error.HResult:X8}."
                : string.Empty;
            throw new InvalidOperationException(
                $"MathType OLE insertion failed at stage '{stage}'.{hresult} {error.Message}",
                error);
        }
        finally
        {
            nativePreview?.Dispose();
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(createdObjectBookmarkRange);
            Release(createdObjectBookmark);
            Release(sourceNumberTemplateField);
            Release(shape);
            Release(displaySpacingAnchor);
            Release(insertion);
            Release(selection);
            Release(document);
        }
    }

    private OfficeObjectResult InsertMathTypeOleWithoutActivation(
        OfficeSessionDocument session,
        string mathMl,
        string emfPath,
        MathTypeWordOpenXml.NumberTemplate? explicitNumberTemplate,
        bool numberingSectionStatePrepared,
        int? expectedInlineShapeIndex,
        bool useNativePreview,
        string? isolatedNativePreviewWmfPath,
        float isolatedNativePreviewWidthPt,
        float isolatedNativePreviewHeightPt,
        int isolatedNativePreviewWordPosition,
        string? createdObjectBookmarkName,
        bool replaceAtExactInsertion)
    {
        if (string.IsNullOrWhiteSpace(mathMl)
            || !mathMl.TrimStart().StartsWith("<math", StringComparison.Ordinal))
            throw new InvalidDataException(
                "VisualTeX did not provide valid MathML for MathType OLE insertion.");
        if (string.IsNullOrWhiteSpace(emfPath) || !File.Exists(emfPath))
            throw new FileNotFoundException(
                "VisualTeX did not provide a valid vector preview for MathType OLE insertion.",
                emfPath);

        var metadata = session.ToMetadata();
        metadata.Validate();
        if (string.IsNullOrWhiteSpace(metadata.Latex))
            throw new InvalidDataException(
                "VisualTeX did not provide LaTeX source for MathType OLE insertion.");

        var inline = string.Equals(
            session.DisplayMode,
            "inline",
            StringComparison.OrdinalIgnoreCase);

        // This is intentionally the same standalone Equation.DSMT4 construction
        // used by VisualTeX's direct MathType insertion. It does not call DoVerb,
        // OLEFormat.Object, MathPage MTSetEqnFromLangStr, or open a MathType window.
        var generated = MathTypeMtefCodec.CreateEquationNative(mathMl, inline);
        var compoundFile = MathTypeOleStorage.CreateStandaloneCompoundFile(generated);
        var generatedMathMl = MathTypeOleStorage.ReadMathMl(compoundFile);
        var expectedSignature = MathTypeMtefCodec.SemanticSignature(mathMl);
        if (!MathTypeMathMlRoundTripMatches(expectedSignature, generatedMathMl))
        {
            var detail = string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                ? $" expected='{expectedSignature}' generated='{MathTypeMtefCodec.SemanticSignature(generatedMathMl)}' generatedMathMl='{generatedMathMl}'"
                : string.Empty;
            throw new InvalidDataException(
                $"VisualTeX generated invalid standalone MathType MTEF for '{metadata.Latex}'.{detail}");
        }

        MathTypeNativePreviewRenderer.Result? nativePreview = null;
        byte[] previewWmf;
        float widthPt;
        float heightPt;
        var wordPosition = 0;
        var renderRoot = Path.GetDirectoryName(emfPath) ?? Path.GetTempPath();
        var presentationEmfPath = emfPath;
        var ownsPresentationEmfPath = false;
        try
        {
            if (!string.IsNullOrWhiteSpace(isolatedNativePreviewWmfPath)
                && File.Exists(isolatedNativePreviewWmfPath)
                && isolatedNativePreviewWidthPt > 0
                && isolatedNativePreviewHeightPt > 0)
            {
                previewWmf = File.ReadAllBytes(isolatedNativePreviewWmfPath!);
                widthPt = isolatedNativePreviewWidthPt;
                heightPt = isolatedNativePreviewHeightPt;
                wordPosition = isolatedNativePreviewWordPosition;
                presentationEmfPath = MathTypeWordOpenXml.ConvertPlaceableWmfToEnhancedMetafile(
                    isolatedNativePreviewWmfPath!,
                    widthPt,
                    heightPt,
                    renderRoot);
                ownsPresentationEmfPath = true;
            }
            else if (useNativePreview
                && MathTypeNativePreviewRenderer.TryRender(
                    generated.Mtef,
                    renderRoot,
                    out var renderedNativePreview))
            {
                nativePreview = renderedNativePreview;
                previewWmf = File.ReadAllBytes(nativePreview.WmfPath);
                widthPt = nativePreview.WidthPt;
                heightPt = nativePreview.HeightPt;
                wordPosition = nativePreview.WordPosition;
                presentationEmfPath = MathTypeWordOpenXml.ConvertPlaceableWmfToEnhancedMetafile(
                    nativePreview.WmfPath,
                    widthPt,
                    heightPt,
                    renderRoot);
                ownsPresentationEmfPath = true;
            }
            else
            {
                widthPt = (float)Math.Max(1d, (session.ExportResult?.Width ?? 200d) * 0.75d);
                heightPt = (float)Math.Max(1d, (session.ExportResult?.Height ?? 60d) * 0.75d);
                previewWmf = MathTypeWordOpenXml.ConvertEnhancedMetafileToPlaceableWmf(
                    emfPath,
                    widthPt,
                    heightPt);
            }

            // Do not add a Windows OlePres stream to Equation.DSMT4. Genuine
            // MathType storage carries Equation Native while Word owns the external
            // WMF presentation; an injected OlePres stream can override that visible
            // preview with an incompatible blank cache.
        }
        catch
        {
            nativePreview?.Dispose();
            nativePreview = null;
            if (ownsPresentationEmfPath)
            {
                try { File.Delete(presentationEmfPath); } catch { }
            }
            throw;
        }

        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Range? displaySpacingAnchor = null;
        InlineShape? shape = null;
        Bookmark? createdObjectBookmark = null;
        Range? createdObjectBookmarkRange = null;
        Field? sourceNumberTemplateField = null;
        UndoRecord? undoRecord = null;
        var sourceParagraphCount = -1;
        var paragraphCountBeforeDisplayPreparation = -1;
        var insertionStart = -1;
        var createdSectionBreakCodeStart = -1;
        var stage = "initialize";
        try
        {
            undoRecord = BeginUndoRecord(
                inline
                    ? "VisualTeX Insert MathType OLE Inline Formula"
                    : session.Numbered
                        ? "VisualTeX Insert MathType OLE Numbered Display Formula"
                        : "VisualTeX Insert MathType OLE Display Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            selection = _application.Selection;

            stage = "resolve-captured-insertion";
            insertion = ResolveSessionInsertionRange(document, session, selection);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);

            if (!inline && session.Numbered)
            {
                stage = "repair-incomplete-number-row";
                ClearIncompleteMathTypeNumberRowAtInsertion(document, insertion);
            }

            MathTypeWordOpenXml.NumberTemplate? numberTemplate = explicitNumberTemplate;
            if (!inline && session.Numbered && numberTemplate is null)
            {
                stage = "resolve-number-template";
                sourceNumberTemplateField = FindNearestMathTypePlaceRefField(
                    document,
                    insertion.Start,
                    excludeStart: -1,
                    excludeEnd: -1);
                if (sourceNumberTemplateField is not null)
                    numberTemplate = ReadMathTypePlaceRefTemplate(
                        document,
                        sourceNumberTemplateField);
                else
                {
                    var documentNumberFormat = EquationNumberFormat.Resolve(
                        WordEquationNumbering.GetEquationNumberFormatId(document));
                    numberTemplate = MathTypeWordOpenXml.CreateVisualTeXNumberTemplate(
                        documentNumberFormat.Id);
                    if (!numberingSectionStatePrepared
                        && documentNumberFormat.UsesHeading)
                    {
                        var logicalInsertionStart = insertion.Start;
                        var insertedSectionLength = EnsureMathTypeHeadingScopeState(
                            document,
                            logicalInsertionStart,
                            documentNumberFormat,
                            out createdSectionBreakCodeStart);
                        if (insertedSectionLength > 0)
                        {
                            var shiftedStart = Math.Min(
                                document.Content.End,
                                logicalInsertionStart + insertedSectionLength);
                            insertion.SetRange(shiftedStart, shiftedStart);
                        }
                    }
                }
            }

            if (!inline)
            {
                stage = "prepare-display-row";
                paragraphCountBeforeDisplayPreparation = ReadDocumentParagraphCount(document);
                displaySpacingAnchor = insertion.Duplicate;
                var displayInsertion = ResolveStandaloneMathTypeDisplayInsertionRange(
                    document,
                    insertion,
                    replaceAtExactInsertion);
                Release(insertion);
                insertion = displayInsertion;
            }

            insertionStart = insertion.Start;
            stage = "build-flat-opc";
            var wordOpenXml = MathTypeWordOpenXml.CreateWithPlaceableWmf(
                compoundFile,
                previewWmf,
                widthPt,
                heightPt,
                display: !inline,
                numberTemplate,
                session.MathTypeNumberPosition);

            var sourceObjectCount = document.InlineShapes.Count;
            var insertionOrdinal = ResolveInlineShapeInsertionIndex(document, insertionStart);
            sourceParagraphCount = ReadDocumentParagraphCount(document);
            stage = "insert-flat-opc";
            insertion.InsertXML(wordOpenXml);
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                && string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE"),
                    "after-flat-opc",
                    StringComparison.Ordinal))
                throw new COMException(
                    "Injected MathType Flat OPC failure for rollback acceptance.",
                    unchecked((int)0x8007000E));
            if (document.InlineShapes.Count != sourceObjectCount + 1)
                throw new InvalidOperationException(
                    "Word did not materialize exactly one standalone MathType OLE equation.");

            // Never require the source formula's old InlineShapes ordinal here.
            // Deleting a numbered table can legitimately change that ordinal in a
            // real document. Resolve the new object from the insertion position first.
            shape = FindMathTypeOleByRange(
                    document,
                    $"{RangeReferencePrefix}{insertionStart}:{insertionStart + 2}")
                ?? FindMathTypeOleInParagraphAtPosition(document, insertionStart)
                ?? FindMathTypeOleNearPosition(document, insertionStart)
                ?? FindMathTypeOleAtIndex(document, insertionOrdinal)
                ?? throw new InvalidOperationException(
                    "Word inserted the MathType OLE data but VisualTeX could not resolve the new equation at its insertion position.");
            if (!MathTypeOleInterop.IsMathTypeOle(shape))
                throw new InvalidOperationException(
                    "Word did not materialize the standalone equation as Equation.DSMT4.");

            stage = "validate-flat-opc-storage";
            if (MathTypeOleStorage.TryCaptureCompoundFileFromWordOpenXml(
                    shape,
                    out var materializedCompoundFile))
            {
                var materializedMathMl = MathTypeOleStorage.ReadMathMl(materializedCompoundFile);
                if (!MathTypeMathMlRoundTripMatches(expectedSignature, materializedMathMl))
                    throw new InvalidDataException(
                        "Word materialized a different MathType equation than VisualTeX generated.");
            }

            stage = "materialize-native-ole-presentation";
            var presentedShape = RematerializeStandaloneMathTypeOlePresentation(
                document,
                shape,
                compoundFile,
                presentationEmfPath,
                widthPt,
                heightPt);
            Release(shape);
            shape = presentedShape;

            stage = "validate-native-ole-presentation";
            if (MathTypeOleStorage.TryCaptureCompoundFileFromWordOpenXml(
                    shape,
                    out var presentedCompoundFile))
            {
                // Semantic validation already passed before the presentation-only
                // PasteSpecial round trip. Re-read only to verify that Word kept a
                // valid MathType CFB; private MathType glyphs can decode as U+FFFD.
                _ = MathTypeOleStorage.ReadMathMl(presentedCompoundFile);
            }
            if (!TryMeasureInlineOlePreview(shape).HasValue)
                throw new InvalidDataException(
                    "Word materialized the MathType OLE without a visible preview.");

            if (inline && sourceParagraphCount >= 0)
                RepairMathTypeInsertXmlParagraphSplit(
                    document,
                    shape,
                    sourceParagraphCount);

            stage = "apply-native-baseline";
            SetInlineOleWordPosition(shape, wordPosition);
            if (inline)
            {
                RestoreTypingBaselineAfter(shape);
                var shapeRange = shape.Range;
                try { selection.SetRange(shapeRange.End, shapeRange.End); }
                finally { Release(shapeRange); }
            }
            else
            {
                stage = "configure-display-numbering";
                ConfigureNewMathTypeDisplayEquation(
                    document,
                    shape,
                    session.Numbered,
                    session.MathTypeNumberPosition);
                var shapeRange = shape.Range;
                try
                {
                    if (session.Numbered)
                        selection.SetRange(shapeRange.Start, shapeRange.End);
                    else
                        MoveSelectionAfterDisplayFormula(selection, shapeRange);
                }
                finally { Release(shapeRange); }
            }

            if (!inline && displaySpacingAnchor is not null)
            {
                stage = "finalize-display-spacing";
                CompactParagraphBeforeOleDisplayFormula(document, displaySpacingAnchor);
            }

            if (!string.IsNullOrWhiteSpace(createdObjectBookmarkName))
            {
                stage = "bind-created-object-identity";
                TryDeleteBookmark(document, createdObjectBookmarkName!);
                createdObjectBookmarkRange = shape.Range;
                createdObjectBookmark = document.Bookmarks.Add(
                    createdObjectBookmarkName!,
                    createdObjectBookmarkRange);
            }

            _ = expectedInlineShapeIndex; // compatibility-only; identity is range/bookmark based.
            stage = "complete";
            return Result(session, document);
        }
        catch (Exception error)
        {
            if (!inline && document is not null && insertionStart >= 0)
            {
                RollbackStandaloneMathTypeDisplayInsertion(
                    document,
                    insertionStart,
                    paragraphCountBeforeDisplayPreparation);
            }
            else
            {
                TryDelete(shape);
            }
            if (createdSectionBreakCodeStart >= 0 && document is not null)
                RemoveMathTypeSectionBreakFieldAtCodeStart(
                    document,
                    createdSectionBreakCodeStart);
            var hresult = error is COMException
                ? $" HRESULT=0x{error.HResult:X8}."
                : string.Empty;
            throw new InvalidOperationException(
                $"MathType OLE insertion failed at stage '{stage}'.{hresult} {error.Message}",
                error);
        }
        finally
        {
            nativePreview?.Dispose();
            if (ownsPresentationEmfPath)
            {
                try { File.Delete(presentationEmfPath); } catch { }
            }
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(createdObjectBookmarkRange);
            Release(createdObjectBookmark);
            Release(sourceNumberTemplateField);
            Release(shape);
            Release(displaySpacingAnchor);
            Release(insertion);
            Release(selection);
            Release(document);
        }
    }

    private static InlineShape RematerializeStandaloneMathTypeOlePresentation(
        Document document,
        InlineShape sourceShape,
        byte[] compoundFile,
        string emfPath,
        float widthPt,
        float heightPt)
    {
        Range? sourceRange = null;
        Range? insertion = null;
        InlineShape? replacement = null;
        try
        {
            sourceRange = sourceShape.Range;
            var start = sourceRange.Start;
            var sourceCount = document.InlineShapes.Count;
            Exception? lastPasteError = null;
            var pasted = false;
            for (var attempt = 1; attempt <= 3 && !pasted; attempt++)
            {
                try
                {
                    using var transaction = MathTypeOleStorage.BeginStandaloneClipboardTransaction(
                        compoundFile,
                        emfPath,
                        widthPt,
                        heightPt);
                    Release(insertion);
                    insertion = document.Range(start, start);
                    insertion.PasteSpecial(
                        Link: false,
                        DataType: WdPasteDataType.wdPasteOLEObject,
                        Placement: WdOLEPlacement.wdInLine,
                        DisplayAsIcon: false);
                    if (document.InlineShapes.Count != sourceCount + 1)
                        throw new InvalidOperationException(
                            "Word did not create exactly one replacement MathType OLE presentation.");
                    if (transaction.ReplacementStorageWriteCount <= 0)
                        throw new InvalidOperationException(
                            "Word created an OLE object without requesting the VisualTeX MathType CFB storage.");
                    pasted = true;
                }
                catch (COMException error) when (attempt < 3)
                {
                    lastPasteError = error;
                    if (document.InlineShapes.Count != sourceCount)
                        throw new InvalidOperationException(
                            "Word partially inserted a MathType OLE while presentation materialization failed; VisualTeX refused an unsafe retry.",
                            error);
                    Thread.Sleep(60 * attempt);
                }
            }
            if (!pasted)
                throw new InvalidOperationException(
                    "Word repeatedly refused to materialize the standalone MathType OLE presentation.",
                    lastPasteError);

            sourceShape.Delete();
            if (document.InlineShapes.Count != sourceCount)
                throw new InvalidOperationException(
                    "Word changed the MathType OLE object count while materializing its presentation.");
            replacement = FindMathTypeOleNearPosition(document, start)
                ?? throw new InvalidOperationException(
                    "Word pasted the MathType OLE presentation but VisualTeX could not resolve the replacement object.");
            if (!MathTypeOleInterop.IsMathTypeOle(replacement))
                throw new InvalidDataException(
                    "Word changed the standalone MathType OLE class while materializing its presentation.");

            replacement.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            replacement.Width = widthPt;
            replacement.Height = heightPt;
            replacement.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            var result = replacement;
            replacement = null;
            return result;
        }
        finally
        {
            Release(replacement);
            Release(insertion);
            Release(sourceRange);
        }
    }

    public OfficeObjectResult InsertOmml(
        OfficeSessionDocument session,
        string mathMl,
        bool deferNumberingLayout = false,
        bool numberingScaffoldOnly = false,
        bool deferFinalFingerprint = false,
        WordOmmlConverter.BatchSource? ommlBatchSource = null,
        bool preserveExistingDisplayParagraphBoundary = false,
        bool normalizeMathTypeDisplayParagraph = false)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Range? equationRange = null;
        var preservedDisplayParagraphStart = -1;
        var preservedDisplayParagraphCount = -1;
        string sourceFingerprint = string.Empty;
        Bookmark? bookmark = null;
        UndoRecord? undoRecord = null;
        InlineFollowingTextVisibility? inlineFollowingTextVisibility = null;
        var metadataSaved = false;
        var preparedNumberedDisplayHost = false;
        var performanceWatch = Stopwatch.StartNew();
        long performanceCheckpoint = 0;
        try
        {
            undoRecord = BeginUndoRecord(
                session.DisplayMode == "inline"
                    ? "VisualTeX Insert Word OMML Inline Formula"
                    : "VisualTeX Insert Word OMML Display Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            if (session.Numbered)
                EnsureEquationFieldResultsVisible(document);
            ApplyDocumentOmmlMathFont(document, metadata);
            selection = _application.Selection;
            insertion = ResolveSessionInsertionRange(document, session, selection);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            // The paragraph after the last numbered display is intentionally kept
            // as a 1pt typing tail so it does not render as an extra blank line.
            // When that exact tail is reused for the next formula, restore Normal
            // paragraph metrics before materializing OMML; otherwise a genuine
            // wdOMathDisplay inherits the 1pt line height, visually overlaps the
            // previous formula and eventually destabilizes Word in long runs.
            selection.SetRange(insertion.Start, insertion.Start);
            WordEquationNumbering.ExpandCompactTrailingTypingParagraph(selection);
            if (session.DisplayMode == "inline")
            {
                inlineFollowingTextVisibility =
                    CaptureInlineFollowingTextVisibility(insertion);
                var placeholder = PrepareInlineBaselineSentinelBeforeInsert(
                    document,
                    insertion,
                    metadata.FormulaId);
                Release(insertion);
                insertion = placeholder;
                equationRange = ommlBatchSource is not null
                    ? ommlBatchSource.Insert(
                        document,
                        insertion,
                        metadata.FormulaId,
                        display: false,
                        sourceFingerprint: out sourceFingerprint,
                        replaceTarget: true)
                    : WordOmmlConverter.Insert(
                        _application,
                        document,
                        insertion,
                        mathMl,
                        display: false,
                        sourceFingerprint: out sourceFingerprint,
                        replaceTarget: true,
                        transformOmml: omml => ApplyOmmlTypographyXml(
                            omml,
                            session.FontSizePt,
                            metadata),
                        mathFontName: document.OMathFontName);
            }
            else
            {
                // Every block OMML formula, numbered or not, is materialized as
                // genuine Word display math from the first native insertion. A
                // numbered formula is then replaced atomically with Word's native
                // #(SEQ VisualTeXEquation) OMath structure in the same paragraph;
                // there is no Shape/TextBox or extra anchor paragraph.
                if (preserveExistingDisplayParagraphBoundary)
                {
                    preservedDisplayParagraphCount = ReadDocumentParagraphCount(document);
                    preservedDisplayParagraphStart = ReadParagraphStart(insertion);
                }
                var displayInsertion = ResolveDisplayInsertionRange(document, insertion);
                Release(insertion);
                insertion = displayInsertion;
                equationRange = ommlBatchSource is not null
                    ? ommlBatchSource.Insert(
                        document,
                        insertion,
                        metadata.FormulaId,
                        display: true,
                        sourceFingerprint: out sourceFingerprint,
                        replaceTarget: preparedNumberedDisplayHost)
                    : WordOmmlConverter.Insert(
                        _application,
                        document,
                        insertion,
                        mathMl,
                        display: true,
                        sourceFingerprint: out sourceFingerprint,
                        replaceTarget: preparedNumberedDisplayHost,
                        transformOmml: omml => ApplyOmmlTypographyXml(
                            omml,
                            session.FontSizePt,
                            metadata),
                        mathFontName: document.OMathFontName);
            }

            if (normalizeMathTypeDisplayParagraph
                && string.Equals(
                    session.DisplayMode,
                    "block",
                    StringComparison.Ordinal))
            {
                NormalizeConvertedOmmlParagraphFromMathTypeStyle(equationRange);
                TraceAcceptancePerformance(
                    "InsertOmml",
                    "detach-mathtype-style",
                    performanceWatch,
                    ref performanceCheckpoint);
            }
            TraceAcceptancePerformance(
                "InsertOmml",
                "insert-native-omml",
                performanceWatch,
                ref performanceCheckpoint);
            ApplyOmmlTypography(equationRange, session.FontSizePt, metadata);
            metadata.NativeOmmlFingerprint = sourceFingerprint;
            // The bookmark is needed by inline boundary cleanup, but metadata is
            // intentionally not persisted yet: Word can still normalize the OMath
            // while the surrounding layout is finalized.
            bookmark = WordOmmlFormulaStore.Wrap(document, equationRange, metadata);
            if (session.DisplayMode == "inline")
            {
                FinalizeInlineOmmlBoundary(
                    document,
                    equationRange,
                    metadata.FormulaId,
                    moveCaretOutsideMath: true,
                    followingTextVisibility: inlineFollowingTextVisibility);
            }
            else
            {
                if (!deferNumberingLayout)
                {
                    TryReconcileOmml(
                        document,
                        bookmark!,
                        equationRange,
                        metadata,
                        reuseExistingNumberedTableFormatting: false,
                        knownNumberedTable: null,
                        numberingScaffoldOnly: numberingScaffoldOnly,
                        deferNativeOmmlShapeFinalization: false,
                        deferNativeOmmlShapeCreation: false);
                    TraceAcceptancePerformance(
                        "InsertOmml",
                        "reconcile",
                        performanceWatch,
                        ref performanceCheckpoint);

                    if (session.Numbered)
                    {
                        // Numbered OMML reconciliation can move the equation from its
                        // original body paragraph into the center cell of a newly
                        // created 1x3 host and replace the collapsed VTOMML bookmark.
                        // Never keep using the pre-reconcile RCWs after that structural
                        // move. On a freshly started Word instance the old Range can
                        // serialize as an empty/non-math paragraph; the final
                        // fingerprint then fails and the old rollback path can delete
                        // the live center formula while leaving the right-side SEQ.
                        Release(equationRange);
                        equationRange = null;
                        Release(bookmark);
                        bookmark = null;
                        bookmark = WordOmmlFormulaStore.FindByFormulaId(
                                document,
                                metadata.FormulaId)
                            ?? throw new InvalidOperationException(
                                "Word lost the newly inserted numbered OMML bookmark after building its 1x3 host.");
                        equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                        var liveMaths = equationRange.OMaths;
                        try
                        {
                            if (liveMaths.Count != 1)
                                throw new InvalidOperationException(
                                    "Word lost the newly inserted center OMath after building its 1x3 host.");
                        }
                        finally { Release(liveMaths); }
                    }
                }
                if (session.Numbered && !deferNumberingLayout)
                {
                    WordEquationNumbering.CleanupNumberedDisplayInsertionSpacing(
                        document,
                        metadata.FormulaId);
                    TraceAcceptancePerformance(
                        "InsertOmml",
                        "cleanup-numbered-spacing",
                        performanceWatch,
                        ref performanceCheckpoint);
                    MoveSelectionAfterNumberedDisplayFormula(
                        document,
                        selection,
                        equationRange,
                        metadata.FormulaId);
                    TraceAcceptancePerformance(
                        "InsertOmml",
                        "move-numbered-caret",
                        performanceWatch,
                        ref performanceCheckpoint);
                }
                else if (preserveExistingDisplayParagraphBoundary)
                {
                    // Display OMML import may split the empty source paragraph to
                    // either side of the new OMath depending on Word build/layout.
                    // Restore the exact pre-conversion paragraph count using the
                    // source paragraph's captured start instead of assuming the
                    // residual blank paragraph is always after the formula.
                    RepairPreservedDisplayParagraphBoundary(
                        document,
                        equationRange,
                        preservedDisplayParagraphStart,
                        preservedDisplayParagraphCount);
                    selection.SetRange(equationRange.End, equationRange.End);
                }
                else
                {
                    MoveSelectionAfterDisplayFormula(selection, equationRange);
                }
            }
            TraceAcceptancePerformance(
                "InsertOmml",
                "post-reconcile-layout",
                performanceWatch,
                ref performanceCheckpoint);

            // Word can normalize the imported OMML again while typography,
            // inline-boundary cleanup, or numbered layout is finalized. Ordinary
            // single-formula operations persist that final native fingerprint
            // immediately. A format-conversion batch defers it: serializing one
            // equation through Range.WordOpenXML becomes document-sized work as
            // the batch grows, so the caller refreshes every new VTOMML formula
            // from one document-wide WordOpenXML snapshot after all inserts.
            if (session.Numbered
                && string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                && string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_VSTO_OMML_FAIL_STAGE"),
                    "after-numbered-insert-reconcile",
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "Injected failure after numbered OMML insertion reconciliation.");
            if (!deferFinalFingerprint)
            {
                try
                {
                    WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                        metadata,
                        equationRange);
                }
                catch (InvalidDataException error) when (
                    session.Numbered
                    && string.Equals(
                        session.DisplayMode,
                        "block",
                        StringComparison.Ordinal)
                    && WordEquationNumbering
                        .HasReusableNumberedNativeOmmlDirectTableHost(
                            document,
                            equationRange,
                            metadata.FormulaId))
                {
                    // A freshly started Word process can transiently serialize the
                    // just-built center OMath Range without its <m:oMath> wrapper on
                    // the first numbered insertion, even though the live 1x3 host is
                    // already structurally complete (one Display OMath in cell 2,
                    // one direct SEQ in cell 3, all FormulaId bookmarks intact).
                    // The semantic fingerprint was computed from the source OMML
                    // before insertion and is still authoritative here. Never turn
                    // this Word serialization race into a destructive insert failure;
                    // genuine host corruption does not satisfy the structural guard
                    // above and therefore still propagates the exception.
                    metadata.NativeOmmlFingerprint = sourceFingerprint;
                    WordDoubleClickHook.TraceMessage(
                        $"numbered-omml-final-fingerprint-transient-fallback formulaId={metadata.FormulaId} error={error.Message}");
                }
            }
            TraceAcceptancePerformance(
                "InsertOmml",
                "final-fingerprint",
                performanceWatch,
                ref performanceCheckpoint);
            if (!WordOmmlFormulaStore.IsCanonicalAnchor(bookmark, equationRange))
            {
                Release(bookmark);
                bookmark = WordOmmlFormulaStore.Wrap(
                    document,
                    equationRange,
                    metadata,
                    replaceExisting: true);
            }
            TraceAcceptancePerformance(
                "InsertOmml",
                "final-anchor",
                performanceWatch,
                ref performanceCheckpoint);
            WordOmmlFormulaStore.SaveNew(document, metadata);
            metadataSaved = true;
            TraceAcceptancePerformance(
                "InsertOmml",
                "final-metadata",
                performanceWatch,
                ref performanceCheckpoint);

            var result = Result(session, document);
            var finalizeNumberedOmmlNumberOutsideTransaction =
                session.Numbered
                && string.Equals(
                    session.DisplayMode,
                    "block",
                    StringComparison.Ordinal)
                && !deferNumberingLayout
                && !numberingScaffoldOnly;
            var directTableAlreadyComplete =
                finalizeNumberedOmmlNumberOutsideTransaction
                && WordEquationNumbering
                    .HasReusableNumberedNativeOmmlDirectTableHost(
                        document,
                        equationRange,
                        metadata.FormulaId);
            if (finalizeNumberedOmmlNumberOutsideTransaction
                && !directTableAlreadyComplete)
            {
                // Only retired/migration structures require a compatibility pass
                // after the insertion transaction. A current direct-SEQ 1x3 host
                // is already fully built and validated; re-enumerating it through
                // the old Shape-named finalizer adds visible latency for no effect.
                EndUndoRecord(undoRecord);
                Release(undoRecord);
                undoRecord = null;
                Release(bookmark);
                bookmark = null;
                Release(equationRange);
                equationRange = null;
                Release(insertion);
                insertion = null;
                Release(selection);
                selection = null;

                var finalizedNumbers =
                    WordEquationNumbering.FinalizeConvertedNumberedOmmlDisplayShapes(
                        document,
                        new[] { metadata.FormulaId });
                if (finalizedNumbers != 1
                    && string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "Word did not finalize the native numbered-OMML #(SEQ) host after insertion.");
                TraceAcceptancePerformance(
                    "InsertOmml",
                    "native-number-finalize",
                    performanceWatch,
                    ref performanceCheckpoint);
            }
            return result;
        }
        catch
        {
            // Never delete through a pre-reconcile bookmark/Range. Numbered OMML
            // host construction can replace both objects while preserving the same
            // FormulaId; deleting through a stale RCW is exactly how a failed first
            // insertion could leave a right-cell SEQ with an empty center cell.
            // Re-resolve the current FormulaId-bound equation first and delete only
            // that proven live OMath. If it cannot be resolved, leave document
            // content intact rather than guessing at a stale range.
            Bookmark? rollbackBookmark = null;
            Range? rollbackRange = null;
            Table? rollbackTable = null;
            try
            {
                if (document is not null)
                {
                    if (session.Numbered
                        && string.Equals(
                            session.DisplayMode,
                            "block",
                            StringComparison.Ordinal))
                    {
                        rollbackTable = WordEquationNumbering.FindNumberedEquationTable(
                            document,
                            metadata.FormulaId);
                        if (rollbackTable is not null)
                        {
                            try { rollbackTable.Delete(); } catch { }
                            Release(rollbackTable);
                            rollbackTable = null;
                        }
                    }

                    // If no complete numbered table could be resolved (for example
                    // failure happened before the scaffold existed), remove only a
                    // freshly re-resolved FormulaId-bound OMath. Never act on the
                    // stale pre-reconcile bookmark/Range held by the caller.
                    rollbackBookmark = WordOmmlFormulaStore.FindByFormulaId(
                        document,
                        metadata.FormulaId);
                    if (rollbackBookmark is not null)
                    {
                        try { rollbackRange = WordOmmlFormulaStore.GetEquationRange(rollbackBookmark); }
                        catch { rollbackRange = null; }
                        try { rollbackBookmark.Delete(); } catch { }
                        if (rollbackRange is not null)
                        {
                            try { rollbackRange.Delete(); } catch { }
                        }
                    }

                    if (session.Numbered)
                    {
                        try
                        {
                            WordEquationNumbering.RemoveFormulaNumberingArtifacts(
                                document,
                                metadata.FormulaId);
                        }
                        catch { }
                    }
                }
            }
            finally
            {
                Release(rollbackTable);
                Release(rollbackRange);
                Release(rollbackBookmark);
            }
            if (metadataSaved && document is not null)
            {
                try { WordOmmlFormulaStore.Delete(document, metadata.FormulaId); } catch { }
            }
            if (preparedNumberedDisplayHost && document is not null)
            {
                try
                {
                    WordEquationNumbering.RemoveFormulaNumberingArtifacts(
                        document,
                        metadata.FormulaId);
                }
                catch { }
            }
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(bookmark);
            Release(equationRange);
            Release(insertion);
            Release(selection);
            Release(document);
        }
    }

    public WordLatexRedrawPlan CaptureLatexRedrawPlan(bool wholeDocument)
    {
        Document? document = null;
        Selection? selection = null;
        Range? scope = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            selection = _application.Selection;
            scope = wholeDocument
                ? document.Content.Duplicate
                : selection.Range.Duplicate;
            if (!wholeDocument && scope.Start == scope.End)
                throw new InvalidOperationException("请先选择包含 LaTeX 代码的 Word 内容。");

            var sourceText = scope.Text ?? string.Empty;
            var spans = WordBulkImportParser.FindFormulaSpans(sourceText);
            if (spans.Count == 0)
                throw new InvalidDataException(
                    wholeDocument
                        ? "当前 Word 文档中没有找到 $...$、$$...$$、\\(...\\) 或 \\[...\\] 公式。"
                        : "所选内容中没有找到 $...$、$$...$$、\\(...\\) 或 \\[...\\] 公式。");

            var plan = new WordLatexRedrawPlan
            {
                DocumentId = DocumentIdentity(document),
                ScopeStart = scope.Start,
                ScopeEnd = scope.End,
                SourceText = sourceText,
                Targets = spans.Select(span => new WordLatexRedrawTarget
                {
                    Id = span.Id,
                    RelativeStart = span.Start,
                    SourceLength = span.Length,
                    Latex = span.Latex,
                    DisplayMode = span.DisplayMode,
                }).ToList(),
            };

            // Resolve the exact Word story ranges before rendering. Word story
            // coordinates can diverge from .NET UTF-16 offsets after supplementary
            // Unicode characters. The exact ranges are also required to inherit the
            // surrounding prose size instead of the deliberately smaller LaTeX source
            // run used by many generated documents.
            var sourceContexts = BuildLatexRedrawSourceContexts(
                plan.SourceText,
                plan.Targets);
            var sourceFontSizes = TryBuildWordOpenXmlFontSizeIndex(
                scope,
                plan.SourceText);
            var wordStoryOffsets = BuildWordStoryOffsetIndex(plan.SourceText);
            var resolvedForFormatting = new List<ResolvedLatexRedrawTarget>(plan.Targets.Count);
            long rangeResolutionTicks = 0;
            long fontResolutionTicks = 0;
            var openXmlFontHitCount = 0;
            var contextComFontHitCount = 0;
            var fontFallbackCount = 0;
            try
            {
                foreach (var target in plan.Targets.OrderBy(item => item.RelativeStart))
                {
                    var expectedSource = plan.SourceText.Substring(
                        target.RelativeStart,
                        target.SourceLength);
                    var targetRelativeEnd = target.RelativeStart + target.SourceLength;
                    var sourceStart = plan.ScopeStart
                        + wordStoryOffsets[target.RelativeStart];
                    var sourceEnd = plan.ScopeStart
                        + wordStoryOffsets[targetRelativeEnd];
                    target.AbsoluteStart = sourceStart;
                    target.AbsoluteEnd = sourceEnd;
                    Range? sourceRange = null;
                    if (sourceStart < plan.ScopeStart
                        || sourceEnd <= sourceStart
                        || sourceEnd > plan.ScopeEnd)
                    {
                        var rangeStarted = Stopwatch.GetTimestamp();
                        sourceRange = ResolveExactLatexSourceRange(
                            document,
                            plan,
                            target,
                            expectedSource,
                            resolvedForFormatting);
                        rangeResolutionTicks += Stopwatch.GetTimestamp() - rangeStarted;
                        sourceStart = target.AbsoluteStart;
                        sourceEnd = target.AbsoluteEnd;
                    }

                    var display = string.Equals(
                        target.DisplayMode,
                        "block",
                        StringComparison.Ordinal);
                    var sourceContext = sourceContexts[target.Id];
                    target.PreserveDisplayParagraphBoundary =
                        display && !sourceContext.HasVisibleSurroundingText;
                    var fontStarted = Stopwatch.GetTimestamp();
                    var fontContextPosition = sourceContext.FontContextRelativePosition;
                    if (sourceFontSizes is not null
                        && fontContextPosition >= 0
                        && fontContextPosition < sourceFontSizes.Length
                        && sourceFontSizes[fontContextPosition] is double openXmlFontSize
                        && openXmlFontSize > 0)
                    {
                        openXmlFontHitCount++;
                        target.FontSizePt = openXmlFontSize;
                    }
                    else if (TryResolveSourceFormulaFontSizeFromContext(
                                 document,
                                 plan,
                                 target,
                                 sourceStart,
                                 sourceEnd,
                                 fontContextPosition,
                                 out var contextFontSize))
                    {
                        contextComFontHitCount++;
                        target.FontSizePt = contextFontSize;
                    }
                    else
                    {
                        fontFallbackCount++;
                        if (sourceRange is null)
                        {
                            var rangeStarted = Stopwatch.GetTimestamp();
                            sourceRange = ResolveExactLatexSourceRange(
                                document,
                                plan,
                                target,
                                expectedSource,
                                resolvedForFormatting);
                            rangeResolutionTicks += Stopwatch.GetTimestamp() - rangeStarted;
                            sourceStart = target.AbsoluteStart;
                            sourceEnd = target.AbsoluteEnd;
                        }
                        target.FontSizePt = ResolveSourceFormulaFontSize(
                            document,
                            sourceRange,
                            display);
                    }
                    fontResolutionTicks += Stopwatch.GetTimestamp() - fontStarted;
                    resolvedForFormatting.Add(new ResolvedLatexRedrawTarget
                    {
                        Target = target,
                        SourceRange = sourceRange!,
                        SourceStart = sourceStart,
                        SourceEnd = sourceEnd,
                        ExpectedSource = expectedSource,
                    });
                }
                if (string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                {
                    var ticksToMilliseconds = 1000d / Stopwatch.Frequency;
                    Console.WriteLine(
                        "    [perf] latex-redraw-capture: "
                        + $"range={rangeResolutionTicks * ticksToMilliseconds:F1}ms, "
                        + $"font={fontResolutionTicks * ticksToMilliseconds:F1}ms, "
                        + $"openXmlFontHits={openXmlFontHitCount}, "
                        + $"contextComFontHits={contextComFontHitCount}, "
                        + $"fontFallbacks={fontFallbackCount}/{plan.Targets.Count}.");
                }
            }
            finally
            {
                foreach (var resolved in resolvedForFormatting)
                    Release(resolved.SourceRange);
            }

            return plan;
        }
        finally
        {
            Release(scope);
            Release(selection);
            Release(document);
        }
    }

    public WordLatexRedrawResult ApplyLatexRedrawPlan(
        WordLatexRedrawPlan plan,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));
        if (prepared is null) throw new ArgumentNullException(nameof(prepared));
        Document? document = null;
        Selection? selection = null;
        Range? validationRange = null;
        UndoRecord? undoRecord = null;
        WordViewState? viewState = null;
        List<ResolvedLatexRedrawTarget>? resolvedTargets = null;
        var insertedFormulaIds = new List<string>();
        var totalInsertMilliseconds = 0L;
        var maxInsertMilliseconds = 0L;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, plan.DocumentId);
            validationRange = document.Range(plan.ScopeStart, plan.ScopeEnd);
            if (!string.Equals(validationRange.Text ?? string.Empty, plan.SourceText, StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "渲染期间 Word 内容发生了变化。为避免替换错误位置，本次重绘已停止，请重新选择后再试。");

            selection = _application.Selection;
            viewState = CaptureViewState();
            // Resolve and validate every Word Range before changing the document.
            // Word story coordinates are not always a one-to-one mapping of .NET
            // UTF-16 string offsets (for example after supplementary Unicode
            // characters, tracked revisions or hidden story markers). Keeping the
            // resolved live ranges also prevents a late locator failure after some
            // formulas have already been replaced.
            resolvedTargets = ResolveLatexRedrawTargets(document, plan, prepared);
            undoRecord = BeginUndoRecord("VisualTeX 重绘 LaTeX 公式");
            foreach (var resolved in resolvedTargets
                         .OrderByDescending(item => item.SourceRange.Start))
            {
                Range? preservedDisplayParagraphRange = null;
                string? preservedFollowingParagraphText = null;
                try
                {
                    var targetRange = resolved.SourceRange;
                    if (!string.Equals(
                            targetRange.Text ?? string.Empty,
                            resolved.ExpectedSource,
                            StringComparison.Ordinal))
                        throw new InvalidOperationException(
                            $"渲染期间公式内容发生变化：{resolved.Target.Latex}");

                    if (resolved.Target.PreserveDisplayParagraphBoundary)
                    {
                        preservedDisplayParagraphRange =
                            DuplicateContainingParagraphRange(targetRange);
                        Range? followingParagraphRange = null;
                        try
                        {
                            followingParagraphRange = DuplicateFollowingParagraphRange(
                                document,
                                preservedDisplayParagraphRange);
                            preservedFollowingParagraphText = followingParagraphRange?.Text;
                        }
                        finally { Release(followingParagraphRange); }
                    }

                    selection.SetRange(targetRange.Start, targetRange.End);
                    selection.Text = string.Empty;
                    selection.Collapse(WdCollapseDirection.wdCollapseStart);
                    var stopwatch = Stopwatch.StartNew();
                    InsertPreparedFormula(
                        document,
                        selection,
                        resolved.Formula,
                        display: string.Equals(
                            resolved.Target.DisplayMode,
                            "block",
                            StringComparison.Ordinal),
                        preserveExistingDisplayParagraphBoundary:
                            resolved.Target.PreserveDisplayParagraphBoundary,
                        preservedDisplayParagraphRange:
                            preservedDisplayParagraphRange,
                        preservedFollowingParagraphText:
                            preservedFollowingParagraphText);
                    stopwatch.Stop();
                    totalInsertMilliseconds += stopwatch.ElapsedMilliseconds;
                    maxInsertMilliseconds = Math.Max(
                        maxInsertMilliseconds,
                        stopwatch.ElapsedMilliseconds);
                    insertedFormulaIds.Add(resolved.Formula.Session.FormulaId);
                }
                finally { Release(preservedDisplayParagraphRange); }
            }

            return new WordLatexRedrawResult
            {
                FormulaCount = insertedFormulaIds.Count,
                TotalInsertMilliseconds = totalInsertMilliseconds,
                MaxInsertMilliseconds = maxInsertMilliseconds,
                FormulaIds = insertedFormulaIds,
            };
        }
        finally
        {
            EndUndoRecord(undoRecord);
            RestoreViewState(document, viewState, preferredSelection: null);
            Release(undoRecord);
            if (resolvedTargets is not null)
            {
                foreach (var resolved in resolvedTargets)
                    Release(resolved.SourceRange);
            }
            Release(validationRange);
            Release(selection);
            Release(document);
        }
    }

    public int CountFormulaObjectsForLatex(bool wholeDocument, string objectMode)
    {
        Document? document = null;
        Selection? selection = null;
        Range? scope = null;
        List<FormulaToLatexTarget>? targets = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            selection = _application.Selection;
            scope = wholeDocument
                ? document.Content.Duplicate
                : selection.Range.Duplicate;
            targets = CaptureFormulaToLatexTargets(
                document,
                scope,
                wholeDocument,
                objectMode,
                refreshOmmlMetadata: false);
            return targets.Count;
        }
        finally
        {
            ReleaseFormulaToLatexTargets(targets);
            Release(scope);
            Release(selection);
            Release(document);
        }
    }

    public WordFormulaToLatexResult ConvertFormulaObjectsToLatex(
        bool wholeDocument,
        string objectMode)
    {
        Document? document = null;
        Selection? selection = null;
        Range? scope = null;
        UndoRecord? undoRecord = null;
        WordViewState? viewState = null;
        List<FormulaToLatexTarget>? targets = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            selection = _application.Selection;
            scope = wholeDocument
                ? document.Content.Duplicate
                : selection.Range.Duplicate;
            targets = CaptureFormulaToLatexTargets(
                document,
                scope,
                wholeDocument,
                objectMode,
                refreshOmmlMetadata: true);
            if (targets.Count == 0)
            {
                var modeLabel = string.Equals(
                    objectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal)
                    ? "VisualTeX OLE"
                    : string.Equals(
                        objectMode,
                        FormulaOleContract.MathTypeOleMode,
                        StringComparison.Ordinal)
                        ? "MathType"
                        : "Word OMML";
                throw new InvalidDataException(
                    wholeDocument
                        ? $"当前 Word 文档中没有找到可转换的 {modeLabel} 公式。"
                        : $"所选内容中没有找到可转换的 {modeLabel} 公式。");
            }

            // Preflight every source before opening a destructive Word undo record.
            // Corrupt/empty metadata must never turn a visible formula into an empty
            // replacement merely because the object itself can still be located.
            foreach (var target in targets)
            {
                ApplyFormulaToLatexEmptySourceInjection(target);
                target.LatexSource = BuildFormulaLatexSource(target.Metadata);
            }

            viewState = CaptureViewState();
            undoRecord = BeginUndoRecord("VisualTeX 公式转为 LaTeX 代码");
            if (undoRecord is null)
                throw new InvalidOperationException(
                    "Word 无法建立公式转 LaTeX 的撤销事务。为避免公式丢失，本次转换未开始。");

            var undoRecordEnded = false;
            var documentMutationStarted = false;
            try
            {
                var convertedIds = new List<string>(targets.Count);
                foreach (var target in targets.OrderByDescending(item => item.Start))
                {
                    ConvertFormulaTargetToLatex(
                        document,
                        target,
                        ref documentMutationStarted);
                    convertedIds.Add(target.Metadata.FormulaId);
                }
                WordEquationNumbering.TryReconcile(document);
                return new WordFormulaToLatexResult
                {
                    FormulaCount = convertedIds.Count,
                    FormulaIds = convertedIds,
                };
            }
            catch (Exception conversionError)
            {
                EndUndoRecord(undoRecord);
                undoRecordEnded = true;
                if (documentMutationStarted
                    && !TryUndoFormulaToLatexConversion(document))
                    throw new InvalidOperationException(
                        "公式转 LaTeX 失败，而且 Word 无法自动撤销本次转换。请立即使用 Ctrl+Z，并保留当前文档以便排查。",
                        conversionError);
                throw;
            }
            finally
            {
                if (!undoRecordEnded)
                    EndUndoRecord(undoRecord);
            }
        }
        finally
        {
            RestoreViewState(document, viewState, preferredSelection: null);
            Release(undoRecord);
            ReleaseFormulaToLatexTargets(targets);
            Release(scope);
            Release(selection);
            Release(document);
        }
    }

    private static List<FormulaToLatexTarget> CaptureFormulaToLatexTargets(
        Document document,
        Range scope,
        bool wholeDocument,
        string objectMode,
        bool refreshOmmlMetadata)
    {
        if (!string.Equals(
                objectMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal)
            && !string.Equals(
                objectMode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal)
            && !string.Equals(
                objectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal))
            throw new ArgumentOutOfRangeException(
                nameof(objectMode),
                objectMode,
                "Only VisualTeX OLE, Word OMML and MathType formulas can be restored to LaTeX code.");

        var targets = new List<FormulaToLatexTarget>();
        try
        {
            if (string.Equals(
                    objectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                InlineShapes? shapes = null;
                try
                {
                    shapes = document.InlineShapes;
                    for (var index = 1; index <= shapes.Count; index++)
                    {
                        InlineShape? shape = null;
                        Range? formulaRange = null;
                        try
                        {
                            shape = shapes[index];
                            if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                            var metadata = WordFormulaMetadataReader.TryRead(shape);
                            if (metadata is null) continue;
                            formulaRange = shape.Range;
                            if (!FormulaRangeMatchesScope(
                                    formulaRange,
                                    scope,
                                    wholeDocument))
                                continue;
                            targets.Add(new FormulaToLatexTarget
                            {
                                Metadata = metadata,
                                ObjectMode = FormulaOleContract.NativeOleMode,
                                Start = formulaRange.Start,
                                End = formulaRange.End,
                                FormulaRange = formulaRange,
                                OleShape = shape,
                            });
                            formulaRange = null;
                            shape = null;
                        }
                        finally
                        {
                            Release(formulaRange);
                            Release(shape);
                        }
                    }
                }
                finally { Release(shapes); }
                return targets;
            }

            if (string.Equals(
                    objectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal))
            {
                InlineShapes? shapes = null;
                try
                {
                    shapes = document.InlineShapes;
                    for (var index = 1; index <= shapes.Count; index++)
                    {
                        InlineShape? shape = null;
                        Range? formulaRange = null;
                        try
                        {
                            shape = shapes[index];
                            if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                            formulaRange = shape.Range;
                            if (!FormulaRangeMatchesScope(
                                    formulaRange,
                                    scope,
                                    wholeDocument))
                                continue;
                            var sourceMathMl = MathTypeOleStorage.ReadMathMl(shape);
                            var metadata = MathTypeOleInterop.ReadMetadata(
                                document.Application,
                                shape,
                                sourceMathMl);
                            targets.Add(new FormulaToLatexTarget
                            {
                                Metadata = metadata,
                                ObjectMode = FormulaOleContract.MathTypeOleMode,
                                Start = formulaRange.Start,
                                End = formulaRange.End,
                                FormulaRange = formulaRange,
                                OleShape = shape,
                            });
                            formulaRange = null;
                            shape = null;
                        }
                        finally
                        {
                            Release(formulaRange);
                            Release(shape);
                        }
                    }
                }
                finally { Release(shapes); }
                return targets;
            }

            foreach (var formulaId in WordOmmlFormulaStore.FormulaIds(document))
            {
                Bookmark? bookmark = null;
                Range? formulaRange = null;
                try
                {
                    bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                    if (bookmark is null) continue;
                    var metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
                    if (metadata is null) continue;
                    if (refreshOmmlMetadata)
                        metadata = WordOmmlNativeSource.RefreshForVisualTeX(
                            document,
                            bookmark,
                            metadata);
                    formulaRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                    if (!FormulaRangeMatchesScope(
                            formulaRange,
                            scope,
                            wholeDocument))
                        continue;
                    // Word can leave multiple collapsed VTOMML anchors near one
                    // physical OMath after undo/paste or an interrupted edit. A
                    // formula-to-LaTeX operation is defined by the visible native
                    // equation, not by the number of stale logical anchors.
                    if (targets.Any(target => FormulaRangesMatch(
                            target.FormulaRange,
                            formulaRange)))
                        continue;
                    targets.Add(new FormulaToLatexTarget
                    {
                        Metadata = metadata,
                        ObjectMode = FormulaOleContract.WordOmmlMode,
                        Start = formulaRange.Start,
                        End = formulaRange.End,
                        FormulaRange = formulaRange,
                        OmmlBookmark = bookmark,
                    });
                    formulaRange = null;
                    bookmark = null;
                }
                finally
                {
                    Release(formulaRange);
                    Release(bookmark);
                }
            }

            OMaths? nativeMaths = null;
            try
            {
                nativeMaths = document.OMaths;
                for (var index = 1; index <= nativeMaths.Count; index++)
                {
                    OMath? math = null;
                    Range? formulaRange = null;
                    try
                    {
                        math = nativeMaths[index];
                        formulaRange = math.Range.Duplicate;
                        if (!FormulaRangeMatchesScope(
                                formulaRange,
                                scope,
                                wholeDocument))
                            continue;
                        if (targets.Any(target => FormulaRangesMatch(
                                target.FormulaRange,
                                formulaRange)))
                            continue;
                        var metadata = WordOmmlNativeSource.CreateForNative(
                            document,
                            formulaRange);
                        targets.Add(new FormulaToLatexTarget
                        {
                            Metadata = metadata,
                            ObjectMode = FormulaOleContract.WordOmmlMode,
                            Start = formulaRange.Start,
                            End = formulaRange.End,
                            FormulaRange = formulaRange,
                        });
                        formulaRange = null;
                    }
                    finally
                    {
                        Release(formulaRange);
                        Release(math);
                    }
                }
            }
            finally { Release(nativeMaths); }
            return targets;
        }
        catch
        {
            ReleaseFormulaToLatexTargets(targets);
            throw;
        }
    }

    private static bool FormulaRangesMatch(Range left, Range right) =>
        left.Start == right.Start && left.End == right.End
        || left.Start <= right.Start && left.End >= right.End
        || right.Start <= left.Start && right.End >= left.End;

    private static bool FormulaRangeMatchesScope(
        Range formulaRange,
        Range scope,
        bool wholeDocument)
    {
        if (wholeDocument) return true;
        var scopeStart = scope.Start;
        var scopeEnd = scope.End;
        var formulaStart = formulaRange.Start;
        var formulaEnd = formulaRange.End;
        if (scopeStart == scopeEnd)
            return scopeStart >= formulaStart && scopeStart <= formulaEnd;
        return formulaStart >= scopeStart && formulaEnd <= scopeEnd;
    }

    private void ConvertFormulaTargetToLatex(
        Document document,
        FormulaToLatexTarget target,
        ref bool documentMutationStarted,
        IReadOnlyDictionary<string, int>? knownReferenceCounts = null,
        bool preserveCrossReferences = false)
    {
        var metadata = target.Metadata;
        var latexSource = target.LatexSource;
        if (string.IsNullOrWhiteSpace(latexSource))
            throw new InvalidDataException(
                $"公式 {metadata.FormulaId} 没有可安全恢复的 LaTeX 源码。");
        var formulaStart = target.FormulaRange.Start;
        Table? numberedTable = null;
        var numberedNativeHashSequence = false;
        Range? tableRange = null;
        Range? convertedTableRange = null;
        Range? insertion = null;
        Range? inserted = null;
        try
        {
            if (string.Equals(
                    target.ObjectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal))
            {
                if (target.OleShape is null)
                    throw new InvalidDataException(
                        "MathType 公式转 LaTeX 时无法定位源 Equation.DSMT4 对象。");

                documentMutationStarted = true;
                var referenceAliases = metadata.Numbered
                    ? MathTypeEquationReferences.CaptureReferenceBookmarkAliases(
                        document,
                        target.OleShape)
                    : Array.Empty<string>();
                if (!preserveCrossReferences && referenceAliases.Count > 0)
                    MathTypeEquationReferences.FreezeReferencesToPlainText(
                        document,
                        referenceAliases);

                Range? liveSourceRange = null;
                try
                {
                    liveSourceRange = target.OleShape.Range;
                    var formatTarget = new WordFormulaFormatConversionTarget
                    {
                        SourceFormulaId = metadata.FormulaId,
                        SourceObjectId = $"{RangeReferencePrefix}{liveSourceRange.Start}:{liveSourceRange.End}",
                        SourceStart = liveSourceRange.Start,
                        Latex = metadata.Latex,
                        DisplayMode = metadata.DisplayMode,
                        Numbered = metadata.Numbered,
                        FontSizePt = FormulaFontSize.Normalize(
                            metadata.FontSizePt ?? FormulaFontSize.DefaultPt),
                        Metadata = metadata,
                    };
                    var insertionStart = DeleteSimpleSourceHost(
                        document,
                        FormulaOleContract.MathTypeOleMode,
                        formatTarget);
                    ThrowIfFormulaToLatexFailureInjected(target);
                    insertion = document.Range(insertionStart, insertionStart);
                    insertion.Text = latexSource;
                    inserted = document.Range(
                        insertionStart,
                        insertionStart + latexSource.Length);
                    VerifyLatexSourceRange(inserted, latexSource, metadata.FormulaId);
                    NormalizeLatexSourceRange(inserted, metadata);
                    return;
                }
                finally { Release(liveSourceRange); }
            }

            numberedTable = TryGetVisualTeXNumberedTable(
                target.FormulaRange,
                metadata);
            if (numberedTable is not null)
            {
                tableRange = numberedTable.Range;
                formulaStart = tableRange.Start;
            }
            else if (metadata.Numbered
                && string.Equals(
                    target.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal)
                && string.Equals(
                    metadata.DisplayMode,
                    "block",
                    StringComparison.Ordinal))
            {
                numberedNativeHashSequence =
                    WordEquationNumbering
                        .HasReusableNumberedNativeOmmlHashSequenceHost(
                            document,
                            target.FormulaRange,
                            metadata.FormulaId);
            }

            if (metadata.Numbered)
            {
                documentMutationStarted = true;
                if (!preserveCrossReferences)
                {
                    WordEquationNumbering.FreezeFormulaCrossReferences(
                        document,
                        metadata.FormulaId,
                        knownReferenceCounts);
                }
                if (numberedNativeHashSequence)
                {
                    // The current OMML number is part of the same OMath that will
                    // be deleted below. Remove only its three alias bookmarks now;
                    // running the general unnumbering path would first replace the
                    // mathematical field and can spill its SEQ result into a stray
                    // body paragraph during OMML→MathType/VisualTeX conversion.
                    WordEquationNumbering
                        .RemoveNativeOmmlHashSequenceAliasesForReplacement(
                            document,
                            metadata.FormulaId);
                }
                else
                {
                    WordEquationNumbering.RemoveFormulaNumberingArtifacts(
                        document,
                        metadata.FormulaId);
                }
            }
            RemoveInlineBaselineSentinel(document, metadata.FormulaId);
            if (string.Equals(
                    target.ObjectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
                RemoveInlineOleTypingAnchorAfter(target.FormulaRange);

            if (numberedTable is not null)
            {
                // Do not delete the table and then recreate a Range from its old
                // coordinates. In saved/reopened Word documents Table.Delete can
                // leave the old insertion point in a transient table boundary:
                // Range.Text then returns normally but writes nothing. Ask Word to
                // dismantle the table first and use the live Range it returns.
                object separator = WdTableFieldSeparator.wdSeparateByParagraphs;
                object nestedTables = false;
                convertedTableRange = numberedTable.ConvertToText(
                    ref separator,
                    ref nestedTables);
                formulaStart = convertedTableRange.Start;
                ThrowIfFormulaToLatexFailureInjected(target);
                convertedTableRange.Text = latexSource + "\r";
                inserted = document.Range(
                    formulaStart,
                    formulaStart + latexSource.Length);
            }
            else
            {
                documentMutationStarted = true;
                if (string.Equals(
                        target.ObjectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal))
                {
                    target.OleShape?.Delete();
                }
                else
                {
                    target.FormulaRange.Delete();
                }
                ThrowIfFormulaToLatexFailureInjected(target);
                insertion = document.Range(formulaStart, formulaStart);
                insertion.Text = latexSource;
                inserted = document.Range(
                    formulaStart,
                    formulaStart + latexSource.Length);
                if (metadata.Numbered)
                    NormalizeFormerNumberedFormulaParagraph(document, inserted);
            }
            TraceFormulaToLatexBridgeState(
                document,
                inserted,
                latexSource,
                metadata.FormulaId,
                "before-detach");
            DetachLatexSourceFromVisualTeXNumberingFrame(inserted, metadata);
            TraceFormulaToLatexBridgeState(
                document,
                inserted,
                latexSource,
                metadata.FormulaId,
                "after-detach");
            VerifyLatexSourceRange(inserted, latexSource, metadata.FormulaId);
            NormalizeLatexSourceRange(inserted, metadata);

            // Keep OMML metadata/bookmarks alive until the plain-text replacement
            // has been verified. If writing fails, the enclosing custom Undo record
            // can restore the original equation together with its ownership anchor.
            if (string.Equals(
                    target.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                try { target.OmmlBookmark?.Delete(); } catch { }
                WordOmmlFormulaStore.Delete(document, metadata.FormulaId);
            }
        }
        finally
        {
            Release(inserted);
            Release(insertion);
            Release(convertedTableRange);
            Release(tableRange);
            Release(numberedTable);
        }
    }

    private static void DetachLatexSourceFromVisualTeXNumberingFrame(
        Range inserted,
        FormulaMetadata metadata)
    {
        if (!metadata.Numbered) return;
        Frames? frames = null;
        try
        {
            frames = inserted.Frames;
            for (var index = frames.Count; index >= 1; index--)
            {
                Frame? frame = null;
                try
                {
                    frame = frames[index];
                    var clippedVisualTeXCaptionFrame =
                        frame.Width <= 1f
                        && frame.Height <= 1f
                        && frame.LockAnchor
                        && !frame.TextWrap
                        && frame.RelativeHorizontalPosition
                            == WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                        && frame.RelativeVerticalPosition
                            == WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    if (clippedVisualTeXCaptionFrame)
                        frame.Delete();
                }
                finally { Release(frame); }
            }
        }
        finally { Release(frames); }

        Frames? remaining = null;
        try
        {
            remaining = inserted.Frames;
            for (var index = 1; index <= remaining.Count; index++)
            {
                Frame? frame = null;
                try
                {
                    frame = remaining[index];
                    if (frame.Width <= 1f && frame.Height <= 1f)
                        throw new InvalidDataException(
                            $"公式 {metadata.FormulaId} 的 LaTeX 源码仍位于隐藏编号框架中。为避免源码不可见，转换已回滚。");
                }
                finally { Release(frame); }
            }
        }
        finally { Release(remaining); }
    }

    private static string BuildFormulaLatexSource(FormulaMetadata metadata)
    {
        var latex = string.IsNullOrWhiteSpace(metadata.Latex)
            ? string.Join("\n", metadata.Lines.Select(line => line.Latex))
            : metadata.Latex;
        latex = (latex ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Trim();
        if (string.IsNullOrWhiteSpace(latex))
            throw new InvalidDataException(
                $"公式 {metadata.FormulaId} 的 LaTeX 元数据为空。为避免删除原公式，转换已中止。");
        if (string.Equals(
                metadata.DisplayMode,
                "block",
                StringComparison.Ordinal))
        {
            latex = FormulaEquationTag.Attach(latex, metadata.EquationTag);
            return "$$" + latex + "$$";
        }
        latex = FormulaEquationTag.Extract(latex).Latex
            .Replace('\n', ' ');
        return "$" + latex + "$";
    }

    private static string NormalizeFormulaToLatexVerificationText(string value) =>
        (value ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Replace('\v', '\n');

    private static void TraceFormulaToLatexBridgeState(
        Document document,
        Range inserted,
        string expected,
        string formulaId,
        string stage)
    {
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_LATEX_BRIDGE"),
                "1",
                StringComparison.Ordinal))
            return;

        Frames? frames = null;
        Range? content = null;
        try
        {
            var actual = inserted.Text ?? string.Empty;
            frames = inserted.Frames;
            content = document.Content;
            var documentText = content.Text ?? string.Empty;
            var expectedOffset = documentText.IndexOf(expected, StringComparison.Ordinal);
            var expectedStart = expectedOffset >= 0
                ? content.Start + expectedOffset
                : -1;
            WordDoubleClickHook.TraceMessage(
                $"formula-to-latex-bridge-state stage={stage} formulaId={formulaId} range={inserted.Start}:{inserted.End} actualLength={actual.Length} frames={frames.Count} expectedStart={expectedStart}");
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage(
                $"formula-to-latex-bridge-state-failed stage={stage} formulaId={formulaId} error={error.Message}");
        }
        finally
        {
            Release(content);
            Release(frames);
        }
    }

    private static void VerifyLatexSourceRange(
        Range inserted,
        string expected,
        string formulaId)
    {
        var actual = NormalizeFormulaToLatexVerificationText(
            inserted.Text ?? string.Empty);
        var normalizedExpected = NormalizeFormulaToLatexVerificationText(expected);
        if (!string.Equals(actual, normalizedExpected, StringComparison.Ordinal))
        {
            var actualCodes = string.Join(
                ",",
                actual.Take(128).Select(ch => $"U+{(int)ch:X4}"));
            var expectedCodes = string.Join(
                ",",
                normalizedExpected.Take(128).Select(ch => $"U+{(int)ch:X4}"));
            WordDoubleClickHook.TraceMessage(
                $"formula-to-latex-verify-mismatch formulaId={formulaId} range={inserted.Start}:{inserted.End} actualLength={actual.Length} expectedLength={normalizedExpected.Length} actualCodes={actualCodes} expectedCodes={expectedCodes}");
            throw new InvalidDataException(
                $"公式 {formulaId} 的 LaTeX 写回校验失败。Word 实际写入内容与预期源码不一致。");
        }

        Frames? frames = null;
        try
        {
            frames = inserted.Frames;
            if (frames.Count > 0)
                throw new InvalidDataException(
                    $"公式 {formulaId} 的 LaTeX 源码被 Word Frame 包围。为避免源码不可见，本次转换已撤销。");
        }
        finally { Release(frames); }
    }

    private static void ApplyFormulaToLatexEmptySourceInjection(
        FormulaToLatexTarget target)
    {
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            return;
        var requested = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_FORMULA_TO_LATEX_EMPTY_SOURCE");
        if (string.IsNullOrWhiteSpace(requested)) return;
        if (!string.Equals(requested, "1", StringComparison.Ordinal)
            && !string.Equals(
                requested,
                target.Metadata.FormulaId,
                StringComparison.OrdinalIgnoreCase))
            return;
        target.Metadata.Latex = string.Empty;
        foreach (var line in target.Metadata.Lines)
            line.Latex = string.Empty;
    }

    private static void ThrowIfFormulaToLatexFailureInjected(
        FormulaToLatexTarget target)
    {
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            return;
        var requested = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_FORMULA_TO_LATEX_FAIL_AFTER_DELETE");
        if (string.IsNullOrWhiteSpace(requested)) return;
        if (!string.Equals(requested, "1", StringComparison.Ordinal)
            && !string.Equals(
                requested,
                target.Metadata.FormulaId,
                StringComparison.OrdinalIgnoreCase))
            return;
        throw new InvalidOperationException(
            "Injected formula-to-LaTeX failure after deleting the source object.");
    }

    private static bool TryUndoFormulaToLatexConversion(Document document)
    {
        object times = 1;
        try { return document.Undo(ref times); }
        catch { return false; }
    }

    private static Table? TryGetVisualTeXNumberedTable(
        Range formulaRange,
        FormulaMetadata metadata)
    {
        if (!metadata.Numbered) return null;
        try
        {
            if (!(bool)formulaRange.get_Information(WdInformation.wdWithInTable)
                || formulaRange.Tables.Count == 0)
                return null;
            var table = formulaRange.Tables[1];
            if (table.Columns.Count < 3)
            {
                Release(table);
                return null;
            }
            return table;
        }
        catch { return null; }
    }

    private static void NormalizeFormerNumberedFormulaParagraph(
        Document document,
        Range latexRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefix = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        try
        {
            paragraphs = latexRange.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (latexRange.Start > paragraphRange.Start)
            {
                prefix = document.Range(paragraphRange.Start, latexRange.Start);
                var prefixText = prefix.Text ?? string.Empty;
                if (prefixText.All(character => character is '\t' or '\v' or ' '))
                    prefix.Delete();
            }
            format = paragraphRange.ParagraphFormat;
            format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            // A numbered 1x3 formula can be adjacent to VisualTeX's mandatory
            // compact table separator (1pt exact line spacing). When the table is
            // removed during OMML→LaTeX conversion Word may reuse that paragraph's
            // pPr for the newly inserted visible source text. Font normalization
            // alone is not enough: 10.5pt LaTeX rendered inside a 1pt exact line
            // box is visibly crushed. Repair only this unmistakable compact-tail
            // signature; preserve ordinary/custom user line spacing otherwise.
            if (format.LineSpacingRule == WdLineSpacing.wdLineSpaceExactly
                && format.LineSpacing <= 2.01f)
            {
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                format.SpaceBefore = 0f;
                format.SpaceAfter = 0f;
                try { format.DisableLineHeightGrid = -1; } catch { }
            }
            tabStops = format.TabStops;
            tabStops.ClearAll();
        }
        finally
        {
            Release(tabStops);
            Release(format);
            Release(prefix);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void NormalizeLatexSourceRange(
        Range range,
        FormulaMetadata metadata)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? paragraphFormat = null;
        try
        {
            font = range.Font;
            font.Hidden = 0;
            font.Position = 0;
            font.Superscript = 0;
            font.Subscript = 0;
            var size = FormulaFontSize.ResolveSemanticFontSize(metadata);
            if (size > 0) font.Size = (float)size;

            // Also protect unnumbered/cross-format LaTeX restores that happen to
            // land in a compact VisualTeX structural paragraph. Only the 1-2pt
            // exact-spacing sentinel is normalized, so intentional user paragraph
            // formatting remains untouched.
            paragraphs = range.Paragraphs;
            if (paragraphs.Count == 1)
            {
                paragraph = paragraphs[1];
                paragraphFormat = paragraph.Range.ParagraphFormat;
                if (paragraphFormat.LineSpacingRule == WdLineSpacing.wdLineSpaceExactly
                    && paragraphFormat.LineSpacing <= 2.01f)
                {
                    paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                    paragraphFormat.SpaceBefore = 0f;
                    paragraphFormat.SpaceAfter = 0f;
                    try { paragraphFormat.DisableLineHeightGrid = -1; } catch { }
                }
            }
        }
        finally
        {
            Release(paragraphFormat);
            Release(paragraph);
            Release(paragraphs);
            Release(font);
        }
    }

    private static void ReleaseFormulaToLatexTargets(
        IEnumerable<FormulaToLatexTarget>? targets)
    {
        if (targets is null) return;
        foreach (var target in targets)
        {
            Release(target.OmmlBookmark);
            Release(target.OleShape);
            Release(target.FormulaRange);
        }
    }

    private static List<ResolvedLatexRedrawTarget> ResolveLatexRedrawTargets(
        Document document,
        WordLatexRedrawPlan plan,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared)
    {
        var resolved = new List<ResolvedLatexRedrawTarget>(plan.Targets.Count);
        try
        {
            foreach (var target in plan.Targets.OrderBy(item => item.RelativeStart))
            {
                if (!prepared.TryGetValue(target.Id, out var formula))
                    throw new InvalidDataException($"缺少公式 {target.Id} 的渲染结果。");
                if (target.RelativeStart < 0
                    || target.SourceLength <= 0
                    || target.RelativeStart + target.SourceLength > plan.SourceText.Length)
                    throw new InvalidDataException($"公式 {target.Id} 的源文本范围无效。");

                var expectedSource = plan.SourceText.Substring(
                    target.RelativeStart,
                    target.SourceLength);
                var sourceRange = ResolveExactLatexSourceRange(
                    document,
                    plan,
                    target,
                    expectedSource,
                    resolved);
                resolved.Add(new ResolvedLatexRedrawTarget
                {
                    Target = target,
                    Formula = formula,
                    SourceRange = sourceRange,
                    SourceStart = sourceRange.Start,
                    SourceEnd = sourceRange.End,
                    ExpectedSource = expectedSource,
                });
            }
            return resolved;
        }
        catch
        {
            foreach (var item in resolved)
                Release(item.SourceRange);
            throw;
        }
    }

    private static Range ResolveExactLatexSourceRange(
        Document document,
        WordLatexRedrawPlan plan,
        WordLatexRedrawTarget target,
        string expectedSource,
        IReadOnlyList<ResolvedLatexRedrawTarget> alreadyResolved)
    {
        var hasResolvedCoordinates =
            target.AbsoluteStart >= plan.ScopeStart
            && target.AbsoluteEnd > target.AbsoluteStart;
        var approximateStart = hasResolvedCoordinates
            ? target.AbsoluteStart
            : plan.ScopeStart + target.RelativeStart;
        var approximateEnd = hasResolvedCoordinates
            ? target.AbsoluteEnd
            : approximateStart + target.SourceLength;
        Range? direct = null;
        try
        {
            if (approximateStart >= plan.ScopeStart
                && approximateEnd >= approximateStart
                && approximateEnd <= plan.ScopeEnd)
            {
                direct = document.Range(approximateStart, approximateEnd);
                if (string.Equals(
                        direct.Text ?? string.Empty,
                        expectedSource,
                        StringComparison.Ordinal)
                    && !OverlapsResolvedLatexRange(direct, alreadyResolved))
                {
                    target.AbsoluteStart = approximateStart;
                    target.AbsoluteEnd = approximateEnd;
                    var result = direct;
                    direct = null;
                    return result;
                }
            }
        }
        finally { Release(direct); }

        const int localSearchRadius = 1024;
        var localStart = Math.Max(plan.ScopeStart, approximateStart - localSearchRadius);
        var localEnd = Math.Min(plan.ScopeEnd, approximateEnd + localSearchRadius);
        var located = FindExactLatexSourceRange(
            document,
            localStart,
            localEnd,
            approximateStart,
            expectedSource,
            alreadyResolved);
        if (located is not null)
        {
            target.AbsoluteStart = located.Start;
            target.AbsoluteEnd = located.End;
            return located;
        }

        located = FindExactLatexSourceRange(
            document,
            plan.ScopeStart,
            plan.ScopeEnd,
            approximateStart,
            expectedSource,
            alreadyResolved);
        if (located is not null)
        {
            target.AbsoluteStart = located.Start;
            target.AbsoluteEnd = located.End;
            return located;
        }

        throw new InvalidOperationException(
            $"无法在原位置附近重新定位公式：{target.Latex}。为避免替换错误内容，本次重绘已停止。");
    }

    private static Range? FindExactLatexSourceRange(
        Document document,
        int searchStart,
        int searchEnd,
        int approximateStart,
        string expectedSource,
        IReadOnlyList<ResolvedLatexRedrawTarget> alreadyResolved)
    {
        if (searchEnd <= searchStart || expectedSource.Length == 0) return null;
        var findText = BuildWordFindAnchor(expectedSource);
        if (findText.Length == 0) return null;

        Range? search = null;
        Range? best = null;
        var bestDistance = int.MaxValue;
        try
        {
            search = document.Range(searchStart, searchEnd);
            while (search.Start < searchEnd)
            {
                Find? find = null;
                var matched = false;
                try
                {
                    find = search.Find;
                    find.ClearFormatting();
                    find.Text = findText;
                    find.Forward = true;
                    find.Wrap = WdFindWrap.wdFindStop;
                    find.Format = false;
                    find.MatchCase = true;
                    find.MatchWholeWord = false;
                    find.MatchWildcards = false;
                    find.MatchSoundsLike = false;
                    find.MatchAllWordForms = false;
                    matched = find.Execute();
                }
                finally { Release(find); }
                if (!matched) break;

                var matchStart = search.Start;
                var nextSearchStart = Math.Min(
                    searchEnd,
                    Math.Max(matchStart + 1, search.End));
                var candidate = TryCreateExactLatexRangeAt(
                    document,
                    matchStart,
                    searchEnd,
                    expectedSource);
                if (candidate is not null)
                {
                    if (!OverlapsResolvedLatexRange(candidate, alreadyResolved))
                    {
                        var distance = Math.Abs(candidate.Start - approximateStart);
                        if (distance < bestDistance)
                        {
                            Release(best);
                            best = candidate;
                            candidate = null;
                            bestDistance = distance;
                        }
                    }
                    Release(candidate);
                }
                search.SetRange(nextSearchStart, searchEnd);
            }

            var result = best;
            best = null;
            return result;
        }
        finally
        {
            Release(best);
            Release(search);
        }
    }

    private static Range? TryCreateExactLatexRangeAt(
        Document document,
        int start,
        int maximumEnd,
        string expectedSource)
    {
        const int maximumCoordinateAdjustment = 256;
        for (var adjustment = 0;
             adjustment <= maximumCoordinateAdjustment;
             adjustment++)
        {
            var deltas = adjustment == 0
                ? new[] { 0 }
                : new[] { -adjustment, adjustment };
            foreach (var delta in deltas)
            {
                var end = start + expectedSource.Length + delta;
                if (end <= start || end > maximumEnd) continue;
                Range? candidate = null;
                try
                {
                    candidate = document.Range(start, end);
                    if (string.Equals(
                            candidate.Text ?? string.Empty,
                            expectedSource,
                            StringComparison.Ordinal))
                    {
                        var result = candidate;
                        candidate = null;
                        return result;
                    }
                }
                catch (COMException)
                {
                    // Keep trying nearby Word story coordinates.
                }
                finally { Release(candidate); }
            }
        }
        return null;
    }

    private static bool OverlapsResolvedLatexRange(
        Range candidate,
        IReadOnlyList<ResolvedLatexRedrawTarget> alreadyResolved)
    {
        // Read the COM coordinates once. The previous implementation read
        // Range.Start/End again for every earlier formula, which turns a
        // 1000-formula document into hundreds of thousands of cross-process COM
        // calls even though the overlap comparison itself is trivial.
        var candidateStart = candidate.Start;
        var candidateEnd = candidate.End;
        foreach (var resolved in alreadyResolved)
        {
            if (candidateStart < resolved.SourceEnd
                && resolved.SourceStart < candidateEnd)
                return true;
        }
        return false;
    }

    private static string BuildWordFindAnchor(string source)
    {
        const int maximumFindTextLength = 180;
        var builder = new StringBuilder(Math.Min(source.Length, maximumFindTextLength));
        foreach (var character in source)
        {
            var token = character switch
            {
                '^' => "^^",
                '\r' => "^p",
                '\v' => "^l",
                '\t' => "^t",
                _ => character.ToString(),
            };
            if (builder.Length > 0
                && builder.Length + token.Length > maximumFindTextLength)
                break;
            builder.Append(token);
        }
        return builder.ToString();
    }

    private static double?[]? TryBuildWordOpenXmlFontSizeIndex(
        Range scope,
        string sourceText)
    {
        const string packageNamespace =
            "http://schemas.microsoft.com/office/2006/xmlPackage";
        const string wordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        try
        {
            var package = XDocument.Parse(
                scope.WordOpenXML,
                LoadOptions.PreserveWhitespace);
            XNamespace packageNs = packageNamespace;
            XNamespace word = wordNamespace;
            var documentPart = package
                .Descendants(packageNs + "part")
                .FirstOrDefault(part => string.Equals(
                    (string?)part.Attribute(packageNs + "name"),
                    "/word/document.xml",
                    StringComparison.OrdinalIgnoreCase));
            var body = documentPart?
                .Descendants(word + "body")
                .FirstOrDefault();
            if (body is null) return null;

            var reconstructed = new StringBuilder(sourceText.Length + 16);
            var fontSizes = new List<double?>(sourceText.Length + 16);
            foreach (var paragraph in body.Descendants(word + "p"))
            {
                var paragraphFontSize = ReadWordOpenXmlFontSize(
                    paragraph.Element(word + "pPr")?.Element(word + "rPr"),
                    word);
                foreach (var token in paragraph.Descendants())
                {
                    string? tokenText = null;
                    if (token.Name == word + "t"
                        || token.Name == word + "delText")
                    {
                        tokenText = token.Value;
                    }
                    else if (token.Name == word + "tab")
                    {
                        tokenText = "\t";
                    }
                    else if (token.Name == word + "br"
                        || token.Name == word + "cr")
                    {
                        tokenText = "\v";
                    }
                    else if (token.Name == word + "noBreakHyphen")
                    {
                        tokenText = "\u2011";
                    }
                    else if (token.Name == word + "softHyphen")
                    {
                        tokenText = "\u00AD";
                    }
                    if (tokenText is null) continue;

                    var run = token.Ancestors(word + "r").FirstOrDefault();
                    var runFontSize = ReadWordOpenXmlFontSize(
                        run?.Element(word + "rPr"),
                        word)
                        ?? paragraphFontSize;
                    reconstructed.Append(tokenText);
                    for (var index = 0; index < tokenText.Length; index++)
                        fontSizes.Add(runFontSize);
                }
                reconstructed.Append('\r');
                fontSizes.Add(paragraphFontSize);
            }

            var reconstructedText = reconstructed.ToString();
            if (string.Equals(
                    reconstructedText,
                    sourceText,
                    StringComparison.Ordinal))
                return fontSizes.ToArray();

            // A Word Open XML fragment can include one mandatory trailing empty
            // paragraph that is outside the requested story range. Accept only that
            // exact, unambiguous difference; every other mismatch falls back to COM.
            if (reconstructedText.Length == sourceText.Length + 1
                && reconstructedText[reconstructedText.Length - 1] == '\r'
                && string.Equals(
                    reconstructedText.Substring(0, sourceText.Length),
                    sourceText,
                    StringComparison.Ordinal))
            {
                fontSizes.RemoveAt(fontSizes.Count - 1);
                return fontSizes.ToArray();
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    private static double? ReadWordOpenXmlFontSize(
        XElement? runProperties,
        XNamespace word)
    {
        if (runProperties is null) return null;
        var value = (string?)runProperties
            .Element(word + "sz")?
            .Attribute(word + "val")
            ?? (string?)runProperties
                .Element(word + "szCs")?
                .Attribute(word + "val");
        if (!double.TryParse(
                value,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out var halfPoints)
            || halfPoints <= 0
            || double.IsNaN(halfPoints)
            || double.IsInfinity(halfPoints))
            return null;
        return halfPoints / 2d;
    }

    private static IReadOnlyDictionary<string, LatexRedrawSourceContext>
        BuildLatexRedrawSourceContexts(
            string sourceText,
            IReadOnlyList<WordLatexRedrawTarget> targets)
    {
        var contexts = new Dictionary<string, LatexRedrawSourceContext>(
            targets.Count,
            StringComparer.Ordinal);
        var formulaMask = new bool[sourceText.Length];
        foreach (var target in targets)
        {
            var start = Math.Max(0, Math.Min(sourceText.Length, target.RelativeStart));
            var end = Math.Max(start, Math.Min(
                sourceText.Length,
                target.RelativeStart + target.SourceLength));
            for (var index = start; index < end; index++)
                formulaMask[index] = true;
        }

        foreach (var target in targets)
        {
            var start = Math.Max(0, Math.Min(sourceText.Length, target.RelativeStart));
            var end = Math.Max(start, Math.Min(
                sourceText.Length,
                target.RelativeStart + target.SourceLength));
            var paragraphStart = FindSourceParagraphStart(sourceText, start);
            var paragraphEnd = FindSourceParagraphEnd(sourceText, end);
            var previousInParagraph = FindPreviousVisibleSourcePosition(
                sourceText,
                start - 1,
                paragraphStart,
                formulaMask);
            var nextInParagraph = FindNextVisibleSourcePosition(
                sourceText,
                end,
                paragraphEnd,
                formulaMask);
            var hasVisibleSurroundingText =
                previousInParagraph >= 0 || nextInParagraph >= 0;
            var display = string.Equals(
                target.DisplayMode,
                "block",
                StringComparison.Ordinal);

            int fontContextPosition;
            if (!display || hasVisibleSurroundingText)
            {
                fontContextPosition = previousInParagraph >= 0
                    ? previousInParagraph
                    : nextInParagraph;
            }
            else
            {
                var previousParagraph = FindPreviousVisibleSourcePosition(
                    sourceText,
                    paragraphStart - 1,
                    0,
                    formulaMask);
                var nextParagraph = FindNextVisibleSourcePosition(
                    sourceText,
                    paragraphEnd,
                    sourceText.Length,
                    formulaMask);
                fontContextPosition = previousParagraph >= 0
                    ? previousParagraph
                    : nextParagraph;
            }

            contexts[target.Id] = new LatexRedrawSourceContext
            {
                HasVisibleSurroundingText = hasVisibleSurroundingText,
                FontContextRelativePosition = fontContextPosition,
            };
        }
        return contexts;
    }

    private static int FindSourceParagraphStart(string sourceText, int position)
    {
        for (var index = Math.Min(position, sourceText.Length) - 1;
             index >= 0;
             index--)
        {
            if (IsSourceParagraphBoundary(sourceText[index]))
                return index + 1;
        }
        return 0;
    }

    private static int FindSourceParagraphEnd(string sourceText, int position)
    {
        for (var index = Math.Max(0, position);
             index < sourceText.Length;
             index++)
        {
            if (IsSourceParagraphBoundary(sourceText[index]))
                return index;
        }
        return sourceText.Length;
    }

    private static int FindPreviousVisibleSourcePosition(
        string sourceText,
        int startPosition,
        int minimumPosition,
        IReadOnlyList<bool> formulaMask)
    {
        for (var index = Math.Min(startPosition, sourceText.Length - 1);
             index >= Math.Max(0, minimumPosition);
             index--)
        {
            if (char.IsLowSurrogate(sourceText[index])
                && index > minimumPosition
                && char.IsHighSurrogate(sourceText[index - 1]))
                index--;
            if (formulaMask[index]) continue;
            if (char.IsHighSurrogate(sourceText[index])
                && index + 1 < sourceText.Length
                && char.IsLowSurrogate(sourceText[index + 1])
                && formulaMask[index + 1])
                continue;
            if (IsVisibleSourceCharacter(sourceText[index])) return index;
        }
        return -1;
    }

    private static int FindNextVisibleSourcePosition(
        string sourceText,
        int startPosition,
        int maximumPosition,
        IReadOnlyList<bool> formulaMask)
    {
        var maximum = Math.Min(sourceText.Length, maximumPosition);
        for (var index = Math.Max(0, startPosition);
             index < maximum;
             index++)
        {
            if (char.IsLowSurrogate(sourceText[index])
                && index > 0
                && char.IsHighSurrogate(sourceText[index - 1]))
                continue;
            if (formulaMask[index]) continue;
            if (char.IsHighSurrogate(sourceText[index])
                && index + 1 < sourceText.Length
                && char.IsLowSurrogate(sourceText[index + 1])
                && formulaMask[index + 1])
                continue;
            if (IsVisibleSourceCharacter(sourceText[index])) return index;
            if (char.IsHighSurrogate(sourceText[index])
                && index + 1 < maximum
                && char.IsLowSurrogate(sourceText[index + 1]))
                index++;
        }
        return -1;
    }

    private static bool IsSourceParagraphBoundary(char character) =>
        character is '\r' or '\n' or '\v' or '\a';

    private static bool IsVisibleSourceCharacter(char character)
    {
        if (character is '\r' or '\n' or '\t' or '\v' or '\a'
            or '\u0001' or '\u200B' or '\u200C')
            return false;
        return !char.IsWhiteSpace(character);
    }

    private static bool TryResolveSourceFormulaFontSizeFromContext(
        Document document,
        WordLatexRedrawPlan plan,
        WordLatexRedrawTarget target,
        int sourceStart,
        int sourceEnd,
        int contextRelativePosition,
        out float fontSizePt)
    {
        fontSizePt = FormulaFontSize.DefaultPt;
        if (contextRelativePosition < 0
            || contextRelativePosition >= plan.SourceText.Length)
            return false;

        var targetEnd = target.RelativeStart + target.SourceLength;
        int wordPosition;
        if (contextRelativePosition < target.RelativeStart)
        {
            wordPosition = sourceStart - CountWordStoryCharacters(
                plan.SourceText,
                contextRelativePosition,
                target.RelativeStart);
        }
        else if (contextRelativePosition >= targetEnd)
        {
            wordPosition = sourceEnd + CountWordStoryCharacters(
                plan.SourceText,
                targetEnd,
                contextRelativePosition);
        }
        else
        {
            return false;
        }

        if (wordPosition < plan.ScopeStart || wordPosition >= plan.ScopeEnd)
            return false;

        Range? probe = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            probe = document.Range(wordPosition, wordPosition + 1);
            var expectedText = ReadSourceCharacterAt(
                plan.SourceText,
                contextRelativePosition);
            var probeText = probe.Text ?? string.Empty;
            if (!string.Equals(
                    probeText,
                    expectedText,
                    StringComparison.Ordinal)
                || !ContainsVisibleBodyText(probeText))
                return false;
            font = probe.Font;
            return TryResolveWordFontSize(font.Size, out fontSizePt);
        }
        catch (COMException)
        {
            return false;
        }
        finally
        {
            Release(font);
            Release(probe);
        }
    }

    private static int[] BuildWordStoryOffsetIndex(string sourceText)
    {
        var offsets = new int[sourceText.Length + 1];
        var sourceIndex = 0;
        var wordOffset = 0;
        while (sourceIndex < sourceText.Length)
        {
            offsets[sourceIndex] = wordOffset;
            if (char.IsHighSurrogate(sourceText[sourceIndex])
                && sourceIndex + 1 < sourceText.Length
                && char.IsLowSurrogate(sourceText[sourceIndex + 1]))
            {
                // Word counts a valid UTF-16 surrogate pair as one story
                // character. The intermediate boundary is not a valid formula
                // boundary, but mapping it to the pair start keeps the table total.
                offsets[sourceIndex + 1] = wordOffset;
                sourceIndex += 2;
                wordOffset++;
                offsets[sourceIndex] = wordOffset;
                continue;
            }
            sourceIndex++;
            wordOffset++;
            offsets[sourceIndex] = wordOffset;
        }
        return offsets;
    }

    private static int CountWordStoryCharacters(
        string sourceText,
        int startPosition,
        int endPosition)
    {
        var start = Math.Max(0, Math.Min(sourceText.Length, startPosition));
        var end = Math.Max(start, Math.Min(sourceText.Length, endPosition));
        var count = 0;
        for (var index = start; index < end; index++, count++)
        {
            if (char.IsHighSurrogate(sourceText[index])
                && index + 1 < end
                && char.IsLowSurrogate(sourceText[index + 1]))
                index++;
        }
        return count;
    }

    private static string ReadSourceCharacterAt(string sourceText, int position)
    {
        if (position < 0 || position >= sourceText.Length) return string.Empty;
        if (char.IsHighSurrogate(sourceText[position])
            && position + 1 < sourceText.Length
            && char.IsLowSurrogate(sourceText[position + 1]))
            return sourceText.Substring(position, 2);
        return sourceText[position].ToString();
    }

    private static float ResolveSourceFormulaFontSize(
        Document document,
        Range source,
        bool display)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            paragraphs = source.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
            }

            var paragraphStart = paragraphRange?.Start ?? source.Start;
            var paragraphBodyEnd = Math.Max(
                paragraphStart,
                (paragraphRange?.End ?? source.End) - 1);

            // Inline formulas should inherit the prose beside them. Generated
            // documents often deliberately make the raw $...$ source smaller than
            // the surrounding text, so the source run itself is only a fallback.
            if (!display || HasVisibleSurroundingText(source))
            {
                if (TryResolveNearbyVisibleFontSize(
                        document,
                        source.Start - 1,
                        out var previousInline,
                        minimumPosition: paragraphStart,
                        step: -1))
                    return previousInline;
                if (TryResolveNearbyVisibleFontSize(
                        document,
                        source.End,
                        out var nextInline,
                        maximumPosition: paragraphBodyEnd,
                        step: 1))
                    return nextInline;
            }

            // A display formula normally occupies its own paragraph. In that case
            // inherit from the nearest visible prose outside the source paragraph,
            // preferring the preceding paragraph as Word users generally expect.
            if (TryResolveNearbyVisibleFontSize(
                    document,
                    paragraphStart - 1,
                    out var previousParagraph,
                    minimumPosition: 0,
                    step: -1))
                return previousParagraph;
            if (TryResolveNearbyVisibleFontSize(
                    document,
                    paragraphRange?.End ?? source.End,
                    out var nextParagraph,
                    maximumPosition: int.MaxValue,
                    step: 1))
                return nextParagraph;

            font = source.Font;
            if (TryResolveWordFontSize(font.Size, out var direct)) return direct;
            Release(font);
            font = null;

            if (paragraphRange is not null)
            {
                if (paragraphRange.End > paragraphRange.Start)
                    paragraphRange.End -= 1;
                font = paragraphRange.Font;
                if (TryResolveWordFontSize(font.Size, out var paragraphSize))
                    return paragraphSize;
            }
            return FormulaFontSize.DefaultPt;
        }
        finally
        {
            Release(font);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool TryResolveNearbyVisibleFontSize(
        Document document,
        int startPosition,
        out float fontSizePt,
        int minimumPosition = 0,
        int maximumPosition = int.MaxValue,
        int step = 1)
    {
        fontSizePt = FormulaFontSize.DefaultPt;
        if (step is not (-1 or 1)) return false;

        Range? content = null;
        try
        {
            content = document.Content;
            var contentStart = content.Start;
            var contentEnd = Math.Max(contentStart, content.End - 1);
            var lowerBound = Math.Max(contentStart, minimumPosition);
            var upperBound = Math.Min(contentEnd, maximumPosition);
            if (upperBound < lowerBound
                || (step < 0 && startPosition < lowerBound)
                || (step > 0 && startPosition > upperBound))
                return false;
            var position = startPosition;
            const int maximumProbeCharacters = 256;
            for (var probeIndex = 0;
                 probeIndex < maximumProbeCharacters
                 && position >= lowerBound
                 && position <= upperBound;
                 probeIndex++, position += step)
            {
                Range? probe = null;
                Microsoft.Office.Interop.Word.Font? font = null;
                try
                {
                    probe = document.Range(position, Math.Min(position + 1, content.End));
                    if (!ContainsVisibleBodyText(probe.Text)) continue;
                    font = probe.Font;
                    if (TryResolveWordFontSize(font.Size, out fontSizePt))
                        return true;
                }
                catch (COMException)
                {
                    // Keep probing neighboring Word story coordinates.
                }
                finally
                {
                    Release(font);
                    Release(probe);
                }
            }
            return false;
        }
        finally { Release(content); }
    }

    public WordBulkInsertResult InsertBulkDocument(
        WordBulkImportDocument source,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared,
        string? expectedDocumentId,
        string? sourceObjectId)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (prepared is null) throw new ArgumentNullException(nameof(prepared));
        Document? document = null;
        Selection? selection = null;
        Range? rollbackRange = null;
        UndoRecord? undoRecord = null;
        WordOmmlConverter.BatchSource? ommlBatchSource = null;
        var deferredOmmlMetadata = new List<FormulaMetadata>();
        var insertedFormulaIds = new List<string>();
        var insertionStart = -1;
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, expectedDocumentId);
            try
            {
                previousScreenUpdating = _application.ScreenUpdating;
                _application.ScreenUpdating = false;
                screenUpdatingSuspended = true;
            }
            catch { }
            selection = _application.Selection;
            Range? sourceRange = null;
            try
            {
                sourceRange = ResolveSourceRange(document, sourceObjectId, selection);
                selection.SetRange(sourceRange.Start, sourceRange.End);
            }
            finally { Release(sourceRange); }
            if (selection.Range.Start != selection.Range.End)
                selection.Text = string.Empty;
            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
            insertionStart = selection.Start;
            var nativeOmmlBulk = prepared.Count > 0
                && prepared.Values.All(item => string.Equals(
                    item.Session.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal));
            var ommlPreparedByFormulaId = prepared.Values
                .Where(item => string.Equals(
                    item.Session.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
                .ToDictionary(
                    item => item.Session.FormulaId,
                    item => item,
                    StringComparer.OrdinalIgnoreCase);
            if (ommlPreparedByFormulaId.Count > 0)
            {
                var requestedMathFonts = ommlPreparedByFormulaId.Values
                    .Select(item => ResolveDocumentOmmlMathFont(
                        item.Session.ToMetadata().FormulaLetterFont))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                if (requestedMathFonts.Length != 1)
                    throw new InvalidDataException(
                        "同一次 Word OMML 批量操作不能请求多个文档级数学字体。" );
                var documentMathMetadata = ommlPreparedByFormulaId.Values
                    .First()
                    .Session
                    .ToMetadata();
                documentMathMetadata.Validate();
                ApplyDocumentOmmlMathFont(document, documentMathMetadata);
            }
            var ommlFormulas = ommlPreparedByFormulaId.Values
                .Select(item => (
                    FormulaId: item.Session.FormulaId,
                    MathMl: item.MathMl
                        ?? throw new InvalidDataException(
                            $"公式 {item.Session.FormulaId} 没有可用的 MathML。")))
                .ToList();
            if (ommlFormulas.Count > 0 && !nativeOmmlBulk)
            {
                var batchSourceStopwatch = Stopwatch.StartNew();
                ommlBatchSource = WordOmmlConverter.CreateBatchSource(
                    _application,
                    ommlFormulas,
                    (formulaId, omml) =>
                    {
                        if (!ommlPreparedByFormulaId.TryGetValue(
                                formulaId,
                                out var preparedFormula))
                            throw new InvalidDataException(
                                $"公式 {formulaId} 的 OMML 字体配置不存在。" );
                        var typographyMetadata = preparedFormula.Session.ToMetadata();
                        typographyMetadata.Validate();
                        return ApplyOmmlTypographyXml(
                            omml,
                            preparedFormula.Session.FontSizePt,
                            typographyMetadata);
                    },
                    mathFontName: document.OMathFontName);
                batchSourceStopwatch.Stop();
                if (string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                    Console.WriteLine(
                        $"    [perf] BulkOmml.CreateBatchSource: {batchSourceStopwatch.ElapsedMilliseconds}ms formulas={ommlFormulas.Count}");
                document.Activate();
                Release(selection);
                selection = _application.Selection;
                selection.SetRange(insertionStart, insertionStart);
            }
            undoRecord = BeginUndoRecord("VisualTeX 批量导入 LaTeX / Markdown");

            var nativeOleBulk = prepared.Count > 0
                && prepared.Values.All(item => string.Equals(
                    item.Session.ObjectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal));
            if (nativeOleBulk)
            {
                InsertBulkOleDocumentTwoPhase(
                    document,
                    selection,
                    source,
                    prepared,
                    insertedFormulaIds);
            }
            else if (nativeOmmlBulk)
            {
                InsertBulkOmmlDocumentOneShot(
                    document,
                    selection,
                    source,
                    prepared,
                    insertedFormulaIds,
                    deferredOmmlMetadata);
            }
            else
            {
            for (var blockIndex = 0; blockIndex < source.Blocks.Count; blockIndex++)
            {
                var block = source.Blocks[blockIndex];
                var nextKind = blockIndex + 1 < source.Blocks.Count
                    ? source.Blocks[blockIndex + 1].Kind
                    : (WordBulkBlockKind?)null;
                if (block.Kind == WordBulkBlockKind.DisplayFormula)
                {
                    var formulaRun = block.Runs.Single(run => run.IsFormula);
                    if (!prepared.TryGetValue(formulaRun.Id, out var formula))
                        throw new InvalidDataException(
                            $"缺少行间公式 {formulaRun.Id} 的渲染结果。");
                    InsertPreparedFormula(
                        document,
                        selection,
                        formula,
                        display: true,
                        ommlBatchSource: ommlBatchSource,
                        deferredOmmlMetadata: ommlBatchSource is null
                            ? null
                            : deferredOmmlMetadata,
                        bulkImport: true);
                    insertedFormulaIds.Add(formula.Session.FormulaId);
                    continue;
                }

                EnsureWritableParagraph(selection);
                var paragraphStart = selection.Start;
                var pendingInlineFormulas = new List<(int Start, PreparedWordBulkFormula Formula)>();
                foreach (var run in block.Runs)
                {
                    if (!run.IsFormula)
                    {
                        InsertNativeTextRun(document, selection, run);
                        continue;
                    }
                    if (!prepared.TryGetValue(run.Id, out var formula))
                        throw new InvalidDataException(
                            $"缺少行内公式 {run.Id} 的渲染结果。");

                    // Write the complete native paragraph before materializing
                    // inline formulas. Word keeps a caret collapsed at an OMML
                    // range end inside the math zone, so typing the following
                    // text immediately would absorb that text into <m:oMath>.
                    // Replacing one-character placeholders from right to left
                    // also ensures paragraph/list formatting is applied before
                    // OLE baseline offsets are calculated and persisted.
                    var placeholderStart = selection.Start;
                    selection.TypeText(BulkInlineFormulaPlaceholder);
                    pendingInlineFormulas.Add((placeholderStart, formula));
                }
                selection.TypeParagraph();
                var paragraphEnd = selection.Start;
                ApplyBulkParagraphFormatting(
                    document,
                    paragraphStart,
                    paragraphEnd,
                    block);

                for (var formulaIndex = pendingInlineFormulas.Count - 1;
                     formulaIndex >= 0;
                     formulaIndex--)
                {
                    var pending = pendingInlineFormulas[formulaIndex];
                    selection.SetRange(pending.Start, pending.Start + BulkInlineFormulaPlaceholder.Length);
                    selection.Text = string.Empty;
                    selection.Collapse(WdCollapseDirection.wdCollapseStart);
                    InsertPreparedFormula(
                        document,
                        selection,
                        pending.Formula,
                        display: false,
                        ommlBatchSource: ommlBatchSource,
                        deferredOmmlMetadata: ommlBatchSource is null
                            ? null
                            : deferredOmmlMetadata,
                        bulkImport: true);
                    insertedFormulaIds.Add(pending.Formula.Session.FormulaId);
                }

                MoveSelectionAfterBulkParagraph(document, selection, paragraphStart);
                ResetNextParagraphFormatting(selection, block.Kind, nextKind);
            }
            }

            if (deferredOmmlMetadata.Count > 0)
            {
                var metadataStopwatch = Stopwatch.StartNew();
                WordOmmlFormulaStore.SaveNewBatch(document, deferredOmmlMetadata);
                metadataStopwatch.Stop();
                if (string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                    Console.WriteLine(
                        $"    [perf] BulkOmml.SaveNewBatch: {metadataStopwatch.ElapsedMilliseconds}ms formulas={deferredOmmlMetadata.Count}");
            }

            return new WordBulkInsertResult
            {
                BlockCount = source.Blocks.Count,
                FormulaCount = insertedFormulaIds.Count,
                FormulaIds = insertedFormulaIds,
            };
        }
        catch
        {
            if (document is not null)
            {
                foreach (var formulaId in insertedFormulaIds)
                {
                    try { RemoveInlineBaselineSentinel(document, formulaId); } catch { }
                    try { WordOmmlFormulaStore.Delete(document, formulaId); } catch { }
                }
                if (insertionStart >= 0)
                {
                    try
                    {
                        var contentEnd = document.Content.End;
                        var rollbackEnd = selection is null
                            ? insertionStart
                            : Math.Min(Math.Max(insertionStart, selection.Start), contentEnd);
                        rollbackRange = document.Range(insertionStart, rollbackEnd);
                        rollbackRange.Delete();
                    }
                    catch { }
                }
            }
            throw;
        }
        finally
        {
            try { document?.Activate(); } catch { }
            ommlBatchSource?.Dispose();
            if (screenUpdatingSuspended)
            {
                try { _application.ScreenUpdating = previousScreenUpdating; } catch { }
            }
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(rollbackRange);
            Release(selection);
            Release(document);
        }
    }

    private void InsertBulkOmmlDocumentOneShot(
        Document document,
        Selection selection,
        WordBulkImportDocument source,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared,
        ICollection<string> insertedFormulaIds,
        ICollection<FormulaMetadata> deferredOmmlMetadata)
    {
        const string wordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string mathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        XNamespace word = wordNamespace;
        XNamespace math = mathNamespace;
        var marker = "VTBIEND_" + Guid.NewGuid().ToString("N");
        var body = new XElement(word + "body");
        var orderedFormulas = new List<(
            PreparedWordBulkFormula Formula,
            bool Display)>();

        void AppendFormula(
            XElement paragraph,
            PreparedWordBulkFormula formula,
            bool display)
        {
            var mathMl = formula.MathMl;
            if (string.IsNullOrWhiteSpace(mathMl))
                throw new InvalidDataException(
                    $"公式 {formula.Session.FormulaId} 没有可用的 MathML。");
            var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl!);
            var equation = XElement.Parse(
                WordOmmlConverter.ExtractSingleOMath(omml),
                LoadOptions.PreserveWhitespace);
            var metadata = formula.Session.ToMetadata();
            metadata.Validate();
            ApplyBulkOmmlTypographyXml(
                equation,
                formula.Session.FontSizePt,
                metadata,
                word,
                math);
            paragraph.Add(equation);
            orderedFormulas.Add((formula, display));
        }

        foreach (var block in source.Blocks)
        {
            var paragraph = new XElement(word + "p");
            var paragraphProperties =
                BuildBulkOmmlParagraphPropertiesXml(block, word);
            if (paragraphProperties is not null)
                paragraph.Add(paragraphProperties);

            if (block.Kind == WordBulkBlockKind.DisplayFormula)
            {
                var formulaRun = block.Runs.Single(run => run.IsFormula);
                if (!prepared.TryGetValue(formulaRun.Id, out var formula))
                    throw new InvalidDataException(
                        $"缺少行间公式 {formulaRun.Id} 的渲染结果。");
                AppendFormula(paragraph, formula, display: true);
            }
            else
            {
                foreach (var run in block.Runs)
                {
                    if (!run.IsFormula)
                    {
                        paragraph.Add(BuildBulkOmmlTextRunXml(
                            run,
                            block.Kind,
                            word));
                        continue;
                    }
                    if (!prepared.TryGetValue(run.Id, out var formula))
                        throw new InvalidDataException(
                            $"缺少行内公式 {run.Id} 的渲染结果。");
                    AppendFormula(paragraph, formula, display: false);
                }
            }
            body.Add(paragraph);
        }
        body.Add(
            new XElement(
                word + "p",
                BuildBulkOmmlTextRunXml(
                    new WordBulkRun { Text = marker },
                    WordBulkBlockKind.Paragraph,
                    word)));
        var root = new XElement(
            word + "document",
            new XAttribute(XNamespace.Xmlns + "w", wordNamespace),
            new XAttribute(XNamespace.Xmlns + "m", mathNamespace),
            body);
        var documentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + root.ToString(SaveOptions.DisableFormatting);

        WordOmmlConverter.WholeDocumentSource? wholeDocumentSource = null;
        Range? insertionAnchor = null;
        Range? markerSearch = null;
        Microsoft.Office.Interop.Word.Find? find = null;
        Range? insertedRange = null;
        OMaths? maths = null;
        try
        {
            EnsureWritableParagraph(selection);
            insertionAnchor = selection.Range.Duplicate;
            insertionAnchor.Collapse(WdCollapseDirection.wdCollapseStart);
            var insertionStart = insertionAnchor.Start;
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                try
                {
                    File.WriteAllText(
                        Path.Combine(Path.GetTempPath(), "visualtex-bulk-omml-oneshot.xml"),
                        documentXml,
                        new UTF8Encoding(false));
                }
                catch { }
            }
            var sourceStopwatch = Stopwatch.StartNew();
            wholeDocumentSource = WordOmmlConverter.CreateWholeDocumentSource(
                _application,
                documentXml);
            sourceStopwatch.Stop();
            document.Activate();
            var insertStopwatch = Stopwatch.StartNew();
            insertedRange = wholeDocumentSource.Insert(
                document,
                insertionAnchor);
            insertStopwatch.Stop();

            markerSearch = insertedRange.Duplicate;
            find = markerSearch.Find;
            find.ClearFormatting();
            find.Text = marker;
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindStop;
            if (!find.Execute())
                throw new InvalidOperationException(
                    "Word 未能定位批量 OMML 导入的结束标记。");
            var insertionEnd = markerSearch.Start;
            markerSearch.Text = string.Empty;
            Release(insertedRange);
            insertedRange = document.Range(insertionStart, insertionEnd);
            ApplyBulkOmmlDestinationParagraphFormatting(
                document,
                insertedRange,
                source);
            maths = insertedRange.OMaths;
            if (maths.Count != orderedFormulas.Count)
                throw new InvalidDataException(
                    $"批量 OMML 一次性写入生成了 {maths.Count} 个公式，"
                    + $"预期 {orderedFormulas.Count} 个。");
            // Word normalizes imported OMML (especially aligned/equation-array
            // structures) after insertion. First finalize inline/display types,
            // then fingerprint the actual Word XML in one batch. Persisting the
            // pre-insertion transform fingerprint makes the very first reopen look
            // like a native Word edit and replaces the original aligned LaTeX with
            // a lossy matrix reconstruction.
            for (var index = 0; index < orderedFormulas.Count; index++)
            {
                OMath? mathObject = null;
                try
                {
                    mathObject = maths[index + 1];
                    var targetType = orderedFormulas[index].Display
                        ? WdOMathType.wdOMathDisplay
                        : WdOMathType.wdOMathInline;
                    if (mathObject.Type != targetType)
                        mathObject.Type = targetType;
                }
                finally { Release(mathObject); }
            }

            var insertedFingerprints = ComputeBulkInsertedOmmlFingerprints(
                insertedRange,
                orderedFormulas.Count);
            for (var index = 0; index < orderedFormulas.Count; index++)
            {
                var ordered = orderedFormulas[index];
                OMath? mathObject = null;
                Range? equationRange = null;
                Bookmark? bookmark = null;
                try
                {
                    mathObject = maths[index + 1];
                    equationRange = mathObject.Range.Duplicate;
                    var metadata = ordered.Formula.Session.ToMetadata();
                    metadata.Validate();
                    if (insertedFingerprints.Count == orderedFormulas.Count)
                        metadata.NativeOmmlFingerprint = insertedFingerprints[index];
                    else
                        WordOmmlNativeSource.StampFingerprint(metadata, equationRange);
                    bookmark = WordOmmlFormulaStore.Wrap(
                        document,
                        equationRange,
                        metadata,
                        replaceExisting: false);
                    deferredOmmlMetadata.Add(metadata);
                    insertedFormulaIds.Add(metadata.FormulaId);
                }
                finally
                {
                    Release(bookmark);
                    Release(equationRange);
                    Release(mathObject);
                }
            }
            selection.SetRange(insertionEnd, insertionEnd);
            selection.Collapse(WdCollapseDirection.wdCollapseStart);
            ResetSelectionTransientFormatting(selection);
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    [perf] BulkOmml.OneShotSource: {sourceStopwatch.ElapsedMilliseconds}ms; "
                    + $"copy={insertStopwatch.ElapsedMilliseconds}ms formulas={orderedFormulas.Count}");
        }
        finally
        {
            Release(maths);
            Release(insertedRange);
            Release(find);
            Release(markerSearch);
            Release(insertionAnchor);
            wholeDocumentSource?.Dispose();
        }
    }

    private static IReadOnlyList<string> ComputeBulkInsertedOmmlFingerprints(
        Range insertedRange,
        int expectedCount)
    {
        try
        {
            XNamespace math =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var document = XDocument.Parse(
                insertedRange.WordOpenXML,
                LoadOptions.PreserveWhitespace);
            var equations = document
                .Descendants(math + "oMath")
                .Where(equation => !equation.Ancestors(math + "oMath").Any())
                .ToArray();
            if (equations.Length != expectedCount)
                return Array.Empty<string>();
            return equations
                .Select(equation => WordOmmlConverter.ComputeOmmlFingerprint(
                    equation.ToString(SaveOptions.DisableFormatting)))
                .ToArray();
        }
        catch
        {
            // Exceptional fallback: stamp each exact equation range separately.
            // The normal path performs one WordOpenXML read for the whole batch.
            return Array.Empty<string>();
        }
    }

    private static void ApplyBulkOmmlDestinationParagraphFormatting(
        Document document,
        Range insertedRange,
        WordBulkImportDocument source)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = insertedRange.Paragraphs;
            if (paragraphs.Count < source.Blocks.Count)
                throw new InvalidDataException(
                    $"批量 OMML 一次性写入生成了 {paragraphs.Count} 个段落，"
                    + $"预期至少 {source.Blocks.Count} 个。");
            for (var index = 0; index < source.Blocks.Count; index++)
            {
                var block = source.Blocks[index];
                if (block.Kind is not (
                        WordBulkBlockKind.Heading
                        or WordBulkBlockKind.Bullet
                        or WordBulkBlockKind.Numbered))
                    continue;
                Release(paragraphRange);
                paragraphRange = null;
                Release(paragraph);
                paragraph = null;
                paragraph = paragraphs[index + 1];
                paragraphRange = paragraph.Range;
                ApplyBulkParagraphFormatting(
                    document,
                    paragraphRange.Start,
                    paragraphRange.End,
                    block);
            }
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static XElement? BuildBulkOmmlParagraphPropertiesXml(
        WordBulkBlock block,
        XNamespace word)
    {
        var properties = new XElement(word + "pPr");
        switch (block.Kind)
        {
            case WordBulkBlockKind.Heading:
                var headingLevel = Math.Max(1, Math.Min(4, block.Level <= 0 ? 1 : block.Level));
                properties.Add(
                    new XElement(
                        word + "pStyle",
                        new XAttribute(word + "val", $"Heading{headingLevel}")));
                break;
            case WordBulkBlockKind.Bullet:
                var bulletLevel = Math.Max(1, Math.Min(9, block.Level + 1));
                properties.Add(
                    new XElement(
                        word + "pStyle",
                        new XAttribute(
                            word + "val",
                            bulletLevel == 1 ? "ListBullet" : $"ListBullet{bulletLevel}")));
                break;
            case WordBulkBlockKind.Numbered:
                var numberLevel = Math.Max(1, Math.Min(9, block.Level + 1));
                properties.Add(
                    new XElement(
                        word + "pStyle",
                        new XAttribute(
                            word + "val",
                            numberLevel == 1 ? "ListNumber" : $"ListNumber{numberLevel}")));
                break;
            case WordBulkBlockKind.Quote:
                properties.Add(
                    new XElement(
                        word + "ind",
                        new XAttribute(word + "left", "360"),
                        new XAttribute(word + "right", "180")));
                break;
            case WordBulkBlockKind.Code:
                properties.Add(
                    new XElement(
                        word + "ind",
                        new XAttribute(word + "left", "360")),
                    new XElement(
                        word + "spacing",
                        new XAttribute(word + "before", "60"),
                        new XAttribute(word + "after", "60")));
                break;
            case WordBulkBlockKind.DisplayFormula:
                properties.Add(
                    new XElement(
                        word + "jc",
                        new XAttribute(word + "val", "center")));
                break;
        }
        return properties.HasElements ? properties : null;
    }

    private static XElement BuildBulkOmmlTextRunXml(
        WordBulkRun run,
        WordBulkBlockKind blockKind,
        XNamespace word)
    {
        var result = new XElement(word + "r");
        var properties = new XElement(word + "rPr");
        var bold = run.Bold;
        var italic = run.Italic || blockKind == WordBulkBlockKind.Quote;
        // A document fragment inserted at a formatted Word caret inherits the
        // destination paragraph mark unless every semantic toggle is explicit.
        // In particular, omitting <w:i> for ordinary prose makes an italic caret
        // turn the whole imported document italic. Emit both the enabled and
        // disabled forms so source semantics, never destination typing state,
        // determine the imported native Word text.
        properties.Add(
            new XElement(
                word + "b",
                new XAttribute(word + "val", bold ? "1" : "0")),
            new XElement(
                word + "bCs",
                new XAttribute(word + "val", bold ? "1" : "0")),
            new XElement(
                word + "i",
                new XAttribute(word + "val", italic ? "1" : "0")),
            new XElement(
                word + "iCs",
                new XAttribute(word + "val", italic ? "1" : "0")),
            new XElement(
                word + "strike",
                new XAttribute(word + "val", run.Strike ? "1" : "0")),
            new XElement(
                word + "dstrike",
                new XAttribute(word + "val", "0")),
            new XElement(
                word + "u",
                new XAttribute(word + "val", run.Underline ? "single" : "none")),
            new XElement(
                word + "vertAlign",
                new XAttribute(word + "val", "baseline")),
            new XElement(
                word + "position",
                new XAttribute(word + "val", "0")),
            new XElement(
                word + "vanish",
                new XAttribute(word + "val", "0")));
        if (run.Code || blockKind == WordBulkBlockKind.Code)
            properties.Add(
                new XElement(
                    word + "rFonts",
                    new XAttribute(word + "ascii", "Consolas"),
                    new XAttribute(word + "hAnsi", "Consolas"),
                    new XAttribute(word + "eastAsia", "Microsoft YaHei UI")));
        result.Add(properties);

        var normalized = (run.Text ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n');
        var text = new StringBuilder();
        void FlushText()
        {
            if (text.Length == 0) return;
            result.Add(
                new XElement(
                    word + "t",
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    text.ToString()));
            text.Clear();
        }
        foreach (var character in normalized)
        {
            if (character == '\n')
            {
                FlushText();
                result.Add(new XElement(word + "br"));
            }
            else if (character == '\t')
            {
                FlushText();
                result.Add(new XElement(word + "tab"));
            }
            else
            {
                text.Append(character);
            }
        }
        FlushText();
        return result;
    }

    private static void ApplyBulkOmmlTypographyXml(
        XElement equation,
        double fontSizePt,
        FormulaMetadata metadata,
        XNamespace word,
        XNamespace math)
    {
        var halfPoints = ((int)Math.Round(
            FormulaFontSize.Normalize(fontSizePt) * 2.0))
            .ToString(System.Globalization.CultureInfo.InvariantCulture);

        // Keep ordinary variables, digits, Greek letters, operators and all native
        // structures as genuine Office Math. Their glyph selection and spacing must
        // come from the document-level m:mathFont and its OpenType MATH table—not
        // from per-run w:rFonts, run splitting, or m:nor normal-text conversion.
        foreach (var mathRun in equation.DescendantsAndSelf(math + "r"))
        {
            var mathProperties = mathRun.Element(math + "rPr");
            var normalText = mathProperties?.Element(math + "nor") is not null;
            var properties = mathRun.Element(word + "rPr");
            if (properties is null)
            {
                properties = new XElement(word + "rPr");
                if (mathProperties is not null) mathProperties.AddAfterSelf(properties);
                else mathRun.AddFirst(properties);
            }

            if (!normalText)
            {
                // A native math run must inherit the document's Office Math font.
                // Explicit run fonts bypass that mechanism and can flatten MATH-table
                // italic correction, operator spacing and extensible constructions.
                properties.Element(word + "rFonts")?.Remove();
            }
            else if (ContainsChineseOmmlText(mathRun, math))
            {
                // m:nor is retained only when it already carries real text semantics
                // (for example MathML mtext / LaTeX \text{中文}). Select only the
                // East-Asian text face; do not override the mathematical ASCII runs.
                var fonts = properties.Element(word + "rFonts");
                if (fonts is null)
                {
                    fonts = new XElement(word + "rFonts");
                    properties.AddFirst(fonts);
                }
                fonts.SetAttributeValue(word + "eastAsiaTheme", null);
                fonts.SetAttributeValue(
                    word + "eastAsia",
                    ResolveOmmlChineseFont(metadata.FormulaChineseFont));
            }

            ApplyNativeOmmlSizeAndPosition(properties, halfPoints, word);
        }

        // Fraction bars, radicals, delimiters and other structures may carry
        // m:ctrlPr/w:rPr independently of visible runs. Size them consistently,
        // while removing any control-level font override so the same MATH font
        // drives glyph assembly and spacing throughout the complete OMath tree.
        foreach (var controlProperties in equation.DescendantsAndSelf(math + "ctrlPr"))
        {
            var properties = controlProperties.Element(word + "rPr");
            if (properties is null)
            {
                properties = new XElement(word + "rPr");
                controlProperties.Add(properties);
            }
            properties.Element(word + "rFonts")?.Remove();
            ApplyNativeOmmlSizeAndPosition(properties, halfPoints, word);
        }

        // Genuine Word display OMML obtains fraction, radical, matrix and large-
        // operator sizing from the OpenType MATH table. Never apply the retired
        // numbered-inline 1.5x numerator/denominator compensation here.

        // Numbering is paragraph structure, never mathematical content. Keep the
        // OMath tree free of generated '#(...)' wrappers so Word does not turn a
        // normal formula into a full-width equation-array control. The visible REF
        // number is created outside OMath by WordEquationNumbering, exactly like the
        // accepted VisualTeX/MathType OLE tab layout.
    }

    private static bool ContainsChineseOmmlText(
        XElement mathRun,
        XNamespace math)
    {
        var text = string.Concat(mathRun.Elements(math + "t").Select(item => item.Value));
        for (var index = 0; index < text.Length;)
        {
            var width = char.IsHighSurrogate(text[index])
                && index + 1 < text.Length
                && char.IsLowSurrogate(text[index + 1])
                    ? 2
                    : 1;
            var codePoint = width == 2
                ? char.ConvertToUtf32(text[index], text[index + 1])
                : text[index];
            if (IsChineseOmmlCodePoint(codePoint)) return true;
            index += width;
        }
        return false;
    }

    private static void ApplyNativeOmmlSizeAndPosition(
        XElement properties,
        string halfPoints,
        XNamespace word)
    {
        var size = properties.Element(word + "sz");
        if (size is null)
        {
            size = new XElement(word + "sz");
            properties.Add(size);
        }
        size.SetAttributeValue(word + "val", halfPoints);

        var complexSize = properties.Element(word + "szCs");
        if (complexSize is null)
        {
            complexSize = new XElement(word + "szCs");
            properties.Add(complexSize);
        }
        complexSize.SetAttributeValue(word + "val", halfPoints);

        var position = properties.Element(word + "position");
        if (position is null)
        {
            position = new XElement(word + "position");
            properties.Add(position);
        }
        position.SetAttributeValue(word + "val", "0");
    }

    private void InsertBulkOleDocumentTwoPhase(
        Document document,
        Selection selection,
        WordBulkImportDocument source,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared,
        ICollection<string> insertedFormulaIds)
    {
        var pendingFormulas = new List<(
            int Start,
            PreparedWordBulkFormula Formula,
            bool Display)>();
        var endBookmarkName = "VTBI_" + Guid.NewGuid().ToString("N");
        Bookmarks? bookmarks = null;
        Bookmark? endBookmark = null;
        Range? endRange = null;
        try
        {
            for (var blockIndex = 0; blockIndex < source.Blocks.Count; blockIndex++)
            {
                var block = source.Blocks[blockIndex];
                var nextKind = blockIndex + 1 < source.Blocks.Count
                    ? source.Blocks[blockIndex + 1].Kind
                    : (WordBulkBlockKind?)null;
                if (block.Kind == WordBulkBlockKind.DisplayFormula)
                {
                    var formulaRun = block.Runs.Single(run => run.IsFormula);
                    if (!prepared.TryGetValue(formulaRun.Id, out var formula))
                        throw new InvalidDataException(
                            $"缺少行间公式 {formulaRun.Id} 的渲染结果。");
                    EnsureWritableParagraph(selection);
                    var placeholderStart = selection.Start;
                    selection.TypeText(BulkInlineFormulaPlaceholder);
                    pendingFormulas.Add((placeholderStart, formula, true));
                    selection.TypeParagraph();
                    ResetNextParagraphFormatting(selection, block.Kind, nextKind);
                    continue;
                }

                EnsureWritableParagraph(selection);
                var paragraphStart = selection.Start;
                foreach (var run in block.Runs)
                {
                    if (!run.IsFormula)
                    {
                        InsertNativeTextRun(document, selection, run);
                        continue;
                    }
                    if (!prepared.TryGetValue(run.Id, out var formula))
                        throw new InvalidDataException(
                            $"缺少行内公式 {run.Id} 的渲染结果。");
                    var placeholderStart = selection.Start;
                    selection.TypeText(BulkInlineFormulaPlaceholder);
                    pendingFormulas.Add((placeholderStart, formula, false));
                }
                selection.TypeParagraph();
                var paragraphEnd = selection.Start;
                ApplyBulkParagraphFormatting(
                    document,
                    paragraphStart,
                    paragraphEnd,
                    block);
                ResetNextParagraphFormatting(selection, block.Kind, nextKind);
            }

            endRange = selection.Range.Duplicate;
            endRange.Collapse(WdCollapseDirection.wdCollapseStart);
            bookmarks = document.Bookmarks;
            endBookmark = bookmarks.Add(endBookmarkName, endRange);

            foreach (var pending in pendingFormulas
                         .OrderByDescending(item => item.Start))
            {
                Range? preservedDisplayParagraphRange = null;
                Range? selectionRange = null;
                try
                {
                    selection.SetRange(
                        pending.Start,
                        pending.Start + BulkInlineFormulaPlaceholder.Length);
                    selection.Text = string.Empty;
                    selection.Collapse(WdCollapseDirection.wdCollapseStart);
                    if (pending.Display)
                    {
                        selectionRange = selection.Range;
                        preservedDisplayParagraphRange =
                            DuplicateContainingParagraphRange(selectionRange);
                    }
                    InsertPreparedFormula(
                        document,
                        selection,
                        pending.Formula,
                        display: pending.Display,
                        preserveExistingDisplayParagraphBoundary: pending.Display,
                        preservedDisplayParagraphRange: preservedDisplayParagraphRange,
                        bulkImport: true);
                    insertedFormulaIds.Add(pending.Formula.Session.FormulaId);
                }
                finally
                {
                    Release(selectionRange);
                    Release(preservedDisplayParagraphRange);
                }
            }

            if (bookmarks.Exists(endBookmarkName))
            {
                Release(endBookmark);
                endBookmark = bookmarks[endBookmarkName];
                Release(endRange);
                endRange = endBookmark.Range;
                selection.SetRange(endRange.Start, endRange.End);
                selection.Collapse(WdCollapseDirection.wdCollapseStart);
                endBookmark.Delete();
            }
        }
        finally
        {
            if (endBookmark is not null)
            {
                try { endBookmark.Delete(); } catch { }
            }
            Release(endRange);
            Release(endBookmark);
            Release(bookmarks);
        }
    }

    private void InsertPreparedFormula(
        Document document,
        Selection selection,
        PreparedWordBulkFormula prepared,
        bool display,
        bool preserveExistingDisplayParagraphBoundary = false,
        Range? preservedDisplayParagraphRange = null,
        string? preservedFollowingParagraphText = null,
        WordOmmlConverter.BatchSource? ommlBatchSource = null,
        ICollection<FormulaMetadata>? deferredOmmlMetadata = null,
        bool bulkImport = false)
    {
        var session = prepared.Session;
        session.DisplayMode = display ? "block" : "inline";
        session.Numbered = false;
        var metadata = session.ToMetadata();
        metadata.Validate();
        var nativeOmml = string.Equals(
            session.ObjectMode,
            FormulaOleContract.WordOmmlMode,
            StringComparison.Ordinal);
        var mathTypeOle = string.Equals(
            session.ObjectMode,
            FormulaOleContract.MathTypeOleMode,
            StringComparison.Ordinal);
        Range? insertion = null;
        Range? equationRange = null;
        string sourceFingerprint = string.Empty;
        Bookmark? bookmark = null;
        InlineShape? shape = null;
        try
        {
            if (mathTypeOle)
            {
                if (string.IsNullOrWhiteSpace(prepared.MathMl))
                    throw new InvalidDataException(
                        $"公式 {metadata.FormulaId} 没有可用于 MathType 的 MathML。" );
                if (string.IsNullOrWhiteSpace(prepared.EmfPath))
                    throw new InvalidDataException(
                        $"公式 {metadata.FormulaId} 没有可用于 MathType 的矢量预览。" );
                var nativePreview = prepared.MathTypeNativePreview;
                InsertMathTypeOle(
                    session,
                    prepared.MathMl!,
                    prepared.EmfPath,
                    isolatedNativePreviewWmfPath: nativePreview?.WmfPath,
                    isolatedNativePreviewWidthPt: nativePreview?.WidthPt ?? 0,
                    isolatedNativePreviewHeightPt: nativePreview?.HeightPt ?? 0,
                    isolatedNativePreviewWordPosition: nativePreview?.WordPosition ?? 0,
                    isolatedNativePreviewAttempted: prepared.MathTypeNativePreviewAttempted);
                if (display
                    && preserveExistingDisplayParagraphBoundary
                    && preservedDisplayParagraphRange is not null)
                {
                    RemoveGeneratedParagraphAfterMathTypeRedraw(
                        document,
                        preservedDisplayParagraphRange,
                        preservedFollowingParagraphText);
                }
                return;
            }

            if (nativeOmml)
                ApplyDocumentOmmlMathFont(document, metadata);

            var usePreservedDisplayParagraph =
                display
                && preserveExistingDisplayParagraphBoundary
                && preservedDisplayParagraphRange is not null;
            if (display)
            {
                if (!nativeOmml)
                {
                    Range? spacingAnchor = null;
                    try
                    {
                        spacingAnchor = usePreservedDisplayParagraph
                            ? preservedDisplayParagraphRange!.Duplicate
                            : selection.Range.Duplicate;
                        CompactParagraphBeforeOleDisplayFormula(document, spacingAnchor);
                    }
                    finally { Release(spacingAnchor); }
                }
                if (usePreservedDisplayParagraph)
                    FormatExistingDisplayParagraph(
                        preservedDisplayParagraphRange!,
                        preserveNativeOmmlSpacing: nativeOmml);
                else
                    EnsureBlankDisplayParagraph(
                        selection,
                        preserveNativeOmmlSpacing: nativeOmml);
            }
            insertion = usePreservedDisplayParagraph
                ? preservedDisplayParagraphRange!.Duplicate
                : selection.Range.Duplicate;
            if (!usePreservedDisplayParagraph)
                insertion.Collapse(WdCollapseDirection.wdCollapseEnd);

            if (nativeOmml)
            {
                using var nativeOmmlScreenUpdating =
                    NativeOmmlScreenUpdatingScope.Suspend(_application);
                var ommlPerformance = string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                    ? Stopwatch.StartNew()
                    : null;
                long ommlCheckpoint = 0;
                void TraceOmmlStage(string stage)
                {
                    if (ommlPerformance is null) return;
                    var elapsed = ommlPerformance.ElapsedMilliseconds;
                    Console.WriteLine(
                        $"    [perf] BulkOmml.Formula.{stage}: +{elapsed - ommlCheckpoint}ms ({elapsed}ms total) display={display}");
                    ommlCheckpoint = elapsed;
                }
                var mathMl = prepared.MathMl;
                if (string.IsNullOrWhiteSpace(mathMl))
                    throw new InvalidDataException(
                        $"公式 {metadata.FormulaId} 没有可用的 MathML。" );
                if (!display)
                {
                    var placeholder = PrepareInlineBaselineSentinelBeforeInsert(
                        document,
                        insertion,
                        metadata.FormulaId,
                        createBookmark: !bulkImport);
                    Release(insertion);
                    insertion = placeholder;
                }
                TraceOmmlStage("prepare-boundary");
                equationRange = ommlBatchSource is not null
                    ? ommlBatchSource.Insert(
                        document,
                        insertion,
                        metadata.FormulaId,
                        display,
                        sourceFingerprint: out sourceFingerprint,
                        replaceTarget: !display)
                    : WordOmmlConverter.Insert(
                        _application,
                        document,
                        insertion,
                        mathMl!,
                        display,
                        sourceFingerprint: out sourceFingerprint,
                        replaceTarget: !display,
                        transformOmml: omml => ApplyOmmlTypographyXml(
                            omml,
                            session.FontSizePt,
                            metadata),
                        mathFontName: document.OMathFontName);
                TraceOmmlStage("insert");
                ApplyOmmlTypography(equationRange, session.FontSizePt, metadata);
                TraceOmmlStage("font-size");
                metadata.NativeOmmlFingerprint = sourceFingerprint;
                bookmark = WordOmmlFormulaStore.Wrap(
                    document,
                    equationRange,
                    metadata,
                    replaceExisting: !bulkImport);
                TraceOmmlStage("wrap-bookmark");
                if (ommlBatchSource is not null && deferredOmmlMetadata is not null)
                    deferredOmmlMetadata.Add(metadata);
                // Do not persist the converter-side fingerprint yet. The native
                // OMath can still change during boundary/layout finalization.
                TraceOmmlStage("metadata-deferred");
                if (display)
                {
                    using var targetedNumberingMutation =
                        WordEquationNumbering.BeginTargetedNumberingMutation(
                            metadata.FormulaId,
                            metadata.Numbered);
                    TryReconcileOmml(document, bookmark, equationRange, metadata);
                    if (!preserveExistingDisplayParagraphBoundary)
                        MoveSelectionAfterDisplayFormula(selection, equationRange);
                }
                else if (bulkImport)
                {
                    RemoveBulkInlineOmmlTemporaryBoundary(
                        document,
                        equationRange,
                        metadata.FormulaId);
                }
                else
                {
                    FinalizeInlineOmmlBoundary(
                        document,
                        equationRange,
                        metadata.FormulaId,
                        moveCaretOutsideMath: true);
                }
                TraceOmmlStage("finalize-boundary");

                // The Word-native equation may be normalized after insertion by
                // BuildUp, typography, boundary cleanup, or display layout. Stamp
                // the live final OMath rather than retaining the converter's
                // provisional source fingerprint, and repair the collapsed anchor
                // to the equation's final start before the operation returns.
                WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                    metadata,
                    equationRange);
                if (!WordOmmlFormulaStore.IsCanonicalAnchor(bookmark, equationRange))
                {
                    Release(bookmark);
                    bookmark = WordOmmlFormulaStore.Wrap(
                        document,
                        equationRange,
                        metadata,
                        replaceExisting: true);
                }
                if (ommlBatchSource is null || deferredOmmlMetadata is null)
                    WordOmmlFormulaStore.SaveNew(document, metadata);
                TraceOmmlStage("finalize-identity");
                return;
            }

            if (!string.Equals(
                    session.ObjectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    $"批量导入不支持公式对象格式 {session.ObjectMode}。" );
            if (string.IsNullOrWhiteSpace(prepared.PngPath)
                || string.IsNullOrWhiteSpace(prepared.EmfPath))
                throw new InvalidDataException(
                    $"公式 {metadata.FormulaId} 没有可用的 OLE 预览。" );
            shape = AddOleObject(document, insertion);
            InitializeOle(shape, metadata, prepared.EmfPath!, prepared.PngPath!);
            Configure(
                shape,
                metadata,
                (session.ExportResult?.Width ?? 200) * 0.75f,
                (session.ExportResult?.Height ?? 60) * 0.75f,
                prepared.PngPath!,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                !display,
                nativeOleKnown: true,
                trustExportDimensions: bulkImport);
            if (display)
            {
                if (!bulkImport)
                    TryReconcileShape(document, shape, metadata);
                if (!preserveExistingDisplayParagraphBoundary)
                {
                    Range? shapeRange = null;
                    try
                    {
                        shapeRange = shape.Range;
                        MoveSelectionAfterDisplayFormula(selection, shapeRange);
                    }
                    finally { Release(shapeRange); }
                }
            }
            else if (!bulkImport)
            {
                RestoreTypingBaselineAfter(shape);
            }
            BindOleIdentityBookmark(shape, metadata.FormulaId);
        }
        catch
        {
            if (bookmark is not null) TryDelete(bookmark, deleteContents: true);
            else TryDelete(equationRange);
            TryDelete(shape);
            try { WordOmmlFormulaStore.Delete(document, metadata.FormulaId); } catch { }
            throw;
        }
        finally
        {
            Release(shape);
            Release(bookmark);
            Release(equationRange);
            Release(insertion);
        }
    }

    private static void MoveSelectionAfterBulkParagraph(
        Document document,
        Selection selection,
        int paragraphStart)
    {
        Range? anchor = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            anchor = document.Range(paragraphStart, paragraphStart);
            paragraphs = anchor.Paragraphs;
            if (paragraphs.Count == 0)
                throw new InvalidDataException("Word 未能定位批量导入段落。");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            selection.SetRange(paragraphRange.End, paragraphRange.End);
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(anchor);
        }
    }

    private static void EnsureWritableParagraph(Selection selection)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        try
        {
            paragraphs = selection.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            range = paragraph.Range;
            if (!ContainsVisibleBodyText(range.Text))
                return;
            if (selection.Start >= range.End - 1)
                selection.TypeParagraph();
        }
        finally
        {
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void CompactParagraphBeforeOleDisplayFormula(
        Document document,
        Range insertion)
    {
        Range? anchor = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = Math.Max(contentStart, document.Content.End - 1);
            var position = Math.Min(Math.Max(insertion.Start, contentStart), contentEnd);
            anchor = document.Range(position, position);
            paragraphs = anchor.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            // A caret at the start of an empty paragraph belongs to that empty
            // paragraph, while the display formula will still be visually tied
            // to the preceding prose. Resolve the preceding paragraph mark in
            // that case so the local spacing adjustment is applied to the text
            // the reader actually sees above the equation.
            if (!ContainsVisibleBodyText(paragraphRange.Text)
                && position > contentStart)
            {
                Release(paragraphRange);
                paragraphRange = null;
                Release(paragraph);
                paragraph = null;
                Release(paragraphs);
                paragraphs = null;
                Release(anchor);
                anchor = null;

                anchor = document.Range(position - 1, position - 1);
                paragraphs = anchor.Paragraphs;
                if (paragraphs.Count == 0) return;
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
            }

            if (!ContainsVisibleBodyText(paragraphRange.Text)) return;
            format = paragraph.Format;
            format.SpaceAfter = ParagraphBeforeOleDisplaySpaceAfterPoints;
        }
        finally
        {
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(anchor);
        }
    }

    private static Range DuplicateContainingParagraphRange(Range sourceRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = sourceRange.Paragraphs;
            if (paragraphs.Count == 0)
                throw new InvalidDataException("Word 未能定位行间公式所在段落。");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            var result = paragraphRange;
            paragraphRange = null;
            return result;
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static Range? DuplicateFollowingParagraphRange(
        Document document,
        Range paragraphRange)
    {
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? result = null;
        try
        {
            if (paragraphRange.End >= document.Content.End) return null;
            var start = paragraphRange.End;
            var end = Math.Min(document.Content.End, start + 1);
            probe = document.Range(start, end);
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            result = paragraph.Range.Duplicate;
            var value = result;
            result = null;
            return value;
        }
        finally
        {
            Release(result);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    private static void RemoveGeneratedParagraphAfterMathTypeRedraw(
        Document document,
        Range preservedFormulaParagraphRange,
        string? preservedFollowingParagraphText)
    {
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaRange = null;
        Range? nextProbe = null;
        Paragraphs? nextParagraphs = null;
        Paragraph? nextParagraph = null;
        Range? nextRange = null;
        InlineShapes? nextShapes = null;
        Fields? nextFields = null;
        Range? afterNextProbe = null;
        Paragraphs? afterNextParagraphs = null;
        Paragraph? afterNextParagraph = null;
        Range? afterNextRange = null;
        try
        {
            formulaParagraphs = preservedFormulaParagraphRange.Paragraphs;
            if (formulaParagraphs.Count != 1) return;
            formulaParagraph = formulaParagraphs[1];
            formulaRange = formulaParagraph.Range.Duplicate;
            var formulaShapes = formulaRange.InlineShapes;
            var hasMathTypeOle = false;
            try
            {
                for (var index = 1; index <= formulaShapes.Count; index++)
                {
                    InlineShape? shape = null;
                    try
                    {
                        shape = formulaShapes[index];
                        if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                        hasMathTypeOle = true;
                        break;
                    }
                    finally { Release(shape); }
                }
            }
            finally { Release(formulaShapes); }
            if (!hasMathTypeOle) return;
            if (formulaRange.End >= document.Content.End) return;

            nextProbe = document.Range(
                formulaRange.End,
                Math.Min(document.Content.End, formulaRange.End + 1));
            nextParagraphs = nextProbe.Paragraphs;
            if (nextParagraphs.Count != 1) return;
            nextParagraph = nextParagraphs[1];
            nextRange = nextParagraph.Range.Duplicate;
            if (ContainsVisibleBodyText(nextRange.Text)) return;
            nextShapes = nextRange.InlineShapes;
            nextFields = nextRange.Fields;
            if (nextShapes.Count > 0 || nextFields.Count > 0) return;
            if (preservedFollowingParagraphText is null) return;

            // If the source already had a blank paragraph immediately after the
            // display formula, this empty paragraph is user-authored and must stay.
            if (string.Equals(
                    nextRange.Text ?? string.Empty,
                    preservedFollowingParagraphText,
                    StringComparison.Ordinal))
                return;
            if (nextRange.End >= document.Content.End) return;

            afterNextProbe = document.Range(
                nextRange.End,
                Math.Min(document.Content.End, nextRange.End + 1));
            afterNextParagraphs = afterNextProbe.Paragraphs;
            if (afterNextParagraphs.Count != 1) return;
            afterNextParagraph = afterNextParagraphs[1];
            afterNextRange = afterNextParagraph.Range.Duplicate;
            if (!string.Equals(
                    afterNextRange.Text ?? string.Empty,
                    preservedFollowingParagraphText,
                    StringComparison.Ordinal))
                return;

            // Word's MathType Flat OPC generated exactly one extra empty paragraph
            // between the formula and the paragraph that originally followed the
            // LaTeX source. Delete only that proven generated paragraph.
            nextRange.Delete();
        }
        finally
        {
            Release(afterNextRange);
            Release(afterNextParagraph);
            Release(afterNextParagraphs);
            Release(afterNextProbe);
            Release(nextFields);
            Release(nextShapes);
            Release(nextRange);
            Release(nextParagraph);
            Release(nextParagraphs);
            Release(nextProbe);
            Release(formulaRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
        }
    }

    private static bool IsIncompleteMathTypeNumberRow(Range paragraphRange)
    {
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        var placeRefCount = 0;
        try
        {
            shapes = paragraphRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                if (MathTypeOleInterop.IsMathTypeOle(shape)) return false;
            }

            fields = paragraphRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    placeRefCount++;
            }
            if (placeRefCount != 1) return false;

            var text = paragraphRange.Text ?? string.Empty;
            foreach (var character in text)
            {
                if (character == '\r' || character == '\t'
                    || character == '\u0013' || character == '\u0014' || character == '\u0015'
                    || character == '\u0001' || character == '\uFFFC'
                    || char.IsWhiteSpace(character))
                    continue;
                return false;
            }
            return true;
        }
        catch { return false; }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(shape);
            Release(shapes);
        }
    }

    private static bool ClearIncompleteMathTypeNumberRowAtInsertion(
        Document document,
        Range insertion)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? body = null;
        try
        {
            paragraphs = insertion.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (insertion.Start < paragraphRange.Start || insertion.Start > paragraphRange.End)
                return false;
            if (!IsIncompleteMathTypeNumberRow(paragraphRange)) return false;

            var start = paragraphRange.Start;
            var bodyEnd = Math.Max(start, paragraphRange.End - 1);
            body = document.Range(start, bodyEnd);
            body.Delete();
            insertion.SetRange(start, start);
            return true;
        }
        finally
        {
            Release(body);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void RollbackStandaloneMathTypeDisplayInsertion(
        Document document,
        int insertionStart,
        int paragraphCountBeforePreparation)
    {
        Range? anchor = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? body = null;
        Range? paragraphMark = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = Math.Max(contentStart, document.Content.End - 1);
            var position = Math.Max(contentStart, Math.Min(insertionStart, contentEnd));
            anchor = document.Range(position, position);
            paragraphs = anchor.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            // A failed standalone MathType insertion owns the whole generated
            // display row. Delete every MathType OLE in that row first so a
            // partially completed PasteSpecial cannot survive the rollback.
            shapes = paragraphRange.InlineShapes;
            for (var index = shapes.Count; index >= 1; index--)
            {
                Release(shape);
                shape = shapes[index];
                if (MathTypeOleInterop.IsMathTypeOle(shape))
                    shape.Delete();
            }
            Release(shape);
            shape = null;
            Release(shapes);
            shapes = null;

            // Remove the MTPlaceRef/tabs left by the Flat OPC row.  Only clear a
            // row that is now structurally empty or a native number-only row;
            // never delete arbitrary user prose on an exception path.
            if (IsIncompleteMathTypeNumberRow(paragraphRange)
                || !ContainsVisibleBodyText(paragraphRange.Text))
            {
                var start = paragraphRange.Start;
                var end = Math.Max(start, paragraphRange.End - 1);
                body = document.Range(start, end);
                body.Delete();
            }

            // ResolveStandaloneMathTypeDisplayInsertionRange may have created one
            // dedicated blank paragraph before InsertXML. Restore the original
            // paragraph count when that paragraph was created solely for the failed
            // transaction; preserve a blank paragraph that already existed.
            if (paragraphCountBeforePreparation >= 0
                && ReadDocumentParagraphCount(document) > paragraphCountBeforePreparation)
            {
                Release(paragraphRange);
                paragraphRange = paragraph.Range;
                if (!ContainsVisibleBodyText(paragraphRange.Text))
                {
                    paragraphMark = document.Range(
                        Math.Max(paragraphRange.Start, paragraphRange.End - 1),
                        paragraphRange.End);
                    if (string.Equals(paragraphMark.Text, "\r", StringComparison.Ordinal))
                        paragraphMark.Delete();
                }
            }
        }
        catch
        {
            // Rollback is best-effort and must never mask the original Word/OLE
            // failure. The caller rethrows the original exception with stage data.
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(paragraphMark);
            Release(body);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(anchor);
        }
    }

    private static Range ResolveStandaloneMathTypeDisplayInsertionRange(
        Document document,
        Range anchor,
        bool replaceAtExactInsertion = false)
    {
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? paragraphShapes = null;
        Fields? paragraphFields = null;
        try
        {
            // A collapsed Range at a paragraph/table boundary is frequently
            // reported by Word as belonging to the PREVIOUS paragraph. That is
            // catastrophic during right-to-left format conversion: a following
            // block equation can then be appended to the previous MathType row.
            // Resolve ownership with a one-character forward probe so the anchor
            // is interpreted in the paragraph that actually owns the insertion
            // position in the remaining document.
            var contentStart = document.Content.Start;
            var probeContentEnd = document.Content.End;
            var position = Math.Max(
                contentStart,
                Math.Min(anchor.Start, Math.Max(contentStart, probeContentEnd - 1)));
            var probeEnd = Math.Min(probeContentEnd, position + 1);
            probe = document.Range(position, probeEnd);
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count == 0) return anchor.Duplicate;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            paragraphShapes = paragraphRange.InlineShapes;
            paragraphFields = paragraphRange.Fields;
            var hasInlineShapes = paragraphShapes.Count > 0;
            var hasFields = paragraphFields.Count > 0;
            if (IsMathTypeSectionStateParagraph(paragraphRange)
                && !hasInlineShapes
                && !ContainsVisibleBodyText(paragraphRange.Text))
            {
                // A collapsed Range exactly on a paragraph boundary is often
                // reported by Word as belonging to the previous paragraph.  When
                // that previous paragraph is MathType's hidden section state, the
                // user caret is logically in the following paragraph, so never
                // snap back to the hidden paragraph start.
                var nextParagraphStart = Math.Max(
                    document.Content.Start,
                    Math.Min(
                        paragraphRange.End,
                        Math.Max(document.Content.Start, document.Content.End - 1)));
                return document.Range(nextParagraphStart, nextParagraphStart);
            }
            if (!ContainsVisibleBodyText(paragraphRange.Text)
                && !hasInlineShapes
                && !hasFields)
                return document.Range(paragraphRange.Start, paragraphRange.Start);

            if (replaceAtExactInsertion)
            {
                // Format conversion has already deleted the original display host;
                // anchor.Start is the exact replacement position. At a paragraph
                // boundary Word reports the following paragraph as the owner of the
                // collapsed range. Inserting after that paragraph reverses adjacent
                // equations during right-to-left conversion. Reserve a fresh row
                // immediately BEFORE the following content instead.
                var exactStart = Math.Max(
                    document.Content.Start,
                    Math.Min(anchor.Start, Math.Max(document.Content.Start, document.Content.End - 1)));
                paragraphRange.InsertParagraphBefore();
                return document.Range(exactStart, exactStart);
            }

            // Flat OPC always carries its own <w:p>. Inserting it immediately before
            // the final paragraph mark does not automatically create a clean display
            // paragraph; Word can put the OLE beside existing prose. Create the one
            // required blank paragraph first, then materialize the MathType object
            // into that paragraph.
            var newParagraphStart = paragraphRange.End;
            paragraphRange.InsertParagraphAfter();
            var contentEnd = document.Content.End;
            newParagraphStart = Math.Max(
                document.Content.Start,
                Math.Min(newParagraphStart, Math.Max(document.Content.Start, contentEnd - 1)));
            return document.Range(newParagraphStart, newParagraphStart);
        }
        finally
        {
            Release(paragraphFields);
            Release(paragraphShapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    private static bool IsMathTypeSectionStateParagraph(Range range)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            fields = range.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static Range PrepareNumberedOmmlReplacementTabPlaceholderPreservingOle(
        Document document,
        InlineShape sourceShape)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? shapeRange = null;
        Range? before = null;
        Range? after = null;
        Range? scaffold = null;
        Range? placeholder = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            shapeRange = sourceShape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The numbered OLE replacement no longer occupies one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (shapeRange.Start < paragraphRange.Start
                || shapeRange.End > paragraphRange.End)
                throw new InvalidOperationException(
                    "The numbered OLE replacement range escaped its paragraph.");

            maths = paragraphRange.OMaths;
            shapes = paragraphRange.InlineShapes;
            fields = paragraphRange.Fields;
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                Console.WriteLine(
                    $"    [diagnostic] numbered OLE→OMML host: paragraph={paragraphRange.Start}:{paragraphRange.End}, shape={shapeRange.Start}:{shapeRange.End}, maths={maths.Count}, shapes={shapes.Count}, fields={fields.Count}.");
                for (var diagnosticIndex = 1; diagnosticIndex <= fields.Count; diagnosticIndex++)
                {
                    Field? diagnosticField = null;
                    Range? diagnosticCode = null;
                    Range? diagnosticResult = null;
                    try
                    {
                        diagnosticField = fields[diagnosticIndex];
                        diagnosticCode = diagnosticField.Code;
                        diagnosticResult = diagnosticField.Result;
                        Console.WriteLine(
                            $"    [diagnostic] numbered OLE→OMML field#{diagnosticIndex}: type={diagnosticField.Type}, code={diagnosticCode.Start}:{diagnosticCode.End} '{diagnosticCode.Text}', result={diagnosticResult.Start}:{diagnosticResult.End} '{diagnosticResult.Text}'.");
                    }
                    finally
                    {
                        Release(diagnosticResult);
                        Release(diagnosticCode);
                        Release(diagnosticField);
                    }
                }
            }
            if (maths.Count != 0
                || shapes.Count != 1
                || !ContainsOnlySourceOleFields(fields, sourceShape))
                throw new InvalidOperationException(
                    "VisualTeX refused to rebuild a numbered OLE paragraph containing another formula, object, or unrelated field.");

            var editableEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            if (shapeRange.End < editableEnd)
            {
                after = document.Range(shapeRange.End, editableEnd);
                if (ContainsVisibleBodyText(after.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text after the numbered OLE formula.");
                after.Text = string.Empty;
            }

            // Delete structural TAB/zero-width runs before the OLE only after the
            // trailing side has been removed, because deleting the prefix shifts the
            // InlineShape's live Word coordinates.
            Release(shapeRange);
            shapeRange = sourceShape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            if (paragraphRange.Start < shapeRange.Start)
            {
                before = document.Range(paragraphRange.Start, shapeRange.Start);
                if (ContainsVisibleBodyText(before.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text before the numbered OLE formula.");
                before.Text = string.Empty;
            }

            Release(shapeRange);
            shapeRange = sourceShape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            var start = paragraphRange.Start;
            scaffold = document.Range(start, start);
            scaffold.Text = "\t" + BulkInlineFormulaPlaceholder + "\t";
            placeholder = document.Range(
                start + 1,
                start + 1 + BulkInlineFormulaPlaceholder.Length);
            var result = placeholder;
            placeholder = null;
            return result;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(maths);
            Release(placeholder);
            Release(scaffold);
            Release(after);
            Release(before);
            Release(shapeRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool ContainsOnlySourceOleFields(
        Fields fields,
        InlineShape sourceShape)
    {
        if (fields.Count == 0) return true;
        Range? shapeRange = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        try
        {
            shapeRange = sourceShape.Range;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(result); result = null;
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                result = field.Result;
                var codeText = code.Text ?? string.Empty;
                var isEmbed = field.Type == WdFieldType.wdFieldEmbed
                    || codeText.IndexOf(
                        "EMBED ",
                        StringComparison.OrdinalIgnoreCase) >= 0;
                // For embedded OLE Word reports InlineShape.Range as the complete
                // field span including the begin/separator/end characters, while
                // Field.Code/Result exclude those boundary characters. Reconstruct
                // that full span instead of expecting Result alone to contain the
                // InlineShape range.
                var fullFieldStart = Math.Max(0, code.Start - 1);
                var fullFieldEnd = result.End + 1;
                var ownsSourceShape = fullFieldStart <= shapeRange.Start
                    && fullFieldEnd >= shapeRange.End;
                if (!isEmbed || !ownsSourceShape)
                    return false;
            }
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(shapeRange);
        }
    }

    private static Range PrepareNumberedOmmlTrueDisplayReplacementPlaceholder(
        Document document,
        Range equationRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? before = null;
        Range? after = null;
        Range? editableRange = null;
        Range? placeholder = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            paragraphs = equationRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The true-display numbered OMML replacement no longer occupies one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable)
                || equationRange.Start < paragraphRange.Start
                || equationRange.End > paragraphRange.End)
                throw new InvalidOperationException(
                    "The true-display numbered OMML replacement escaped its table-free formula paragraph.");

            maths = paragraphRange.OMaths;
            shapes = paragraphRange.InlineShapes;
            fields = paragraphRange.Fields;
            if (maths.Count != 1 || shapes.Count != 0 || fields.Count != 0)
                throw new InvalidOperationException(
                    "VisualTeX refused to replace a true-display OMML paragraph containing another formula, object, or field.");
            var paragraphXml = paragraphRange.WordOpenXML ?? string.Empty;
            if (paragraphXml.IndexOf("<m:oMathPara", StringComparison.OrdinalIgnoreCase) < 0
                || paragraphXml.IndexOf("<m:eqArr", StringComparison.OrdinalIgnoreCase) >= 0
                || paragraphXml.IndexOf("<w:fldChar", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new InvalidOperationException(
                    "The numbered OMML source is not a pure genuine-display paragraph.");

            if (paragraphRange.Start < equationRange.Start)
            {
                before = document.Range(paragraphRange.Start, equationRange.Start);
                if (ContainsVisibleBodyText(before.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text before the true-display OMML formula.");
            }
            var editableEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            if (equationRange.End < editableEnd)
            {
                after = document.Range(equationRange.End, editableEnd);
                if (ContainsVisibleBodyText(after.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text after the true-display OMML formula.");
            }

            var start = paragraphRange.Start;
            editableRange = document.Range(start, editableEnd);
            editableRange.Text = BulkInlineFormulaPlaceholder;
            placeholder = document.Range(
                start,
                start + BulkInlineFormulaPlaceholder.Length);
            var result = placeholder;
            placeholder = null;
            return result;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(maths);
            Release(placeholder);
            Release(editableRange);
            Release(after);
            Release(before);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static Range PrepareNumberedOmmlReplacementTabPlaceholder(
        Document document,
        Range equationRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? before = null;
        Range? after = null;
        Range? editableRange = null;
        Range? placeholder = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            paragraphs = equationRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The numbered OMML replacement no longer occupies one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (equationRange.Start < paragraphRange.Start
                || equationRange.End > paragraphRange.End)
                throw new InvalidOperationException(
                    "The numbered OMML replacement range escaped its paragraph.");

            maths = paragraphRange.OMaths;
            shapes = paragraphRange.InlineShapes;
            fields = paragraphRange.Fields;
            if (maths.Count != 1 || shapes.Count != 0 || fields.Count != 0)
                throw new InvalidOperationException(
                    "VisualTeX refused to rebuild a numbered OMML paragraph containing another formula, object, or field.");

            if (paragraphRange.Start < equationRange.Start)
            {
                before = document.Range(paragraphRange.Start, equationRange.Start);
                if (ContainsVisibleBodyText(before.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text before the numbered OMML formula.");
            }
            var editableEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            if (equationRange.End < editableEnd)
            {
                after = document.Range(equationRange.End, editableEnd);
                if (ContainsVisibleBodyText(after.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text after the numbered OMML formula.");
            }

            editableRange = document.Range(paragraphRange.Start, editableEnd);
            editableRange.Text = "\t" + BulkInlineFormulaPlaceholder + "\t";
            placeholder = document.Range(
                paragraphRange.Start + 1,
                paragraphRange.Start + 1 + BulkInlineFormulaPlaceholder.Length);
            var result = placeholder;
            placeholder = null;
            return result;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(maths);
            Release(placeholder);
            Release(editableRange);
            Release(after);
            Release(before);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static Range ResolveDisplayInsertionRange(
        Document document,
        Range anchor,
        bool replaceAtExactInsertion = false)
    {
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? content = null;
        try
        {
            content = document.Content;
            var lastInsertPosition = Math.Max(content.Start, content.End - 1);
            var exactPosition = Math.Max(
                content.Start,
                Math.Min(anchor.Start, lastInsertPosition));
            if (replaceAtExactInsertion)
            {
                // At a paragraph boundary a collapsed Word range can report the
                // paragraph on either side. Probe one character forward so an
                // in-place format conversion never appends the target formula to
                // the paragraph that originally followed the source equation.
                probe = document.Range(
                    exactPosition,
                    Math.Min(content.End, exactPosition + 1));
                paragraphs = probe.Paragraphs;
            }
            else
            {
                paragraphs = anchor.Paragraphs;
            }
            if (paragraphs.Count == 0)
                return anchor.Duplicate;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            var hasVisibleText = ContainsVisibleBodyText(paragraphRange.Text);
            if (replaceAtExactInsertion)
            {
                if (!hasVisibleText)
                    return document.Range(paragraphRange.Start, paragraphRange.Start);

                // The source display host has been removed completely and the
                // exact replacement point now coincides with the start of the
                // following user paragraph. Reserve one paragraph before that
                // content; this is the formula host, not an additional blank line.
                if (exactPosition <= paragraphRange.Start)
                {
                    paragraphRange.InsertParagraphBefore();
                    return document.Range(exactPosition, exactPosition);
                }
                throw new InvalidOperationException(
                    "VisualTeX refused to insert a converted display formula inside user paragraph text.");
            }

            // Reuse an existing empty paragraph instead of creating another one.
            // At the physical document end, however, paragraphRange.End lies one
            // position beyond the last legal insertion point. Clamping it back to
            // Content.End - 1 places the caret at the preceding OMath boundary and
            // Word absorbs the next display formula into that same equation. Create
            // the following paragraph explicitly and insert at its start instead.
            if (hasVisibleText && paragraphRange.End > lastInsertPosition)
            {
                var nextParagraphStart = paragraphRange.End;
                paragraphRange.InsertParagraphAfter();
                return document.Range(nextParagraphStart, nextParagraphStart);
            }

            // For ordinary body text with a following paragraph, its existing
            // boundary remains a stable display insertion point.
            var position = hasVisibleText
                ? paragraphRange.End
                : paragraphRange.Start;
            position = Math.Max(
                content.Start,
                Math.Min(position, lastInsertPosition));
            return document.Range(position, position);
        }
        finally
        {
            Release(content);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    private static void FormatExistingDisplayParagraph(
        Range paragraphRange,
        bool preserveNativeOmmlSpacing)
    {
        ParagraphFormat? format = null;
        try
        {
            format = paragraphRange.ParagraphFormat;
            format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            if (!preserveNativeOmmlSpacing)
            {
                format.SpaceBefore = 0;
                format.SpaceAfter = 0;
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            }
            try { paragraphRange.ListFormat.RemoveNumbers(); } catch { }
        }
        finally { Release(format); }
    }

    private static void EnsureBlankDisplayParagraph(
        Selection selection,
        bool preserveNativeOmmlSpacing)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        try
        {
            paragraphs = selection.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                range = paragraph.Range;
                if (ContainsVisibleBodyText(range.Text))
                    selection.TypeParagraph();
            }
            selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            if (!preserveNativeOmmlSpacing)
            {
                selection.ParagraphFormat.SpaceBefore = 0;
                selection.ParagraphFormat.SpaceAfter = 0;
                selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            }
        }
        finally
        {
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void MoveSelectionAfterNumberedDisplayFormula(
        Document document,
        Selection selection,
        Range equationRange,
        string formulaId)
    {
        Range? typingRange = null;
        Range? ownerRange = null;
        Range? captionRange = null;
        Range? content = null;
        try
        {
            typingRange =
                WordEquationNumbering.EnsureNormalTypingParagraphAfterNumberedDisplay(
                    document,
                    formulaId);
            if (typingRange is not null)
            {
                selection.SetRange(typingRange.Start, typingRange.Start);
                selection.Collapse(WdCollapseDirection.wdCollapseStart);
                return;
            }

            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(
                document,
                formulaId);
            captionRange = WordEquationNumbering.FindNativeEquationCaptionRange(
                document,
                formulaId);
            content = document.Content;
            var target = equationRange.End;
            if (ownerRange is not null) target = Math.Max(target, ownerRange.End);
            if (captionRange is not null) target = Math.Max(target, captionRange.End);
            target = Math.Max(content.Start, Math.Min(target, content.End));
            selection.SetRange(target, target);
            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        }
        finally
        {
            Release(content);
            Release(captionRange);
            Release(ownerRange);
            Release(typingRange);
        }
    }

    private static int ReadParagraphStart(Range anchor)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = anchor.Paragraphs;
            if (paragraphs.Count == 0) return anchor.Start;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return paragraphRange.Start;
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool IsStructurallyEmptyParagraph(Range range)
    {
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            maths = range.OMaths;
            shapes = range.InlineShapes;
            fields = range.Fields;
            var text = (range.Text ?? string.Empty)
                .Trim('\r', '\a', '\v', '\f', '\t', ' ');
            return text.Length == 0
                && maths.Count == 0
                && shapes.Count == 0
                && fields.Count == 0;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(maths);
        }
    }

    private static void RepairPreservedDisplayParagraphBoundary(
        Document document,
        Range formulaRange,
        int sourceParagraphStart,
        int sourceParagraphCount)
    {
        if (sourceParagraphStart < 0 || sourceParagraphCount < 0) return;
        var currentParagraphCount = ReadDocumentParagraphCount(document);
        if (currentParagraphCount == sourceParagraphCount) return;
        if (currentParagraphCount != sourceParagraphCount + 1)
            throw new InvalidOperationException(
                $"Word changed the paragraph count unexpectedly while converting a display formula to OMML: before={sourceParagraphCount}, after={currentParagraphCount}.");

        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Range? content = null;
        Range? probe = null;
        Paragraphs? candidateParagraphs = null;
        Paragraph? candidateParagraph = null;
        Range? candidateRange = null;
        try
        {
            formulaParagraphs = formulaRange.Paragraphs;
            if (formulaParagraphs.Count != 1)
                throw new InvalidOperationException(
                    "The converted OMML display formula spans multiple paragraphs.");
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range;
            content = document.Content;

            if (formulaParagraphRange.Start > sourceParagraphStart)
            {
                var previousStart = Math.Max(content.Start, formulaParagraphRange.Start - 1);
                probe = document.Range(previousStart, formulaParagraphRange.Start);
                candidateParagraphs = probe.Paragraphs;
                if (candidateParagraphs.Count == 0)
                    throw new InvalidOperationException(
                        "Word inserted an OMML paragraph before the source boundary, but the residual source paragraph could not be resolved.");
                candidateParagraph = candidateParagraphs[1];
                candidateRange = candidateParagraph.Range;
                if (candidateRange.Start != sourceParagraphStart
                    || !IsStructurallyEmptyParagraph(candidateRange))
                    throw new InvalidOperationException(
                        "The paragraph before the converted OMML formula is not the empty source paragraph VisualTeX expected to repair.");
            }
            else if (formulaParagraphRange.Start == sourceParagraphStart)
            {
                if (formulaParagraphRange.End >= content.End)
                    throw new InvalidOperationException(
                        "Word added a display-formula paragraph split at the document end, but no residual paragraph is available to repair.");
                probe = document.Range(
                    formulaParagraphRange.End,
                    Math.Min(content.End, formulaParagraphRange.End + 1));
                candidateParagraphs = probe.Paragraphs;
                if (candidateParagraphs.Count == 0)
                    throw new InvalidOperationException(
                        "Word inserted an OMML paragraph after the source boundary, but the residual paragraph could not be resolved.");
                candidateParagraph = candidateParagraphs[1];
                candidateRange = candidateParagraph.Range;
                if (!IsStructurallyEmptyParagraph(candidateRange))
                    throw new InvalidOperationException(
                        "The paragraph after the converted OMML formula contains user content and was not removed.");
            }
            else
            {
                throw new InvalidOperationException(
                    "The converted OMML display formula moved before its captured source paragraph boundary.");
            }

            candidateRange.Delete();
        }
        finally
        {
            Release(candidateRange);
            Release(candidateParagraph);
            Release(candidateParagraphs);
            Release(probe);
            Release(content);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
        }

        var repairedParagraphCount = ReadDocumentParagraphCount(document);
        if (repairedParagraphCount != sourceParagraphCount)
            throw new InvalidOperationException(
                $"VisualTeX could not restore the display formula's original paragraph structure: expected={sourceParagraphCount}, actual={repairedParagraphCount}.");
    }

    private static void MoveSelectionAfterDisplayFormula(
        Selection selection,
        Range formulaRange)
    {
        selection.SetRange(formulaRange.End, formulaRange.End);
        selection.TypeParagraph();
        selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        selection.ParagraphFormat.LeftIndent = 0;
        selection.ParagraphFormat.FirstLineIndent = 0;
        object normal = WdBuiltinStyle.wdStyleNormal;
        try { selection.Range.set_Style(ref normal); } catch { }
        try { selection.Range.ListFormat.RemoveNumbers(); } catch { }
        ResetSelectionTransientFormatting(selection);
    }

    private static void ResetSelectionTransientFormatting(Selection selection)
    {
        Range? caret = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        Microsoft.Office.Interop.Word.Font? selectionFont = null;
        Microsoft.Office.Interop.Word.Font? paragraphMarkFont = null;
        try
        {
            selectionFont = selection.Font;
            ResetTransientFont(selectionFont);

            // A collapsed Word Selection can report neutral formatting while the
            // paragraph mark still stores the old direct italic/bold state. The
            // next typed character inherits the paragraph mark, not the temporary
            // Selection.Font value. Clear that mark only when the caret is at the
            // paragraph boundary used for subsequent typing.
            if (selection.Start != selection.End) return;
            caret = selection.Range;
            paragraphs = caret.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (selection.Start < paragraphRange.End - 1) return;
            paragraphMark = paragraphRange.Duplicate;
            paragraphMark.SetRange(
                Math.Max(paragraphRange.Start, paragraphRange.End - 1),
                paragraphRange.End);
            paragraphMarkFont = paragraphMark.Font;
            ResetTransientFont(paragraphMarkFont);
        }
        finally
        {
            Release(paragraphMarkFont);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(caret);
            Release(selectionFont);
        }
    }

    private static void ResetTransientFont(Microsoft.Office.Interop.Word.Font font)
    {
        font.Bold = 0;
        font.Italic = 0;
        font.StrikeThrough = 0;
        try { font.DoubleStrikeThrough = 0; } catch { }
        font.Underline = WdUnderline.wdUnderlineNone;
        font.Hidden = 0;
        font.Subscript = 0;
        font.Superscript = 0;
        font.Position = 0;
        try { font.AllCaps = 0; } catch { }
        try { font.SmallCaps = 0; } catch { }
    }

    private static void InsertNativeTextRun(
        Document document,
        Selection selection,
        WordBulkRun run)
    {
        if (string.IsNullOrEmpty(run.Text)) return;
        ResetSelectionTransientFormatting(selection);
        var start = selection.Start;
        selection.TypeText(run.Text);
        Range? inserted = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            inserted = document.Range(start, selection.Start);
            font = inserted.Font;
            font.Bold = run.Bold ? 1 : 0;
            font.Italic = run.Italic ? 1 : 0;
            font.StrikeThrough = run.Strike ? 1 : 0;
            try { font.DoubleStrikeThrough = 0; } catch { }
            font.Underline = run.Underline
                ? WdUnderline.wdUnderlineSingle
                : WdUnderline.wdUnderlineNone;
            font.Hidden = 0;
            font.Subscript = 0;
            font.Superscript = 0;
            font.Position = 0;
            try { font.AllCaps = 0; } catch { }
            try { font.SmallCaps = 0; } catch { }
            if (run.Code)
            {
                font.Name = "Consolas";
                try { font.NameAscii = "Consolas"; } catch { }
                try { font.NameFarEast = "Microsoft YaHei UI"; } catch { }
            }
        }
        finally
        {
            Release(font);
            Release(inserted);
            // Range formatting can update Word's collapsed typing state. Keep
            // placeholders, following runs and the next paragraph neutral; the
            // next source run reapplies its own explicit semantics.
            ResetSelectionTransientFormatting(selection);
        }
    }

    private static void ApplyBulkParagraphFormatting(
        Document document,
        int start,
        int end,
        WordBulkBlock block)
    {
        if (end < start) return;
        Range? range = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ListFormat? listFormat = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = document.Range(start, end);
            paragraphs = range.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            switch (block.Kind)
            {
                case WordBulkBlockKind.Heading:
                    object heading = block.Level switch
                    {
                        <= 1 => WdBuiltinStyle.wdStyleHeading1,
                        2 => WdBuiltinStyle.wdStyleHeading2,
                        3 => WdBuiltinStyle.wdStyleHeading3,
                        _ => WdBuiltinStyle.wdStyleHeading4,
                    };
                    try { range.set_Style(ref heading); } catch { }
                    break;
                case WordBulkBlockKind.Bullet:
                    listFormat = range.ListFormat;
                    listFormat.ApplyBulletDefault();
                    for (var level = 0; level < Math.Min(block.Level, 8); level++)
                        listFormat.ListIndent();
                    break;
                case WordBulkBlockKind.Numbered:
                    listFormat = range.ListFormat;
                    listFormat.ApplyNumberDefault();
                    for (var level = 0; level < Math.Min(block.Level, 8); level++)
                        listFormat.ListIndent();
                    break;
                case WordBulkBlockKind.Quote:
                    paragraph.LeftIndent = 18f;
                    paragraph.RightIndent = 9f;
                    font = range.Font;
                    font.Italic = 1;
                    break;
                case WordBulkBlockKind.Code:
                    font = range.Font;
                    font.Name = "Consolas";
                    try { font.NameAscii = "Consolas"; } catch { }
                    paragraph.LeftIndent = 18f;
                    paragraph.SpaceBefore = 3f;
                    paragraph.SpaceAfter = 3f;
                    break;
            }

        }
        finally
        {
            Release(font);
            Release(listFormat);
            Release(paragraph);
            Release(paragraphs);
            Release(range);
        }
    }

    private static void ResetNextParagraphFormatting(
        Selection selection,
        WordBulkBlockKind current,
        WordBulkBlockKind? next)
    {
        var continuingList =
            current == WordBulkBlockKind.Bullet && next == WordBulkBlockKind.Bullet
            || current == WordBulkBlockKind.Numbered && next == WordBulkBlockKind.Numbered;
        if (continuingList)
        {
            // Preserve list numbering/indentation, but never carry the previous
            // item's final bold/italic run into the next list item.
            ResetSelectionTransientFormatting(selection);
            return;
        }
        try { selection.Range.ListFormat.RemoveNumbers(); } catch { }
        selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        selection.ParagraphFormat.LeftIndent = 0;
        selection.ParagraphFormat.RightIndent = 0;
        selection.ParagraphFormat.FirstLineIndent = 0;
        object normal = WdBuiltinStyle.wdStyleNormal;
        try { selection.Range.set_Style(ref normal); } catch { }
        ResetSelectionTransientFormatting(selection);
    }

    public OfficeObjectResult ReplaceMathTypeOle(
        OfficeSessionDocument session,
        string mathMl,
        string emfPath)
    {
        if (string.IsNullOrWhiteSpace(mathMl)
            || !mathMl.TrimStart().StartsWith("<math", StringComparison.Ordinal))
            throw new InvalidDataException("VisualTeX did not provide valid MathML for MathType OLE.");
        if (string.IsNullOrWhiteSpace(emfPath) || !File.Exists(emfPath))
            throw new FileNotFoundException(
                "VisualTeX did not provide a valid MathType OLE vector preview.",
                emfPath);

        var metadata = session.ToMetadata();
        metadata.Validate();
        if (string.IsNullOrWhiteSpace(metadata.Latex))
            throw new InvalidDataException(
                "VisualTeX did not provide LaTeX source for the MathType OLE update.");

        Document? document = null;
        InlineShape? oldShape = null;
        InlineShape? replacement = null;
        Range? oldRange = null;
        Range? insertion = null;
        Range? finalSelection = null;
        UndoRecord? undoRecord = null;
        WordViewState? viewState = null;
        string? rollbackWordOpenXml = null;
        var rollbackStart = -1;
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        var oldDeleted = false;
        MathTypeDisplayParagraphLayout? displayParagraphLayout = null;
        MathTypeNativePreviewRenderer.Result? sourceNativePreview = null;
        MathTypeNativePreviewRenderer.Result? nativePreview = null;
        var sourceParagraphCount = -1;
        var sourceWasNumbered = false;
        var sourceNumberPosition = "right";
        var numberingLayoutChanged = false;
        MathTypeWordOpenXml.NumberTemplate? sourceNumberTemplate = null;
        var createdEditSectionBreakCodeStart = -1;
        var alignInline = string.Equals(
            session.DisplayMode,
            "inline",
            StringComparison.OrdinalIgnoreCase);
        try
        {
            undoRecord = BeginUndoRecord("VisualTeX Update MathType OLE Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            viewState = CaptureViewState();
            try
            {
                previousScreenUpdating = _application.ScreenUpdating;
                _application.ScreenUpdating = false;
                screenUpdatingSuspended = true;
            }
            catch { }

            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=resolve-source formulaId={session.FormulaId}");
            oldShape = FindMathTypeOleByRange(document, session.SourceObjectId)
                ?? throw new InvalidOperationException(
                    "The MathType OLE equation no longer exists at the captured Word location.");
            if (!MathTypeOleInterop.IsMathTypeOle(oldShape))
                throw new InvalidOperationException(
                    "The selected OLE object is no longer recognized as MathType Equation.DSMT4.");

            oldRange = oldShape.Range.Duplicate;
            var oldStart = oldRange.Start;
            var sourceCount = document.InlineShapes.Count;
            // InsertXML can split the host paragraph for both display and inline
            // Equation.DSMT4 objects. Capture this before deleting the source so an
            // inline re-edit can restore surrounding prose to the same paragraph.
            sourceParagraphCount = ReadDocumentParagraphCount(document);
            if (!alignInline)
            {
                displayParagraphLayout = CaptureMathTypeDisplayParagraphLayout(oldShape);
                sourceWasNumbered = MathTypeOleInterop.TryReadDisplayNumberPosition(
                    oldShape,
                    out sourceNumberPosition);
                numberingLayoutChanged = sourceWasNumbered != session.Numbered
                    || sourceWasNumbered
                        && session.Numbered
                        && !string.Equals(
                            sourceNumberPosition,
                            session.MathTypeNumberPosition,
                            StringComparison.OrdinalIgnoreCase);
                if (sourceWasNumbered)
                    sourceNumberTemplate = ReadMathTypePlaceRefTemplateForShape(
                        document,
                        oldShape,
                        sourceNumberPosition);
                WordDoubleClickHook.TraceMessage(
                    $"mathtype-replace-numbering source={sourceWasNumbered}:{sourceNumberPosition} target={session.Numbered}:{session.MathTypeNumberPosition} changed={numberingLayoutChanged}");
            }

            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=read-source-flat-opc formulaId={session.FormulaId}");
            var sourceFragment = MathTypeWordOpenXml.Read(oldShape);
            var originalProgId = sourceFragment.ProgId;
            var originalWordPosition = ReadInlineOleWordPosition(oldShape);
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=read-source-flat-opc-complete formulaId={session.FormulaId}");

            // Preserve the original Equation.DSMT4 CFB and replace only its MTEF.
            // Geometry is resolved from the rewritten MTEF below so a VisualTeX
            // re-edit uses the exact same MathType-native presentation model as
            // VisualTeX→MathType conversion and direct MathType insertion.
            var targetWidthPt = (float)Math.Max(
                1d,
                (session.ExportResult?.Width ?? 200d) * 0.75d);
            var targetHeightPt = (float)Math.Max(
                1d,
                (session.ExportResult?.Height ?? 60d) * 0.75d);
            var alignToWordTextBaseline = alignInline || session.Numbered;
            var targetWordPosition = alignToWordTextBaseline
                ? CalculateMathTypeOleWordPosition(
                    targetHeightPt,
                    session.ExportResult?.Height ?? 0f,
                    session.ExportResult?.Baseline)
                : 0;
            byte[] previewWmf;

            // Preserve the original Equation.DSMT4 CFB, replace only its MTEF
            // structure, and seed a fresh OLE presentation cache from VisualTeX's
            // current EMF. The source and result are serialized data; no MathType
            // COM server is needed for the semantic or visual update.
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=rewrite-cfb formulaId={session.FormulaId}");
            var rewritten = MathTypeOleStorage.RewriteMathTypeCompoundFile(
                sourceFragment.CompoundFile,
                mathMl,
                string.Equals(session.DisplayMode, "inline", StringComparison.Ordinal));
            var expectedMathTypeSignature = MathTypeMtefCodec.SemanticSignature(mathMl);
            var generatedMathMl = MathTypeOleStorage.ReadMathMl(rewritten.CompoundFile);
            var generatedLatex = MathMlToLatexConverter.Convert(generatedMathMl);
            if (!MathTypeMathMlRoundTripMatches(expectedMathTypeSignature, generatedMathMl))
                throw new InvalidDataException(
                    $"VisualTeX generated invalid MathType MTEF. Expected '{metadata.Latex}', actual '{generatedLatex}'.");

            // MathType edit must not switch back to VisualTeX/MathJax geometry.
            // Render both the source and rewritten MTEF through the isolated
            // MathPage sidecar.  The source native extent tells us whether Word
            // was displaying the original equation at a user/document-specific
            // scale; apply that same scale to the rewritten native presentation.
            // This keeps MathType's own glyph/spacing model and prevents an inline
            // edit from changing the object baseline merely because VisualTeX's
            // frontend export has different pixel geometry.
            var renderRoot = Path.GetDirectoryName(emfPath) ?? Path.GetTempPath();
            var nativePreviewInputs = new Dictionary<string, byte[]>(StringComparer.Ordinal)
            {
                ["source"] = ReadMathTypeMtefFromCompoundFile(sourceFragment.CompoundFile),
                ["rewritten"] = rewritten.Mtef,
            };
            var renderedAllNativePreviews =
                MathTypeNativePreviewRenderer.TryRenderBatch(
                    nativePreviewInputs,
                    renderRoot,
                    out var nativePreviews);
            nativePreviews.TryGetValue("source", out sourceNativePreview);
            nativePreviews.TryGetValue("rewritten", out nativePreview);
            if (renderedAllNativePreviews
                && sourceNativePreview is not null
                && nativePreview is not null)
            {
                var widthScale = CalculateMathTypeNativePresentationScale(
                    sourceFragment.WidthPt,
                    sourceNativePreview.WidthPt);
                var heightScale = CalculateMathTypeNativePresentationScale(
                    sourceFragment.HeightPt,
                    sourceNativePreview.HeightPt);
                previewWmf = File.ReadAllBytes(nativePreview.WmfPath);
                targetWidthPt = Math.Max(1f, nativePreview.WidthPt * widthScale);
                targetHeightPt = Math.Max(1f, nativePreview.HeightPt * heightScale);
                targetWordPosition = alignToWordTextBaseline
                    ? (int)Math.Round(
                        nativePreview.WordPosition * heightScale,
                        MidpointRounding.AwayFromZero)
                    : 0;
                WordDoubleClickHook.TraceMessage(
                    $"mathtype-replace-native-preview formulaId={session.FormulaId} "
                    + $"sourceWord={sourceFragment.WidthPt:0.###}x{sourceFragment.HeightPt:0.###}@{originalWordPosition} "
                    + $"sourceNative={sourceNativePreview.WidthPt:0.###}x{sourceNativePreview.HeightPt:0.###}@{sourceNativePreview.WordPosition} "
                    + $"rewrittenNative={nativePreview.WidthPt:0.###}x{nativePreview.HeightPt:0.###}@{nativePreview.WordPosition} "
                    + $"scale={widthScale:0.###}x{heightScale:0.###} "
                    + $"target={targetWidthPt:0.###}x{targetHeightPt:0.###}@{targetWordPosition}");
            }
            else
            {
                // Older MathType installations may provide only a 32-bit MathPage
                // library, and machines without MathPage have no native renderer at
                // all. Editing must still remain functional: use the already-rendered
                // VisualTeX EMF as a Word-owned presentation while preserving the
                // genuine MathType CFB/MTEF semantics. This is the same safe fallback
                // used by direct insertion and never activates the MathType server.
                previewWmf = MathTypeWordOpenXml.ConvertEnhancedMetafileToPlaceableWmf(
                    emfPath,
                    targetWidthPt,
                    targetHeightPt);
                targetWordPosition = alignToWordTextBaseline
                    ? CalculateMathTypeOleWordPosition(
                        targetHeightPt,
                        session.ExportResult?.Height ?? 0f,
                        session.ExportResult?.Baseline)
                    : 0;
                WordDoubleClickHook.TraceMessage(
                    $"mathtype-replace-preview-fallback formulaId={session.FormulaId} "
                    + $"sourceNative={sourceNativePreview is not null} rewrittenNative={nativePreview is not null} "
                    + $"target={targetWidthPt:0.###}x{targetHeightPt:0.###}@{targetWordPosition}");
            }

            // Rewrite the existing Equation.DSMT4 CFB and its external Word WMF
            // presentation in one offline Flat OPC transaction using the native
            // MathType geometry whenever MathPage is available.
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=rewrite-flat-opc formulaId={session.FormulaId}");
            var replacementWordOpenXml = MathTypeWordOpenXml.RewriteWithPlaceableWmf(
                sourceFragment.WordOpenXml,
                rewritten.CompoundFile,
                previewWmf,
                targetWidthPt,
                targetHeightPt);
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=rewrite-flat-opc-complete formulaId={session.FormulaId}");
            var rewrittenFragment = MathTypeWordOpenXml.Read(replacementWordOpenXml);
            var rewrittenMathMl = MathTypeOleStorage.ReadMathMl(rewrittenFragment.CompoundFile);
            var rewrittenLatex = MathMlToLatexConverter.Convert(rewrittenMathMl);
            if (!MathTypeMathMlRoundTripMatches(expectedMathTypeSignature, rewrittenMathMl))
                throw new InvalidDataException(
                    $"VisualTeX generated invalid MathType Flat OPC. Expected '{metadata.Latex}', actual '{rewrittenLatex}'.");

            rollbackWordOpenXml = sourceFragment.WordOpenXml;
            rollbackStart = oldStart;

            // An inline OLE object occupies exactly one Word character. Remove only
            // that character, then materialize the rewritten Flat OPC at the same
            // insertion point. Surrounding prose is untouched. If InsertXML fails,
            // the catch block restores the original serialized object.
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=delete-source formulaId={session.FormulaId}");
            oldShape.Delete();
            oldDeleted = true;
            insertion = document.Range(oldStart, oldStart);
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=insert-rewritten-flat-opc formulaId={session.FormulaId}");
            insertion.InsertXML(replacementWordOpenXml);
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=insert-rewritten-flat-opc-complete formulaId={session.FormulaId}");

            if (document.InlineShapes.Count != sourceCount)
                throw new InvalidOperationException(
                    "Word changed the inline OLE object count while replacing the MathType equation.");

            replacement = FindMathTypeOleByRange(
                document,
                $"{RangeReferencePrefix}{oldStart}:{oldStart + 1}")
                ?? throw new InvalidOperationException(
                    "Word materialized the rewritten Flat OPC, but VisualTeX could not resolve the replacement MathType equation.");

            if (sourceParagraphCount >= 0)
                RepairMathTypeInsertXmlParagraphSplit(
                    document,
                    replacement,
                    sourceParagraphCount);

            MathTypeDisplayParagraphLayout? detachedNumberLayout = null;
            if (!alignInline)
            {
                // Older VisualTeX builds could misclassify a MathType MTPlaceRef
                // numbered row as inline and leave the equation and its number in
                // two adjacent paragraphs.  Repair that already-damaged shape on
                // the next edit as well: merge only a clean, immediately-following
                // MTPlaceRef-only paragraph and recover MathType's center/right tab
                // stops from that numbering paragraph.
                detachedNumberLayout = RepairDetachedMathTypeNumberParagraph(
                    document,
                    replacement);
            }
            var displayLayoutToRestore = detachedNumberLayout ?? displayParagraphLayout;
            if (!alignInline && numberingLayoutChanged)
            {
                var targetNumberTemplate = session.Numbered
                    ? ResolveMathTypeEditNumberTemplate(
                        document,
                        replacement,
                        sourceWasNumbered ? sourceNumberTemplate : null,
                        out createdEditSectionBreakCodeStart)
                    : null;
                RebuildMathTypeDisplayScaffold(
                    document,
                    replacement,
                    session.Numbered,
                    session.MathTypeNumberPosition,
                    targetNumberTemplate);
                if (displayLayoutToRestore is not null)
                    RestoreMathTypeDisplayParagraphLayout(
                        replacement,
                        displayLayoutToRestore);
                MathTypeEquationNumbering.UpdateEquationNumbers(document);
                WordDoubleClickHook.TraceMessage(
                    $"mathtype-replace-numbering-rebuilt numbered={session.Numbered} position={session.MathTypeNumberPosition}");
            }
            else if (!alignInline && displayLayoutToRestore is not null)
            {
                RestoreMathTypeDisplayParagraphLayout(replacement, displayLayoutToRestore);
            }

            // Keep the U+0001 object-result baseline synchronized with the exact
            // same exported MathType geometry used to build the WMF. Reusing the
            // old object baseline after replacing the preview is what previously
            // let re-edited formulas acquire a second, incompatible layout model.
            SetInlineOleWordPosition(replacement, targetWordPosition);
            if (alignInline)
                RestoreTypingBaselineAfter(replacement);

            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=validate-replacement formulaId={session.FormulaId}");
            var replacementFragment = MathTypeWordOpenXml.Read(replacement);
            if (!string.IsNullOrWhiteSpace(originalProgId)
                && !string.Equals(
                    replacementFragment.ProgId,
                    originalProgId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    $"Word changed the MathType OLE ProgID from '{originalProgId}' to '{replacementFragment.ProgId}'.");
            var replacementMathMl = MathTypeOleStorage.ReadMathMl(replacementFragment.CompoundFile);
            var replacementLatex = MathMlToLatexConverter.Convert(replacementMathMl);
            if (!MathTypeMathMlRoundTripMatches(expectedMathTypeSignature, replacementMathMl))
                throw new InvalidDataException(
                    $"Word materialized the wrong MathType formula. Expected '{metadata.Latex}', actual '{replacementLatex}'.");

            finalSelection = replacement.Range.Duplicate;
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=complete formulaId={session.FormulaId}");
            return Result(session, document);
        }
        catch
        {
            if (oldDeleted && document is not null && rollbackStart >= 0
                && !string.IsNullOrWhiteSpace(rollbackWordOpenXml))
            {
                try
                {
                    TryDelete(replacement);
                    Range? rollbackRange = null;
                    InlineShape? rollbackShape = null;
                    try
                    {
                        rollbackRange = document.Range(rollbackStart, rollbackStart);
                        rollbackRange.InsertXML(rollbackWordOpenXml);
                        if (sourceParagraphCount >= 0)
                        {
                            rollbackShape = FindMathTypeOleByRange(
                                document,
                                $"{RangeReferencePrefix}{rollbackStart}:{rollbackStart + 1}");
                            if (rollbackShape is not null)
                            {
                                RepairMathTypeInsertXmlParagraphSplit(
                                    document,
                                    rollbackShape,
                                    sourceParagraphCount);
                                if (!alignInline && numberingLayoutChanged)
                                {
                                    RebuildMathTypeDisplayScaffold(
                                        document,
                                        rollbackShape,
                                        sourceWasNumbered,
                                        sourceNumberPosition,
                                        sourceWasNumbered ? sourceNumberTemplate : null);
                                    MathTypeEquationNumbering.UpdateEquationNumbers(document);
                                }
                                if (displayParagraphLayout is not null)
                                    RestoreMathTypeDisplayParagraphLayout(
                                        rollbackShape,
                                        displayParagraphLayout);
                            }
                        }
                    }
                    finally
                    {
                        Release(rollbackShape);
                        Release(rollbackRange);
                    }
                }
                catch { }
                if (createdEditSectionBreakCodeStart >= 0)
                {
                    try
                    {
                        RemoveMathTypeSectionBreakFieldAtCodeStart(
                            document,
                            createdEditSectionBreakCodeStart);
                    }
                    catch { }
                }
            }
            else
            {
                TryDelete(replacement);
            }
            throw;
        }
        finally
        {
            nativePreview?.Dispose();
            sourceNativePreview?.Dispose();
            if (screenUpdatingSuspended)
            {
                try { _application.ScreenUpdating = previousScreenUpdating; } catch { }
            }
            RestoreViewState(document, viewState, finalSelection);
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(finalSelection);
            Release(insertion);
            Release(oldRange);
            Release(replacement);
            Release(oldShape);
            Release(document);
        }
    }

    private static byte[] WaitForMathTypeOleMaterialization(
        InlineShape shape,
        string expectedLatex)
    {
        var delaysMs = new[] { 0, 15, 35, 70, 120, 200 };
        Exception? lastError = null;
        string? lastLatex = null;
        foreach (var delayMs in delaysMs)
        {
            if (delayMs > 0) Thread.Sleep(delayMs);
            try
            {
                var compound = MathTypeOleStorage.CaptureCompoundFile(shape);
                if (!MathTypeOleStorage.LooksLikeMathTypeCompoundFile(compound))
                {
                    lastError = new InvalidDataException(
                        "Word's replacement object is not yet a valid MathType compound storage.");
                    continue;
                }
                var mathMl = MathTypeOleStorage.ReadMathMl(compound);
                lastLatex = MathMlToLatexConverter.Convert(mathMl);
                if (MathTypeOleRoundTripMatches(expectedLatex, lastLatex))
                    return compound;
                lastError = new InvalidDataException(
                    $"MathType OLE materialization mismatch. Expected '{expectedLatex}', actual '{lastLatex}'.");
            }
            catch (Exception error)
            {
                lastError = error;
            }
        }

        throw new InvalidDataException(
            $"Word did not finish materializing the rewritten MathType OLE. Expected '{expectedLatex}', last='{lastLatex ?? "<unreadable>"}'.",
            lastError);
    }

    private static bool MathTypeMathMlRoundTripMatches(
        string expectedSignature,
        string actualMathMl) =>
        string.Equals(
            expectedSignature,
            MathTypeMtefCodec.SemanticSignature(actualMathMl),
            StringComparison.Ordinal);

    private static string DescribeSemanticSignatureDifference(
        string expected,
        string actual)
    {
        expected ??= string.Empty;
        actual ??= string.Empty;
        var common = Math.Min(expected.Length, actual.Length);
        var index = 0;
        while (index < common && expected[index] == actual[index]) index++;
        if (index == common && expected.Length == actual.Length)
            return $"identical(length={expected.Length})";

        static string DescribeAt(string value, int position)
        {
            if (position < 0 || position >= value.Length) return "<end>";
            var ch = value[position];
            return $"U+{(int)ch:X4}('{(char.IsControl(ch) ? '?' : ch)}')";
        }

        var start = Math.Max(0, index - 6);
        var expectedEnd = Math.Min(expected.Length, index + 7);
        var actualEnd = Math.Min(actual.Length, index + 7);
        static string Escape(string value) => string.Concat(value.Select(ch =>
            ch >= ' ' && ch <= '~' ? ch.ToString() : $"\\u{(int)ch:X4}"));
        return $"index={index}; expectedLen={expected.Length}; actualLen={actual.Length}; "
            + $"expectedChar={DescribeAt(expected, index)}; actualChar={DescribeAt(actual, index)}; "
            + $"expectedContext='{Escape(expected.Substring(start, expectedEnd - start))}'; "
            + $"actualContext='{Escape(actual.Substring(start, actualEnd - start))}'";
    }

    private static bool MathTypeOleRoundTripMatches(string expectedLatex, string actualLatex)
    {
        static string Normalize(string value) =>
            (value ?? string.Empty)
                .Replace(" ", string.Empty)
                .Replace("{", string.Empty)
                .Replace("}", string.Empty)
                .Trim();
        return string.Equals(
            Normalize(expectedLatex),
            Normalize(actualLatex),
            StringComparison.Ordinal);
    }

    private static InlineShape? FindNewMathTypeOleAtStart(
        Document document,
        int start,
        InlineShape sourceShape)
    {
        InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? range = null;
                var keepCandidate = false;
                try
                {
                    candidate = shapes[index];
                    range = candidate.Range;
                    if (range.Start != start
                        || IsSameComObject(candidate, sourceShape)
                        || !MathTypeOleInterop.IsMathTypeOle(candidate))
                        continue;
                    keepCandidate = true;
                    return candidate;
                }
                catch { }
                finally
                {
                    Release(range);
                    if (!keepCandidate) Release(candidate);
                }
            }
            return null;
        }
        finally { Release(shapes); }
    }

    private static bool IsSameComObject(object left, object right)
    {
        if (ReferenceEquals(left, right)) return true;
        if (!Marshal.IsComObject(left) || !Marshal.IsComObject(right)) return false;
        IntPtr leftIdentity = IntPtr.Zero;
        IntPtr rightIdentity = IntPtr.Zero;
        try
        {
            leftIdentity = Marshal.GetIUnknownForObject(left);
            rightIdentity = Marshal.GetIUnknownForObject(right);
            return leftIdentity == rightIdentity;
        }
        catch
        {
            return false;
        }
        finally
        {
            if (leftIdentity != IntPtr.Zero) Marshal.Release(leftIdentity);
            if (rightIdentity != IntPtr.Zero) Marshal.Release(rightIdentity);
        }
    }

    public OfficeObjectResult ReplaceOle(
        OfficeSessionDocument session,
        string pngPath,
        string emfPath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        InlineShape? oldShape = null;
        Bookmark? oldBookmark = null;
        Range? oldRange = null;
        Range? insertion = null;
        InlineShape? replacement = null;
        Table? numberedTable = null;
        Range? rollbackEquationRange = null;
        Bookmark? rollbackBookmark = null;
        UndoRecord? undoRecord = null;
        FormulaMetadata? originalMetadata = null;
        string? originalOmmlWordOpenXml = null;
        WordViewState? viewState = null;
        Range? finalSelection = null;
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        var oldStart = -1;
        var removedOmml = false;
        var sourceOmmlManagedByVisualTeX = false;
        var sourceWasNumberedManagedOmml = false;
        var capturedNativeDisplayAnchorParagraphStart = -1;
        var sourceIsMathTypeOle = false;
        var performanceWatch = Stopwatch.StartNew();
        long performanceCheckpoint = 0;
        try
        {
            undoRecord = BeginUndoRecord("VisualTeX Convert or Update Native OLE Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            viewState = CaptureViewState();
            try
            {
                previousScreenUpdating = _application.ScreenUpdating;
                _application.ScreenUpdating = false;
                screenUpdatingSuspended = true;
            }
            catch { }
            oldShape = FindByFormulaId(
                document,
                session.FormulaId,
                session.SourceObjectId);
            if (oldShape is null)
            {
                oldShape = FindMathTypeOleByRange(document, session.SourceObjectId);
                sourceIsMathTypeOle = oldShape is not null;
            }
            if (sourceIsMathTypeOle)
            {
                // A MathType source can already have an equation number owned by
                // MathType/Word in the surrounding paragraph.  Never add a second
                // VisualTeX numbering owner during the same edit/conversion, even
                // if a stale or non-UI client submits Numbered=true.
                metadata.Numbered = false;
                metadata.Validate();
            }
            if (!sourceIsMathTypeOle
                && session.DisplayMode == "block"
                && session.Numbered)
                numberedTable = WordEquationNumbering.FindNumberedEquationTable(
                    document,
                    session.FormulaId);
            TraceAcceptancePerformance(
                "ReplaceOle",
                "locate-target",
                performanceWatch,
                ref performanceCheckpoint);
            float oldWidth;
            float oldHeight;
            if (oldShape is not null)
            {
                originalMetadata = WordFormulaMetadataReader.TryRead(oldShape)
                    ?? session.OriginalMetadata;
                var preservesDisplayMode = string.Equals(
                    originalMetadata?.DisplayMode,
                    session.DisplayMode,
                    StringComparison.OrdinalIgnoreCase);
                if (preservesDisplayMode)
                {
                    oldWidth = oldShape.Width;
                    oldHeight = oldShape.Height;
                }
                else
                {
                    oldWidth = (float)Math.Max(
                        1,
                        (originalMetadata?.RenderWidthPx ?? session.ExportResult?.Width ?? 200) * 0.75);
                    oldHeight = (float)Math.Max(
                        1,
                        (originalMetadata?.RenderHeightPx ?? session.ExportResult?.Height ?? 60) * 0.75);
                }
            }
            else
            {
                oldBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    document,
                    session.FormulaId);
                if (oldBookmark is not null)
                {
                    sourceOmmlManagedByVisualTeX = true;
                    originalMetadata = WordOmmlFormulaStore.TryRead(document, oldBookmark)
                        ?? session.OriginalMetadata;
                    oldRange = ResolveOmmlEquationRange(
                        document,
                        oldBookmark,
                        session.SourceObjectId,
                        PreferSessionOmmlResolutionMetadata(
                            originalMetadata,
                            session.OriginalMetadata));
                }
                else
                {
                    originalMetadata = session.OriginalMetadata
                        ?? throw new InvalidOperationException(
                            "The selected Word-native OMML formula no longer exists.");
                    oldRange = ResolveStandaloneOmmlEquationRange(
                        document,
                        session.SourceObjectId,
                        originalMetadata);
                }
                originalOmmlWordOpenXml =
                    WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                        document,
                        oldRange,
                        originalMetadata!.FormulaId);
                if (session.DisplayMode == "inline"
                    && string.Equals(
                        originalMetadata?.DisplayMode,
                        "inline",
                        StringComparison.OrdinalIgnoreCase)
                    && TryReadStoredWordInlineOleSize(
                        originalMetadata,
                        out var storedInlineWidth,
                        out var storedInlineHeight))
                {
                    oldWidth = storedInlineWidth;
                    oldHeight = storedInlineHeight;
                }
                else
                {
                    oldWidth = (float)Math.Max(
                        1,
                        (originalMetadata?.RenderWidthPx ?? session.ExportResult?.Width ?? 200) * 0.75);
                    oldHeight = (float)Math.Max(
                        1,
                        (originalMetadata?.RenderHeightPx ?? session.ExportResult?.Height ?? 60) * 0.75);
                }
            }
            var editedSize = OfficeFormulaSizing.EditedSize(
                oldWidth,
                oldHeight,
                originalMetadata?.RenderWidthPx,
                originalMetadata?.RenderHeightPx,
                session.ExportResult?.Width ?? oldWidth / 0.75f,
                session.ExportResult?.Height ?? oldHeight / 0.75f,
                originalFontSizePt: originalMetadata?.FontSizePt,
                originalRenderFontSizePt: originalMetadata?.RenderFontSizePt);
            StoreWordInlineOleSize(
                metadata,
                editedSize.Width,
                editedSize.Height,
                session.DisplayMode == "inline");
            metadata.Validate();

            // Reusing an inline OLE object preserves its previous COM extent.
            // When the edited formula becomes wider or taller, Word can then
            // scale the new preview back into the old canvas and make every
            // glyph look smaller. Recreate inline objects so the OLE server is
            // initialized with the new natural extent; block objects may still
            // update in place because their outer layout is intentionally
            // controlled by the host paragraph/table.
            if (oldShape is not null
                && !sourceIsMathTypeOle
                && session.DisplayMode != "inline"
                && !FormulaFontPreferencesChanged(originalMetadata, metadata)
                && TryUpdateOle(oldShape, metadata, emfPath, pngPath))
            {
                TraceAcceptancePerformance(
                    "ReplaceOle",
                    "update-native-object",
                    performanceWatch,
                    ref performanceCheckpoint);
                Configure(
                    oldShape,
                    metadata,
                    editedSize.Width,
                    editedSize.Height,
                    pngPath,
                    session.ExportResult?.Height ?? 0,
                    session.ExportResult?.Baseline,
                    session.DisplayMode == "inline");
                TraceAcceptancePerformance(
                    "ReplaceOle",
                    "configure",
                    performanceWatch,
                    ref performanceCheckpoint);
                var numberingHostMayHaveChanged =
                    NumberingOrderMayHaveChanged(originalMetadata, metadata)
                    || numberedTable is not null;
                if (session.DisplayMode == "inline")
                    RestoreTypingBaselineAfter(oldShape);
                else
                    TryReconcileShape(
                        document,
                        oldShape,
                        metadata,
                        numberingHostMayHaveChanged,
                        reuseExistingNumberedTableFormatting: numberedTable is not null,
                        knownNumberedTable: numberedTable);
                TraceAcceptancePerformance(
                    "ReplaceOle",
                    "reconcile",
                    performanceWatch,
                    ref performanceCheckpoint);
                // Ordinary in-place edits keep this exact InlineShape alive and
                // must retain the fast direct-range path. Changing Numbered/display
                // mode, or upgrading a legacy 1x3 OLE host to MathType-style tabs,
                // invalidates the pre-reconcile RCW even though the formula itself
                // survived. Re-resolve only for those structural cases before
                // touching Range or the identity bookmark.
                if (numberingHostMayHaveChanged)
                {
                    Release(oldShape);
                    oldShape = FindByFormulaId(document, metadata.FormulaId)
                        ?? throw new InvalidOperationException(
                            "Word lost the updated VisualTeX OLE formula while changing its numbering host.");
                }
                finalSelection = oldShape.Range.Duplicate;
                BindOleIdentityBookmark(oldShape, metadata.FormulaId);
                TraceAcceptancePerformance(
                    "ReplaceOle",
                    "selection",
                    performanceWatch,
                    ref performanceCheckpoint);
                return Result(session, document);
            }

            if (oldRange is null)
                oldRange = oldShape is not null
                    ? oldShape.Range
                    : throw new InvalidOperationException(
                        "The selected Word OMML formula range is unavailable.");
            oldStart = oldRange.Start;
            if (oldShape is null
                && sourceOmmlManagedByVisualTeX
                && originalMetadata?.Numbered == true
                && string.Equals(
                    originalMetadata.DisplayMode,
                    "block",
                    StringComparison.OrdinalIgnoreCase))
            {
                // A numbered OMML source already owns VTEq_/VTEqCap_/VTEqNum_
                // ranges in this paragraph. Deleting the OMath first lets Word
                // expand the old visible-number bookmark across a newly inserted
                // OLE; the later scaffold repair then deletes that bookmark range
                // and the OLE with it. Remove the generated artifacts while their
                // OMML owner is still intact, then re-resolve the equation before
                // replacing it. External REF fields keep the same bookmark name
                // and reconnect when the target scaffold is rebuilt below.
                sourceWasNumberedManagedOmml = true;
                WordEquationNumbering.TryCaptureNativeDisplayAnchorParagraphBounds(
                    document,
                    originalMetadata.FormulaId,
                    out capturedNativeDisplayAnchorParagraphStart,
                    out _);
                WordEquationNumbering.RemoveFormulaNumberingArtifacts(
                    document,
                    originalMetadata.FormulaId);
                Release(oldRange);
                oldRange = null;
                Release(oldBookmark);
                oldBookmark = WordOmmlFormulaStore.FindByFormulaId(
                        document,
                        originalMetadata.FormulaId)
                    ?? throw new InvalidOperationException(
                        "The numbered OMML source lost its formula bookmark while detaching its generated number.");
                oldRange = ResolveOmmlEquationRange(
                    document,
                    oldBookmark,
                    session.SourceObjectId,
                    PreferSessionOmmlResolutionMetadata(
                        originalMetadata,
                        session.OriginalMetadata));
                oldStart = oldRange.Start;
            }
            if (oldShape is not null)
            {
                insertion = oldRange.Duplicate;
                insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            }
            else
            {
                // Remove the native equation before creating the OLE object.
                // Inserting at OMath.End first allows Word to expand the live
                // math container around the OLE object, leaving the large OMML
                // selection frame and shifting the replacement horizontally.
                if (oldBookmark is not null)
                    oldBookmark.Delete();
                oldRange.Delete();
                if (sourceOmmlManagedByVisualTeX)
                    WordOmmlFormulaStore.Delete(document, session.FormulaId);
                removedOmml = true;
                object insertionStart = oldStart;
                object insertionEnd = oldStart;
                insertion = document.Range(ref insertionStart, ref insertionEnd);
            }
            replacement = AddOleObject(document, insertion);
            InitializeOle(replacement, metadata, emfPath, pngPath);
            Configure(
                replacement,
                metadata,
                editedSize.Width,
                editedSize.Height,
                pngPath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            if (oldShape is not null)
            {
                RemoveInlineBaselineSentinel(
                    document,
                    originalMetadata?.FormulaId ?? metadata.FormulaId);
                RemoveInlineOleTypingAnchorAfter(oldShape);
                oldShape.Delete();
            }
            if (session.DisplayMode == "block" && session.Numbered)
                NormalizeNumberedDisplayCell(replacement);
            if (session.DisplayMode == "inline")
            {
                RestoreTypingBaselineAfter(replacement);
                // OMML -> OLE conversion must leave the insertion point in the
                // ordinary zero-position text run after the formula. Selecting
                // the replacement again in finally restores the shape's negative
                // baseline onto Word's typing caret and contaminates following
                // prose. Existing OLE edits keep the historical object selection.
                finalSelection = oldShape is null
                    ? DuplicateCurrentSelectionRange()
                    : replacement.Range.Duplicate;
            }
            else
            {
                // Numbered reconciliation can invalidate the InlineShape RCW even
                // when the host remains a tab paragraph. Bind the durable VTO_
                // identity before Word touches the paragraph, then always resolve a
                // fresh object afterwards. This is essential for OMML→OLE format
                // conversion: before this bookmark exists, a newly created native
                // OLE can be temporarily undiscoverable by embedded-metadata scans.
                BindOleIdentityBookmark(replacement, metadata.FormulaId);
                TraceNativeOleIdentityState(
                    document,
                    metadata.FormulaId,
                    "before-numbered-reconcile");
                var replacementHostMayHaveChanged =
                    NumberingOrderMayHaveChanged(originalMetadata, metadata)
                    || numberedTable is not null;
                TryReconcileShape(
                    document,
                    replacement,
                    metadata,
                    replacementHostMayHaveChanged);
                TraceNativeOleIdentityState(
                    document,
                    metadata.FormulaId,
                    "after-numbered-reconcile");
                Release(replacement);
                replacement = FindNativeOleByIdentityBookmark(
                        document,
                        metadata.FormulaId)
                    ?? FindByFormulaId(document, metadata.FormulaId)
                    ?? throw new InvalidOperationException(
                        "Word lost the replacement VisualTeX OLE formula while reconciling its numbered display paragraph.");
                finalSelection = replacement.Range.Duplicate;
            }
            if (sourceWasNumberedManagedOmml
                && capturedNativeDisplayAnchorParagraphStart >= 0)
            {
                Range? replacementRange = null;
                var removedOrphanedAnchor = false;
                try
                {
                    replacementRange = replacement.Range;
                    removedOrphanedAnchor =
                        WordEquationNumbering.RemoveCapturedNativeDisplayAnchorParagraphBeforeOle(
                            document,
                            capturedNativeDisplayAnchorParagraphStart,
                            replacementRange);
                }
                finally { Release(replacementRange); }
                if (removedOrphanedAnchor)
                {
                    Release(replacement);
                    replacement = FindNativeOleByIdentityBookmark(
                            document,
                            metadata.FormulaId)
                        ?? FindByFormulaId(document, metadata.FormulaId)
                        ?? throw new InvalidOperationException(
                            "Word lost the VisualTeX OLE after removing its obsolete numbered-OMML anchor paragraph.");
                    if (!string.Equals(
                            session.DisplayMode,
                            "inline",
                            StringComparison.Ordinal))
                    {
                        Release(finalSelection);
                        finalSelection = replacement.Range.Duplicate;
                    }
                    TraceAcceptancePerformance(
                        "ReplaceOle",
                        "remove-orphaned-omml-anchor",
                        performanceWatch,
                        ref performanceCheckpoint);
                }
            }
            BindOleIdentityBookmark(replacement, metadata.FormulaId);
            return Result(session, document);
        }
        catch
        {
            TryDelete(replacement);
            if (removedOmml
                && document is not null
                && originalMetadata is not null
                && oldStart >= 0
                && !string.IsNullOrWhiteSpace(originalOmmlWordOpenXml))
            {
                try
                {
                    rollbackEquationRange = RestoreOmmlReplacementRollback(
                        document,
                        oldStart,
                        originalOmmlWordOpenXml!);
                    if (sourceOmmlManagedByVisualTeX)
                    {
                        rollbackBookmark = WordOmmlFormulaStore.Wrap(
                            document,
                            rollbackEquationRange,
                            originalMetadata);
                        WordOmmlFormulaStore.Save(document, originalMetadata);
                        if (string.Equals(
                                originalMetadata.DisplayMode,
                                "inline",
                                StringComparison.OrdinalIgnoreCase))
                            FinalizeInlineOmmlBoundary(
                                document,
                                rollbackEquationRange,
                                originalMetadata.FormulaId,
                                moveCaretOutsideMath: false);
                        else
                            TryReconcileOmml(
                                document,
                                rollbackBookmark,
                                rollbackEquationRange,
                                originalMetadata);
                    }
                }
                catch { }
            }
            throw;
        }
        finally
        {
            RestoreViewState(document, viewState, finalSelection);
            if (screenUpdatingSuspended)
            {
                try { _application.ScreenUpdating = previousScreenUpdating; } catch { }
            }
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(finalSelection);
            Release(rollbackBookmark);
            Release(rollbackEquationRange);
            Release(numberedTable);
            Release(replacement);
            Release(insertion);
            Release(oldRange);
            Release(oldBookmark);
            Release(oldShape);
            Release(document);
        }
    }

    public OfficeObjectResult ReplaceOmml(
        OfficeSessionDocument session,
        string mathMl)
    {
        Document? document = null;
        try
        {
            var result = ReplaceOmmlCore(
                session,
                mathMl,
                out document,
                out var directTableAlreadyFinalized);
            if (document is not null
                && !directTableAlreadyFinalized
                && session.Numbered
                && string.Equals(
                    session.DisplayMode,
                    "block",
                    StringComparison.Ordinal))
            {
                // Legacy/migration hosts can still require one targeted structural
                // finalization after the replacement transaction has released its
                // OMath RCWs. A current direct-SEQ 1x3 edit is already complete in
                // ReplaceOmmlCore and deliberately skips this extra COM pass.
                var finalizedNativeNumbers =
                    WordEquationNumbering.FinalizeConvertedNumberedOmmlDisplayShapes(
                        document,
                        new[] { session.FormulaId });
                if (finalizedNativeNumbers != 1
                    && string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                {
                    Console.WriteLine(
                        $"    numbered-OMML compatibility finalization deferred formulaId={session.FormulaId}");
                }
            }
            return result;
        }
        finally { Release(document); }
    }

    private OfficeObjectResult ReplaceOmmlCore(
        OfficeSessionDocument session,
        string mathMl,
        out Document? finalizedDocument,
        out bool directTableAlreadyFinalized)
    {
        finalizedDocument = null;
        directTableAlreadyFinalized = false;
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        InlineShape? oldShape = null;
        Bookmark? oldBookmark = null;
        Range? oldRange = null;
        Range? insertion = null;
        Range? equationRange = null;
        string sourceFingerprint = string.Empty;
        Bookmark? replacement = null;
        Table? numberedTable = null;
        Paragraph? replacementParagraph = null;
        Range? replacementParagraphRange = null;
        UndoRecord? undoRecord = null;
        FormulaMetadata? originalOmmlMetadata = null;
        string? originalOmmlWordOpenXml = null;
        var originalOmmlStart = -1;
        var originalOmmlRemoved = false;
        WordViewState? viewState = null;
        Range? finalSelection = null;
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        var metadataSaved = false;
        InlineFollowingTextVisibility? inlineFollowingTextVisibility = null;
        var sourceOmmlManagedByVisualTeX = false;
        var moveCaretOutsideAfterInlineOmmlEdit = false;
        var oldNumberingArtifactsRemoved = false;
        var replaceTargetWithOmml = false;
        var replaceHealthyDirectTableAtomically = false;
        var reuseHealthyDirectTableForUnnumberingOnly = false;
        var rebuiltHealthyDirectTableViaPlaceholder = false;
        var replaceHealthyStandaloneDisplayAtomically = false;
        var reuseHealthyStandaloneDisplayForNumberingOnly = false;
        var rebuiltHealthyStandaloneDisplayViaPlaceholder = false;
        var fastHealthyStandaloneUnnumberedEdit = false;
        var replaceHealthyHashSequenceAtomically = false;
        string? preparedDirectTableOmml = null;
        float? preparedDirectTableDisplayHeightPoints = null;
        string? preparedStandaloneDisplayOmml = null;
        string? preparedHashSequenceOmml = null;
        var operationSucceeded = false;
        var performanceWatch = Stopwatch.StartNew();
        long performanceCheckpoint = 0;
        try
        {
            undoRecord = BeginUndoRecord("VisualTeX Convert or Update Word OMML Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            ApplyDocumentOmmlMathFont(document, metadata);
            // Resolve the source numbering host independently of the target's
            // Numbered flag. A numbered 1x3 formula that is being changed to
            // unnumbered still has to replace only its center OMath first, then
            // dismantle the table through the validated numbering reconciler.
            numberedTable = WordEquationNumbering.FindNumberedEquationTable(
                document,
                session.FormulaId);
            viewState = CaptureViewState();
            var keepScreenUpdatingForInlineEdit = !string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_INLINE_KEEP_SCREEN_UPDATING"),
                    "0",
                    StringComparison.Ordinal)
                && string.Equals(
                    session.DisplayMode,
                    "inline",
                    StringComparison.OrdinalIgnoreCase)
                && string.Equals(session.Mode, "edit", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(
                    session.OriginalMetadata?.NativeOmmlFingerprint);
            if (!keepScreenUpdatingForInlineEdit)
            {
                try
                {
                    previousScreenUpdating = _application.ScreenUpdating;
                    _application.ScreenUpdating = false;
                    screenUpdatingSuspended = true;
                }
                catch { }
            }
            var sourceWasOmml = !string.IsNullOrWhiteSpace(
                session.OriginalMetadata?.NativeOmmlFingerprint);
            oldShape = sourceWasOmml
                ? null
                : FindByFormulaId(
                    document,
                    session.FormulaId,
                    session.SourceObjectId);
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "locate-target",
                performanceWatch,
                ref performanceCheckpoint);
            if (oldShape is not null)
            {
                // Remove old equation-number bookmarks before inserting the
                // adjacent replacement. Word expands a trailing bookmark when
                // content is inserted at its edge; deleting that old bookmark
                // during reconciliation can otherwise delete the new table.
                WordEquationNumbering.RemoveFormulaNumberingArtifacts(
                    document,
                    session.FormulaId,
                    preserveNativeCaptionParagraph: numberedTable is not null);
                oldNumberingArtifactsRemoved = true;
                oldRange = oldShape.Range;
                insertion = oldRange.Duplicate;
                // Insert immediately before the source OLE in its existing
                // paragraph/cell. Creating another paragraph or numbered table
                // here nests layout containers during OLE -> OMML -> OLE
                // round-trips and leaves visible empty paragraph marks behind.
                insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            }
            else
            {
                oldBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    document,
                    session.FormulaId);
                if (oldBookmark is not null)
                {
                    sourceOmmlManagedByVisualTeX = true;
                    originalOmmlMetadata = WordOmmlFormulaStore.TryRead(
                        document,
                        oldBookmark)
                        ?? session.OriginalMetadata;
                    // FormulaId/bookmark is the durable OMML identity. Never replace
                    // an OMath range reconstructed from the editor-opening selection:
                    // Word may clip OMath.Range to that caret/partial probe, which
                    // leaves the old equation around the inserted replacement and
                    // corrupts the paragraph layout. Resolve the complete equation
                    // from its collapsed bookmark immediately before committing.
                    oldRange = ResolveOmmlEquationRange(
                        document,
                        oldBookmark,
                        session.SourceObjectId,
                        PreferSessionOmmlResolutionMetadata(
                            originalOmmlMetadata,
                            session.OriginalMetadata));
                }
                else
                {
                    originalOmmlMetadata = session.OriginalMetadata
                        ?? throw new InvalidOperationException(
                            "The selected Word-native OMML formula no longer exists.");
                    oldRange = ResolveStandaloneOmmlEquationRange(
                        document,
                        session.SourceObjectId,
                        originalOmmlMetadata);
                }
                originalOmmlStart = oldRange.Start;
                TraceAcceptancePerformance(
                    "ReplaceOmml",
                    "source-range-resolve",
                    performanceWatch,
                    ref performanceCheckpoint);
                originalOmmlWordOpenXml =
                    WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                        document,
                        oldRange,
                        originalOmmlMetadata!.FormulaId);
                TraceAcceptancePerformance(
                    "ReplaceOmml",
                    "rollback-capture",
                    performanceWatch,
                    ref performanceCheckpoint);
                insertion = oldRange.Duplicate;
            }
            moveCaretOutsideAfterInlineOmmlEdit = oldShape is null
                && session.DisplayMode == "inline"
                && string.Equals(session.Mode, "edit", StringComparison.OrdinalIgnoreCase);
            if (oldShape is not null)
            {
                var sourceOleMetadata = WordFormulaMetadataReader.TryRead(oldShape)
                    ?? session.OriginalMetadata;
                StoreWordInlineOleSize(
                    metadata,
                    oldShape.Width,
                    oldShape.Height,
                    session.DisplayMode == "inline"
                        && string.Equals(
                            sourceOleMetadata?.DisplayMode,
                            "inline",
                            StringComparison.OrdinalIgnoreCase));
                metadata.Validate();
            }
            else if (session.DisplayMode != "inline")
            {
                StoreWordInlineOleSize(metadata, 0, 0, inline: false);
                metadata.Validate();
            }
            if (session.DisplayMode == "inline")
            {
                inlineFollowingTextVisibility =
                    CaptureInlineFollowingTextVisibility(oldRange);
                PrepareInlineBaselineSentinelAfterFormula(
                    document,
                    oldRange,
                    metadata.FormulaId);
            }
            if (oldShape is null)
            {
                var blockTarget = string.Equals(
                    session.DisplayMode,
                    "block",
                    StringComparison.Ordinal);
                var numberedBlock = session.Numbered && blockTarget;

                // Detect the *source* 1x3 host before branching on the target's
                // Numbered flag. This is required for numbered→unnumbered edits:
                // first replace only the center OMath, then let reconciliation
                // dismantle the still-intact table. Sending that case through the
                // generic placeholder path lets Word absorb a cell/paragraph mark.
                replaceHealthyDirectTableAtomically = blockTarget
                    && numberedTable is not null
                    && ((originalOmmlMetadata?.Numbered == true
                            && (session.Numbered
                                || CanReuseNumberedOmmlForUnnumberingOnly(
                                    originalOmmlMetadata,
                                    metadata)))
                            ? WordEquationNumbering
                                .HasReusableNumberedNativeOmmlDirectTableCenterHost(
                                    numberedTable,
                                    oldRange)
                            : WordEquationNumbering
                                .HasReusableNumberedNativeOmmlDirectTableHost(
                                    document,
                                    oldRange,
                                    metadata.FormulaId));
                TraceAcceptancePerformance(
                    "ReplaceOmml",
                    "direct-host-health",
                    performanceWatch,
                    ref performanceCheckpoint);
                if (replaceHealthyDirectTableAtomically)
                {
                    reuseHealthyDirectTableForUnnumberingOnly =
                        !session.Numbered
                        && CanReuseNumberedOmmlForUnnumberingOnly(
                            originalOmmlMetadata,
                            metadata);
                    if (reuseHealthyDirectTableForUnnumberingOnly)
                    {
                        // Pure Numbered=true -> false toggle: the center OMath is
                        // already the exact desired formula. Keep it live and let the
                        // direct-table dismantler preserve that same OMath through
                        // ConvertToText. A non-empty marker selects the trusted live-
                        // OMath removal path; it is never parsed or inserted.
                        sourceFingerprint = originalOmmlMetadata!.NativeOmmlFingerprint!;
                        preparedDirectTableOmml = originalOmmlWordOpenXml
                            ?? "<visualtex-live-omml/>";
                    }
                    else
                    {
                        // Current numbered OMML is an exact 1x3 host and its content
                        // actually changed. Preserve the table and the right-cell
                        // direct SEQ/bookmarks, but never overwrite the live center
                        // OMath through FormattedText. That later conversion-oriented
                        // path can materialize an empty OMath in ordinary installed
                        // Word documents. Replace only the exact center OMath with one
                        // ordinary character, then import the new professional OMML at
                        // that same range through WordOmmlConverter.Insert.
                        var semanticOmml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
                        semanticOmml = ApplyOmmlTypographyXml(
                            semanticOmml,
                            session.FontSizePt,
                            metadata);
                        sourceFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                            semanticOmml);
                        preparedDirectTableOmml = semanticOmml;
                        TraceAcceptancePerformance(
                            "ReplaceOmml",
                            "direct-content-prep",
                            performanceWatch,
                            ref performanceCheckpoint);
                        if (oldBookmark is not null)
                        {
                            oldBookmark.Delete();
                            Release(oldBookmark);
                            oldBookmark = null;
                        }
                        insertion!.Text = BulkInlineFormulaPlaceholder;
                        insertion.SetRange(
                            originalOmmlStart,
                            originalOmmlStart + BulkInlineFormulaPlaceholder.Length);
                        originalOmmlRemoved = true;
                        rebuiltHealthyDirectTableViaPlaceholder = true;
                    }
                    if (reuseHealthyDirectTableForUnnumberingOnly)
                    {
                        Release(insertion);
                        insertion = oldRange.Duplicate;
                    }
                }
                else if (blockTarget
                    && (session.Numbered || originalOmmlMetadata?.Numbered != false
                        ? HasReusableStandaloneDisplayOmmlHost(document, oldRange)
                        : HasReusableStandaloneDisplayOmmlEditHost(oldRange)))
                {
                    fastHealthyStandaloneUnnumberedEdit = !session.Numbered
                        && originalOmmlMetadata?.Numbered == false
                        && numberedTable is null;
                    if (string.Equals(
                            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                            "1",
                            StringComparison.Ordinal))
                        Console.WriteLine("    [perf] ReplaceOmml.source-path=standalone-display");
                    reuseHealthyStandaloneDisplayForNumberingOnly =
                        CanReuseStandaloneOmmlForNumberingOnly(
                            originalOmmlMetadata,
                            metadata,
                            mathMl,
                            originalOmmlWordOpenXml);
                    if (reuseHealthyStandaloneDisplayForNumberingOnly)
                    {
                        replaceHealthyStandaloneDisplayAtomically = true;
                        if (string.Equals(
                                Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                                "1",
                                StringComparison.Ordinal))
                            Console.WriteLine("    [perf] ReplaceOmml.source-path=standalone-numbering-only-reuse");
                        // A true numbering-only edit does not need to rebuild the
                        // existing professional OMath. Keep it live and let the
                        // direct-SEQ reconciler move that exact formula into cell
                        // (1,2). This is a performance fast path only; correctness
                        // must not depend on the incoming LaTeX staying byte-identical.
                        sourceFingerprint = originalOmmlMetadata!.NativeOmmlFingerprint!;
                        Release(insertion);
                        insertion = oldRange.Duplicate;
                    }
                    else
                    {
                        if (string.Equals(
                                Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                                "1",
                                StringComparison.Ordinal))
                            Console.WriteLine("    [perf] ReplaceOmml.source-path=standalone-placeholder-rebuild");
                        // Do not replace a live standalone OMath through FormattedText.
                        // That path was introduced later for conversion/layout repair
                        // and is not reliable in real, long-lived Word documents: a
                        // valid MathML payload can materialize as an empty OMath. The
                        // mature Word edit path removes the complete resolved equation,
                        // leaves one ordinary placeholder, then asks WordOmmlConverter
                        // to insert the new OMML at that exact text range. Numbered
                        // targets are reconciled into the current 1x3 direct-SEQ host
                        // afterwards, so this does not change the numbering architecture.
                        if (oldBookmark is not null)
                        {
                            oldBookmark.Delete();
                            Release(oldBookmark);
                            oldBookmark = null;
                        }
                        insertion!.Text = BulkInlineFormulaPlaceholder;
                        insertion.SetRange(
                            originalOmmlStart,
                            originalOmmlStart + BulkInlineFormulaPlaceholder.Length);
                        originalOmmlRemoved = true;
                        rebuiltHealthyStandaloneDisplayViaPlaceholder = true;
                    }
                }
                else if (numberedBlock)
                {
                    replaceHealthyHashSequenceAtomically =
                        WordEquationNumbering
                            .HasReusableNumberedNativeOmmlHashSequenceHost(
                                document,
                                oldRange,
                                metadata.FormulaId);
                    if (replaceHealthyHashSequenceAtomically)
                    {
                        // A retired but still valid numbered formula can own one
                        // mathematical #(SEQ) host. Replace that complete OMath
                        // atomically, then let reconciliation migrate it to 1x3.
                        var semanticOmml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
                        semanticOmml = ApplyOmmlTypographyXml(
                            semanticOmml,
                            session.FontSizePt,
                            metadata);
                        sourceFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                            semanticOmml);
                        preparedHashSequenceOmml =
                            WordEquationNumbering
                                .PrepareNumberedNativeOmmlHashSequenceForReplacement(
                                    document,
                                    semanticOmml,
                                    metadata.FormulaId,
                                    oldRange.Start);

                        WordEquationNumbering
                            .RemoveNativeOmmlHashSequenceAliasesForReplacement(
                                document,
                                metadata.FormulaId);
                        if (oldBookmark is not null)
                        {
                            oldBookmark.Delete();
                            Release(oldBookmark);
                            oldBookmark = null;
                        }
                        Release(insertion);
                        insertion = oldRange.Duplicate;
                    }
                    else
                    {
                        // Only legacy paragraph/tab, Shape and malformed eqArr
                        // hosts use the compatibility staging placeholder.
                        WordEquationNumbering.RemoveVisibleEquationNumberForFormula(
                            document,
                            metadata.FormulaId);
                        if (oldBookmark is not null) oldBookmark.Delete();
                        var tabPlaceholder = PrepareNumberedOmmlReplacementTabPlaceholder(
                            document,
                            oldRange);
                        Release(insertion);
                        insertion = tabPlaceholder;
                        originalOmmlRemoved = true;
                        replaceTargetWithOmml = true;
                    }
                }
                else
                {
                    // Never ask Word to overwrite a live standalone OMath directly.
                    // Replace the complete resolved equation with one ordinary
                    // placeholder, then replace that exact character with OMML.
                    if (oldBookmark is not null)
                    {
                        oldBookmark.Delete();
                        Release(oldBookmark);
                        oldBookmark = null;
                    }
                    insertion!.Text = BulkInlineFormulaPlaceholder;
                    insertion.SetRange(
                        originalOmmlStart,
                        originalOmmlStart + BulkInlineFormulaPlaceholder.Length);
                    originalOmmlRemoved = true;
                }
            }
            else if (session.Numbered
                && string.Equals(
                    session.DisplayMode,
                    "block",
                    StringComparison.Ordinal))
            {
                // OLE -> numbered OMML must obey the same host-construction order
                // as fresh insertion and OMML editing. Both ordinary layout TABs
                // must exist before OMath is materialized; adding either boundary
                // afterwards lets Word absorb TAB/REF content into OMath.Range.
                // Keep the source OLE alive until the replacement is valid, but
                // normalize all structural paragraph text around it first and
                // pre-seed TAB + placeholder + TAB immediately before the OLE.
                var tabPlaceholder = PrepareNumberedOmmlReplacementTabPlaceholderPreservingOle(
                    document,
                    oldShape!);
                Release(insertion);
                insertion = tabPlaceholder;
                replaceTargetWithOmml = true;
            }
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "resolve-source",
                performanceWatch,
                ref performanceCheckpoint);

            if (replaceHealthyDirectTableAtomically)
            {
                if (reuseHealthyDirectTableForUnnumberingOnly)
                {
                    equationRange = insertion!.Duplicate;
                    originalOmmlRemoved = true;
                }
                else
                {
                    if (!rebuiltHealthyDirectTableViaPlaceholder
                        || preparedDirectTableOmml is null)
                        throw new InvalidOperationException(
                            "The direct-table OMML replacement payload was not prepared.");
                    equationRange = WordOmmlConverter.Insert(
                        _application,
                        document,
                        insertion!,
                        mathMl,
                        display: true,
                        sourceFingerprint: out sourceFingerprint,
                        replaceTarget: true,
                        transformOmml: omml => ApplyOmmlTypographyXml(
                            omml,
                            session.FontSizePt,
                            metadata),
                        mathFontName: document.OMathFontName);
                    // The source center OMath was replaced by a one-character text
                    // range before insertion. Any later failure restores its captured
                    // WordOpenXML into the same center cell; the pre-existing table
                    // and right-cell direct-SEQ identities are never deleted.
                    originalOmmlRemoved = true;
                }
            }
            else if (replaceHealthyStandaloneDisplayAtomically)
            {
                if (reuseHealthyStandaloneDisplayForNumberingOnly)
                {
                    equationRange = insertion!.Duplicate;
                    // From this point the numbering reconciler is allowed to move
                    // and delete the standalone source paragraph. Mark the original
                    // as rollback-owned so any later structural failure restores the
                    // captured pre-edit WordOpenXML atomically.
                    originalOmmlRemoved = true;
                }
                else
                {
                    if (preparedStandaloneDisplayOmml is null)
                        throw new InvalidOperationException(
                            "The standalone-display OMML replacement payload was not prepared.");
                    equationRange = WordOmmlConverter.ReplaceWithPreparedOmml(
                        _application,
                        document,
                        insertion,
                        preparedStandaloneDisplayOmml,
                        display: true,
                        mathFontName: document.OMathFontName);
                    originalOmmlRemoved = true;
                }
            }
            else if (replaceHealthyHashSequenceAtomically)
            {
                if (preparedHashSequenceOmml is null)
                    throw new InvalidOperationException(
                        "The atomic numbered-OMML replacement payload was not prepared.");
                originalOmmlRemoved = true;
                equationRange = WordOmmlConverter.ReplaceWithPreparedOmml(
                    _application,
                    document,
                    insertion,
                    preparedHashSequenceOmml,
                    display: true,
                    mathFontName: document.OMathFontName);
            }
            else
            {
                equationRange = WordOmmlConverter.Insert(
                    _application,
                    document,
                    insertion,
                    mathMl,
                    session.DisplayMode == "block" && !session.Numbered,
                    sourceFingerprint: out sourceFingerprint,
                    replaceTarget: oldShape is null || replaceTargetWithOmml,
                    transformOmml: omml => ApplyOmmlTypographyXml(
                        omml,
                        session.FontSizePt,
                        metadata),
                    mathFontName: document.OMathFontName);
            }
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "insert-native-omml",
                performanceWatch,
                ref performanceCheckpoint);
            ValidateInsertedOmml(equationRange);
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "validate-native-omml",
                performanceWatch,
                ref performanceCheckpoint);
            if (!reuseHealthyStandaloneDisplayForNumberingOnly
                && !reuseHealthyDirectTableForUnnumberingOnly
                && string.Equals(
                    session.DisplayMode,
                    "inline",
                    StringComparison.OrdinalIgnoreCase))
            {
                // Every replacement payload has already passed through
                // ApplyOmmlTypographyXml before Word materializes the OMath, and
                // ApplyDocumentOmmlMathFont ran at the start of this transaction.
                // Reassigning Range.Font.Size/Position here performs a second full
                // COM walk of the same native tree (~100ms in a 100-OMML document)
                // without changing its requested typography. Inline formulas still
                // need the independent line-grid safeguard for stacked fractions.
                StabilizeInlineOmmlFractionLineGrid(equationRange);
                TraceAcceptancePerformance(
                    "ReplaceOmml",
                    "inline-line-grid",
                    performanceWatch,
                    ref performanceCheckpoint);
            }
            if (replaceHealthyDirectTableAtomically
                && session.Numbered
                && numberedTable is not null)
            {
                preparedDirectTableDisplayHeightPoints =
                    WordOmmlFormulaStore.EstimateHeightPoints(equationRange);
                WordEquationNumbering.ApplyNativeOmmlTableMinimumDisplayHeight(
                    numberedTable,
                    preparedDirectTableDisplayHeightPoints.Value);
            }
            metadata.NativeOmmlFingerprint = sourceFingerprint;
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "stamp-fingerprint",
                performanceWatch,
                ref performanceCheckpoint);
            var useFreshOmmlAnchor = sourceOmmlManagedByVisualTeX
                && originalOmmlRemoved
                && oldBookmark is null
                && (fastHealthyStandaloneUnnumberedEdit
                    || string.Equals(
                        session.DisplayMode,
                        "inline",
                        StringComparison.OrdinalIgnoreCase));
            replacement = useFreshOmmlAnchor
                ? WordOmmlFormulaStore.WrapFreshOmmlReplacement(
                    document,
                    equationRange,
                    metadata)
                : WordOmmlFormulaStore.Wrap(
                    document,
                    equationRange,
                    metadata,
                    replaceExisting: oldBookmark is not null);
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "wrap-formula-anchor",
                performanceWatch,
                ref performanceCheckpoint);
            if (replaceHealthyDirectTableAtomically
                && string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                && string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_VSTO_OMML_FAIL_STAGE"),
                    "after-direct-table-replacement",
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "Injected failure after direct-table OMML replacement.");

            // An OLE -> OMML conversion has no pre-existing CustomXML fallback,
            // so persist provisional metadata before deleting the source OLE. A
            // normal OMML edit keeps its old durable metadata until the final
            // Word-normalized fingerprint is ready below.
            if (oldShape is not null)
            {
                WordOmmlFormulaStore.Save(document, metadata);
                metadataSaved = true;
            }
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "protect-source-metadata",
                performanceWatch,
                ref performanceCheckpoint);

            // Keep the source OLE until replacement and metadata are valid.
            if (oldShape is not null)
            {
                RemoveInlineBaselineSentinel(
                    document,
                    originalOmmlMetadata?.FormulaId ?? metadata.FormulaId);
                RemoveInlineOleTypingAnchorAfter(oldShape);
                oldShape.Delete();
            }
            if (session.DisplayMode == "block"
                && session.Numbered
                && !replaceHealthyDirectTableAtomically
                && !replaceHealthyStandaloneDisplayAtomically
                && !replaceHealthyHashSequenceAtomically)
            {
                // Legacy hosts are staged as an inline OMath before the numbering
                // reconciler atomically materializes the final genuine #(SEQ)
                // display paragraph. A healthy current host was already replaced as
                // one complete numbered OMath and needs no temporary anchor here.
                Release(replacement);
                replacement = WordOmmlFormulaStore.Wrap(
                    document,
                    equationRange,
                    metadata);
            }
            // The replacement range and metadata are already available here.
            // Re-reading both through the new bookmark only to set a temporary
            // caret adds a large COM round-trip and is immediately overwritten
            // by RestoreViewState's final formula selection. Keep the durable
            // inline text boundary normalized directly from the live range.
            if (session.DisplayMode == "inline")
            {
                var finalizedInlineRange = FinalizeInlineOmmlBoundaryFromLiveRange(
                    document,
                    equationRange,
                    metadata.FormulaId,
                    inlineFollowingTextVisibility);
                Release(equationRange);
                equationRange = finalizedInlineRange;
            }
            if (session.DisplayMode == "block"
                && !(replaceHealthyDirectTableAtomically && session.Numbered))
            {
                if (fastHealthyStandaloneUnnumberedEdit)
                {
                    // A validated standalone Display OMath whose numbering state
                    // remains false has no numbering host to build or dismantle.
                    // Preserve only the display paragraph normalization normally
                    // performed at the start of TryReconcileOmml; entering the
                    // numbering reconciler here costs ~100ms and cannot change the
                    // structure of this local content edit.
                    ResetDisplayFormulaPosition(equationRange);
                }
                else
                {
                    TryReconcileOmml(
                        document,
                        replacement!,
                        equationRange,
                        metadata,
                        NumberingOrderMayHaveChanged(session.OriginalMetadata, metadata),
                        reuseExistingNumberedTableFormatting: numberedTable is not null,
                        knownNumberedTable: numberedTable,
                        deferNativeOmmlShapeFinalization:
                            session.Numbered
                            && string.Equals(
                                session.DisplayMode,
                                "block",
                                StringComparison.Ordinal),
                        deferNativeOmmlMetadataPersistence: true,
                        preparedUnnumberedOmml:
                            replaceHealthyDirectTableAtomically && !session.Numbered
                                ? preparedDirectTableOmml
                                : null);
                }
            }
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "reconcile",
                performanceWatch,
                ref performanceCheckpoint);
            if (session.Numbered
                && string.Equals(
                    session.DisplayMode,
                    "block",
                    StringComparison.Ordinal)
                && !replaceHealthyDirectTableAtomically)
            {
                // Creating a direct-SEQ 1x3 host deliberately keeps the original
                // source paragraph until the center OMath, right-cell SEQ and all
                // FormulaId aliases are durable. This applies both to OLE→OMML and
                // to the common standalone OMML unnumbered→numbered edit. Resolve
                // the table directly from the live equation Range that reconciliation
                // just moved into cell (1,2); do not rediscover it from document-wide
                // numbering identities during spacing cleanup.
                if (numberedTable is null)
                {
                    Tables? liveTables = null;
                    try
                    {
                        if ((bool)equationRange.get_Information(
                                WdInformation.wdWithInTable))
                        {
                            liveTables = equationRange.Tables;
                            if (liveTables.Count == 1)
                                numberedTable = liveTables[1];
                        }
                    }
                    catch { }
                    finally { Release(liveTables); }
                }
                WordEquationNumbering.CleanupNumberedDisplayInsertionSpacing(
                    document,
                    metadata.FormulaId,
                    numberedTable);
                TraceAcceptancePerformance(
                    "ReplaceOmml",
                    "cleanup-numbered-source-spacing",
                    performanceWatch,
                    ref performanceCheckpoint);
            }

            // Numbered tab/field reconciliation can shift and recreate ordinary
            // paragraph runs around the OMath. The pre-reconcile Range RCW may then
            // point at a former field/text span even though the native equation is
            // healthy. Resolve the durable FormulaId bookmark again before reading
            // final OpenXML or selecting the replacement.
            var finalizeDirectUnnumberFromLiveRange =
                replaceHealthyDirectTableAtomically
                && !session.Numbered
                && string.Equals(
                    session.DisplayMode,
                    "block",
                    StringComparison.Ordinal);
            Range? liveUnnumberedRange = null;
            if (finalizeDirectUnnumberFromLiveRange)
            {
                liveUnnumberedRange = TryResolveStandaloneOmmlFromLiveRange(equationRange)
                    ?? throw new InvalidOperationException(
                        "Word lost the preserved standalone Display OMath after removing its direct-SEQ table.");
            }

            Release(replacement);
            replacement = null;
            var keepFreshInlineRange = string.Equals(
                    metadata.DisplayMode,
                    "inline",
                    StringComparison.Ordinal)
                && equationRange is not null;
            var keepFreshStandaloneDisplayRange =
                fastHealthyStandaloneUnnumberedEdit
                && equationRange is not null;
            if (finalizeDirectUnnumberFromLiveRange)
            {
                Release(equationRange);
                equationRange = liveUnnumberedRange;
                liveUnnumberedRange = null;
                // The managed VTOMML anchor was deliberately removed before
                // ConvertToText. Rebind it once, after the standalone Display OMath
                // is final, without the generic table-affinity probe that is only
                // needed for an equation that might still live inside a 1x3 host.
                replacement = WordOmmlFormulaStore.WrapFreshOmmlReplacement(
                    document,
                    equationRange,
                    metadata);
            }
            else
            {
                replacement = WordOmmlFormulaStore.FindByFormulaId(
                        document,
                        metadata.FormulaId)
                    ?? throw new InvalidOperationException(
                        "Word lost the replacement OMML formula bookmark while finalizing its layout.");
                if (!keepFreshInlineRange && !keepFreshStandaloneDisplayRange)
                {
                    Release(equationRange);
                    if (metadata.Numbered
                        && string.Equals(
                            metadata.DisplayMode,
                            "block",
                            StringComparison.Ordinal))
                    {
                        equationRange = TryResolveNumberedOmmlFromKnownTable(numberedTable)
                            ?? ResolveNumberedOmmlFromNumberingOwner(
                                document,
                                metadata.FormulaId);
                    }
                    else
                    {
                        equationRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                            document,
                            metadata.FormulaId,
                            metadata);
                    }
                }
            }
            Release(liveUnnumberedRange);
            if (!WordOmmlFormulaStore.IsCanonicalAnchor(replacement, equationRange))
            {
                Release(replacement);
                replacement = WordOmmlFormulaStore.Wrap(
                    document,
                    equationRange,
                    metadata,
                    replaceExisting: true);
            }
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "final-identity-resolve",
                performanceWatch,
                ref performanceCheckpoint);

            // Save identity only after Word has finished its final OMath/layout
            // normalization. Otherwise the stored fingerprint can already be
            // stale before the editor is opened again, and a later bookmark
            // drift becomes unrecoverable.
            if ((rebuiltHealthyDirectTableViaPlaceholder
                    || replaceHealthyStandaloneDisplayAtomically
                    || rebuiltHealthyStandaloneDisplayViaPlaceholder)
                && !string.IsNullOrWhiteSpace(sourceFingerprint))
            {
                // Both healthy standalone paths materialize one complete semantic
                // OMath and then either keep it live or move that same OMath into the
                // direct-SEQ center cell. The Insert path already computed the
                // fingerprint from the exact transformed OMML that Word imported.
                // Re-reading Range.WordOpenXML here is document-sized work and costs
                // ~300ms in a 100-OMML document, undoing the earlier Word performance
                // work for every ordinary content edit.
                metadata.NativeOmmlFingerprint = sourceFingerprint;
            }
            else
            {
                WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                    metadata,
                    equationRange);
            }
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "final-identity-fingerprint",
                performanceWatch,
                ref performanceCheckpoint);
            if (!WordOmmlFormulaStore.IsCanonicalAnchor(replacement, equationRange))
            {
                Release(replacement);
                replacement = WordOmmlFormulaStore.Wrap(
                    document,
                    equationRange,
                    metadata,
                    replaceExisting: true);
            }
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "final-identity-anchor",
                performanceWatch,
                ref performanceCheckpoint);
            if (!WordOmmlFormulaStore.TrySaveKnownCachedPart(document, metadata))
                WordOmmlFormulaStore.Save(document, metadata);
            metadataSaved = true;
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "final-identity-metadata",
                performanceWatch,
                ref performanceCheckpoint);
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "finalize-identity",
                performanceWatch,
                ref performanceCheckpoint);
            finalSelection = equationRange.Duplicate;
            var result = Result(session, document);
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "prepare-result",
                performanceWatch,
                ref performanceCheckpoint);
            directTableAlreadyFinalized =
                (replaceHealthyDirectTableAtomically
                    || replaceHealthyStandaloneDisplayAtomically)
                && session.Numbered
                && string.Equals(
                    session.DisplayMode,
                    "block",
                    StringComparison.Ordinal);
            operationSucceeded = true;
            return result;
        }
        catch
        {
            if (replaceHealthyDirectTableAtomically)
            {
                // This table existed before the edit and may contain every durable
                // number/reference identity. Remove only the provisional formula
                // bookmark and replacement OMath so the captured original OMath can
                // be restored into the same center cell below.
                TryDelete(replacement, deleteContents: false);
                TryDelete(equationRange);
            }
            else
            {
                TryDelete(replacement, deleteContents: true);
                if (replacement is null) TryDelete(equationRange);
            }
            // numberedTable is always a source host resolved before mutation. It is
            // never temporary cleanup state and must not be deleted on failure.
            if (oldShape is not null && numberedTable is null)
                TryDelete(replacementParagraphRange);
            if (document is not null)
            {
                try
                {
                    if (originalOmmlRemoved
                        && originalOmmlStart >= 0
                        && !string.IsNullOrWhiteSpace(originalOmmlWordOpenXml)
                        && originalOmmlMetadata is not null)
                    {
                        Range? restoredRange = null;
                        Bookmark? restoredBookmark = null;
                        try
                        {
                            restoredRange = RestoreOmmlReplacementRollback(
                                document,
                                originalOmmlStart,
                                originalOmmlWordOpenXml!);
                            if (sourceOmmlManagedByVisualTeX)
                            {
                                restoredBookmark = WordOmmlFormulaStore.Wrap(
                                    document,
                                    restoredRange,
                                    originalOmmlMetadata);
                                WordOmmlFormulaStore.Save(document, originalOmmlMetadata);
                                if (originalOmmlMetadata.DisplayMode == "inline")
                                    FinalizeInlineOmmlBoundary(
                                        document,
                                        restoredRange,
                                        originalOmmlMetadata.FormulaId,
                                        moveCaretOutsideMath: false);
                                else
                                {
                                    var restoredHealthyDirectTable =
                                        originalOmmlMetadata.Numbered
                                        && string.Equals(
                                            originalOmmlMetadata.DisplayMode,
                                            "block",
                                            StringComparison.Ordinal)
                                        && WordEquationNumbering
                                            .HasReusableNumberedNativeOmmlDirectTableHost(
                                                document,
                                                restoredRange,
                                                originalOmmlMetadata.FormulaId);
                                    if (!restoredHealthyDirectTable)
                                        TryReconcileOmml(
                                            document,
                                            restoredBookmark,
                                            restoredRange,
                                            originalOmmlMetadata);
                                }
                            }
                            else if (metadataSaved)
                            {
                                WordOmmlFormulaStore.Delete(document, metadata.FormulaId);
                            }
                        }
                        finally
                        {
                            Release(restoredBookmark);
                            Release(restoredRange);
                        }
                    }
                    else if (metadataSaved)
                    {
                        WordOmmlFormulaStore.Delete(document, metadata.FormulaId);
                    }
                    if (oldNumberingArtifactsRemoved && oldShape is not null)
                        WordEquationNumbering.TryReconcile(document);
                }
                catch { }
            }
            throw;
        }
        finally
        {
            RestoreViewState(document, viewState, finalSelection);
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "restore-view",
                performanceWatch,
                ref performanceCheckpoint);
            if (document is not null
                && equationRange is not null
                && moveCaretOutsideAfterInlineOmmlEdit)
            {
                Range? finalInlineRange = null;
                try
                {
                    finalInlineRange = TryResolveInlineOmmlFromLiveRange(equationRange)
                        ?? ResolveCurrentInlineOmmlRange(
                            document,
                            equationRange,
                            metadata.FormulaId);
                    // Word 2021 keeps a collapsed caret at OMath.End on the math
                    // side of the boundary. Its native right-arrow operation is
                    // the reliable way to switch that same position to ordinary
                    // text affinity after RestoreViewState selected the formula.
                    MoveCaretOutsideInlineOmml(finalInlineRange);
                }
                catch
                {
                    // The replacement is already committed. A stale final range
                    // must not roll back a valid equation.
                }
                finally { Release(finalInlineRange); }
            }
            if (screenUpdatingSuspended)
            {
                try { _application.ScreenUpdating = previousScreenUpdating; } catch { }
            }
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "restore-screen-updating",
                performanceWatch,
                ref performanceCheckpoint);
            EndUndoRecord(undoRecord);
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "end-undo",
                performanceWatch,
                ref performanceCheckpoint);
            Release(undoRecord);
            Release(finalSelection);
            Release(replacement);
            Release(equationRange);
            Release(replacementParagraphRange);
            Release(replacementParagraph);
            Release(numberedTable);
            Release(insertion);
            Release(oldRange);
            Release(oldBookmark);
            Release(oldShape);
            if (operationSucceeded && document is not null)
            {
                finalizedDocument = document;
                document = null;
            }
            Release(document);
            TraceAcceptancePerformance(
                "ReplaceOmml",
                "release-rcws",
                performanceWatch,
                ref performanceCheckpoint);
        }
    }

    public OfficeObjectResult Replace(OfficeSessionDocument session, string imagePath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        InlineShape? oldShape = null;
        Range? oldRange = null;
        Range? insertion = null;
        InlineShape? replacement = null;
        UndoRecord? undoRecord = null;
        WordViewState? viewState = null;
        Range? finalSelection = null;
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        try
        {
            undoRecord = BeginUndoRecord("VisualTeX Replace Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            viewState = CaptureViewState();
            try
            {
                previousScreenUpdating = _application.ScreenUpdating;
                _application.ScreenUpdating = false;
                screenUpdatingSuspended = true;
            }
            catch { }
            oldShape = FindByFormulaId(
                    document,
                    session.FormulaId,
                    session.SourceObjectId)
                ?? throw new InvalidOperationException("The target Word formula no longer exists.");
            var oldWidth = oldShape.Width;
            var oldHeight = oldShape.Height;
            var originalMetadata = WordFormulaMetadataReader.TryRead(oldShape)
                ?? session.OriginalMetadata;
            var editedSize = OfficeFormulaSizing.EditedSize(
                oldWidth,
                oldHeight,
                originalMetadata?.RenderWidthPx,
                originalMetadata?.RenderHeightPx,
                session.ExportResult?.Width ?? oldWidth / 0.75f,
                session.ExportResult?.Height ?? oldHeight / 0.75f,
                originalFontSizePt: originalMetadata?.FontSizePt,
                originalRenderFontSizePt: originalMetadata?.RenderFontSizePt);
            oldRange = oldShape.Range;
            insertion = oldRange.Duplicate;
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            object link = false;
            object save = true;
            object rangeObject = insertion;
            replacement = document.InlineShapes.AddPicture(
                imagePath,
                ref link,
                ref save,
                ref rangeObject);
            Configure(
                replacement,
                metadata,
                editedSize.Width,
                editedSize.Height,
                imagePath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            RemoveInlineBaselineSentinel(
                document,
                originalMetadata?.FormulaId ?? metadata.FormulaId);
            RemoveInlineOleTypingAnchorAfter(oldShape);
            oldShape.Delete();
            if (session.DisplayMode == "inline")
                RestoreTypingBaselineAfter(replacement);
            else
                TryReconcileShape(document, replacement, metadata);
            finalSelection = replacement.Range.Duplicate;
            return Result(session, document);
        }
        catch
        {
            TryDelete(replacement);
            throw;
        }
        finally
        {
            RestoreViewState(document, viewState, finalSelection);
            if (screenUpdatingSuspended)
            {
                try { _application.ScreenUpdating = previousScreenUpdating; } catch { }
            }
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(finalSelection);
            Release(replacement);
            Release(insertion);
            Release(oldRange);
            Release(oldShape);
            Release(document);
        }
    }

    private static bool CanReuseStandaloneOmmlForNumberingOnly(
        FormulaMetadata? original,
        FormulaMetadata current,
        string incomingMathMl,
        string? originalWordOpenXml)
    {
        if (original is null
            || original.Numbered
            || !current.Numbered
            || !string.Equals(
                original.DisplayMode,
                "block",
                StringComparison.Ordinal)
            || !string.Equals(
                current.DisplayMode,
                "block",
                StringComparison.Ordinal)
            || string.IsNullOrWhiteSpace(original.NativeOmmlFingerprint))
            return false;

        var originalSize = FormulaFontSize.ResolveSemanticFontSize(original);
        var currentSize = FormulaFontSize.ResolveSemanticFontSize(current);
        if (Math.Abs(originalSize - currentSize) >= 0.001f
            || FormulaFontPreferencesChanged(original, current))
            return false;

        // Legacy Office metadata persisted raw LaTeX as codeFormat="latex",
        // while the current editor internally normalizes the same source mode to
        // "raw". A numbering-only edit must never rebuild a healthy live OMath
        // merely because that representation label changed. Keep the cheap exact
        // metadata path first, then fall back to semantic MathML equivalence for
        // editor normalizations such as e^{iπ} -> \\mathrm{e}^{\\mathrm{i}π}.
        if (string.Equals(original.Latex, current.Latex, StringComparison.Ordinal)
            && CodeFormatsRepresentSameRawLatex(original.CodeFormat, current.CodeFormat))
            return true;

        if (string.IsNullOrWhiteSpace(incomingMathMl)
            || string.IsNullOrWhiteSpace(originalWordOpenXml))
            return false;
        try
        {
            var originalMathMl = WordOmmlConverter.TransformOmmlToMathMl(
                originalWordOpenXml!,
                display: true);
            return string.Equals(
                MathTypeMtefCodec.SemanticSignature(originalMathMl),
                MathTypeMtefCodec.SemanticSignature(incomingMathMl),
                StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
    }

    private static bool CodeFormatsRepresentSameRawLatex(string? left, string? right)
    {
        static string Normalize(string? value) =>
            string.Equals(value, "latex", StringComparison.OrdinalIgnoreCase)
                ? "raw"
                : (value ?? string.Empty).Trim();
        return string.Equals(
            Normalize(left),
            Normalize(right),
            StringComparison.OrdinalIgnoreCase);
    }

    private static bool CanReuseNumberedOmmlForUnnumberingOnly(
        FormulaMetadata? original,
        FormulaMetadata current)
    {
        if (original is null
            || !original.Numbered
            || current.Numbered
            || !string.Equals(
                original.DisplayMode,
                "block",
                StringComparison.Ordinal)
            || !string.Equals(
                current.DisplayMode,
                "block",
                StringComparison.Ordinal)
            || string.IsNullOrWhiteSpace(original.NativeOmmlFingerprint)
            || !string.Equals(original.Latex, current.Latex, StringComparison.Ordinal)
            || !string.Equals(
                original.CodeFormat,
                current.CodeFormat,
                StringComparison.Ordinal))
            return false;

        var originalSize = FormulaFontSize.ResolveSemanticFontSize(original);
        var currentSize = FormulaFontSize.ResolveSemanticFontSize(current);
        return Math.Abs(originalSize - currentSize) < 0.001f
            && !FormulaFontPreferencesChanged(original, current);
    }

    private static bool FormulaFontPreferencesChanged(
        FormulaMetadata? original,
        FormulaMetadata current)
    {
        static string Letter(string? value) =>
            string.IsNullOrWhiteSpace(value) ? "katex" : value!.Trim();
        static string Chinese(string? value) =>
            string.IsNullOrWhiteSpace(value) ? "system" : value!.Trim();
        return !string.Equals(
                Letter(original?.FormulaLetterFont),
                Letter(current.FormulaLetterFont),
                StringComparison.OrdinalIgnoreCase)
            || !string.Equals(
                Chinese(original?.FormulaChineseFont),
                Chinese(current.FormulaChineseFont),
                StringComparison.OrdinalIgnoreCase);
    }

    private static InlineShape AddOleObject(Document document, Range range) =>
        document.InlineShapes.AddOLEObject(
            ClassType: FormulaOleContract.ProgId,
            LinkToFile: false,
            DisplayAsIcon: false,
            Range: range);

    private static void InitializeOle(
        InlineShape shape,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            format = shape.OLEFormat;
            oleObject = WordOleObjectAccessor.GetRunningObject(format);
            if (oleObject is not IVisualTeXFormulaObject formula)
                throw new InvalidOperationException(
                    "The inserted Word object does not expose the VisualTeX native OLE interface.");
            FormulaOleInterop.Initialize(formula, metadata, emfPath, pngPath);
            WordFormulaMetadataReader.CacheMetadata(shape, metadata);
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private static bool TryUpdateOle(
        InlineShape shape,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            try { format = shape.OLEFormat; }
            catch { return false; }
            try { oleObject = WordOleObjectAccessor.GetRunningObject(format); }
            catch { return false; }
            if (oleObject is not IVisualTeXFormulaObject formula) return false;
            FormulaOleInterop.Update(formula, metadata, emfPath, pngPath);
            WordFormulaMetadataReader.CacheMetadata(shape, metadata);
            return true;
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private Range DuplicateCurrentSelectionRange()
    {
        Selection? selection = null;
        Range? range = null;
        try
        {
            selection = _application.Selection;
            range = selection.Range;
            return range.Duplicate;
        }
        finally
        {
            Release(range);
            Release(selection);
        }
    }

    internal WordViewState CaptureFormulaFormatConversionViewState() =>
        CaptureViewState();

    internal void RestoreFormulaFormatConversionViewState(WordViewState? state)
    {
        if (state is null) return;
        Document? document = null;
        try
        {
            document = _application.ActiveDocument;
            RestoreViewState(document, state, preferredSelection: null);
        }
        catch
        {
            // View restoration must never turn an otherwise successful formula
            // conversion into an error. The original document is reactivated by
            // the conversion write path before this method normally runs.
        }
        finally { Release(document); }
    }

    private WordViewState CaptureViewState()
    {
        Selection? selection = null;
        Range? range = null;
        Window? window = null;
        try
        {
            selection = _application.Selection;
            range = selection.Range;
            try { window = _application.ActiveWindow; } catch { }
            int? vertical = null;
            int? horizontal = null;
            try { vertical = window?.VerticalPercentScrolled; } catch { }
            try { horizontal = window?.HorizontalPercentScrolled; } catch { }
            return new WordViewState
            {
                SelectionStart = range.Start,
                SelectionEnd = range.End,
                VerticalPercentScrolled = vertical,
                HorizontalPercentScrolled = horizontal,
            };
        }
        finally
        {
            Release(window);
            Release(range);
            Release(selection);
        }
    }

    private void RestoreViewState(
        Document? document,
        WordViewState? state,
        Range? preferredSelection)
    {
        if (document is null || state is null) return;
        Selection? selection = null;
        Range? fallback = null;
        Range? content = null;
        Window? window = null;
        try
        {
            selection = _application.Selection;
            if (preferredSelection is not null)
            {
                selection.SetRange(preferredSelection.Start, preferredSelection.End);
            }
            else
            {
                content = document.Content;
                var start = Math.Max(content.Start, Math.Min(state.SelectionStart, content.End));
                var end = Math.Max(start, Math.Min(state.SelectionEnd, content.End));
                object startValue = start;
                object endValue = end;
                fallback = document.Range(ref startValue, ref endValue);
                selection.SetRange(fallback.Start, fallback.End);
            }
            try { window = _application.ActiveWindow; } catch { }
            if (window is not null)
            {
                try
                {
                    if (state.HorizontalPercentScrolled.HasValue)
                        window.HorizontalPercentScrolled = state.HorizontalPercentScrolled.Value;
                }
                catch { }
                try
                {
                    if (state.VerticalPercentScrolled.HasValue)
                        window.VerticalPercentScrolled = state.VerticalPercentScrolled.Value;
                }
                catch { }
            }
        }
        catch { }
        finally
        {
            Release(window);
            Release(content);
            Release(fallback);
            Release(selection);
        }
    }

    private static InlineShape? FindMathTypeOleByRange(
        Document document,
        string? sourceObjectId,
        bool allowGlobalFallback = true)
    {
        if (!TryParseRangeReference(sourceObjectId, out var start, out var end))
            return null;

        Range? hintedRange = null;
        InlineShapes? hintedShapes = null;
        Range? content = null;
        Range? localRange = null;
        InlineShapes? localShapes = null;
        InlineShapes? shapes = null;
        try
        {
            try
            {
                hintedRange = document.Range(start, end);
                hintedShapes = hintedRange.InlineShapes;
                for (var index = 1; index <= hintedShapes.Count; index++)
                {
                    InlineShape? candidate = null;
                    try
                    {
                        candidate = hintedShapes[index];
                        if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                        var result = candidate;
                        candidate = null;
                        return result;
                    }
                    finally { Release(candidate); }
                }
            }
            catch
            {
                // Fall back to a document scan below. The source range can move
                // if surrounding prose changes while the VisualTeX editor is open.
            }

            // During whole-document format conversion targets are processed from
            // the end toward the start, so the captured MathType range remains a
            // reliable local locator. After the source has been deleted the exact
            // range is expected to contain no MathType object. Probe only a tiny
            // neighborhood for Word's occasional one/two-character range drift;
            // never enumerate the whole document merely to prove absence.
            try
            {
                content = document.Content;
                var localStart = Math.Max(content.Start, start - 2);
                var localEnd = Math.Min(content.End, Math.Max(end, start) + 2);
                if (localEnd > localStart)
                {
                    localRange = document.Range(localStart, localEnd);
                    localShapes = localRange.InlineShapes;
                    InlineShape? nearestLocal = null;
                    var nearestLocalDistance = int.MaxValue;
                    try
                    {
                        for (var index = 1; index <= localShapes.Count; index++)
                        {
                            InlineShape? candidate = null;
                            Range? candidateRange = null;
                            try
                            {
                                candidate = localShapes[index];
                                if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                                candidateRange = candidate.Range;
                                var distance = Math.Abs(candidateRange.Start - start);
                                if (distance >= nearestLocalDistance) continue;
                                Release(nearestLocal);
                                nearestLocal = candidate;
                                candidate = null;
                                nearestLocalDistance = distance;
                            }
                            finally
                            {
                                Release(candidateRange);
                                Release(candidate);
                            }
                        }
                        if (nearestLocal is not null && nearestLocalDistance <= 2)
                        {
                            var result = nearestLocal;
                            nearestLocal = null;
                            return result;
                        }
                    }
                    finally { Release(nearestLocal); }
                }
            }
            catch { }
            finally
            {
                Release(localShapes);
                localShapes = null;
                Release(localRange);
                localRange = null;
                Release(content);
                content = null;
            }

            if (!allowGlobalFallback) return null;

            shapes = document.InlineShapes;
            InlineShape? nearest = null;
            var nearestDistance = int.MaxValue;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? candidateRange = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    candidateRange = candidate.Range;
                    if (candidateRange.Start == start && candidateRange.End == end)
                    {
                        Release(nearest);
                        nearest = null;
                        var exact = candidate;
                        candidate = null;
                        return exact;
                    }
                    var distance = Math.Abs(candidateRange.Start - start);
                    if (distance >= nearestDistance) continue;
                    Release(nearest);
                    nearest = candidate;
                    candidate = null;
                    nearestDistance = distance;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }

            // A range is only a transient locator for third-party OLE. Avoid
            // accidentally editing a different MathType equation after large
            // document changes: permit only a very small local shift.
            if (nearest is not null && nearestDistance <= 2)
            {
                var result = nearest;
                nearest = null;
                return result;
            }
            Release(nearest);
            return null;
        }
        finally
        {
            Release(shapes);
            Release(localShapes);
            Release(localRange);
            Release(content);
            Release(hintedShapes);
            Release(hintedRange);
        }
    }

    private static InlineShape? FindByFormulaId(
        Document document,
        string formulaId,
        string? sourceObjectIdHint = null,
        bool allowGlobalFallback = true)
    {
        Bookmarks? identityBookmarks = null;
        Bookmark? identityBookmark = null;
        Range? identityRange = null;
        InlineShapes? identityShapes = null;
        Range? hintedRange = null;
        Range? content = null;
        InlineShapes? hintedShapes = null;
        Range? localRange = null;
        InlineShapes? localShapes = null;
        InlineShapes? shapes = null;
        try
        {
            // New and recently-read VisualTeX OLE formulas carry a durable VTO_
            // identity bookmark. Resolve that exact owner first. This avoids a
            // document-wide InlineShapes + embedded-metadata scan for every item
            // in a large format-conversion batch and also avoids stale COM
            // collection failures after many delete/insert mutations.
            try
            {
                identityBookmarks = document.Bookmarks;
                var identityName = WordFormulaMetadataReader.IdentityBookmarkName(formulaId);
                if (identityBookmarks.Exists(identityName))
                {
                    identityBookmark = identityBookmarks[identityName];
                    identityRange = identityBookmark.Range;
                    identityShapes = identityRange.InlineShapes;
                    for (var index = 1; index <= identityShapes.Count; index++)
                    {
                        InlineShape? candidate = null;
                        try
                        {
                            candidate = identityShapes[index];
                            if (!WordFormulaMetadataReader.IsNativeOle(candidate)) continue;
                            var metadata = WordFormulaMetadataReader.TryRead(candidate);
                            if (!string.Equals(
                                    metadata?.FormulaId,
                                    formulaId,
                                    StringComparison.OrdinalIgnoreCase))
                                continue;
                            var result = candidate;
                            candidate = null;
                            return result;
                        }
                        finally { Release(candidate); }
                    }
                }
            }
            catch
            {
                // Older documents or Word bookmark repair can invalidate this
                // fast locator. Preserve the range-hint/global compatibility path.
            }
            finally
            {
                Release(identityShapes);
                identityShapes = null;
                Release(identityRange);
                identityRange = null;
                Release(identityBookmark);
                identityBookmark = null;
                Release(identityBookmarks);
                identityBookmarks = null;
            }

            if (TryParseRangeReference(sourceObjectIdHint, out var start, out var end))
            {
                try
                {
                    content = document.Content;
                    if (start >= content.Start && end >= start && end <= content.End)
                    {
                        object startValue = start;
                        object endValue = end;
                        hintedRange = document.Range(ref startValue, ref endValue);
                        hintedShapes = hintedRange.InlineShapes;
                        for (var index = 1; index <= hintedShapes.Count; index++)
                        {
                            InlineShape? candidate = null;
                            try
                            {
                                candidate = hintedShapes[index];
                                var metadata = WordFormulaMetadataReader.TryRead(candidate);
                                if (string.Equals(
                                        metadata?.FormulaId,
                                        formulaId,
                                        StringComparison.OrdinalIgnoreCase))
                                {
                                    var result = candidate;
                                    candidate = null;
                                    return result;
                                }
                            }
                            finally { Release(candidate); }
                        }
                    }
                }
                catch
                {
                    // The document may have shifted while the editor was open.
                    // Fall back to the durable FormulaId scan below.
                }
                finally
                {
                    Release(hintedShapes);
                    hintedShapes = null;
                    Release(hintedRange);
                    hintedRange = null;
                    Release(content);
                    content = null;
                }
            }

            if (TryParseRangeReference(
                    sourceObjectIdHint,
                    out var localStartHint,
                    out var localEndHint))
            {
                try
                {
                    content = document.Content;
                    var localStart = Math.Max(content.Start, localStartHint - 32);
                    // Word can defer OLE deletion until after the temporary LaTeX
                    // bridge is inserted, shifting the still-live object forward.
                    // Search a bounded neighborhood rather than the full document.
                    var localEnd = Math.Min(
                        content.End,
                        Math.Max(localEndHint, localStartHint) + 16384);
                    if (localEnd > localStart)
                    {
                        localRange = document.Range(localStart, localEnd);
                        localShapes = localRange.InlineShapes;
                        for (var index = 1; index <= localShapes.Count; index++)
                        {
                            InlineShape? candidate = null;
                            try
                            {
                                candidate = localShapes[index];
                                if (!WordFormulaMetadataReader.IsNativeOle(candidate))
                                    continue;
                                var metadata = WordFormulaMetadataReader.TryRead(candidate);
                                if (!string.Equals(
                                        metadata?.FormulaId,
                                        formulaId,
                                        StringComparison.OrdinalIgnoreCase))
                                    continue;
                                var result = candidate;
                                candidate = null;
                                return result;
                            }
                            finally { Release(candidate); }
                        }
                    }
                }
                catch { }
                finally
                {
                    Release(localShapes);
                    localShapes = null;
                    Release(localRange);
                    localRange = null;
                    Release(content);
                    content = null;
                }
            }

            if (!allowGlobalFallback) return null;

            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                try
                {
                    shape = shapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (string.Equals(
                            metadata?.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        var result = shape;
                        shape = null;
                        return result;
                    }
                }
                finally { Release(shape); }
            }
            return null;
        }
        finally
        {
            Release(shapes);
            Release(hintedShapes);
            Release(hintedRange);
            Release(localShapes);
            Release(localRange);
            Release(content);
            Release(identityShapes);
            Release(identityRange);
            Release(identityBookmark);
            Release(identityBookmarks);
        }
    }

    private UndoRecord? BeginUndoRecord(string name)
    {
        UndoRecord? undoRecord = null;
        try
        {
            undoRecord = _application.UndoRecord;
            // Word exposes one process-wide Custom Undo Record stack. Starting a
            // second record from a helper that is already running inside a parent
            // transaction lets the inner EndCustomRecord close/reshape the outer
            // transaction. Format conversion deliberately wraps each destructive
            // source replacement in an outer record, while mature insertion paths
            // such as InsertMathTypeOle also call this helper. Treat an active
            // record as borrowed: the inner operation participates in it but must
            // neither start nor end another record.
            if (undoRecord.IsRecordingCustomRecord || undoRecord.CustomRecordLevel > 0)
            {
                Release(undoRecord);
                return null;
            }
            undoRecord.StartCustomRecord(name);
            return undoRecord;
        }
        catch
        {
            Release(undoRecord);
            return null;
        }
    }

    private static void EndUndoRecord(UndoRecord? undoRecord)
    {
        if (undoRecord is null) return;
        try { undoRecord.EndCustomRecord(); } catch { }
    }

    private static Range DuplicateOleRangeByFormulaId(
        Document document,
        string formulaId)
    {
        InlineShape? refreshedShape = null;
        Range? refreshedRange = null;
        try
        {
            refreshedShape = FindByFormulaId(document, formulaId)
                ?? throw new InvalidOperationException(
                    "Word lost the updated VisualTeX OLE formula during numbered-layout migration.");
            refreshedRange = refreshedShape.Range.Duplicate;
            var result = refreshedRange;
            refreshedRange = null;
            return result;
        }
        finally
        {
            Release(refreshedRange);
            Release(refreshedShape);
        }
    }

    private static void TryReconcileShape(
        Document document,
        InlineShape shape,
        FormulaMetadata metadata,
        bool numberingOrderMayHaveChanged = true,
        bool reuseExistingNumberedTableFormatting = false,
        Table? knownNumberedTable = null,
        bool numberingScaffoldOnly = false)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            if (string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
            {
                RemoveInlineBaselineSentinel(document, metadata.FormulaId);
                ResetDisplayFormulaPosition(range);
            }
            if (numberingScaffoldOnly && metadata.Numbered)
            {
                WordEquationNumbering.BuildFormulaNumberingScaffoldForConversion(
                    document,
                    range,
                    shape.Height,
                    metadata,
                    knownNumberedTable);
            }
            else
            {
                WordEquationNumbering.TryReconcileFormula(
                    document,
                    range,
                    shape.Height,
                    metadata,
                    numberingOrderMayHaveChanged,
                    reuseExistingNumberedTableFormatting,
                    knownNumberedTable);
            }
        }
        finally { Release(range); }
    }

    private static bool NumberingOrderMayHaveChanged(
        FormulaMetadata? originalMetadata,
        FormulaMetadata metadata)
    {
        if (originalMetadata is null) return true;
        return originalMetadata.Numbered != metadata.Numbered
            || !string.Equals(
                originalMetadata.DisplayMode,
                metadata.DisplayMode,
                StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryReadStoredWordInlineOleSize(
        FormulaMetadata? metadata,
        out float width,
        out float height)
    {
        width = 0;
        height = 0;
        if (metadata?.WordInlineOleWidthPt is not > 0
            || metadata.WordInlineOleHeightPt is not > 0)
            return false;
        width = (float)metadata.WordInlineOleWidthPt.Value;
        height = (float)metadata.WordInlineOleHeightPt.Value;
        return width > 0
            && height > 0
            && !float.IsNaN(width)
            && !float.IsInfinity(width)
            && !float.IsNaN(height)
            && !float.IsInfinity(height);
    }

    private static void StoreWordInlineOleSize(
        FormulaMetadata metadata,
        float width,
        float height,
        bool inline)
    {
        if (!inline
            || width <= 0
            || height <= 0
            || float.IsNaN(width)
            || float.IsInfinity(width)
            || float.IsNaN(height)
            || float.IsInfinity(height))
        {
            metadata.WordInlineOleWidthPt = null;
            metadata.WordInlineOleHeightPt = null;
            return;
        }
        metadata.WordInlineOleWidthPt = width;
        metadata.WordInlineOleHeightPt = height;
    }

    private static float ReadOmmlFontSize(
        Bookmark bookmark,
        FormulaMetadata metadata)
    {
        Range? range = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            font = range.Font;
            return TryNormalizeDefinedWordFontSize(font.Size, out var size)
                ? size
                : FormulaFontSize.ResolveSemanticFontSize(metadata);
        }
        catch
        {
            return FormulaFontSize.ResolveSemanticFontSize(metadata);
        }
        finally
        {
            Release(font);
            Release(range);
        }
    }

    private static void ApplyOmmlTypography(
        Range equationRange,
        double fontSizePt,
        FormulaMetadata metadata)
    {
        Document? document = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            document = equationRange.Document;
            ApplyDocumentOmmlMathFont(document, metadata);

            var normalized = FormulaFontSize.Normalize(fontSizePt);
            font = equationRange.Font;
            font.Position = 0;
            font.Size = normalized;
            try { font.SizeBi = normalized; } catch { }
        }
        finally
        {
            Release(font);
            Release(document);
        }

        if (string.Equals(
                metadata.DisplayMode,
                "inline",
                StringComparison.OrdinalIgnoreCase))
            StabilizeInlineOmmlFractionLineGrid(equationRange);
    }

    private static bool IsNumberedBlockOmml(FormulaMetadata metadata) =>
        metadata.Numbered
        && string.Equals(
            metadata.DisplayMode,
            "block",
            StringComparison.OrdinalIgnoreCase);

    private static bool TryNormalizeDefinedWordFontSize(
        float value,
        out float fontSizePt)
    {
        fontSizePt = FormulaFontSize.DefaultPt;
        // Word uses 9,999,999 for mixed/undefined Font.Size. Fall back to the
        // persisted semantic size instead of treating that sentinel as a real value.
        if (value <= 0f
            || value >= 1000f
            || float.IsNaN(value)
            || float.IsInfinity(value))
            return false;
        fontSizePt = FormulaFontSize.Normalize(value);
        return true;
    }

    // Word exposes one native OpenType MATH face per document. This is intentionally
    // document-scoped: changing it reflows every existing OMath in the document; it is
    // not a per-formula font override and must never be emulated with m:nor/run splitting.
    private static void ApplyDocumentOmmlMathFont(
        Document document,
        FormulaMetadata metadata)
    {
        var requested = ResolveDocumentOmmlMathFont(metadata.FormulaLetterFont);
        if (string.Equals(
                requested,
                WordOfficeMathFontLoader.LatinModernMathFamily,
                StringComparison.OrdinalIgnoreCase))
            WordOfficeMathFontLoader.EnsureLoaded();

        string current;
        try { current = document.OMathFontName ?? string.Empty; }
        catch (COMException error)
        {
            throw new InvalidOperationException(
                "Word could not read this document's native Office Math font setting.",
                error);
        }
        if (string.Equals(current, requested, StringComparison.OrdinalIgnoreCase))
            return;

        try { document.OMathFontName = requested; }
        catch (COMException error)
        {
            throw new InvalidOperationException(
                $"Word could not set this document's native Office Math font to '{requested}'.",
                error);
        }

        string applied;
        try { applied = document.OMathFontName ?? string.Empty; }
        catch (COMException error)
        {
            throw new InvalidOperationException(
                "Word could not verify the document-level Office Math font after applying it.",
                error);
        }
        if (!string.Equals(applied, requested, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"Word rejected '{requested}' as the document's native Office Math font. "
                + "Install a valid OpenType math font or repair the VisualTeX Office integration.");
        }
    }

    private static string ResolveDocumentOmmlMathFont(string? preference)
    {
        // Word native OMath has one document-wide math font and requires an
        // OpenType MATH table. VisualTeX's text-face choices (Times, Palatino,
        // Helvetica/Arial) cannot safely replace that math font, so they fall back
        // to the bundled Latin Modern Math instead of flattening individual runs.
        return (preference ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "cambria" => "Cambria Math",
            "stix" => "STIX Two Math",
            _ => WordOfficeMathFontLoader.LatinModernMathFamily,
        };
    }

    private static string ApplyOmmlTypographyXml(
        string omml,
        double fontSizePt,
        FormulaMetadata metadata)
    {
        if (string.IsNullOrWhiteSpace(omml))
            throw new InvalidDataException("VisualTeX produced an empty OMML payload.");
        var equation = XElement.Parse(
            omml,
            LoadOptions.PreserveWhitespace);
        XNamespace word =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace math =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        ApplyBulkOmmlTypographyXml(
            equation,
            fontSizePt,
            metadata,
            word,
            math);
        return equation.ToString(SaveOptions.DisableFormatting);
    }

    private static string ResolveOmmlChineseFont(string? preference) =>
        (preference ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "songti" => "SimSun",
            "kaiti" => "KaiTi",
            "heiti" => "SimHei",
            _ => "Microsoft YaHei",
        };

    private static bool IsChineseOmmlCodePoint(int codePoint) =>
        codePoint is >= 0x2E80 and <= 0x2FFF
        || codePoint is >= 0x3000 and <= 0x303F
        || codePoint is >= 0x31C0 and <= 0x31EF
        || codePoint is >= 0x3400 and <= 0x4DBF
        || codePoint is >= 0x4E00 and <= 0x9FFF
        || codePoint is >= 0xF900 and <= 0xFAFF
        || codePoint is >= 0xFF00 and <= 0xFFEF
        || codePoint is >= 0x20000 and <= 0x2FA1F;

    private static void StabilizeInlineOmmlFractionLineGrid(Range equationRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? paragraphFormat = null;
        try
        {
            maths = equationRange.OMaths;
            if (maths.Count != 1) return;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathInline) return;

            sections = equationRange.Sections;
            if (sections.Count == 0) return;
            section = sections[1];
            pageSetup = section.PageSetup;
            if (pageSetup.LayoutMode is not (
                    WdLayoutMode.wdLayoutModeLineGrid
                    or WdLayoutMode.wdLayoutModeGrid
                    or WdLayoutMode.wdLayoutModeGenko))
                return;
            if (pageSetup.LinesPage <= 0) return;

            var equationXml = WordOmmlConverter.ExtractSingleOMath(equationRange.WordOpenXML);
            var equation = XDocument.Parse(equationXml, LoadOptions.None);
            const string officeMathNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            if (!equation.Descendants(XName.Get("f", officeMathNamespace)).Any())
                return;

            paragraphs = equationRange.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            paragraphFormat = paragraphRange.ParagraphFormat;
            if (paragraphFormat.DisableLineHeightGrid == -1) return;

            // Word can quantize an inline stacked fraction to an extra document-
            // grid row when a descender (for example lowercase p/g/q/y) crosses
            // a line-grid threshold. The equation itself is still valid, but its
            // line box can suddenly double. Disable only line-grid snapping for
            // this paragraph; preserve its line-spacing rule, spacing before/
            // after, indents and all other user formatting.
            paragraphFormat.DisableLineHeightGrid = -1;
        }
        catch (COMException)
        {
            // Layout-grid compatibility is protective. Unsupported section/page
            // settings must never invalidate an already inserted native OMath.
        }
        catch (InvalidDataException)
        {
            // If Word exposes incomplete OpenXML during a transient layout pass,
            // keep the native equation rather than failing the Office operation.
        }
        catch (System.Xml.XmlException)
        {
            // A transient WordOpenXML snapshot must not invalidate the formula.
        }
        finally
        {
            Release(paragraphFormat);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(math);
            Release(maths);
        }
    }

    private static bool HasReusableStandaloneDisplayOmmlEditHost(
        Range formulaRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        Fields? fields = null;
        InlineShapes? shapes = null;
        try
        {
            // Same-state unnumbered Display edits replace only this OMath and do
            // not touch its paragraph topology. Avoid the full paragraph/prefix/
            // suffix validation required before an unnumbered→numbered structural
            // conversion; prove only the object we will mutate.
            if ((bool)formulaRange.get_Information(WdInformation.wdWithInTable))
                return false;
            maths = formulaRange.OMaths;
            if (maths.Count != 1) return false;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay) return false;
            mathRange = math.Range.Duplicate;
            if (mathRange.Start != formulaRange.Start
                || mathRange.End != formulaRange.End)
                return false;
            fields = formulaRange.Fields;
            shapes = formulaRange.InlineShapes;
            return fields.Count == 0 && shapes.Count == 0;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(shapes);
            Release(fields);
            Release(mathRange);
            Release(math);
            Release(maths);
        }
    }

    private static bool HasReusableStandaloneDisplayOmmlHost(
        Document document,
        Range formulaRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefix = null;
        Range? suffix = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            if ((bool)formulaRange.get_Information(WdInformation.wdWithInTable))
                return false;
            maths = formulaRange.OMaths;
            if (maths.Count != 1) return false;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay) return false;
            mathRange = math.Range.Duplicate;
            if (mathRange.Start != formulaRange.Start
                || mathRange.End != formulaRange.End)
                return false;
            fields = formulaRange.Fields;
            shapes = formulaRange.InlineShapes;
            if (fields.Count != 0 || shapes.Count != 0) return false;

            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (formulaRange.Start < paragraphRange.Start
                || formulaRange.End > paragraphRange.End)
                return false;
            prefix = document.Range(paragraphRange.Start, formulaRange.Start);
            suffix = document.Range(formulaRange.End, paragraphRange.End);
            return !ContainsVisibleBodyText(prefix.Text)
                && !ContainsVisibleBodyText(suffix.Text);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(suffix);
            Release(prefix);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(mathRange);
            Release(math);
            Release(maths);
        }
    }

    private static Range? TryResolveStandaloneOmmlFromLiveRange(Range? liveRange)
    {
        if (liveRange is null) return null;
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        try
        {
            if ((bool)liveRange.get_Information(WdInformation.wdWithInTable))
                return null;
            maths = liveRange.OMaths;
            if (maths.Count != 1) return null;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay) return null;
            mathRange = math.Range.Duplicate;
            var result = mathRange;
            mathRange = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(mathRange);
            Release(math);
            Release(maths);
        }
    }

    private static Range? TryResolveNumberedOmmlFromKnownTable(Table? table)
    {
        if (table is null) return null;
        Cell? formulaCell = null;
        Range? formulaCellRange = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? equationRange = null;
        try
        {
            if (table.Rows.Count != 1 || table.Columns.Count != 3)
                return null;
            formulaCell = table.Cell(1, 2);
            formulaCellRange = formulaCell.Range;
            maths = formulaCellRange.OMaths;
            if (maths.Count != 1) return null;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay) return null;
            equationRange = math.Range.Duplicate;
            var result = equationRange;
            equationRange = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(equationRange);
            Release(math);
            Release(maths);
            Release(formulaCellRange);
            Release(formulaCell);
        }
    }

    private static Range ResolveNumberedOmmlFromNumberingOwner(
        Document document,
        string formulaId)
    {
        Range? ownerRange = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? equationRange = null;
        InlineShapes? shapes = null;
        Tables? ownerTables = null;
        Table? ownerTable = null;
        Cell? formulaCell = null;
        Range? formulaCellRange = null;
        OMaths? tableMaths = null;
        OMath? tableMath = null;
        try
        {
            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(
                    document,
                    formulaId)
                ?? throw new InvalidDataException(
                    $"The numbered OMML formula {formulaId} has no numbering owner paragraph.");
            if ((bool)ownerRange.get_Information(WdInformation.wdWithInTable))
            {
                // Current native numbered OMML deliberately owns one exact 1x3
                // table: true Display OMath in cell (1,2), ordinary direct SEQ
                // number in cell (1,3). Resolve the live center OMath from that
                // physical owner instead of applying the retired table-free/#SEQ
                // invariant. FindNumberingOwnerRange already resolved the table
                // through the durable VTEq_ visible-number bookmark.
                ownerTables = ownerRange.Tables;
                if (ownerTables.Count != 1)
                    throw new InvalidDataException(
                        $"The numbered OMML table owner for {formulaId} is ambiguous.");
                ownerTable = ownerTables[1];
                if (ownerTable.Rows.Count != 1 || ownerTable.Columns.Count != 3)
                    throw new InvalidDataException(
                        $"The numbered OMML table owner for {formulaId} is not the managed 1x3 host.");
                formulaCell = ownerTable.Cell(1, 2);
                formulaCellRange = formulaCell.Range;
                tableMaths = formulaCellRange.OMaths;
                if (tableMaths.Count != 1)
                    throw new InvalidDataException(
                        $"The numbered OMML center cell for {formulaId} contains {tableMaths.Count} equations instead of exactly one.");
                tableMath = tableMaths[1];
                if (tableMath.Type != WdOMathType.wdOMathDisplay)
                    throw new InvalidDataException(
                        $"The numbered OMML center cell for {formulaId} is not true Word Display math.");
                if (formulaCellRange.Fields.Count != 0)
                    throw new InvalidDataException(
                        $"The numbered OMML center cell for {formulaId} contains a number field.");
                equationRange = tableMath.Range.Duplicate;
                var tableResult = equationRange;
                equationRange = null;
                return tableResult;
            }
            shapes = ownerRange.InlineShapes;
            if (shapes.Count != 0)
                throw new InvalidDataException(
                    $"The numbered OMML owner for {formulaId} contains an inline object.");
            maths = ownerRange.OMaths;
            if (maths.Count != 1)
            {
                if (string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                {
                    OMaths? documentMaths = null;
                    try
                    {
                        documentMaths = document.OMaths;
                        Console.WriteLine(
                            $"    [diagnostic] numbered OMML owner formulaId={formulaId} range={ownerRange.Start}:{ownerRange.End} textCodes={string.Join(",", (ownerRange.Text ?? string.Empty).Select(character => $"U+{(int)character:X4}"))} ownerMaths={maths.Count} documentMaths={documentMaths.Count}");
                        for (var index = 1; index <= documentMaths.Count; index++)
                        {
                            OMath? candidate = null;
                            Range? candidateRange = null;
                            Paragraphs? candidateParagraphs = null;
                            Paragraph? candidateParagraph = null;
                            Range? candidateParagraphRange = null;
                            try
                            {
                                candidate = documentMaths[index];
                                candidateRange = candidate.Range;
                                candidateParagraphs = candidateRange.Paragraphs;
                                candidateParagraph = candidateParagraphs[1];
                                candidateParagraphRange = candidateParagraph.Range;
                                Console.WriteLine(
                                    $"    [diagnostic] OMath#{index} type={candidate.Type} range={candidateRange.Start}:{candidateRange.End} textCodes={string.Join(",", (candidateRange.Text ?? string.Empty).Select(character => $"U+{(int)character:X4}"))} paragraph={candidateParagraphRange.Start}:{candidateParagraphRange.End} paragraphCodes={string.Join(",", (candidateParagraphRange.Text ?? string.Empty).Select(character => $"U+{(int)character:X4}"))}");
                            }
                            finally
                            {
                                Release(candidateParagraphRange);
                                Release(candidateParagraph);
                                Release(candidateParagraphs);
                                Release(candidateRange);
                                Release(candidate);
                            }
                        }
                    }
                    finally { Release(documentMaths); }
                }
                throw new InvalidDataException(
                    $"The numbered OMML owner for {formulaId} contains {maths.Count} OMath equations instead of exactly one.");
            }
            math = maths[1];
            equationRange = math.Range.Duplicate;
            var result = equationRange;
            equationRange = null;
            return result;
        }
        finally
        {
            Release(tableMath);
            Release(tableMaths);
            Release(formulaCellRange);
            Release(formulaCell);
            Release(ownerTable);
            Release(ownerTables);
            Release(shapes);
            Release(equationRange);
            Release(math);
            Release(maths);
            Release(ownerRange);
        }
    }

    private static void TryReconcileOmml(
        Document document,
        Bookmark bookmark,
        Range equationRange,
        FormulaMetadata metadata,
        bool numberingOrderMayHaveChanged = true,
        bool reuseExistingNumberedTableFormatting = false,
        Table? knownNumberedTable = null,
        bool numberingScaffoldOnly = false,
        bool deferNativeOmmlShapeFinalization = false,
        bool deferNativeOmmlShapeCreation = false,
        bool deferNativeOmmlMetadataPersistence = false,
        string? preparedUnnumberedOmml = null)
    {
        var display = string.Equals(
            metadata.DisplayMode,
            "block",
            StringComparison.Ordinal);
        if (display)
            ResetDisplayFormulaPosition(equationRange);
        else
            RemoveInlineBaselineSentinel(document, metadata.FormulaId);
        // The exact equation range is already available. Re-reading it through
        // the bookmark and enumerating all document OMaths only to estimate a
        // height made this local operation scale with total formula count.
        var height = (float)Math.Max(
            11,
            FormulaFontSize.ResolveSemanticFontSize(metadata) * 1.55);
        if (!metadata.Numbered)
        {
            // An edit can simultaneously toggle a numbered native-OMML formula to
            // unnumbered. In the current direct-SEQ 1x3 host that structural cleanup
            // must dismantle the table back to one standalone wdOMathDisplay. The
            // old early return left the now-unnumbered formula stranded in its 1x3
            // shell. Re-enter the normal numbering reconciler; when no artifacts
            // exist it is still a cheap local paragraph-format operation.
            WordEquationNumbering.TryReconcileFormula(
                document,
                equationRange,
                height,
                metadata,
                numberingOrderMayHaveChanged,
                reuseExistingNumberedTableFormatting,
                knownNumberedTable,
                deferNativeOmmlShapeFinalization,
                deferNativeOmmlShapeCreation,
                deferNativeOmmlMetadataPersistence,
                preparedUnnumberedOmml);
            return;
        }
        if (numberingScaffoldOnly)
        {
            WordEquationNumbering.BuildFormulaNumberingScaffoldForConversion(
                document,
                equationRange,
                height,
                metadata,
                knownNumberedTable);
        }
        else
        {
            WordEquationNumbering.TryReconcileFormula(
                document,
                equationRange,
                height,
                metadata,
                numberingOrderMayHaveChanged,
                reuseExistingNumberedTableFormatting,
                knownNumberedTable,
                deferNativeOmmlShapeFinalization,
                deferNativeOmmlShapeCreation,
                deferNativeOmmlMetadataPersistence,
                preparedUnnumberedOmml);
        }
    }

    private static void Configure(
        InlineShape shape,
        FormulaMetadata metadata,
        float maxWidth,
        float maxHeight,
        string imagePath,
        float exportedHeight,
        float? exportedBaseline,
        bool alignInline,
        bool nativeOleKnown = false,
        bool trustExportDimensions = false)
    {
        // maxWidth/maxHeight are already the SVG's physical size after the
        // 96 dpi CSS-pixel to 72 dpi Word-point conversion. A 12 pt minimum
        // width scales narrow inline formulas (notably x) far above their
        // semantic font size. Keep only a one-point safety floor.
        var width = Math.Max(1f, maxWidth);
        var height = Math.Max(1f, maxHeight);
        if (!trustExportDimensions)
        {
            using var image = Image.FromFile(imagePath);
            var ratio = image.Width / (float)Math.Max(1, image.Height);
            height = width / ratio;
            if (maxHeight > 0 && height > maxHeight)
            {
                height = maxHeight;
                width = height * ratio;
            }
        }
        // An OLE object is initially created with the placeholder preview's 4:1
        // aspect ratio. Setting only Width while aspect-ratio locking is enabled
        // therefore distorts the real formula. Apply both natural dimensions
        // explicitly, then lock the resolved ratio for later user resizing.
        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
        shape.Width = width;
        shape.Height = height;
        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
        if (!nativeOleKnown && !WordFormulaMetadataReader.IsNativeOle(shape))
        {
            var encoded = FormulaMetadataCodec.Encode(metadata);
            shape.Title = encoded;
            shape.AlternativeText = encoded;
        }
        if (alignInline)
            ApplyInlineBaseline(
                shape,
                shape.Height,
                exportedHeight,
                exportedBaseline,
                FormulaFontSize.ResolveSemanticFontSize(metadata));
        else
            ResetDisplayFormulaPosition(shape);
    }

    private static bool ShouldAlignInline(InlineShape shape, FormulaMetadata metadata)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            return ShouldAlignInline(range, metadata);
        }
        finally { Release(range); }
    }

    private static bool ShouldAlignInline(Range formulaRange, FormulaMetadata metadata)
    {
        if (string.Equals(metadata.DisplayMode, "inline", StringComparison.Ordinal))
            return true;
        return HasVisibleSurroundingText(formulaRange);
    }

    private static bool HasVisibleSurroundingText(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? before = null;
        Range? after = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (formulaRange.Start > paragraphRange.Start)
            {
                before = paragraphRange.Duplicate;
                before.SetRange(paragraphRange.Start, formulaRange.Start);
                if (ContainsVisibleBodyText(before.Text)) return true;
            }
            if (formulaRange.End < paragraphRange.End)
            {
                after = paragraphRange.Duplicate;
                after.SetRange(formulaRange.End, paragraphRange.End);
                if (ContainsVisibleBodyText(after.Text)) return true;
            }
            return false;
        }
        catch { return false; }
        finally
        {
            Release(after);
            Release(before);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool ContainsVisibleBodyText(string? value)
    {
        if (string.IsNullOrEmpty(value)) return false;
        foreach (var character in value!)
        {
            if (character is '\r' or '\n' or '\t' or '\v' or '\a' or '\u0001' or '\u200B' or '\u200C')
                continue;
            if (!char.IsWhiteSpace(character)) return true;
        }
        return false;
    }

    private readonly struct InlineOlePreviewMetrics
    {
        internal InlineOlePreviewMetrics(float inkHeightRatio, float bottomWhitespaceRatio)
        {
            InkHeightRatio = inkHeightRatio;
            BottomWhitespaceRatio = bottomWhitespaceRatio;
        }

        internal float InkHeightRatio { get; }
        internal float BottomWhitespaceRatio { get; }
    }

    private static float? ReadDefinedShapeFontPosition(InlineShape shape)
    {
        Range? range = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = shape.Range;
            font = range.Font;
            var position = font.Position;
            return position == (int)WdConstants.wdUndefined
                || position < -256
                || position > 256
                    ? null
                    : position;
        }
        catch { return null; }
        finally
        {
            Release(font);
            Release(range);
        }
    }

    private static int ResolveInlineShapeInsertionIndex(Document document, int position)
    {
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? range = null;
        try
        {
            shapes = document.InlineShapes;
            var preceding = 0;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(range);
                range = null;
                Release(shape);
                shape = shapes[index];
                range = shape.Range;
                if (range.Start >= position) break;
                preceding++;
            }
            return preceding + 1;
        }
        finally
        {
            Release(range);
            Release(shape);
            Release(shapes);
        }
    }

    private static InlineShape? FindMathTypeOleAtIndex(Document document, int index)
    {
        InlineShapes? shapes = null;
        InlineShape? candidate = null;
        try
        {
            shapes = document.InlineShapes;
            if (index < 1 || index > shapes.Count) return null;
            candidate = shapes[index];
            if (!MathTypeOleInterop.IsMathTypeOle(candidate)) return null;
            var result = candidate;
            candidate = null;
            return result;
        }
        catch { return null; }
        finally
        {
            Release(candidate);
            Release(shapes);
        }
    }

    private static InlineShape? FindMathTypeOleInParagraphAtPosition(
        Document document,
        int position)
    {
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? shapes = null;
        try
        {
            var safePosition = Math.Max(
                document.Content.Start,
                Math.Min(document.Content.End - 1, position));
            probe = document.Range(safePosition, safePosition);
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            shapes = paragraphRange.InlineShapes;
            InlineShape? match = null;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    if (match is not null)
                    {
                        Release(match);
                        return null;
                    }
                    match = candidate;
                    candidate = null;
                }
                finally { Release(candidate); }
            }
            return match;
        }
        catch { return null; }
        finally
        {
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    private static InlineShape? FindMathTypeOleInLocalWindow(
        Document document,
        int position)
    {
        Range? window = null;
        InlineShapes? shapes = null;
        InlineShape? best = null;
        var bestDistance = int.MaxValue;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = document.Content.End;
            var start = Math.Max(contentStart, position - 8);
            var end = Math.Min(contentEnd, position + 256);
            if (end <= start) end = Math.Min(contentEnd, start + 1);
            window = document.Range(start, end);
            shapes = window.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? candidateRange = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    candidateRange = candidate.Range;
                    var distance = Math.Abs(candidateRange.Start - position);
                    if (distance >= bestDistance) continue;
                    Release(best);
                    best = candidate;
                    candidate = null;
                    bestDistance = distance;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }
            return best;
        }
        catch
        {
            Release(best);
            return null;
        }
        finally
        {
            Release(shapes);
            Release(window);
        }
    }

    private static InlineShape? FindMathTypeOleNearPosition(
        Document document,
        int position)
    {
        InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            InlineShape? best = null;
            var bestDistance = int.MaxValue;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? range = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    range = candidate.Range;
                    var distance = Math.Abs(range.Start - position);
                    if (distance > 3 || distance >= bestDistance) continue;
                    Release(best);
                    best = candidate;
                    candidate = null;
                    bestDistance = distance;
                }
                catch { }
                finally
                {
                    Release(range);
                    Release(candidate);
                }
            }
            return best;
        }
        finally { Release(shapes); }
    }

    private static void BuildIndependentMathTypeDisplayScaffold(
        Document document,
        InlineShape shape,
        bool numbered,
        string mathTypeNumberPosition,
        MathTypeWordOpenXml.NumberTemplate? numberTemplate)
    {
        Range? shapeRange = null;
        Range? separator = null;
        Field? placeRef = null;
        try
        {
            var numberOnLeft = numbered && string.Equals(
                mathTypeNumberPosition,
                "left",
                StringComparison.OrdinalIgnoreCase);
            var numberOnRight = numbered && string.Equals(
                mathTypeNumberPosition,
                "right",
                StringComparison.OrdinalIgnoreCase);
            if (numbered && !numberOnLeft && !numberOnRight)
                throw new InvalidDataException(
                    "MathType equation number position must be left or right.");
            if (numbered && numberTemplate is null)
                throw new InvalidDataException(
                    "A numbered MathType display equation requires an MTPlaceRef template.");

            shapeRange = shape.Range;
            if (!numberOnLeft)
            {
                // MathType's display style centers the equation on its center tab.
                separator = document.Range(shapeRange.Start, shapeRange.Start);
                separator.Text = "\t";
                Release(separator);
                separator = null;
                Release(shapeRange);
                shapeRange = shape.Range;
            }

            if (!numbered) return;

            if (numberOnLeft)
            {
                // Do not create MTPlaceRef first and then insert a tab at its end.
                // Word can expand the field's boundary to absorb that character,
                // leaving no ordinary separator between the number and the OLE.
                // Materialize the tab first, then create MTPlaceRef before it so
                // the final structure is unambiguously FIELD + TAB + OLE.
                var numberInsertionPosition = shapeRange.Start;
                separator = document.Range(
                    numberInsertionPosition,
                    numberInsertionPosition);
                separator.Text = "\t";
                Release(separator);
                separator = null;
                placeRef = CreateIndependentMathTypePlaceRef(
                    document,
                    numberInsertionPosition,
                    numberTemplate!);
            }
            else
            {
                separator = document.Range(shapeRange.End, shapeRange.End);
                separator.Text = "\t";
                var placeRefPosition = separator.End;
                placeRef = CreateIndependentMathTypePlaceRef(
                    document,
                    placeRefPosition,
                    numberTemplate!);
            }
            try { placeRef.ShowCodes = false; } catch { }
        }
        finally
        {
            Release(placeRef);
            Release(separator);
            Release(shapeRange);
        }
    }

    private static MathTypeWordOpenXml.NumberTemplate ReadMathTypePlaceRefTemplateForShape(
        Document document,
        InlineShape shape,
        string numberPosition)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Field? placeRef = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display equation does not occupy one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var numberOnLeft = string.Equals(
                numberPosition,
                "left",
                StringComparison.OrdinalIgnoreCase);
            placeRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft)
                ?? throw new InvalidOperationException(
                    "The numbered MathType equation has no readable MTPlaceRef field.");
            return ReadMathTypePlaceRefTemplate(document, placeRef);
        }
        finally
        {
            Release(placeRef);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static MathTypeWordOpenXml.NumberTemplate ResolveMathTypeEditNumberTemplate(
        Document document,
        InlineShape shape,
        MathTypeWordOpenXml.NumberTemplate? sourceTemplate,
        out int createdSectionBreakCodeStart)
    {
        createdSectionBreakCodeStart = -1;
        if (sourceTemplate is not null)
            return sourceTemplate;

        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Field? nearest = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display equation does not occupy one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            nearest = FindNearestMathTypePlaceRefField(
                document,
                shapeRange.Start,
                paragraphRange.Start,
                paragraphRange.End);
            var documentNumberFormat = EquationNumberFormat.Resolve(
                WordEquationNumbering.GetEquationNumberFormatId(document));
            var template = nearest is not null
                ? ReadMathTypePlaceRefTemplate(document, nearest)
                : MathTypeWordOpenXml.CreateVisualTeXNumberTemplate(
                    documentNumberFormat.Id);

            if (documentNumberFormat.UsesHeading
                && MathTypeNumberTemplateUsesHeading(template))
            {
                EnsureMathTypeHeadingScopeState(
                    document,
                    shapeRange.Start,
                    documentNumberFormat,
                    out createdSectionBreakCodeStart);
            }
            return template;
        }
        finally
        {
            Release(nearest);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static void RebuildMathTypeDisplayScaffold(
        Document document,
        InlineShape shape,
        bool numbered,
        string numberPosition,
        MathTypeWordOpenXml.NumberTemplate? numberTemplate)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? paragraphShapes = null;
        Field? placeRef = null;
        Range? placeRefCode = null;
        Range? placeRefResult = null;
        Range? fieldSpan = null;
        Range? prefix = null;
        Range? suffix = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "MathType numbering can only be changed for a standalone display equation.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            paragraphShapes = paragraphRange.InlineShapes;
            if (paragraphShapes.Count != 1)
                throw new InvalidOperationException(
                    "MathType numbering cannot be changed in a paragraph containing other inline objects.");

            // Remove only the native MathType number field. Nested SEQ fields live
            // inside its code range and disappear with the outer MACROBUTTON.
            placeRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft: true)
                ?? FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft: false);
            if (placeRef is not null)
            {
                placeRefCode = placeRef.Code;
                placeRefResult = placeRef.Result;
                var fieldStart = Math.Max(
                    paragraphRange.Start,
                    placeRefCode.Start - 1);
                var fieldEnd = Math.Min(
                    Math.Max(paragraphRange.Start, paragraphRange.End - 1),
                    placeRefResult.End + 1);
                fieldSpan = document.Range(fieldStart, Math.Max(fieldStart, fieldEnd));
                fieldSpan.Delete();
            }

            Release(shapeRange);
            shapeRange = shape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            var bodyEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            prefix = document.Range(paragraphRange.Start, shapeRange.Start);
            suffix = document.Range(shapeRange.End, bodyEnd);
            if (!IsMathTypeDisplayScaffoldWhitespace(prefix.Text)
                || !IsMathTypeDisplayScaffoldWhitespace(suffix.Text))
                throw new InvalidOperationException(
                    "MathType numbering was not changed because the display paragraph contains user text outside the equation.");

            // Delete from the end first so the OLE range remains stable while its
            // surrounding tabs are normalized back to a bare equation object.
            if (suffix.End > suffix.Start) suffix.Delete();
            Release(shapeRange);
            shapeRange = shape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            Release(prefix);
            prefix = document.Range(paragraphRange.Start, shapeRange.Start);
            if (prefix.End > prefix.Start) prefix.Delete();

            BuildIndependentMathTypeDisplayScaffold(
                document,
                shape,
                numbered,
                numberPosition,
                numberTemplate);
            ConfigureNewMathTypeDisplayEquation(
                document,
                shape,
                numbered,
                numberPosition);
        }
        finally
        {
            Release(suffix);
            Release(prefix);
            Release(fieldSpan);
            Release(placeRefResult);
            Release(placeRefCode);
            Release(placeRef);
            Release(paragraphShapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static bool IsMathTypeDisplayScaffoldWhitespace(string? text)
    {
        if (string.IsNullOrEmpty(text)) return true;
        foreach (var character in text)
        {
            if (char.IsWhiteSpace(character)
                || character is '\t' or '\r' or '\n' or '\v'
                    or '\u0013' or '\u0014' or '\u0015')
                continue;
            return false;
        }
        return true;
    }

    private static Field CreateIndependentMathTypePlaceRef(
        Document document,
        int position,
        MathTypeWordOpenXml.NumberTemplate template)
    {
        if (template.Segments.Count == 0)
            throw new InvalidDataException("MathType MTPlaceRef numbering template is empty.");

        Range? insertion = null;
        Range? outerCode = null;
        Field? outer = null;
        Field? nested = null;
        try
        {
            insertion = document.Range(position, position);
            outer = document.Fields.Add(
                insertion,
                WdFieldType.wdFieldMacroButton,
                "MTPlaceRef",
                false);

            foreach (var segment in template.Segments)
            {
                Release(outerCode);
                outerCode = outer.Code;
                Release(insertion);
                insertion = document.Range(outerCode.End, outerCode.End);
                if (!segment.IsField)
                {
                    if (!string.IsNullOrEmpty(segment.Value))
                        insertion.InsertAfter(segment.Value);
                    continue;
                }

                Release(nested);
                nested = document.Fields.Add(
                    insertion,
                    WdFieldType.wdFieldEmpty,
                    segment.Value.Trim(),
                    false);
            }

            try { outer.ShowCodes = false; } catch { }
            var result = outer;
            outer = null;
            return result;
        }
        finally
        {
            Release(nested);
            Release(outerCode);
            Release(insertion);
            Release(outer);
        }
    }

    private static void ConfigureNewMathTypeDisplayEquation(
        Document document,
        InlineShape shape,
        bool numbered,
        string mathTypeNumberPosition,
        bool updateNestedNumberFields = true)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        TabStops? tabs = null;
        TabStop? tab = null;
        Field? placeRef = null;
        Range? shapePrefix = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "A new MathType display equation must occupy one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            EnsureMathTypeNativeStyles(document);
            object displayStyle = "MTDisplayEquation";
            // Clear direct formatting inherited from the source/blank paragraph
            // first, then apply MathType's native paragraph style. Paragraph.Reset
            // can itself restore Normal, so applying MTDisplayEquation before Reset
            // silently lost the native style in real Word.
            try { paragraph.Reset(); } catch { }
            paragraphRange.set_Style(ref displayStyle);
            format = paragraph.Format;
            tabs = format.TabStops;
            var hasCenter = false;
            var hasRight = false;
            for (var index = 1; index <= tabs.Count; index++)
            {
                Release(tab);
                tab = tabs[index];
                hasCenter |= tab.Alignment == WdTabAlignment.wdAlignTabCenter;
                hasRight |= tab.Alignment == WdTabAlignment.wdAlignTabRight;
            }
            if (!hasCenter || !hasRight)
                throw new InvalidOperationException(
                    "MTDisplayEquation does not contain MathType's center/right tab stops.");
            try { paragraphRange.ListFormat.RemoveNumbers(); } catch { }

            // Do not infer the OLE boundary from Paragraph.Range.Text. When Word's
            // field-code view is enabled, an embedded equation is exposed as
            // FIELD-BEGIN + " EMBED Equation..." rather than U+0001 even though
            // InlineShape.Range still identifies the exact OLE story interval.
            // Validate the real characters before that interval instead; this is
            // invariant under Alt+F9 / Field.ShowCodes.
            shapePrefix = document.Range(paragraphRange.Start, shapeRange.Start);
            if (!numbered)
            {
                if (!string.Equals(shapePrefix.Text, "\t", StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "The unnumbered MathType display equation does not begin with Word's native center tab before its OLE object.");
                return;
            }

            var numberOnLeft = string.Equals(
                mathTypeNumberPosition,
                "left",
                StringComparison.OrdinalIgnoreCase);
            var numberOnRight = string.Equals(
                mathTypeNumberPosition,
                "right",
                StringComparison.OrdinalIgnoreCase);
            if (!numberOnLeft && !numberOnRight)
                throw new InvalidDataException(
                    "MathType equation number position must be left or right.");

            placeRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft)
                ?? throw new InvalidOperationException(
                    "The numbered MathType display equation has no MTPlaceRef field on the requested side of its OLE object.");
            Range? placeRefCode = null;
            Range? placeRefResult = null;
            Range? separator = null;
            try
            {
                placeRefCode = placeRef.Code;
                placeRefResult = placeRef.Result;
                var fieldStart = Math.Max(paragraphRange.Start, placeRefCode.Start - 1);
                var fieldEnd = Math.Min(paragraphRange.End, placeRefResult.End + 1);
                if (numberOnLeft)
                {
                    if (fieldEnd > shapeRange.Start)
                        throw new InvalidOperationException(
                            "The MathType left equation number is not positioned before the equation object.");
                    // An empty-result MTPlaceRef can report a field-end boundary
                    // that coincides with the following OLE field. The one story
                    // character immediately before InlineShape.Range is still the
                    // actual separator TAB in both result and field-code views.
                    Release(shapePrefix);
                    shapePrefix = document.Range(
                        Math.Max(paragraphRange.Start, shapeRange.Start - 1),
                        shapeRange.Start);
                    if (!string.Equals(shapePrefix.Text, "\t", StringComparison.Ordinal))
                        throw new InvalidOperationException(
                            "The MathType left-numbered display equation has no tab between its number and equation.");
                }
                else
                {
                    if (fieldStart < shapeRange.End)
                        throw new InvalidOperationException(
                            "The MathType right equation number is not positioned after the equation object.");
                    separator = document.Range(shapeRange.End, fieldStart);
                    if ((separator.Text ?? string.Empty).IndexOf('\t') < 0)
                        throw new InvalidOperationException(
                            "The MathType right-numbered display equation has no tab between its equation and number.");
                    if (!string.Equals(shapePrefix.Text, "\t", StringComparison.Ordinal))
                    {
                        var prefix = shapePrefix.Text ?? string.Empty;
                        var codes = string.Join(
                            ",",
                            prefix.Select(ch => $"U+{(int)ch:X4}"));
                        throw new InvalidOperationException(
                            $"The MathType right-numbered display equation does not begin with Word's native center tab before its OLE object. paragraph={paragraphRange.Start}-{paragraphRange.End}; shape={shapeRange.Start}-{shapeRange.End}; prefix={codes}.");
                    }
                }
                if (updateNestedNumberFields)
                    UpdateNestedMathTypeNumberFields(placeRef);
            }
            finally
            {
                Release(separator);
                Release(placeRefResult);
                Release(placeRefCode);
            }
        }
        finally
        {
            Release(shapePrefix);
            Release(placeRef);
            Release(tab);
            Release(tabs);
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static string StripWordFieldControlCharacters(string text) =>
        (text ?? string.Empty)
            .Replace("\u0013", string.Empty)
            .Replace("\u0014", string.Empty)
            .Replace("\u0015", string.Empty)
            .Replace("\u0001", string.Empty)
            .Replace("\uFFFC", string.Empty)
            .Replace("\r", string.Empty);

    private static MathTypeWordOpenXml.NumberTemplate ReadMathTypePlaceRefTemplate(
        Document document,
        Field source)
    {
        _ = document;
        Range? sourceCode = null;
        try
        {
            sourceCode = source.Code;
            var sourceText = sourceCode.Text ?? string.Empty;
            var macroIndex = sourceText.IndexOf(
                "MTPlaceRef",
                StringComparison.OrdinalIgnoreCase);
            if (macroIndex < 0)
                throw new InvalidDataException(
                    "The source MathType numbering field is not MTPlaceRef.");

            // Word exposes nested fields inside Field.Code using the same field
            // control characters stored in the document: U+0013 begin, optional
            // U+0014 separate, and U+0015 end. Parse that stream directly so
            // literal MathType punctuation such as '(', '.', '-' and ')' is kept
            // byte-for-byte instead of inferred from COM Range coordinates.
            var pattern = sourceText.Substring(
                macroIndex + "MTPlaceRef".Length);
            var template = new MathTypeWordOpenXml.NumberTemplate();
            var literal = new StringBuilder();
            for (var index = 0; index < pattern.Length;)
            {
                if (pattern[index] != '\u0013')
                {
                    literal.Append(pattern[index]);
                    index++;
                    continue;
                }

                if (literal.Length > 0)
                {
                    template.Segments.Add(
                        MathTypeWordOpenXml.NumberSegment.Text(literal.ToString()));
                    literal.Clear();
                }
                var end = pattern.IndexOf('\u0015', index + 1);
                if (end < 0)
                    throw new InvalidDataException(
                        "MathType MTPlaceRef contains an unterminated nested Word field.");
                var fieldBody = pattern.Substring(index + 1, end - index - 1);
                var separate = fieldBody.IndexOf('\u0014');
                if (separate >= 0)
                    fieldBody = fieldBody.Substring(0, separate);
                if (string.IsNullOrWhiteSpace(fieldBody))
                    throw new InvalidDataException(
                        "MathType MTPlaceRef contains an empty nested Word field.");
                template.Segments.Add(
                    MathTypeWordOpenXml.NumberSegment.Field(fieldBody));
                index = end + 1;
            }
            if (literal.Length > 0)
                template.Segments.Add(
                    MathTypeWordOpenXml.NumberSegment.Text(literal.ToString()));
            if (template.Segments.Count == 0)
                throw new InvalidDataException(
                    "MathType MTPlaceRef numbering template has no usable segments.");
            return template;
        }
        finally { Release(sourceCode); }
    }

    private static bool TryReadReusableMathTypePlaceRefTemplate(
        Document document,
        Field source,
        out MathTypeWordOpenXml.NumberTemplate? template)
    {
        template = null;
        Range? visibleNumberRange = null;
        Range? sourceCode = null;
        Fields? nestedFields = null;
        try
        {
            if (!MathTypeEquationReferences.TryGetVisibleNumberRange(
                    document,
                    source,
                    out visibleNumberRange)
                || visibleNumberRange is null)
                return false;
            if (string.IsNullOrWhiteSpace(
                    MathTypeEquationReferences.ReadVisibleNumberText(source)))
                return false;

            sourceCode = source.Code;
            nestedFields = sourceCode.Fields;
            // A native MTPlaceRef always owns at least the hidden MTEqn increment
            // plus the visible MTEqn current-value field. Heading-aware templates
            // additionally own MTChap/MTSec fields. A literal-only outer field such
            // as "(.)" is a detached/corrupt scaffold and is never reusable.
            if (nestedFields.Count < 2) return false;

            var candidate = ReadMathTypePlaceRefTemplate(document, source);
            var hasHiddenEquationIncrement = candidate.Segments.Any(segment =>
                segment.IsField
                && segment.Value.IndexOf("SEQ MTEqn", StringComparison.OrdinalIgnoreCase) >= 0
                && segment.Value.IndexOf("\\h", StringComparison.OrdinalIgnoreCase) >= 0);
            var hasVisibleEquationValue = candidate.Segments.Any(segment =>
                segment.IsField
                && segment.Value.IndexOf("SEQ MTEqn", StringComparison.OrdinalIgnoreCase) >= 0
                && segment.Value.IndexOf("\\c", StringComparison.OrdinalIgnoreCase) >= 0);
            if (!hasHiddenEquationIncrement || !hasVisibleEquationValue)
                return false;

            template = candidate;
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(nestedFields);
            Release(sourceCode);
            Release(visibleNumberRange);
        }
    }

    private static bool MathTypeNumberTemplateUsesHeading(
        MathTypeWordOpenXml.NumberTemplate template) =>
        template.Segments.Any(segment =>
            segment.IsField
            && (segment.Value.IndexOf("SEQ MTChap", StringComparison.OrdinalIgnoreCase) >= 0
                || segment.Value.IndexOf("SEQ MTSec", StringComparison.OrdinalIgnoreCase) >= 0));

    private static Field? FindMathTypePlaceRefFieldForShape(
        Range range,
        Range shapeRange,
        bool numberOnLeft)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? resultRange = null;
        Field? best = null;
        var bestDistance = int.MaxValue;
        try
        {
            fields = range.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(resultRange);
                resultRange = null;
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                resultRange = field.Result;
                var fieldStart = Math.Max(range.Start, code.Start - 1);
                var fieldEnd = Math.Min(range.End, resultRange.End + 1);
                int distance;
                if (numberOnLeft)
                {
                    if (fieldEnd > shapeRange.Start) continue;
                    distance = shapeRange.Start - fieldEnd;
                }
                else
                {
                    if (fieldStart < shapeRange.End) continue;
                    distance = fieldStart - shapeRange.End;
                }
                if (distance >= bestDistance) continue;
                Release(best);
                best = field;
                field = null;
                bestDistance = distance;
            }
            return best;
        }
        finally
        {
            Release(resultRange);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static Field? FindNearestMathTypePlaceRefField(
        Document document,
        int position,
        int excludeStart,
        int excludeEnd)
    {
        Range? content = null;
        Range? localBeforeRange = null;
        Range? localAfterRange = null;
        Range? beforeRange = null;
        Range? afterRange = null;
        Fields? localBeforeFields = null;
        Fields? localAfterFields = null;
        Fields? beforeFields = null;
        Fields? afterFields = null;
        Field? localBefore = null;
        Field? localAfter = null;
        Field? before = null;
        Field? after = null;
        try
        {
            content = document.Content;
            var safePosition = Math.Max(
                content.Start,
                Math.Min(position, content.End));

            // Ordinary insertion is usually adjacent to the previous MathType
            // equation. Search only a bounded character window first so adding the
            // Nth formula does not enumerate every field created by formulas 1..N-1.
            // A document-wide directional fallback remains for a genuinely distant
            // custom MathType numbering template.
            var localStart = Math.Max(content.Start, safePosition - 4096);
            var localEnd = Math.Min(content.End, safePosition + 4096);
            var localBeforeStart = -1;
            var localAfterStart = -1;
            if (safePosition > localStart)
            {
                localBeforeRange = document.Range(localStart, safePosition);
                localBeforeFields = localBeforeRange.Fields;
                localBefore = FindNearestMathTypePlaceRefFieldInCollection(
                    localBeforeFields,
                    safePosition,
                    excludeStart,
                    excludeEnd,
                    reverse: true,
                    stopAfterFirstMatch: true,
                    out localBeforeStart);
            }
            if (safePosition < localEnd)
            {
                localAfterRange = document.Range(safePosition, localEnd);
                localAfterFields = localAfterRange.Fields;
                localAfter = FindNearestMathTypePlaceRefFieldInCollection(
                    localAfterFields,
                    safePosition,
                    excludeStart,
                    excludeEnd,
                    reverse: false,
                    stopAfterFirstMatch: true,
                    out localAfterStart);
            }
            if (localBefore is not null || localAfter is not null)
            {
                if (localAfter is null
                    || localBefore is not null
                    && Math.Abs(localBeforeStart - safePosition)
                        <= Math.Abs(localAfterStart - safePosition))
                {
                    var localResult = localBefore;
                    localBefore = null;
                    return localResult;
                }
                var localAfterResult = localAfter;
                localAfter = null;
                return localAfterResult;
            }

            if (safePosition > content.Start)
            {
                beforeRange = document.Range(content.Start, safePosition);
                beforeFields = beforeRange.Fields;
                before = FindNearestMathTypePlaceRefFieldInCollection(
                    beforeFields,
                    safePosition,
                    excludeStart,
                    excludeEnd,
                    reverse: true,
                    stopAfterFirstMatch: true,
                    out var beforeStart);
                if (before is not null && safePosition - beforeStart == 0)
                {
                    var exact = before;
                    before = null;
                    return exact;
                }
            }

            if (safePosition < content.End)
            {
                afterRange = document.Range(safePosition, content.End);
                afterFields = afterRange.Fields;
                after = FindNearestMathTypePlaceRefFieldInCollection(
                    afterFields,
                    safePosition,
                    excludeStart,
                    excludeEnd,
                    reverse: false,
                    stopAfterFirstMatch: true,
                    out var afterStart);

                if (before is null)
                {
                    var result = after;
                    after = null;
                    return result;
                }
                if (after is null)
                {
                    var result = before;
                    before = null;
                    return result;
                }

                Range? beforeCode = null;
                try
                {
                    beforeCode = before.Code;
                    var beforeStart = Math.Max(content.Start, beforeCode.Start - 1);
                    if (Math.Abs(beforeStart - safePosition)
                        <= Math.Abs(afterStart - safePosition))
                    {
                        var result = before;
                        before = null;
                        return result;
                    }
                    var afterResult = after;
                    after = null;
                    return afterResult;
                }
                finally { Release(beforeCode); }
            }

            var onlyBefore = before;
            before = null;
            return onlyBefore;
        }
        finally
        {
            Release(after);
            Release(before);
            Release(localAfter);
            Release(localBefore);
            Release(afterFields);
            Release(beforeFields);
            Release(localAfterFields);
            Release(localBeforeFields);
            Release(afterRange);
            Release(beforeRange);
            Release(localAfterRange);
            Release(localBeforeRange);
            Release(content);
        }
    }

    private static Field? FindNearestMathTypePlaceRefFieldInCollection(
        Fields fields,
        int position,
        int excludeStart,
        int excludeEnd,
        bool reverse,
        bool stopAfterFirstMatch,
        out int matchedStart)
    {
        Field? field = null;
        Range? code = null;
        Field? best = null;
        var bestDistance = int.MaxValue;
        matchedStart = -1;
        try
        {
            var index = reverse ? fields.Count : 1;
            var end = reverse ? 1 : fields.Count;
            var step = reverse ? -1 : 1;
            for (; reverse ? index >= end : index <= end; index += step)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                var fieldStart = Math.Max(0, code.Start - 1);
                if (fieldStart >= excludeStart && fieldStart < excludeEnd) continue;
                var distance = Math.Abs(fieldStart - position);
                if (distance >= bestDistance) continue;
                Release(best);
                best = field;
                field = null;
                bestDistance = distance;
                matchedStart = fieldStart;
                if (stopAfterFirstMatch) break;
            }
            return best;
        }
        finally
        {
            Release(code);
            Release(field);
        }
    }

    internal static bool HasMathTypeSectionBreak(
        Document document,
        int beforePosition = int.MaxValue)
    {
        Range? content = null;
        Range? searchRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            content = document.Content;
            var searchEnd = beforePosition == int.MaxValue
                ? content.End
                : Math.Max(
                    content.Start,
                    Math.Min(content.End, beforePosition + 1));
            if (searchEnd <= content.Start) return false;
            searchRange = document.Range(content.Start, searchEnd);
            fields = searchRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(searchRange);
            Release(content);
        }
    }

    private static bool HasMathTypeSectionBreakBetween(
        Document document,
        int afterPosition,
        int beforePosition)
    {
        Range? content = null;
        Range? searchRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            content = document.Content;
            var start = Math.Max(
                content.Start,
                Math.Min(Math.Min(afterPosition, beforePosition), content.End));
            var end = Math.Max(
                start,
                Math.Min(Math.Max(afterPosition, beforePosition) + 1, content.End));
            if (end <= start) return false;
            searchRange = document.Range(start, end);
            fields = searchRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(searchRange);
            Release(content);
        }
    }

    internal static int EnsureMathTypeHeadingScopeState(
        Document document,
        int formulaPosition,
        EquationNumberFormat format) =>
        EnsureMathTypeHeadingScopeState(
            document,
            formulaPosition,
            format,
            out _);

    private static int EnsureMathTypeHeadingScopeState(
        Document document,
        int formulaPosition,
        EquationNumberFormat format,
        out int createdSectionBreakCodeStart,
        ResolvedEquationHeadingScope? preResolvedHeadingScope = null)
    {
        createdSectionBreakCodeStart = -1;
        if (!format.UsesHeading) return 0;

        var resolvedScope = preResolvedHeadingScope;
        var scope = resolvedScope is null
            ? WordEquationNumbering.ResolveHeadingScopeAtPosition(
                document,
                formulaPosition,
                format.Id)
            : (
                resolvedScope.ScopeStart,
                resolvedScope.ScopeEnd,
                resolvedScope.NumberText);
        // No real Heading paragraph exists before this equation. VisualTeX's
        // heading-aware numbering intentionally uses a zero prefix in that scope
        // (0.1 / 0.0-1). Do not manufacture MathType chapter 1 state.
        if (scope.ScopeStart == int.MinValue || scope.ScopeEnd == int.MinValue)
            return 0;
        if (scope.ScopeEnd > formulaPosition)
            return 0;

        var parts = (scope.NumberText ?? string.Empty)
            .Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0
            || !int.TryParse(
                parts[0],
                System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture,
                out var chapter))
            return 0;
        var section = 0;
        if (format.HeadingLevel >= 2
            && parts.Length >= 2
            && !int.TryParse(
                parts[1],
                System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture,
                out section))
            return 0;
        if (chapter == 0 && section == 0)
            return 0;

        // Scope state belongs immediately after the real Word heading, not before
        // whichever equation happens to be converted first. This is crucial for
        // descending whole-document conversion: every equation in the same scope
        // must remain after the same native MathType section state.
        if (HasMathTypeSectionBreakBetween(document, scope.ScopeEnd, formulaPosition))
            return 0;

        WordDoubleClickHook.TraceMessage(
            $"mathtype-heading-scope-state chapter={chapter} section={section} heading={scope.ScopeStart}:{scope.ScopeEnd} formula={formulaPosition}");
        var insertedLength = InsertMathTypeSectionBreakState(
            document,
            scope.ScopeEnd,
            chapter,
            section);
        if (insertedLength <= 0) return insertedLength;

        createdSectionBreakCodeStart = FindMathTypeSectionBreakCodeStartBetween(
            document,
            scope.ScopeEnd,
            Math.Min(document.Content.End, formulaPosition + insertedLength));
        if (createdSectionBreakCodeStart < 0)
            throw new InvalidOperationException(
                "VisualTeX inserted MathType heading state but could not identify the new section field for transactional rollback.");
        return insertedLength;
    }

    private static int FindMathTypeSectionBreakCodeStartBetween(
        Document document,
        int afterPosition,
        int beforePosition)
    {
        Range? content = null;
        Range? searchRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        var best = int.MaxValue;
        try
        {
            content = document.Content;
            var start = Math.Max(
                content.Start,
                Math.Min(Math.Min(afterPosition, beforePosition), content.End));
            var end = Math.Max(
                start,
                Math.Min(Math.Max(afterPosition, beforePosition) + 1, content.End));
            if (end <= start) return -1;
            searchRange = document.Range(start, end);
            fields = searchRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                best = Math.Min(best, code.Start);
            }
            return best == int.MaxValue ? -1 : best;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(searchRange);
            Release(content);
        }
    }

    private static void InsertIsolatedMathTypeSectionParagraph(
        Document targetDocument,
        int insertionPosition,
        string sectionFlatOpc)
    {
        Microsoft.Office.Interop.Word.Application? application = null;
        Documents? documents = null;
        Document? stagingDocument = null;
        Range? stagingInsertion = null;
        Field? stagingSection = null;
        Range? stagingCode = null;
        Range? stagingResult = null;
        Range? stagingFull = null;
        Paragraphs? stagingParagraphs = null;
        Paragraph? stagingParagraph = null;
        Range? stagingParagraphRange = null;
        Range? targetBreak = null;
        Range? targetProbe = null;
        Paragraphs? targetParagraphs = null;
        Paragraph? targetParagraph = null;
        Range? targetParagraphRange = null;
        try
        {
            application = targetDocument.Application;
            documents = application.Documents;
            stagingDocument = documents.Add(Visible: false);
            stagingInsertion = stagingDocument.Range(
                stagingDocument.Content.Start,
                stagingDocument.Content.Start);
            stagingInsertion.InsertXML(sectionFlatOpc);
            stagingSection = FindFirstMathTypeSectionBreakField(stagingDocument)
                ?? throw new InvalidOperationException(
                    "The isolated MathType section-state document did not materialize MTEditEquationSection2.");
            stagingCode = stagingSection.Code;
            stagingResult = stagingSection.Result;
            stagingFull = stagingDocument.Range(
                Math.Max(stagingDocument.Content.Start, stagingCode.Start - 1),
                Math.Min(stagingDocument.Content.End, stagingResult.End + 1));
            stagingParagraphs = stagingFull.Paragraphs;
            if (stagingParagraphs.Count != 1)
                throw new InvalidOperationException(
                    "The isolated MathType section-state field spans multiple staging paragraphs.");
            stagingParagraph = stagingParagraphs[1];
            stagingParagraphRange = stagingParagraph.Range;
            if (stagingParagraphRange.InlineShapes.Count != 0)
                throw new InvalidOperationException(
                    "The isolated MathType section-state staging paragraph unexpectedly contains an OLE object.");

            targetBreak = targetDocument.Range(insertionPosition, insertionPosition);
            targetBreak.Text = "\r";
            targetProbe = targetDocument.Range(
                insertionPosition,
                Math.Min(targetDocument.Content.End, insertionPosition + 1));
            targetParagraphs = targetProbe.Paragraphs;
            if (targetParagraphs.Count != 1)
                throw new InvalidOperationException(
                    "Word did not create a dedicated target paragraph for MathType section state.");
            targetParagraph = targetParagraphs[1];
            targetParagraphRange = targetParagraph.Range;
            targetParagraphRange.FormattedText = stagingParagraphRange.FormattedText;
        }
        finally
        {
            Release(targetParagraphRange);
            Release(targetParagraph);
            Release(targetParagraphs);
            Release(targetProbe);
            Release(targetBreak);
            Release(stagingParagraphRange);
            Release(stagingParagraph);
            Release(stagingParagraphs);
            Release(stagingFull);
            Release(stagingResult);
            Release(stagingCode);
            Release(stagingSection);
            Release(stagingInsertion);
            if (stagingDocument is not null)
            {
                try { stagingDocument.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(stagingDocument);
            Release(documents);
            Release(application);
        }
    }

    internal static int EnsureDefaultMathTypeSectionBreak(
        Document document,
        int beforePosition = int.MaxValue)
    {
        if (HasMathTypeSectionBreak(document, beforePosition)) return 0;
        return InsertMathTypeSectionBreakState(
            document,
            beforePosition,
            chapter: 1,
            section: 1);
    }

    private static int InsertMathTypeSectionBreakState(
        Document document,
        int beforePosition,
        int chapter,
        int section)
    {
        EnsureMathTypeNativeStyles(document);

        var contentEndBefore = document.Content.End;
        var paragraphCountBefore = ReadDocumentParagraphCount(document);
        var label = string.Equals(
            System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName,
            "zh",
            StringComparison.OrdinalIgnoreCase)
            ? $"公式章 {chapter} 节 {section}"
            : $"Equation Chapter {chapter} Section {section}";
        var breakXml = MathTypeWordOpenXml.CreateSectionBreakFlatOpc(
            label,
            chapter,
            section);
        Range? insertion = null;
        Field? sectionBreak = null;
        Range? sectionCode = null;
        Range? sectionResult = null;
        Range? sectionFull = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        Range? sectionParagraphSplit = null;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = document.Content.End;
            var insertionPosition = beforePosition == int.MaxValue
                ? contentStart
                : Math.Max(
                    contentStart,
                    Math.Min(beforePosition, Math.Max(contentStart, contentEnd - 1)));
            insertion = document.Range(insertionPosition, insertionPosition);
            // Always anchor the hidden section state at the containing paragraph
            // boundary. A shape/bookmark Start inside an MTDisplayEquation row can
            // otherwise place MTEditEquationSection2 after the row's leading tab.
            paragraphs = insertion.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
                insertionPosition = paragraphRange.Start;
                insertion.SetRange(insertionPosition, insertionPosition);
            }
            Release(paragraphRange);
            paragraphRange = null;
            Release(paragraph);
            paragraph = null;
            Release(paragraphs);
            paragraphs = null;

            // Target-document InsertXML can merge this complete <w:p> field tree
            // back into the following MTDisplayEquation paragraph. Materialize it
            // in an isolated hidden Word document first, then transfer the single
            // finished paragraph with FormattedText into a reserved target row.
            InsertIsolatedMathTypeSectionParagraph(
                document,
                insertionPosition,
                breakXml);
            sectionBreak = FindFirstMathTypeSectionBreakField(document)
                ?? throw new InvalidOperationException(
                    "Word did not materialize MathType's MTEditEquationSection2 field.");
            sectionCode = sectionBreak.Code;
            sectionResult = sectionBreak.Result;
            sectionFull = document.Range(
                Math.Max(document.Content.Start, sectionCode.Start - 1),
                Math.Min(document.Content.End, sectionResult.End + 1));

            // Do not infer isolation from the total paragraph count. Word can add
            // a paragraph while still leaving MTEditEquationSection2 and the first
            // Equation.DSMT4 in the same row. Inspect the field's actual paragraph
            // and split immediately after the outer field whenever anything else
            // remains before that paragraph mark.
            paragraphs = sectionFull.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "MathType chapter/section break materialized across multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var paragraphContentEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            if (sectionFull.End < paragraphContentEnd)
            {
                sectionParagraphSplit = document.Range(sectionFull.End, sectionFull.End);
                sectionParagraphSplit.Text = "\r";
            }

            Release(paragraphRange);
            paragraphRange = null;
            Release(paragraph);
            paragraph = null;
            Release(paragraphs);
            paragraphs = null;

            Release(sectionFull);
            sectionFull = null;
            Release(sectionResult);
            sectionResult = null;
            Release(sectionCode);
            sectionCode = null;
            sectionCode = sectionBreak.Code;
            sectionResult = sectionBreak.Result;
            sectionFull = document.Range(
                Math.Max(document.Content.Start, sectionCode.Start - 1),
                Math.Min(document.Content.End, sectionResult.End + 1));
            paragraphs = sectionFull.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "MathType chapter/section break materialized across multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.InlineShapes.Count != 0)
                throw new InvalidOperationException(
                    "MathType chapter/section state still shares a paragraph with an OLE equation.");
            for (var fieldIndex = 1; fieldIndex <= paragraphRange.Fields.Count; fieldIndex++)
            {
                Field? paragraphField = null;
                Range? paragraphFieldCode = null;
                try
                {
                    paragraphField = paragraphRange.Fields[fieldIndex];
                    paragraphFieldCode = paragraphField.Code;
                    var fieldCodeText = paragraphFieldCode.Text ?? string.Empty;
                    if (fieldCodeText.IndexOf("MTPlaceRef", StringComparison.OrdinalIgnoreCase) >= 0
                        || fieldCodeText.IndexOf("EMBED ", StringComparison.OrdinalIgnoreCase) >= 0)
                        throw new InvalidOperationException(
                            "MathType chapter/section state still shares a paragraph with equation-owned fields.");
                }
                finally
                {
                    Release(paragraphFieldCode);
                    Release(paragraphField);
                }
            }
            paragraphMark = document.Range(
                Math.Max(paragraphRange.Start, paragraphRange.End - 1),
                paragraphRange.End);
            if (!string.Equals(paragraphMark.Text, "\r", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "MathType chapter/section break has no isolated paragraph mark.");

            // MTEquationSection is a hidden red *character* style.  The paragraph
            // mark must stay with the hidden field.  Deleting that mark merges the
            // hidden character formatting into the following user paragraph, so
            // the next field/text insertion inherits MathType's red state.
            object sectionStyle = "MTEquationSection";
            paragraphRange.set_Style(ref sectionStyle);
            UpdateNestedMathTypeNumberFields(sectionBreak);
            return document.Content.End - contentEndBefore;
        }
        finally
        {
            Release(sectionParagraphSplit);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(sectionFull);
            Release(sectionResult);
            Release(sectionCode);
            Release(sectionBreak);
            Release(insertion);
        }
    }

    private static Field? FindFirstMathTypeSectionBreakField(Document document)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Field? result = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                result = field;
                field = null;
                break;
            }
            return result;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    internal static void RemoveAllMathTypeSectionBreakFields(Document document)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        var starts = new List<int>();
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    starts.Add(code.Start);
            }
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }

        foreach (var start in starts.OrderByDescending(value => value))
            RemoveMathTypeSectionBreakFieldAtCodeStart(document, start);
    }

    private static void RemoveMathTypeSectionBreakFieldAtCodeStart(
        Document document,
        int codeStart)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        Range? full = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (code.Start != codeStart
                    || (code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                paragraphs = code.Paragraphs;
                if (paragraphs.Count == 1)
                {
                    paragraph = paragraphs[1];
                    paragraphRange = paragraph.Range.Duplicate;
                }

                // Delete the exact section field created by this operation. Never
                // delete "the first" MathType section field: later-chapter failures
                // must not destroy an earlier chapter's valid native state.
                try { field.Delete(); }
                catch
                {
                    result = field.Result;
                    var start = Math.Max(document.Content.Start, code.Start - 1);
                    var end = Math.Min(
                        document.Content.End,
                        Math.Max(code.End, result.End) + 1);
                    full = document.Range(start, end);
                    full.Delete();
                }

                if (paragraphRange is not null
                    && paragraphRange.InlineShapes.Count == 0)
                {
                    try
                    {
                        object normalStyle = WdBuiltinStyle.wdStyleNormal;
                        paragraphRange.set_Style(ref normalStyle);
                        paragraphRange.Font.Hidden = 0;
                        paragraphRange.Delete();
                    }
                    catch { }
                }
                return;
            }
        }
        catch { }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(full);
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static void EnsureMathTypeNativeStyles(Document document)
    {
        Styles? styles = null;
        Style? displayStyle = null;
        Style? sectionStyle = null;
        ParagraphFormat? displayFormat = null;
        TabStops? displayTabs = null;
        TabStop? tab = null;
        PageSetup? pageSetup = null;
        Microsoft.Office.Interop.Word.Font? sectionFont = null;
        try
        {
            styles = document.Styles;
            object displayName = "MTDisplayEquation";
            try { displayStyle = styles.get_Item(ref displayName); }
            catch
            {
                object paragraphType = WdStyleType.wdStyleTypeParagraph;
                displayStyle = styles.Add("MTDisplayEquation", ref paragraphType);
                // Match MathType's own behavior: create the style only when the
                // document does not already have it. Never overwrite an existing
                // native MathType style, because its tab geometry may have been
                // adapted by MathType for the current document/template.
                displayFormat = displayStyle.ParagraphFormat;
                displayFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                displayFormat.SpaceBefore = 0;
                displayFormat.SpaceAfter = 0;
                displayFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                displayTabs = displayFormat.TabStops;
                displayTabs.ClearAll();
                pageSetup = document.PageSetup;
                var usableWidth = Math.Max(
                    72f,
                    pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin);
                tab = displayTabs.Add(
                    usableWidth / 2f,
                    WdTabAlignment.wdAlignTabCenter,
                    WdTabLeader.wdTabLeaderSpaces);
                Release(tab);
                tab = displayTabs.Add(
                    usableWidth,
                    WdTabAlignment.wdAlignTabRight,
                    WdTabLeader.wdTabLeaderSpaces);
                Release(tab);
                tab = null;
            }

            object sectionName = "MTEquationSection";
            try { sectionStyle = styles.get_Item(ref sectionName); }
            catch
            {
                object characterType = WdStyleType.wdStyleTypeCharacter;
                sectionStyle = styles.Add("MTEquationSection", ref characterType);
            }
            sectionFont = sectionStyle.Font;
            sectionFont.Hidden = -1;
            sectionFont.Color = WdColor.wdColorRed;
        }
        finally
        {
            Release(sectionFont);
            Release(pageSetup);
            Release(tab);
            Release(displayTabs);
            Release(displayFormat);
            Release(sectionStyle);
            Release(displayStyle);
            Release(styles);
        }
    }

    private static void NormalizeMathTypeDisplayParagraphFormat(
        Document document,
        Paragraph paragraph)
    {
        ParagraphFormat? format = null;
        TabStops? tabs = null;
        TabStop? tab = null;
        PageSetup? pageSetup = null;
        try
        {
            format = paragraph.Format;
            format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            format.LeftIndent = 0;
            format.RightIndent = 0;
            format.FirstLineIndent = 0;
            format.SpaceBefore = 0;
            format.SpaceAfter = 0;
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            tabs = format.TabStops;
            tabs.ClearAll();
            pageSetup = document.PageSetup;
            var usableWidth = Math.Max(
                72f,
                pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin);
            tab = tabs.Add(
                usableWidth / 2f,
                WdTabAlignment.wdAlignTabCenter,
                WdTabLeader.wdTabLeaderSpaces);
            Release(tab);
            tab = tabs.Add(
                usableWidth,
                WdTabAlignment.wdAlignTabRight,
                WdTabLeader.wdTabLeaderSpaces);
        }
        finally
        {
            Release(pageSetup);
            Release(tab);
            Release(tabs);
            Release(format);
        }
    }

    private static string ReadMathTypeNumberPositionPreference(Document document)
    {
        object? propertiesObject = null;
        object? propertyObject = null;
        try
        {
            propertiesObject = document.CustomDocumentProperties;
            if (propertiesObject is null) return "right";
            dynamic properties = propertiesObject;
            try
            {
                propertyObject = properties["MTEqnNumsOnRight"];
                dynamic property = propertyObject;
                var value = property.Value;
                if (value is bool right) return right ? "right" : "left";
                if (value is int integer) return integer != 0 ? "right" : "left";
                var text = Convert.ToString((object)value);
                if (bool.TryParse(text, out bool parsed))
                    return parsed ? "right" : "left";
            }
            catch { }
            return "right";
        }
        finally
        {
            Release(propertyObject);
            Release(propertiesObject);
        }
    }

    private static void UpdateNestedMathTypeNumberFields(Field outer)
    {
        Range? code = null;
        Fields? fields = null;
        Field? field = null;
        try
        {
            code = outer.Code;
            fields = code.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(field);
                field = fields[index];
                try { field.Update(); } catch { }
            }
            try { outer.Update(); } catch { }
        }
        finally
        {
            Release(field);
            Release(fields);
            Release(code);
        }
    }

    private static MathTypeDisplayParagraphLayout? CaptureMathTypeDisplayParagraphLayout(
        InlineShape shape)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            return CaptureMathTypeDisplayParagraphLayout(paragraph);
        }
        catch { return null; }
        finally
        {
            Release(tabStop);
            Release(tabStops);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static MathTypeDisplayParagraphLayout? CaptureMathTypeDisplayParagraphLayout(
        Paragraph paragraph)
    {
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        try
        {
            format = paragraph.Format;
            var layout = new MathTypeDisplayParagraphLayout
            {
                Alignment = format.Alignment,
                LeftIndent = format.LeftIndent,
                RightIndent = format.RightIndent,
                FirstLineIndent = format.FirstLineIndent,
                SpaceBefore = format.SpaceBefore,
                SpaceAfter = format.SpaceAfter,
                LineSpacingRule = format.LineSpacingRule,
                LineSpacing = format.LineSpacing,
                KeepTogether = format.KeepTogether,
                KeepWithNext = format.KeepWithNext,
                WidowControl = format.WidowControl,
                PageBreakBefore = format.PageBreakBefore,
            };
            tabStops = format.TabStops;
            for (var index = 1; index <= tabStops.Count; index++)
            {
                Release(tabStop);
                tabStop = tabStops[index];
                if (tabStop.Alignment == WdTabAlignment.wdAlignTabLeft) continue;
                layout.SpecialTabStops.Add((tabStop.Position, tabStop.Alignment, tabStop.Leader));
            }
            return layout;
        }
        catch { return null; }
        finally
        {
            Release(tabStop);
            Release(tabStops);
            Release(format);
        }
    }

    private static int ReadDocumentParagraphCount(Document document)
    {
        Paragraphs? paragraphs = null;
        try
        {
            paragraphs = document.Paragraphs;
            return paragraphs.Count;
        }
        finally { Release(paragraphs); }
    }

    private static void RepairMathTypeInsertXmlParagraphSplit(
        Document document,
        InlineShape shape,
        int sourceParagraphCount)
    {
        var currentParagraphCount = ReadDocumentParagraphCount(document);
        if (currentParagraphCount == sourceParagraphCount) return;
        if (currentParagraphCount != sourceParagraphCount + 1)
            throw new InvalidOperationException(
                $"Word changed the paragraph count unexpectedly while materializing a MathType OLE: before={sourceParagraphCount}, after={currentParagraphCount}.");

        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The materialized MathType OLE spans multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.End <= paragraphRange.Start)
                throw new InvalidOperationException(
                    "The materialized MathType OLE paragraph has no paragraph mark to repair.");
            paragraphMark = document.Range(paragraphRange.End - 1, paragraphRange.End);
            if (!string.Equals(paragraphMark.Text, "\r", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "Word inserted a MathType OLE paragraph boundary that VisualTeX could not identify safely.");
            paragraphMark.Delete();
        }
        finally
        {
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }

        var repairedParagraphCount = ReadDocumentParagraphCount(document);
        if (repairedParagraphCount != sourceParagraphCount)
            throw new InvalidOperationException(
                $"VisualTeX could not restore the original MathType OLE paragraph structure: expected={sourceParagraphCount}, actual={repairedParagraphCount}.");
    }

    private static MathTypeDisplayParagraphLayout? RepairDetachedMathTypeNumberParagraph(
        Document document,
        InlineShape shape)
    {
        Range? shapeRange = null;
        Paragraphs? shapeParagraphs = null;
        Paragraph? shapeParagraph = null;
        Range? shapeParagraphRange = null;
        Range? nextProbe = null;
        Paragraphs? nextParagraphs = null;
        Paragraph? nextParagraph = null;
        Range? nextRange = null;
        Range? paragraphMark = null;
        try
        {
            shapeRange = shape.Range;
            shapeParagraphs = shapeRange.Paragraphs;
            if (shapeParagraphs.Count != 1) return null;
            shapeParagraph = shapeParagraphs[1];
            shapeParagraphRange = shapeParagraph.Range;
            if (shapeParagraphRange.End >= document.Content.End) return null;

            // Probe the first character after the formula paragraph mark.  A
            // non-empty one-character range reliably belongs to the following
            // paragraph, unlike a collapsed boundary range which Word can resolve
            // to either side depending on field state.
            nextProbe = document.Range(
                shapeParagraphRange.End,
                Math.Min(document.Content.End, shapeParagraphRange.End + 1));
            nextParagraphs = nextProbe.Paragraphs;
            if (nextParagraphs.Count < 1) return null;
            nextParagraph = nextParagraphs[1];
            nextRange = nextParagraph.Range;
            if (!IsDetachedMathTypeNumberParagraph(nextRange)) return null;

            var numberingLayout = CaptureMathTypeDisplayParagraphLayout(nextParagraph);
            paragraphMark = document.Range(
                shapeParagraphRange.End - 1,
                shapeParagraphRange.End);
            if (!string.Equals(paragraphMark.Text, "\r", StringComparison.Ordinal))
                return null;
            paragraphMark.Delete();
            return numberingLayout;
        }
        catch
        {
            // This is compatibility recovery for documents already damaged by an
            // older VisualTeX build.  Never make an otherwise valid MathType edit
            // fail merely because an adjacent paragraph only resembles MTPlaceRef.
            return null;
        }
        finally
        {
            Release(paragraphMark);
            Release(nextRange);
            Release(nextParagraph);
            Release(nextParagraphs);
            Release(nextProbe);
            Release(shapeParagraphRange);
            Release(shapeParagraph);
            Release(shapeParagraphs);
            Release(shapeRange);
        }
    }

    private static bool IsDetachedMathTypeNumberParagraph(Range range)
    {
        InlineShapes? inlineShapes = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            inlineShapes = range.InlineShapes;
            if (inlineShapes.Count != 0) return false;

            fields = range.Fields;
            var hasPlaceRef = false;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    hasPlaceRef = true;
            }
            if (!hasPlaceRef) return false;

            var text = (range.Text ?? string.Empty)
                .Replace("\r", string.Empty)
                .Replace("\a", string.Empty)
                .Replace("\t", string.Empty)
                .Replace("\u0013", string.Empty)
                .Replace("\u0014", string.Empty)
                .Replace("\u0015", string.Empty)
                .Trim();
            if (text.Length == 0) return true;

            var sawNumber = false;
            foreach (var character in text)
            {
                if (char.IsWhiteSpace(character)) continue;
                if (char.IsDigit(character))
                {
                    sawNumber = true;
                    continue;
                }
                if ("()[]{}.,:;-–—/\\".IndexOf(character) >= 0) continue;
                return false;
            }
            return sawNumber;
        }
        catch { return false; }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(inlineShapes);
        }
    }

    private static void RestoreMathTypeDisplayParagraphLayout(
        InlineShape shape,
        MathTypeDisplayParagraphLayout layout)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            format = paragraph.Format;
            format.Alignment = layout.Alignment;
            format.LeftIndent = layout.LeftIndent;
            format.RightIndent = layout.RightIndent;
            format.FirstLineIndent = layout.FirstLineIndent;
            format.SpaceBefore = layout.SpaceBefore;
            format.SpaceAfter = layout.SpaceAfter;
            format.LineSpacingRule = layout.LineSpacingRule;
            try { format.LineSpacing = layout.LineSpacing; } catch { }
            format.KeepTogether = layout.KeepTogether;
            format.KeepWithNext = layout.KeepWithNext;
            format.WidowControl = layout.WidowControl;
            format.PageBreakBefore = layout.PageBreakBefore;

            if (layout.SpecialTabStops.Count == 0) return;
            tabStops = format.TabStops;
            foreach (var special in layout.SpecialTabStops)
            {
                var exists = false;
                for (var index = 1; index <= tabStops.Count; index++)
                {
                    Release(tabStop);
                    tabStop = tabStops[index];
                    if (Math.Abs(tabStop.Position - special.Position) <= 0.5f
                        && tabStop.Alignment == special.Alignment
                        && tabStop.Leader == special.Leader)
                    {
                        exists = true;
                        break;
                    }
                }
                if (exists) continue;
                Release(tabStop);
                tabStop = tabStops.Add(special.Position, special.Alignment, special.Leader);
            }
        }
        finally
        {
            Release(tabStop);
            Release(tabStops);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static byte[] ReadMathTypeMtefFromCompoundFile(byte[] compoundFile)
    {
        var equationNative = MathTypeOleStorage.ReadEquationNative(compoundFile);
        if (equationNative.Length < 12)
            throw new InvalidDataException("MathType Equation Native is too short for MTEF extraction.");
        var headerLength = BitConverter.ToUInt16(equationNative, 0);
        var mtefLength = checked((int)BitConverter.ToUInt32(equationNative, 8));
        if (headerLength < 12
            || mtefLength <= 0
            || headerLength + mtefLength > equationNative.Length)
            throw new InvalidDataException("MathType Equation Native contains an invalid MTEF extent.");
        var mtef = new byte[mtefLength];
        Buffer.BlockCopy(equationNative, headerLength, mtef, 0, mtefLength);
        return mtef;
    }

    private static bool TryRenderMathTypeNativePreviewFromCompoundFile(
        byte[] compoundFile,
        string outputDirectory,
        out MathTypeNativePreviewRenderer.Result? result)
    {
        result = null;
        try
        {
            var mtef = ReadMathTypeMtefFromCompoundFile(compoundFile);
            if (!MathTypeNativePreviewRenderer.TryRender(
                    mtef,
                    outputDirectory,
                    out var rendered))
                return false;
            result = rendered;
            return true;
        }
        catch { return false; }
    }

    private static float CalculateMathTypeNativePresentationScale(
        float sourceWordExtent,
        float? sourceNativeExtent)
    {
        if (!(sourceWordExtent > 0) || sourceNativeExtent is not > 0) return 1f;
        var scale = sourceWordExtent / sourceNativeExtent.Value;
        if (!(scale > 0) || float.IsNaN(scale) || float.IsInfinity(scale)) return 1f;
        return Math.Max(0.25f, Math.Min(4f, scale));
    }

    private static int ReadInlineOleWordPosition(InlineShape shape)
    {
        Range? shapeRange = null;
        Range? probe = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        Document? document = null;
        try
        {
            shapeRange = shape.Range;
            document = shapeRange.Document;
            // Word exposes an embedded MathType object as an EMBED field.  The
            // InlineShape.Range therefore spans the field instruction plus the
            // single U+0001 object/result character.  Genuine MathType keeps the
            // field instruction at Position=0 and applies the vertical offset only
            // to U+0001.  Reading range.Start therefore returns the wrong baseline.
            for (var position = shapeRange.Start; position < shapeRange.End; position++)
            {
                Release(font);
                font = null;
                Release(probe);
                probe = document.Range(position, position + 1);
                if (!string.Equals(probe.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                font = probe.Font;
                var wordPosition = font.Position;
                if (wordPosition != (int)WdConstants.wdUndefined
                    && wordPosition >= -256
                    && wordPosition <= 256)
                    return wordPosition;
            }
        }
        catch { }
        finally
        {
            Release(font);
            Release(probe);
            Release(shapeRange);
            Release(document);
        }
        return (int)Math.Round(ReadDefinedShapeFontPosition(shape) ?? 0f);
    }

    private static int CalculateMathTypeOleWordPosition(
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline)
    {
        if (!(actualHeightPoints > 0)
            || !(exportedHeight > 0)
            || !exportedBaseline.HasValue
            || float.IsNaN(exportedBaseline.Value)
            || float.IsInfinity(exportedBaseline.Value)
            || exportedBaseline.Value < 0
            || exportedBaseline.Value >= exportedHeight)
            return 0;

        // MathType's Word integration stores the OLE object's character position
        // from the equation baseline toward the bottom of the picture. Keep this
        // calculation local to MathType instead of sharing VisualTeX's ordinary
        // inline-OLE optical alignment code; changes to VisualTeX inline layout must
        // never alter Equation.DSMT4 placement again.
        var baselineFromBottomPoints =
            actualHeightPoints * (exportedHeight - exportedBaseline.Value) / exportedHeight;
        var rounded = Math.Max(
            0,
            (int)Math.Round(baselineFromBottomPoints, MidpointRounding.AwayFromZero));

        // The offline WMF presentation includes Word's one-point character-box
        // allowance. Native MathType samples (ordinary fractions, radicals and
        // mixed inline equations) place the U+0001 object one point above the raw
        // picture descent, while the WMF/dxaOrig dimensions retain the full bound.
        return -Math.Max(0, rounded - 1);
    }

    private static void SetInlineOleWordPosition(InlineShape shape, int position)
    {
        Range? shapeRange = null;
        Range? probe = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        Document? document = null;
        try
        {
            shapeRange = shape.Range;
            document = shapeRange.Document;
            // Match genuine MathType field formatting exactly: EMBED instruction
            // characters remain on the paragraph baseline; only the U+0001 object
            // result receives the equation's vertical offset.  Applying Position to
            // shape.Range shifts the whole field and can leak into following prose.
            font = shapeRange.Font;
            font.Position = 0;
            Release(font);
            font = null;

            var clamped = Math.Max(-256, Math.Min(256, position));
            for (var index = shapeRange.Start; index < shapeRange.End; index++)
            {
                Release(probe);
                probe = document.Range(index, index + 1);
                if (!string.Equals(probe.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                font = probe.Font;
                font.Position = clamped;
                return;
            }
        }
        finally
        {
            Release(font);
            Release(probe);
            Release(shapeRange);
            Release(document);
        }
    }

    private static InlineOlePreviewMetrics? TryMeasureInlineOlePreview(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            var bytes = range.EnhMetaFileBits as byte[];
            if (bytes is null || bytes.Length == 0) return null;
            using var stream = new System.IO.MemoryStream(bytes, writable: false);
            using var metafile = new System.Drawing.Imaging.Metafile(stream);
            return MeasureMetafilePreview(metafile);
        }
        catch { return null; }
        finally { Release(range); }
    }

    private static InlineOlePreviewMetrics? TryMeasureMetafilePreview(string emfPath)
    {
        try
        {
            using var metafile = new System.Drawing.Imaging.Metafile(emfPath);
            return MeasureMetafilePreview(metafile);
        }
        catch { return null; }
    }

    private static InlineOlePreviewMetrics? MeasureMetafilePreview(
        System.Drawing.Imaging.Metafile metafile)
    {
        const int width = 640;
        const int height = 240;
        using var bitmap = new Bitmap(width, height);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            graphics.DrawImage(metafile, 0, 0, width, height);
        }

        var minY = height;
        var maxY = -1;
        for (var y = 0; y < height; y++)
        {
            for (var x = 0; x < width; x++)
            {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.R >= 245 && pixel.G >= 245 && pixel.B >= 245) continue;
                minY = Math.Min(minY, y);
                maxY = Math.Max(maxY, y);
            }
        }
        if (maxY < minY) return null;

        var inkHeightRatio = (maxY - minY + 1f) / height;
        var bottomWhitespaceRatio = (height - 1f - maxY) / height;
        if (!(inkHeightRatio > 0.01f)
            || float.IsNaN(inkHeightRatio)
            || float.IsInfinity(inkHeightRatio))
            return null;
        return new InlineOlePreviewMetrics(inkHeightRatio, bottomWhitespaceRatio);
    }

    private static (float Width, float Height) CalculateMathTypeEditedPresentationSize(
        float oldWidth,
        float oldHeight,
        InlineOlePreviewMetrics? sourcePreview,
        InlineOlePreviewMetrics? editedPreview,
        float? newRenderWidth,
        float? newRenderHeight,
        double? originalRenderWidth,
        double? originalRenderHeight,
        double? originalFontSizePt,
        double? originalRenderFontSizePt)
    {
        var fallback = OfficeFormulaSizing.EditedSize(
            oldWidth,
            oldHeight,
            originalRenderWidth,
            originalRenderHeight,
            newRenderWidth ?? oldWidth / 0.75f,
            newRenderHeight ?? oldHeight / 0.75f,
            originalFontSizePt: originalFontSizePt,
            originalRenderFontSizePt: originalRenderFontSizePt);
        if (!sourcePreview.HasValue || !editedPreview.HasValue
            || !(oldHeight > 0)
            || !(sourcePreview.Value.InkHeightRatio > 0.01f)
            || !(editedPreview.Value.InkHeightRatio > 0.01f))
            return fallback;

        var height = oldHeight
            * sourcePreview.Value.InkHeightRatio
            / editedPreview.Value.InkHeightRatio;
        if (!(height > 0) || float.IsNaN(height) || float.IsInfinity(height))
            return fallback;

        var aspect = newRenderWidth is > 0 && newRenderHeight is > 0
            ? newRenderWidth.Value / newRenderHeight.Value
            : fallback.Width / Math.Max(0.01f, fallback.Height);
        if (!(aspect > 0) || float.IsNaN(aspect) || float.IsInfinity(aspect))
            return fallback;

        // Preserve the native MathType glyph scale, not the outer OLE box. A
        // MathType preview typically contains appreciable ascent/descent padding,
        // whereas VisualTeX's EMF is tightly cropped. Keeping the same outer
        // height would therefore enlarge the visible glyphs by 20–40 percent.
        height = Math.Max(1f, Math.Min(oldHeight * 4f, Math.Max(oldHeight * 0.25f, height)));
        return (Math.Max(1f, height * aspect), height);
    }

    private static int CalculateMathTypeInlineWordPosition(
        int sourceWordPosition,
        float sourceHeight,
        float targetHeight,
        InlineOlePreviewMetrics? sourcePreview,
        InlineOlePreviewMetrics? targetPreview)
    {
        if (!sourcePreview.HasValue || !targetPreview.HasValue)
            return sourceWordPosition;

        var sourceBottomWhitespace = sourceHeight * sourcePreview.Value.BottomWhitespaceRatio;
        var targetBottomWhitespace = targetHeight * targetPreview.Value.BottomWhitespaceRatio;
        var correction = (int)Math.Round(
            sourceBottomWhitespace - targetBottomWhitespace,
            MidpointRounding.AwayFromZero);
        return Math.Max(-256, Math.Min(256, sourceWordPosition + correction));
    }

    private static void ResetShapeFontPosition(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            ResetRangeFontPosition(range);
        }
        finally { Release(range); }
    }

    private static void ResetDisplayFormulaPosition(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            ResetDisplayFormulaPosition(range);
        }
        finally { Release(range); }
    }

    private static void ResetDisplayFormulaPosition(Range formulaRange)
    {
        ResetRangeFontPosition(formulaRange);
        ResetParagraphTypingPosition(formulaRange);
    }

    private static void NormalizeFollowingInlineProseBaseline(Range formulaRange)
    {
        Document? document = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? precedingWindow = null;
        Range? preceding = null;
        Range? trailingHostWindow = null;
        Range? trailing = null;
        InlineShapes? shapes = null;
        InlineShape? nextShape = null;
        Range? nextShapeRange = null;
        OMaths? maths = null;
        OMath? nextMath = null;
        Range? nextMathRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            document = formulaRange.Document;
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var paragraphBodyEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);

            // Only the nearby prose can determine the insertion baseline. Capturing
            // every InlineShape/OMath in a paragraph made appending the Nth inline
            // formula rescan formulas 1..N-1 twice. Keep the format probe bounded;
            // hidden OLE field-code ranges inside this small window are still
            // excluded by the same host-aware logic used previously.
            const int maximumProbeCharacters = 256;
            var precedingStart = Math.Max(
                paragraphRange.Start,
                formulaRange.Start - maximumProbeCharacters);
            if (formulaRange.Start > precedingStart)
            {
                precedingWindow = document.Range(precedingStart, formulaRange.Start);
                var excludedHostRanges = CaptureNonProseHostRanges(precedingWindow);
                preceding = FindOrdinaryVisibleCharacterRange(
                    document,
                    formulaRange.Start - 1,
                    precedingStart,
                    step: -1,
                    excludedHostRanges);
            }

            var targetPosition = 0;
            if (preceding is not null)
            {
                font = preceding.Font;
                var precedingPosition = font.Position;
                if (precedingPosition != (int)WdConstants.wdUndefined
                    && precedingPosition >= -256
                    && precedingPosition <= 256)
                    targetPosition = precedingPosition;
                Release(font);
                font = null;
            }

            var trailingEnd = Math.Max(formulaRange.End, paragraphBodyEnd);
            if (trailingEnd <= formulaRange.End) return;
            trailingHostWindow = document.Range(formulaRange.End, trailingEnd);

            // Range collections are in document order. The first following host is
            // sufficient; do not enumerate every object in the paragraph.
            shapes = trailingHostWindow.InlineShapes;
            if (shapes.Count > 0)
            {
                nextShape = shapes[1];
                nextShapeRange = nextShape.Range;
                trailingEnd = Math.Min(trailingEnd, nextShapeRange.Start);
            }
            maths = trailingHostWindow.OMaths;
            if (maths.Count > 0)
            {
                nextMath = maths[1];
                nextMathRange = nextMath.Range;
                trailingEnd = Math.Min(trailingEnd, nextMathRange.Start);
            }

            if (trailingEnd <= formulaRange.End) return;
            trailing = document.Range(formulaRange.End, trailingEnd);
            if (!ContainsVisibleBodyText(trailing.Text)) return;
            font = trailing.Font;
            var currentPosition = font.Position;
            if (currentPosition == (int)WdConstants.wdUndefined
                || currentPosition != targetPosition)
                font.Position = targetPosition;
        }
        catch
        {
            // Baseline repair is best-effort and must not interrupt Word input.
        }
        finally
        {
            Release(font);
            Release(nextMathRange);
            Release(nextMath);
            Release(maths);
            Release(nextShapeRange);
            Release(nextShape);
            Release(shapes);
            Release(trailing);
            Release(trailingHostWindow);
            Release(preceding);
            Release(precedingWindow);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(document);
        }
    }

    private static void ResetParagraphTypingPosition(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        Range? nextCharacter = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.End > paragraphRange.Start)
            {
                paragraphMark = paragraphRange.Duplicate;
                paragraphMark.SetRange(paragraphRange.End - 1, paragraphRange.End);
                ResetRangeFontPosition(paragraphMark);
            }

            if (formulaRange.End >= paragraphRange.End) return;
            nextCharacter = paragraphRange.Duplicate;
            nextCharacter.SetRange(
                formulaRange.End,
                Math.Min(formulaRange.End + 1, paragraphRange.End));
            if (nextCharacter.Text is "\v" or "\r" or "\n")
                ResetRangeFontPosition(nextCharacter);
        }
        catch
        {
            // Baseline restoration is best-effort and must not invalidate the
            // formula that has already been inserted or resized.
        }
        finally
        {
            Release(nextCharacter);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static string InlineBaselineBookmarkName(string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var parsed))
            throw new InvalidDataException("VisualTeX formulaId must be a UUID.");
        return InlineBaselineBookmarkPrefix + parsed.ToString("N");
    }


    private Range PrepareInlineBaselineSentinelBeforeInsert(
        Document document,
        Range insertionRange,
        string formulaId,
        bool createBookmark = true)
    {
        if (createBookmark)
            RemoveInlineBaselineSentinel(document, formulaId);
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? placeholders = null;
        Range? formulaPlaceholder = null;
        Range? guard = null;
        Range? sentinel = null;
        try
        {
            placeholders = insertionRange.Duplicate;
            placeholders.Collapse(WdCollapseDirection.wdCollapseEnd);
            var start = placeholders.Start;
            // Word can absorb the first ordinary character immediately after a
            // newly materialized inline OMath. Keep two temporary hidden ordinary
            // spaces only while importing. FinalizeInlineOmmlBoundary removes both
            // and deletes the temporary VTBL bookmark before the operation returns.
            placeholders.Text = BulkInlineFormulaPlaceholder
                + InlineMathGuard
                + InlineBaselineSentinel;
            formulaPlaceholder = document.Range(start, start + BulkInlineFormulaPlaceholder.Length);
            guard = document.Range(
                formulaPlaceholder.End,
                formulaPlaceholder.End + InlineMathGuard.Length);
            sentinel = document.Range(
                guard.End,
                guard.End + InlineBaselineSentinel.Length);
            ConfigureInlineBaselineSentinel(guard);
            ConfigureInlineBaselineSentinel(sentinel);
            if (createBookmark)
            {
                bookmarks = document.Bookmarks;
                bookmark = bookmarks.Add(InlineBaselineBookmarkName(formulaId), sentinel);
            }
            var result = formulaPlaceholder;
            formulaPlaceholder = null;
            return result;
        }
        finally
        {
            Release(sentinel);
            Release(guard);
            Release(formulaPlaceholder);
            Release(placeholders);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private int PrepareInlineBaselineSentinelAfterFormula(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? sentinel = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? probe = null;
        try
        {
            var name = InlineBaselineBookmarkName(formulaId);
            var temporaryNativeMathBoundary = RangeContainsMath(formulaRange);
            bookmarks = document.Bookmarks;
            if (bookmarks.Exists(name))
            {
                bookmark = bookmarks[name];
                sentinel = bookmark.Range;
                if (!temporaryNativeMathBoundary
                    && IsUsableInlineBaselineSentinel(sentinel, formulaRange))
                    return NormalizeInlineBaselineBoundary(
                        document,
                        formulaRange,
                        formulaId);
                bookmark.Delete();
                if (IsKnownInlineBaselineSentinel(sentinel.Text))
                    sentinel.Delete();
                Release(sentinel);
                sentinel = null;
                Release(bookmark);
                bookmark = null;
            }

            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0)
                throw new InvalidOperationException("Word could not locate the inline formula paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var insertionPosition = Math.Max(formulaRange.End, paragraphRange.Start);
            var finalPosition = Math.Max(insertionPosition, paragraphRange.End - 1);
            for (var position = insertionPosition; position <= finalPosition; position++)
            {
                Release(probe);
                object probeStart = position;
                object probeEnd = Math.Min(position + 1, paragraphRange.End);
                probe = document.Range(ref probeStart, ref probeEnd);
                if (RangeContainsMath(probe)) continue;
                insertionPosition = position;
                break;
            }

            if (temporaryNativeMathBoundary)
                return CreateInlineOmmlTemporaryBoundary(
                    document,
                    bookmarks,
                    name,
                    insertionPosition);

            object sentinelStart = insertionPosition;
            object sentinelEnd = insertionPosition;
            sentinel = document.Range(ref sentinelStart, ref sentinelEnd);
            bookmark = bookmarks.Add(name, sentinel);
            Release(sentinel);
            sentinel = null;
            Release(bookmark);
            bookmark = null;
            return NormalizeInlineBaselineBoundary(
                document,
                formulaRange,
                formulaId);
        }
        finally
        {
            Release(probe);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(sentinel);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static int CreateInlineOmmlTemporaryBoundary(
        Document document,
        Bookmarks bookmarks,
        string bookmarkName,
        int insertionPosition)
    {
        Range? boundary = null;
        Range? guard = null;
        Range? sentinel = null;
        Bookmark? bookmark = null;
        try
        {
            boundary = document.Range(insertionPosition, insertionPosition);
            boundary.Text = InlineMathGuard + InlineBaselineSentinel;
            guard = document.Range(
                insertionPosition,
                insertionPosition + InlineMathGuard.Length);
            sentinel = document.Range(
                guard.End,
                guard.End + InlineBaselineSentinel.Length);
            ConfigureInlineBaselineSentinel(guard);
            ConfigureInlineBaselineSentinel(sentinel);
            bookmark = bookmarks.Add(bookmarkName, sentinel);
            return sentinel.End;
        }
        finally
        {
            Release(bookmark);
            Release(sentinel);
            Release(guard);
            Release(boundary);
        }
    }

    private static bool IsKnownInlineBaselineSentinel(string? text) =>
        string.IsNullOrEmpty(text)
        || string.Equals(text, InlineOleTypingAnchor, StringComparison.Ordinal)
        || string.Equals(text, InlineBaselineSentinel, StringComparison.Ordinal)
        || string.Equals(text, LegacyInlineBaselineSentinel, StringComparison.Ordinal)
        || string.Equals(
            text,
            LegacyInlineNonbreakingBaselineSentinel,
            StringComparison.Ordinal);

    private static bool IsUsableInlineBaselineSentinel(
        Range sentinel,
        Range formulaRange)
    {
        if (sentinel.Start != sentinel.End
            && !IsKnownInlineBaselineSentinel(sentinel.Text))
            return false;
        // Current OLE bookmarks own one zero-width typing anchor. Collapsed and
        // legacy one-character bookmarks are accepted long enough to be migrated
        // by NormalizeInlineBaselineBoundary.
        return sentinel.Start >= formulaRange.End
            && sentinel.Start <= formulaRange.End + 8;
    }

    private static bool IsHiddenTextRange(Range range)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            return font.Hidden != 0;
        }
        catch { return false; }
        finally { Release(font); }
    }

    private static InlineFollowingTextVisibility?
        CaptureInlineFollowingTextVisibility(Range boundaryRange)
    {
        Document? document = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? following = null;
        Range? trailingFormulaCharacter = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            document = boundaryRange.Document;
            paragraphs = boundaryRange.Paragraphs;
            if (paragraphs.Count == 0) return null;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var start = Math.Max(boundaryRange.End, paragraphRange.Start);
            // Paragraph.Range includes its final paragraph mark. That mark is not
            // following prose and must never contribute to the character count:
            // Word 2021 can move the temporary OMML guards while replacing the
            // equation, and restoring one character too far then unhides a guard
            // as a visible trailing ASCII space.
            var end = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            if (end <= start) return null;
            // Older Word 2021 builds can serialize VisualTeX's hidden ordinary
            // guard as a one-character range whose Text is empty. Trim only such
            // hidden, known boundary artifacts from the tail before recording the
            // amount of real following prose. This also repairs documents touched
            // by builds that left the temporary guards behind.
            while (end > start)
            {
                Range? trailingBoundary = null;
                try
                {
                    trailingBoundary = document.Range(end - 1, end);
                    if (!IsHiddenTextRange(trailingBoundary)
                        || !IsKnownInlineBaselineSentinel(trailingBoundary.Text))
                        break;
                    end--;
                }
                finally { Release(trailingBoundary); }
            }
            if (end <= start) return null;
            following = document.Range(start, end);
            font = following.Font;
            var hidden = font.Hidden;
            if (hidden is not (0 or -1)) return null;

            // A previous buggy inline OMML update can leave its hidden ASCII
            // guard as the last native-math character and propagate Hidden=-1 to
            // all following prose. Such a run was ordinary visible body text
            // before VisualTeX touched it, so repair it while replacing again.
            var inheritedVisualTeXHiddenState = false;
            if (boundaryRange.End > boundaryRange.Start)
            {
                trailingFormulaCharacter = document.Range(
                    boundaryRange.End - 1,
                    boundaryRange.End);
                inheritedVisualTeXHiddenState = string.Equals(
                        trailingFormulaCharacter.Text,
                        InlineMathGuard,
                        StringComparison.Ordinal)
                    && IsHiddenTextRange(trailingFormulaCharacter)
                    && RangeContainsMath(trailingFormulaCharacter);
            }

            return new InlineFollowingTextVisibility
            {
                CharacterCount = end - start,
                Hidden = inheritedVisualTeXHiddenState ? 0 : hidden,
            };
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(font);
            Release(trailingFormulaCharacter);
            Release(following);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(document);
        }
    }

    private static void RestoreInlineFollowingTextVisibility(
        Range formulaRange,
        InlineFollowingTextVisibility? visibility)
    {
        if (visibility is null || visibility.CharacterCount <= 0) return;
        Document? document = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? following = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            document = formulaRange.Document;
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var start = Math.Max(formulaRange.End, paragraphRange.Start);
            var paragraphBodyEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            var end = Math.Min(
                paragraphBodyEnd,
                start + visibility.CharacterCount);
            if (end <= start) return;
            following = document.Range(start, end);
            font = following.Font;
            font.Hidden = visibility.Hidden;
        }
        catch
        {
            // Visibility restoration is protective cleanup. A stale paragraph
            // range must never invalidate the already committed native formula.
        }
        finally
        {
            Release(font);
            Release(following);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(document);
        }
    }

    private static int NormalizeInlineBaselineBoundary(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? sentinel = null;
        try
        {
            var name = InlineBaselineBookmarkName(formulaId);
            bookmarks = document.Bookmarks;
            if (bookmarks.Exists(name))
            {
                bookmark = bookmarks[name];
                sentinel = bookmark.Range;

                if (sentinel.Start == formulaRange.End
                    && string.Equals(
                        sentinel.Text,
                        InlineOleTypingAnchor,
                        StringComparison.Ordinal))
                {
                    ConfigureInlineOleTypingAnchor(sentinel, formulaRange);
                    return sentinel.End;
                }

                // Replace every historical/collapsed VTBL representation with
                // the current zero-width OLE typing anchor. The bookmark owns the
                // marker, so deleting it cannot remove user-authored prose.
                var ownedBoundaryText = sentinel.Text;
                var ownsBoundaryCharacter = sentinel.Start < sentinel.End
                    && IsKnownInlineBaselineSentinel(ownedBoundaryText);
                bookmark.Delete();
                if (ownsBoundaryCharacter)
                    sentinel.Delete();
            }

            // An interrupted insertion can also leave one sacrificial guard at
            // OMath.Range.End. Remove only VisualTeX legacy characters or hidden
            // ordinary spaces; visible user-authored spacing is never touched.
            for (var attempt = 0; attempt < 2; attempt++)
            {
                if (!TryDeleteTemporaryInlineBoundaryAt(
                        document,
                        formulaRange.End))
                    break;
            }

            var boundaryPosition = formulaRange.End;
            object boundaryStart = boundaryPosition;
            object boundaryEnd = Math.Min(
                document.Content.End,
                boundaryPosition + 1);
            Release(sentinel);
            sentinel = document.Range(ref boundaryStart, ref boundaryEnd);
            // Assigning Text on a collapsed Range at InlineShape.End can retain
            // object-side affinity immediately after OMML -> OLE conversion and
            // invalidate the new OLE object. Insert before the following prose or
            // paragraph mark so the anchor is unambiguously outside the object.
            sentinel.InsertBefore(InlineOleTypingAnchor);
            sentinel.SetRange(
                boundaryPosition,
                boundaryPosition + InlineOleTypingAnchor.Length);
            ConfigureInlineOleTypingAnchor(sentinel, formulaRange);
            Release(bookmark);
            bookmark = bookmarks.Add(name, sentinel);
            return sentinel.End;
        }
        finally
        {
            Release(sentinel);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static bool RangeContainsMath(Range range)
    {
        OMaths? maths = null;
        try
        {
            maths = range.OMaths;
            return maths.Count > 0;
        }
        catch { return false; }
        finally { Release(maths); }
    }

    private int EnsureInlineBaselineSentinel(
        Range formulaRange,
        string formulaId,
        bool placeOutsideNativeMath = false)
    {
        Document? document = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? sentinel = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            bookmarks = document.Bookmarks;
            var name = InlineBaselineBookmarkName(formulaId);
            if (bookmarks.Exists(name))
            {
                bookmark = bookmarks[name];
                sentinel = bookmark.Range;
                if (IsUsableInlineBaselineSentinel(sentinel, formulaRange))
                {
                    ResetParagraphTypingPosition(formulaRange);
                    return NormalizeInlineBaselineBoundary(
                        document,
                        formulaRange,
                        formulaId);
                }
            }
        }
        finally
        {
            Release(sentinel);
            Release(bookmark);
            Release(bookmarks);
            Release(document);
        }

        document = _application.ActiveDocument
            ?? throw new InvalidOperationException("No active Word document.");
        try
        {
            var result = PrepareInlineBaselineSentinelAfterFormula(
                document,
                formulaRange,
                formulaId);
            if (placeOutsideNativeMath)
                ResetParagraphTypingPosition(formulaRange);
            return result;
        }
        finally { Release(document); }
    }

    private static void RemoveInlineOleTypingAnchorAfter(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            RemoveInlineOleTypingAnchorAfter(range);
        }
        finally { Release(range); }
    }

    private static void RemoveInlineOleTypingAnchorAfter(Range formulaRange)
    {
        Document? document = null;
        Range? content = null;
        Range? anchor = null;
        try
        {
            document = formulaRange.Document;
            content = document.Content;
            var position = formulaRange.End;
            var contentStart = content.Start;
            var contentEnd = content.End;
            if (position < contentStart || position >= contentEnd) return;
            anchor = document.Range(position, Math.Min(position + 1, contentEnd));
            if (!string.Equals(
                    anchor.Text,
                    InlineOleTypingAnchor,
                    StringComparison.Ordinal))
                return;
            // Word can report an ordinary character immediately after an OLE as
            // math-affiliated because the adjacent object participates in the
            // same layout run. This helper is called only for a confirmed
            // VisualTeX inline OLE, whose first U+200C at Range.End is owned by us.
            anchor.Delete();
        }
        catch
        {
            // The formula may have been deleted by Word between range capture and
            // cleanup. Orphan-anchor removal is best-effort in rollback paths.
        }
        finally
        {
            Release(anchor);
            Release(content);
            Release(document);
        }
    }

    private static void RemoveInlineBaselineSentinel(Document document, string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? sentinel = null;
        try
        {
            var name = InlineBaselineBookmarkName(formulaId);
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            sentinel = bookmark.Range;
            var sentinelStart = sentinel.Start;
            var ownedBoundaryHasWidth = sentinel.Start < sentinel.End
                && IsKnownInlineBaselineSentinel(sentinel.Text);
            bookmark.Delete();
            if (ownedBoundaryHasWidth)
                sentinel.Delete();

            // Remove a temporary guard from interrupted insertions. Probe both
            // sides of the former bookmark because deleting the bookmarked marker
            // shifts the following text one position to the left.
            TryDeleteTemporaryInlineBoundaryAt(document, sentinelStart - 1);
            TryDeleteTemporaryInlineBoundaryAt(document, sentinelStart);
        }
        catch
        {
            // A stale or externally edited sentinel must never block formula work.
        }
        finally
        {
            Release(sentinel);
            Release(bookmark);
            Release(bookmarks);
        }
    }


    private static bool TryDeleteTemporaryInlineBoundaryAt(
        Document document,
        int position)
    {
        Range? candidate = null;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = document.Content.End;
            if (position < contentStart || position >= contentEnd) return false;
            object candidateStart = position;
            object candidateEnd = Math.Min(contentEnd, position + 1);
            candidate = document.Range(ref candidateStart, ref candidateEnd);
            var text = candidate.Text;
            var removable =
                string.Equals(text, BulkInlineFormulaPlaceholder, StringComparison.Ordinal)
                || string.Equals(text, LegacyInlineMathGuard, StringComparison.Ordinal)
                || string.Equals(text, LegacyInlineBaselineSentinel, StringComparison.Ordinal)
                || string.Equals(text, InlineMathGuard, StringComparison.Ordinal)
                    && IsHiddenTextRange(candidate);
            // U+00A0 is also the correct Word representation of explicit LaTeX
            // spacing (for example `~` and `\ `). Delete it only when the VTBL
            // bookmark itself owns that legacy marker, never by proximity.
            if (!removable || RangeContainsMath(candidate)) return false;
            candidate.Delete();
            return true;
        }
        catch { return false; }
        finally { Release(candidate); }
    }

    private static void ConfigureInlineBaselineSentinel(Range range)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            font.Position = 0;
            // This formatting is applied only to temporary insertion guards.
            // Both characters are deleted before the operation is committed.
            font.Hidden = -1;
        }
        finally { Release(font); }
    }

    private static void ConfigureInlineOleTypingAnchor(
        Range range,
        Range formulaRange)
    {
        Range? formattingSource = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            // While replacing native OMML, Word can still treat the boundary
            // character at OMath.Range.End as part of the math zone. Applying
            // ordinary-text properties such as Subscript/Superscript then raises
            // 0x800A1863 and aborts the update. That boundary is only a temporary
            // insertion guard; the persistent OLE typing anchor is configured
            // again after the new InlineShape exists outside native math.
            if (RangeContainsMath(range))
            {
                ConfigureInlineBaselineSentinel(range);
                return;
            }

            formattingSource = FindInlineTypingFormatSource(formulaRange, range);
            if (formattingSource is not null)
                CopyInlineTypingCharacterFormatting(formattingSource, range);

            font = range.Font;
            if (formattingSource is null)
                font.Position = 0;
            font.Hidden = 0;
            try { font.Subscript = 0; } catch { }
            try { font.Superscript = 0; } catch { }
            try { font.Spacing = 0; } catch { }
            try { font.Scaling = 100; } catch { }
        }
        finally
        {
            Release(font);
            Release(formattingSource);
        }
    }

    private static Range? FindInlineTypingFormatSource(
        Range formulaRange,
        Range typingAnchor)
    {
        Document? document = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? precedingWindow = null;
        Range? followingWindow = null;
        try
        {
            document = formulaRange.Document;
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return null;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var paragraphBodyEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            const int maximumProbeCharacters = 256;

            // The prose before an inline formula is the authoritative typing
            // style. Restrict host discovery to the same bounded region that can
            // actually be probed; scanning every OLE in a long formula paragraph
            // made caret restoration grow linearly with formula count.
            var precedingStart = Math.Max(
                paragraphRange.Start,
                formulaRange.Start - maximumProbeCharacters);
            if (formulaRange.Start > precedingStart)
            {
                precedingWindow = document.Range(precedingStart, formulaRange.Start);
                var precedingHosts = CaptureNonProseHostRanges(precedingWindow);
                var preceding = FindOrdinaryVisibleCharacterRange(
                    document,
                    formulaRange.Start - 1,
                    precedingStart,
                    step: -1,
                    precedingHosts);
                if (preceding is not null) return preceding;
            }

            // Paragraph-leading formulas have no preceding prose. In that case,
            // inherit from the first nearby ordinary character after the formula.
            var followingStart = Math.Max(typingAnchor.End, formulaRange.End);
            var followingEnd = Math.Min(
                paragraphBodyEnd,
                followingStart + maximumProbeCharacters);
            if (followingEnd <= followingStart) return null;
            followingWindow = document.Range(followingStart, followingEnd);
            var followingHosts = CaptureNonProseHostRanges(followingWindow);
            return FindOrdinaryVisibleCharacterRange(
                document,
                followingStart,
                followingEnd,
                step: 1,
                followingHosts);
        }
        finally
        {
            Release(followingWindow);
            Release(precedingWindow);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(document);
        }
    }

    private static Range? FindOrdinaryVisibleCharacterRange(
        Document document,
        int startPosition,
        int boundaryPosition,
        int step,
        IReadOnlyList<NonProseHostRange> excludedHostRanges)
    {
        if (step is not (-1 or 1)) return null;
        const int maximumProbeCharacters = 256;
        var position = startPosition;
        for (var probeIndex = 0;
             probeIndex < maximumProbeCharacters
             && (step > 0
                 ? position < boundaryPosition
                 : position >= boundaryPosition);
             probeIndex++, position += step)
        {
            Range? probe = null;
            try
            {
                if (position < 0) break;
                if (excludedHostRanges.Any(host => host.Contains(position)))
                    continue;
                probe = document.Range(position, position + 1);
                if (!ContainsVisibleBodyText(probe.Text)) continue;
                var result = probe.Duplicate;
                return result;
            }
            catch (COMException)
            {
                // Keep probing nearby ordinary prose.
            }
            finally
            {
                Release(probe);
            }
        }
        return null;
    }

    private static List<NonProseHostRange> CaptureNonProseHostRanges(
        Range paragraphRange)
    {
        var result = new List<NonProseHostRange>();
        InlineShapes? shapes = null;
        OMaths? maths = null;
        try
        {
            shapes = paragraphRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                try
                {
                    shape = shapes[index];
                    range = shape.Range;
                    result.Add(new NonProseHostRange(range.Start, range.End));
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }

            maths = paragraphRange.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    result.Add(new NonProseHostRange(range.Start, range.End));
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
            return result;
        }
        finally
        {
            Release(maths);
            Release(shapes);
        }
    }

    private static void CopyInlineTypingCharacterFormatting(
        Range source,
        Range target)
    {
        Microsoft.Office.Interop.Word.Font? sourceFont = null;
        Microsoft.Office.Interop.Word.Font? targetFont = null;
        try
        {
            var targetStart = target.Start;
            var alreadyTypingAnchor = target.End == targetStart + InlineOleTypingAnchor.Length
                && string.Equals(
                    target.Text,
                    InlineOleTypingAnchor,
                    StringComparison.Ordinal);
            if (!alreadyTypingAnchor)
            {
                target.SetRange(targetStart, target.End);
                target.Text = InlineOleTypingAnchor;
            }
            target.SetRange(
                targetStart,
                targetStart + InlineOleTypingAnchor.Length);

            sourceFont = source.Font;
            targetFont = target.Font;
            targetFont.Name = sourceFont.Name;
            try { targetFont.NameAscii = sourceFont.NameAscii; } catch { }
            try { targetFont.NameFarEast = sourceFont.NameFarEast; } catch { }
            try { targetFont.NameOther = sourceFont.NameOther; } catch { }
            targetFont.Size = sourceFont.Size;
            targetFont.Bold = sourceFont.Bold;
            targetFont.Italic = sourceFont.Italic;
            var sourcePosition = sourceFont.Position;
            targetFont.Position = sourcePosition == (int)WdConstants.wdUndefined
                ? 0
                : sourcePosition;
            try { targetFont.Underline = sourceFont.Underline; } catch { }
            try { targetFont.Color = sourceFont.Color; } catch { }
        }
        finally
        {
            Release(targetFont);
            Release(sourceFont);
        }
    }

    private static void ResetRangeFontPosition(Range range)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            font.Position = 0;
        }
        finally { Release(font); }
    }

    private static void ApplyInlineBaseline(
        InlineShape shape,
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline,
        double semanticFontSizePoints) =>
        ApplyInlineBaseline(
            shape,
            actualHeightPoints,
            exportedHeight,
            exportedBaseline,
            existingFontPosition: null,
            sourceSemanticFontSizePoints: semanticFontSizePoints,
            targetSemanticFontSizePoints: semanticFontSizePoints);

    private static void ApplyInlineBaseline(
        InlineShape shape,
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline,
        float? existingFontPosition,
        double sourceSemanticFontSizePoints,
        double targetSemanticFontSizePoints)
    {
        Range? range = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = shape.Range;
            font = range.Font;
            font.Position = WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                actualHeightPoints,
                exportedHeight,
                exportedBaseline,
                existingFontPosition,
                sourceSemanticFontSizePoints,
                targetSemanticFontSizePoints);
        }
        finally
        {
            Release(font);
            Release(range);
        }
    }

    private void RestoreTypingBaselineAfterMathTypeConversion(InlineShape shape)
    {
        Range? range = null;
        Selection? selection = null;
        try
        {
            range = shape.Range;
            // Format conversion replaces an existing inline equation at the same
            // text boundary. Its surrounding prose already owns the correct
            // character baseline, so do not rescan/rewrite the entire following
            // prose segment here. Re-establish only the paragraph end and the
            // collapsed insertion format immediately after the new MathType OLE.
            ResetParagraphTypingPosition(range);
            selection = _application.Selection;
            selection.SetRange(range.End, range.End);
            ApplyInlineTypingFormattingToSelection(selection, range);
        }
        catch
        {
            // The MathType OLE is already structurally valid. A transient Word
            // insertion-format refusal must not invalidate the conversion.
        }
        finally
        {
            Release(selection);
            Release(range);
        }
    }

    private void RestoreTypingBaselineAfter(
        InlineShape shape,
        bool ensureTypingAnchor = false)
    {
        Document? document = null;
        Range? range = null;
        try
        {
            range = shape.Range;
            document = range.Document;
            NormalizeFollowingInlineProseBaseline(range);
            ResetParagraphTypingPosition(range);
            var caretPosition = range.End;
            if (ensureTypingAnchor)
            {
                var metadata = WordFormulaMetadataReader.TryRead(shape);
                if (metadata is not null)
                    caretPosition = EnsureInlineBaselineSentinel(
                        range,
                        metadata.FormulaId);
            }

            // The caret must inherit the ordinary prose run, not a hard-coded
            // baseline. The persistent anchor above makes that formatting stable
            // across the first keystroke, arrow navigation and later mouse clicks.
            RestoreTypingCaretAt(document, caretPosition, range);
        }
        finally
        {
            Release(range);
            Release(document);
        }
    }

    internal void NormalizeInlineOleParagraphBaselinesBeforeSave(Document document)
    {
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? range = null;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = null;
                Release(range);
                range = null;
                try
                {
                    shape = shapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (!string.Equals(
                            metadata?.DisplayMode,
                            "inline",
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    range = shape.Range;
                    // Normalize every inline OLE segment in document order. The
                    // previous one-per-paragraph gate repaired only the first OLE,
                    // allowing a polluted run between later formulas to become the
                    // formatting source for every subsequent typing anchor.
                    NormalizeFollowingInlineProseBaseline(range);
                    ResetParagraphTypingPosition(range);
                }
                catch
                {
                    // Saving must remain available when one stale or externally
                    // edited object cannot be inspected.
                }
            }
        }
        finally
        {
            Release(range);
            Release(shape);
            Release(shapes);
        }
    }

    private static Range ResolveCurrentInlineOmmlRange(
        Document document,
        Range fallbackRange,
        string formulaId)
    {
        Bookmark? formulaBookmark = null;
        try
        {
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            return formulaBookmark is null
                ? fallbackRange.Duplicate
                : WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
        }
        catch
        {
            return fallbackRange.Duplicate;
        }
        finally { Release(formulaBookmark); }
    }

    private static void RemoveBulkInlineOmmlTemporaryBoundary(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        _ = formulaId;
        // Bulk import creates every formula with a fresh id and already owns the
        // live OMath range. Do not create or search a document-level VTBL
        // bookmark: Word 2021 makes Bookmarks.Exists progressively slower as the
        // imported document grows. The two guards are immediately adjacent to
        // the live OMath end and can be removed locally.
        for (var attempt = 0; attempt < 4; attempt++)
        {
            if (!TryDeleteInlineOmmlGuardAtBoundary(
                    document,
                    formulaRange,
                    forceTextReplacement: false))
                break;
        }
    }

    private static void RemoveInlineOmmlTemporaryBoundary(
        Document document,
        Range formulaRange,
        string formulaId,
        bool forceTextReplacement = true)
    {
        // VTBL is only a temporary insertion guard for native OMath. Persisting
        // even a collapsed bookmark makes Word show a dotted placeholder when
        // "Show bookmarks" is enabled. Remove the bookmark and every temporary
        // character before returning; VTOMML remains the durable formula identity.
        // Re-resolve the final OMath after every deletion: Word mutates live Range
        // coordinates when the bookmarked sentinel is removed, so a cached End can
        // skip the last hidden ASCII guard and leave the caret immediately before it.
        Range? currentRange = null;
        try
        {
            RemoveInlineBaselineSentinel(document, formulaId);
            if (!forceTextReplacement)
            {
                // Bulk import already owns a live range for the just-inserted
                // OMath and writes complete paragraphs before replacing formula
                // placeholders. Avoid re-opening the VTOMML bookmark for every
                // sacrificial guard: Word updates this live range as the hidden
                // runs are deleted, and the bulk spacing acceptance verifies that
                // neither guard survives. The stricter single-insert path below
                // still re-resolves the durable bookmark after every deletion.
                for (var attempt = 0; attempt < 4; attempt++)
                {
                    if (!TryDeleteInlineOmmlGuardAtBoundary(
                            document,
                            formulaRange,
                            forceTextReplacement: false))
                        break;
                }
                return;
            }

            for (var attempt = 0; attempt < 8; attempt++)
            {
                Release(currentRange);
                currentRange = ResolveCurrentInlineOmmlRange(
                    document,
                    formulaRange,
                    formulaId);
                if (!TryDeleteInlineOmmlGuardAtBoundary(
                        document,
                        currentRange,
                        forceTextReplacement: true))
                    break;
            }
        }
        finally { Release(currentRange); }
    }

    private static bool TryDeleteInlineOmmlGuardAtBoundary(
        Document document,
        Range formulaRange,
        bool forceTextReplacement)
    {
        // Word 2021 has two boundary representations after FormattedText/InsertFile:
        // the hidden guard can remain outside at OMath.End, or it can be swallowed
        // as the final hidden m:r and make OMath.End advance by one. Probe both
        // locations, deleting only VisualTeX legacy markers or hidden ASCII space.
        if (TryDeleteInlineOmmlGuardAt(
                document,
                formulaRange.End,
                forceTextReplacement))
            return true;
        return formulaRange.End > formulaRange.Start
            && TryDeleteInlineOmmlGuardAt(
                document,
                formulaRange.End - 1,
                forceTextReplacement);
    }

    private static bool TryDeleteInlineOmmlGuardAt(
        Document document,
        int position,
        bool forceTextReplacement)
    {
        Range? candidate = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = document.Content.End;
            if (position < contentStart || position >= contentEnd) return false;
            candidate = document.Range(position, Math.Min(position + 1, contentEnd));
            var text = candidate.Text;
            var legacyGuard =
                string.Equals(text, LegacyInlineMathGuard, StringComparison.Ordinal)
                || string.Equals(text, LegacyInlineBaselineSentinel, StringComparison.Ordinal);
            var privateUseFormulaPlaceholder = string.Equals(
                text,
                BulkInlineFormulaPlaceholder,
                StringComparison.Ordinal);
            font = candidate.Font;
            var hiddenAsciiGuard = string.Equals(
                    text,
                    InlineMathGuard,
                    StringComparison.Ordinal)
                && font.Hidden != 0;
            // Word 2021 (including build 16.0.14332) can expose the swallowed
            // hidden ASCII guard at OMath.End as a one-character Range with an
            // empty Text value. This helper is called only at End/End-1 of the
            // formula, so Hidden + empty is the same VisualTeX-owned temporary
            // boundary, not arbitrary user prose.
            var hiddenEmptyGuard = string.IsNullOrEmpty(text)
                && font.Hidden != 0;
            if (!legacyGuard
                && !privateUseFormulaPlaceholder
                && !hiddenAsciiGuard
                && !hiddenEmptyGuard)
                return false;

            // Immediately after OMath.BuildUp, Word can report this external guard
            // as math-affiliated even though it is outside OMath.Range. The guard
            // was created by PrepareInlineBaselineSentinelBeforeInsert and sits at
            // the re-resolved final OMath.End. Assigning empty text is more reliable
            // than Range.Delete at this boundary; some Word builds return from
            // Delete without actually removing the hidden run.
            if (forceTextReplacement)
                candidate.Text = string.Empty;
            else
                candidate.Delete();
            return true;
        }
        catch { return false; }
        finally
        {
            Release(font);
            Release(candidate);
        }
    }

    private static Range? TryResolveInlineOmmlFromLiveRange(Range formulaRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? resolved = null;
        try
        {
            maths = formulaRange.OMaths;
            if (maths.Count != 1) return null;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathInline) return null;
            resolved = math.Range.Duplicate;
            var result = resolved;
            resolved = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(resolved);
            Release(math);
            Release(maths);
        }
    }

    private Range FinalizeInlineOmmlBoundaryFromLiveRange(
        Document document,
        Range formulaRange,
        string formulaId,
        InlineFollowingTextVisibility? followingTextVisibility)
    {
        Range? currentRange = null;
        try
        {
            // ReplaceOmml has just materialized this exact native OMath and still
            // owns its live Range. Word updates that RCW as adjacent guard text is
            // removed, so repeatedly reopening VTOMML after every guard is pure
            // document-size work. Keep the same robust hidden-guard Text=""
            // removal used by the single-edit path, but probe End/End-1 locally.
            RemoveInlineBaselineSentinel(document, formulaId);
            for (var attempt = 0; attempt < 8; attempt++)
            {
                if (!TryDeleteInlineOmmlGuardAtBoundary(
                        document,
                        formulaRange,
                        forceTextReplacement: true))
                    break;
            }

            currentRange = TryResolveInlineOmmlFromLiveRange(formulaRange);
            if (currentRange is null)
            {
                // Unexpected Word range invalidation keeps the older defensive
                // bookmark-based behavior as a correctness fallback.
                FinalizeInlineOmmlBoundary(
                    document,
                    formulaRange,
                    formulaId,
                    moveCaretOutsideMath: false,
                    followingTextVisibility);
                return ResolveCurrentInlineOmmlRange(
                    document,
                    formulaRange,
                    formulaId);
            }

            RestoreInlineFollowingTextVisibility(
                currentRange,
                followingTextVisibility);
            ResetParagraphTypingPosition(currentRange);
            var result = currentRange;
            currentRange = null;
            return result;
        }
        finally { Release(currentRange); }
    }

    private void FinalizeInlineOmmlBoundary(
        Document document,
        Range formulaRange,
        string formulaId,
        bool moveCaretOutsideMath,
        InlineFollowingTextVisibility? followingTextVisibility = null)
    {
        Range? currentRange = null;
        try
        {
            RemoveInlineOmmlTemporaryBoundary(document, formulaRange, formulaId);
            currentRange = ResolveCurrentInlineOmmlRange(
                document,
                formulaRange,
                formulaId);
            RestoreInlineFollowingTextVisibility(
                currentRange,
                followingTextVisibility);
            ResetParagraphTypingPosition(currentRange);
            if (moveCaretOutsideMath)
                MoveCaretOutsideInlineOmml(currentRange);
        }
        finally { Release(currentRange); }
    }

    private void MoveCaretOutsideInlineOmml(Range formulaRange)
    {
        Selection? selection = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            selection = _application.Selection;
            selection.SetRange(formulaRange.End, formulaRange.End);

            // A collapsed range at OMath.End still has mathematical caret
            // affinity. Word's native right-arrow operation switches that same
            // coordinate to the ordinary text side without inserting or skipping
            // a character. Following prose is therefore typed outside the OMath.
            _ = selection.MoveRight(
                WdUnits.wdCharacter,
                1,
                WdMovementType.wdMove);
            font = selection.Font;
            font.Position = 0;
            font.Hidden = 0;
        }
        finally
        {
            Release(font);
            Release(selection);
        }
    }

    private void RestoreTypingCaretAt(
        Document document,
        int caretPosition,
        Range formulaRange)
    {
        Range? content = null;
        Selection? selection = null;
        try
        {
            content = document.Content;
            var safePosition = Math.Max(
                content.Start,
                Math.Min(caretPosition, content.End));
            selection = _application.Selection;
            selection.SetRange(safePosition, safePosition);
            ApplyInlineTypingFormattingToSelection(selection, formulaRange);
        }
        catch
        {
            // Structural caret placement remains useful even if Word rejects a
            // transient insertion-format mutation at an unusual protected range.
        }
        finally
        {
            Release(selection);
            Release(content);
        }
    }

    private void RestoreTypingBaselineAfter(Range formulaRange) =>
        RestoreTypingBaselineAfter(formulaRange, null);

    private void RestoreTypingBaselineAfter(Range formulaRange, int? caretPosition)
    {
        Range? caret = null;
        Selection? selection = null;
        try
        {
            ResetParagraphTypingPosition(formulaRange);
            caret = formulaRange.Duplicate;
            if (caretPosition.HasValue)
                caret.SetRange(caretPosition.Value, caretPosition.Value);
            else
                caret.Collapse(WdCollapseDirection.wdCollapseEnd);

            selection = _application.Selection;
            selection.SetRange(caret.Start, caret.End);
            ApplyInlineTypingFormattingToSelection(selection, formulaRange);
        }
        catch
        {
            // Keep the caret outside the formula even if Word refuses to mutate
            // insertion formatting at an unusual protected boundary.
        }
        finally
        {
            Release(selection);
            Release(caret);
        }
    }

    private static OfficeObjectResult Result(OfficeSessionDocument session, Document document) =>
        new()
        {
            FormulaId = session.FormulaId,
            DocumentId = DocumentIdentity(document),
            ObjectId = session.FormulaId,
        };

    private static string RangeReference(Range range) =>
        $"{RangeReferencePrefix}{range.Start}:{range.End}";

    private static FormulaMetadata? PreferSessionOmmlResolutionMetadata(
        FormulaMetadata? storedMetadata,
        FormulaMetadata? sessionSnapshot)
    {
        if (sessionSnapshot is not null
            && !string.IsNullOrWhiteSpace(sessionSnapshot.NativeOmmlFingerprint)
            && (storedMetadata is null
                || string.Equals(
                    storedMetadata.FormulaId,
                    sessionSnapshot.FormulaId,
                    StringComparison.OrdinalIgnoreCase)))
            return sessionSnapshot;
        return storedMetadata ?? sessionSnapshot;
    }

    private static Range ResolveOmmlEquationRange(
        Document document,
        Bookmark bookmark,
        string? sourceObjectId,
        FormulaMetadata? metadata)
    {
        try
        {
            return WordOmmlFormulaStore.GetEquationRange(bookmark);
        }
        catch (InvalidDataException)
        {
            Range? fallback = null;
            try
            {
                fallback = TryResolveOmmlRangeReference(document, sourceObjectId);
                if (fallback is null) throw;
                if (!string.IsNullOrWhiteSpace(metadata?.NativeOmmlFingerprint))
                {
                    string fingerprint;
                    try
                    {
                        fingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                            fallback.WordOpenXML);
                    }
                    catch { throw; }
                    if (!string.Equals(
                            fingerprint,
                            metadata!.NativeOmmlFingerprint,
                            StringComparison.OrdinalIgnoreCase))
                        throw;
                }
                var result = fallback;
                fallback = null;
                return result;
            }
            finally { Release(fallback); }
        }
    }

    private static Range ResolveStandaloneOmmlEquationRange(
        Document document,
        string? sourceObjectId,
        FormulaMetadata metadata)
    {
        Range? hinted = null;
        OMaths? maths = null;
        Range? matched = null;
        var matchCount = 0;
        try
        {
            hinted = TryResolveOmmlRangeReference(document, sourceObjectId);
            if (hinted is not null
                && OmmlFingerprintMatches(document, hinted, metadata))
            {
                var result = hinted;
                hinted = null;
                return result;
            }

            if (string.IsNullOrWhiteSpace(metadata.NativeOmmlFingerprint))
                throw new InvalidOperationException(
                    "The selected Word-native OMML formula could not be relocated.");

            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? candidate = null;
                try
                {
                    math = maths[index];
                    candidate = math.Range.Duplicate;
                    if (!OmmlFingerprintMatches(document, candidate, metadata))
                        continue;
                    matchCount++;
                    Release(matched);
                    matched = candidate;
                    candidate = null;
                }
                finally
                {
                    Release(candidate);
                    Release(math);
                }
            }
            if (matchCount != 1 || matched is null)
                throw new InvalidOperationException(
                    matchCount == 0
                        ? "The selected Word-native OMML formula no longer exists."
                        : "The Word document contains multiple identical native OMML formulas; please select the target again.");
            var unique = matched;
            matched = null;
            return unique;
        }
        finally
        {
            Release(matched);
            Release(maths);
            Release(hinted);
        }
    }

    private static bool OmmlFingerprintMatches(
        Document document,
        Range equationRange,
        FormulaMetadata metadata)
    {
        if (string.IsNullOrWhiteSpace(metadata.NativeOmmlFingerprint))
            return true;
        try
        {
            var xml = WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                document,
                equationRange,
                metadata.FormulaId);
            var fingerprint = WordOmmlConverter.ComputeOmmlFingerprint(xml);
            return string.Equals(
                fingerprint,
                metadata.NativeOmmlFingerprint,
                StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    private static Range? TryResolveOmmlRangeReference(
        Document document,
        string? sourceObjectId)
    {
        if (!TryParseRangeReference(sourceObjectId, out var start, out var end))
            return null;
        Range? content = null;
        Range? candidate = null;
        OMaths? maths = null;
        OMath? math = null;
        try
        {
            content = document.Content;
            if (start < content.Start || end < start || end > content.End)
                return null;
            object startValue = start;
            object endValue = end;
            candidate = document.Range(ref startValue, ref endValue);
            maths = candidate.OMaths;
            if (maths.Count != 1) return null;
            math = maths[1];
            var result = math.Range.Duplicate;
            return result;
        }
        catch { return null; }
        finally
        {
            Release(math);
            Release(maths);
            Release(candidate);
            Release(content);
        }
    }

    private static Range ResolveSessionInsertionRange(
        Document document,
        OfficeSessionDocument session,
        Selection selection)
    {
        var sourceRange = ResolveSourceRange(
            document,
            session.SourceObjectId,
            selection);
        if (!string.Equals(session.Mode, "create", StringComparison.OrdinalIgnoreCase))
            return sourceRange;
        try
        {
            return ResolveCreateInsertionRange(document, sourceRange);
        }
        finally { Release(sourceRange); }
    }

    private static Range ResolveCreateInsertionRange(
        Document document,
        Range sourceRange)
    {
        Tables? tables = null;
        Table? table = null;
        Range? safeTypingRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Range? captionRange = null;
        Range? content = null;
        try
        {
            if (!(bool)sourceRange.get_Information(WdInformation.wdWithInTable))
                return sourceRange.Duplicate;
            tables = sourceRange.Tables;
            if (tables.Count == 0) return sourceRange.Duplicate;
            table = tables[1];
            var formulaId = TryGetNumberedFormulaId(document, table);
            if (string.IsNullOrWhiteSpace(formulaId))
                return sourceRange.Duplicate;

            // When the captured caret belongs to a legacy VisualTeX numbered table,
            // place a new formula after that formula's dedicated native SEQ paragraph.
            // In Office 2019, using the live Selection at commit time can otherwise
            // insert new content between the old table and caption, reordering the
            // old number and corrupting adjacent cells.
            safeTypingRange =
                WordEquationNumbering.EnsureNormalTypingParagraphAfterNumberedDisplay(
                    document,
                    formulaId!);
            if (safeTypingRange is not null)
            {
                var result = safeTypingRange;
                safeTypingRange = null;
                return result;
            }

            bookmarks = document.Bookmarks;
            var captionName = WordEquationNumbering.NativeCaptionBookmarkName(
                formulaId!);
            if (!bookmarks.Exists(captionName))
                return sourceRange.Duplicate;
            captionBookmark = bookmarks[captionName];
            captionRange = captionBookmark.Range;
            content = document.Content;
            var lastInsertPosition = Math.Max(content.Start, content.End - 1);
            var position = Math.Max(
                content.Start,
                Math.Min(captionRange.End, lastInsertPosition));
            return document.Range(position, position);
        }
        catch
        {
            // Falling back to the captured range is safer than consulting the
            // mutable live Selection again. Normal document tables must retain
            // their existing in-cell insertion behavior.
            return sourceRange.Duplicate;
        }
        finally
        {
            Release(content);
            Release(captionRange);
            Release(captionBookmark);
            Release(bookmarks);
            Release(safeTypingRange);
            Release(table);
            Release(tables);
        }
    }

    private static string? TryGetNumberedFormulaId(
        Document document,
        Table table)
    {
        Range? tableRange = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Bookmark? ommlBookmark = null;
        try
        {
            tableRange = table.Range;
            shapes = tableRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                var metadata = WordFormulaMetadataReader.TryRead(shape);
                if (metadata?.Numbered == true
                    && string.Equals(
                        metadata.DisplayMode,
                        "block",
                        StringComparison.Ordinal))
                    return metadata.FormulaId;
            }

            ommlBookmark = WordOmmlFormulaStore.FindAtRange(document, tableRange);
            if (ommlBookmark is null) return null;
            var ommlMetadata = WordOmmlFormulaStore.TryRead(document, ommlBookmark);
            return ommlMetadata?.Numbered == true
                && string.Equals(
                    ommlMetadata.DisplayMode,
                    "block",
                    StringComparison.Ordinal)
                ? ommlMetadata.FormulaId
                : null;
        }
        finally
        {
            Release(ommlBookmark);
            Release(shape);
            Release(shapes);
            Release(tableRange);
        }
    }

    private static Range ResolveSourceRange(
        Document document,
        string? sourceObjectId,
        Selection selection)
    {
        if (!TryParseRangeReference(sourceObjectId, out var start, out var end))
            return selection.Range.Duplicate;
        Range? content = null;
        try
        {
            content = document.Content;
            if (start < 0 || end < start || end > content.End)
                throw new InvalidOperationException(
                    "The Word insertion range selected when the formula editor opened is no longer valid.");
            object startValue = start;
            object endValue = end;
            return document.Range(ref startValue, ref endValue);
        }
        finally { Release(content); }
    }

    private static bool TryParseRangeReference(
        string? value,
        out int start,
        out int end)
    {
        start = 0;
        end = 0;
        if (string.IsNullOrWhiteSpace(value)) return false;
        var reference = value!;
        if (!reference.StartsWith(RangeReferencePrefix, StringComparison.Ordinal))
            return false;
        var payload = reference.Substring(RangeReferencePrefix.Length);
        var separator = payload.IndexOf(':');
        if (separator <= 0 || separator >= payload.Length - 1) return false;
        return int.TryParse(payload.Substring(0, separator), out start)
            && int.TryParse(payload.Substring(separator + 1), out end);
    }

    private static void EnsureSourceDocument(
        Document document,
        string? expectedIdentity)
    {
        if (string.IsNullOrWhiteSpace(expectedIdentity)) return;
        var actual = DocumentIdentity(document);
        if (!string.Equals(actual, expectedIdentity, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "The active Word document changed while the VisualTeX editor was open.");
    }

    private static string DocumentIdentity(Document document)
    {
        try
        {
            var fullName = ReadWordStateWithRetry(() => document.FullName);
            if (!string.IsNullOrWhiteSpace(fullName)) return fullName;
        }
        catch { }
        return ReadWordStateWithRetry(() => document.Name);
    }

    private static void EnsureWritable(Document document)
    {
        if (ReadWordStateWithRetry(() => document.ReadOnly))
            throw new UnauthorizedAccessException("The active Word document is read-only.");
    }

    private static void EnsureEquationFieldResultsVisible(Document document)
    {
        Window? window = null;
        Microsoft.Office.Interop.Word.View? view = null;
        try
        {
            window = document.ActiveWindow;
            if (window is null) return;
            view = window.View;
            if (!view.ShowFieldCodes) return;

            // VisualTeX equation numbers and references are genuine Word fields.
            // When a Word window is left in Alt+F9/ShowFieldCodes mode, Word expands
            // the complete SEQ/MACROBUTTON instruction inside the numbered row,
            // making a healthy formula appear corrupted and stretching the layout.
            // Per-field ShowCodes=false cannot override the window-level setting,
            // so user-initiated VisualTeX numbering operations normalize the active
            // view back to rendered field results. The user can still press Alt+F9
            // afterwards if they intentionally want to inspect field instructions.
            view.ShowFieldCodes = false;
        }
        catch
        {
            // Hidden/protected automation windows can reject View mutations. The
            // formula operation itself remains valid; only interactive presentation
            // normalization is best-effort in that environment.
        }
        finally
        {
            Release(view);
            Release(window);
        }
    }

    private static T ReadWordStateWithRetry<T>(Func<T> read)
    {
        const int rpcCallRejected = unchecked((int)0x80010001);
        const int rpcServerCallRetryLater = unchecked((int)0x8001010A);
        const int officeBusy = unchecked((int)0x800AC472);
        const int maximumAttempts = 40;
        for (var attempt = 0; ; attempt++)
        {
            try { return read(); }
            catch (COMException error)
                when ((error.HResult == rpcCallRejected
                        || error.HResult == rpcServerCallRetryLater
                        || error.HResult == officeBusy)
                    && attempt < maximumAttempts - 1)
            {
                // Word can reject harmless state reads for a few UI turns while a
                // just-created document, OLE server or imported OMath is settling.
                // Retry only the idempotent property read; callers still execute
                // every document mutation exactly once.
                System.Threading.Thread.Sleep(50);
            }
        }
    }

    private static bool HasLeadingTab(Document document, Range formulaRange)
    {
        if (formulaRange.Start <= 0) return false;
        Range? preceding = null;
        try
        {
            object start = formulaRange.Start - 1;
            object end = formulaRange.Start;
            preceding = document.Range(ref start, ref end);
            if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal)) return true;
            if (!string.Equals(preceding.Text, "\v", StringComparison.Ordinal)
                || formulaRange.Start <= 1)
                return false;
            preceding.SetRange(formulaRange.Start - 2, formulaRange.Start - 1);
            return string.Equals(preceding.Text, "\t", StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
        finally { Release(preceding); }
    }

    private static void NormalizeNumberedDisplayCell(Range formulaRange)
    {
        Document? document = null;
        Table? table = null;
        Columns? columns = null;
        Cell? centerCell = null;
        Range? cellRange = null;
        Range? character = null;
        try
        {
            if (!(bool)formulaRange.get_Information(WdInformation.wdWithInTable)
                || formulaRange.Tables.Count == 0)
                return;
            document = formulaRange.Document;
            table = formulaRange.Tables[1];
            columns = table.Columns;
            if (columns.Count < 3) return;
            centerCell = table.Cell(1, 2);
            cellRange = centerCell.Range;

            // A display OMath inserted next to the source OLE can leave one
            // manual line break on each side. Delete only those exact control
            // characters, scanning backwards so Word's shifting ranges cannot
            // expand across and remove the replacement formula object.
            for (var position = cellRange.End - 2;
                 position >= cellRange.Start;
                 position--)
            {
                if (position >= formulaRange.Start
                    && position < formulaRange.End)
                    continue;
                object characterStart = position;
                object characterEnd = position + 1;
                character = document.Range(
                    ref characterStart,
                    ref characterEnd);
                if (string.Equals(character.Text, "\v", StringComparison.Ordinal))
                    character.Delete();
                Release(character);
                character = null;
            }
        }
        finally
        {
            Release(character);
            Release(cellRange);
            Release(centerCell);
            Release(columns);
            Release(table);
            Release(document);
        }
    }

    private static void NormalizeNumberedDisplayCell(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            NormalizeNumberedDisplayCell(range);
        }
        finally { Release(range); }
    }

    private static Range RestoreOmmlReplacementRollback(
        Document document,
        int position,
        string wordOpenXml)
    {
        Range? content = null;
        Range? insertion = null;
        Range? probe = null;
        OMaths? maths = null;
        Range? best = null;
        var bestDistance = int.MaxValue;
        try
        {
            content = document.Content;
            var safePosition = Math.Max(content.Start, Math.Min(position, content.End));
            object insertionStart = safePosition;
            object insertionEnd = Math.Min(
                content.End,
                safePosition + BulkInlineFormulaPlaceholder.Length);
            insertion = document.Range(ref insertionStart, ref insertionEnd);
            if (string.Equals(
                    insertion.Text,
                    BulkInlineFormulaPlaceholder,
                    StringComparison.Ordinal))
                insertion.Text = string.Empty;
            insertion.SetRange(safePosition, safePosition);
            insertion.InsertXML(wordOpenXml);

            object probeStart = Math.Max(content.Start, safePosition - 1);
            object probeEnd = Math.Min(document.Content.End, safePosition + 8);
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
                    var distance = Math.Abs(range.Start - safePosition);
                    if (distance >= bestDistance) continue;
                    Release(best);
                    best = range.Duplicate;
                    bestDistance = distance;
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
            if (best is null)
                throw new InvalidOperationException(
                    "Word could not restore the original OMML equation after a failed replacement.");
            var result = best;
            best = null;
            return result;
        }
        finally
        {
            Release(best);
            Release(maths);
            Release(probe);
            Release(insertion);
            Release(content);
        }
    }

    private static void ValidateInsertedOmml(Range equationRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        try
        {
            maths = equationRange.OMaths;
            if (maths.Count != 1)
                throw new InvalidOperationException(
                    "Word did not create exactly one native OMML equation.");
            math = maths[1];
            mathRange = math.Range;
            var wordOpenXml = mathRange.WordOpenXML;
            if (mathRange.End <= mathRange.Start
                || string.IsNullOrWhiteSpace(wordOpenXml)
                || wordOpenXml.IndexOf("oMath", StringComparison.Ordinal) < 0)
                throw new InvalidOperationException(
                    "Word returned an empty native OMML equation.");
            // Validate the structure after Word has imported and normalized the
            // OMML. Pre-insertion XML checks cannot catch empty slots that Word
            // introduces while materializing the native equation tree.
            WordOmmlConverter.ValidateMaterializedOmml(wordOpenXml);
        }
        finally
        {
            Release(mathRange);
            Release(math);
            Release(maths);
        }
    }

    private static void TryDelete(InlineShape? shape)
    {
        if (shape is null) return;
        try { shape.Delete(); } catch { }
    }

    private static void TryDelete(Table? table)
    {
        if (table is null) return;
        try { table.Delete(); } catch { }
    }

    private static void TryDelete(Bookmark? bookmark, bool deleteContents)
    {
        if (bookmark is null) return;
        Range? range = null;
        try
        {
            if (deleteContents) range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            bookmark.Delete();
            if (deleteContents) range?.Delete();
        }
        catch { }
        finally { Release(range); }
    }

    private static void TryDelete(Range? range)
    {
        if (range is null) return;
        try { range.Delete(); } catch { }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        // Office may return the same RCW to the host and to this service.
        // FinalReleaseComObject would invalidate every shared reference in the
        // add-in AppDomain, so release only the reference acquired here.
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
