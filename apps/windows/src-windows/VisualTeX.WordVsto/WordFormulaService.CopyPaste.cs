using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    internal sealed class WordFormulaCopyGroupItem
    {
        internal FormulaMetadata Metadata { get; set; } = new();
        internal WdStoryType SourceStoryType { get; set; }
        internal int SourceStart { get; set; }
        internal int SourceEnd { get; set; }
        internal WordCharacterFormatting? BodyFormatting { get; set; }
        internal string? VisibleNumber { get; set; }
    }

    private sealed class PastedNativeOleGroupCandidate
    {
        internal WordFormulaCopyGroupItem Item { get; set; } = new();
        internal InlineShape Shape { get; set; } = null!;
        internal int Start { get; set; }
    }

    internal sealed class WordFormulaCopySnapshot
    {
        internal string ObjectMode { get; set; } = string.Empty;
        internal FormulaMetadata Metadata { get; set; } = new();
        internal string SourceDocumentId { get; set; } = string.Empty;
        internal WdStoryType SourceStoryType { get; set; }
        internal int SourceStart { get; set; }
        internal int SourceEnd { get; set; }
        internal string MathTypeNumberPosition { get; set; } = "right";
        internal string? MathTypeContentSignature { get; set; }
        internal MathTypeWordOpenXml.NumberTemplate? MathTypeNumberTemplate { get; set; }
        internal string[] MathTypeNumberFieldResults { get; set; } = Array.Empty<string>();
        internal MathTypeDisplayParagraphLayout? MathTypeParagraphLayout { get; set; }
        internal int? OmmlNumberOrdinal { get; set; }
        internal string? OmmlNumberPrefix { get; set; }
        internal WordCharacterFormatting? BodyFormatting { get; set; }
        internal string TrackingDocumentId { get; set; } = string.Empty;
        internal int KnownInlineShapeCount { get; set; }
        internal int KnownOmmlCount { get; set; }
        internal int KnownDocumentEnd { get; set; }
        internal string? VisibleNumber { get; set; }
        internal List<WordFormulaCopyGroupItem> GroupItems { get; set; } = new();
        internal int CopiedSelectionStart { get; set; }
        internal int CopiedSelectionEnd { get; set; }
    }

    internal sealed class NumberedHostDeleteGuard
    {
        internal string DocumentId { get; set; } = string.Empty;
        internal string FormulaId { get; set; } = string.Empty;
        internal int SelectionStart { get; set; }
        internal int SelectionEnd { get; set; }
        internal int DocumentEnd { get; set; }
    }

    internal enum PastedFormulaRepairResult
    {
        NotReady,
        NotApplicable,
        Repaired,
    }

    internal WordFormulaCopySnapshot? CaptureSelectedFormulaForCopy()
    {
        using var perf = WordSelectionPerformance.Start("copy-capture");
        Document? document = null;
        Selection? selection = null;
        Range? selectionRange = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? formulaRange = null;
        Bookmark? ommlBookmark = null;
        Table? numberedTable = null;
        Cell? numberCell = null;
        Range? numberCellRange = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null) return null;
            selection = _application.Selection;
            selectionRange = selection.Range;
            ReadFormulaObjectCounts(document, out var inlineShapeCount, out var ommlCount);
            var documentId = DocumentIdentity(document);

            shapes = selectionRange.InlineShapes;
            if (shapes.Count > 1)
            {
                return CaptureSelectedNativeOleGroupForCopy(
                    document,
                    selectionRange,
                    inlineShapeCount,
                    ommlCount,
                    documentId);
            }
            if (shapes.Count == 1)
            {
                shape = shapes[1];
                formulaRange = shape.Range;
                if (WordFormulaMetadataReader.IsNativeOle(shape))
                {
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata is null) return null;
                    return new WordFormulaCopySnapshot
                    {
                        ObjectMode = FormulaOleContract.NativeOleMode,
                        Metadata = CloneFormulaMetadata(metadata),
                        SourceDocumentId = documentId,
                        SourceStoryType = formulaRange.StoryType,
                        SourceStart = formulaRange.Start,
                        SourceEnd = formulaRange.End,
                        BodyFormatting = TryCaptureParagraphMarkFormatting(formulaRange),
                        TrackingDocumentId = documentId,
                        KnownInlineShapeCount = inlineShapeCount,
                        KnownOmmlCount = ommlCount,
                        KnownDocumentEnd = document.Content.End,
                        VisibleNumber = metadata.Numbered ? ReadPastedOleVisibleNumber(shape, metadata.FormulaId) : null,
                    };
                }

                if (!MathTypeOleInterop.IsMathTypeOle(shape)) return null;
                var mathTypeBytes = MathTypeOleStorage.CaptureCompoundFile(shape);
                var mathTypeMathMl = MathTypeOleStorage.ReadMathMl(mathTypeBytes);
                var mathTypeMetadata = MathTypeOleInterop.ReadMetadata(_application, shape, mathTypeMathMl, mathTypeBytes);
                var numberPosition = "right";
                MathTypeWordOpenXml.NumberTemplate? numberTemplate = null;
                var numberFieldResults = Array.Empty<string>();
                if (mathTypeMetadata.Numbered)
                {
                    if (!MathTypeOleInterop.TryReadDisplayNumberPosition(shape, out numberPosition))
                        throw new InvalidDataException(
                            "The copied MathType equation has no readable number position.");
                    numberTemplate = ReadMathTypePlaceRefTemplateForShape(
                        document,
                        shape,
                        numberPosition);
                    numberFieldResults = CaptureMathTypePlaceRefFieldResults(
                        document,
                        shape,
                        numberPosition);
                }
                return new WordFormulaCopySnapshot
                {
                    ObjectMode = FormulaOleContract.MathTypeOleMode,
                    Metadata = CloneFormulaMetadata(mathTypeMetadata),
                    SourceDocumentId = documentId,
                    SourceStoryType = formulaRange.StoryType,
                    SourceStart = formulaRange.Start,
                    SourceEnd = formulaRange.End,
                    MathTypeNumberPosition = numberPosition,
                    MathTypeContentSignature = MathTypeMtefCodec.SemanticSignature(mathTypeMathMl),
                    MathTypeNumberTemplate = numberTemplate,
                    MathTypeNumberFieldResults = numberFieldResults,
                    MathTypeParagraphLayout = string.Equals(
                        mathTypeMetadata.DisplayMode,
                        "block",
                        StringComparison.Ordinal)
                        ? CaptureMathTypeDisplayParagraphLayout(shape)
                        : null,
                    BodyFormatting = TryCaptureParagraphMarkFormatting(formulaRange),
                    TrackingDocumentId = documentId,
                    KnownInlineShapeCount = inlineShapeCount,
                    KnownOmmlCount = ommlCount,
                    KnownDocumentEnd = document.Content.End,
                };
            }

            formulaRange = TryResolveNativeOmmlAtRange(document, selectionRange);
            if (formulaRange is null) return null;
            ommlBookmark = WordOmmlFormulaStore.FindAtRange(document, formulaRange);
            if (ommlBookmark is null) return null;
            var ommlMetadata = WordOmmlFormulaStore.TryRead(document, ommlBookmark);
            if (ommlMetadata is null) return null;

            WordCharacterFormatting? bodyFormatting = null;
            int? ommlNumberOrdinal = null;
            string? ommlNumberPrefix = null;
            if (ommlMetadata.Numbered)
            {
                if (WordEquationNumbering.TryReadManagedDirectTableNumberPlan(
                        document,
                        ommlMetadata.FormulaId,
                        out var sourceOrdinal,
                        out var sourcePrefix))
                {
                    ommlNumberOrdinal = sourceOrdinal;
                    ommlNumberPrefix = sourcePrefix;
                }
                numberedTable = WordEquationNumbering.FindNumberedEquationTable(
                    document,
                    ommlMetadata.FormulaId);
                if (numberedTable is not null)
                {
                    numberCell = numberedTable.Cell(1, 3);
                    numberCellRange = numberCell.Range;
                    bodyFormatting = TryCaptureParagraphMarkFormatting(numberCellRange);
                }
            }
            bodyFormatting ??= TryCaptureParagraphMarkFormatting(formulaRange);
            return new WordFormulaCopySnapshot
            {
                ObjectMode = FormulaOleContract.WordOmmlMode,
                Metadata = CloneFormulaMetadata(ommlMetadata),
                SourceDocumentId = documentId,
                SourceStoryType = formulaRange.StoryType,
                SourceStart = formulaRange.Start,
                SourceEnd = formulaRange.End,
                OmmlNumberOrdinal = ommlNumberOrdinal,
                OmmlNumberPrefix = ommlNumberPrefix,
                BodyFormatting = bodyFormatting,
                TrackingDocumentId = documentId,
                KnownInlineShapeCount = inlineShapeCount,
                KnownOmmlCount = ommlCount,
                KnownDocumentEnd = document.Content.End,
            };
        }
        finally
        {
            Release(numberCellRange);
            Release(numberCell);
            Release(numberedTable);
            Release(ommlBookmark);
            Release(formulaRange);
            Release(shape);
            Release(shapes);
            Release(selectionRange);
            Release(selection);
            Release(document);
        }
    }

    private static WordFormulaCopySnapshot? CaptureSelectedNativeOleGroupForCopy(
        Document document,
        Range selectionRange,
        int inlineShapeCount,
        int ommlCount,
        string documentId)
    {
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? formulaRange = null;
        Range? numberingOwner = null;
        var items = new List<WordFormulaCopyGroupItem>();
        var formulaIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            shapes = selectionRange.InlineShapes;
            if (shapes.Count <= 1) return null;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(numberingOwner);
                numberingOwner = null;
                Release(formulaRange);
                formulaRange = null;
                Release(shape);
                shape = shapes[index];
                if (!WordFormulaMetadataReader.IsNativeOle(shape)) return null;
                var metadata = WordFormulaMetadataReader.TryRead(shape);
                if (metadata is null
                    || !metadata.Numbered
                    || !string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal)
                    || !formulaIds.Add(metadata.FormulaId))
                    return null;
                formulaRange = shape.Range;
                if (!WordEquationNumbering.HasCompleteFormulaNumberingArtifacts(
                        document,
                        metadata.FormulaId)
                    || !WordEquationNumbering.FormulaRangeOwnsNumberingArtifacts(
                        document,
                        formulaRange,
                        metadata.FormulaId))
                    return null;
                numberingOwner = WordEquationNumbering.FindNumberingOwnerRange(
                    document,
                    metadata.FormulaId);
                if (numberingOwner is null
                    || numberingOwner.StoryType != selectionRange.StoryType
                    || selectionRange.Start > numberingOwner.Start
                    || selectionRange.End < numberingOwner.End)
                    return null;
                items.Add(new WordFormulaCopyGroupItem
                {
                    Metadata = CloneFormulaMetadata(metadata),
                    SourceStoryType = formulaRange.StoryType,
                    SourceStart = formulaRange.Start,
                    SourceEnd = formulaRange.End,
                    BodyFormatting = TryCaptureParagraphMarkFormatting(formulaRange),
                    VisibleNumber = ReadPastedOleVisibleNumber(shape, metadata.FormulaId),
                });
            }
            if (items.Count != shapes.Count || items.Count <= 1) return null;
            var first = items[0];
            return new WordFormulaCopySnapshot
            {
                ObjectMode = FormulaOleContract.NativeOleMode,
                Metadata = CloneFormulaMetadata(first.Metadata),
                SourceDocumentId = documentId,
                SourceStoryType = selectionRange.StoryType,
                SourceStart = selectionRange.Start,
                SourceEnd = selectionRange.End,
                BodyFormatting = first.BodyFormatting,
                TrackingDocumentId = documentId,
                KnownInlineShapeCount = inlineShapeCount,
                KnownOmmlCount = ommlCount,
                KnownDocumentEnd = document.Content.End,
                VisibleNumber = first.VisibleNumber,
                GroupItems = items,
                CopiedSelectionStart = selectionRange.Start,
                CopiedSelectionEnd = selectionRange.End,
            };
        }
        finally
        {
            Release(numberingOwner);
            Release(formulaRange);
            Release(shape);
            Release(shapes);
        }
    }

    internal NumberedHostDeleteGuard? CaptureNumberedHostDeleteGuard(Selection selection)
    {
        if (selection is null || selection.Start == selection.End) return null;
        Document? document = null;
        Range? selectionRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        Bookmarks? localBookmarks = null;
        Bookmark? localBookmark = null;
        Bookmarks? documentBookmarks = null;
        Bookmark? identityBookmark = null;
        Range? identityRange = null;
        try
        {
            document = selection.Document;
            selectionRange = selection.Range;
            paragraphs = selectionRange.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            // This guard is deliberately narrow: only a full numbered display
            // paragraph selection (including its paragraph mark) is eligible.
            // Partial formula/number selections do not exhibit Word's viewport jump
            // and must not participate in view restoration.
            if (selectionRange.Start != paragraphRange.Start
                || selectionRange.End != paragraphRange.End)
                return null;
            shapes = paragraphRange.InlineShapes;
            if (shapes.Count != 1) return null;
            shape = shapes[1];
            if (!WordFormulaMetadataReader.IsNativeOle(shape)) return null;
            shapeRange = shape.Range;

            string? formulaId = null;
            localBookmarks = paragraphRange.Bookmarks;
            for (var index = 1; index <= localBookmarks.Count; index++)
            {
                Release(localBookmark);
                localBookmark = localBookmarks[index];
                if (!WordEquationNumbering.TryFormulaIdFromEquationBookmark(
                        localBookmark.Name,
                        out var candidateId))
                    continue;
                if (formulaId is not null
                    && !string.Equals(formulaId, candidateId, StringComparison.OrdinalIgnoreCase))
                    return null;
                formulaId = candidateId;
            }
            if (string.IsNullOrWhiteSpace(formulaId)
                || !WordEquationNumbering.HasCompleteFormulaNumberingArtifacts(
                    document,
                    formulaId!)
                || !WordEquationNumbering.FormulaRangeOwnsNumberingArtifacts(
                    document,
                    shapeRange,
                    formulaId!))
                return null;

            documentBookmarks = document.Bookmarks;
            var identityName = WordFormulaMetadataReader.IdentityBookmarkName(formulaId!);
            if (!documentBookmarks.Exists(identityName)) return null;
            identityBookmark = documentBookmarks[identityName];
            identityRange = identityBookmark.Range;
            if (identityRange.StoryType != shapeRange.StoryType
                || identityRange.Start != shapeRange.Start
                || identityRange.End != shapeRange.End)
                return null;

            return new NumberedHostDeleteGuard
            {
                DocumentId = DocumentIdentity(document),
                FormulaId = formulaId!,
                SelectionStart = selectionRange.Start,
                SelectionEnd = selectionRange.End,
                DocumentEnd = document.Content.End,
            };
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(identityRange);
            Release(identityBookmark);
            Release(documentBookmarks);
            Release(localBookmark);
            Release(localBookmarks);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(selectionRange);
            Release(document);
        }
    }

    internal bool IsPendingNumberedHostDeletionTransition(
        Selection selection,
        NumberedHostDeleteGuard guard)
    {
        if (selection is null || guard is null) return false;
        Document? document = null;
        Bookmarks? bookmarks = null;
        try
        {
            document = selection.Document;
            if (!string.Equals(
                    DocumentIdentity(document),
                    guard.DocumentId,
                    StringComparison.OrdinalIgnoreCase)
                || selection.Start != selection.End
                || selection.Start != guard.SelectionStart
                || document.Content.End != guard.DocumentEnd)
                return false;
            bookmarks = document.Bookmarks;
            return bookmarks.Exists(
                WordFormulaMetadataReader.IdentityBookmarkName(guard.FormulaId));
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(bookmarks);
            Release(document);
        }
    }

    internal bool IsCompletedNumberedHostDeletion(
        Selection selection,
        NumberedHostDeleteGuard guard)
    {
        if (selection is null || guard is null) return false;
        Document? document = null;
        Bookmarks? bookmarks = null;
        try
        {
            document = selection.Document;
            if (!string.Equals(
                    DocumentIdentity(document),
                    guard.DocumentId,
                    StringComparison.OrdinalIgnoreCase)
                || selection.Start != selection.End
                || selection.Start != guard.SelectionStart
                || document.Content.End >= guard.DocumentEnd)
                return false;
            bookmarks = document.Bookmarks;
            // The visible paragraph deletion removes the VTO physical owner while
            // the hidden SEQ caption can legitimately survive until a later explicit
            // numbering refresh. That exact loss proves this was the guarded delete,
            // not an ordinary caret move or text replacement.
            return !bookmarks.Exists(
                WordFormulaMetadataReader.IdentityBookmarkName(guard.FormulaId));
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(bookmarks);
            Release(document);
        }
    }

    internal void TrackCopySnapshotDocument(WordFormulaCopySnapshot snapshot)
    {
        if (snapshot is null) return;
        Document? document = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null) return;
            var documentId = DocumentIdentity(document);
            if (string.Equals(
                    snapshot.TrackingDocumentId,
                    documentId,
                    StringComparison.OrdinalIgnoreCase))
                return;
            ReadFormulaObjectCounts(document, out var inlineShapeCount, out var ommlCount);
            snapshot.TrackingDocumentId = documentId;
            snapshot.KnownInlineShapeCount = inlineShapeCount;
            snapshot.KnownOmmlCount = ommlCount;
            WordDoubleClickHook.TraceMessage(
                $"copy-snapshot-target document={documentId} inline={inlineShapeCount} omml={ommlCount}");
        }
        finally { Release(document); }
    }

    internal void RefreshCopySnapshotAfterExplicitMutation(WordFormulaCopySnapshot snapshot)
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is not null) RefreshCopySnapshotCounts(document, snapshot);
        }
        finally { Release(document); }
    }

    internal int ReadCopyWatchDocumentEnd(WordFormulaCopySnapshot snapshot)
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null
                || !string.Equals(
                    DocumentIdentity(document),
                    snapshot.TrackingDocumentId,
                    StringComparison.OrdinalIgnoreCase))
                return -1;
            return document.Content.End;
        }
        catch { return -1; }
        finally { Release(document); }
    }

    // Delayed rendering of a native OLE clipboard can change its sequence on
    // Paste. Inventory growth only permits a bounded payload check; the actual
    // repair must still prove the copied FormulaId/content before any write.
    internal bool HasPotentialCopiedInsertion(WordFormulaCopySnapshot snapshot)
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null || !string.Equals(DocumentIdentity(document),
                    snapshot.TrackingDocumentId, StringComparison.OrdinalIgnoreCase)) return false;
            ReadFormulaObjectCounts(document, out var oleCount, out var mathCount);
            return snapshot.ObjectMode == FormulaOleContract.WordOmmlMode
                ? mathCount > snapshot.KnownOmmlCount
                    || (snapshot.Metadata.DisplayMode == "inline" && mathCount == snapshot.KnownOmmlCount
                        && document.Content.End > snapshot.KnownDocumentEnd)
                : oleCount > snapshot.KnownInlineShapeCount;
        }
        finally { Release(document); }
    }

    internal PastedFormulaRepairResult RepairPastedFormula(
        WordFormulaCopySnapshot snapshot)
    {
        if (snapshot is null) return PastedFormulaRepairResult.NotApplicable;
        using var perf = WordSelectionPerformance.Start("paste-repair");
        Document? document = null;
        Selection? selection = null;
        InlineShape? shape = null;
        Range? formulaRange = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null) return PastedFormulaRepairResult.NotReady;
            if (document.ReadOnly || document.ProtectionType != WdProtectionType.wdNoProtection)
                return PastedFormulaRepairResult.NotApplicable;
            var documentId = DocumentIdentity(document);
            if (!string.Equals(
                    snapshot.TrackingDocumentId,
                    documentId,
                    StringComparison.OrdinalIgnoreCase))
                return PastedFormulaRepairResult.NotApplicable;

            ReadFormulaObjectCounts(document, out var inlineShapeCount, out var ommlCount);
            var isOleMode = string.Equals(
                    snapshot.ObjectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal)
                || string.Equals(
                    snapshot.ObjectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal);
            if (inlineShapeCount < snapshot.KnownInlineShapeCount || ommlCount < snapshot.KnownOmmlCount)
            {
                // Deletion/Undo establishes a new baseline. Do not require the
                // next paste to exceed an old count from before that deletion.
                snapshot.KnownInlineShapeCount = inlineShapeCount;
                snapshot.KnownOmmlCount = ommlCount;
                return PastedFormulaRepairResult.NotApplicable;
            }
            if (isOleMode && inlineShapeCount <= snapshot.KnownInlineShapeCount)
                return PastedFormulaRepairResult.NotApplicable;
            if (string.Equals(
                    snapshot.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal)
                && ommlCount <= snapshot.KnownOmmlCount)
            {
                // Native Word merges directly adjacent inline OMath objects on
                // paste. Their count stays constant even though a complete copied
                // formula was inserted. Prove the two content fingerprints before
                // separating that exact local boundary; never infer it from size alone.
                if (snapshot.Metadata.DisplayMode == "inline"
                    && TryRepairMergedCopiedOmml(document, snapshot))
                {
                    RefreshCopySnapshotCounts(document, snapshot);
                    return PastedFormulaRepairResult.Repaired;
                }
                return PastedFormulaRepairResult.NotApplicable;
            }

            selection = _application.Selection;
            if (snapshot.GroupItems.Count > 1)
            {
                if (!string.Equals(
                        snapshot.ObjectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal))
                    return PastedFormulaRepairResult.NotApplicable;
                if (inlineShapeCount < snapshot.KnownInlineShapeCount + snapshot.GroupItems.Count)
                    return PastedFormulaRepairResult.NotReady;
                return RepairPastedNativeOleGroup(document, selection, snapshot);
            }
            if (string.Equals(
                    snapshot.ObjectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                shape = FindNearestPastedInlineShape(
                    document,
                    selection,
                    snapshot,
                    WordFormulaMetadataReader.IsNativeOle,
                    requireSourceFormulaId: true);
                if (shape is null) return PastedFormulaRepairResult.NotReady;
                formulaRange = shape.Range;
                if (snapshot.Metadata.DisplayMode == "block")
                    snapshot.BodyFormatting?.ApplyToParagraphMark(formulaRange);
                var metadata = CloneForPastedFormula(snapshot.Metadata);
                BindOleIdentityBookmark(shape, metadata.FormulaId);
                try { WordFormulaMetadataReader.Write(shape, metadata); }
                catch { WordFormulaMetadataReader.CacheMetadata(shape, metadata); }
                if (string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
                {
                    var copiedNumber = ReadPastedOleVisibleNumber(shape, snapshot.Metadata.FormulaId)
                        ?? snapshot.VisibleNumber;
                    RebindPastedOleVisibleNumber(document, shape, snapshot.Metadata.FormulaId, metadata.FormulaId);
                    perf?.Mark("identity-and-local-ref");
                    if (metadata.Numbered)
                    {
                        // Give the copy its own SEQ/REF owner, but preserve the
                        // displayed number. No global renumber/reference refresh
                        // belongs in Paste; Update Numbers performs it explicitly.
                        WordEquationNumbering.BuildFormulaNumberingScaffoldForConversion(
                            document, formulaRange, shape.Height, metadata,
                            plannedOrdinal: 1, plannedPrefix: string.Empty, deferFieldUpdate: true);
                        RestorePastedOleVisibleNumber(document, metadata.FormulaId, copiedNumber);
                    }
                    else
                        TryReconcileShape(document, shape, metadata, numberingOrderMayHaveChanged: false);
                    perf?.Mark("local-number-owner");
                    // Numbering inserts a leading TAB before a display OLE. Word can
                    // expand a just-created bookmark over that new character, so bind
                    // the durable VTO identity once more after the final scaffold is stable.
                    BindOleIdentityBookmark(shape, metadata.FormulaId);
                }
                else
                {
                    RemoveInlineBaselineSentinel(document, metadata.FormulaId);
                    var boundary = EnsureInlineBaselineSentinel(
                        formulaRange,
                        metadata.FormulaId);
                    if (selection.Start == formulaRange.End)
                    {
                        PositionSelectionAfterInlineTypingAnchor(
                            selection,
                            formulaRange,
                            boundary);
                        ApplyInlineTypingFormattingToSelection(selection, formulaRange);
                    }
                }
                RepairLocalCopiedOleIdentityBookmarks(shape);
                RefreshCopySnapshotCounts(document, snapshot);
                WordDoubleClickHook.TraceMessage(
                    $"copy-paste-repaired mode=nativeOle formulaId={metadata.FormulaId} "
                    + $"range={formulaRange.Start}:{formulaRange.End} numbered={metadata.Numbered}");
                return PastedFormulaRepairResult.Repaired;
            }

            if (string.Equals(
                    snapshot.ObjectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal))
            {
                shape = FindNearestPastedInlineShape(
                    document,
                    selection,
                    snapshot,
                    MathTypeOleInterop.IsMathTypeOle,
                    requireSourceFormulaId: false);
                if (shape is null) return PastedFormulaRepairResult.NotReady;
                formulaRange = shape.Range;
                if (snapshot.Metadata.DisplayMode == "block")
                    snapshot.BodyFormatting?.ApplyToParagraphMark(formulaRange);
                var metadata = snapshot.Metadata;
                if (string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
                {
                    // A complete native row is already independent after Word
                    // Paste. Rebuilding it needlessly destroys/recreates MTPlaceRef.
                    // OLE-only Paste, however, has no number field; build one local
                    // scaffold without updating the sequence, then restore the exact
                    // copied nested field results. Word can also leave a freshly-built
                    // MTPlaceRef in a stale code-visible runtime state even though
                    // Field.ShowCodes reports false, so normalize that presentation
                    // explicitly before accepting the row as healthy.
                    var hasCompleteScaffold = HasCompletePastedMathTypeScaffold(
                        document,
                        shape,
                        metadata.Numbered,
                        snapshot.MathTypeNumberPosition);
                    var presentationHealthy = !metadata.Numbered
                        || hasCompleteScaffold
                            && TryRestorePastedMathTypeNumberPresentation(
                                document,
                                shape,
                                snapshot.MathTypeNumberPosition,
                                snapshot.MathTypeNumberFieldResults);
                    if (!hasCompleteScaffold || !presentationHealthy)
                    {
                        RebuildMathTypeDisplayScaffold(
                            document, shape, metadata.Numbered,
                            snapshot.MathTypeNumberPosition, snapshot.MathTypeNumberTemplate,
                            updateNumberFields: false);
                        if (metadata.Numbered
                            && !TryRestorePastedMathTypeNumberPresentation(
                                document,
                                shape,
                                snapshot.MathTypeNumberPosition,
                                snapshot.MathTypeNumberFieldResults))
                            throw new InvalidDataException(
                                "The copied MathType equation number could not be restored without renumbering the document.");
                        if (snapshot.MathTypeParagraphLayout is not null)
                            RestoreMathTypeDisplayParagraphLayout(shape, snapshot.MathTypeParagraphLayout);
                    }
                    perf?.Mark("mathtype-preserve-number-scaffold");
                }
                else
                {
                    SetInlineOleWordPosition(
                        shape,
                        (int)Math.Round(ReadDefinedShapeFontPosition(shape) ?? 0f));
                    RestoreTypingBaselineAfter(shape);
                }
                RefreshCopySnapshotCounts(document, snapshot);
                WordDoubleClickHook.TraceMessage(
                    $"copy-paste-repaired mode=mathTypeOle range={formulaRange.Start}:{formulaRange.End} "
                    + $"numbered={metadata.Numbered} side={snapshot.MathTypeNumberPosition}");
                return PastedFormulaRepairResult.Repaired;
            }

            if (!string.Equals(
                    snapshot.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
                return PastedFormulaRepairResult.NotApplicable;

            formulaRange = FindNearestPastedOmmlRange(document, selection, snapshot);
            if (formulaRange is null) return PastedFormulaRepairResult.NotReady;
            if (!string.IsNullOrEmpty(snapshot.Metadata.NativeOmmlFingerprint)
                && WordOmmlConverter.ComputeOmmlFingerprint(formulaRange.WordOpenXML)
                    != snapshot.Metadata.NativeOmmlFingerprint)
            {
                // Pasting INSIDE Word's native math editor can splice a source
                // across existing atoms. A partial result is not an independent
                // copy and must never receive the full source's cached metadata.
                RefreshCopySnapshotCounts(document, snapshot);
                WordDoubleClickHook.TraceMessage("paste-repair-skipped reason=partial-native-math-input");
                return PastedFormulaRepairResult.NotApplicable;
            }
            // Word promotes inline math pasted into an empty paragraph to display
            // math. Preserve the real OMath.Type as well as the copied metadata.
            var typedRange = RestoreCopiedOmmlDisplayMode(formulaRange, snapshot.Metadata.DisplayMode);
            Release(formulaRange);
            formulaRange = typedRange;
            if (snapshot.Metadata.DisplayMode == "block")
                snapshot.BodyFormatting?.ApplyToParagraphMark(formulaRange);
            var pastedMetadata = CloneForPastedFormula(snapshot.Metadata);
            pastedMetadata.NativeOmmlFingerprint = null;
            WordOmmlNativeSource.StampFingerprint(pastedMetadata, formulaRange);
            var bookmark = WordOmmlFormulaStore.Wrap(
                document,
                formulaRange,
                pastedMetadata,
                replaceExisting: true);
            try
            {
                WordOmmlFormulaStore.Save(document, pastedMetadata);
                // Paste creates only this formula's independent host. Preserve the
                // copied number presentation and leave sequence reordering to the
                // explicit Update Numbers command, exactly like VisualTeX OLE and
                // MathType OLE. Current direct-table OMML carries an exact ordinal
                // + heading prefix snapshot; legacy inputs still build only a local
                // scaffold and never renumber the surrounding document here.
                TryReconcileOmml(
                    document,
                    bookmark,
                    formulaRange,
                    pastedMetadata,
                    numberingOrderMayHaveChanged: false,
                    reuseExistingNumberedTableFormatting: false,
                    numberingScaffoldOnly: pastedMetadata.Numbered,
                    plannedNumberOrdinal: snapshot.OmmlNumberOrdinal,
                    plannedNumberPrefix: snapshot.OmmlNumberPrefix);
            }
            finally { Release(bookmark); }

            Release(formulaRange);
            formulaRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                pastedMetadata.FormulaId,
                pastedMetadata);
            WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                pastedMetadata,
                formulaRange);
            WordOmmlFormulaStore.Save(document, pastedMetadata);
            RefreshCopySnapshotCounts(document, snapshot);
            WordDoubleClickHook.TraceMessage(
                $"copy-paste-repaired mode=wordOmml formulaId={pastedMetadata.FormulaId} "
                + $"range={formulaRange.Start}:{formulaRange.End} numbered={pastedMetadata.Numbered}");
            return PastedFormulaRepairResult.Repaired;
        }
        finally
        {
            Release(formulaRange);
            Release(shape);
            Release(selection);
            Release(document);
        }
    }

    private PastedFormulaRepairResult RepairPastedNativeOleGroup(
        Document document,
        Selection selection,
        WordFormulaCopySnapshot snapshot)
    {
        if (snapshot.GroupItems.Count <= 1)
            return PastedFormulaRepairResult.NotApplicable;
        Range? caretAnchor = null;
        var candidates = new List<PastedNativeOleGroupCandidate>();
        var repairedFormulaIds = new List<string>();
        try
        {
            caretAnchor = selection.Range.Duplicate;
            caretAnchor.SetRange(selection.End, selection.End);
            if (!TryCollectPastedNativeOleGroup(
                    document,
                    selection,
                    snapshot,
                    candidates,
                    out var pastedStart,
                    out var pastedEnd))
                return PastedFormulaRepairResult.NotReady;

            RemoveUnownedCopiedNativeCaptionParagraphs(
                document,
                pastedStart,
                pastedEnd);

            foreach (var candidate in candidates.OrderByDescending(item => item.Start))
            {
                Range? formulaRange = null;
                try
                {
                    formulaRange = candidate.Shape.Range;
                    candidate.Item.BodyFormatting?.ApplyToParagraphMark(formulaRange);
                    var sourceId = candidate.Item.Metadata.FormulaId;
                    var copiedNumber = ReadPastedOleVisibleNumber(candidate.Shape, sourceId)
                        ?? candidate.Item.VisibleNumber;
                    var metadata = CloneForPastedFormula(candidate.Item.Metadata);
                    BindOleIdentityBookmark(candidate.Shape, metadata.FormulaId);
                    try { WordFormulaMetadataReader.Write(candidate.Shape, metadata); }
                    catch { WordFormulaMetadataReader.CacheMetadata(candidate.Shape, metadata); }
                    RebindPastedOleVisibleNumber(
                        document,
                        candidate.Shape,
                        sourceId,
                        metadata.FormulaId);
                    WordEquationNumbering.BuildFormulaNumberingScaffoldForConversion(
                        document,
                        formulaRange,
                        candidate.Shape.Height,
                        metadata,
                        plannedOrdinal: 1,
                        plannedPrefix: string.Empty,
                        deferFieldUpdate: true);
                    RestorePastedOleVisibleNumber(document, metadata.FormulaId, copiedNumber);
                    BindOleIdentityBookmark(candidate.Shape, metadata.FormulaId);
                    RepairLocalCopiedOleIdentityBookmarks(candidate.Shape);
                    repairedFormulaIds.Add(metadata.FormulaId);
                    WordDoubleClickHook.TraceMessage(
                        $"copy-paste-group-item-repaired sourceFormulaId={sourceId} "
                        + $"formulaId={metadata.FormulaId} range={formulaRange.Start}:{formulaRange.End} "
                        + $"number={copiedNumber}");
                }
                finally { Release(formulaRange); }
            }

            RefreshCopySnapshotCounts(document, snapshot);
            var caret = caretAnchor.Start;
            foreach (var formulaId in repairedFormulaIds)
            {
                Range? caption = null;
                try
                {
                    caption = WordEquationNumbering.FindNativeEquationCaptionRange(
                        document,
                        formulaId);
                    if (caption is not null
                        && caption.StoryType == WdStoryType.wdMainTextStory)
                        caret = Math.Max(caret, caption.End);
                }
                finally { Release(caption); }
            }
            caret = Math.Max(document.Content.Start,
                Math.Min(caret, document.Content.End - 1));
            selection.SetRange(caret, caret);
            WordDoubleClickHook.TraceMessage(
                $"copy-paste-group-repaired formulas={candidates.Count} range={pastedStart}:{pastedEnd}");
            return PastedFormulaRepairResult.Repaired;
        }
        finally
        {
            foreach (var candidate in candidates)
                Release(candidate.Shape);
            Release(caretAnchor);
        }
    }

    private static bool TryCollectPastedNativeOleGroup(
        Document document,
        Selection selection,
        WordFormulaCopySnapshot snapshot,
        ICollection<PastedNativeOleGroupCandidate> output,
        out int pastedStart,
        out int pastedEnd)
    {
        pastedStart = -1;
        pastedEnd = -1;
        Range? content = null;
        Range? search = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        var best = new Dictionary<string, (InlineShape Shape, int Start, int End, int Distance)>(
            StringComparer.OrdinalIgnoreCase);
        try
        {
            var itemById = snapshot.GroupItems.ToDictionary(
                item => item.Metadata.FormulaId,
                item => item,
                StringComparer.OrdinalIgnoreCase);
            if (itemById.Count != snapshot.GroupItems.Count) return false;
            content = document.Content;
            var anchor = Math.Max(content.Start,
                Math.Min(selection.Start, Math.Max(content.Start, content.End - 1)));
            var copiedLength = Math.Max(1, snapshot.CopiedSelectionEnd - snapshot.CopiedSelectionStart);
            var radius = Math.Min(250_000, Math.Max(4096, copiedLength * 3 + 2048));
            var searchStart = Math.Max(content.Start, anchor - radius);
            var searchEnd = Math.Min(content.End, anchor + 1024);
            search = document.Range(searchStart, searchEnd);
            shapes = search.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shapeRange);
                shapeRange = null;
                Release(shape);
                shape = shapes[index];
                if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                var metadata = WordFormulaMetadataReader.TryReadCachedPreview(shape)
                    ?? WordFormulaMetadataReader.TryReadEmbeddedNativeOle(shape);
                if (metadata is null
                    || !itemById.TryGetValue(metadata.FormulaId, out var item))
                    continue;
                shapeRange = shape.Range;
                if (IsCurrentOriginalNativeOleGroupItem(document, shapeRange, item, snapshot))
                    continue;
                var distance = DistanceToRange(anchor, shapeRange.Start, shapeRange.End);
                if (best.TryGetValue(metadata.FormulaId, out var previous)
                    && previous.Distance <= distance)
                    continue;
                if (best.TryGetValue(metadata.FormulaId, out previous))
                    Release(previous.Shape);
                best[metadata.FormulaId] = (
                    shape,
                    shapeRange.Start,
                    shapeRange.End,
                    distance);
                shape = null;
            }
            if (best.Count != snapshot.GroupItems.Count) return false;

            foreach (var item in snapshot.GroupItems)
            {
                if (!best.TryGetValue(item.Metadata.FormulaId, out var match))
                    return false;
                output.Add(new PastedNativeOleGroupCandidate
                {
                    Item = item,
                    Shape = match.Shape,
                    Start = match.Start,
                });
                best.Remove(item.Metadata.FormulaId);

                Release(paragraphRange);
                paragraphRange = null;
                Release(paragraph);
                paragraph = null;
                Release(paragraphs);
                paragraphs = null;
                shapeRange = match.Shape.Range;
                paragraphs = shapeRange.Paragraphs;
                if (paragraphs.Count != 1) return false;
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
                pastedStart = pastedStart < 0
                    ? paragraphRange.Start
                    : Math.Min(pastedStart, paragraphRange.Start);
                pastedEnd = Math.Max(pastedEnd, paragraphRange.End);
            }
            return pastedStart >= 0 && pastedEnd > pastedStart;
        }
        finally
        {
            foreach (var orphan in best.Values)
                Release(orphan.Shape);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(search);
            Release(content);
        }
    }

    private static bool IsCurrentOriginalNativeOleGroupItem(
        Document document,
        Range candidate,
        WordFormulaCopyGroupItem item,
        WordFormulaCopySnapshot snapshot)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? owner = null;
        try
        {
            if (string.Equals(
                    DocumentIdentity(document),
                    snapshot.SourceDocumentId,
                    StringComparison.OrdinalIgnoreCase))
            {
                bookmarks = document.Bookmarks;
                var name = WordFormulaMetadataReader.IdentityBookmarkName(item.Metadata.FormulaId);
                if (bookmarks.Exists(name))
                {
                    bookmark = bookmarks[name];
                    owner = bookmark.Range;
                    if (owner.StoryType == candidate.StoryType
                        && owner.Start == candidate.Start
                        && owner.End == candidate.End)
                        return true;
                }
                return candidate.StoryType == item.SourceStoryType
                    && candidate.Start == item.SourceStart
                    && candidate.End == item.SourceEnd;
            }
            return false;
        }
        finally
        {
            Release(owner);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void RemoveUnownedCopiedNativeCaptionParagraphs(
        Document document,
        int start,
        int end)
    {
        Range? scan = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Frames? frames = null;
        Tables? tables = null;
        Bookmarks? bookmarks = null;
        var deletions = new List<(int Start, int End)>();
        try
        {
            scan = document.Range(
                Math.Max(document.Content.Start, start),
                Math.Min(document.Content.End, end));
            paragraphs = scan.Paragraphs;
            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Release(bookmarks); bookmarks = null;
                Release(tables); tables = null;
                Release(frames); frames = null;
                Release(code); code = null;
                Release(field); field = null;
                Release(fields); fields = null;
                Release(maths); maths = null;
                Release(shapes); shapes = null;
                Release(paragraphRange); paragraphRange = null;
                Release(paragraph); paragraph = paragraphs[index];
                paragraphRange = paragraph.Range;
                shapes = paragraphRange.InlineShapes;
                maths = paragraphRange.OMaths;
                fields = WordFormulaHost.GetLocalFields(paragraphRange);
                frames = paragraphRange.Frames;
                tables = paragraphRange.Tables;
                bookmarks = paragraphRange.Bookmarks;
                if (shapes.Count != 0
                    || maths.Count != 0
                    || fields.Count != 1
                    || frames.Count != 1
                    || tables.Count != 0
                    || bookmarks.Count != 0)
                    continue;
                field = fields[1];
                code = field.Code;
                var fieldCode = (code.Text ?? string.Empty).Trim();
                if (!fieldCode.StartsWith("SEQ VisualTeXEquation ", StringComparison.OrdinalIgnoreCase))
                    continue;
                deletions.Add((paragraphRange.Start, paragraphRange.End));
            }
        }
        finally
        {
            Release(bookmarks);
            Release(tables);
            Release(frames);
            Release(code);
            Release(field);
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(scan);
        }

        foreach (var deletion in deletions.OrderByDescending(item => item.Start))
        {
            Range? range = null;
            try
            {
                range = document.Range(deletion.Start, deletion.End);
                range.Delete();
            }
            finally { Release(range); }
        }
        if (deletions.Count > 0)
            WordDoubleClickHook.TraceMessage(
                $"copy-paste-group-caption-copies-removed count={deletions.Count} range={start}:{end}");
    }

    private static Range RestoreCopiedOmmlDisplayMode(Range range, string displayMode)
    {
        OMaths? maths = null;
        OMath? math = null;
        try
        {
            maths = range.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException("A copied OMML range must identify exactly one equation.");
            math = maths[1];
            var expected = displayMode == "inline" ? WdOMathType.wdOMathInline : WdOMathType.wdOMathDisplay;
            if (math.Type != expected) math.Type = expected;
            return math.Range;
        }
        finally { Release(math); Release(maths); }
    }

    private static void ReadFormulaObjectCounts(
        Document document,
        out int inlineShapeCount,
        out int ommlCount)
    {
        InlineShapes? shapes = null;
        OMaths? maths = null;
        try
        {
            shapes = document.InlineShapes;
            maths = document.OMaths;
            inlineShapeCount = shapes.Count;
            ommlCount = maths.Count;
        }
        finally
        {
            Release(maths);
            Release(shapes);
        }
    }

    private static void RefreshCopySnapshotCounts(
        Document document,
        WordFormulaCopySnapshot snapshot)
    {
        ReadFormulaObjectCounts(document, out var inlineShapeCount, out var ommlCount);
        snapshot.TrackingDocumentId = DocumentIdentity(document);
        snapshot.KnownInlineShapeCount = inlineShapeCount;
        snapshot.KnownOmmlCount = ommlCount;
        snapshot.KnownDocumentEnd = document.Content.End;
    }

    private static FormulaMetadata CloneFormulaMetadata(FormulaMetadata metadata) =>
        FormulaMetadataCodec.DeserializeJson(
            FormulaMetadataCodec.SerializeJson(metadata))
        ?? throw new InvalidOperationException("Unable to clone formula metadata for copy/paste.");

    private static FormulaMetadata CloneForPastedFormula(FormulaMetadata source)
    {
        var now = DateTimeOffset.UtcNow.ToString("O");
        var clone = WordFormulaMetadataReader.CloneWithFormulaId(
            source,
            Guid.NewGuid().ToString("D"));
        clone.Lines = clone.Lines
            .Select(line => new FormulaLine
            {
                Id = Guid.NewGuid().ToString("D"),
                Latex = line.Latex,
            })
            .ToList();
        clone.CreatedWithVersion = "1.2.7";
        clone.UpdatedWithVersion = "1.2.7";
        clone.CreatedAt = now;
        clone.UpdatedAt = now;
        clone.Validate();
        return clone;
    }

    private static WordCharacterFormatting? TryCaptureParagraphMarkFormatting(Range range)
    {
        try { return WordCharacterFormatting.CaptureParagraphMark(range); }
        catch { return null; }
    }

    private static Range GetPastedFormulaProbeRange(Document document, Range selection, Paragraph paragraph)
    {
        var current = paragraph.Range;
        Range? previousCharacter = null;
        Range? previousRange = null;
        Paragraphs? previousParagraphs = null;
        Paragraph? previousParagraph = null;
        Tables? tables = null;
        Table? table = null;
        try
        {
            // Whole-paragraph/table Paste puts the caret AFTER the pasted host.
            // Inspect only that immediately preceding host, never the whole document.
            if (selection.Start == current.Start && current.Start > 0)
            {
                previousCharacter = current.Duplicate;
                previousCharacter.SetRange(current.Start - 1, current.Start);
                tables = previousCharacter.Tables;
                if (tables.Count == 1)
                {
                    table = tables[1];
                    previousRange = table.Range;
                }
                else
                {
                    previousParagraphs = previousCharacter.Paragraphs;
                    previousParagraph = previousParagraphs[1];
                    previousRange = previousParagraph.Range;
                }
                current.SetRange(previousRange.Start, current.End);
            }
            var result = current;
            current = null!;
            return result;
        }
        finally
        {
            Release(table); Release(tables); Release(previousRange);
            Release(previousParagraph); Release(previousParagraphs);
            Release(previousCharacter); Release(current);
        }
    }

    private static string[] CaptureMathTypePlaceRefFieldResults(
        Document document,
        InlineShape shape,
        string side)
    {
        Range? shapeRange = null;
        Range? paragraphRange = null;
        Field? placeRef = null;
        Range? code = null;
        Fields? nestedFields = null;
        Field? nested = null;
        Range? result = null;
        try
        {
            shapeRange = shape.Range;
            paragraphRange = shapeRange.Paragraphs[1].Range;
            placeRef = FindMathTypePlaceRefFieldForShape(
                paragraphRange,
                shapeRange,
                string.Equals(side, "left", StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidDataException(
                    "The copied numbered MathType equation has no MTPlaceRef field.");
            if (!TryReadCompleteMathTypePlaceRefTemplate(document, placeRef, out _))
                throw new InvalidDataException(
                    "The copied MathType equation has an incomplete MTPlaceRef field tree.");
            code = placeRef.Code;
            nestedFields = code.Fields;
            var values = new string[nestedFields.Count];
            for (var index = 1; index <= nestedFields.Count; index++)
            {
                Release(result);
                result = null;
                Release(nested);
                nested = nestedFields[index];
                result = nested.Result;
                values[index - 1] = result.Text ?? string.Empty;
            }
            return values;
        }
        finally
        {
            Release(result);
            Release(nested);
            Release(nestedFields);
            Release(code);
            Release(placeRef);
            Release(paragraphRange);
            Release(shapeRange);
        }
    }

    private static bool TryRestorePastedMathTypeNumberPresentation(
        Document document,
        InlineShape shape,
        string side,
        IReadOnlyList<string> copiedFieldResults)
    {
        if (copiedFieldResults is null || copiedFieldResults.Count == 0)
            return false;
        Range? shapeRange = null;
        Range? paragraphRange = null;
        Field? placeRef = null;
        Range? code = null;
        Fields? nestedFields = null;
        Field? nested = null;
        Range? result = null;
        try
        {
            shapeRange = shape.Range;
            paragraphRange = shapeRange.Paragraphs[1].Range;
            placeRef = FindMathTypePlaceRefFieldForShape(
                paragraphRange,
                shapeRange,
                string.Equals(side, "left", StringComparison.OrdinalIgnoreCase));
            if (placeRef is null
                || !TryReadCompleteMathTypePlaceRefTemplate(document, placeRef, out _))
                return false;
            code = placeRef.Code;
            nestedFields = code.Fields;
            if (nestedFields.Count != copiedFieldResults.Count)
                return false;
            for (var index = 1; index <= nestedFields.Count; index++)
            {
                Release(result);
                result = null;
                Release(nested);
                nested = nestedFields[index];
                result = nested.Result;
                var expected = copiedFieldResults[index - 1] ?? string.Empty;
                if (!string.Equals(result.Text ?? string.Empty, expected, StringComparison.Ordinal))
                    result.Text = expected;
                try { nested.ShowCodes = false; } catch { }
            }

            var viewShowsCodes = false;
            try { viewShowsCodes = document.ActiveWindow.View.ShowFieldCodes; } catch { }
            if (viewShowsCodes)
                return true;

            // Word can materialize a newly-created MTPlaceRef with its runtime
            // code cache still visible even though Field.ShowCodes already reports
            // false. A real false -> true -> false transition refreshes that cache
            // without updating any SEQ field, so the copied number stays unchanged.
            try { placeRef.ShowCodes = true; } catch { return false; }
            try { placeRef.ShowCodes = false; } catch { return false; }

            Release(paragraphRange);
            paragraphRange = null;
            Release(shapeRange);
            shapeRange = shape.Range;
            paragraphRange = shapeRange.Paragraphs[1].Range;
            Range? visibleScaffold = null;
            try
            {
                visibleScaffold = string.Equals(side, "left", StringComparison.OrdinalIgnoreCase)
                    ? document.Range(paragraphRange.Start, shapeRange.Start)
                    : document.Range(
                        shapeRange.End,
                        Math.Max(shapeRange.End, paragraphRange.End - 1));
                var visibleText = visibleScaffold.Text ?? string.Empty;
                return visibleText.IndexOf('\u0013') < 0
                    && visibleText.IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0
                    && visibleText.IndexOf(
                        "SEQ MT",
                        StringComparison.OrdinalIgnoreCase) < 0;
            }
            finally { Release(visibleScaffold); }
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(result);
            Release(nested);
            Release(nestedFields);
            Release(code);
            Release(placeRef);
            Release(paragraphRange);
            Release(shapeRange);
        }
    }

    private static bool HasCompletePastedMathTypeScaffold(Document document, InlineShape shape, bool numbered, string side)
    {
        Range? shapeRange = null; Range? paragraph = null; Field? number = null; Range? prefix = null;
        try
        {
            shapeRange = shape.Range;
            paragraph = shapeRange.Paragraphs[1].Range;
            if (paragraph.InlineShapes.Count != 1) return false;
            if (numbered)
            {
                number = FindMathTypePlaceRefFieldForShape(paragraph, shapeRange, side == "left");
                return number is not null && TryReadCompleteMathTypePlaceRefTemplate(document, number, out _);
            }
            prefix = document.Range(paragraph.Start, shapeRange.Start);
            return prefix.Text == "\t";
        }
        finally { Release(prefix); Release(number); Release(paragraph); Release(shapeRange); }
    }

    private static string? ReadPastedOleVisibleNumber(InlineShape shape, string formulaId)
    {
        Range? range = null;
        Range? paragraph = null;
        Fields? fields = null;
        try
        {
            range = shape.Range;
            paragraph = range.Paragraphs[1].Range;
            fields = paragraph.Fields;
            var name = "VTEqNum_" + Guid.Parse(formulaId).ToString("N");
            for (var i = 1; i <= fields.Count; i++)
            {
                Field? field = null; Range? code = null; Range? result = null;
                try
                {
                    field = fields[i];
                    if (field.Type != WdFieldType.wdFieldRef) continue;
                    code = field.Code;
                    if (!(code.Text ?? string.Empty).Contains(name)) continue;
                    result = field.Result;
                    return result.Text;
                }
                finally { Release(result); Release(code); Release(field); }
            }
            return null;
        }
        finally { Release(fields); Release(paragraph); Release(range); }
    }

    private static void RestorePastedOleVisibleNumber(Document document, string formulaId, string? number)
    {
        if (number is null) return;
        Bookmarks? bookmarks = null; Bookmark? bookmark = null; Range? range = null; Fields? fields = null;
        try
        {
            bookmarks = document.Bookmarks;
            var suffix = Guid.Parse(formulaId).ToString("N");
            // Preserve the cached SEQ value as well as its visible REF. Otherwise
            // the next ordinary content edit refreshes REF from the seed ordinal 1
            // and loses the copied chapter prefix even without global renumbering.
            // The SEQ instruction stays live; explicit Update Numbers recalculates it.
            bookmark = bookmarks["VTEqCap_" + suffix];
            range = bookmark.Range;
            fields = range.Fields;
            var restoredCaption = false;
            for (var i = 1; i <= fields.Count; i++)
            {
                Field? field = null; Range? result = null; Bookmark? rebound = null;
                try
                {
                    field = fields[i];
                    if (field.Type != WdFieldType.wdFieldSequence) continue;
                    result = field.Result;
                    if (result.Text != number) result.Text = number;
                    rebound = bookmarks.Add("VTEqNum_" + suffix, result);
                    restoredCaption = true;
                    break;
                }
                finally { Release(rebound); Release(result); Release(field); }
            }
            if (!restoredCaption) throw new InvalidDataException("The copied formula has no independent SEQ owner.");
            Release(fields); fields = null; Release(range); range = null; Release(bookmark); bookmark = null;
            bookmark = bookmarks["VTEq_" + suffix];
            range = bookmark.Range;
            fields = range.Fields;
            for (var i = 1; i <= fields.Count; i++)
            {
                Field? field = null; Range? code = null; Range? result = null;
                try
                {
                    field = fields[i];
                    if (field.Type != WdFieldType.wdFieldRef) continue;
                    code = field.Code;
                    if (!(code.Text ?? string.Empty).Contains("VTEqNum_" + Guid.Parse(formulaId).ToString("N"))) continue;
                    result = field.Result;
                    if (result.Text != number) result.Text = number;
                    return;
                }
                finally { Release(result); Release(code); Release(field); }
            }
            throw new InvalidDataException("The copied formula has no independent visible number reference.");
        }
        finally { Release(fields); Release(range); Release(bookmark); Release(bookmarks); }
    }

    private static void RebindPastedOleVisibleNumber(Document document, InlineShape shape, string sourceId, string copiedId)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? numberRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            numberRange = document.Range(shapeRange.End, paragraphRange.End - 1);
            var text = numberRange.Text ?? string.Empty;
            if (!text.StartsWith("\t(", StringComparison.Ordinal) || !text.EndsWith(")", StringComparison.Ordinal)) return;
            fields = numberRange.Fields;
            if (fields.Count != 1) return;
            field = fields[1];
            code = field.Code;
            var originalName = "VTEqNum_" + Guid.Parse(sourceId).ToString("N");
            if (field.Type != WdFieldType.wdFieldRef
                || !(code.Text ?? string.Empty).Contains(originalName)) return;
            code.Text = code.Text.Replace(originalName, "VTEqNum_" + Guid.Parse(copiedId).ToString("N"));
            // Keep the copied field's formatting and delimiters. The reconciler
            // supplies a new hidden SEQ owner and updates this local REF afterwards.
            bookmarks = document.Bookmarks;
            bookmark = bookmarks.Add("VTEq_" + Guid.Parse(copiedId).ToString("N"), numberRange);
        }
        finally
        {
            Release(bookmark); Release(bookmarks); Release(code); Release(field); Release(fields);
            Release(numberRange); Release(paragraphRange); Release(paragraph); Release(paragraphs); Release(shapeRange);
        }
    }

    private static InlineShape? FindNearestPastedInlineShape(
        Document document,
        Selection selection,
        WordFormulaCopySnapshot snapshot,
        Func<InlineShape, bool> predicate,
        bool requireSourceFormulaId)
    {
        Range? selectionRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? shapes = null;
        InlineShape? candidate = null;
        InlineShape? best = null;
        Range? candidateRange = null;
        try
        {
            selectionRange = selection.Range;
            paragraphs = selectionRange.Paragraphs;
            if (paragraphs.Count == 0) return null;
            paragraph = paragraphs[1];
            paragraphRange = GetPastedFormulaProbeRange(document, selectionRange, paragraph);
            shapes = paragraphRange.InlineShapes;
            var anchor = selectionRange.Start;
            var bestDistance = int.MaxValue;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(candidateRange);
                candidateRange = null;
                Release(candidate);
                candidate = shapes[index];
                if (!predicate(candidate)) continue;
                candidateRange = candidate.Range;
                if (requireSourceFormulaId && IsCurrentOriginalManagedOle(document, candidateRange, snapshot)) continue;
                if (requireSourceFormulaId)
                {
                    // At an adjacent paste Word can expand the previous copy's
                    // VTO bookmark over BOTH objects. Its alias is not the new
                    // object's identity; inspect the physical payload/cache first.
                    var candidateMetadata = WordFormulaMetadataReader.TryReadCachedPreview(candidate)
                        ?? WordFormulaMetadataReader.TryReadEmbeddedNativeOle(candidate);
                    if (candidateMetadata is null
                        || !string.Equals(
                            candidateMetadata.FormulaId,
                            snapshot.Metadata.FormulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                }
                else
                {
                    // MathType has no persistent FormulaId. A newly pasted object
                    // may occupy the old source coordinates when pasted in front;
                    // compare its native content, not those now-stale coordinates.
                    var actualMathMl = MathTypeOleStorage.ReadMathMl(candidate);
                    if (!string.Equals(MathTypeMtefCodec.SemanticSignature(actualMathMl),
                            snapshot.MathTypeContentSignature, StringComparison.Ordinal))
                        continue;
                }
                var distance = DistanceToRange(anchor, candidateRange.Start, candidateRange.End);
                if ((distance > 96 && !(snapshot.Metadata.DisplayMode == "block" && shapes.Count == 1))
                    || distance >= bestDistance) continue;
                Release(best);
                best = candidate;
                candidate = null;
                bestDistance = distance;
            }
            var result = best;
            best = null;
            return result;
        }
        finally
        {
            Release(candidateRange);
            Release(candidate);
            Release(best);
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(selectionRange);
        }
    }

    private static void RepairLocalCopiedOleIdentityBookmarks(InlineShape pastedShape)
    {
        Range? shapeRange = null;
        Range? localRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        InlineShapes? shapes = null;
        Document? document = null;
        Bookmarks? bookmarks = null;
        var hosts = new Dictionary<string, List<InlineShape>>(StringComparer.OrdinalIgnoreCase);
        try
        {
            shapeRange = pastedShape.Range;
            document = shapeRange.Document;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            localRange = paragraph.Range;
            shapes = localRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = shapes[index];
                try
                {
                    if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                    var metadata = WordFormulaMetadataReader.TryReadCachedPreview(shape);
                    if (metadata is null) continue;
                    if (!hosts.TryGetValue(metadata.FormulaId, out var entries))
                        hosts.Add(metadata.FormulaId, entries = new List<InlineShape>());
                    entries.Add(shape);
                    shape = null;
                }
                finally { Release(shape); }
            }
            bookmarks = document.Bookmarks;
            foreach (var entry in hosts)
            {
                // A still-unrekeyed duplicate is ambiguous; never move its source
                // alias. Only a uniquely identified local physical owner qualifies.
                if (entry.Value.Count != 1) continue;
                var name = WordFormulaMetadataReader.IdentityBookmarkName(entry.Key);
                if (!bookmarks.Exists(name)) continue;
                Bookmark? bookmark = null;
                Range? owner = null;
                Range? candidate = null;
                try
                {
                    bookmark = bookmarks[name];
                    owner = bookmark.Range;
                    candidate = entry.Value[0].Range;
                    if (owner.StoryType != candidate.StoryType
                        || owner.Start > candidate.Start || owner.End < candidate.End
                        || (owner.Start == candidate.Start && owner.End == candidate.End))
                        continue;
                    BindOleIdentityBookmark(entry.Value[0], entry.Key);
                }
                finally { Release(candidate); Release(owner); Release(bookmark); }
            }
        }
        finally
        {
            foreach (var shape in hosts.Values.SelectMany(items => items)) Release(shape);
            Release(bookmarks); Release(document); Release(shapes);
            Release(localRange); Release(paragraph); Release(paragraphs); Release(shapeRange);
        }
    }

    private static bool IsCurrentOriginalManagedOle(
        Document document,
        Range candidate,
        WordFormulaCopySnapshot snapshot)
    {
        if (!string.Equals(DocumentIdentity(document), snapshot.SourceDocumentId, StringComparison.OrdinalIgnoreCase))
            return false;
        if (!string.Equals(
                snapshot.ObjectMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal))
            return IsOriginalCopySource(document, candidate, snapshot);

        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? owner = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = WordFormulaMetadataReader.IdentityBookmarkName(snapshot.Metadata.FormulaId);
            if (!bookmarks.Exists(name))
                return IsOriginalCopySource(document, candidate, snapshot);
            bookmark = bookmarks[name];
            owner = bookmark.Range;
            return owner.StoryType == candidate.StoryType
                && owner.Start == candidate.Start
                && owner.End == candidate.End;
        }
        catch
        {
            return IsOriginalCopySource(document, candidate, snapshot);
        }
        finally
        {
            Release(owner);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static Range? FindNearestPastedOmmlRange(
        Document document,
        Selection selection,
        WordFormulaCopySnapshot snapshot)
    {
        Range? selectionRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        Range? best = null;
        Bookmark? managed = null;
        try
        {
            selectionRange = selection.Range;
            paragraphs = selectionRange.Paragraphs;
            if (paragraphs.Count == 0) return null;
            paragraph = paragraphs[1];
            paragraphRange = GetPastedFormulaProbeRange(document, selectionRange, paragraph);
            maths = paragraphRange.OMaths;
            var anchor = selectionRange.Start;
            var bestDistance = int.MaxValue;
            for (var index = 1; index <= maths.Count; index++)
            {
                Release(managed);
                managed = null;
                Release(mathRange);
                mathRange = null;
                Release(math);
                math = maths[index];
                mathRange = math.Range;
                managed = WordOmmlFormulaStore.FindAtRange(document, mathRange);
                if (managed is not null
                    && string.Equals(DocumentIdentity(document), snapshot.SourceDocumentId, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (IsOriginalCopySource(document, mathRange, snapshot)) continue;
                var distance = DistanceToRange(anchor, mathRange.Start, mathRange.End);
                if ((distance > 96 && !(snapshot.Metadata.DisplayMode == "block" && maths.Count == 1))
                    || distance >= bestDistance) continue;
                Release(best);
                best = mathRange.Duplicate;
                bestDistance = distance;
            }
            var result = best;
            best = null;
            return result;
        }
        finally
        {
            Release(managed);
            Release(best);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(selectionRange);
        }
    }

    private static bool IsOriginalCopySource(
        Document document,
        Range candidate,
        WordFormulaCopySnapshot snapshot)
    {
        if (!string.Equals(
                DocumentIdentity(document),
                snapshot.SourceDocumentId,
                StringComparison.OrdinalIgnoreCase))
            return false;
        return candidate.StoryType == snapshot.SourceStoryType
            && candidate.Start == snapshot.SourceStart
            && candidate.End == snapshot.SourceEnd;
    }

    private static int DistanceToRange(int position, int start, int end)
    {
        if (position >= start && position <= end) return 0;
        return position < start ? start - position : position - end;
    }
}
