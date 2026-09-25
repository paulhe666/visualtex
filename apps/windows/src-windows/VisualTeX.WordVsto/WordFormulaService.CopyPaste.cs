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
        internal int BlankParagraphsBefore { get; set; }
        internal WordFormulaHostKind CoreHostKind { get; set; }
        internal string? CoreSemanticSignature { get; set; }
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
        internal WordFormulaHostKind? CoreHostKind { get; set; }
        internal string? CoreSemanticSignature { get; set; }
        internal string? CoreMathMl { get; set; }
        internal string? CoreLatex { get; set; }
        internal bool UsesHostCore { get; set; }
        internal bool CorePendingPasteArmed { get; set; }
        internal string CorePendingPasteDocumentId { get; set; } = string.Empty;
        internal WdStoryType CorePendingPasteStoryType { get; set; }
        internal int CorePendingPasteStart { get; set; }
        internal int CorePendingPasteEnd { get; set; }
        internal bool CorePendingOmmlMergeTarget { get; set; }
        internal int CorePendingOmmlTargetStart { get; set; }
        internal int CorePendingOmmlTargetEnd { get; set; }
        internal string? CorePendingOmmlTargetFormulaId { get; set; }
        internal string? CorePendingOmmlTargetSemanticSignature { get; set; }
        internal string? CorePendingOmmlTargetLatex { get; set; }
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
        using var perf =
            WordSelectionPerformance.Start("copy-capture");
        Document? document = null;
        Selection? selection = null;
        Range? selectionRange = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? formulaRange = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null)
                return null;

            selection = _application.Selection;
            selectionRange = selection.Range;

            if (TryCaptureHostCoreCopySnapshot(
                    document,
                    selectionRange,
                    out var coreSnapshot,
                    out var touchesHostCore))
                return coreSnapshot;
            if (touchesHostCore)
                return null;

            // OMML and VisualTeX are completely owned by CopyPasteCore above.
            // This legacy branch is intentionally MathType-only.
            shapes = selectionRange.InlineShapes;
            if (shapes.Count != 1)
                return null;

            shape = shapes[1];
            if (!MathTypeOleInterop.IsMathTypeOle(shape))
                return null;

            ReadFormulaObjectCounts(
                document,
                out var inlineShapeCount,
                out var ommlCount);
            var documentId =
                DocumentIdentity(document);
            formulaRange = shape.Range;

            var mathTypeBytes =
                MathTypeOleStorage.CaptureCompoundFile(
                    shape);
            var mathTypeMathMl =
                MathTypeOleStorage.ReadMathMl(
                    mathTypeBytes);
            var mathTypeMetadata =
                MathTypeOleInterop.ReadMetadata(
                    _application,
                    shape,
                    mathTypeMathMl,
                    mathTypeBytes);
            var numberPosition = "right";
            MathTypeWordOpenXml.NumberTemplate?
                numberTemplate = null;
            var numberFieldResults =
                Array.Empty<string>();

            if (mathTypeMetadata.Numbered)
            {
                if (!MathTypeOleInterop
                        .TryReadDisplayNumberPosition(
                            shape,
                            out numberPosition))
                    throw new InvalidDataException(
                        "The copied MathType equation has no readable number position.");

                numberTemplate =
                    ReadMathTypePlaceRefTemplateForShape(
                        document,
                        shape,
                        numberPosition);
                numberFieldResults =
                    CaptureMathTypePlaceRefFieldResults(
                        document,
                        shape,
                        numberPosition);
            }

            return new WordFormulaCopySnapshot
            {
                ObjectMode =
                    FormulaOleContract.MathTypeOleMode,
                Metadata =
                    CloneFormulaMetadata(
                        mathTypeMetadata),
                SourceDocumentId = documentId,
                SourceStoryType =
                    formulaRange.StoryType,
                SourceStart =
                    formulaRange.Start,
                SourceEnd =
                    formulaRange.End,
                MathTypeNumberPosition =
                    numberPosition,
                MathTypeContentSignature =
                    MathTypeMtefCodec.SemanticSignature(
                        mathTypeMathMl),
                MathTypeNumberTemplate =
                    numberTemplate,
                MathTypeNumberFieldResults =
                    numberFieldResults,
                MathTypeParagraphLayout =
                    string.Equals(
                        mathTypeMetadata.DisplayMode,
                        "block",
                        StringComparison.Ordinal)
                        ? CaptureMathTypeDisplayParagraphLayout(
                            shape)
                        : null,
                BodyFormatting =
                    TryCaptureParagraphMarkFormatting(
                        formulaRange),
                TrackingDocumentId =
                    documentId,
                KnownInlineShapeCount =
                    inlineShapeCount,
                KnownOmmlCount =
                    ommlCount,
                KnownDocumentEnd =
                    document.Content.End,
            };
        }
        finally
        {
            Release(formulaRange);
            Release(shape);
            Release(shapes);
            Release(selectionRange);
            Release(selection);
            Release(document);
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
            if (snapshot.UsesHostCore)
            {
                snapshot.TrackingDocumentId = documentId;
                snapshot.KnownDocumentEnd = document.Content.End;
                snapshot.CorePendingPasteArmed = false;
                snapshot.CorePendingPasteDocumentId = string.Empty;
                WordDoubleClickHook.TraceMessage(
                    $"copy-snapshot-core-target document={documentId}");
                return;
            }
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
            if (document is null) return;
            if (snapshot.UsesHostCore)
            {
                snapshot.TrackingDocumentId = DocumentIdentity(document);
                snapshot.KnownDocumentEnd = document.Content.End;
                snapshot.CorePendingPasteArmed = false;
                snapshot.CorePendingPasteDocumentId = string.Empty;
                return;
            }
            RefreshCopySnapshotCounts(document, snapshot);
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
        if (snapshot.UsesHostCore)
            return HasPotentialHostCoreCopiedInsertion(snapshot);

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
        if (snapshot is null)
            return PastedFormulaRepairResult.NotApplicable;
        if (snapshot.UsesHostCore)
            return RepairPastedFormulaWithHostCore(
                snapshot);

        // The non-core path is deliberately MathType-only. OMML/VisualTeX paste
        // never reaches inventory counts, nearest-object recovery, fingerprints,
        // or legacy numbering repair.
        if (!string.Equals(
                snapshot.ObjectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal))
            return PastedFormulaRepairResult.NotApplicable;

        using var perf =
            WordSelectionPerformance.Start(
                "paste-repair-mathtype");
        Document? document = null;
        Selection? selection = null;
        InlineShape? shape = null;
        Range? formulaRange = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null)
                return PastedFormulaRepairResult.NotReady;
            if (document.ReadOnly
                || document.ProtectionType !=
                    WdProtectionType.wdNoProtection)
                return PastedFormulaRepairResult.NotApplicable;
            if (!string.Equals(
                    snapshot.TrackingDocumentId,
                    DocumentIdentity(document),
                    StringComparison.OrdinalIgnoreCase))
                return PastedFormulaRepairResult.NotApplicable;

            ReadFormulaObjectCounts(
                document,
                out var inlineShapeCount,
                out var ommlCount);
            if (inlineShapeCount <
                    snapshot.KnownInlineShapeCount
                || ommlCount <
                    snapshot.KnownOmmlCount)
            {
                snapshot.KnownInlineShapeCount =
                    inlineShapeCount;
                snapshot.KnownOmmlCount =
                    ommlCount;
                return PastedFormulaRepairResult.NotApplicable;
            }
            if (inlineShapeCount <=
                snapshot.KnownInlineShapeCount)
                return PastedFormulaRepairResult.NotApplicable;

            selection = _application.Selection;
            shape = FindNearestPastedInlineShape(
                document,
                selection,
                snapshot,
                MathTypeOleInterop.IsMathTypeOle,
                requireSourceFormulaId: false);
            if (shape is null)
                return PastedFormulaRepairResult.NotReady;

            formulaRange = shape.Range;
            if (snapshot.Metadata.DisplayMode ==
                "block")
                snapshot.BodyFormatting?
                    .ApplyToParagraphMark(
                        formulaRange);

            var metadata = snapshot.Metadata;
            if (string.Equals(
                    metadata.DisplayMode,
                    "block",
                    StringComparison.Ordinal))
            {
                var hasCompleteScaffold =
                    HasCompletePastedMathTypeScaffold(
                        document,
                        shape,
                        metadata.Numbered,
                        snapshot.MathTypeNumberPosition);
                var presentationHealthy =
                    !metadata.Numbered
                    || hasCompleteScaffold
                    && TryRestorePastedMathTypeNumberPresentation(
                        document,
                        shape,
                        snapshot.MathTypeNumberPosition,
                        snapshot.MathTypeNumberFieldResults);

                if (!hasCompleteScaffold
                    || !presentationHealthy)
                {
                    RebuildMathTypeDisplayScaffold(
                        document,
                        shape,
                        metadata.Numbered,
                        snapshot.MathTypeNumberPosition,
                        snapshot.MathTypeNumberTemplate,
                        updateNumberFields: false);
                    if (metadata.Numbered
                        && !TryRestorePastedMathTypeNumberPresentation(
                            document,
                            shape,
                            snapshot.MathTypeNumberPosition,
                            snapshot.MathTypeNumberFieldResults))
                        throw new InvalidDataException(
                            "The copied MathType equation number could not be restored without renumbering the document.");

                    if (snapshot.MathTypeParagraphLayout
                        is not null)
                        RestoreMathTypeDisplayParagraphLayout(
                            shape,
                            snapshot.MathTypeParagraphLayout);
                }
                perf?.Mark(
                    "mathtype-preserve-number-scaffold");
            }
            else
            {
                SetInlineOleWordPosition(
                    shape,
                    (int)Math.Round(
                        ReadDefinedShapeFontPosition(
                            shape) ?? 0f));
                RestoreTypingBaselineAfter(
                    shape);
            }

            RefreshCopySnapshotCounts(
                document,
                snapshot);
            WordDoubleClickHook.TraceMessage(
                $"copy-paste-repaired mode=mathTypeOle "
                + $"range={formulaRange.Start}:{formulaRange.End} "
                + $"numbered={metadata.Numbered} "
                + $"side={snapshot.MathTypeNumberPosition}");
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

    private static InlineShape? FindNearestPastedInlineShape(
        Document document,
        Selection selection,
        WordFormulaCopySnapshot snapshot,
        Func<InlineShape, bool> predicate,
        bool requireSourceFormulaId)
    {
        // This compatibility locator is MathType-only. OMML/VisualTeX paste is
        // resolved from the exact pre-paste Selection window in CopyPasteCore.
        if (requireSourceFormulaId
            || !string.Equals(
                snapshot.ObjectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal))
            return null;

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
            if (paragraphs.Count == 0)
                return null;

            paragraph = paragraphs[1];
            paragraphRange =
                GetPastedFormulaProbeRange(
                    document,
                    selectionRange,
                    paragraph);
            shapes = paragraphRange.InlineShapes;

            var anchor = selectionRange.Start;
            var bestDistance = int.MaxValue;
            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                Release(candidateRange);
                candidateRange = null;
                Release(candidate);
                candidate = shapes[index];

                if (!predicate(candidate))
                    continue;

                var actualMathMl =
                    MathTypeOleStorage.ReadMathMl(
                        candidate);
                if (!string.Equals(
                        MathTypeMtefCodec.SemanticSignature(
                            actualMathMl),
                        snapshot.MathTypeContentSignature,
                        StringComparison.Ordinal))
                    continue;

                candidateRange = candidate.Range;
                var distance =
                    DistanceToRange(
                        anchor,
                        candidateRange.Start,
                        candidateRange.End);
                if ((distance > 96
                     && !(snapshot.Metadata.DisplayMode == "block"
                          && shapes.Count == 1))
                    || distance >= bestDistance)
                    continue;

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

    private static int DistanceToRange(
        int position,
        int start,
        int end)
    {
        if (position >= start && position <= end)
            return 0;
        return position < start
            ? start - position
            : position - end;
    }
}
