using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    internal sealed class WordFormulaCopySnapshot
    {
        internal string ObjectMode { get; set; } = string.Empty;
        internal FormulaMetadata Metadata { get; set; } = new();
        internal string SourceDocumentId { get; set; } = string.Empty;
        internal WdStoryType SourceStoryType { get; set; }
        internal int SourceStart { get; set; }
        internal int SourceEnd { get; set; }
        internal string MathTypeNumberPosition { get; set; } = "right";
        internal MathTypeWordOpenXml.NumberTemplate? MathTypeNumberTemplate { get; set; }
        internal MathTypeDisplayParagraphLayout? MathTypeParagraphLayout { get; set; }
        internal WordCharacterFormatting? BodyFormatting { get; set; }
        internal string TrackingDocumentId { get; set; } = string.Empty;
        internal int KnownInlineShapeCount { get; set; }
        internal int KnownOmmlCount { get; set; }
    }

    internal enum PastedFormulaRepairResult
    {
        NotReady,
        NotApplicable,
        Repaired,
    }

    internal WordFormulaCopySnapshot? CaptureSelectedFormulaForCopy()
    {
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
                    };
                }

                if (!MathTypeOleInterop.IsMathTypeOle(shape)) return null;
                var mathTypeMetadata = MathTypeOleInterop.ReadMetadata(_application, shape);
                var numberPosition = "right";
                MathTypeWordOpenXml.NumberTemplate? numberTemplate = null;
                if (mathTypeMetadata.Numbered)
                {
                    if (!MathTypeOleInterop.TryReadDisplayNumberPosition(shape, out numberPosition))
                        throw new InvalidDataException(
                            "The copied MathType equation has no readable number position.");
                    numberTemplate = ReadMathTypePlaceRefTemplateForShape(
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
                    MathTypeNumberTemplate = numberTemplate,
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
                };
            }

            formulaRange = TryResolveNativeOmmlAtRange(document, selectionRange);
            if (formulaRange is null) return null;
            ommlBookmark = WordOmmlFormulaStore.FindAtRange(document, formulaRange);
            if (ommlBookmark is null) return null;
            var ommlMetadata = WordOmmlFormulaStore.TryRead(document, ommlBookmark);
            if (ommlMetadata is null) return null;

            WordCharacterFormatting? bodyFormatting = null;
            if (ommlMetadata.Numbered)
            {
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
                BodyFormatting = bodyFormatting,
                TrackingDocumentId = documentId,
                KnownInlineShapeCount = inlineShapeCount,
                KnownOmmlCount = ommlCount,
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

    internal PastedFormulaRepairResult RepairPastedFormula(
        WordFormulaCopySnapshot snapshot)
    {
        if (snapshot is null) return PastedFormulaRepairResult.NotApplicable;
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
                    RebindPastedOleVisibleNumber(document, shape, snapshot.Metadata.FormulaId, metadata.FormulaId);
                    TryReconcileShape(
                        document,
                        shape,
                        metadata,
                        numberingOrderMayHaveChanged: metadata.Numbered,
                        reuseExistingNumberedTableFormatting: false);
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
                    // A whole-row paste already includes MTPlaceRef and tabs.
                    // Rebuild that copied scaffold, rather than append a second number.
                    RebuildMathTypeDisplayScaffold(
                        document,
                        shape,
                        metadata.Numbered,
                        snapshot.MathTypeNumberPosition,
                        snapshot.MathTypeNumberTemplate);
                    if (snapshot.MathTypeParagraphLayout is not null)
                        RestoreMathTypeDisplayParagraphLayout(
                            shape,
                            snapshot.MathTypeParagraphLayout);
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
                TryReconcileOmml(
                    document,
                    bookmark,
                    formulaRange,
                    pastedMetadata,
                    numberingOrderMayHaveChanged: pastedMetadata.Numbered,
                    reuseExistingNumberedTableFormatting: false);
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
        clone.CreatedWithVersion = "1.2.6";
        clone.UpdatedWithVersion = "1.2.6";
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
                if (IsCurrentOriginalManagedOle(document, candidateRange, snapshot)) continue;
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
