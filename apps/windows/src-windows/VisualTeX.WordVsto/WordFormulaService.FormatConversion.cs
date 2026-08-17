using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    public WordFormulaFormatConversionPlan CaptureFormulaFormatConversionPlan(
        bool wholeDocument,
        string sourceMode,
        string targetMode)
    {
        ValidateSimpleFormatConversionPair(sourceMode, targetMode);
        Document? document = null;
        Selection? selection = null;
        Range? scope = null;
        InlineShapes? shapes = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            selection = _application.Selection;
            scope = wholeDocument
                ? document.Content.Duplicate
                : selection.Range.Duplicate;

            var plan = new WordFormulaFormatConversionPlan
            {
                DocumentId = DocumentIdentity(document),
                SourceMode = sourceMode,
                TargetMode = targetMode,
                WholeDocument = wholeDocument,
            };

            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                try
                {
                    shape = shapes[index];
                    FormulaMetadata? metadata = null;
                    var mathTypeNumberPosition = "right";
                    string sourceFormulaId;

                    if (string.Equals(
                            sourceMode,
                            FormulaOleContract.NativeOleMode,
                            StringComparison.Ordinal))
                    {
                        if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                        metadata = WordFormulaMetadataReader.TryRead(shape);
                        if (metadata is null) continue;
                        sourceFormulaId = metadata.FormulaId;
                    }
                    else
                    {
                        if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                        var mathMl = MathTypeOleStorage.ReadMathMl(shape);
                        metadata = MathTypeOleInterop.ReadMetadata(
                            _application,
                            shape,
                            mathMl);
                        if (MathTypeOleInterop.TryReadDisplayNumberPosition(
                                shape,
                                out var detectedPosition))
                            mathTypeNumberPosition = detectedPosition;
                        sourceFormulaId = metadata.FormulaId;
                    }

                    range = shape.Range;
                    if (!FormulaRangeMatchesScope(range, scope, wholeDocument))
                        continue;

                    var latex = string.IsNullOrWhiteSpace(metadata.Latex)
                        ? string.Join("\n", metadata.Lines.Select(line => line.Latex))
                        : metadata.Latex;
                    latex = (latex ?? string.Empty).Trim();
                    if (latex.Length == 0)
                        throw new InvalidDataException(
                            "A source formula has no recoverable LaTeX and was not converted.");

                    plan.Targets.Add(new WordFormulaFormatConversionTarget
                    {
                        Id = Guid.NewGuid().ToString("D"),
                        SourceFormulaId = sourceFormulaId,
                        SourceObjectId = $"{RangeReferencePrefix}{range.Start}:{range.End}",
                        SourceStart = range.Start,
                        Latex = latex,
                        DisplayMode = metadata.DisplayMode,
                        Numbered = metadata.Numbered,
                        MathTypeNumberPosition = mathTypeNumberPosition,
                        FontSizePt = FormulaFontSize.Normalize(metadata.FontSizePt),
                        Metadata = metadata,
                    });
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }

            return plan;
        }
        finally
        {
            Release(shapes);
            Release(scope);
            Release(selection);
            Release(document);
        }
    }

    public WordFormulaFormatConversionResult ApplyFormulaFormatConversionPlan(
        WordFormulaFormatConversionPlan plan,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));
        if (prepared is null) throw new ArgumentNullException(nameof(prepared));
        ValidateSimpleFormatConversionPair(plan.SourceMode, plan.TargetMode);

        Document? document = null;
        Document? visualTeXRollbackBuffer = null;
        Selection? selection = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, plan.DocumentId);
            selection = _application.Selection;
            if (string.Equals(
                    plan.SourceMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                visualTeXRollbackBuffer = _application.Documents.Add(Visible: false);
                document.Activate();
            }

            foreach (var target in plan.Targets)
            {
                if (!prepared.TryGetValue(target.Id, out var formula))
                    throw new InvalidDataException(
                        $"Missing rendered payload for formula '{target.Latex}'.");
                ValidatePreparedFormatConversionTarget(plan.TargetMode, target, formula);
                ValidateSimpleSourceHost(document, plan.SourceMode, target);
            }

            var result = new WordFormulaFormatConversionResult();
            var initialTargetObjectCount = CountSimpleFormatObjects(document, plan.TargetMode);
            var traceObjectCounts = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_COUNTS"),
                "1",
                StringComparison.Ordinal);
            foreach (var target in plan.Targets.OrderByDescending(item => item.SourceStart))
            {
                UndoRecord? formulaUndoRecord = null;
                string? createdTargetBookmarkName = null;
                VisualTeXRollbackSnapshot? visualTeXRollbackSnapshot = null;
                var undoRecordEnded = false;
                var mutationStarted = false;
                var stage = "capture-rollback-snapshot";
                try
                {
                    if (visualTeXRollbackBuffer is not null)
                    {
                        visualTeXRollbackSnapshot = CaptureVisualTeXRollbackSnapshot(
                            document,
                            visualTeXRollbackBuffer,
                            target);
                        document.Activate();
                    }
                    stage = "begin-undo";
                    formulaUndoRecord = BeginUndoRecord("VisualTeX Convert Formula Format");
                    if (formulaUndoRecord is null)
                        throw new InvalidOperationException(
                            "Word 无法建立单公式转换撤销事务。为避免转换失败时丢失原公式，本次转换已停止。");

                    var formula = prepared[target.Id];
                    var targetObjectCountBefore = CountSimpleFormatObjects(
                        document,
                        plan.TargetMode);
                    if (traceObjectCounts)
                        TraceSimpleFormatObjectCounts(document, target, "before-delete");
                    mutationStarted = true;
                    stage = "delete-source";
                    var insertionStart = DeleteSimpleSourceHost(
                        document,
                        plan.SourceMode,
                        target);
                    EnsureSimpleFormatSourceRemoved(document, plan.SourceMode, target, "delete-source");
                    if (traceObjectCounts)
                        TraceSimpleFormatObjectCounts(document, target, "after-delete");
                    ThrowIfSimpleFormatConversionFailureInjected(target);
                    var content = document.Content;
                    try
                    {
                        insertionStart = Math.Max(
                            content.Start,
                            Math.Min(insertionStart, content.End));
                    }
                    finally { Release(content); }

                    stage = "prepare-target-selection";
                    selection.SetRange(insertionStart, insertionStart);
                    selection.Collapse(WdCollapseDirection.wdCollapseStart);

                    var session = formula.Session;
                    session.Mode = "create";
                    session.SourceDocumentId = plan.DocumentId;
                    session.SourceObjectId =
                        $"{RangeReferencePrefix}{insertionStart}:{insertionStart}";
                    session.DisplayMode = target.DisplayMode;
                    session.ObjectMode = plan.TargetMode;
                    session.Numbered = target.Numbered;
                    session.MathTypeNumberPosition = target.MathTypeNumberPosition;
                    session.FontSizePt = target.FontSizePt;
                    session.OriginalMetadata = null;

                    stage = "insert-target";
                    if (string.Equals(
                            plan.TargetMode,
                            FormulaOleContract.MathTypeOleMode,
                            StringComparison.Ordinal))
                    {
                        createdTargetBookmarkName =
                            "VTMT_" + target.Id.Replace("-", string.Empty);
                        InsertMathTypeOle(
                            session,
                            formula.MathMl!,
                            formula.EmfPath!,
                            createdTargetBookmarkName);
                    }
                    else
                    {
                        InsertOle(
                            session,
                            formula.PngPath!,
                            formula.EmfPath!);
                    }
                    EnsureSimpleFormatSourceRemoved(document, plan.SourceMode, target, "insert-target");
                    if (traceObjectCounts)
                        TraceSimpleFormatObjectCounts(document, target, "after-insert-before-commit");
                    stage = "commit-undo";
                    EndUndoRecord(formulaUndoRecord);
                    undoRecordEnded = true;
                    stage = "verify-target";
                    if (!string.IsNullOrWhiteSpace(createdTargetBookmarkName))
                    {
                        EnsureBookmarkedMathTypeTargetSurvived(
                            document,
                            createdTargetBookmarkName,
                            target);
                    }
                    else
                    {
                        WaitForSimpleTargetObjectCountToStabilize(
                            document,
                            plan.TargetMode,
                            targetObjectCountBefore + 1,
                            target);
                    }
                    EnsureSimpleFormatSourceRemoved(
                        document,
                        plan.SourceMode,
                        target,
                        "post-transaction");
                    if (traceObjectCounts)
                        TraceSimpleFormatObjectCounts(document, target, "after-commit-stable");
                    result.FormulaCount++;
                }
                catch (Exception error)
                {
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-item-failed stage={stage} formulaId={target.SourceFormulaId} latex={target.Latex} error={error}");
                    if (!undoRecordEnded)
                    {
                        EndUndoRecord(formulaUndoRecord);
                        undoRecordEnded = true;
                    }
                    if (mutationStarted)
                    {
                        if (!TryUndoFormulaToLatexConversion(document))
                            throw new InvalidOperationException(
                                $"公式“{target.Latex}”转换失败，而且 Word 无法自动恢复原公式。请立即停止编辑当前文档。",
                                error);
                        try
                        {
                            ValidateSimpleSourceHost(document, plan.SourceMode, target);
                        }
                        catch (Exception restoreError)
                        {
                            if (visualTeXRollbackSnapshot is null)
                                throw new InvalidOperationException(
                                    $"公式“{target.Latex}”转换失败；Word 执行了撤销，但原公式宿主没有完整恢复。",
                                    new AggregateException(error, restoreError));
                            try
                            {
                                RestoreVisualTeXRollbackSnapshot(
                                    document,
                                    visualTeXRollbackSnapshot,
                                    target);
                                ValidateSimpleSourceHost(document, plan.SourceMode, target);
                            }
                            catch (Exception snapshotRestoreError)
                            {
                                WordDoubleClickHook.TraceMessage(
                                    $"format-conversion-snapshot-restore-failed formulaId={target.SourceFormulaId} error={snapshotRestoreError}");
                                throw new InvalidOperationException(
                                    $"公式“{target.Latex}”转换失败；Word 撤销和 VisualTeX 结构快照恢复都未能完整恢复原公式宿主。",
                                    new AggregateException(error, restoreError, snapshotRestoreError));
                            }
                        }
                    }
                    result.FailedFormulaCount++;
                    result.Failures.Add($"{target.Latex}: [{stage}] {error.Message}");
                    break;
                }
                finally
                {
                    if (document is not null
                        && !string.IsNullOrWhiteSpace(createdTargetBookmarkName))
                        TryDeleteBookmark(document, createdTargetBookmarkName);
                    if (visualTeXRollbackSnapshot is not null)
                        Release(visualTeXRollbackSnapshot.Payload);
                    if (!undoRecordEnded) EndUndoRecord(formulaUndoRecord);
                    Release(formulaUndoRecord);
                }
            }

            if (result.FailedFormulaCount == 0)
            {
                foreach (var target in plan.Targets)
                {
                    if (!IsSimpleFormatSourcePresent(document, plan.SourceMode, target))
                        continue;
                    result.FailedFormulaCount++;
                    result.FormulaCount = Math.Max(0, result.FormulaCount - 1);
                    result.Failures.Add(
                        $"{target.Latex}: Word restored the source formula after the conversion transaction completed.");
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-source-reappeared formulaId={target.SourceFormulaId} latex={target.Latex}");
                    break;
                }
            }
            if (result.FailedFormulaCount == 0)
            {
                var finalTargetObjectCount = CountSimpleFormatObjects(document, plan.TargetMode);
                var expectedTargetObjectCount = initialTargetObjectCount + result.FormulaCount;
                if (finalTargetObjectCount != expectedTargetObjectCount)
                {
                    result.FailedFormulaCount++;
                    result.Failures.Add(
                        $"Target formula count mismatch after conversion: expected {expectedTargetObjectCount}, actual {finalTargetObjectCount}. Word removed or failed to retain a converted formula.");
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-target-count-mismatch targetMode={plan.TargetMode} initial={initialTargetObjectCount} converted={result.FormulaCount} expected={expectedTargetObjectCount} actual={finalTargetObjectCount}");
                }
            }
            if (result.FormulaCount > 0)
            {
                if (string.Equals(
                        plan.TargetMode,
                        FormulaOleContract.MathTypeOleMode,
                        StringComparison.Ordinal))
                    MathTypeEquationNumbering.UpdateEquationNumbers(document);
                else
                    WordEquationNumbering.TryReconcile(document);
            }
            return result;
        }
        finally
        {
            if (visualTeXRollbackBuffer is not null)
            {
                try { visualTeXRollbackBuffer.Close(WdSaveOptions.wdDoNotSaveChanges); }
                catch { }
            }
            Release(visualTeXRollbackBuffer);
            Release(selection);
            Release(document);
        }
    }

    private sealed class VisualTeXRollbackSnapshot
    {
        public int InsertionStart { get; set; }
        public float FormulaHeightPoints { get; set; }
        public Range Payload { get; set; } = null!;
    }

    private VisualTeXRollbackSnapshot CaptureVisualTeXRollbackSnapshot(
        Document sourceDocument,
        Document rollbackBuffer,
        WordFormulaFormatConversionTarget target)
    {
        InlineShape? shape = null;
        Range? shapeRange = null;
        Range? bufferContent = null;
        Range? bufferInsertion = null;
        Range? bufferPayload = null;
        Table? table = null;
        Range? tableRange = null;
        try
        {
            shape = FindByFormulaId(
                    sourceDocument,
                    target.SourceFormulaId,
                    target.SourceObjectId)
                ?? throw new InvalidOperationException(
                    "The VisualTeX source formula moved before its rollback snapshot was captured.");
            shapeRange = shape.Range;
            var insertionStart = shapeRange.Start;
            if (target.Numbered
                && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
            {
                table = WordEquationNumbering.FindNumberedEquationTable(
                        sourceDocument,
                        target.SourceFormulaId)
                    ?? throw new InvalidOperationException(
                        "The numbered VisualTeX source lost its table before its rollback snapshot was captured.");
                tableRange = table.Range;
                insertionStart = tableRange.Start;
            }

            // Do not clear and reuse the same Word range after it has hosted an
            // embedded OLE object. Word can leave that range in a non-editable COM
            // state, making the next snapshot fail with 0x800A1710. Append every
            // snapshot at the untouched end of the hidden buffer instead; only the
            // current snapshot range is retained for rollback.
            bufferContent = rollbackBuffer.Content;
            var payloadStart = Math.Max(bufferContent.Start, bufferContent.End - 1);
            Release(bufferContent);
            bufferContent = null;
            bufferInsertion = rollbackBuffer.Range(payloadStart, payloadStart);
            bufferInsertion.FormattedText = shapeRange.FormattedText;
            Release(bufferInsertion);
            bufferInsertion = null;
            bufferContent = rollbackBuffer.Content;
            var payloadEnd = Math.Max(payloadStart, bufferContent.End - 1);
            bufferPayload = rollbackBuffer.Range(payloadStart, payloadEnd);
            var snapshot = new VisualTeXRollbackSnapshot
            {
                InsertionStart = insertionStart,
                FormulaHeightPoints = shape.Height,
                Payload = bufferPayload,
            };
            bufferPayload = null;
            return snapshot;
        }
        finally
        {
            Release(bufferPayload);
            Release(bufferInsertion);
            Release(bufferContent);
            Release(tableRange);
            Release(table);
            Release(shapeRange);
            Release(shape);
        }
    }

    private void RestoreVisualTeXRollbackSnapshot(
        Document document,
        VisualTeXRollbackSnapshot snapshot,
        WordFormulaFormatConversionTarget target)
    {
        InlineShape? shape = null;
        Range? insertion = null;
        Range? shapeRange = null;
        Range? content = null;
        try
        {
            document.Activate();
            shape = FindByFormulaId(
                document,
                target.SourceFormulaId,
                target.SourceObjectId);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-snapshot-restore-stage stage=existing-source formulaId={target.SourceFormulaId} found={shape is not null}");
            if (shape is null)
            {
                content = document.Content;
                var start = Math.Max(
                    content.Start,
                    Math.Min(snapshot.InsertionStart, content.End));
                insertion = document.Range(start, start);
                insertion.FormattedText = snapshot.Payload.FormattedText;
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-snapshot-restore-stage stage=insert-ole formulaId={target.SourceFormulaId} start={start} inlineShapes={document.InlineShapes.Count}");
                Release(insertion);
                insertion = null;
                shape = FindByFormulaId(
                        document,
                        target.SourceFormulaId,
                        $"{RangeReferencePrefix}{start}:{start + 1}")
                    ?? throw new InvalidOperationException(
                        "VisualTeX restored the rollback OLE payload, but Word did not expose the restored source formula.");
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-snapshot-restore-stage stage=resolve-ole formulaId={target.SourceFormulaId} found=true");
            }

            BindOleIdentityBookmark(shape, target.SourceFormulaId);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-snapshot-restore-stage stage=bind-id formulaId={target.SourceFormulaId}");
            var standaloneShape = EnsureRollbackDisplayFormulaOwnParagraph(
                document,
                shape,
                target);
            Release(shape);
            shape = standaloneShape;
            shapeRange = shape.Range;
            WordEquationNumbering.RemoveFormulaNumberingArtifacts(
                document,
                target.SourceFormulaId);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-snapshot-restore-stage stage=remove-numbering formulaId={target.SourceFormulaId}");
            WordEquationNumbering.TryReconcileFormula(
                document,
                shapeRange,
                snapshot.FormulaHeightPoints,
                target.Metadata,
                numberingOrderMayHaveChanged: true,
                reuseExistingNumberedTableFormatting: false,
                knownNumberedTable: null);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-snapshot-restore-stage stage=reconcile formulaId={target.SourceFormulaId}");
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-source-restored-from-snapshot formulaId={target.SourceFormulaId} range={shapeRange.Start}:{shapeRange.End} latex={target.Latex}");
        }
        finally
        {
            Release(content);
            Release(shapeRange);
            Release(insertion);
            Release(shape);
        }
    }

    private InlineShape EnsureRollbackDisplayFormulaOwnParagraph(
        Document document,
        InlineShape sourceShape,
        WordFormulaFormatConversionTarget target)
    {
        var formulaId = target.SourceFormulaId;
        var sourceObjectId = target.SourceObjectId;
        var rollbackBridge = BuildFormulaLatexSource(target.Metadata);
        var shape = sourceShape;
        var ownsShape = false;
        try
        {
            for (var pass = 0; pass < 3; pass++)
            {
                Range? shapeRange = null;
                Paragraphs? paragraphs = null;
                Paragraph? paragraph = null;
                Range? paragraphRange = null;
                Range? prefix = null;
                Range? suffix = null;
                Range? split = null;
                try
                {
                    shapeRange = shape.Range;
                    paragraphs = shapeRange.Paragraphs;
                    if (paragraphs.Count != 1)
                        throw new InvalidOperationException(
                            "The restored VisualTeX display formula spans multiple paragraphs before numbering recovery.");
                    paragraph = paragraphs[1];
                    paragraphRange = paragraph.Range;
                    prefix = document.Range(
                        paragraphRange.Start,
                        Math.Max(paragraphRange.Start, shapeRange.Start));
                    suffix = document.Range(
                        Math.Min(shapeRange.End, paragraphRange.End),
                        Math.Max(Math.Min(shapeRange.End, paragraphRange.End), paragraphRange.End - 1));
                    var prefixText = prefix.Text ?? string.Empty;
                    var suffixText = suffix.Text ?? string.Empty;
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-snapshot-restore-paragraph formulaId={formulaId} pass={pass} shape={shapeRange.Start}:{shapeRange.End} paragraph={paragraphRange.Start}:{paragraphRange.End} prefixCodes={DescribeWordCharacters(prefixText)} suffixCodes={DescribeWordCharacters(suffixText)}");

                    var changed = false;
                    if (IsRollbackBridgeResidual(prefixText, rollbackBridge))
                    {
                        prefix.Delete();
                        changed = true;
                    }
                    else if (IsRollbackBridgeResidual(suffixText, rollbackBridge))
                    {
                        suffix.Delete();
                        changed = true;
                    }
                    else if (!IsRollbackParagraphAdornment(prefixText))
                    {
                        split = document.Range(shapeRange.Start, shapeRange.Start);
                        split.InsertBefore("\r");
                        changed = true;
                    }
                    Release(split);
                    split = null;

                    if (!changed && !IsRollbackParagraphAdornment(suffixText))
                    {
                        split = document.Range(shapeRange.End, shapeRange.End);
                        split.InsertAfter("\r");
                        changed = true;
                    }
                    if (!changed)
                    {
                        if (ownsShape)
                        {
                            var result = shape;
                            shape = null!;
                            ownsShape = false;
                            return result;
                        }
                        var resolved = FindByFormulaId(document, formulaId, sourceObjectId)
                            ?? throw new InvalidOperationException(
                                "The restored VisualTeX display formula disappeared while validating its paragraph.");
                        return resolved;
                    }
                }
                finally
                {
                    Release(split);
                    Release(suffix);
                    Release(prefix);
                    Release(paragraphRange);
                    Release(paragraph);
                    Release(paragraphs);
                    Release(shapeRange);
                }

                if (ownsShape) Release(shape);
                shape = FindByFormulaId(document, formulaId, sourceObjectId)
                    ?? throw new InvalidOperationException(
                        "The restored VisualTeX display formula disappeared while separating its paragraph.");
                ownsShape = true;
            }
            throw new InvalidOperationException(
                "VisualTeX could not isolate the restored display formula into its own paragraph.");
        }
        catch
        {
            if (ownsShape) Release(shape);
            throw;
        }
    }

    private static bool IsRollbackBridgeResidual(string? text, string bridge)
    {
        var normalizedText = NormalizeFormulaToLatexVerificationText(text ?? string.Empty).Trim();
        var normalizedBridge = NormalizeFormulaToLatexVerificationText(bridge).Trim();
        return normalizedText.Length > 0
            && string.Equals(normalizedText, normalizedBridge, StringComparison.Ordinal);
    }

    private static bool IsRollbackParagraphAdornment(string? text)
    {
        foreach (var character in text ?? string.Empty)
        {
            if (character is '\t' or '\r' or '\n' or '\v'
                or '\u200b' or '\u200c' or '\u200d' or '\ufeff'
                || char.IsWhiteSpace(character))
                continue;
            return false;
        }
        return true;
    }

    private static string DescribeWordCharacters(string text) =>
        string.Join(",", (text ?? string.Empty).Take(64).Select(ch => $"U+{(int)ch:X4}"));

    private static int CountSimpleFormatObjects(Document document, string mode)
    {
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        try
        {
            var count = 0;
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                if (string.Equals(mode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal))
                {
                    if (WordFormulaMetadataReader.IsNativeOle(shape)) count++;
                }
                else if (string.Equals(mode, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal)
                         && MathTypeOleInterop.IsMathTypeOle(shape))
                {
                    count++;
                }
            }
            return count;
        }
        finally
        {
            Release(shape);
            Release(shapes);
        }
    }

    private static void EnsureBookmarkedMathTypeTargetSurvived(
        Document document,
        string bookmarkName,
        WordFormulaFormatConversionTarget target)
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        Exception? lastError = null;
        while (watch.ElapsedMilliseconds < 750)
        {
            Bookmark? bookmark = null;
            Range? bookmarkRange = null;
            InlineShapes? shapes = null;
            InlineShape? shape = null;
            try
            {
                if (!document.Bookmarks.Exists(bookmarkName))
                {
                    lastError = new InvalidOperationException(
                        "The temporary MathType identity bookmark disappeared after the formula transaction.");
                }
                else
                {
                    bookmark = document.Bookmarks[bookmarkName];
                    bookmarkRange = bookmark.Range;
                    shapes = bookmarkRange.InlineShapes;
                    var mathTypeCount = 0;
                    for (var index = 1; index <= shapes.Count; index++)
                    {
                        Release(shape);
                        shape = shapes[index];
                        if (MathTypeOleInterop.IsMathTypeOle(shape))
                            mathTypeCount++;
                    }
                    if (mathTypeCount == 1)
                    {
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-bookmarked-target-stable formulaId={target.SourceFormulaId} bookmark={bookmarkName} range={bookmarkRange.Start}:{bookmarkRange.End} elapsedMs={watch.ElapsedMilliseconds}");
                        return;
                    }
                    lastError = new InvalidOperationException(
                        $"The temporary MathType identity bookmark contains {mathTypeCount} MathType objects instead of exactly one.");
                }
            }
            catch (Exception error)
            {
                lastError = error;
            }
            finally
            {
                Release(shape);
                Release(shapes);
                Release(bookmarkRange);
                Release(bookmark);
            }
            Thread.Sleep(25);
        }

        throw new InvalidOperationException(
            "Word did not retain the exact MathType replacement object after the formula transaction.",
            lastError);
    }

    private static void WaitForSimpleTargetObjectCountToStabilize(
        Document document,
        string targetMode,
        int expectedCount,
        WordFormulaFormatConversionTarget target)
    {
        const int requiredStableSamples = 2;
        var stableSamples = 0;
        var watch = System.Diagnostics.Stopwatch.StartNew();
        var lastCount = -1;
        while (watch.ElapsedMilliseconds < 1200)
        {
            lastCount = CountSimpleFormatObjects(document, targetMode);
            if (lastCount == expectedCount)
            {
                stableSamples++;
                if (stableSamples >= requiredStableSamples)
                {
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-target-stable formulaId={target.SourceFormulaId} targetMode={targetMode} count={lastCount} elapsedMs={watch.ElapsedMilliseconds}");
                    return;
                }
            }
            else
            {
                stableSamples = 0;
            }
            Thread.Sleep(25);
        }

        throw new InvalidOperationException(
            $"Word did not stabilize the converted target formula after the formula transaction. Expected target count {expectedCount}, actual {lastCount}.");
    }

    private static void TraceSimpleFormatObjectCounts(
        Document document,
        WordFormulaFormatConversionTarget target,
        string stage)
    {
        try
        {
            var visualTeXCount = CountSimpleFormatObjects(
                document,
                FormulaOleContract.NativeOleMode);
            var mathTypeCount = CountSimpleFormatObjects(
                document,
                FormulaOleContract.MathTypeOleMode);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-counts stage={stage} formulaId={target.SourceFormulaId} VT={visualTeXCount} MT={mathTypeCount} latex={target.Latex}");
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-counts-failed stage={stage} formulaId={target.SourceFormulaId} error={error.Message}");
        }
    }

    private bool IsSimpleFormatSourcePresent(
        Document document,
        string sourceMode,
        WordFormulaFormatConversionTarget target)
    {
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        try
        {
            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                // Verification must not depend on the old identity bookmark or
                // captured source range: either can disappear or shift while Word
                // restructures OLE paragraphs. Scan the live document and compare
                // the FormulaId stored inside each VisualTeX OLE instead.
                shapes = document.InlineShapes;
                for (var index = 1; index <= shapes.Count; index++)
                {
                    Release(shape);
                    shape = shapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata is not null
                        && string.Equals(
                            metadata.FormulaId,
                            target.SourceFormulaId,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        Range? liveRange = null;
                        try
                        {
                            liveRange = shape.Range;
                            WordDoubleClickHook.TraceMessage(
                                $"format-conversion-source-match formulaId={target.SourceFormulaId} index={index} liveRange={liveRange.Start}:{liveRange.End} sourceHint={target.SourceObjectId} latex={target.Latex}");
                        }
                        catch { }
                        finally { Release(liveRange); }
                        return true;
                    }
                }
                return false;
            }
            shape = FindMathTypeOleByRange(document, target.SourceObjectId);
            return shape is not null && MathTypeOleInterop.IsMathTypeOle(shape);
        }
        catch { return false; }
        finally
        {
            Release(shape);
            Release(shapes);
        }
    }

    private void EnsureSimpleFormatSourceRemoved(
        Document document,
        string sourceMode,
        WordFormulaFormatConversionTarget target,
        string stage)
    {
        var present = IsSimpleFormatSourcePresent(document, sourceMode, target);
        if (string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
        {
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-source-check stage={stage} formulaId={target.SourceFormulaId} present={present} latex={target.Latex}");
        }
        if (present)
            throw new InvalidOperationException(
                $"Word still contains the source formula after stage '{stage}'. The current formula conversion was rolled back.");
    }

    private static void ThrowIfSimpleFormatConversionFailureInjected(
        WordFormulaFormatConversionTarget target)
    {
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            return;
        var requested = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE");
        if (string.IsNullOrWhiteSpace(requested)) return;
        var matches = string.Equals(requested, "1", StringComparison.Ordinal)
            || string.Equals(requested, target.SourceFormulaId, StringComparison.OrdinalIgnoreCase)
            || string.Equals(requested, "numbered", StringComparison.OrdinalIgnoreCase)
                && target.Numbered;
        if (matches)
            throw new InvalidOperationException(
                "Injected format-conversion failure after deleting the source host.");
    }

    private static void ValidateSimpleFormatConversionPair(
        string sourceMode,
        string targetMode)
    {
        var visualTeXToMathType =
            string.Equals(sourceMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal)
            && string.Equals(targetMode, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal);
        var mathTypeToVisualTeX =
            string.Equals(sourceMode, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal)
            && string.Equals(targetMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal);
        if (!visualTeXToMathType && !mathTypeToVisualTeX)
            throw new ArgumentOutOfRangeException(
                nameof(targetMode),
                $"Unsupported simple format conversion: {sourceMode} -> {targetMode}.");
    }

    private static void ValidatePreparedFormatConversionTarget(
        string targetMode,
        WordFormulaFormatConversionTarget target,
        PreparedWordBulkFormula formula)
    {
        if (string.IsNullOrWhiteSpace(formula.MathMl))
            throw new InvalidDataException(
                $"Formula '{target.Latex}' did not produce MathML.");
        if (string.IsNullOrWhiteSpace(formula.EmfPath)
            || !File.Exists(formula.EmfPath))
            throw new FileNotFoundException(
                $"Formula '{target.Latex}' did not produce an EMF preview.",
                formula.EmfPath);
        if (string.Equals(
                targetMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal)
            && (string.IsNullOrWhiteSpace(formula.PngPath)
                || !File.Exists(formula.PngPath)))
            throw new FileNotFoundException(
                $"Formula '{target.Latex}' did not produce a PNG preview.",
                formula.PngPath);
    }

    private void ValidateSimpleSourceHost(
        Document document,
        string sourceMode,
        WordFormulaFormatConversionTarget target)
    {
        InlineShape? shape = null;
        Table? table = null;
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                shape = FindByFormulaId(
                        document,
                        target.SourceFormulaId,
                        target.SourceObjectId)
                    ?? throw new InvalidOperationException(
                        "The VisualTeX source formula moved before conversion started.");
                if (!WordFormulaMetadataReader.IsNativeOle(shape))
                    throw new InvalidOperationException(
                        "The source object is no longer a VisualTeX OLE formula.");
                shapeRange = shape.Range;
                if (target.Numbered
                    && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                {
                    table = WordEquationNumbering.FindNumberedEquationTable(
                            document,
                            target.SourceFormulaId)
                        ?? throw new InvalidOperationException(
                            "The numbered VisualTeX source no longer owns its numbering table.");
                    if (!IsSafeVisualTeXNumberingTableForConversion(table, shapeRange))
                        throw new InvalidOperationException(
                            "The VisualTeX numbering table contains nonempty extra rows or another formula object; conversion was refused before modifying the document.");
                }
                return;
            }

            shape = FindMathTypeOleByRange(document, target.SourceObjectId)
                ?? throw new InvalidOperationException(
                    "The MathType source formula moved before conversion started.");
            if (!MathTypeOleInterop.IsMathTypeOle(shape))
                throw new InvalidOperationException(
                    "The source object is no longer Equation.DSMT4.");
            if (!string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                return;
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display formula no longer occupies one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.InlineShapes.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display paragraph contains another inline object; conversion was refused.");
            if (!IsSafeMathTypeDisplayParagraph(paragraphRange))
                throw new InvalidOperationException(
                    "The MathType display paragraph contains ordinary user text; conversion was refused to avoid deleting prose.");
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
            Release(table);
            Release(shape);
        }
    }

    private int DeleteSimpleSourceHost(
        Document document,
        string sourceMode,
        WordFormulaFormatConversionTarget target)
    {
        InlineShape? shape = null;
        Range? shapeRange = null;
        Table? table = null;
        Range? tableRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? contentRange = null;
        try
        {
            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                // Do not treat InlineShape.Delete() as a committed replacement.
                // Word can defer OLE removal until something is actually written
                // at the same range, which caused real documents to report success
                // while leaving one VisualTeX object behind. Reuse the mature
                // VisualTeX OLE -> LaTeX replacement path: it deletes the complete
                // source host/numbering, writes the recoverable source text, and
                // verifies that Word materialized that text. The text is only a
                // transaction-local bridge; Numbered/DisplayMode were captured
                // before this call and are used independently for target creation.
                shape = FindByFormulaId(
                        document,
                        target.SourceFormulaId,
                        target.SourceObjectId)
                    ?? throw new InvalidOperationException(
                        "The VisualTeX source formula moved before replacement.");
                shapeRange = shape.Range;
                var start = shapeRange.Start;
                if (target.Numbered
                    && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                {
                    table = TryGetVisualTeXNumberedTable(shapeRange, target.Metadata)
                        ?? throw new InvalidOperationException(
                            "The numbered VisualTeX source lost its table before replacement.");
                    if (table.Rows.Count != 1)
                    {
                        if (!IsSafeVisualTeXNumberingTableForConversion(table, shapeRange))
                            throw new InvalidOperationException(
                                "The VisualTeX numbering table contains nonempty extra rows or another formula object; conversion was refused before modifying the document.");
                        TrimEmptyVisualTeXNumberingRows(table, shapeRange);
                        Release(shapeRange);
                        shapeRange = null;
                        Release(shape);
                        shape = FindByFormulaId(
                                document,
                                target.SourceFormulaId,
                                target.SourceObjectId)
                            ?? throw new InvalidOperationException(
                                "The VisualTeX source formula disappeared while normalizing an empty numbering-table row.");
                        shapeRange = shape.Range;
                        Release(table);
                        table = TryGetVisualTeXNumberedTable(shapeRange, target.Metadata)
                            ?? throw new InvalidOperationException(
                                "The numbered VisualTeX source lost its table after empty-row normalization.");
                    }
                    tableRange = table.Range.Duplicate;
                    start = tableRange.Start;
                }

                var latexSource = BuildFormulaLatexSource(target.Metadata);
                var sourceTarget = new FormulaToLatexTarget
                {
                    Metadata = target.Metadata,
                    ObjectMode = FormulaOleContract.NativeOleMode,
                    LatexSource = latexSource,
                    Start = shapeRange.Start,
                    End = shapeRange.End,
                    FormulaRange = shapeRange,
                    OleShape = shape,
                };
                shapeRange = null;
                shape = null;
                var mutationStarted = false;
                try
                {
                    ConvertFormulaTargetToLatex(
                        document,
                        sourceTarget,
                        ref mutationStarted);

                    // Word can acknowledge InlineShape.Delete() yet keep the EMBED
                    // field alive until a later structural edit. In that case the
                    // verified LaTeX bridge is inserted immediately before the OLE,
                    // shifting the same FormulaId forward by exactly the bridge text
                    // length. Once the recoverable LaTeX text has been verified it is
                    // safe to force-commit removal by replacing the live OLE field
                    // range with empty text inside the same custom Undo transaction.
                    if (IsSimpleFormatSourcePresent(document, sourceMode, target))
                    {
                        InlineShape? deferredShape = null;
                        Range? deferredRange = null;
                        try
                        {
                            deferredShape = FindByFormulaId(
                                    document,
                                    target.SourceFormulaId,
                                    target.SourceObjectId)
                                ?? throw new InvalidOperationException(
                                    "Word deferred VisualTeX OLE removal, but the live source object could not be resolved.");
                            deferredRange = deferredShape.Range;
                            WordDoubleClickHook.TraceMessage(
                                $"format-conversion-force-remove-deferred-source formulaId={target.SourceFormulaId} range={deferredRange.Start}:{deferredRange.End} latexLength={latexSource.Length}");
                            deferredRange.Text = string.Empty;
                        }
                        finally
                        {
                            Release(deferredRange);
                            Release(deferredShape);
                        }
                    }
                    EnsureSimpleFormatSourceRemoved(
                        document,
                        sourceMode,
                        target,
                        "latex-bridge");

                    // Remove only the verified temporary LaTeX characters. For a
                    // former numbered table the mature converter deliberately adds
                    // one paragraph mark; keep that paragraph as the clean insertion
                    // host for the new MathType display equation.
                    Range? latexRange = null;
                    try
                    {
                        latexRange = document.Range(
                            start,
                            Math.Min(document.Content.End, start + latexSource.Length));
                        if (!string.Equals(
                                NormalizeFormulaToLatexVerificationText(latexRange.Text ?? string.Empty),
                                NormalizeFormulaToLatexVerificationText(latexSource),
                                StringComparison.Ordinal))
                            throw new InvalidDataException(
                                "The temporary LaTeX bridge changed before target insertion.");
                        latexRange.Delete();
                    }
                    finally { Release(latexRange); }
                }
                finally
                {
                    ReleaseFormulaToLatexTargets(new[] { sourceTarget });
                }
                return start;
            }

            shape = FindMathTypeOleByRange(document, target.SourceObjectId)
                ?? throw new InvalidOperationException(
                    "The MathType source formula moved before replacement.");
            shapeRange = shape.Range;
            var mathTypeStart = shapeRange.Start;
            if (!string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
            {
                RemoveInlineOleTypingAnchorAfter(shape);
                shape.Delete();
                return mathTypeStart;
            }

            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display source no longer occupies one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (!IsSafeMathTypeDisplayParagraph(paragraphRange))
                throw new InvalidOperationException(
                    "The MathType display paragraph contains ordinary user text; conversion was stopped.");
            contentRange = paragraphRange.Duplicate;
            mathTypeStart = contentRange.Start;
            var text = contentRange.Text ?? string.Empty;
            if (contentRange.End > contentRange.Start
                && text.Length > 0
                && (text[text.Length - 1] == '\r' || text[text.Length - 1] == '\a'))
                contentRange.SetRange(contentRange.Start, contentRange.End - 1);
            contentRange.Delete();
            return mathTypeStart;
        }
        finally
        {
            Release(contentRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(tableRange);
            Release(table);
            Release(shapeRange);
            Release(shape);
        }
    }

    private static bool IsSafeVisualTeXNumberingTableForConversion(
        Table table,
        Range formulaRange)
    {
        Rows? rows = null;
        Columns? columns = null;
        Range? tableRange = null;
        InlineShapes? tableShapes = null;
        Row? row = null;
        Range? rowRange = null;
        try
        {
            columns = table.Columns;
            if (columns.Count != 3) return false;
            tableRange = table.Range;
            tableShapes = tableRange.InlineShapes;
            if (tableShapes.Count != 1) return false;

            rows = table.Rows;
            var formulaRowFound = false;
            for (var index = 1; index <= rows.Count; index++)
            {
                Release(rowRange);
                rowRange = null;
                Release(row);
                row = rows[index];
                rowRange = row.Range;
                var ownsFormula = formulaRange.Start >= rowRange.Start
                    && formulaRange.Start < rowRange.End;
                if (ownsFormula)
                {
                    if (formulaRowFound) return false;
                    formulaRowFound = true;
                    continue;
                }
                if (!IsStructurallyEmptyVisualTeXNumberingRow(rowRange))
                    return false;
            }
            return formulaRowFound;
        }
        finally
        {
            Release(rowRange);
            Release(row);
            Release(tableShapes);
            Release(tableRange);
            Release(columns);
            Release(rows);
        }
    }

    private static bool IsStructurallyEmptyVisualTeXNumberingRow(Range rowRange)
    {
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        try
        {
            shapes = rowRange.InlineShapes;
            if (shapes.Count != 0) return false;
            maths = rowRange.OMaths;
            if (maths.Count != 0) return false;
            fields = rowRange.Fields;
            if (fields.Count != 0) return false;
            bookmarks = rowRange.Bookmarks;
            if (bookmarks.Count != 0) return false;
            foreach (var character in rowRange.Text ?? string.Empty)
            {
                if (character is '\r' or '\a' or '\n' or '\v'
                    or '\u200b' or '\u200c' or '\u200d' or '\ufeff'
                    || char.IsWhiteSpace(character))
                    continue;
                return false;
            }
            return true;
        }
        finally
        {
            Release(bookmarks);
            Release(fields);
            Release(maths);
            Release(shapes);
        }
    }

    private static void TrimEmptyVisualTeXNumberingRows(
        Table table,
        Range formulaRange)
    {
        if (!IsSafeVisualTeXNumberingTableForConversion(table, formulaRange))
            throw new InvalidOperationException(
                "The VisualTeX numbering table contains nonempty extra rows or another formula object.");

        Rows? rows = null;
        Row? row = null;
        Range? rowRange = null;
        try
        {
            rows = table.Rows;
            for (var index = rows.Count; index >= 1; index--)
            {
                Release(rowRange);
                rowRange = null;
                Release(row);
                row = rows[index];
                rowRange = row.Range;
                if (formulaRange.Start >= rowRange.Start
                    && formulaRange.Start < rowRange.End)
                    continue;
                row.Delete();
            }
            if (table.Rows.Count != 1 || table.Columns.Count != 3)
                throw new InvalidOperationException(
                    "Word did not normalize the VisualTeX numbering table to one 1x3 formula row.");
        }
        finally
        {
            Release(rowRange);
            Release(row);
            Release(rows);
        }
    }

    private static bool IsSafeMathTypeDisplayParagraph(Range paragraphRange)
    {
        var text = paragraphRange.Text ?? string.Empty;
        foreach (var character in text)
        {
            if (char.IsWhiteSpace(character) || char.IsDigit(character)) continue;
            if (character < ' ') continue;
            if (character is '\u0001' or '\u0013' or '\u0014' or '\u0015') continue;
            if ("()[]{}.-–—_:;,+/\\".IndexOf(character) >= 0) continue;
            return false;
        }
        return true;
    }

    private static void RemoveDetachedVisualTeXNumberingArtifacts(
        Document document,
        string formulaId)
    {
        // The visible 1x3 numbered host has already been deleted. Only the hidden
        // native SEQ caption/frame remains. Delete that detached structure without
        // touching any former table/cell Range.
        TryDeleteBookmark(document, WordEquationNumbering.NativeNumberBookmarkName(formulaId));

        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Range? captionRange = null;
        Frames? frames = null;
        Frame? frame = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = WordEquationNumbering.NativeCaptionBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName)) return;
            captionBookmark = bookmarks[captionName];
            captionRange = captionBookmark.Range;
            try
            {
                frames = captionRange.Frames;
                if (frames.Count > 0)
                {
                    frame = frames[1];
                    frame.Delete();
                    Release(frame);
                    frame = null;
                    Release(frames);
                    frames = null;
                    Release(captionRange);
                    captionRange = captionBookmark.Range;
                }
            }
            catch
            {
                // The caption may already have lost its clipping frame. Deleting
                // the bookmarked caption contents is sufficient in that case.
            }
            captionRange.Delete();
        }
        finally
        {
            Release(frame);
            Release(frames);
            Release(captionRange);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    private static void TryDeleteBookmark(Document document, string name)
    {
        if (document is null || string.IsNullOrWhiteSpace(name)) return;
        try
        {
            if (!document.Bookmarks.Exists(name)) return;
            Bookmark? bookmark = null;
            try
            {
                bookmark = document.Bookmarks[name];
                bookmark.Delete();
            }
            finally { Release(bookmark); }
        }
        catch
        {
            // The source host itself has already been removed. A collapsed stale
            // bookmark is harmless and must never turn cleanup into a conversion failure.
        }
    }
}
