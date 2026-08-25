using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    private static bool HasAdjacentFormatConversionSourceHosts(
        IReadOnlyList<WordFormulaFormatConversionTarget> targets)
    {
        var ranges = new List<(int Start, int End)>();
        foreach (var target in targets)
        {
            if (!TryParseRangeReference(
                    target.SourceObjectId,
                    out var start,
                    out var end))
                continue;
            ranges.Add((start, end));
        }
        ranges.Sort((left, right) => left.Start.CompareTo(right.Start));
        for (var index = 1; index < ranges.Count; index++)
        {
            // Shared boundaries are the normal zero-gap OLE+OLE case. Treat an
            // overlap conservatively as adjacent too: Word bookmarks around
            // overlapping formula hosts are not safe forward locators after an
            // earlier source is replaced.
            if (ranges[index - 1].End >= ranges[index].Start)
                return true;
        }
        return false;
    }

    private static IReadOnlyList<IReadOnlyList<WordFormulaFormatConversionTarget>>
        BuildAdjacentInlineOmmlGroups(
            IReadOnlyList<WordFormulaFormatConversionTarget> targets)
    {
        var sorted = targets
            .Where(target => string.Equals(
                target.DisplayMode,
                "inline",
                StringComparison.Ordinal))
            .Select(target =>
            {
                var valid = TryParseRangeReference(
                    target.SourceObjectId,
                    out var start,
                    out var end);
                return (Target: target, Valid: valid, Start: start, End: end);
            })
            .Where(item => item.Valid)
            .OrderBy(item => item.Start)
            .ToArray();
        var groups = new List<IReadOnlyList<WordFormulaFormatConversionTarget>>();
        var current = new List<(WordFormulaFormatConversionTarget Target, int Start, int End)>();
        foreach (var item in sorted)
        {
            if (current.Count == 0)
            {
                current.Add((item.Target, item.Start, item.End));
                continue;
            }
            var previous = current[current.Count - 1];
            if (previous.End == item.Start)
            {
                current.Add((item.Target, item.Start, item.End));
                continue;
            }
            if (current.Count > 1)
                groups.Add(current.Select(entry => entry.Target).ToArray());
            current.Clear();
            current.Add((item.Target, item.Start, item.End));
        }
        if (current.Count > 1)
            groups.Add(current.Select(entry => entry.Target).ToArray());
        return groups;
    }

    private int ApplyAdjacentMathTypeToOmmlGroup(
        Document document,
        WordFormulaFormatConversionPlan plan,
        IReadOnlyList<WordFormulaFormatConversionTarget> group,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared,
        WordOmmlConverter.BatchSource ommlBatchSource)
    {
        if (group.Count < 2)
            throw new ArgumentOutOfRangeException(nameof(group));
        var ordered = group.OrderBy(target => target.SourceStart).ToArray();
        if (ordered.Any(target => !string.Equals(
                target.DisplayMode,
                "inline",
                StringComparison.Ordinal)))
            throw new InvalidOperationException(
                "Only adjacent inline MathType formulas can use grouped OMML conversion.");

        UndoRecord? undoRecord = null;
        Range? insertion = null;
        IReadOnlyList<Range>? insertedRanges = null;
        var undoEnded = false;
        var metadataSaved = new List<string>();
        try
        {
            undoRecord = BeginUndoRecord("VisualTeX Convert Adjacent MathType Formulas");
            if (undoRecord is null)
                throw new InvalidOperationException(
                    "Word 无法建立相邻公式转换撤销事务。为避免转换失败时丢失原公式，本次转换已停止。");

            var targetCountBefore = CountSimpleFormatObjects(
                document,
                FormulaOleContract.WordOmmlMode);
            var insertionStart = ordered[0].SourceStart;
            foreach (var source in ordered.OrderByDescending(target => target.SourceStart))
            {
                _ = DeleteSimpleSourceHost(
                    document,
                    plan.SourceMode,
                    source,
                    knownReferenceCounts: null);
                EnsureSimpleFormatSourceRemoved(
                    document,
                    plan.SourceMode,
                    source,
                    "adjacent-group-delete-source");
                ThrowIfSimpleFormatConversionFailureInjected(source);
            }

            Range? content = null;
            try
            {
                content = document.Content;
                insertionStart = Math.Max(
                    content.Start,
                    Math.Min(insertionStart, content.End));
            }
            finally { Release(content); }
            insertion = document.Range(insertionStart, insertionStart);

            var formulaIds = ordered
                .Select(target => prepared[target.Id].Session.FormulaId)
                .ToArray();
            insertedRanges = ommlBatchSource.InsertAdjacentInlineGroup(
                _application,
                document,
                insertion,
                formulaIds);
            document.Activate();
            if (insertedRanges.Count != ordered.Length)
                throw new InvalidOperationException(
                    $"Word retained {insertedRanges.Count}/{ordered.Length} adjacent OMML formulas.");

            for (var index = 0; index < ordered.Length; index++)
            {
                var target = ordered[index];
                var formula = prepared[target.Id];
                var session = formula.Session;
                session.Mode = "create";
                session.SourceDocumentId = plan.DocumentId;
                session.SourceObjectId =
                    $"{RangeReferencePrefix}{insertedRanges[index].Start}:{insertedRanges[index].End}";
                session.DisplayMode = "inline";
                session.ObjectMode = FormulaOleContract.WordOmmlMode;
                session.Numbered = false;
                session.MathTypeNumberPosition = target.MathTypeNumberPosition;
                session.FontSizePt = target.FontSizePt;
                session.OriginalMetadata = null;

                var metadata = session.ToMetadata();
                metadata.NativeOmmlFingerprint =
                    ommlBatchSource.GetSourceFingerprint(session.FormulaId);
                ApplyOmmlTypography(
                    insertedRanges[index],
                    session.FontSizePt,
                    metadata);
                Bookmark? bookmark = null;
                try
                {
                    bookmark = WordOmmlFormulaStore.Wrap(
                        document,
                        insertedRanges[index],
                        metadata);
                    if (!WordOmmlFormulaStore.IsCanonicalAnchor(
                            bookmark,
                            insertedRanges[index]))
                    {
                        Release(bookmark);
                        bookmark = WordOmmlFormulaStore.Wrap(
                            document,
                            insertedRanges[index],
                            metadata,
                            replaceExisting: true);
                    }
                    WordOmmlFormulaStore.SaveNew(document, metadata);
                    metadataSaved.Add(metadata.FormulaId);
                }
                finally { Release(bookmark); }
            }

            EndUndoRecord(undoRecord);
            undoEnded = true;
            var targetCountAfter = CountSimpleFormatObjects(
                document,
                FormulaOleContract.WordOmmlMode);
            if (targetCountAfter != targetCountBefore + ordered.Length)
                throw new InvalidOperationException(
                    $"Word retained {targetCountAfter - targetCountBefore}/{ordered.Length} OMML targets after adjacent-group conversion.");
            foreach (var source in ordered)
                EnsureSimpleFormatSourceRemoved(
                    document,
                    plan.SourceMode,
                    source,
                    "adjacent-group-post-transaction");

            WordDoubleClickHook.TraceMessage(
                $"format-conversion-adjacent-omml-group-complete count={ordered.Length} start={ordered[0].SourceStart} end={ordered[ordered.Length - 1].SourceStart}");
            return ordered.Length;
        }
        catch
        {
            if (!undoEnded)
            {
                EndUndoRecord(undoRecord);
                undoEnded = true;
            }
            if (!TryUndoFormulaToLatexConversion(document))
                throw new InvalidOperationException(
                    "相邻 MathType→OMML 转换失败，而且 Word 无法自动恢复原公式。请立即停止编辑当前文档。");
            foreach (var formulaId in metadataSaved)
            {
                try { WordOmmlFormulaStore.Delete(document, formulaId); } catch { }
            }
            foreach (var source in ordered)
                ValidateSimpleSourceHost(document, plan.SourceMode, source);
            throw;
        }
        finally
        {
            if (insertedRanges is not null)
                foreach (var range in insertedRanges) Release(range);
            Release(insertion);
            if (!undoEnded) EndUndoRecord(undoRecord);
            Release(undoRecord);
        }
    }

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
            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                var repairedCaptionFrames =
                    WordEquationNumbering.RepairLeakedNativeCaptionFrames(document);
                if (repairedCaptionFrames > 0)
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-preflight-repaired-caption-frames count={repairedCaptionFrames}");
            }
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
                NumberFormatId = WordEquationNumbering.GetEquationNumberFormatId(document),
            };

            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                CaptureOmmlFormulaFormatConversionTargets(
                    document,
                    scope,
                    wholeDocument,
                    plan);
                return plan;
            }

            shapes = document.InlineShapes;
            var tracePlanPerformance = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF"),
                "1",
                StringComparison.Ordinal);
            IReadOnlyList<MathTypeWordOpenXml.OleSnapshot>? bulkOleSnapshots = null;
            var bulkOleSnapshotIndex = 0;
            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal))
            {
                Range? packageRange = null;
                var bulkWatch = tracePlanPerformance
                    ? System.Diagnostics.Stopwatch.StartNew()
                    : null;
                try
                {
                    packageRange = document.Content;
                    bulkOleSnapshots = MathTypeWordOpenXml.ReadOleSnapshots(packageRange);
                    if (bulkWatch is not null)
                    {
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-plan-bulk-ole-snapshot count={bulkOleSnapshots.Count} elapsedMs={bulkWatch.ElapsedMilliseconds}");
                    }
                }
                catch (Exception bulkError)
                {
                    bulkOleSnapshots = null;
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-plan-bulk-ole-snapshot-fallback error={bulkError.GetType().Name}:{bulkError.Message}");
                }
                finally { Release(packageRange); }
            }
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                var planPerfWatch = tracePlanPerformance
                    ? System.Diagnostics.Stopwatch.StartNew()
                    : null;
                long planPerfLastMs = 0;
                void TracePlanPerf(string perfStage)
                {
                    if (planPerfWatch is null) return;
                    var totalMs = planPerfWatch.ElapsedMilliseconds;
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-plan-perf stage={perfStage} source={sourceMode} index={index}/{shapes.Count} deltaMs={totalMs - planPerfLastMs} totalMs={totalMs}");
                    planPerfLastMs = totalMs;
                }
                try
                {
                    shape = shapes[index];
                    TracePlanPerf("shape-access");
                    MathTypeWordOpenXml.OleSnapshot? bulkOleSnapshot = null;
                    if (bulkOleSnapshots is not null
                        && shape.Type is WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                            or WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                    {
                        if (bulkOleSnapshotIndex < bulkOleSnapshots.Count)
                            bulkOleSnapshot = bulkOleSnapshots[bulkOleSnapshotIndex];
                        bulkOleSnapshotIndex++;
                    }
                    FormulaMetadata? metadata = null;
                    string? sourceMathMl = null;
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
                        TracePlanPerf("identify-mathtype");
                        if (bulkOleSnapshot is not null
                            && bulkOleSnapshot.CompoundFile.Length > 0
                            && (string.Equals(
                                    bulkOleSnapshot.ProgId,
                                    "Equation.DSMT4",
                                    StringComparison.OrdinalIgnoreCase)
                                || bulkOleSnapshot.ProgId.StartsWith(
                                    "Equation.",
                                    StringComparison.OrdinalIgnoreCase)))
                        {
                            try
                            {
                                sourceMathMl = MathTypeOleStorage.ReadMathMl(
                                    bulkOleSnapshot.CompoundFile);
                                TracePlanPerf("read-mathml-bulk");
                            }
                            catch
                            {
                                sourceMathMl = MathTypeOleStorage.ReadMathMl(shape);
                                TracePlanPerf("read-mathml-fallback");
                            }
                        }
                        else
                        {
                            sourceMathMl = MathTypeOleStorage.ReadMathMl(shape);
                            TracePlanPerf("read-mathml-fallback");
                        }
                        metadata = MathTypeOleInterop.ReadMetadata(
                            _application,
                            shape,
                            sourceMathMl);
                        TracePlanPerf("read-metadata");
                        if (MathTypeOleInterop.TryReadDisplayNumberPosition(
                                shape,
                                out var detectedPosition))
                            mathTypeNumberPosition = detectedPosition;
                        TracePlanPerf("number-position");
                        sourceFormulaId = metadata.FormulaId;
                    }

                    range = shape.Range;
                    TracePlanPerf("shape-range");
                    if (!FormulaRangeMatchesScope(range, scope, wholeDocument))
                        continue;

                    var latex = string.IsNullOrWhiteSpace(metadata.Latex)
                        ? string.Join("\n", metadata.Lines.Select(line => line.Latex))
                        : metadata.Latex;
                    latex = (latex ?? string.Empty).Trim();
                    if (latex.Length == 0)
                        throw new InvalidDataException(
                            "A source formula has no recoverable LaTeX and was not converted.");

                    if (string.Equals(
                            sourceMode,
                            FormulaOleContract.NativeOleMode,
                            StringComparison.Ordinal)
                        && metadata.Numbered
                        && string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
                    {
                        // VisualTeX numbered formulas do not carry MathType's side
                        // preference in their own metadata.  Conversion must follow
                        // the current document's native MathType preference instead
                        // of silently forcing every target to the right.
                        mathTypeNumberPosition = ReadMathTypeNumberPositionPreference(document);
                    }

                    plan.Targets.Add(new WordFormulaFormatConversionTarget
                    {
                        Id = Guid.NewGuid().ToString("D"),
                        SourceFormulaId = sourceFormulaId,
                        SourceObjectId = $"{RangeReferencePrefix}{range.Start}:{range.End}",
                        SourceStart = range.Start,
                        Latex = latex,
                        SourceMathMl = sourceMathMl,
                        SourceIsManagedOmml = false,
                        DisplayMode = metadata.DisplayMode,
                        Numbered = metadata.Numbered,
                        MathTypeNumberPosition = mathTypeNumberPosition,
                        FontSizePt = FormulaFontSize.Normalize(metadata.FontSizePt),
                        Metadata = metadata,
                    });
                    TracePlanPerf("target-added");
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

    private static void CaptureOmmlFormulaFormatConversionTargets(
        Document document,
        Range scope,
        bool wholeDocument,
        WordFormulaFormatConversionPlan plan)
    {
        if (!wholeDocument
            && TryCaptureSingleOmmlFormatConversionTarget(document, scope, plan))
            return;

        List<FormulaToLatexTarget>? sources = null;
        try
        {
            sources = CaptureFormulaToLatexTargets(
                document,
                scope,
                wholeDocument,
                FormulaOleContract.WordOmmlMode,
                refreshOmmlMetadata: true);
            foreach (var source in sources)
            {
                var metadata = source.Metadata;
                var range = source.FormulaRange;
                var wordOpenXml = WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                    document,
                    range,
                    metadata.FormulaId);
                var sourceMathMl = WordOmmlConverter.TransformOmmlToMathMl(
                    wordOpenXml,
                    display: string.Equals(
                        metadata.DisplayMode,
                        "block",
                        StringComparison.Ordinal));
                var latex = MathMlToLatexConverter.Convert(sourceMathMl).Trim();
                if (latex.Length == 0)
                    throw new InvalidDataException(
                        "A Word OMML source formula could not be converted to editable LaTeX for preview rendering.");

                plan.Targets.Add(new WordFormulaFormatConversionTarget
                {
                    Id = Guid.NewGuid().ToString("D"),
                    SourceFormulaId = metadata.FormulaId,
                    SourceObjectId = $"{RangeReferencePrefix}{range.Start}:{range.End}",
                    SourceStart = range.Start,
                    Latex = latex,
                    SourceMathMl = sourceMathMl,
                    SourceIsManagedOmml = source.OmmlBookmark is not null,
                    DisplayMode = metadata.DisplayMode,
                    Numbered = metadata.Numbered,
                    MathTypeNumberPosition = metadata.Numbered
                        ? ReadMathTypeNumberPositionPreference(document)
                        : "right",
                    FontSizePt = FormulaFontSize.Normalize(metadata.FontSizePt),
                    Metadata = metadata,
                });
            }
        }
        finally
        {
            ReleaseFormulaToLatexTargets(sources);
        }
    }

    private static bool TryCaptureSingleOmmlFormatConversionTarget(
        Document document,
        Range scope,
        WordFormulaFormatConversionPlan plan)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? equationRange = null;
        Bookmark? bookmark = null;
        try
        {
            maths = scope.OMaths;
            if (maths.Count != 1) return false;
            math = maths[1];
            equationRange = math.Range.Duplicate;
            bookmark = WordOmmlFormulaStore.FindAtRange(document, equationRange);

            FormulaMetadata metadata;
            var sourceIsManagedOmml = bookmark is not null;
            if (bookmark is not null)
            {
                var stored = WordOmmlFormulaStore.TryRead(document, bookmark);
                if (stored is null) return false;
                metadata = WordOmmlNativeSource.RefreshForVisualTeX(
                    document,
                    bookmark,
                    stored);
                Release(equationRange);
                equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            }
            else
            {
                metadata = WordOmmlNativeSource.CreateForNative(
                    document,
                    equationRange);
            }

            var wordOpenXml = WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                document,
                equationRange,
                metadata.FormulaId);
            var sourceMathMl = WordOmmlConverter.TransformOmmlToMathMl(
                wordOpenXml,
                display: string.Equals(
                    metadata.DisplayMode,
                    "block",
                    StringComparison.Ordinal));
            var latex = MathMlToLatexConverter.Convert(sourceMathMl).Trim();
            if (latex.Length == 0)
                throw new InvalidDataException(
                    "The selected Word OMML source could not be converted to editable LaTeX.");

            plan.Targets.Add(new WordFormulaFormatConversionTarget
            {
                Id = Guid.NewGuid().ToString("D"),
                SourceFormulaId = metadata.FormulaId,
                SourceObjectId = $"{RangeReferencePrefix}{equationRange.Start}:{equationRange.End}",
                SourceStart = equationRange.Start,
                Latex = latex,
                SourceMathMl = sourceMathMl,
                SourceIsManagedOmml = sourceIsManagedOmml,
                DisplayMode = metadata.DisplayMode,
                Numbered = metadata.Numbered,
                MathTypeNumberPosition = metadata.Numbered
                    ? ReadMathTypeNumberPositionPreference(document)
                    : "right",
                FontSizePt = FormulaFontSize.Normalize(metadata.FontSizePt),
                Metadata = metadata,
            });
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-single-omml-local-capture managed={sourceIsManagedOmml} range={equationRange.Start}:{equationRange.End}");
            return true;
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Fall back to the mature document-level enumerator when Word does
            // not expose exactly one stable local OMath at the current selection.
            return false;
        }
        finally
        {
            Release(bookmark);
            Release(equationRange);
            Release(math);
            Release(maths);
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
        WordOmmlConverter.BatchSource? ommlBatchSource = null;
        var forwardSourceBookmarks = new Dictionary<string, string>(StringComparer.Ordinal);
        var retainedMathTypeTargetBookmarks = new Dictionary<string, string>(StringComparer.Ordinal);
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        Window? batchWindow = null;
        var previousViewType = WdViewType.wdPrintView;
        var batchViewSuspended = false;
        var previousPagination = true;
        var paginationSuspended = false;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, plan.DocumentId);
            // Freeze the document's numbering preset at capture time. Source-host
            // deletion can temporarily remove every numbered VisualTeX structure;
            // rebuilding the target must not fall back to the registry/default
            // continuous sequence during that empty intermediate state.
            WordEquationNumbering.RestoreEquationNumberFormatForConversion(
                document,
                plan.NumberFormatId);
            if (plan.Targets.Count > 1)
            {
                try
                {
                    previousScreenUpdating = _application.ScreenUpdating;
                    _application.ScreenUpdating = false;
                    screenUpdatingSuspended = true;
                }
                catch { }
            }
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
            var targetIsMathType = string.Equals(
                plan.TargetMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal);
            var targetIsVisualTeX = string.Equals(
                plan.TargetMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal);
            var targetIsOmml = string.Equals(
                plan.TargetMode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal);
            var canFinalizeSingleMathTypeNumberLocally =
                targetIsMathType
                && plan.Targets.Count == 1
                && plan.Targets[0].Numbered
                && MathTypeEquationNumbering.CountPlaceRefFields(document) == 0;
            if (targetIsOmml && plan.Targets.Count > 1)
            {
                try
                {
                    batchWindow = _application.ActiveWindow;
                    if (batchWindow is not null)
                    {
                        previousViewType = batchWindow.View.Type;
                        if (previousViewType != WdViewType.wdNormalView)
                        {
                            batchWindow.View.Type = WdViewType.wdNormalView;
                            batchViewSuspended = true;
                        }
                    }
                    previousPagination = _application.Options.Pagination;
                    if (previousPagination)
                    {
                        _application.Options.Pagination = false;
                        paginationSuspended = true;
                    }
                }
                catch
                {
                    // View/pagination suspension is only a batch performance aid.
                    // Conversion correctness must not depend on it.
                }
            }
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-runtime assembly={typeof(WordFormulaService).Assembly.Location} source={plan.SourceMode} target={plan.TargetMode} targetIsVisualTeX={targetIsVisualTeX} targetIsMathType={targetIsMathType}");
            var initialTargetObjectCount = CountSimpleFormatObjects(document, plan.TargetMode);
            var initialSourceObjectCount = CountSimpleFormatObjects(document, plan.SourceMode);
            var referenceAliasesByTargetId =
                new Dictionary<string, IReadOnlyList<EquationReferenceBookmarkAlias>>(StringComparer.Ordinal);
            foreach (var target in plan.Targets.Where(item => item.Numbered))
            {
                IReadOnlyList<EquationReferenceBookmarkAlias> aliases;
                if (string.Equals(
                        plan.SourceMode,
                        FormulaOleContract.MathTypeOleMode,
                        StringComparison.Ordinal))
                {
                    InlineShape? sourceShape = null;
                    try
                    {
                        sourceShape = FindMathTypeOleByRange(
                            document,
                            target.SourceObjectId,
                            allowGlobalFallback: false);
                        if (sourceShape is null) continue;
                        aliases = MathTypeEquationReferences.CaptureFormatConversionAliasesFromMathType(
                            document,
                            sourceShape);
                    }
                    finally { Release(sourceShape); }
                }
                else
                {
                    aliases = MathTypeEquationReferences.CaptureFormatConversionAliasesFromVisualTeX(
                        document,
                        target.SourceFormulaId);
                }
                if (aliases.Count == 0) continue;
                referenceAliasesByTargetId[target.Id] = aliases;
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-reference-aliases-captured sourceMode={plan.SourceMode} formulaId={target.SourceFormulaId} aliases={string.Join(",", aliases.Select(alias => alias.Name))}");
            }
            var capturedReferenceFormatting =
                MathTypeEquationReferences.CaptureReferenceCharacterFormatting(
                    document,
                    referenceAliasesByTargetId.Values
                        .SelectMany(items => items)
                        .Select(alias => alias.Name)
                        .Distinct(StringComparer.OrdinalIgnoreCase));
            if (targetIsOmml && plan.Targets.Count > 0)
            {
                var ommlFormulas = plan.Targets.Select(target =>
                {
                    var formula = prepared[target.Id];
                    return (
                        FormulaId: formula.Session.FormulaId,
                        MathMl: formula.MathMl
                            ?? throw new InvalidDataException(
                                $"Missing MathML for OMML conversion target '{target.Latex}'."));
                }).ToList();
                var batchSourceWatch = System.Diagnostics.Stopwatch.StartNew();
                ommlBatchSource = WordOmmlConverter.CreateBatchSource(
                    _application,
                    ommlFormulas);
                document.Activate();
                Release(selection);
                selection = _application.Selection;
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-omml-batch-source-created formulas={ommlFormulas.Count} elapsedMs={batchSourceWatch.ElapsedMilliseconds}");
            }
            IReadOnlyDictionary<string, int>? sourceReferenceCounts = null;
            IReadOnlyDictionary<int, ResolvedEquationHeadingScope>? batchHeadingScopes = null;
            HashSet<int>? preparedHeadingScopeStarts = null;
            if (targetIsMathType && plan.Targets.Any(target => target.Numbered))
            {
                var batchNumberFormatId = EquationNumberFormat.Resolve(plan.NumberFormatId).Id;
                var batchNumberFormat = EquationNumberFormat.Resolve(batchNumberFormatId);
                if (batchNumberFormat.UsesHeading)
                {
                    batchHeadingScopes = WordEquationNumbering.CaptureHeadingScopesAtPositions(
                        document,
                        batchNumberFormatId,
                        plan.Targets
                            .Where(target => target.Numbered)
                            .Select(target => target.SourceStart));
                    preparedHeadingScopeStarts = new HashSet<int>();
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-heading-scope-cache format={batchNumberFormatId} formulas={batchHeadingScopes.Count} scopes={batchHeadingScopes.Values.Select(scope => scope.ScopeStart).Distinct().Count()}");
                }
            }
            if ((string.Equals(
                     plan.SourceMode,
                     FormulaOleContract.NativeOleMode,
                     StringComparison.Ordinal)
                 || string.Equals(
                     plan.SourceMode,
                     FormulaOleContract.WordOmmlMode,
                     StringComparison.Ordinal))
                && plan.Targets.Any(target => target.Numbered)
                && WordEquationNumbering.TryGetHealthyEquationReferenceCounts(
                    document,
                    out var capturedReferenceCounts))
            {
                sourceReferenceCounts = capturedReferenceCounts;
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-reference-counts-fast-path source={plan.SourceMode} formulas={sourceReferenceCounts.Count}");
            }
            var hasAdjacentSourceHosts = targetIsOmml
                && HasAdjacentFormatConversionSourceHosts(plan.Targets);
            var useForwardSourceBookmarks = targetIsOmml
                && plan.Targets.Count > 1
                && !hasAdjacentSourceHosts;
            if (hasAdjacentSourceHosts)
            {
                WordDoubleClickHook.TraceMessage(
                    "format-conversion-adjacent-source-safe-order target=OMML order=descending bookmarks=disabled");
            }
            if (useForwardSourceBookmarks)
            {
                Bookmarks? sourceBookmarks = null;
                try
                {
                    sourceBookmarks = document.Bookmarks;
                    foreach (var target in plan.Targets)
                    {
                        if (!TryParseRangeReference(
                                target.SourceObjectId,
                                out var sourceStart,
                                out var sourceEnd))
                            throw new InvalidDataException(
                                $"MathType source range is invalid for forward conversion: {target.SourceObjectId}");
                        Range? sourceRange = null;
                        try
                        {
                            sourceRange = document.Range(sourceStart, sourceEnd);
                            var bookmarkName = "VTFC_" + target.Id.Replace("-", string.Empty);
                            if (sourceBookmarks.Exists(bookmarkName))
                            {
                                Bookmark? existing = null;
                                try
                                {
                                    existing = sourceBookmarks[bookmarkName];
                                    existing.Delete();
                                }
                                finally { Release(existing); }
                            }
                            sourceBookmarks.Add(bookmarkName, sourceRange);
                            forwardSourceBookmarks[target.Id] = bookmarkName;
                        }
                        finally { Release(sourceRange); }
                    }
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-forward-source-bookmarks count={forwardSourceBookmarks.Count}");
                }
                finally { Release(sourceBookmarks); }
            }

            var adjacentInlineOmmlGroups = targetIsOmml
                && string.Equals(
                    plan.SourceMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)
                ? BuildAdjacentInlineOmmlGroups(plan.Targets)
                : Array.Empty<IReadOnlyList<WordFormulaFormatConversionTarget>>();
            var adjacentOmmlGroupByMember =
                new Dictionary<string, IReadOnlyList<WordFormulaFormatConversionTarget>>(
                    StringComparer.Ordinal);
            foreach (var group in adjacentInlineOmmlGroups)
            {
                foreach (var member in group)
                    adjacentOmmlGroupByMember[member.Id] = group;
            }
            var processedAdjacentOmmlTargets = new HashSet<string>(StringComparer.Ordinal);
            if (adjacentInlineOmmlGroups.Count > 0)
            {
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-adjacent-omml-groups groups={adjacentInlineOmmlGroups.Count} formulas={adjacentOmmlGroupByMember.Count}");
            }

            var traceObjectCounts = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_COUNTS"),
                "1",
                StringComparison.Ordinal);
            var tracePerformance = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF"),
                "1",
                StringComparison.Ordinal);
            var orderedTargets = useForwardSourceBookmarks
                && forwardSourceBookmarks.Count == plan.Targets.Count
                ? plan.Targets.OrderBy(item => item.SourceStart)
                : plan.Targets.OrderByDescending(item => item.SourceStart);
            foreach (var target in orderedTargets)
            {
                UndoRecord? formulaUndoRecord = null;
                string? createdTargetBookmarkName = null;
                VisualTeXRollbackSnapshot? visualTeXRollbackSnapshot = null;
                var undoRecordEnded = false;
                var mutationStarted = false;
                var stage = "capture-rollback-snapshot";
                var perfWatch = tracePerformance
                    ? System.Diagnostics.Stopwatch.StartNew()
                    : null;
                long perfLastMs = 0;
                void TracePerf(string perfStage)
                {
                    if (perfWatch is null) return;
                    var totalMs = perfWatch.ElapsedMilliseconds;
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-perf stage={perfStage} formulaId={target.SourceFormulaId} numbered={target.Numbered} deltaMs={totalMs - perfLastMs} totalMs={totalMs}");
                    perfLastMs = totalMs;
                }
                try
                {
                    if (adjacentOmmlGroupByMember.TryGetValue(target.Id, out var adjacentGroup))
                    {
                        if (processedAdjacentOmmlTargets.Contains(target.Id))
                            continue;
                        var groupAnchor = adjacentGroup
                            .OrderByDescending(member => member.SourceStart)
                            .First();
                        if (!string.Equals(target.Id, groupAnchor.Id, StringComparison.Ordinal))
                            continue;
                        stage = "adjacent-omml-group";
                        if (ommlBatchSource is null)
                            throw new InvalidOperationException(
                                "The OMML batch source is unavailable for an adjacent MathType formula group.");
                        var convertedGroupCount = ApplyAdjacentMathTypeToOmmlGroup(
                            document,
                            plan,
                            adjacentGroup,
                            prepared,
                            ommlBatchSource);
                        result.FormulaCount += convertedGroupCount;
                        foreach (var member in adjacentGroup)
                            processedAdjacentOmmlTargets.Add(member.Id);
                        continue;
                    }

                    if (forwardSourceBookmarks.TryGetValue(target.Id, out var forwardBookmarkName))
                    {
                        Bookmarks? bookmarks = null;
                        Bookmark? bookmark = null;
                        Range? liveSourceRange = null;
                        try
                        {
                            bookmarks = document.Bookmarks;
                            if (!bookmarks.Exists(forwardBookmarkName))
                                throw new InvalidOperationException(
                                    $"Word lost the temporary source locator for formula {target.SourceFormulaId}.");
                            bookmark = bookmarks[forwardBookmarkName];
                            liveSourceRange = bookmark.Range;
                            target.SourceObjectId =
                                $"{RangeReferencePrefix}{liveSourceRange.Start}:{liveSourceRange.End}";
                            WordDoubleClickHook.TraceMessage(
                                $"format-conversion-forward-source-live formulaId={target.SourceFormulaId} range={liveSourceRange.Start}:{liveSourceRange.End}");
                        }
                        finally
                        {
                            Release(liveSourceRange);
                            Release(bookmark);
                            Release(bookmarks);
                        }
                    }
                    if (visualTeXRollbackBuffer is not null)
                    {
                        visualTeXRollbackSnapshot = CaptureVisualTeXRollbackSnapshot(
                            document,
                            visualTeXRollbackBuffer,
                            target);
                        document.Activate();
                    }
                    TracePerf("rollback-snapshot");
                    stage = "begin-undo";
                    formulaUndoRecord = BeginUndoRecord("VisualTeX Convert Formula Format");
                    if (formulaUndoRecord is null)
                        throw new InvalidOperationException(
                            "Word 无法建立单公式转换撤销事务。为避免转换失败时丢失原公式，本次转换已停止。");

                    var formula = prepared[target.Id];
                    // MathType replacements have an exact per-object VTMT bookmark,
                    // while VisualTeX OLE replacements bind their durable VTO_
                    // identity bookmark before InsertOle returns. Both can therefore
                    // be verified locally. Re-enumerating document.InlineShapes for
                    // every target is O(N^2) and eventually destabilizes Word after
                    // dozens of delete/insert mutations. OMML still uses the older
                    // count stabilization until it gets an equivalent local anchor.
                    var targetObjectCountBefore = targetIsMathType || targetIsVisualTeX
                        ? -1
                        : CountSimpleFormatObjects(document, plan.TargetMode);
                    if (traceObjectCounts)
                        TraceSimpleFormatObjectCounts(document, target, "before-delete");
                    mutationStarted = true;
                    stage = "delete-source";
                    var sourceIsOmml = string.Equals(
                        plan.SourceMode,
                        FormulaOleContract.WordOmmlMode,
                        StringComparison.Ordinal);
                    var useDirectSingleNativeOmmlDelete = targetIsMathType
                        && plan.Targets.Count == 1
                        && sourceIsOmml
                        && !target.Numbered
                        && !target.SourceIsManagedOmml
                        && !HasLocalVisualTeXOmmlAnchor(document, target);
                    var useDirectSingleManagedNumberedOmmlDelete = targetIsMathType
                        && plan.Targets.Count == 1
                        && sourceIsOmml
                        && target.Numbered
                        && target.SourceIsManagedOmml
                        && string.Equals(
                            target.DisplayMode,
                            "block",
                            StringComparison.Ordinal);
                    var preserveFormulaCrossReferences =
                        referenceAliasesByTargetId.ContainsKey(target.Id);
                    var insertionStart = useDirectSingleNativeOmmlDelete
                        ? DeleteSingleNativeOmmlSourceDirect(document, target)
                        : useDirectSingleManagedNumberedOmmlDelete
                            ? DeleteSingleManagedNumberedOmmlSourceDirect(
                                document,
                                target,
                                sourceReferenceCounts,
                                preserveFormulaCrossReferences)
                            : DeleteSimpleSourceHost(
                                document,
                                plan.SourceMode,
                                target,
                                sourceReferenceCounts,
                                preserveFormulaCrossReferences);
                    if (useDirectSingleNativeOmmlDelete)
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-direct-omml-delete formulaId={target.SourceFormulaId} start={insertionStart}");
                    else if (useDirectSingleManagedNumberedOmmlDelete)
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-direct-numbered-omml-delete formulaId={target.SourceFormulaId} start={insertionStart}");
                    TracePerf("delete-source-host");
                    if (forwardSourceBookmarks.TryGetValue(target.Id, out var liveSourceBookmark))
                        EnsureForwardMathTypeSourceRemoved(
                            document,
                            liveSourceBookmark,
                            target,
                            "delete-source");
                    else
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
                    if (targetIsMathType)
                    {
                        createdTargetBookmarkName =
                            "VTMT_" + target.Id.Replace("-", string.Empty);
                        ResolvedEquationHeadingScope? preResolvedHeadingScope = null;
                        if (target.Numbered && batchHeadingScopes is not null)
                            batchHeadingScopes.TryGetValue(
                                target.SourceStart,
                                out preResolvedHeadingScope);
                        var batchNativePreview = formula.MathTypeNativePreview;
                        InsertMathTypeOle(
                            session,
                            formula.MathMl!,
                            formula.EmfPath!,
                            createdTargetBookmarkName,
                            preResolvedHeadingScope,
                            preparedHeadingScopeStarts,
                            batchNativePreview?.WmfPath,
                            batchNativePreview?.WidthPt ?? 0,
                            batchNativePreview?.HeightPt ?? 0,
                            batchNativePreview?.WordPosition ?? 0,
                            formula.MathTypeNativePreviewAttempted,
                            reuseExistingInlineTypingBoundary:
                                useDirectSingleNativeOmmlDelete,
                            updateCreatedMathTypeNumberFields:
                                canFinalizeSingleMathTypeNumberLocally);
                    }
                    else if (string.Equals(
                                 plan.TargetMode,
                                 FormulaOleContract.WordOmmlMode,
                                 StringComparison.Ordinal))
                    {
                        // Keep format conversion on the same performance architecture as the
                        // pre-MathType numbered-formula work: materialize formula content first,
                        // then repair numbering once after the batch. Building a table + hidden
                        // SEQ caption + visible REF while MathType numbering structures still
                        // remain later in the document forces Word to repaginate/maintain fields
                        // on every item and turns a 50-formula conversion into O(N^2) work.
                        // Metadata still records Numbered=true, so the final numbering refresh
                        // below can rebuild the standard VisualTeX structure without changing
                        // formula semantics or the target OMML payload.
                        InsertOmml(
                            session,
                            formula.MathMl!,
                            deferNumberingLayout: true,
                            deferFinalFingerprint: true,
                            ommlBatchSource: ommlBatchSource,
                            preserveExistingDisplayParagraphBoundary: true);
                    }
                    else
                    {
                        InsertOle(
                            session,
                            formula.PngPath!,
                            formula.EmfPath!,
                            deferNumberingLayout: targetIsVisualTeX);
                    }
                    TracePerf("insert-target");
                    if (forwardSourceBookmarks.TryGetValue(target.Id, out var insertedSourceBookmark))
                        EnsureForwardMathTypeSourceRemoved(
                            document,
                            insertedSourceBookmark,
                            target,
                            "insert-target");
                    else
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
                        if (referenceAliasesByTargetId.ContainsKey(target.Id))
                            retainedMathTypeTargetBookmarks[target.Id] = createdTargetBookmarkName;
                    }
                    else if (targetIsVisualTeX)
                    {
                        EnsureBookmarkedVisualTeXTargetSurvived(
                            document,
                            session.FormulaId,
                            target,
                            insertionStart);
                    }
                    else
                    {
                        WaitForSimpleTargetObjectCountToStabilize(
                            document,
                            plan.TargetMode,
                            targetObjectCountBefore + 1,
                            target);
                    }
                    if (forwardSourceBookmarks.TryGetValue(target.Id, out var committedSourceBookmark))
                        EnsureForwardMathTypeSourceRemoved(
                            document,
                            committedSourceBookmark,
                            target,
                            "post-transaction");
                    else
                        EnsureSimpleFormatSourceRemoved(
                            document,
                            plan.SourceMode,
                            target,
                            "post-transaction");
                    if (useDirectSingleManagedNumberedOmmlDelete)
                    {
                        // Keep the managed OMML metadata intact until the new
                        // Equation.DSMT4 object has survived its Word transaction.
                        // That preserves the mature Undo rollback path if insertion
                        // fails; only the now-dead source identity is removed here.
                        stage = "delete-source-metadata";
                        WordOmmlFormulaStore.Delete(
                            document,
                            target.SourceFormulaId);
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-direct-numbered-omml-metadata-deleted formulaId={target.SourceFormulaId}");
                    }
                    TracePerf("verify-target");
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
                            RemoveResidualFormatConversionBridgeAfterRollback(
                                document,
                                plan.SourceMode,
                                target);
                            ValidateSimpleSourceHost(document, plan.SourceMode, target);
                        }
                        catch (Exception restoreError)
                        {
                            if (visualTeXRollbackSnapshot is null)
                                throw new InvalidOperationException(
                                    $"公式“{target.Latex}”转换失败；Word 执行了撤销，但原公式宿主或临时 LaTeX bridge 没有完整恢复。",
                                    new AggregateException(error, restoreError));
                            try
                            {
                                RestoreVisualTeXRollbackSnapshot(
                                    document,
                                    visualTeXRollbackSnapshot,
                                    target);
                                ValidateSimpleSourceHost(document, plan.SourceMode, target);
                                RemoveResidualFormatConversionBridgeAfterRollback(
                                    document,
                                    plan.SourceMode,
                                    target);
                                ValidateSimpleSourceHost(document, plan.SourceMode, target);
                            }
                            catch (Exception snapshotRestoreError)
                            {
                                WordDoubleClickHook.TraceMessage(
                                    $"format-conversion-snapshot-restore-failed formulaId={target.SourceFormulaId} error={snapshotRestoreError}");
                                throw new InvalidOperationException(
                                    $"公式“{target.Latex}”转换失败；Word 撤销和 VisualTeX 结构快照恢复都未能完整恢复原公式宿主且清除临时 LaTeX bridge。",
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
                        && !string.IsNullOrWhiteSpace(createdTargetBookmarkName)
                        && !retainedMathTypeTargetBookmarks.ContainsKey(target.Id))
                        TryDeleteBookmark(document, createdTargetBookmarkName);
                    if (document is not null
                        && forwardSourceBookmarks.TryGetValue(target.Id, out var sourceLocatorBookmark))
                        TryDeleteBookmark(document, sourceLocatorBookmark);
                    if (visualTeXRollbackSnapshot is not null)
                        Release(visualTeXRollbackSnapshot.Payload);
                    if (!undoRecordEnded) EndUndoRecord(formulaUndoRecord);
                    Release(formulaUndoRecord);
                }
            }

            var finalizeWatch = tracePerformance
                ? System.Diagnostics.Stopwatch.StartNew()
                : null;
            long finalizeLastMs = 0;
            void TraceFinalize(string finalizeStage)
            {
                if (finalizeWatch is null) return;
                var totalMs = finalizeWatch.ElapsedMilliseconds;
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-finalize-perf stage={finalizeStage} sourceMode={plan.SourceMode} targetMode={plan.TargetMode} deltaMs={totalMs - finalizeLastMs} totalMs={totalMs}");
                finalizeLastMs = totalMs;
            }

            if (result.FailedFormulaCount == 0)
            {
                if (forwardSourceBookmarks.Count == plan.Targets.Count)
                {
                    var finalSourceObjectCount = CountSimpleFormatObjects(document, plan.SourceMode);
                    var expectedSourceObjectCount = Math.Max(
                        0,
                        initialSourceObjectCount - result.FormulaCount);
                    if (finalSourceObjectCount != expectedSourceObjectCount)
                    {
                        result.FailedFormulaCount++;
                        result.Failures.Add(
                            $"Source formula count mismatch after forward conversion: expected {expectedSourceObjectCount}, actual {finalSourceObjectCount}.");
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-source-count-mismatch sourceMode={plan.SourceMode} initial={initialSourceObjectCount} converted={result.FormulaCount} expected={expectedSourceObjectCount} actual={finalSourceObjectCount}");
                    }
                }
                else
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
            }
            TraceFinalize("source-residual-check");
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
            TraceFinalize("target-count-check");
            var convertedOmmlFormulaIds = targetIsOmml
                ? plan.Targets
                    .Select(target => prepared[target.Id].Session.FormulaId)
                    .Where(formulaId => !string.IsNullOrWhiteSpace(formulaId))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray()
                : Array.Empty<string>();
            var convertedOmmlNumberingMetadata = targetIsOmml
                ? plan.Targets
                    .Where(target => target.Numbered
                        && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                    .Select(target => prepared[target.Id].Session.ToMetadata())
                    .ToArray()
                : Array.Empty<FormulaMetadata>();
            if (result.FormulaCount > 0)
            {
                if (targetIsMathType)
                {
                    if (!canFinalizeSingleMathTypeNumberLocally)
                        MathTypeEquationNumbering.UpdateEquationNumbers(document);
                    else
                        WordDoubleClickHook.TraceMessage(
                            "format-conversion-numbering-local-mathtype-single finalized=1");
                }
                else if (targetIsOmml)
                {
                    WordEquationNumbering.RestoreEquationNumberFormatForConversion(
                        document,
                        plan.NumberFormatId);
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-number-format-restored format={plan.NumberFormatId}");
                    // The conversion loop materializes only OMML content. Finalize the known
                    // numbered subset end-to-start, then update all generated SEQ/REF fields in
                    // one batch. This mirrors the pre-1.2.5 numbering strategy: local structural
                    // work on known formulas, one incremental refresh, full Reconcile only as a
                    // malformed-document fallback.
                    if (WordEquationNumbering.TryBuildConvertedOmmlNumberingBatch(
                            document,
                            convertedOmmlNumberingMetadata,
                            out var builtNumbered)
                        && WordEquationNumbering.TryFinalizeHealthyConversionNumbering(
                            document,
                            out var finalizedNumbered))
                    {
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-numbering-local-batch targetMode={plan.TargetMode} built={builtNumbered} finalized={finalizedNumbered}");
                    }
                    else
                    {
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-numbering-local-batch-fallback targetMode={plan.TargetMode}");
                        var fallbackFinalizedNumbered = WordEquationNumbering.UpdateEquationNumbers(document);
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-numbering-fallback-finalized targetMode={plan.TargetMode} numbered={fallbackFinalizedNumbered}");
                    }
                }
                else
                {
                    WordEquationNumbering.TryReconcile(document);
                }
            }
            if (referenceAliasesByTargetId.Count > 0 && result.FailedFormulaCount == 0)
            {
                try
                {
                    var restoredAliases = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var entry in referenceAliasesByTargetId)
                    {
                        int restored;
                        if (targetIsMathType)
                        {
                            if (!retainedMathTypeTargetBookmarks.TryGetValue(
                                    entry.Key,
                                    out var targetBookmarkName))
                                throw new InvalidDataException(
                                    $"Missing converted MathType target locator for reference aliases on {entry.Key}.");
                            restored = MathTypeEquationReferences.RestoreFormatConversionAliasesToMathType(
                                document,
                                targetBookmarkName,
                                entry.Value);
                        }
                        else
                        {
                            if (!prepared.TryGetValue(entry.Key, out var convertedFormula))
                                throw new InvalidDataException(
                                    $"Missing converted VisualTeX payload for reference aliases on {entry.Key}.");
                            restored = MathTypeEquationReferences.RestoreFormatConversionAliasesToVisualTeX(
                                document,
                                convertedFormula.Session.FormulaId,
                                entry.Value);
                        }
                        if (restored != entry.Value.Count)
                            throw new InvalidDataException(
                                $"Restored {restored}/{entry.Value.Count} equation-reference aliases for conversion target {entry.Key}.");
                        foreach (var alias in entry.Value)
                            restoredAliases.Add(alias.Name);
                    }
                    var refreshedReferences = MathTypeEquationReferences.RefreshReferences(
                        document,
                        restoredAliases);
                    var expectedFormattingCount = capturedReferenceFormatting.Values.Sum(items => items.Count);
                    var restoredFormattingCount =
                        MathTypeEquationReferences.RestoreReferenceCharacterFormatting(
                            document,
                            capturedReferenceFormatting);
                    if (restoredFormattingCount != expectedFormattingCount)
                        throw new InvalidDataException(
                            $"Restored {restoredFormattingCount}/{expectedFormattingCount} equation-reference formatting snapshots.");
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-reference-aliases-restored sourceMode={plan.SourceMode} targetMode={plan.TargetMode} aliases={restoredAliases.Count} refreshedReferences={refreshedReferences} restoredFormatting={restoredFormattingCount}");
                }
                catch (Exception referenceAliasError)
                {
                    result.FailedFormulaCount++;
                    result.Failures.Add(
                        $"Equation-reference compatibility finalization failed: {referenceAliasError.Message}");
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-reference-alias-finalize-failed error={referenceAliasError}");
                }
            }
            TraceFinalize("numbering-reconcile");
            if (targetIsOmml
                && result.FormulaCount > 0
                && result.FailedFormulaCount == 0)
            {
                try
                {
                    var refreshed = WordOmmlNativeSource.RefreshFingerprintsFromDocumentOpenXml(
                        document,
                        convertedOmmlFormulaIds);
                    if (refreshed != convertedOmmlFormulaIds.Length)
                        throw new InvalidDataException(
                            $"OMML fingerprint finalization refreshed {refreshed}/{convertedOmmlFormulaIds.Length} converted formulas.");
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-omml-fingerprints-finalized count={refreshed}");
                }
                catch (Exception fingerprintError)
                {
                    result.FailedFormulaCount++;
                    result.Failures.Add(
                        $"OMML fingerprint finalization failed: {fingerprintError.Message}");
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-omml-fingerprint-finalize-failed error={fingerprintError}");
                }
            }
            TraceFinalize("omml-fingerprint-refresh");
            return result;
        }
        finally
        {
            if (document is not null && retainedMathTypeTargetBookmarks.Count > 0)
            {
                foreach (var bookmarkName in retainedMathTypeTargetBookmarks.Values)
                {
                    try { TryDeleteBookmark(document, bookmarkName); } catch { }
                }
            }
            if (document is not null && forwardSourceBookmarks.Count > 0)
            {
                foreach (var bookmarkName in forwardSourceBookmarks.Values)
                {
                    try { TryDeleteBookmark(document, bookmarkName); } catch { }
                }
            }
            try { document?.Activate(); } catch { }
            ommlBatchSource?.Dispose();
            if (paginationSuspended)
            {
                try { _application.Options.Pagination = previousPagination; } catch { }
            }
            if (batchViewSuspended && batchWindow is not null)
            {
                try { batchWindow.View.Type = previousViewType; } catch { }
            }
            Release(batchWindow);
            if (screenUpdatingSuspended)
            {
                try { _application.ScreenUpdating = previousScreenUpdating; } catch { }
            }
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
                    target.SourceObjectId,
                    allowGlobalFallback: false)
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

    internal static void RemoveResidualFormatConversionBridgeAfterRollback(
        Document document,
        string sourceMode,
        WordFormulaFormatConversionTarget target)
    {
        if (!string.Equals(sourceMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal)
            && !string.Equals(sourceMode, FormulaOleContract.WordOmmlMode, StringComparison.Ordinal))
            return;

        var bridge = BuildFormulaLatexSource(target.Metadata);
        var normalizedBridge = NormalizeFormulaToLatexVerificationText(bridge);
        if (normalizedBridge.Length == 0) return;

        InlineShape? shape = null;
        Range? sourceRange = null;
        Table? table = null;
        Range? hostRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? content = null;
        Range? candidate = null;
        Range? gap = null;
        Range? residual = null;
        try
        {
            if (string.Equals(sourceMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal))
            {
                shape = FindByFormulaId(
                        document,
                        target.SourceFormulaId,
                        target.SourceObjectId)
                    ?? throw new InvalidOperationException(
                        "The VisualTeX source formula was not present while validating rollback bridge cleanup.");
                sourceRange = shape.Range.Duplicate;
            }
            else
            {
                sourceRange = ResolveSimpleOmmlSourceRange(document, target)
                    ?? throw new InvalidOperationException(
                        "The Word OMML source formula was not present while validating rollback bridge cleanup.");
            }

            hostRange = sourceRange.Duplicate;
            if (target.Numbered
                && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
            {
                table = TryGetVisualTeXNumberedTable(sourceRange, target.Metadata);
                if (table is not null)
                {
                    Release(hostRange);
                    hostRange = table.Range.Duplicate;
                }
            }

            paragraphs = sourceRange.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
            }

            content = document.Content;
            var sourceStart = sourceRange.Start;
            var sourceEnd = sourceRange.End;
            var hostStart = hostRange.Start;
            var hostEnd = hostRange.End;
            var paragraphStart = paragraphRange?.Start ?? sourceStart;
            var paragraphEnd = paragraphRange?.End ?? sourceEnd;
            var anchorStart = Math.Min(Math.Min(sourceStart, hostStart), paragraphStart);
            var anchorEnd = Math.Max(Math.Max(sourceEnd, hostEnd), paragraphEnd);
            const int MaxAdornmentLength = 4;
            var windowStart = Math.Max(
                content.Start,
                anchorStart - normalizedBridge.Length - MaxAdornmentLength);
            var windowEnd = Math.Min(
                content.End,
                anchorEnd + normalizedBridge.Length + MaxAdornmentLength);
            var matches = new List<(int Start, int End)>();

            // Do not derive Word Range coordinates from Range.Text offsets here.
            // Around OLE/field runs Word exposes hidden field characters in the
            // document coordinate space that are absent from Range.Text, so the
            // two indices are not interchangeable. Probe the actual Word Range
            // coordinates directly instead.
            for (var candidateStart = windowStart;
                 candidateStart + bridge.Length <= windowEnd;
                 candidateStart++)
            {
                Release(candidate);
                candidate = document.Range(
                    candidateStart,
                    candidateStart + bridge.Length);
                if (!string.Equals(
                        NormalizeFormulaToLatexVerificationText(candidate.Text ?? string.Empty),
                        normalizedBridge,
                        StringComparison.Ordinal))
                    continue;

                var candidateEnd = candidateStart + bridge.Length;
                // A rollback bridge is transaction-local text adjacent to the
                // restored source host. Never delete an occurrence that overlaps
                // the restored formula itself or ordinary prose farther away.
                int gapStart;
                int gapEnd;
                if (candidateEnd <= hostStart)
                {
                    gapStart = candidateEnd;
                    gapEnd = hostStart;
                }
                else if (candidateStart >= hostEnd)
                {
                    gapStart = hostEnd;
                    gapEnd = candidateStart;
                }
                else
                {
                    continue;
                }
                if (gapEnd - gapStart > MaxAdornmentLength) continue;
                Release(gap);
                gap = document.Range(gapStart, gapEnd);
                if (!IsRollbackParagraphAdornment(gap.Text)) continue;
                matches.Add((candidateStart, candidateEnd));
            }

            if (matches.Count == 0) return;
            if (matches.Count > 1)
                throw new InvalidDataException(
                    "Rollback restored the source formula but left multiple adjacent LaTeX bridge candidates; cleanup was refused to avoid deleting user text.");

            var match = matches[0];
            residual = document.Range(match.Start, match.End);
            if (!IsRollbackBridgeResidual(residual.Text, bridge))
                throw new InvalidDataException(
                    "The rollback LaTeX bridge changed before cleanup and was not deleted.");
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-rollback-remove-latex-bridge formulaId={target.SourceFormulaId} range={match.Start}:{match.End} length={bridge.Length}");
            residual.Delete();
            Release(residual);
            residual = null;

            // Re-scan the same bounded neighborhood. A successful rollback is
            // not complete merely because the OLE/OMML source came back; none of
            // this conversion's temporary bridge text may survive beside it.
            windowEnd = Math.Min(content.End, windowEnd);
            for (var candidateStart = windowStart;
                 candidateStart + bridge.Length <= windowEnd;
                 candidateStart++)
            {
                Release(candidate);
                candidate = document.Range(
                    candidateStart,
                    candidateStart + bridge.Length);
                if (string.Equals(
                        NormalizeFormulaToLatexVerificationText(candidate.Text ?? string.Empty),
                        normalizedBridge,
                        StringComparison.Ordinal))
                    throw new InvalidDataException(
                        "A temporary LaTeX bridge still remains beside the restored formula after rollback cleanup.");
            }
        }
        finally
        {
            Release(residual);
            Release(gap);
            Release(candidate);
            Release(content);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(hostRange);
            Release(table);
            Release(sourceRange);
            Release(shape);
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
            if (string.Equals(mode, FormulaOleContract.WordOmmlMode, StringComparison.Ordinal))
            {
                Release(shape);
                shape = null;
                Release(shapes);
                shapes = null;
                OMaths? maths = null;
                try
                {
                    maths = document.OMaths;
                    return maths.Count;
                }
                finally { Release(maths); }
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

    private static void EnsureBookmarkedVisualTeXTargetSurvived(
        Document document,
        string formulaId,
        WordFormulaFormatConversionTarget target,
        int insertionStart)
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        Exception? lastError = null;
        while (watch.ElapsedMilliseconds < 750)
        {
            InlineShape? shape = null;
            try
            {
                shape = FindByFormulaId(
                    document,
                    formulaId,
                    sourceObjectIdHint: null,
                    allowGlobalFallback: false);
                if (shape is null)
                {
                    // Word can collapse a bookmark boundary when two embedded OLE
                    // objects are directly adjacent and the left object is replaced
                    // after the right one. The freshly inserted VisualTeX OLE still
                    // owns the correct cache FormulaId, so recover only in a small
                    // neighborhood around the exact insertion point and immediately
                    // rebind the durable VTO_ bookmark. The session FormulaId is a
                    // newly generated GUID, therefore a unique native OLE carrying
                    // that exact cached id is sufficient proof that the target
                    // survived; do not let a second unstable bookmark COM read turn
                    // a successful adjacent replacement into a false rollback.
                    shape = FindFreshVisualTeXTargetNearPosition(
                        document,
                        insertionStart,
                        formulaId);
                    if (shape is not null)
                    {
                        BindOleIdentityBookmark(shape, formulaId);
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-bookmarked-visualtex-target-stable sourceFormulaId={target.SourceFormulaId} formulaId={formulaId} localAnchorRepair=True elapsedMs={watch.ElapsedMilliseconds}");
                        return;
                    }
                }
                if (shape is not null && WordFormulaMetadataReader.IsNativeOle(shape))
                {
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata is not null
                        && string.Equals(metadata.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase))
                    {
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-bookmarked-visualtex-target-stable sourceFormulaId={target.SourceFormulaId} formulaId={formulaId} localAnchorRepair=False elapsedMs={watch.ElapsedMilliseconds}");
                        return;
                    }
                }
                lastError = new InvalidOperationException(
                    "The VisualTeX identity bookmark did not resolve to the exact converted OLE object.");
            }
            catch (Exception error)
            {
                lastError = error;
            }
            finally { Release(shape); }
            Thread.Sleep(25);
        }
        throw new InvalidOperationException(
            "Word did not retain the exact VisualTeX replacement object after the formula transaction.",
            lastError);
    }

    private static InlineShape? FindFreshVisualTeXTargetNearPosition(
        Document document,
        int insertionStart,
        string formulaId)
    {
        Range? content = null;
        Range? localRange = null;
        InlineShapes? shapes = null;
        InlineShape? result = null;
        try
        {
            content = document.Content;
            var start = Math.Max(content.Start, insertionStart - 2);
            var end = Math.Min(content.End, insertionStart + 64);
            if (end <= start) return null;
            localRange = document.Range(start, end);
            shapes = localRange.InlineShapes;
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-local-vt-probe formulaId={formulaId} insertionStart={insertionStart} range={start}:{end} shapes={shapes.Count}");
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(candidate)) continue;
                    var cached = WordFormulaMetadataReader.TryReadCachedPreview(candidate);
                    Range? candidateRange = null;
                    try
                    {
                        candidateRange = candidate.Range;
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-local-vt-candidate formulaId={formulaId} index={index}/{shapes.Count} range={candidateRange.Start}:{candidateRange.End} cachedFormulaId={cached?.FormulaId ?? "<null>"}");
                    }
                    finally { Release(candidateRange); }
                    if (!string.Equals(
                            cached?.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (result is not null)
                        throw new InvalidOperationException(
                            "More than one fresh VisualTeX target near the insertion point carries the same FormulaId.");
                    result = candidate;
                    candidate = null;
                }
                finally { Release(candidate); }
            }
            return result;
        }
        catch
        {
            Release(result);
            throw;
        }
        finally
        {
            Release(shapes);
            Release(localRange);
            Release(content);
        }
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
            var ommlCount = CountSimpleFormatObjects(
                document,
                FormulaOleContract.WordOmmlMode);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-counts stage={stage} formulaId={target.SourceFormulaId} VT={visualTeXCount} MT={mathTypeCount} OMML={ommlCount} latex={target.Latex}");
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
                // Batch conversion processes targets from the end of the document
                // toward the start, so the captured source range remains a strong
                // local hint. Resolve the durable VTO_ bookmark first and then a
                // bounded neighborhood around that hint; never enumerate every OLE
                // just to prove one already-deleted source is absent.
                shape = FindByFormulaId(
                    document,
                    target.SourceFormulaId,
                    target.SourceObjectId,
                    allowGlobalFallback: false);
                if (shape is null) return false;
                Range? liveRange = null;
                try
                {
                    liveRange = shape.Range;
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-source-match formulaId={target.SourceFormulaId} liveRange={liveRange.Start}:{liveRange.End} sourceHint={target.SourceObjectId} latex={target.Latex}");
                }
                catch { }
                finally { Release(liveRange); }
                return true;
            }
            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                Range? ommlRange = null;
                try
                {
                    ommlRange = ResolveSimpleOmmlSourceRange(document, target);
                    return ommlRange is not null;
                }
                finally { Release(ommlRange); }
            }
            shape = FindMathTypeOleByRange(
                document,
                target.SourceObjectId,
                allowGlobalFallback: false);
            return shape is not null && MathTypeOleInterop.IsMathTypeOle(shape);
        }
        catch { return false; }
        finally
        {
            Release(shape);
            Release(shapes);
        }
    }

    private static Range? ResolveSimpleOmmlSourceRange(
        Document document,
        WordFormulaFormatConversionTarget target)
    {
        Bookmark? bookmark = null;
        try
        {
            bookmark = WordOmmlFormulaStore.FindByFormulaId(
                document,
                target.SourceFormulaId);
            if (bookmark is not null)
                return WordOmmlFormulaStore.GetEquationRange(bookmark);
            if (target.SourceIsManagedOmml)
                return null;
            return TryResolveOmmlRangeReference(document, target.SourceObjectId);
        }
        catch { return null; }
        finally { Release(bookmark); }
    }

    private static void EnsureForwardMathTypeSourceRemoved(
        Document document,
        string bookmarkName,
        WordFormulaFormatConversionTarget target,
        string stage)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        InlineShapes? shapes = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(bookmarkName))
            {
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-forward-source-check stage={stage} formulaId={target.SourceFormulaId} bookmarkMissing=True present=False");
                return;
            }

            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range;
            shapes = range.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    throw new InvalidOperationException(
                        $"Word still contains the MathType source formula after stage '{stage}'. The current formula conversion was rolled back.");
                }
                finally { Release(candidate); }
            }
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-forward-source-check stage={stage} formulaId={target.SourceFormulaId} range={range.Start}:{range.End} present=False");
        }
        finally
        {
            Release(shapes);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
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
        var ommlToMathType =
            string.Equals(sourceMode, FormulaOleContract.WordOmmlMode, StringComparison.Ordinal)
            && string.Equals(targetMode, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal);
        var mathTypeToOmml =
            string.Equals(sourceMode, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal)
            && string.Equals(targetMode, FormulaOleContract.WordOmmlMode, StringComparison.Ordinal);
        if (!visualTeXToMathType
            && !mathTypeToVisualTeX
            && !ommlToMathType
            && !mathTypeToOmml)
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
        if (string.Equals(
                targetMode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal))
            return;
        if (string.Equals(
                targetMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal)
            && formula.MathTypeNativePreviewAttempted
            && formula.MathTypeNativePreview is not null)
            return;
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
                        target.SourceObjectId,
                        allowGlobalFallback: false)
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

            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                shapeRange = ResolveSimpleOmmlSourceRange(document, target)
                    ?? throw new InvalidOperationException(
                        "The Word OMML source formula moved before conversion started.");
                OMaths? maths = null;
                try
                {
                    maths = shapeRange.OMaths;
                    if (maths.Count != 1)
                        throw new InvalidOperationException(
                            "The source range no longer contains exactly one Word OMath equation.");
                }
                finally { Release(maths); }
                if (target.Numbered
                    && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                {
                    table = TryGetVisualTeXNumberedTable(shapeRange, target.Metadata)
                        ?? throw new InvalidOperationException(
                            "The numbered OMML source no longer owns its VisualTeX numbering table.");
                    if (!IsSafeOmmlNumberingTableForConversion(table, shapeRange))
                        throw new InvalidOperationException(
                            "The OMML numbering table contains ordinary user content or another formula; conversion was refused before modifying the document.");
                }
                return;
            }

            shape = FindMathTypeOleByRange(
                    document,
                    target.SourceObjectId,
                    allowGlobalFallback: false)
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

    private static bool HasLocalVisualTeXOmmlAnchor(
        Document document,
        WordFormulaFormatConversionTarget target)
    {
        Range? content = null;
        Range? probe = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            content = document.Content;
            var start = Math.Max(content.Start, target.SourceStart - 3);
            var end = Math.Min(
                content.End,
                Math.Max(target.SourceStart + 3, target.SourceStart + 1));
            probe = document.Range(start, end);
            bookmarks = probe.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmark);
                bookmark = bookmarks[index];
                var name = bookmark.Name ?? string.Empty;
                if (name.StartsWith(
                        WordOmmlFormulaStore.BookmarkPrefix,
                        StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
        catch
        {
            // Direct deletion is only a performance optimization. If Word refuses
            // a local bookmark safety probe, fall back to the mature managed-OMML
            // replacement path rather than risk orphaning VisualTeX metadata.
            return true;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(probe);
            Release(content);
        }
    }

    private Range DeleteSingleNativeOmmlSourceRange(
        Document document,
        WordFormulaFormatConversionTarget target)
    {
        var sourceRange = ResolveSimpleOmmlSourceRange(document, target)
            ?? throw new InvalidOperationException(
                "The native Word OMML source moved before direct replacement.");
        if (target.SourceIsManagedOmml || target.Numbered)
        {
            Release(sourceRange);
            throw new InvalidOperationException(
                "Direct OMML deletion is restricted to one unnumbered native Word equation.");
        }
        return sourceRange;
    }

    private int DeleteSingleNativeOmmlSourceDirect(
        Document document,
        WordFormulaFormatConversionTarget target)
    {
        Range? sourceRange = null;
        try
        {
            sourceRange = DeleteSingleNativeOmmlSourceRange(document, target);
            var start = sourceRange.Start;
            sourceRange.Delete();
            return start;
        }
        finally { Release(sourceRange); }
    }

    private int DeleteSingleManagedNumberedOmmlSourceDirect(
        Document document,
        WordFormulaFormatConversionTarget target,
        IReadOnlyDictionary<string, int>? knownReferenceCounts,
        bool preserveCrossReferences)
    {
        if (!target.SourceIsManagedOmml
            || !target.Numbered
            || !string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
            throw new InvalidOperationException(
                "The managed numbered OMML fast path requires one numbered VisualTeX display OMath.");

        Range? sourceRange = null;
        Table? table = null;
        Range? convertedTableRange = null;
        try
        {
            sourceRange = ResolveSimpleOmmlSourceRange(document, target)
                ?? throw new InvalidOperationException(
                    "The managed numbered OMML source moved before direct replacement.");
            table = TryGetVisualTeXNumberedTable(sourceRange, target.Metadata)
                ?? throw new InvalidOperationException(
                    "The managed numbered OMML source lost its numbering table before direct replacement.");
            if (!IsSafeOmmlNumberingTableForConversion(table, sourceRange))
                throw new InvalidOperationException(
                    "The managed numbered OMML table contains user content or another formula and cannot use the direct replacement path.");

            if (!preserveCrossReferences)
            {
                WordEquationNumbering.FreezeFormulaCrossReferences(
                    document,
                    target.SourceFormulaId,
                    knownReferenceCounts);
            }
            WordEquationNumbering.RemoveFormulaNumberingArtifacts(
                document,
                target.SourceFormulaId);

            // Reuse Word's stable table-to-text dismantling operation, but skip the
            // mature temporary-LaTeX bridge. Replacing the converted 1x3 row with a
            // single paragraph mark removes the OMath and table in one local edit
            // while leaving an exact, ordinary insertion paragraph for MathType.
            object separator = WdTableFieldSeparator.wdSeparateByParagraphs;
            object nestedTables = false;
            convertedTableRange = table.ConvertToText(
                ref separator,
                ref nestedTables);
            var start = convertedTableRange.Start;
            convertedTableRange.Text = "\r";
            return start;
        }
        finally
        {
            Release(convertedTableRange);
            Release(table);
            Release(sourceRange);
        }
    }

    private int DeleteSimpleSourceHost(
        Document document,
        string sourceMode,
        WordFormulaFormatConversionTarget target,
        IReadOnlyDictionary<string, int>? knownReferenceCounts = null,
        bool preserveCrossReferences = false)
    {
        InlineShape? shape = null;
        Range? shapeRange = null;
        Table? table = null;
        Range? tableRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? contentRange = null;
        Bookmark? ommlBookmark = null;
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
                        target.SourceObjectId,
                        allowGlobalFallback: false)
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
                                target.SourceObjectId,
                                allowGlobalFallback: false)
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
                        ref mutationStarted,
                        knownReferenceCounts,
                        preserveCrossReferences);

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
                                    target.SourceObjectId,
                                    allowGlobalFallback: false)
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

            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                shapeRange = ResolveSimpleOmmlSourceRange(document, target)
                    ?? throw new InvalidOperationException(
                        "The Word OMML source formula moved before replacement.");
                var start = shapeRange.Start;
                if (target.Numbered
                    && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                {
                    table = TryGetVisualTeXNumberedTable(shapeRange, target.Metadata)
                        ?? throw new InvalidOperationException(
                            "The numbered OMML source lost its numbering table before replacement.");
                    if (!IsSafeOmmlNumberingTableForConversion(table, shapeRange))
                        throw new InvalidOperationException(
                            "The OMML numbering table contains ordinary user content or another formula; conversion was refused before modifying the document.");
                    tableRange = table.Range.Duplicate;
                    start = tableRange.Start;
                }

                ommlBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    document,
                    target.SourceFormulaId);
                var latexSource = BuildFormulaLatexSource(target.Metadata);
                var sourceTarget = new FormulaToLatexTarget
                {
                    Metadata = target.Metadata,
                    ObjectMode = FormulaOleContract.WordOmmlMode,
                    LatexSource = latexSource,
                    Start = shapeRange.Start,
                    End = shapeRange.End,
                    FormulaRange = shapeRange,
                    OmmlBookmark = ommlBookmark,
                };
                shapeRange = null;
                ommlBookmark = null;
                var mutationStarted = false;
                try
                {
                    ConvertFormulaTargetToLatex(
                        document,
                        sourceTarget,
                        ref mutationStarted,
                        knownReferenceCounts,
                        preserveCrossReferences);
                    EnsureSimpleFormatSourceRemoved(
                        document,
                        sourceMode,
                        target,
                        "omml-latex-bridge");

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
                                "The temporary OMML LaTeX bridge changed before target insertion.");
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

            shape = FindMathTypeOleByRange(
                    document,
                    target.SourceObjectId,
                    allowGlobalFallback: false)
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
            Release(ommlBookmark);
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

    private static bool IsSafeOmmlNumberingTableForConversion(
        Table table,
        Range formulaRange)
    {
        Rows? rows = null;
        Columns? columns = null;
        Range? tableRange = null;
        OMaths? tableMaths = null;
        InlineShapes? tableShapes = null;
        Row? row = null;
        Range? rowRange = null;
        OMaths? rowMaths = null;
        try
        {
            columns = table.Columns;
            if (columns.Count != 3) return false;
            tableRange = table.Range;
            tableMaths = tableRange.OMaths;
            if (tableMaths.Count != 1) return false;
            tableShapes = tableRange.InlineShapes;
            if (tableShapes.Count != 0) return false;

            rows = table.Rows;
            var formulaRowFound = false;
            for (var index = 1; index <= rows.Count; index++)
            {
                Release(rowMaths);
                rowMaths = null;
                Release(rowRange);
                rowRange = null;
                Release(row);
                row = rows[index];
                rowRange = row.Range;
                var ownsFormula = formulaRange.Start >= rowRange.Start
                    && formulaRange.Start < rowRange.End;
                if (!ownsFormula)
                {
                    if (!IsStructurallyEmptyVisualTeXNumberingRow(rowRange))
                        return false;
                    continue;
                }

                if (formulaRowFound) return false;
                formulaRowFound = true;
                rowMaths = rowRange.OMaths;
                if (rowMaths.Count != 1) return false;
                if (!IsSafeOmmlNumberingRowOutsideFormula(rowRange, formulaRange))
                    return false;
            }
            return formulaRowFound;
        }
        finally
        {
            Release(rowMaths);
            Release(rowRange);
            Release(row);
            Release(tableShapes);
            Release(tableMaths);
            Release(tableRange);
            Release(columns);
            Release(rows);
        }
    }

    private static bool IsSafeOmmlNumberingRowOutsideFormula(
        Range rowRange,
        Range formulaRange)
    {
        if (formulaRange.Start < rowRange.Start || formulaRange.End > rowRange.End)
            return false;
        Document? document = null;
        Range? before = null;
        Range? after = null;
        try
        {
            document = rowRange.Document;
            before = document.Range(rowRange.Start, formulaRange.Start);
            after = document.Range(formulaRange.End, rowRange.End);
            return IsSafeMathTypeDisplayParagraph(before)
                && IsSafeMathTypeDisplayParagraph(after);
        }
        finally
        {
            Release(after);
            Release(before);
            Release(document);
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
