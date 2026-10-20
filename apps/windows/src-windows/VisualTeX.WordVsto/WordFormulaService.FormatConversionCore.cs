using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    public WordFormulaFormatConversionPlan CaptureFormulaFormatConversionPlan(
        bool wholeDocument,
        string sourceMode,
        string targetMode) =>
        CaptureFormulaFormatConversionPlanCore(
            wholeDocument,
            sourceMode,
            targetMode);

    public WordFormulaFormatConversionResult ApplyFormulaFormatConversionPlan(
        WordFormulaFormatConversionPlan plan,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared) =>
        ApplyFormulaFormatConversionPlanCore(
            plan,
            prepared);

    internal WordFormulaFormatConversionPlan CaptureFormulaFormatConversionPlanCore(
        bool wholeDocument,
        string sourceMode,
        string targetMode)
    {
        using var numberFormatScope =
            WordEquationNumberFormatAuthority.BeginOperationScope();
        ValidateCoreConversionPair(sourceMode, targetMode);

        Document? document = null;
        Selection? selection = null;
        Range? scope = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);

            scope = wholeDocument
                ? document.Content.Duplicate
                : (selection = _application.Selection).Range.Duplicate;

            var sourceKind = ObjectModeToHostKind(sourceMode);
            var targetKind = ObjectModeToHostKind(targetMode);
            IReadOnlyList<WordFormulaHostDescriptor> hosts;
            Dictionary<WordFormulaHostDescriptor, WordMathTypeHostAdapter.Semantic>? mathTypeSemantics = null;

            if (!wholeDocument
                && scope.Start == scope.End)
            {
                WordFormulaHostDescriptor? local =
                    sourceKind == WordFormulaHostKind.MathType
                        ? WordMathTypeHostAdapter.ResolveLocal(
                            _application,
                            document,
                            scope)
                        : WordFormulaHostLookup.ResolveLocal(
                            document,
                            scope,
                            sourceKind);
                hosts = local is null
                    ? Array.Empty<WordFormulaHostDescriptor>()
                    : new[] { local };
            }
            else if (sourceKind == WordFormulaHostKind.MathType)
            {
                mathTypeSemantics = new Dictionary<WordFormulaHostDescriptor, WordMathTypeHostAdapter.Semantic>();
                hosts = WordMathTypeHostAdapter.CaptureScopeIndex(
                    _application,
                    document,
                    scope,
                    mathTypeSemantics);
            }
            else
            {
                hosts =
                    WordFormulaHostResolver
                        .CaptureScopeKindIndex(
                            document,
                            scope,
                            sourceKind);
            }

            var targetMathTypeNumberPosition =
                string.Equals(
                    targetMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)
                    ? GetMathTypeNumberPositionPreference()
                    : "right";

            var plan = new WordFormulaFormatConversionPlan
            {
                DocumentId = DocumentIdentity(document),
                WritableValidated = true,
                SourceMode = sourceMode,
                TargetMode = targetMode,
                WholeDocument = wholeDocument,
                NumberFormatId =
                    WordEquationNumberFormatAuthority.Resolve(
                        document),
            };

            foreach (var host in hosts
                         .OrderBy(item => item.Range.Start))
            {
                ReportOperationProgress("正在读取源公式…", plan.Targets.Count, hosts.Count);
                var target = CaptureCoreConversionTarget(
                    document,
                    host,
                    sourceKind,
                    targetKind,
                    targetMathTypeNumberPosition,
                    mathTypeSemantics is null ? null : mathTypeSemantics[host]);
                plan.Targets.Add(target);
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

    internal WordFormulaFormatConversionResult ApplyFormulaFormatConversionPlanCore(
        WordFormulaFormatConversionPlan plan,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared)
    {
        using var numberFormatScope =
            WordEquationNumberFormatAuthority.BeginOperationScope();
        if (plan is null)
            throw new ArgumentNullException(nameof(plan));
        if (prepared is null)
            throw new ArgumentNullException(nameof(prepared));

        ValidateCoreConversionPair(
            plan.SourceMode,
            plan.TargetMode);

        var preparedWrites = new Dictionary<string, WordFormulaHostWriteRequest>(StringComparer.Ordinal);
        var resolvedSources = new Dictionary<string, WordFormulaHostDescriptor>(StringComparer.Ordinal);
        var nativeSources = new Dictionary<string, WordNativeConversionSource>(StringComparer.Ordinal);
        var nativeNumberReferences = new Dictionary<string, IReadOnlyDictionary<string, int>>(StringComparer.Ordinal);
        WordEquationReferenceFields.NativeOmmlReferenceIndex? nativeReferenceIndex = null;
        var mathTypeNumberAliases = new Dictionary<string, IReadOnlyList<EquationReferenceBookmarkAlias>>(StringComparer.Ordinal);
        IReadOnlyDictionary<string, int[]>? referencedBookmarks = null;
        WordNativeOmmlFieldRefresh? nativeNumberRefresh = null;
        WordFormulaNumberingKernel.VisualTeXConversionNumberPlan? visualTeXNumbers = null;
        WordNativeOmmlNumbering.NumberedOmmlBatch? numberedOmml = null;
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(
                document,
                plan.DocumentId);

            if (plan.Targets.Count == 0)
                return new WordFormulaFormatConversionResult();

            var missing = plan.Targets
                .Where(target => !prepared.ContainsKey(target.Id))
                .Select(target => target.Id)
                .ToArray();
            if (missing.Length > 0)
                throw new InvalidDataException(
                    "Prepared conversion output is missing for one or more formulas.");

            var sourceKind =
                ObjectModeToHostKind(plan.SourceMode);
            var targetKind =
                ObjectModeToHostKind(plan.TargetMode);
            var forwardOrder = sourceKind == WordFormulaHostKind.Omml
                || (sourceKind == WordFormulaHostKind.MathType && targetKind == WordFormulaHostKind.VisualTeX);

            // VisualTeX's numbered OLE uses a private VTEqNum_<FormulaId>
            // bookmark only while that host family exists. Pure OMML must remove
            // it, but existing user REF fields must not be stranded on the old
            // name. Capture external-reference counts before the transaction,
            // while the source bookmarks are still healthy. Generated visible
            // number REF fields are excluded by CaptureReferenceCounts.
            IReadOnlyDictionary<string, int> sourceReferenceCounts =
                sourceKind == WordFormulaHostKind.VisualTeX
                && targetKind == WordFormulaHostKind.Omml
                    ? WordEquationReferenceFields
                        .CaptureReferenceCounts(
                            document)
                    : new Dictionary<string, int>(
                        StringComparer.OrdinalIgnoreCase);

            // Normal format conversion deliberately uses Word's own Custom
            // UndoRecord as the rollback boundary. The older full-document
            // recovery snapshot parses and canonicalizes the complete Flat OPC
            // package before any mutation; that made unrelated legal package
            // parts part of the success path and could reject an otherwise valid
            // conversion. Keep the strict snapshot only as an opt-in diagnostic
            // mode for rollback investigations, never as a production prerequisite.
            WordDocumentEditSnapshot? rollbackSnapshot = null;
            if (plan.Targets.Count > 1
                && string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_STRICT_FORMAT_ROLLBACK_SNAPSHOT"),
                    "1",
                    StringComparison.Ordinal))
            {
                using var snapshotMetric =
                    WordOperationMetrics.Measure(
                        "FormatConversion.RollbackSnapshot.Capture");
                var ownedBookmarkNames =
                    plan.Targets
                        .Select(target => target.SourceFormulaId)
                        .Where(value => Guid.TryParse(value, out _))
                        .SelectMany(value =>
                            WordBookmarkRecoverySnapshot.NamesForFormula(value))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToArray();
                rollbackSnapshot =
                    new WordDocumentEditSnapshot(
                        document,
                        ownedBookmarkNames);
            }

            if (targetKind == WordFormulaHostKind.VisualTeX && plan.Targets.Any(item => item.Numbered))
                visualTeXNumbers = WordFormulaNumberingKernel.VisualTeXConversionNumberPlan.Capture(
                    document, plan.Targets.Where(item => item.Numbered).Select(item => (item.Id, item.SourceStart)), plan.NumberFormatId);

            var displaySectionTabs = targetKind == WordFormulaHostKind.VisualTeX
                ? CaptureConversionSectionTabStops(document)
                : Array.Empty<(int Start, int End, float Center, float Right)>();

            // Native source reads (including WordOpenXML) are preflight, not
            // target mutation. In Word, exporting that XML can close an active
            // Custom UndoRecord. Resolve and prepare all selected sources in this
            // same synchronous UI turn before starting the one batch record.
            foreach (var target in plan.Targets.OrderByDescending(item => item.SourceStart))
            {
                ReportOperationProgress("正在核对源公式并准备替换…", resolvedSources.Count, plan.Targets.Count);
                using var preparationMetric = WordOperationMetrics.Measure("FormatConversion.PrepareBoundSource." + sourceKind + "To" + targetKind);
                WordFormulaHostDescriptor source;
                if (sourceKind == WordFormulaHostKind.MathType)
                {
                    Range? captured = null;
                    try
                    {
                        captured = WordFormulaOperationLocator.ResolveCapturedRange(document, target.SourceObjectId);
                        var binding = WordMathTypeHostAdapter.BindCapturedHost(_application, document, captured);
                        nativeSources.Add(target.Id, binding);
                        source = binding.Host;
                        if (targetKind == WordFormulaHostKind.VisualTeX && target.Numbered)
                        {
                            // MathType's references (GOTOBUTTON/REF) name the
                            // ZEqnNum bookmark on its number, which is deleted
                            // with the MathType paragraph. Keep those names.
                            referencedBookmarks ??= MathTypeEquationReferences.CaptureReferencedBookmarks(document);
                            var aliases = referencedBookmarks.Count == 0
                                ? Array.Empty<EquationReferenceBookmarkAlias>()
                                : MathTypeEquationReferences.CaptureFormatConversionAliasesFromMathType(
                                    document, binding.MathTypeShape!, referencedBookmarks);
                            if (aliases.Count > 0)
                                mathTypeNumberAliases.Add(target.Id, aliases);
                        }
                        var semanticFont = FormulaFontSize.ResolveSemanticFontSize(source.Metadata!);
                        var currentFont = source.Display
                            ? semanticFont
                            : ReadMathTypeInlinePresentationFontSize(binding.MathTypeShape!) ?? semanticFont;
                        // One fresh read proves the input that produced the
                        // prepared render; it also supplies the physical owner
                        // consumed below. No second resolver/decoder is needed.
                        if (!string.Equals(binding.MathMl, target.SourceMathMl, StringComparison.Ordinal)
                            || !string.Equals(source.DisplayMode, target.DisplayMode, StringComparison.Ordinal)
                            || source.Numbering.Numbered != target.Numbered
                            || semanticFont != FormulaFontSize.ResolveSemanticFontSize(target.Metadata)
                            || currentFont != target.FontSizePt)
                        {
                            WordDoubleClickHook.TraceMessage(
                                $"format-conversion-source-changed start={target.SourceStart} mathMlEqual={string.Equals(binding.MathMl, target.SourceMathMl, StringComparison.Ordinal)} "
                                + $"display={target.DisplayMode}/{source.DisplayMode} numbered={target.Numbered}/{source.Numbering.Numbered} "
                                + $"semanticFont={FormulaFontSize.ResolveSemanticFontSize(target.Metadata)}/{semanticFont} presentationFont={target.FontSizePt}/{currentFont}");
                            throw new InvalidOperationException("The MathType source changed while its conversion was being prepared.");
                        }
                    }
                    finally { Release(captured); }
                }
                else if (sourceKind == WordFormulaHostKind.VisualTeX)
                {
                    Range? captured = null;
                    try
                    {
                        captured = WordFormulaOperationLocator.ResolveCapturedRange(document, target.SourceObjectId);
                        var binding = WordNativeConversionSource.BindVisualTeX(document, captured);
                        nativeSources.Add(target.Id, binding);
                        source = binding.Host;
                        var metadata = source.Metadata!;
                        var numbered = source.Display && (metadata.Numbered || source.Numbering.Numbered);
                        if (!string.Equals(metadata.FormulaId, target.SourceFormulaId, StringComparison.OrdinalIgnoreCase)
                            || !string.Equals(metadata.Latex, target.Latex, StringComparison.Ordinal)
                            || !string.Equals(source.DisplayMode, target.DisplayMode, StringComparison.Ordinal)
                            || numbered != target.Numbered
                            || FormulaFontSize.ResolveSemanticFontSize(metadata) != target.FontSizePt
                            || !string.Equals(metadata.FormulaLetterFont, target.Metadata.FormulaLetterFont, StringComparison.Ordinal)
                            || !string.Equals(metadata.FormulaChineseFont, target.Metadata.FormulaChineseFont, StringComparison.Ordinal))
                            throw new InvalidOperationException("The VisualTeX source changed while its conversion was being prepared.");
                    }
                    finally { Release(captured); }
                }
                else
                {
                    source = WordFormulaOperationLocator.ResolveCapturedHost(
                        _application, document, target.SourceObjectId, sourceKind, target.SourceFormulaId);
                    if (forwardOrder)
                        nativeSources.Add(target.Id, new WordNativeConversionSource(document, source));
                }
                resolvedSources.Add(target.Id, source);
                if (sourceKind == WordFormulaHostKind.Omml
                    && targetKind == WordFormulaHostKind.VisualTeX
                    && target.Numbered)
                {
                    // Word's _Ref target lives inside the OMath being deleted.
                    nativeReferenceIndex ??= WordEquationReferenceFields.NativeOmmlReferenceIndex.Capture(document);
                    var references = nativeReferenceIndex.CountInside(source);
                    if (references.Count > 0)
                        nativeNumberReferences.Add(target.Id, references);
                }
                if (targetKind is not (WordFormulaHostKind.Omml or WordFormulaHostKind.VisualTeX))
                    continue;
                var preparedFormula = prepared[target.Id];
                PrepareCoreTargetSession(preparedFormula.Session, plan, target);
                var request = BuildHostWriteRequest(preparedFormula.Session,
                    preparedFormula.MathMl, preparedFormula.PngPath, preparedFormula.EmfPath);
                request.BatchMutation = true;
                if (targetKind == WordFormulaHostKind.VisualTeX && target.Numbered)
                    request.PreparedVisualTeXNumbering = visualTeXNumbers!.For(target.Id);
                if (targetKind == WordFormulaHostKind.VisualTeX && source.Display && !source.WithinTable)
                {
                    var section = displaySectionTabs.First(item => item.Start <= source.Range.Start && item.End > source.Range.Start);
                    request.DisplayCenterTabPoints = section.Center;
                    request.DisplayRightTabPoints = section.Right;
                }
                preparedWrites.Add(target.Id, request);
                if (targetKind == WordFormulaHostKind.Omml)
                {
                    var display = string.Equals(request.DisplayMode, "block", StringComparison.OrdinalIgnoreCase);
                    request.PreparedOmml = WordOmmlConverter.TransformMathMlToOmml(request.MathMl!);
                    if (display && request.Numbered)
                    {
                        numberedOmml ??= WordNativeOmmlNumbering.NumberedOmmlBatch.Capture(
                            document,
                            plan.Targets.Where(item => item.Numbered).Select(item => item.SourceStart));
                        request.PreparedOmml = numberedOmml.Prepare(
                            source.Range.Start, request.PreparedOmml);
                    }
                    if (sourceKind == WordFormulaHostKind.MathType
                        || (sourceKind == WordFormulaHostKind.VisualTeX && !source.Display))
                    {
                        Range? exact = null;
                        try
                        {
                            exact = WordFormulaHostSemanticReader.CreateRange(document, source.Range);
                            request.PreparedOleReplacement = WordOleToOmmlReplacement.Prepare(
                                document, exact, request.PreparedOmml, display,
                                retiredIdentityBookmark: !Guid.TryParse(source.FormulaId, out _) ? null
                                    : sourceKind == WordFormulaHostKind.MathType
                                        ? WordFormulaIdentityStore.MathTypeBookmarkName(source.FormulaId!)
                                        : WordFormulaMetadataReader.IdentityBookmarkName(source.FormulaId!),
                                replaceMathTypeNumber: sourceKind == WordFormulaHostKind.MathType && source.Numbering.Numbered)
                                ?? throw new InvalidDataException("The conversion preflight does not own an exact external OLE.");
                        }
                        finally { Release(exact); }
                    }
                }
            }

            if (targetKind == WordFormulaHostKind.Omml && plan.Targets.Any(item => item.Numbered))
            {
                using var numberingMetric = WordOperationMetrics.Measure("FormatConversion.OmmlNumbering.Prepare");
                nativeNumberRefresh = WordNativeOmmlFieldRefresh.Prepare(document,
                    plan.Targets.Where(item => item.Numbered).Min(item => item.SourceStart));
            }

            // A numbered MathType target that opens a chapter/section scope needs
            // MathType's native section paragraph. Opening or closing the hidden
            // staging document ends Word's active custom UndoRecord, so prepare one
            // parameterizable section before the conversion transaction, as the
            // MathType bulk skeleton does. Declared first, it is closed last.
            using var mathTypeSection = targetKind == WordFormulaHostKind.MathType
                && plan.Targets.Any(item => item.Numbered)
                && EquationNumberFormat.Resolve(WordEquationNumberFormatAuthority.Resolve(document)).UsesHeading
                    ? new MathTypeSectionStaging(document, CreateMathTypeSectionFlatOpc(0, 0))
                    : null;

            // A visible Print Layout window re-lays out the unconverted OMML of
            // the page on screen after every host write, even with
            // ScreenUpdating off: ~1-5 s per formula on page 1 of a long paper.
            // A minimized window lays out nothing; geometry stays Print Layout
            // exact (Draft/Web views round inline objects to 0.5 pt).
            using var screenUpdating = NativeOmmlScreenUpdatingScope.Suspend(_application);
            using var minimized = plan.Targets.Count >= MinimizeWindowFormulaThreshold
                ? MinimizedWindowScope.Enter(document, OperationWindowMinimized)
                : null;
            return WordFormulaMutationTransaction.Execute(
                _application,
                document,
                "VisualTeX Convert Formula Format",
                () =>
                {
                    var result =
                        new WordFormulaFormatConversionResult();
                    var referenceTargetReplacements =
                        new Dictionary<string, string>(
                            StringComparer.OrdinalIgnoreCase);

                    var mutationOrder = forwardOrder
                        ? plan.Targets.OrderBy(item => item.SourceStart)
                        : plan.Targets.OrderByDescending(item => item.SourceStart);
                    foreach (var target in mutationOrder)
                    {
                        ReportOperationProgress("正在写入目标公式…", result.FormulaCount, plan.Targets.Count);
                        var preparedFormula =
                            prepared[target.Id];
                        var source = resolvedSources[target.Id];
                        nativeSources.TryGetValue(target.Id, out var sourceOwner);
                        sourceOwner?.RefreshAddresses();
                        WordFormulaMutationTransaction.TraceBoundary(_application, "conversion-preflight-source");

                        WordFormulaHostDescriptor? replacedHost =
                            null;
                        if (targetKind is
                            WordFormulaHostKind.Omml
                            or WordFormulaHostKind.VisualTeX)
                        {
                            var request = preparedWrites[target.Id];
                            replacedHost =
                                WordFormulaHostMutationKernel
                                    .ReplaceInActiveTransaction(
                                        _application,
                                        document,
                                        source,
                                        request,
                                        sourceAlreadyResolved: true,
                                        allowDirectCanonicalOmmlRemoval:
                                            sourceKind ==
                                            WordFormulaHostKind.Omml,
                                        deferTargetNumberFinalization: true,
                                        preparedSource: sourceOwner);
                            if (targetKind == WordFormulaHostKind.Omml && target.Numbered)
                                nativeNumberRefresh!.AddInserted(replacedHost);

                            if (sourceKind ==
                                    WordFormulaHostKind.VisualTeX
                                && targetKind ==
                                    WordFormulaHostKind.Omml
                                && target.Numbered
                                && Guid.TryParse(
                                    target.SourceFormulaId,
                                    out var sourceFormulaId))
                            {
                                var oldReferenceTarget =
                                    "VTEqNum_"
                                    + sourceFormulaId.ToString("N");
                                if (sourceReferenceCounts.TryGetValue(
                                        oldReferenceTarget,
                                        out var referenceCount)
                                    && referenceCount > 0)
                                {
                                    var nativeReferenceTarget =
                                        WordEquationReferenceFields
                                            .CreateNativeOmmlNumberReferenceBookmark(
                                                document,
                                                replacedHost);
                                    referenceTargetReplacements.Add(
                                        oldReferenceTarget,
                                        nativeReferenceTarget);
                                }
                            }
                        }
                        else if (targetKind ==
                                 WordFormulaHostKind.MathType)
                        {
                            if (sourceKind ==
                                WordFormulaHostKind.MathType)
                                throw new InvalidOperationException(
                                    "MathType-to-MathType is not a format conversion.");

                            var mathMl =
                                preparedFormula.MathMl
                                ?? target.SourceMathMl
                                ?? throw new InvalidDataException(
                                    $"MathType target '{target.Latex}' has no MathML.");

                            PrepareCoreTargetSession(
                                preparedFormula.Session,
                                plan,
                                target);

                            _ =
                                WordFormulaHostMutationKernel
                                    .ReplaceSourceWithExternalTargetInActiveTransaction(
                                        _application,
                                        document,
                                        source,
                                        insertion =>
                                        {
                                            var session =
                                                preparedFormula.Session;
                                            var oldSourceObjectId =
                                                session.SourceObjectId;
                                            try
                                            {
                                                session.SourceObjectId =
                                                    RangeReference(
                                                        insertion);
                                                var native =
                                                    preparedFormula
                                                        .MathTypeNativePreview;
                                                return InsertMathTypeOle(
                                                    session,
                                                    mathMl,
                                                    preparedFormula.EmfPath,
                                                    isolatedNativePreviewWmfPath:
                                                        native?.WmfPath,
                                                    isolatedNativePreviewWidthPt:
                                                        native?.WidthPt ?? 0,
                                                    isolatedNativePreviewHeightPt:
                                                        native?.HeightPt ?? 0,
                                                    isolatedNativePreviewWordPosition:
                                                        native?.WordPosition ?? 0,
                                                    isolatedNativePreviewAttempted:
                                                        preparedFormula
                                                            .MathTypeNativePreviewAttempted,
                                                    preserveExistingDisplayParagraphBoundary:
                                                        string.Equals(
                                                            target.DisplayMode,
                                                            "block",
                                                            StringComparison.OrdinalIgnoreCase),
                                                    preserveCapturedInsertion:
                                                        true,
                                                    boundedBatchMutation:
                                                        true,
                                                    preparedStandaloneGenerated:
                                                        preparedFormula.MathTypeStandaloneGenerated,
                                                    preparedStandaloneCompoundFile:
                                                        preparedFormula.MathTypeStandaloneCompoundFile,
                                                    preparedStandaloneSemanticSignature:
                                                        preparedFormula.MathTypeStandaloneSemanticSignature,
                                                    preparedBatchMathTypeSection:
                                                        mathTypeSection?.Paragraph);
                                            }
                                            finally
                                            {
                                                session.SourceObjectId =
                                                    oldSourceObjectId;
                                            }
                                        },
                                        sourceAlreadyResolved: true,
                                        allowDirectCanonicalOmmlRemoval:
                                            sourceKind ==
                                            WordFormulaHostKind.Omml,
                                        preparedSource: sourceOwner);
                        }
                        else
                        {
                            throw new NotSupportedException(
                                $"Unsupported conversion target {targetKind}.");
                        }

                        // A source binding is consumed by this replacement;
                        // retaining deleted OLE owners until batch finalization
                        // unnecessarily extends their COM/storage lifetime.
                        sourceOwner?.Dispose();
                        result.FormulaCount++;
                        ReportOperationProgress("正在写入目标公式…", result.FormulaCount, plan.Targets.Count);
                    }

                    // OMML sources and MathType-to-VisualTeX follow live Word
                    // owners forward; other paths retain reverse order. Each edit
                    // not relayout the growing converted tail. Targets still cannot
                    // finalize their visible ordinals one host at a time: formulas
                    // that belong earlier in the document may not exist in the
                    // target family yet. Recompute the canonical Word order once,
                    // after the complete batch has materialized. This is the same
                    // numbering kernel used by the explicit refresh command; no
                    // conversion-specific numbering structure is introduced.
                    if (plan.Targets.Any(target =>
                            target.Numbered))
                    {
                        if (targetKind == WordFormulaHostKind.Omml)
                        {
                            // No post-mutation host/XML/heading rediscovery. The
                            // preflight suffix plan contains affected existing
                            // native equations plus the newly materialized hosts.
                            using var numberingMetric = WordOperationMetrics.Measure("FormatConversion.OmmlNumbering.Apply");
                            nativeNumberRefresh!.Apply(plan.Targets.Count(target => target.Numbered));
                        }
                        else if (targetKind == WordFormulaHostKind.VisualTeX)
                        {
                            visualTeXNumbers!.ApplyExistingTables(document);
                            _ = WordFormulaNumberingKernel.RefreshCanonicalNumbers(
                                document, kindFilter: targetKind,
                                minimumExpectedNumbered: plan.Targets.Count(target =>
                                    target.Numbered && !target.SourceWithinTable),
                                preparedVisualTeX: visualTeXNumbers);
                        }
                        else if (targetKind ==
                                 WordFormulaHostKind.MathType)
                        {
                            // Batch MathType insertion deliberately skips
                            // per-item nested field updates and document-wide
                            // nearest-template scans. Re-evaluate the native
                            // MTPlaceRef/MTEqn stream once, in final document
                            // order, after every target has materialized.
                            _ =
                                MathTypeEquationNumbering
                                    .UpdateEquationNumbers(
                                        document);
                        }
                    }

                    if (nativeNumberReferences.Count > 0)
                    {
                        // VisualTeX numbers are final only after the batch
                        // refresh above; bind references to the fresh hosts.
                        var convertedHosts = WordFormulaHostResolver.CaptureDocumentIndex(document).VisualTeX;
                        foreach (var pending in nativeNumberReferences)
                        {
                            var formulaId = preparedWrites[pending.Key].FormulaId;
                            var host = convertedHosts.Single(candidate => string.Equals(
                                candidate.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase));
                            WordFormulaHostMutationKernel.MigrateConvertedNumberReferences(document, host, pending.Value);
                        }
                    }

                    if (mathTypeNumberAliases.Count > 0)
                    {
                        // Bind each kept MathType name to the converted number:
                        // ZEqnNum owns "(1)", a number-only alias owns "1".
                        // The references themselves stay as the user has them.
                        var convertedHosts = WordFormulaHostResolver.CaptureDocumentIndex(document).VisualTeX;
                        var aliasNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var pending in mathTypeNumberAliases)
                        {
                            var formulaId = preparedWrites[pending.Key].FormulaId;
                            var host = convertedHosts.Single(candidate => string.Equals(
                                candidate.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase));
                            foreach (var alias in pending.Value)
                            {
                                WordVisualTeXParagraphNumbering.EnsureReferenceBookmark(
                                    document, host, alias.Name,
                                    includeParentheses: alias.Span == EquationReferenceBookmarkSpan.VisibleNumber);
                                aliasNames.Add(alias.Name);
                            }
                        }
                        _ = WordEquationReferenceFields.UpdateReferences(document, aliasNames);
                    }

                    if (referenceTargetReplacements.Count > 0)
                    {
                        _ =
                            WordEquationReferenceFields
                                .MigrateReferenceTargets(
                                    document,
                                    referenceTargetReplacements,
                                    sourceReferenceCounts);
                    }

                    return result;
                },
                rollbackSnapshot,
                verifyRollback: () =>
                {
                    foreach (var write in preparedWrites.Values)
                        write.PreparedOleReplacement?.VerifyRollback();
                });
        }
        finally
        {
            foreach (var nativeSource in nativeSources.Values) nativeSource.Dispose();
            visualTeXNumbers?.Dispose();
            nativeNumberRefresh?.Dispose();
            foreach (var write in preparedWrites.Values)
                write.PreparedOleReplacement?.Dispose();
            Release(document);
        }
    }

    private static (int Start, int End, float Center, float Right)[] CaptureConversionSectionTabStops(Document document)
    {
        using var metric = WordOperationMetrics.Measure("FormatConversion.CaptureSectionTabGeometry");
        var result = new List<(int Start, int End, float Center, float Right)>();
        Sections? sections = null;
        try
        {
            sections = document.Sections;
            for (var index = 1; index <= sections.Count; index++)
            {
                Section? section = null;
                Range? range = null;
                PageSetup? setup = null;
                try
                {
                    section = sections[index];
                    range = section.Range;
                    setup = section.PageSetup;
                    var tabs = WordEquationNumbering.CalculateEquationTabStops(
                        setup.PageWidth, setup.LeftMargin, setup.RightMargin, 0, 0);
                    result.Add((range.Start, range.End, tabs.Center, tabs.Right));
                }
                finally { Release(setup); Release(range); Release(section); }
            }
        }
        finally { Release(sections); }
        return result.ToArray();
    }

    private WordFormulaFormatConversionTarget CaptureCoreConversionTarget(
        Document document,
        WordFormulaHostDescriptor host,
        WordFormulaHostKind sourceKind,
        WordFormulaHostKind targetKind,
        string targetMathTypeNumberPosition,
        WordMathTypeHostAdapter.Semantic? capturedMathTypeSemantic = null)
    {
        using var operationMetric =
            WordOperationMetrics.Measure(
                "FormatConversion.CaptureTarget."
                + sourceKind
                + "To"
                + targetKind);
        FormulaMetadata metadata;
        string latex;
        string? mathMl = null;
        bool numberedIntent;

        if (sourceKind == WordFormulaHostKind.Omml)
        {
            var numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            host.Numbering = numbering;
            numberedIntent = numbering.Numbered;
            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    host);
            metadata =
                BuildReadOnlyOmmlSessionMetadata(
                    document,
                    host,
                    payload);
            latex = payload.Latex;
            mathMl = payload.MathMl;
        }
        else if (sourceKind ==
                 WordFormulaHostKind.VisualTeX)
        {
            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    host);
            metadata = CloneCoreMetadata(
                payload.Metadata
                ?? throw new InvalidDataException(
                    "A VisualTeX source has no embedded metadata."));
            metadata.DisplayMode =
                host.DisplayMode;

            // Embedded VisualTeX metadata is the durable user intent. Word can
            // damage/move the visible tab-number scaffold through bookmark
            // insertion gravity while the OLE itself still correctly says
            // Numbered=true. Preserve numbering when either the canonical
            // structure is healthy or the authoritative OLE metadata says the
            // display formula is numbered. Apply will reconcile a metadata-only
            // numbered source inside the same conversion transaction before
            // detaching/replacing it.
            // Rendering needs intent, not another physical scaffold proof.
            // Explicit numbered metadata already decides the existing OR rule;
            // inspect structure only when metadata does not decide it. Exact
            // numbering ownership is resolved once in write preflight.
            numberedIntent = host.Display && (metadata.Numbered
                || WordFormulaNumberingResolver.ResolveLocal(document, host).Numbered);
            latex = metadata.Latex;
        }
        else if (sourceKind ==
                 WordFormulaHostKind.MathType)
        {
            // Scope capture already read this exact native payload synchronously.
            // Collapsed local selection still performs the existing exact reader;
            // Apply re-resolves sources after rendering, independently of this data.
            var semantic = capturedMathTypeSemantic
                ?? WordMathTypeHostAdapter.ReadSemantic(
                    _application,
                    document,
                    host);
            metadata =
                CloneCoreMetadata(
                    semantic.Metadata);
            latex = metadata.Latex;
            mathMl = semantic.MathMl;
            numberedIntent = host.Numbering.Numbered;
        }
        else
        {
            throw new NotSupportedException(
                $"Unsupported conversion source {sourceKind}.");
        }

        if (string.IsNullOrWhiteSpace(latex))
            throw new InvalidDataException(
                "The source formula has no LaTeX representation.");

        var formulaId =
            Guid.TryParse(host.FormulaId, out var owned)
                ? owned.ToString("D")
                : Guid.TryParse(metadata.FormulaId, out var semanticId)
                    ? semanticId.ToString("D")
                    : Guid.NewGuid().ToString("D");
        metadata.FormulaId = formulaId;
        metadata.DisplayMode =
            host.DisplayMode;
        metadata.Numbered = numberedIntent;
        metadata.Validate();

        var semanticFontSize =
            FormulaFontSize.ResolveSemanticFontSize(
                metadata);
        var fontSize =
            semanticFontSize;

        if (sourceKind == WordFormulaHostKind.MathType
            && string.Equals(
                host.DisplayMode,
                "inline",
                StringComparison.OrdinalIgnoreCase)
            && targetKind is WordFormulaHostKind.Omml
                or WordFormulaHostKind.VisualTeX)
        {
            InlineShape? sourceShape = null;
            try
            {
                sourceShape = FindMathTypeOleByRange(
                    document,
                    RangeReferenceFromAddress(host.Range),
                    allowGlobalFallback: false);
                if (sourceShape is not null)
                {
                    var wordPresentationFontSize =
                        ReadMathTypeInlinePresentationFontSize(
                            sourceShape);
                    if (wordPresentationFontSize is > 0)
                    {
                        // For both pure OMML and VisualTeX, preserve the font
                        // size the user actually sees in Word. MathType's MTEF
                        // Full size can differ from its Word OLE presentation
                        // size (for example 12 pt internally but 10.5 pt in the
                        // document). Persisting the hidden MTEF size into
                        // VisualTeX metadata makes the next edit/apply jump back
                        // to 12 pt. Store the Word presentation size as the
                        // target's durable font size instead.
                        fontSize =
                            wordPresentationFontSize.Value;
                    }
                }
            }
            finally
            {
                Release(sourceShape);
            }
        }

        return new WordFormulaFormatConversionTarget
        {
            Id = Guid.NewGuid().ToString("D"),
            SourceFormulaId = formulaId,
            SourceObjectId =
                RangeReferenceFromAddress(host.Range),
            SourceStart = host.Range.Start,
            Latex = latex,
            SourceMathMl = mathMl,
            SourceIsManagedOmml =
                sourceKind == WordFormulaHostKind.Omml
                && !string.IsNullOrWhiteSpace(host.FormulaId),
            SourceWithinTable = host.WithinTable,
            DisplayMode = host.DisplayMode,
            Numbered = numberedIntent,
            // Placement is consumed only by a MathType destination. A source
            // MathType equation cannot also be that destination in this matrix.
            MathTypeNumberPosition = targetMathTypeNumberPosition,
            FontSizePt = fontSize,
            Metadata = metadata,
        };
    }

    private static float? ReadMathTypeInlinePresentationFontSize(
        InlineShape shape)
    {
        using var metric = WordOperationMetrics.Measure("FormatConversion.MathTypeInline.ReadPresentationFont");
        Range? shapeRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? result = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            shapeRange = shape.Range;
            fields = shapeRange.Fields;
            // The object character belongs to the EMBED field's result, not
            // its instruction. Reading the whole field's font is ambiguous when
            // those runs differ; scanning every instruction character is wasteful.
            if (fields.Count == 0)
                result = shapeRange.Duplicate;
            else
            {
                field = fields[1];
                result = field.Result;
            }
            var text = result.Text;
            if (text != "\u0001" && text != "\uFFFC") return null;
            font = result.Font;
            return TryNormalizeDefinedWordFontSize(font.Size, out var size) ? size : null;
        }
        finally
        {
            Release(font);
            Release(result);
            Release(field);
            Release(fields);
            Release(shapeRange);
        }
    }

    private static void PrepareCoreTargetSession(
        OfficeSessionDocument session,
        WordFormulaFormatConversionPlan plan,
        WordFormulaFormatConversionTarget target)
    {
        session.Mode = "create";
        session.Host = "word";
        session.FormulaId =
            target.SourceFormulaId;
        session.SourceDocumentId =
            plan.DocumentId;
        session.SourceObjectId =
            target.SourceObjectId;
        session.DisplayMode =
            target.DisplayMode;
        session.ObjectMode =
            plan.TargetMode;
        session.Numbered =
            target.Numbered;
        session.MathTypeNumberPosition =
            target.MathTypeNumberPosition;
        session.FontSizePt =
            target.FontSizePt;
        session.OriginalMetadata =
            CloneCoreMetadata(
                target.Metadata);
    }

    private static FormulaMetadata CloneCoreMetadata(
        FormulaMetadata metadata) =>
        FormulaMetadataCodec.DeserializeJson(
            FormulaMetadataCodec.SerializeJson(
                metadata))
        ?? throw new InvalidDataException(
            "Unable to clone formula metadata for conversion.");

    private static void ValidateCoreConversionPair(
        string sourceMode,
        string targetMode)
    {
        _ = ObjectModeToHostKind(sourceMode);
        _ = ObjectModeToHostKind(targetMode);
        if (string.Equals(
                sourceMode,
                targetMode,
                StringComparison.Ordinal))
            throw new InvalidOperationException(
                "The source and target formula formats are identical.");
    }
}
