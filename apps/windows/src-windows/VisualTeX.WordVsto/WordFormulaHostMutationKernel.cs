using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;
using WordApplication = Microsoft.Office.Interop.Word.Application;

namespace VisualTeX.WordVsto;

/// <summary>
/// Canonical structural mutation entry point for every OMML/VisualTeX host.
///
/// Callers prepare semantic/rendered target data. This kernel owns Word host
/// replacement, identity, numbering-container preservation, local validation,
/// and rollback.
/// </summary>
internal static class WordFormulaHostMutationKernel
{
    internal static WordFormulaHostDescriptor Replace(
        WordApplication application,
        Document document,
        WordFormulaHostDescriptor source,
        WordFormulaHostWriteRequest target,
        Action<WordFormulaHostDescriptor>? finalize = null)
    {
        ValidateTargetRequest(target);

        WordNativeNumberingAdoptionPlan? nativeNumberInsertion = null;
        try
        {
            // Number attachment is an insertion into Word's native SEQ system,
            // even when the mathematical body already existed as unnumbered
            // OMML. Resolve the complete numbering context BEFORE deleting or
            // rewriting the source host. Otherwise an edit-toggle follows a
            // different path from direct creation: field-free predecessors do
            // not participate in SEQ and a post-delete heading lookup can see a
            // different boundary. One adoption plan now owns both operations.
            var sourceAlreadyOwnsNativeNumber =
                source.Kind == WordFormulaHostKind.Omml
                && source.Display
                && (source.Numbering.Numbered
                    || WordOmmlHostWriter.HasFieldFreeWordNativeNumberHost(
                        document,
                        source));
            if (target.Kind == WordFormulaHostKind.Omml
                && target.Numbered
                && string.Equals(
                    target.DisplayMode,
                    "block",
                    StringComparison.OrdinalIgnoreCase)
                && target.PreparedOmml is null
                && !sourceAlreadyOwnsNativeNumber)
            {
                var semanticOmml =
                    WordOmmlConverter.TransformMathMlToOmml(
                        target.MathMl!);
                nativeNumberInsertion =
                    WordNativeNumberingAdoptionPlan.Prepare(
                        document,
                        source.Range.Start,
                        WordEquationNumberFormatAuthority.Resolve(
                            document));
                target.PreparedOmml =
                    nativeNumberInsertion.PrepareInsertionOmml(
                        semanticOmml);
            }

            // Existing native edits need the same read-before-write contract as
            // native creation. Their numbered-host XML cannot be exported after
            // opening Undo, even though the source range was already resolved.
            // A number toggle uses the same exact-native replacement transport;
            // only its complete number topology is supplied by the adoption plan
            // above instead of preserving the source topology.
            using var nativeEdit =
                nativeNumberInsertion is not null
                && source.Kind == WordFormulaHostKind.Omml
                    ? WordNativeOmmlEdit.PrepareNumberAttachment(
                        document,
                        source,
                        target,
                        target.PreparedOmml!)
                    : WordNativeOmmlEdit.Supports(source, target)
                        ? WordNativeOmmlEdit.Prepare(document, source, target)
                        : null;

            Action? verifyRollback =
                nativeEdit is null
                && nativeNumberInsertion is null
                    ? null
                    : () =>
                    {
                        nativeEdit?.VerifyRollback();
                        nativeNumberInsertion?.VerifyRollback();
                    };

            return WordFormulaMutationTransaction.Execute(
                application,
                document,
                "VisualTeX Replace Formula Host",
                () =>
                {
                    var result = nativeEdit is not null
                        ? nativeEdit.Apply()
                        : ReplaceInsideTransaction(
                            application,
                            document,
                            source,
                            target,
                            sourceAlreadyResolved: true,
                            allowDirectCanonicalOmmlRemoval: false,
                            deferTargetNumberFinalization:
                                nativeNumberInsertion is not null);

                    if (nativeNumberInsertion is not null)
                    {
                        Range? exact = null;
                        OMaths? maths = null;
                        OMath? native = null;
                        try
                        {
                            // Field adoption can change the serialized width of
                            // every earlier/current number. Keep the inserted
                            // native OMath as the physical identity, not the
                            // pre-refresh numeric Range endpoints.
                            exact =
                                WordFormulaHostSemanticReader.CreateRange(
                                    document,
                                    result.Range);
                            maths = exact.OMaths;
                            if (maths.Count != 1)
                                throw new InvalidDataException(
                                    "The edited numbered OMML lost its exact native owner before field adoption.");
                            native = maths[1];

                            nativeNumberInsertion.Apply(result);

                            Release(exact);
                            exact = native.Range.Duplicate;
                            result.Range =
                                new WordFormulaRangeAddress
                                {
                                    StoryType = exact.StoryType,
                                    Start = exact.Start,
                                    End = exact.End,
                                };
                            result =
                                WordNativeOmmlNumbering.RefreshCanonicalNativeFields(
                                    document,
                                    result);
                        }
                        finally
                        {
                            Release(native);
                            Release(maths);
                            Release(exact);
                        }
                    }

                    finalize?.Invoke(result);
                    return result;
                },
                verifyRollback: verifyRollback);
        }
        finally
        {
            nativeNumberInsertion?.Dispose();
        }
    }

    internal static WordFormulaHostDescriptor ReplaceInActiveTransaction(
        WordApplication application,
        Document document,
        WordFormulaHostDescriptor source,
        WordFormulaHostWriteRequest target,
        bool sourceAlreadyResolved = false,
        bool allowDirectCanonicalOmmlRemoval = false,
        bool deferTargetNumberFinalization = false,
        WordNativeConversionSource? preparedSource = null)
    {
        using var operationMetric =
            WordOperationMetrics.Measure(
                "FormatConversion.Replace."
                + source.Kind
                + "To"
                + target.Kind);
        ValidateTargetRequest(target);
        RequireActiveTransaction(application);
        return ReplaceInsideTransaction(
            application,
            document,
            source,
            target,
            sourceAlreadyResolved,
            allowDirectCanonicalOmmlRemoval,
            deferTargetNumberFinalization,
            preparedSource);
    }

    internal static WordFormulaHostDescriptor AdoptUnownedOmmlForMutation(
        Document document,
        WordFormulaHostDescriptor host,
        string? formulaId)
    {
        if (host.Kind != WordFormulaHostKind.Omml)
            return host;
        if (!Guid.TryParse(formulaId, out var parsed))
            throw new InvalidDataException(
                "OMML mutation requires one operation-scoped UUID FormulaId.");

        // Native OMML has no owned/unowned state. The ID exists only so the
        // current editor/mutation request can correlate its response; nothing is
        // written into the Word document.
        host.FormulaId = parsed.ToString("D");
        host.Numbering =
            WordFormulaNumberingResolver.ResolveLocal(
                document,
                host);
        return host;
    }

    internal static WordFormulaHostDescriptor CreateNativeOmml(
        WordApplication application, Document document, string? sourceObjectId,
        WordFormulaHostWriteRequest request, Action<WordFormulaHostDescriptor>? finalize = null)
    {
        ValidateTargetRequest(request);
        if (request.Kind != WordFormulaHostKind.Omml)
            throw new ArgumentException("The native create operation only owns OMML.");
        using var location = WordFormulaOperationLocator.PrepareCreateInsertion(
            document, sourceObjectId, request.DisplayMode);
        Range? evidence = null;
        Range? preparedTarget = null;
        MathTypeLocalEditScope? recovery = null;
        WordNativeNumberingAdoptionPlan? numbers = null;
        try
        {
            evidence = location.EvidenceRange;
            recovery = MathTypeLocalEditScope.Capture(document, evidence);
            location.CaptureBoundaryBookmarks(document, recovery.ParagraphXml);
            var omml = WordOmmlConverter.TransformMathMlToOmml(request.MathMl!);
            if (string.Equals(request.DisplayMode, "inline", StringComparison.OrdinalIgnoreCase))
            {
                // Word can absorb a caret at an existing native math boundary.
                // Prepare the complete neighbouring mathematical bodies BEFORE
                // Undo, then replace that proven span instead of risking loss of
                // the old final character during a collapsed InsertXML.
                var merged = WordOmmlHostWriter.PrepareAdjacentInlineNativeMerge(document, evidence, omml);
                preparedTarget = merged.Target;
                omml = merged.Omml;
            }
            if (request.Numbered)
            {
                var numberFormatId =
                    WordEquationNumberFormatAuthority.Resolve(
                        document);
                numbers =
                    WordNativeNumberingAdoptionPlan.Prepare(
                        document,
                        location.Position,
                        numberFormatId);
                omml = numbers.PrepareInsertionOmml(omml);
            }
            // Everything above is read-only preflight, including the strict
            // recovery evidence. A display paragraph, equation and field updates
            // below are one owned record; no WordOpenXML getter runs in it.
            return WordFormulaMutationTransaction.Execute(application, document,
                "VisualTeX Insert Formula Host", () =>
                {
                    Range? insertion = null;
                    Range? exact = null;
                    OMaths? maths = null;
                    OMath? native = null;
                    try
                    {
                        // Detach the original seam before either the paragraph
                        // split or equation import can expand that empty bookmark.
                        location.BindBoundaryBookmarks(document, evidence, detach: true);
                        insertion = preparedTarget?.Duplicate ?? location.Materialize();
                        var written = WordOmmlHostWriter.InsertPreparedNative(document, insertion, request, omml).Host;
                        written = WordFormulaMutationValidator.ValidateInsertedHost(document, request, written);
                        exact = WordFormulaHostSemanticReader.CreateRange(document, written.Range);
                        location.BindBoundaryBookmarks(document, exact, detach: false);
                        if (numbers is null)
                        {
                            finalize?.Invoke(written);
                            return written;
                        }
                        maths = exact.OMaths;
                        native = maths[1];
                        numbers.Apply(written);
                        Release(exact); exact = native.Range;
                        written.Range = new WordFormulaRangeAddress
                        { StoryType = exact.StoryType, Start = exact.Start, End = exact.End };
                        written = WordNativeOmmlNumbering.RefreshCanonicalNativeFields(document, written);
                        finalize?.Invoke(written);
                        return written;
                    }
                    finally { Release(native); Release(maths); Release(exact); Release(insertion); }
                }, verifyRollback: () =>
                {
                    recovery.VerifyRollback();
                    numbers?.VerifyRollback();
                });
        }
        finally { numbers?.Dispose(); recovery?.Dispose(); Release(preparedTarget); Release(evidence); }
    }

    internal static WordFormulaHostDescriptor Insert(
        WordApplication application,
        Document document,
        Range insertion,
        WordFormulaHostWriteRequest target)
    {
        ValidateTargetRequest(target);

        return WordFormulaMutationTransaction.Execute(
            application,
            document,
            "VisualTeX Insert Formula Host",
            () => InsertInsideTransaction(
                application,
                document,
                insertion,
                target));
    }

    internal static WordFormulaHostDescriptor InsertInActiveTransaction(
        WordApplication application,
        Document document,
        Range insertion,
        WordFormulaHostWriteRequest target)
    {
        ValidateTargetRequest(target);
        RequireActiveTransaction(application);
        return InsertInsideTransaction(
            application,
            document,
            insertion,
            target);
    }

    private static WordFormulaHostDescriptor InsertInsideTransaction(
        WordApplication application,
        Document document,
        Range insertion,
        WordFormulaHostWriteRequest target)
    {
        PrepareTargetIdentity(null, target);
        var write = InsertHost(
            application,
            document,
            insertion,
            target);

        // Low-level writers already validate the physical host they just
        // materialized. VisualTeX additionally reads its embedded OLE payload
        // back and verifies FormulaId/LaTeX/display mode. For a batch mutation,
        // repeating the generic resolver/semantic validation for every formula
        // creates an O(N^2)-like Word COM/layout cost as the document fills with
        // OLE hosts. Interactive create/edit keeps the independent validator.
        var validated =
            target.BatchMutation
                ? write.Host
                : WordFormulaMutationValidator.ValidateInsertedHost(
                    document,
                    target,
                    write.Host);

        if (target.Kind == WordFormulaHostKind.Omml)
        {
            if (target.Numbered)
            {
                // During a batch, create/retain the canonical
                // native numbering host now but defer field finalization until
                // the final document-order refresh.
                return WordFormulaNumberingKernel.AttachCanonical(
                    document,
                    validated,
                    deferFinalization:
                        target.BatchMutation);
            }

            // The OMML writer owns the complete unnumbered physical host.
            return validated;
        }

        if (target.Numbered)
            validated =
                WordFormulaNumberingKernel.AttachCanonical(
                    document,
                    validated,
                    deferFinalization: target.BatchMutation,
                    preparedNumberPlan: target.PreparedVisualTeXNumbering);

        if (target.BatchMutation)
        {
            // Numbered batch targets are finalized/validated once after every
            // batch insert has reached its final document position.
            // Unnumbered VisualTeX targets need no numbering-container walk.
            return validated;
        }

        WordFormulaNumberingKernel.ValidatePreservedContainer(
            document,
            validated,
            target.Numbered);

        validated =
            WordFormulaMutationValidator.ValidateInsertedHost(
                document,
                target,
                validated);
        return validated;
    }

    internal static T ReplaceSourceWithExternalTargetInActiveTransaction<T>(
        WordApplication application,
        Document document,
        WordFormulaHostDescriptor source,
        Func<Range, T> externalWriter,
        bool sourceAlreadyResolved = false,
        bool allowDirectCanonicalOmmlRemoval = false,
        WordNativeConversionSource? preparedSource = null)
    {
        using var operationMetric =
            WordOperationMetrics.Measure(
                "FormatConversion.ReplaceExternalFrom."
                + source.Kind);
        if (externalWriter is null)
            throw new ArgumentNullException(nameof(externalWriter));
        if (source.Kind == WordFormulaHostKind.MathType)
            throw new NotSupportedException(
                "MathType-to-MathType replacement is not owned by the OMML/VisualTeX core.");

        RequireActiveTransaction(application);
        var fresh =
            sourceAlreadyResolved
                ? source
                : ReResolveExactSource(
                    application,
                    document,
                    source);
        var numbering =
            sourceAlreadyResolved
                ? fresh.Numbering
                : WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    fresh);
        if (numbering.ContainerKind ==
            WordFormulaNumberingContainerKind.Legacy)
            throw new InvalidDataException(
                "The source formula uses a retired numbering topology and must be canonicalized before cross-format replacement.");

        var directCanonicalOmmlRemoval =
            allowDirectCanonicalOmmlRemoval
            && fresh.Kind == WordFormulaHostKind.Omml
            && fresh.Display
            && numbering.ContainerKind ==
                WordFormulaNumberingContainerKind.CanonicalNativeOmml;

        if (numbering.Numbered
            && !directCanonicalOmmlRemoval)
        {
            fresh =
                WordFormulaNumberingKernel.DetachCanonical(
                    document,
                    fresh,
                    numbering,
                    sourceWillBeRemovedImmediately: true,
                    preparedBody: preparedSource?.VisualTeXBodyNumbering);
            numbering =
                fresh.Numbering;
            if (numbering.Numbered)
                throw new InvalidDataException(
                    "The source numbering container survived canonical detach.");
        }

        var normalizeDisplayOmmlBoundary =
            fresh.Kind == WordFormulaHostKind.Omml
            && fresh.Display;
        var normalizeDisplayVisualTeXBoundary =
            fresh.Kind == WordFormulaHostKind.VisualTeX
            && fresh.Display
            && !fresh.WithinTable;
        var rightBoundaryBeforeDelete =
            normalizeDisplayOmmlBoundary
                ? null
                : CaptureRightBoundaryText(
                    document,
                    fresh.Range,
                    16);

        Range? insertion = null;
        try
        {
            insertion =
                normalizeDisplayOmmlBoundary
                    ? WordOmmlHostWriter.DeleteExactToExternalDisplayBoundary(
                        document,
                        fresh)
                    : normalizeDisplayVisualTeXBoundary
                        ? WordVisualTeXHostWriter.DeleteExactToExternalDisplayBoundary(
                            document,
                            fresh,
                            preparedSource)
                        : DeleteHostExact(
                            document,
                            fresh,
                            preparedSource);

            if (!normalizeDisplayOmmlBoundary)
            {
                var rightBoundaryAfterDelete =
                    CaptureTextAt(
                        document,
                        insertion.Start,
                        16);
                if (!string.Equals(
                        rightBoundaryBeforeDelete,
                        rightBoundaryAfterDelete,
                        StringComparison.Ordinal))
                    throw new InvalidDataException(
                        "Deleting the source host changed adjacent user text before the external target writer ran. "
                        + $"before=[{(rightBoundaryBeforeDelete ?? string.Empty).Replace("\r", "\\r").Replace("\n", "\\n")}] "
                        + $"after=[{rightBoundaryAfterDelete.Replace("\r", "\\r").Replace("\n", "\\n")}] "
                        + $"source={fresh.Range.Start}:{fresh.Range.End} insertion={insertion.Start}");
            }

            return externalWriter(insertion);
        }
        finally { Release(insertion); }
    }

    internal static T ReplaceSourceWithExternalTarget<T>(
        WordApplication application,
        Document document,
        WordFormulaHostDescriptor source,
        string undoName,
        Func<Range, T> externalWriter)
    {
        if (externalWriter is null)
            throw new ArgumentNullException(nameof(externalWriter));
        if (source.Kind == WordFormulaHostKind.MathType)
            throw new NotSupportedException(
                "MathType-to-MathType replacement is not owned by the OMML/VisualTeX core.");

        return WordFormulaMutationTransaction.Execute(
            application,
            document,
            undoName,
            () =>
            {
                var fresh = ReResolveExactSource(
                    application,
                    document,
                    source);
                var numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        fresh);
                if (numbering.ContainerKind ==
                    WordFormulaNumberingContainerKind.Legacy)
                    throw new InvalidDataException(
                        "The source formula uses a retired numbering topology and must be canonicalized before cross-format replacement.");

                if (numbering.Numbered)
                {
                    fresh =
                        WordFormulaNumberingKernel.DetachCanonical(
                            document,
                            fresh);
                    numbering =
                        WordFormulaNumberingResolver.ResolveLocal(
                            document,
                            fresh);
                    if (numbering.Numbered)
                        throw new InvalidDataException(
                            "The source numbering container survived canonical detach.");
                }

                Range? insertion = null;
                try
                {
                    insertion =
                        fresh.Kind == WordFormulaHostKind.Omml
                        && fresh.Display
                            ? WordOmmlHostWriter.DeleteExactToExternalDisplayBoundary(
                                document,
                                fresh)
                            : fresh.Kind == WordFormulaHostKind.VisualTeX
                              && fresh.Display
                              && !fresh.WithinTable
                                ? WordVisualTeXHostWriter.DeleteExactToExternalDisplayBoundary(
                                    document,
                                    fresh)
                                : DeleteHostExact(
                                    document,
                                    fresh);
                    return externalWriter(insertion);
                }
                finally { Release(insertion); }
            });
    }

    internal static void Delete(
        WordApplication application,
        Document document,
        WordFormulaHostDescriptor source)
    {
        WordFormulaMutationTransaction.Execute(
            application,
            document,
            "VisualTeX Delete Formula Host",
            () =>
            {
                var fresh = ReResolveExactSource(
                    application,
                    document,
                    source);
                var numbering =
                    fresh.Kind == WordFormulaHostKind.MathType
                        ? fresh.Numbering
                        : WordFormulaNumberingResolver.ResolveLocal(
                            document,
                            fresh);

                if (fresh.Kind != WordFormulaHostKind.MathType
                    && numbering.ContainerKind ==
                        WordFormulaNumberingContainerKind.Legacy)
                    throw new InvalidDataException(
                        "The formula uses a retired numbering topology. Canonicalize it before deletion.");

                if (fresh.Kind != WordFormulaHostKind.MathType
                    && numbering.Numbered)
                {
                    // Detach first so table/cell ownership is removed by one
                    // numbering operation; the ordinary host can then be deleted
                    // by the same exact writer used everywhere else.
                    fresh =
                        WordFormulaNumberingKernel.DetachCanonical(
                            document,
                            fresh);
                }

                Range? anchor = null;
                try
                {
                    anchor = DeleteHostExact(
                        document,
                        fresh);
                }
                finally { Release(anchor); }

                return true;
            });
    }

    private static WordFormulaHostDescriptor ReplaceInsideTransaction(
        WordApplication application,
        Document document,
        WordFormulaHostDescriptor source,
        WordFormulaHostWriteRequest target,
        bool sourceAlreadyResolved,
        bool allowDirectCanonicalOmmlRemoval = false,
        bool deferTargetNumberFinalization = false,
        WordNativeConversionSource? preparedSource = null)
    {
        using var operationMetric =
            WordOperationMetrics.Measure(
                "FormatConversion.ReplaceInsideTransaction."
                + source.Kind
                + "->"
                + target.Kind);
        var freshSource =
            sourceAlreadyResolved
                ? source
                : ReResolveExactSource(
                    application,
                    document,
                    source);
        var sourceNumbering =
            freshSource.Numbering;

        // A real VisualTeX OLE can still carry authoritative Numbered=true after
        // Word bookmark gravity has moved/damaged its visible number scaffold.
        // Format conversion preserves that semantic state. Before detaching or
        // replacing such a host, recover the existing numbering through the one
        // canonical ReconcileFormula path rather than treating it as unnumbered
        // or inventing conversion-specific cleanup rules.
        if (freshSource.Kind ==
                WordFormulaHostKind.VisualTeX
            && freshSource.Display
            && target.Numbered
            && !sourceNumbering.Numbered
            && freshSource.Metadata?.Numbered == true)
        {
            freshSource =
                WordFormulaNumberingKernel
                    .RecoverVisualTeXNumberingForConversion(
                        document,
                        freshSource);
            sourceNumbering =
                freshSource.Numbering;
        }

        if (freshSource.Kind != WordFormulaHostKind.MathType
            && sourceNumbering.ContainerKind ==
                WordFormulaNumberingContainerKind.Legacy)
            throw new InvalidDataException(
                "The source formula uses a retired numbering topology. Canonicalize it before host replacement.");

        PrepareTargetIdentity(
            freshSource,
            target);

        // Inline OMML -> inline OMML editing does not need to delete the native
        // OMath and rebuild its entire containing paragraph. The captured exact
        // host is already authoritative and inline OMaths cannot own an equation
        // number, so replace that exact range directly. This keeps surrounding
        // user text untouched and avoids the marker/paragraph-InsertXML path used
        // when a host family or paragraph topology actually changes.
        if (freshSource.Kind == WordFormulaHostKind.Omml
            && !freshSource.Display
            && target.Kind == WordFormulaHostKind.Omml
            && string.Equals(
                target.DisplayMode,
                "inline",
                StringComparison.OrdinalIgnoreCase)
            && !target.Numbered)
        {
            return WordOmmlHostWriter
                .ReplaceExactUnnumberedOmmlHost(
                    document,
                    freshSource,
                    target)
                .Host;
        }

        // Word-native numbered OMML owns its number suffix inside the same
        // OMath host. Preserve that host when editing OMML -> OMML so existing
        // Word _Ref bookmarks continue to target the exact number range. The
        // same body-only replacement also remains mandatory for third-party
        // field-free '#(2)' hosts, whose numbering is entirely Word/user-owned
        // and deliberately exposed to the editor as Numbered=false.
        var preserveFieldFreeWordNumber =
            !target.Numbered
            && WordOmmlHostWriter.HasFieldFreeWordNativeNumberHost(
                document,
                freshSource);
        var preserveManagedWordNumber =
            target.Numbered
            && sourceNumbering.ContainerKind ==
                WordFormulaNumberingContainerKind.CanonicalNativeOmml;
        if (freshSource.Kind == WordFormulaHostKind.Omml
            && freshSource.Display
            && target.Kind == WordFormulaHostKind.Omml
            && string.Equals(
                target.DisplayMode,
                "block",
                StringComparison.OrdinalIgnoreCase)
            && (preserveFieldFreeWordNumber
                || preserveManagedWordNumber))
        {
            var preserved =
                WordOmmlHostWriter.ReplacePreservingWordNativeNumberHost(
                    document,
                    freshSource,
                    target);
            return preserved.Host;
        }

        var sourceUsesSelfContainedVisualTeXNumbering =
            target.Kind == WordFormulaHostKind.VisualTeX
            && freshSource.Kind == WordFormulaHostKind.VisualTeX
            && sourceNumbering.ContainerKind ==
                WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
            && WordVisualTeXParagraphNumbering.IsSelfContainedHost(
                document,
                freshSource);

        var directCanonicalOmmlRemoval =
            allowDirectCanonicalOmmlRemoval
            && freshSource.Kind == WordFormulaHostKind.Omml
            && target.Kind != WordFormulaHostKind.Omml
            && freshSource.Display
            && sourceNumbering.ContainerKind ==
                WordFormulaNumberingContainerKind.CanonicalNativeOmml;

        var preserveCanonicalContainer =
            target.Numbered
            // Numbered OMML always owns Word's native #(SEQ) OMath host. Never
            // preserve an external VisualTeX table/cell scaffold when the target
            // family is OMML; detach it first, then attach native numbering.
            && target.Kind != WordFormulaHostKind.Omml
            && freshSource.Display
            && string.Equals(
                target.DisplayMode,
                "block",
                StringComparison.OrdinalIgnoreCase)
            && sourceNumbering.ContainerKind is
                (WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
                 or WordFormulaNumberingContainerKind.CanonicalBodyTable
                 or WordFormulaNumberingContainerKind.CanonicalUserTableCell);

        // If the target is becoming unnumbered/inline, dismantle the source
        // numbering container before deleting the host. This turns the remainder
        // of the operation into the exact same ordinary host replacement used for
        // every other formula.
        if (freshSource.Kind != WordFormulaHostKind.MathType
            && sourceNumbering.Numbered
            && !preserveCanonicalContainer
            && !directCanonicalOmmlRemoval)
        {
            freshSource =
                WordFormulaNumberingKernel.DetachCanonical(
                    document,
                    freshSource,
                    sourceNumbering,
                    sourceWillBeRemovedImmediately:
                        target.Kind != freshSource.Kind,
                    preparedBody: preparedSource?.VisualTeXBodyNumbering);
            sourceNumbering =
                freshSource.Numbering;
            if (sourceNumbering.Numbered)
                throw new InvalidDataException(
                    "The source numbering container survived canonical detach.");
        }

        Range? insertion = null;
        try
        {
            if (ShouldReplaceExternalHostInPlaceWithOmml(
                    freshSource,
                    sourceNumbering,
                    target))
            {
                RemoveExternalHostIdentityForOmmlReplacement(
                    document,
                    freshSource);
                insertion =
                    WordFormulaHostSemanticReader.CreateRange(
                        document,
                        freshSource.Range);
            }
            else
            {
                insertion =
                    directCanonicalOmmlRemoval
                        ? WordOmmlHostWriter.DeleteExactToPlainTextBoundary(
                            document,
                            freshSource)
                        : freshSource.Kind == WordFormulaHostKind.Omml
                          && target.Kind != WordFormulaHostKind.Omml
                            ? WordOmmlHostWriter.DeleteExactToPlainTextBoundary(
                                document,
                                freshSource)
                        : freshSource.Kind == WordFormulaHostKind.VisualTeX
                          && freshSource.Display
                          && target.Kind == WordFormulaHostKind.Omml
                            ? WordVisualTeXHostWriter.DeleteExactToExternalDisplayBoundary(
                                document,
                                freshSource,
                                preparedSource)
                            : DeleteHostExact(
                                document,
                                freshSource,
                                preparedSource);
            }

            // A prepared native target at the normalized plain boundary needs
            // only the native insertion writer. The generic marker/paragraph
            // route exports live WordOpenXML and closes the active Undo record.
            var written = target.Kind == WordFormulaHostKind.Omml
                && target.PreparedOmml is not null
                && target.PreparedOleReplacement is null
                    ? WordOmmlHostWriter.InsertPreparedNative(
                        document, insertion, target, target.PreparedOmml)
                    : InsertHost(application, document, insertion, target,
                        removeOmmlBoundaryGuard: freshSource.Kind == WordFormulaHostKind.Omml
                            && target.Kind == WordFormulaHostKind.VisualTeX);
            // The low-level writers already prove the exact physical host they
            // just materialized. VisualTeX additionally reads its embedded OLE
            // payload back and validates FormulaId/LaTeX/display mode. During a
            // format-conversion batch, immediately resolving the same host again
            // through WordFormulaMutationValidator is duplicate COM/OLE work.
            // Interactive create/edit keeps the independent validator.
            var validated =
                target.BatchMutation
                    ? written.Host
                    : WordFormulaMutationValidator.ValidateInsertedHost(
                        document,
                        target,
                        written.Host);

            if (target.Kind == WordFormulaHostKind.Omml)
            {
                if (target.Numbered)
                {
                    return WordFormulaNumberingKernel.AttachCanonical(
                        document,
                        validated,
                        deferTargetNumberFinalization);
                }

                // The OMML writer has already established one physical OMath.
                // No semantic/fingerprint comparison participates in this path.
                return validated;
            }

            if (target.BatchMutation
                && !target.Numbered
                && target.Kind == WordFormulaHostKind.VisualTeX)
            {
                // The VisualTeX writer has already read the embedded OLE payload
                // back and the target has no numbering container to finalize.
                return validated;
            }

            if (preserveCanonicalContainer)
            {
                // The number field and VTEqNum target are intentionally untouched.
                // Replacing the OLE itself can nevertheless make Word reapply the
                // ordinary centered block paragraph format. For the canonical body
                // tab host, restore only paragraph geometry before validating the
                // preserved SEQ/REF container.
                if (sourceNumbering.ContainerKind ==
                        WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
                    && target.Kind == WordFormulaHostKind.VisualTeX
                    && !sourceUsesSelfContainedVisualTeXNumbering)
                {
                    Range? preservedRange = null;
                    try
                    {
                        preservedRange =
                            WordFormulaHostSemanticReader.CreateRange(
                                document,
                                validated.Range);
                        WordEquationNumbering
                            .NormalizePreservedVisualTeXNumberedTabParagraph(
                                document,
                                preservedRange,
                                target.HeightPoints,
                                target.Metadata
                                ?? throw new InvalidDataException(
                                    "A preserved numbered VisualTeX target has no metadata."));
                    }
                    finally { Release(preservedRange); }
                }

                WordFormulaNumberingKernel.ValidatePreservedContainer(
                    document,
                    validated,
                    shouldBeNumbered: true);
            }
            else if (target.Numbered)
            {
                validated =
                    WordFormulaNumberingKernel.AttachCanonical(
                        document,
                        validated,
                        deferTargetNumberFinalization,
                        target.PreparedVisualTeXNumbering);
            }

            if (deferTargetNumberFinalization
                && target.Numbered
                && !preserveCanonicalContainer)
            {
                // Batch conversion validates every numbering container in one
                // final document-order refresh. The formula host itself has
                // already passed the local semantic validation above, so a
                // second identical host read here only duplicates COM/OLE work.
                return validated;
            }

            WordFormulaNumberingKernel.ValidatePreservedContainer(
                document,
                validated,
                target.Numbered);

            // Number attachment can move the host into a canonical container.
            // Re-read the actual host after that move and validate its semantics
            // again. No global count participates.
            validated =
                WordFormulaMutationValidator.ValidateInsertedHost(
                    document,
                    target,
                    validated);

            return validated;
        }
        finally
        {
            Release(insertion);
        }
    }

    private static WordFormulaHostDescriptor ReResolveExactSource(
        WordApplication application,
        Document document,
        WordFormulaHostDescriptor source)
    {
        Range? sourceRange = null;
        try
        {
            sourceRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    source.Range);
            var resolved =
                source.Kind == WordFormulaHostKind.MathType
                    ? WordMathTypeHostAdapter.ResolveLocal(
                        application,
                        document,
                        sourceRange)
                    : WordFormulaHostLookup.ResolveLocal(
                        document,
                        sourceRange,
                        source.Kind);
            if (resolved is null)
                throw new InvalidDataException(
                    "The source formula host no longer exists at its captured range.");

            if (resolved.Range.StoryType != source.Range.StoryType
                || resolved.Range.Start != source.Range.Start
                || resolved.Range.End != source.Range.End)
                throw new InvalidDataException(
                    "The source formula host moved before mutation.");

            if (source.Kind == WordFormulaHostKind.VisualTeX
                && !string.IsNullOrWhiteSpace(source.FormulaId)
                && !string.IsNullOrWhiteSpace(resolved.FormulaId)
                && !string.Equals(
                    source.FormulaId,
                    resolved.FormulaId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "The source formula identity changed before mutation.");

            if (source.Kind == WordFormulaHostKind.VisualTeX
                && resolved.FormulaId is null
                && Guid.TryParse(source.FormulaId, out _))
                resolved.FormulaId = source.FormulaId;

            if (resolved.Kind != WordFormulaHostKind.MathType)
                resolved.Numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        resolved);
            return resolved;
        }
        finally { Release(sourceRange); }
    }

    private static void PrepareTargetIdentity(
        WordFormulaHostDescriptor? source,
        WordFormulaHostWriteRequest target)
    {
        if (target.Kind == WordFormulaHostKind.Omml)
        {
            // Word OMath carries no VisualTeX durable identity. The editor may
            // keep its own operation-scoped FormulaId, but the host request must
            // not turn it into document state.
            target.FormulaId = null;
            return;
        }

        var formulaId =
            source?.Kind != WordFormulaHostKind.MathType
            && Guid.TryParse(source?.FormulaId, out var sourceId)
                ? sourceId.ToString("D")
                : Guid.TryParse(target.FormulaId, out var targetId)
                    ? targetId.ToString("D")
                    : Guid.TryParse(
                        target.Metadata?.FormulaId,
                        out var metadataId)
                        ? metadataId.ToString("D")
                        : Guid.NewGuid().ToString("D");

        target.FormulaId = formulaId;
        if (target.Metadata is not null)
        {
            target.Metadata.FormulaId = formulaId;
            target.Metadata.DisplayMode = target.DisplayMode;
            target.Metadata.Numbered = target.Numbered;
            target.Metadata.Validate();
        }
    }

    private static bool ShouldReplaceExternalHostInPlaceWithOmml(
        WordFormulaHostDescriptor source,
        WordFormulaNumberingDescriptor sourceNumbering,
        WordFormulaHostWriteRequest target)
    {
        if (target.Kind != WordFormulaHostKind.Omml)
            return false;

        if (source.Kind == WordFormulaHostKind.VisualTeX)
        {
            // Inline OLE can be replaced in-place without changing the user's
            // paragraph text boundary. A display OLE cannot: its canonical host
            // owns a leading center-tab scaffold, and asking Word to overwrite
            // the live EMBED range can materialize that boundary as TAB + BR
            // before the new oMathPara. Display conversion therefore goes
            // through WordVisualTeXHostWriter.DeleteExactToExternalDisplayBoundary
            // and writes OMML only after the owned display scaffold is gone.
            return !source.Display;
        }

        // A prepared MathType plan owns its exact sibling number scaffold as
        // well as the OLE atom. Keep one physical substitution instead of
        // clearing the whole historical paragraph. Unprepared legacy callers
        // still require their established numbering detach path.
        return source.Kind == WordFormulaHostKind.MathType
            && (!sourceNumbering.Numbered || target.PreparedOleReplacement is not null);
    }

    private static void RemoveExternalHostIdentityForOmmlReplacement(
        Document document,
        WordFormulaHostDescriptor source)
    {
        switch (source.Kind)
        {
            case WordFormulaHostKind.VisualTeX:
                WordFormulaIdentityStore.RemoveVisualTeX(
                    document,
                    source.FormulaId);
                break;
            case WordFormulaHostKind.MathType:
                WordFormulaIdentityStore.RemoveMathType(
                    document,
                    source.FormulaId);
                break;
            default:
                throw new ArgumentException(
                    "In-place OMML replacement requires an external OLE source.");
        }
    }

    private static WordFormulaHostWriteResult InsertHost(
        WordApplication application,
        Document document,
        Range insertion,
        WordFormulaHostWriteRequest target,
        bool removeOmmlBoundaryGuard = false) =>
        target.Kind switch
        {
            WordFormulaHostKind.Omml =>
                WordOmmlHostWriter.Insert(
                    application,
                    document,
                    insertion,
                    target),
            WordFormulaHostKind.VisualTeX =>
                WordVisualTeXHostWriter.Insert(
                    document,
                    insertion,
                    target,
                    removeOmmlBoundaryGuard),
            _ => throw new NotSupportedException(
                $"The rebuilt OMML/VisualTeX core cannot write {target.Kind}."),
        };

    private static Range DeleteHostExact(
        Document document,
        WordFormulaHostDescriptor source,
        WordNativeConversionSource? preparedSource = null) =>
        source.Kind switch
        {
            WordFormulaHostKind.Omml =>
                WordOmmlHostWriter.DeleteExact(
                    document,
                    source),
            WordFormulaHostKind.VisualTeX =>
                WordVisualTeXHostWriter.DeleteExact(
                    document,
                    source,
                    preparedSource),
            WordFormulaHostKind.MathType =>
                WordMathTypeHostAdapter.DeleteExact(
                    document,
                    source,
                    preparedSource?.MathTypeShape),
            _ => throw new NotSupportedException(
                $"The rebuilt OMML/VisualTeX core cannot delete {source.Kind}."),
        };

    private static string CaptureRightBoundaryText(
        Document document,
        WordFormulaRangeAddress source,
        int maximumLength) =>
        CaptureTextAt(
            document,
            source.End,
            maximumLength);

    private static string CaptureTextAt(
        Document document,
        int start,
        int maximumLength)
    {
        Range? content = null;
        Range? probe = null;
        try
        {
            content = document.Content;
            var safeStart = Math.Max(
                content.Start,
                Math.Min(start, content.End));
            var safeEnd = Math.Min(
                content.End,
                safeStart + Math.Max(0, maximumLength));
            if (safeEnd <= safeStart)
                return string.Empty;
            probe = document.Range(
                safeStart,
                safeEnd);
            return probe.Text ?? string.Empty;
        }
        finally
        {
            Release(probe);
            Release(content);
        }
    }

    private static void ValidateTargetRequest(
        WordFormulaHostWriteRequest target)
    {
        if (target.Kind is not
            (WordFormulaHostKind.Omml
             or WordFormulaHostKind.VisualTeX))
            throw new NotSupportedException(
                "This mutation kernel owns OMML and VisualTeX targets only.");

        var inline = string.Equals(
            target.DisplayMode,
            "inline",
            StringComparison.OrdinalIgnoreCase);
        var block = string.Equals(
            target.DisplayMode,
            "block",
            StringComparison.OrdinalIgnoreCase);
        if (!inline && !block)
            throw new InvalidDataException(
                "Formula host display mode must be inline or block.");
        if (inline && target.Numbered)
            throw new InvalidDataException(
                "An inline formula cannot own a display equation number.");

        if (target.Kind == WordFormulaHostKind.Omml
            && string.IsNullOrWhiteSpace(target.MathMl))
            throw new InvalidDataException(
                "OMML target requires MathML.");
        if (target.Kind == WordFormulaHostKind.VisualTeX
            && target.Metadata is null)
            throw new InvalidDataException(
                "VisualTeX target requires authoritative embedded metadata.");
    }

    private static void RequireActiveTransaction(
        WordApplication application)
    {
        _ = application
            ?? throw new ArgumentNullException(nameof(application));

        // Word 2021 can report IsRecordingCustomRecord=false from a newly
        // acquired UndoRecord RCW even while the exact outer core transaction
        // that called us is active. Transaction ownership is therefore a host
        // invariant, not a second COM query. Only WordFormulaMutationTransaction
        // may establish this state.
        if (!WordFormulaMutationTransaction.IsActiveOnCurrentThread)
            throw new InvalidOperationException(
                "The batch mutation primitive requires one active VisualTeX core transaction.");
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
