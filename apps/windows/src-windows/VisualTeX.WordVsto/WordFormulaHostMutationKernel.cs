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
        WordFormulaHostWriteRequest target)
    {
        ValidateTargetRequest(target);

        return WordFormulaMutationTransaction.Execute(
            application,
            document,
            "VisualTeX Replace Formula Host",
            () => ReplaceInsideTransaction(
                application,
                document,
                source,
                target,
                sourceAlreadyResolved: true));
    }

    internal static WordFormulaHostDescriptor ReplaceInActiveTransaction(
        WordApplication application,
        Document document,
        WordFormulaHostDescriptor source,
        WordFormulaHostWriteRequest target)
    {
        ValidateTargetRequest(target);
        RequireActiveTransaction(application);
        return ReplaceInsideTransaction(
            application,
            document,
            source,
            target,
            sourceAlreadyResolved: false);
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
        var validated =
            target.Kind == WordFormulaHostKind.Omml
            && write.SemanticPostconditionValidated
                ? write.Host
                : WordFormulaMutationValidator.ValidateInsertedHost(
                    document,
                    target,
                    write.Host);

        if (target.Kind == WordFormulaHostKind.Omml)
        {
            if (target.Numbered)
            {
                // Pure OMML numbering is already materialized by the writer.
                // AttachCanonical only refreshes Word fields and validates the
                // final native #(SEQ) container.
                return WordFormulaNumberingKernel.AttachCanonical(
                    document,
                    validated);
            }

            // The OMML writer owns the complete unnumbered host and has already
            // validated the exact materialized WordOpenXML above. There is no
            // separate numbering container to rediscover for a fresh insertion.
            return validated;
        }

        if (target.Numbered)
            validated =
                WordFormulaNumberingKernel.AttachCanonical(
                    document,
                    validated);

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
        Func<Range, T> externalWriter)
    {
        if (externalWriter is null)
            throw new ArgumentNullException(nameof(externalWriter));
        if (source.Kind == WordFormulaHostKind.MathType)
            throw new NotSupportedException(
                "MathType-to-MathType replacement is not owned by the OMML/VisualTeX core.");

        RequireActiveTransaction(application);
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
                            fresh)
                        : DeleteHostExact(
                            document,
                            fresh);

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
        bool sourceAlreadyResolved)
    {
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

        var sourceUsesSelfContainedVisualTeXNumbering =
            freshSource.Kind == WordFormulaHostKind.VisualTeX
            && sourceNumbering.ContainerKind ==
                WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
            && WordVisualTeXParagraphNumbering.IsSelfContainedHost(
                document,
                freshSource);

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
            && !preserveCanonicalContainer)
        {
            freshSource =
                WordFormulaNumberingKernel.DetachCanonical(
                    document,
                    freshSource);
            sourceNumbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    freshSource);
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
                    freshSource.Kind == WordFormulaHostKind.Omml
                    && target.Kind != WordFormulaHostKind.Omml
                        ? WordOmmlHostWriter.DeleteExactToPlainTextBoundary(
                            document,
                            freshSource)
                        : freshSource.Kind == WordFormulaHostKind.VisualTeX
                          && freshSource.Display
                          && target.Kind == WordFormulaHostKind.Omml
                            ? WordVisualTeXHostWriter.DeleteExactToExternalDisplayBoundary(
                                document,
                                freshSource)
                            : DeleteHostExact(
                                document,
                                freshSource);
            }

            var written = InsertHost(
                application,
                document,
                insertion,
                target);
            var validated =
                target.Kind == WordFormulaHostKind.Omml
                && written.SemanticPostconditionValidated
                    ? written.Host
                    : WordFormulaMutationValidator.ValidateInsertedHost(
                        document,
                        target,
                        written.Host);

            if (freshSource.Kind == WordFormulaHostKind.Omml
                && target.Kind != WordFormulaHostKind.Omml)
            {
                validated =
                    WordOmmlHostWriter.RemovePlainTextBoundaryGuard(
                        document,
                        validated);
                validated =
                    WordFormulaMutationValidator.ValidateInsertedHost(
                        document,
                        target,
                        validated);
            }

            if (target.Kind == WordFormulaHostKind.Omml)
            {
                if (target.Numbered)
                {
                    return WordFormulaNumberingKernel.AttachCanonical(
                        document,
                        validated);
                }

                // The OMML writer has already validated the exact materialized
                // WordOpenXML. An unnumbered pure OMath has no external
                // container to rediscover.
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
                        validated);
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
                    : WordFormulaHostResolver.ResolveLocal(
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

        // MathType owns a separate numbering implementation. For numbered
        // MathType, keep its established detach/delete path so its native fields
        // are removed before Word-native OMML numbering is attached.
        return source.Kind == WordFormulaHostKind.MathType
            && !sourceNumbering.Numbered;
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
        WordFormulaHostWriteRequest target) =>
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
                    target),
            _ => throw new NotSupportedException(
                $"The rebuilt OMML/VisualTeX core cannot write {target.Kind}."),
        };

    private static Range DeleteHostExact(
        Document document,
        WordFormulaHostDescriptor source) =>
        source.Kind switch
        {
            WordFormulaHostKind.Omml =>
                WordOmmlHostWriter.DeleteExact(
                    document,
                    source),
            WordFormulaHostKind.VisualTeX =>
                WordVisualTeXHostWriter.DeleteExact(
                    document,
                    source),
            WordFormulaHostKind.MathType =>
                WordMathTypeHostAdapter.DeleteExact(
                    document,
                    source),
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
