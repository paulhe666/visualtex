using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    private bool TryCaptureHostCoreCopySnapshot(
        Document document,
        Range selectionRange,
        out WordFormulaCopySnapshot? snapshot,
        out bool touchesHostCore)
    {
        snapshot = null;
        touchesHostCore = false;

        WordFormulaHostDescriptor? direct = null;
        try
        {
            direct = TryResolveSingleHostCoreCopySource(
                document,
                selectionRange);
            if (direct is not null)
            {
                touchesHostCore = true;
                snapshot = CaptureSingleHostCoreCopySnapshot(
                    document,
                    direct);
                return true;
            }
        }
        finally
        {
            // Descriptors contain no live COM objects.
        }

        WordFormulaDocumentIndex index;
        try
        {
            index = WordFormulaHostResolver.CaptureScopeIndex(
                document,
                selectionRange);
        }
        catch
        {
            return false;
        }

        var hosts = index.Omml
            .Concat(index.VisualTeX)
            .Where(host =>
                host.Range.StoryType == selectionRange.StoryType
                && host.Range.Start >= selectionRange.Start
                && host.Range.End <= selectionRange.End)
            .OrderBy(host => host.Range.Start)
            .ToArray();

        touchesHostCore =
            hosts.Length > 0
            || index.Omml.Count > 0
            || index.VisualTeX.Count > 0;
        if (hosts.Length == 0)
            return false;

        if (hosts.Select(host => host.Kind).Distinct().Count() != 1)
            return false;

        if (hosts.Length == 1)
        {
            snapshot = CaptureSingleHostCoreCopySnapshot(
                document,
                hosts[0]);
            return true;
        }

        snapshot = CaptureHostCoreCopyGroupSnapshot(
            document,
            selectionRange,
            hosts);
        return true;
    }

    private static WordFormulaHostDescriptor? TryResolveSingleHostCoreCopySource(
        Document document,
        Range selectionRange)
    {
        WordFormulaHostDescriptor? visualTeX = null;
        WordFormulaHostDescriptor? omml = null;

        try
        {
            visualTeX = WordFormulaHostResolver.ResolveLocal(
                document,
                selectionRange,
                WordFormulaHostKind.VisualTeX);
        }
        catch (InvalidDataException)
        {
            // A selection spanning more than one host is handled by the bounded
            // scope index below. Never fall back to a document-wide search.
        }

        try
        {
            omml = WordFormulaHostResolver.ResolveLocal(
                document,
                selectionRange,
                WordFormulaHostKind.Omml);
        }
        catch (InvalidDataException)
        {
            // Same bounded-group handling as above.
        }

        if (visualTeX is not null && omml is not null)
            return null;
        return visualTeX ?? omml;
    }

    private WordFormulaCopySnapshot CaptureSingleHostCoreCopySnapshot(
        Document document,
        WordFormulaHostDescriptor host)
    {
        host.Numbering = WordFormulaNumberingResolver.ResolveLocal(
            document,
            host);
        if (host.Numbering.ContainerKind ==
            WordFormulaNumberingContainerKind.Legacy)
            throw new InvalidDataException(
                "The copied formula still uses a retired numbering topology.");

        var payload = WordFormulaHostSemanticReader.Read(
            document,
            host);
        var metadata = CreateHostCoreSnapshotMetadata(
            host,
            payload);
        var documentId = DocumentIdentity(document);

        return new WordFormulaCopySnapshot
        {
            ObjectMode = host.Kind == WordFormulaHostKind.VisualTeX
                ? FormulaOleContract.NativeOleMode
                : FormulaOleContract.WordOmmlMode,
            Metadata = metadata,
            SourceDocumentId = documentId,
            SourceStoryType = host.Range.StoryType,
            SourceStart = host.Range.Start,
            SourceEnd = host.Range.End,
            BodyFormatting = TryCaptureHostCoreParagraphFormatting(
                document,
                host.Range),
            TrackingDocumentId = documentId,
            KnownDocumentEnd = document.Content.End,
            VisibleNumber = ReadHostCoreVisibleNumber(
                document,
                host.Numbering),
            CoreHostKind = host.Kind,
            CoreSemanticSignature = HostCoreSemanticSignature(payload),
            CoreMathMl = payload.MathMl,
            CoreLatex = payload.Latex,
            UsesHostCore = true,
        };
    }

    private WordFormulaCopySnapshot CaptureHostCoreCopyGroupSnapshot(
        Document document,
        Range selectionRange,
        IReadOnlyList<WordFormulaHostDescriptor> hosts)
    {
        if (hosts.Count < 2)
            throw new ArgumentException(
                "A host-core copy group requires at least two formulas.",
                nameof(hosts));

        var items = new List<WordFormulaCopyGroupItem>(
            hosts.Count);

        foreach (var source in hosts)
        {
            source.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    source);
            if (source.Numbering.ContainerKind ==
                WordFormulaNumberingContainerKind.Legacy)
                throw new InvalidDataException(
                    "The copied formula group contains a retired numbering topology.");

            var payload = WordFormulaHostSemanticReader.Read(
                document,
                source);
            var metadata = CreateHostCoreSnapshotMetadata(
                source,
                payload);
            items.Add(new WordFormulaCopyGroupItem
            {
                Metadata = metadata,
                SourceStoryType = source.Range.StoryType,
                SourceStart = source.Range.Start,
                SourceEnd = source.Range.End,
                BodyFormatting =
                    TryCaptureHostCoreParagraphFormatting(
                        document,
                        source.Range),
                VisibleNumber = ReadHostCoreVisibleNumber(
                    document,
                    source.Numbering),
                CoreHostKind = source.Kind,
                CoreSemanticSignature =
                    HostCoreSemanticSignature(payload),
            });
        }

        var first = items[0];
        var firstHost = hosts[0];
        var firstPayload =
            WordFormulaHostSemanticReader.Read(
                document,
                firstHost);
        var documentId = DocumentIdentity(document);

        return new WordFormulaCopySnapshot
        {
            ObjectMode =
                firstHost.Kind == WordFormulaHostKind.VisualTeX
                    ? FormulaOleContract.NativeOleMode
                    : FormulaOleContract.WordOmmlMode,
            Metadata = CloneFormulaMetadata(first.Metadata),
            SourceDocumentId = documentId,
            SourceStoryType = selectionRange.StoryType,
            SourceStart = selectionRange.Start,
            SourceEnd = selectionRange.End,
            BodyFormatting = first.BodyFormatting,
            TrackingDocumentId = documentId,
            KnownDocumentEnd = document.Content.End,
            VisibleNumber = first.VisibleNumber,
            GroupItems = items,
            CopiedSelectionStart = selectionRange.Start,
            CopiedSelectionEnd = selectionRange.End,
            CoreHostKind = firstHost.Kind,
            CoreSemanticSignature =
                HostCoreSemanticSignature(firstPayload),
            CoreMathMl = firstPayload.MathMl,
            CoreLatex = firstPayload.Latex,
            UsesHostCore = true,
        };
    }

    internal void ArmHostCorePaste(
        WordFormulaCopySnapshot snapshot)
    {
        if (snapshot is null || !snapshot.UsesHostCore)
            return;

        Document? document = null;
        Selection? selection = null;
        Range? selected = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null)
                return;

            selection = _application.Selection;
            selected = selection.Range;
            snapshot.CorePendingPasteDocumentId =
                DocumentIdentity(document);
            snapshot.CorePendingPasteStoryType =
                selected.StoryType;
            snapshot.CorePendingPasteStart =
                selected.Start;
            snapshot.CorePendingPasteEnd =
                selected.End;
            snapshot.CorePendingPasteArmed = true;
            snapshot.TrackingDocumentId =
                snapshot.CorePendingPasteDocumentId;
            snapshot.KnownDocumentEnd =
                document.Content.End;

            snapshot.CorePendingOmmlMergeTarget = false;
            snapshot.CorePendingOmmlTargetStart = 0;
            snapshot.CorePendingOmmlTargetEnd = 0;
            snapshot.CorePendingOmmlTargetFormulaId = null;
            snapshot.CorePendingOmmlTargetSemanticSignature = null;
            snapshot.CorePendingOmmlTargetLatex = null;

            // Word's native model treats a paste at an inline OMath boundary as
            // insertion into that same mathematical zone. Capture only that local
            // pre-paste host so the post-paste core can recognize the natural
            // merge without inventing a hidden separator or a second fake host.
            if (snapshot.CoreHostKind == WordFormulaHostKind.Omml
                && string.Equals(
                    snapshot.Metadata.DisplayMode,
                    "inline",
                    StringComparison.Ordinal)
                && selected.Start == selected.End)
            {
                try
                {
                    var mergeTarget =
                        WordFormulaHostResolver.ResolveLocal(
                            document,
                            selected,
                            WordFormulaHostKind.Omml);
                    if (mergeTarget is not null
                        && string.Equals(
                            mergeTarget.DisplayMode,
                            "inline",
                            StringComparison.Ordinal))
                    {
                        var targetPayload =
                            WordFormulaHostSemanticReader.Read(
                                document,
                                mergeTarget);
                        snapshot.CorePendingOmmlMergeTarget = true;
                        snapshot.CorePendingOmmlTargetStart =
                            mergeTarget.Range.Start;
                        snapshot.CorePendingOmmlTargetEnd =
                            mergeTarget.Range.End;
                        snapshot.CorePendingOmmlTargetFormulaId =
                            null;
                        snapshot.CorePendingOmmlTargetSemanticSignature =
                            HostCoreSemanticSignature(targetPayload);
                        snapshot.CorePendingOmmlTargetLatex =
                            targetPayload.Latex;
                    }
                }
                catch (InvalidDataException)
                {
                    // A caret touching two independent hosts is ambiguous. Do not
                    // guess which host Word will merge; repair will require an
                    // independently materialized pasted host instead.
                }
            }

            WordDoubleClickHook.TraceMessage(
                $"copy-paste-core-armed kind={snapshot.CoreHostKind} "
                + $"range={selected.Start}:{selected.End} "
                + $"story={selected.StoryType}");
        }
        finally
        {
            Release(selected);
            Release(selection);
            Release(document);
        }
    }

    private bool HasPotentialHostCoreCopiedInsertion(
        WordFormulaCopySnapshot snapshot)
    {
        if (!snapshot.UsesHostCore
            || !snapshot.CorePendingPasteArmed)
            return false;

        Document? document = null;
        Selection? selection = null;
        Range? selected = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null
                || !string.Equals(
                    DocumentIdentity(document),
                    snapshot.CorePendingPasteDocumentId,
                    StringComparison.OrdinalIgnoreCase))
                return false;

            selection = _application.Selection;
            selected = selection.Range;
            if (selected.StoryType !=
                snapshot.CorePendingPasteStoryType)
                return false;

            var currentEnd = Math.Max(
                selected.Start,
                selected.End);
            return currentEnd >
                snapshot.CorePendingPasteStart;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(selected);
            Release(selection);
            Release(document);
        }
    }

    private PastedFormulaRepairResult RepairPastedFormulaWithHostCore(
        WordFormulaCopySnapshot snapshot)
    {
        if (!snapshot.CorePendingPasteArmed
            || snapshot.CoreHostKind is null)
            return PastedFormulaRepairResult.NotApplicable;

        Document? document = null;
        Selection? selection = null;
        Range? selected = null;
        Range? pastedRegion = null;
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
                    DocumentIdentity(document),
                    snapshot.CorePendingPasteDocumentId,
                    StringComparison.OrdinalIgnoreCase))
                return PastedFormulaRepairResult.NotApplicable;

            selection = _application.Selection;
            selected = selection.Range;
            if (selected.StoryType !=
                snapshot.CorePendingPasteStoryType)
                return PastedFormulaRepairResult.NotApplicable;

            var pastedStart =
                snapshot.CorePendingPasteStart;
            var pastedEnd = Math.Max(
                selected.Start,
                selected.End);
            if (pastedEnd <= pastedStart)
                return PastedFormulaRepairResult.NotReady;

            pastedRegion =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    new WordFormulaRangeAddress
                    {
                        StoryType =
                            snapshot.CorePendingPasteStoryType,
                        Start = pastedStart,
                        End = pastedEnd,
                    });

            var index =
                WordFormulaHostResolver.CaptureScopeIndex(
                    document,
                    pastedRegion);
            var candidates =
                index.ForKind(snapshot.CoreHostKind.Value)
                    .Where(host =>
                        host.Range.StoryType ==
                            pastedRegion.StoryType
                        && host.Range.Start >=
                            pastedRegion.Start
                        && host.Range.End <=
                            pastedRegion.End)
                    .OrderBy(host => host.Range.Start)
                    .ToArray();

            if (snapshot.GroupItems.Count > 1)
            {
                var groupResult =
                    RepairPastedHostCoreGroup(
                        document,
                        snapshot,
                        candidates);
                if (groupResult ==
                    PastedFormulaRepairResult.Repaired)
                    CompleteHostCorePaste(
                        document,
                        snapshot);
                return groupResult;
            }

            if (snapshot.CoreHostKind ==
                WordFormulaHostKind.VisualTeX)
            {
                var result =
                    RepairOnePastedVisualTeXHost(
                        document,
                        snapshot,
                        candidates);
                if (result ==
                    PastedFormulaRepairResult.Repaired)
                    CompleteHostCorePaste(
                        document,
                        snapshot);
                return result;
            }

            if (snapshot.CoreHostKind ==
                WordFormulaHostKind.Omml)
            {
                var result =
                    RepairOnePastedOmmlHost(
                        document,
                        snapshot,
                        candidates,
                        pastedRegion);
                if (result ==
                    PastedFormulaRepairResult.Repaired)
                    CompleteHostCorePaste(
                        document,
                        snapshot);
                return result;
            }

            return PastedFormulaRepairResult.NotApplicable;
        }
        finally
        {
            Release(pastedRegion);
            Release(selected);
            Release(selection);
            Release(document);
        }
    }

    private PastedFormulaRepairResult RepairOnePastedVisualTeXHost(
        Document document,
        WordFormulaCopySnapshot snapshot,
        IReadOnlyList<WordFormulaHostDescriptor> candidates)
    {
        var matches =
            candidates
                .Where(candidate =>
                    CandidateVisualTeXMatchesCopy(
                        document,
                        candidate,
                        snapshot.Metadata.FormulaId,
                        snapshot.CoreSemanticSignature))
                .ToArray();

        if (matches.Length == 0)
            return PastedFormulaRepairResult.NotReady;
        if (matches.Length != 1)
            throw new InvalidDataException(
                "The local paste region contains more than one physical VisualTeX copy of the tracked source.");

        RekeyPastedVisualTeXHost(
            document,
            matches[0],
            snapshot.Metadata,
            snapshot.CoreSemanticSignature,
            snapshot.VisibleNumber);
        return PastedFormulaRepairResult.Repaired;
    }

    private PastedFormulaRepairResult RepairOnePastedOmmlHost(
        Document document,
        WordFormulaCopySnapshot snapshot,
        IReadOnlyList<WordFormulaHostDescriptor> candidates,
        Range pastedRegion)
    {
        var matches =
            candidates
                .Where(candidate =>
                    CandidateOmmlMatchesCopy(
                        document,
                        candidate,
                        snapshot.CoreSemanticSignature))
                .ToArray();

        if (matches.Length == 0)
        {
            if (string.Equals(
                    snapshot.Metadata.DisplayMode,
                    "inline",
                    StringComparison.Ordinal)
                && TryRepairMergedPastedOmmlWithHostCore(
                    document,
                    snapshot,
                    pastedRegion))
                return PastedFormulaRepairResult.Repaired;

            return PastedFormulaRepairResult.NotReady;
        }

        if (matches.Length != 1)
            throw new InvalidDataException(
                "The local paste region contains more than one independent OMML equation matching the copied source.");

        var pasted = matches[0];
        pasted.Numbering =
            WordFormulaNumberingResolver.ResolveLocal(
                document,
                pasted);

        if (snapshot.Metadata.Numbered)
        {
            if (!pasted.Display
                || pasted.Numbering.ContainerKind !=
                    WordFormulaNumberingContainerKind.CanonicalNativeOmml)
                throw new InvalidDataException(
                    "Word did not paste the numbered OMML as one native #(SEQ) host.");
        }
        else if (pasted.Numbering.Numbered)
        {
            throw new InvalidDataException(
                "Word pasted an unnumbered OMML source as a numbered equation.");
        }

        // Do not assign FormulaId/VTEqNum during Paste. Bookmark writes and field
        // updates become separate Word undo records even inside a Custom
        // UndoRecord. The pasted native OMath is already complete Word content;
        // VisualTeX adopts private identity only when the user later edits or
        // explicitly references that equation.
        return PastedFormulaRepairResult.Repaired;
    }

    private PastedFormulaRepairResult RepairPastedHostCoreGroup(
        Document document,
        WordFormulaCopySnapshot snapshot,
        IReadOnlyList<WordFormulaHostDescriptor> candidates)
    {
        if (candidates.Count != snapshot.GroupItems.Count)
            return PastedFormulaRepairResult.NotReady;

        if (snapshot.CoreHostKind ==
            WordFormulaHostKind.VisualTeX)
        {
            var mapped =
                new List<(WordFormulaCopyGroupItem Item, WordFormulaHostDescriptor Host)>();
            foreach (var item in snapshot.GroupItems)
            {
                var matches = candidates
                    .Where(candidate =>
                        CandidateVisualTeXMatchesCopy(
                            document,
                            candidate,
                            item.Metadata.FormulaId,
                            item.CoreSemanticSignature))
                    .ToArray();
                if (matches.Length == 0)
                    return PastedFormulaRepairResult.NotReady;
                if (matches.Length != 1)
                    throw new InvalidDataException(
                        "A copied VisualTeX group item does not map to exactly one host inside the local paste region.");
                mapped.Add((item, matches[0]));
            }

            if (mapped.Select(pair => pair.Host.Range.Start)
                    .Distinct()
                    .Count()
                != mapped.Count)
                throw new InvalidDataException(
                    "Two copied VisualTeX group items resolved to the same pasted host.");

            foreach (var pair in mapped
                         .OrderByDescending(
                             item => item.Host.Range.Start))
            {
                RekeyPastedVisualTeXHost(
                    document,
                    pair.Host,
                    pair.Item.Metadata,
                    pair.Item.CoreSemanticSignature,
                    pair.Item.VisibleNumber);
            }

            return PastedFormulaRepairResult.Repaired;
        }

        if (snapshot.CoreHostKind ==
            WordFormulaHostKind.Omml)
        {
            var orderedItems =
                snapshot.GroupItems
                    .OrderBy(item => item.SourceStart)
                    .ToArray();
            var orderedHosts =
                candidates
                    .OrderBy(host => host.Range.Start)
                    .ToArray();

            for (var index = 0;
                 index < orderedItems.Length;
                 index++)
            {
                if (!CandidateOmmlMatchesCopy(
                        document,
                        orderedHosts[index],
                        orderedItems[index]
                            .CoreSemanticSignature))
                    return PastedFormulaRepairResult.NotReady;
            }

            for (var index = 0;
                 index < orderedHosts.Length;
                 index++)
            {
                var pasted = orderedHosts[index];
                pasted.Numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        pasted);
                if (orderedItems[index].Metadata.Numbered)
                {
                    if (!pasted.Display
                        || pasted.Numbering.ContainerKind !=
                            WordFormulaNumberingContainerKind.CanonicalNativeOmml)
                        throw new InvalidDataException(
                            "A pasted numbered OMML group item is not Word-native #(SEQ).");
                }
                else if (pasted.Numbering.Numbered)
                {
                    throw new InvalidDataException(
                        "A pasted unnumbered OMML group item unexpectedly became numbered.");
                }
            }

            return PastedFormulaRepairResult.Repaired;
        }

        return PastedFormulaRepairResult.NotApplicable;
    }

    private static bool CandidateVisualTeXMatchesCopy(
        Document document,
        WordFormulaHostDescriptor candidate,
        string sourceFormulaId,
        string? expectedSemanticSignature)
    {
        try
        {
            if (candidate.Kind !=
                WordFormulaHostKind.VisualTeX)
                return false;

            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    candidate);
            var embedded = payload.Metadata;
            return embedded is not null
                && string.Equals(
                    embedded.FormulaId,
                    sourceFormulaId,
                    StringComparison.OrdinalIgnoreCase)
                && string.Equals(
                    HostCoreSemanticSignature(payload),
                    expectedSemanticSignature,
                    StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
    }

    private static bool CandidateOmmlMatchesCopy(
        Document document,
        WordFormulaHostDescriptor candidate,
        string? expectedSemanticSignature)
    {
        try
        {
            if (candidate.Kind !=
                WordFormulaHostKind.Omml)
                return false;
            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    candidate);
            return string.Equals(
                HostCoreSemanticSignature(payload),
                expectedSemanticSignature,
                StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
    }

    private static void RekeyPastedVisualTeXHost(
        Document document,
        WordFormulaHostDescriptor candidate,
        FormulaMetadata copiedMetadata,
        string? expectedSemanticSignature,
        string? copiedVisibleNumber)
    {
        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            shape = GetExactVisualTeXShape(
                document,
                candidate);
            var embedded =
                WordFormulaMetadataReader
                    .TryReadEmbeddedNativeOle(shape)
                ?? throw new InvalidDataException(
                    "The pasted VisualTeX host has no authoritative embedded metadata.");

            if (!string.Equals(
                    embedded.FormulaId,
                    copiedMetadata.FormulaId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "The pasted VisualTeX host no longer carries the copied source identity.");

            var fresh =
                CloneForPastedFormula(embedded);
            fresh.DisplayMode =
                copiedMetadata.DisplayMode;
            fresh.Numbered =
                copiedMetadata.Numbered;
            fresh.NativeOmmlFingerprint = null;
            fresh.Validate();

            RemoveLocalVisualTeXIdentityBookmarks(
                shape,
                embedded.FormulaId);
            WordFormulaMetadataReader.Write(
                shape,
                fresh);
            WordFormulaIdentityStore.BindVisualTeX(
                shape,
                fresh.FormulaId);

            shapeRange = shape.Range.Duplicate;
            var refreshed =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    shapeRange,
                    WordFormulaHostKind.VisualTeX)
                ?? throw new InvalidDataException(
                    "The pasted VisualTeX host could not be resolved after re-keying.");
            refreshed.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    refreshed);

            if (fresh.Numbered)
            {
                refreshed =
                    WordFormulaNumberingKernel.RekeyPastedNumbering(
                        document,
                        refreshed,
                        copiedVisibleNumber);
            }
            else if (refreshed.Numbering.Numbered
                     || refreshed.Numbering.ContainerKind !=
                         WordFormulaNumberingContainerKind.None)
            {
                throw new InvalidDataException(
                    "An unnumbered VisualTeX copy retained a numbering container.");
            }

            var verified =
                WordFormulaHostSemanticReader.Read(
                    document,
                    refreshed);
            if (!string.Equals(
                    HostCoreSemanticSignature(verified),
                    expectedSemanticSignature,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The VisualTeX copy changed semantic content while receiving fresh identity.");

            if (fresh.Numbered)
            {
                var numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        refreshed);
                if (!numbering.Numbered
                    || numbering.ContainerKind is not
                        (WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
                         or WordFormulaNumberingContainerKind.CanonicalBodyTable
                         or WordFormulaNumberingContainerKind.CanonicalUserTableCell))
                    throw new InvalidDataException(
                        "The pasted VisualTeX number is not canonical after local re-keying.");
            }

            WordDoubleClickHook.TraceMessage(
                $"copy-paste-core-repaired kind=VisualTeX "
                + $"formulaId={fresh.FormulaId} "
                + $"range={refreshed.Range.Start}:{refreshed.Range.End} "
                + $"numbered={fresh.Numbered}");
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
        }
    }

    private static InlineShape GetExactVisualTeXShape(
        Document document,
        WordFormulaHostDescriptor host)
    {
        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? found = null;
        try
        {
            range =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            shapes = range.InlineShapes;
            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                InlineShape? candidate = null;
                Range? candidateRange = null;
                try
                {
                    candidate = shapes[index];
                    if (!WordFormulaMetadataReader
                            .IsNativeOle(candidate))
                        continue;
                    candidateRange =
                        candidate.Range.Duplicate;
                    if (!WordFormulaHostSemanticReader
                            .SameAddress(
                                candidateRange,
                                host.Range))
                        continue;
                    if (found is not null)
                        throw new InvalidDataException(
                            "More than one VisualTeX OLE occupies the pasted host range.");
                    found = candidate;
                    candidate = null;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }

            if (found is null)
                throw new InvalidDataException(
                    "The pasted VisualTeX host moved before local re-keying.");

            var returned = found;
            found = null;
            return returned;
        }
        finally
        {
            Release(found);
            Release(shapes);
            Release(range);
        }
    }

    private static void RemoveLocalVisualTeXIdentityBookmarks(
        InlineShape shape,
        string copiedFormulaId)
    {
        Range? shapeRange = null;
        Range? probe = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        try
        {
            shapeRange = shape.Range;
            probe = shapeRange.Duplicate;
            try
            {
                probe.MoveStart(
                    WdUnits.wdCharacter,
                    -1);
            }
            catch { }
            try
            {
                probe.MoveEnd(
                    WdUnits.wdCharacter,
                    1);
            }
            catch { }

            bookmarks = probe.Bookmarks;
            for (var index = bookmarks.Count;
                 index >= 1;
                 index--)
            {
                Release(bookmarkRange);
                bookmarkRange = null;
                Release(bookmark);
                bookmark = bookmarks[index];
                if (!WordFormulaMetadataReader
                        .TryFormulaIdFromIdentityBookmark(
                            bookmark.Name,
                            out var formulaId)
                    || !string.Equals(
                        formulaId,
                        copiedFormulaId,
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                bookmarkRange =
                    bookmark.Range.Duplicate;
                var ownsShape =
                    (bookmarkRange.Start ==
                         shapeRange.Start
                     && bookmarkRange.End ==
                         shapeRange.End)
                    || (bookmarkRange.Start ==
                            bookmarkRange.End
                        && bookmarkRange.Start ==
                            shapeRange.Start)
                    || (bookmarkRange.Start <=
                            shapeRange.Start
                        && bookmarkRange.End >=
                            shapeRange.End);
                if (ownsShape)
                    bookmark.Delete();
            }
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(probe);
            Release(shapeRange);
        }
    }

    private static FormulaMetadata CloneForPastedFormula(
        FormulaMetadata source)
    {
        var now =
            DateTimeOffset.UtcNow.ToString("O");
        var clone =
            WordFormulaMetadataReader.CloneWithFormulaId(
                source,
                Guid.NewGuid().ToString("D"));
        clone.Lines = clone.Lines
            .Select(line => new FormulaLine
            {
                Id = Guid.NewGuid().ToString("D"),
                Latex = line.Latex,
            })
            .ToList();
        clone.NativeOmmlFingerprint = null;
        clone.CreatedWithVersion = "1.2.7";
        clone.UpdatedWithVersion = "1.2.7";
        clone.CreatedAt = now;
        clone.UpdatedAt = now;
        clone.Validate();
        return clone;
    }

    private static FormulaMetadata CreateHostCoreSnapshotMetadata(
        WordFormulaHostDescriptor host,
        WordFormulaSemanticPayload payload)
    {
        FormulaMetadata metadata;
        if (host.Kind ==
            WordFormulaHostKind.VisualTeX)
        {
            metadata = CloneFormulaMetadata(
                payload.Metadata
                ?? throw new InvalidDataException(
                    "The VisualTeX copy source has no authoritative embedded metadata."));
        }
        else
        {
            var formulaId =
                Guid.TryParse(
                    host.FormulaId,
                    out var parsed)
                    ? parsed.ToString("D")
                    : Guid.NewGuid().ToString("D");
            var now =
                DateTimeOffset.UtcNow.ToString("O");
            metadata = new FormulaMetadata
            {
                FormulaId = formulaId,
                Latex = payload.Latex,
                Lines = new List<FormulaLine>
                {
                    new()
                    {
                        Id =
                            Guid.NewGuid()
                                .ToString("D"),
                        Latex = payload.Latex,
                    },
                },
                CodeFormat = "latex",
                DisplayMode = host.DisplayMode,
                Numbered = false,
                CreatedWithVersion = "1.2.7",
                UpdatedWithVersion = "1.2.7",
                CreatedAt = now,
                UpdatedAt = now,
            };
        }

        metadata.DisplayMode =
            host.DisplayMode;
        metadata.Numbered =
            host.Numbering.Numbered;
        metadata.NativeOmmlFingerprint = null;
        metadata.Validate();
        return metadata;
    }

    private static string HostCoreSemanticSignature(
        WordFormulaSemanticPayload payload)
    {
        var latex =
            NormalizeHostCoreSemanticText(
                payload.Latex);
        if (latex.Length > 0)
            return "latex:" + latex;

        return "mathml:"
            + NormalizeHostCoreSemanticText(
                payload.MathMl ?? string.Empty);
    }

    private static string NormalizeHostCoreSemanticText(
        string value)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;
        return string.Concat(
            value.Where(
                character =>
                    !char.IsWhiteSpace(character)));
    }

    private static string? ReadHostCoreVisibleNumber(
        Document document,
        WordFormulaNumberingDescriptor numbering)
    {
        if (!numbering.Numbered
            || numbering.NumberRange is null)
            return null;

        Range? number = null;
        try
        {
            number =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    numbering.NumberRange);
            return (number.Text ?? string.Empty)
                .Trim();
        }
        finally
        {
            Release(number);
        }
    }

    private static WordCharacterFormatting?
        TryCaptureHostCoreParagraphFormatting(
            Document document,
            WordFormulaRangeAddress address)
    {
        Range? range = null;
        try
        {
            range =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    address);
            return TryCaptureParagraphMarkFormatting(
                range);
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(range);
        }
    }

    private static bool TryRepairMergedPastedOmmlWithHostCore(
        Document document,
        WordFormulaCopySnapshot snapshot,
        Range pastedRegion)
    {
        if (!snapshot.CorePendingOmmlMergeTarget)
            return false;

        Range? probe = null;
        try
        {
            probe =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    new WordFormulaRangeAddress
                    {
                        StoryType =
                            snapshot.CorePendingPasteStoryType,
                        Start =
                            snapshot.CorePendingOmmlTargetStart,
                        End =
                            snapshot.CorePendingOmmlTargetStart,
                    });

            var merged =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    probe,
                    WordFormulaHostKind.Omml);
            if (merged is null
                || !string.Equals(
                    merged.DisplayMode,
                    "inline",
                    StringComparison.Ordinal))
                return false;

            // The exact physical OMath must now own both the old host position and
            // the complete region inserted by this native Paste. That is Word's
            // natural "same math zone" result; no hidden delimiter is synthesized.
            if (merged.Range.StoryType !=
                    snapshot.CorePendingPasteStoryType
                || merged.Range.Start >
                    snapshot.CorePendingOmmlTargetStart
                || merged.Range.Start >
                    pastedRegion.Start
                || merged.Range.End <
                    pastedRegion.End)
                return false;

            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    merged);
            var afterSignature =
                HostCoreSemanticSignature(payload);
            if (string.Equals(
                    afterSignature,
                    snapshot.CorePendingOmmlTargetSemanticSignature,
                    StringComparison.Ordinal))
                return false;

            var afterLatex =
                NormalizeHostCoreSemanticText(
                    payload.Latex);
            var beforeLatex =
                NormalizeHostCoreSemanticText(
                    snapshot.CorePendingOmmlTargetLatex
                    ?? string.Empty);
            var copiedLatex =
                NormalizeHostCoreSemanticText(
                    snapshot.CoreLatex
                    ?? snapshot.Metadata.Latex
                    ?? string.Empty);
            if (beforeLatex.Length > 0
                && afterLatex.IndexOf(
                    beforeLatex,
                    StringComparison.Ordinal) < 0)
                return false;
            if (copiedLatex.Length > 0
                && afterLatex.IndexOf(
                    copiedLatex,
                    StringComparison.Ordinal) < 0)
                return false;

            // OMML has no VisualTeX ownership layer. Native Paste repair is
            // validation-only; the merged OMath remains pure Word content.
            WordDoubleClickHook.TraceMessage(
                $"copy-paste-core-natural-omml-merge "
                + $"range={merged.Range.Start}:{merged.Range.End}");
            return true;
        }
        catch (InvalidDataException)
        {
            return false;
        }
        finally
        {
            Release(probe);
        }
    }

    internal bool RollbackHostCorePaste(
        WordFormulaCopySnapshot snapshot,
        out string? error)
    {
        error = null;
        if (snapshot is null || !snapshot.UsesHostCore)
            return false;

        Document? document = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null
                || !string.Equals(
                    DocumentIdentity(document),
                    snapshot.CorePendingPasteDocumentId,
                    StringComparison.OrdinalIgnoreCase))
            {
                error =
                    "The active document is not the document that owns the paste transaction.";
                return false;
            }

            object times = 1;
            if (!document.Undo(ref times))
            {
                error =
                    "Word returned false while undoing the paste transaction.";
                return false;
            }

            return true;
        }
        catch (Exception undoError)
        {
            error =
                $"{undoError.GetType().Name}:{undoError.Message}";
            return false;
        }
        finally
        {
            CancelHostCorePaste(snapshot);
            Release(document);
        }
    }

    internal void CancelHostCorePaste(
        WordFormulaCopySnapshot snapshot)
    {
        if (snapshot is null || !snapshot.UsesHostCore)
            return;
        snapshot.CorePendingPasteArmed = false;
        snapshot.CorePendingPasteDocumentId =
            string.Empty;
        snapshot.CorePendingPasteStart = 0;
        snapshot.CorePendingPasteEnd = 0;
        snapshot.CorePendingOmmlMergeTarget = false;
        snapshot.CorePendingOmmlTargetStart = 0;
        snapshot.CorePendingOmmlTargetEnd = 0;
        snapshot.CorePendingOmmlTargetFormulaId = null;
        snapshot.CorePendingOmmlTargetSemanticSignature = null;
        snapshot.CorePendingOmmlTargetLatex = null;
    }

    private static void CompleteHostCorePaste(
        Document document,
        WordFormulaCopySnapshot snapshot)
    {
        snapshot.CorePendingPasteArmed = false;
        snapshot.CorePendingPasteDocumentId =
            string.Empty;
        snapshot.CorePendingPasteStart = 0;
        snapshot.CorePendingPasteEnd = 0;
        snapshot.CorePendingOmmlMergeTarget = false;
        snapshot.CorePendingOmmlTargetStart = 0;
        snapshot.CorePendingOmmlTargetEnd = 0;
        snapshot.CorePendingOmmlTargetFormulaId = null;
        snapshot.CorePendingOmmlTargetSemanticSignature = null;
        snapshot.CorePendingOmmlTargetLatex = null;
        snapshot.TrackingDocumentId =
            DocumentIdentity(document);
        snapshot.KnownDocumentEnd =
            document.Content.End;
    }
}
