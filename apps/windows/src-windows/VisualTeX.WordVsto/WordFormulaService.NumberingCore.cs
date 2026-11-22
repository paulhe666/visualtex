using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using static VisualTeX.WordVsto.WordMathTypeNumbering;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    internal int UpdateEquationNumbersCore()
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);
            EnsureEquationFieldResultsVisible(document);

            // Native host discovery, literal-heading resolution and bookmark
            // evidence are read-only preflight. WordOpenXML must not split the
            // number-refresh UndoRecord after any fields have been changed.
            using var nativeNumbers = WordNativeNumberingAdoptionPlan.Prepare(document);
            return ExecuteFieldRefreshEdit(
                document,
                "VisualTeX Update Equation Numbers",
                () =>
                {
                    // MathType keeps its native MTPlaceRef/MTEqn subsystem.
                    // OMML/VisualTeX use only the canonical numbering kernel.
                    var count =
                        MathTypeEquationNumbering.UpdateEquationNumbers(
                            document);
                    nativeNumbers.Apply();
                    count += nativeNumbers.Count;
                    // Retain the established VisualTeX field-stream updater,
                    // including standalone canonical sequences and REF refresh.
                    // It need not rediscover/serialize native OMaths again.
                    count += WordFormulaNumberingKernel.RefreshCanonicalNumbers(
                        document, kindFilter: WordFormulaHostKind.VisualTeX);
                    // display formulas now in another column get that column's
                    // tab stops (documents with newspaper columns only)
                    ReflowDisplayTabStops(document, apply: true, multiColumnOnly: true);
                    return count;
                }, verifyRollback: nativeNumbers.VerifyRollback);
        }
        finally
        {
            Release(document);
        }
    }

    internal int SetEquationNumberFormatCore(
        string formatId)
    {
        WordEquationNumberFormatAuthority.InvalidateOperationScope();
        Document? document = null;
        Selection? selection = null;
        Range? originalSelection = null;
        Bookmarks? bookmarks = null;
        Bookmark? tracking = null;
        Range? tracked = null;
        var trackingName =
            "VTFmt_"
            + Guid.NewGuid().ToString("N");
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);
            EnsureEquationFieldResultsVisible(document);

            var resolvedFormat =
                EquationNumberFormat.Resolve(
                    formatId).Id;
            using var mathTypePlan =
                MathTypeEquationNumbering
                    .PrepareEquationNumberFormat(
                        document,
                        resolvedFormat);
            // The same prepared native section paragraph can be retargeted to
            // each resolved chapter/section by the existing field writer. Build
            // it before Undo; creating/editing another Word document mid-record
            // silently ends the target record even when no XML is read there.
            using var mathTypeSection = mathTypePlan.Count > 0
                && EquationNumberFormat.Resolve(resolvedFormat).UsesHeading
                ? new MathTypeSectionStaging(document, CreateMathTypeSectionFlatOpc(0, 0))
                : null;
            using var nativeNumbers = WordNativeNumberingAdoptionPlan.Prepare(
                document, numberFormat: resolvedFormat);
            var hasVisualTeXNumbers = WordNumberingFieldCollection.HasVisualTeXNumbering(document);

            selection = _application.Selection;
            originalSelection =
                selection.Range.Duplicate;
            var formatVariable =
                WordEquationNumberFormat
                    .CaptureEquationNumberFormatVariable(
                        document);

            var result =
                ExecuteFieldRefreshEdit(
                    document,
                    "VisualTeX Set Equation Number Format",
                    () =>
                    {
                        var count = 0;
                        if (mathTypePlan.ParagraphStarts.Count > 0)
                        {
                            count +=
                                MathTypeEquationNumbering
                                    .SetEquationNumberFormat(
                                        document,
                                        mathTypePlan,
                                        currentParagraphStarts =>
                                            EnsureExistingMathTypeHeadingScopes(
                                                document,
                                                resolvedFormat,
                                                currentParagraphStarts,
                                                mathTypeSection?.Paragraph));
                        }

                        // Track the user's Selection only after MathType has
                        // finished rewriting its field trees. This bookmark is
                        // transient UI state and is always deleted below.
                        bookmarks = document.Bookmarks;
                        tracking = bookmarks.Add(
                            trackingName,
                            originalSelection);

                        WordEquationNumberFormat.RestoreEquationNumberFormatForConversion(document, resolvedFormat);
                        nativeNumbers.Apply();
                        count += nativeNumbers.Count;
                        // Native number ownership/serialization was prepared
                        // outside Undo. Do not build an unrelated all-OMath/OLE
                        // document index merely to discover that there are no VT
                        // numbered objects. Existing VT layouts keep their own
                        // family-specific writer and validation when present.
                        if (hasVisualTeXNumbers)
                            count += WordFormulaNumberingKernel.RefreshCanonicalNumbers(
                                document, resolvedFormat, kindFilter: WordFormulaHostKind.VisualTeX);

                        Release(tracking);
                        tracking = null;
                        if (bookmarks.Exists(
                                trackingName))
                        {
                            tracking =
                                bookmarks[trackingName];
                            tracked =
                                tracking.Range.Duplicate;
                            selection.SetRange(
                                tracked.Start,
                                tracked.End);
                            tracking.Delete();
                        }

                        return count;
                    },
                    restoreNonUndoState: () =>
                        WordEquationNumberFormat
                            .RestoreEquationNumberFormatVariable(
                                document,
                                formatVariable),
                    verifyRollback: nativeNumbers.VerifyRollback);

            WordEquationNumberFormat.SetDefaultEquationNumberFormatPreference(resolvedFormat);
            return result;
        }
        finally
        {
            try
            {
                if (bookmarks is not null
                    && bookmarks.Exists(
                        trackingName))
                {
                    Release(tracking);
                    tracking =
                        bookmarks[trackingName];
                    tracking.Delete();
                }
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }

            Release(tracked);
            Release(tracking);
            Release(bookmarks);
            Release(originalSelection);
            Release(selection);
            Release(document);
        }
    }

    internal void InsertEquationReferenceCore(
        Document document,
        Selection selection,
        EquationReferenceTarget target,
        EquationReferenceStyle style,
        WdColor color)
    {
        using var metric =
            VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure(
                "Reference.Insert");
        if (target.Source ==
            EquationReferenceSource.MathType)
        {
            // A MathType reference adds one field at the caret and, the first time an equation is
            // referenced, one bookmark on its number. Its own undo entries are the whole rollback.
            // (This edit used to take the whole-document recovery snapshot: the complete document
            // exported as Word XML and parsed three times, and every inline object walked twice,
            // before each reference - a minute and more in a long thesis full of MathType objects.)
            ExecuteLocalReferenceEdit(
                document,
                "VisualTeX Insert MathType Equation Reference",
                () =>
                    MathTypeEquationReferences.InsertReference(
                        document,
                        selection,
                        target,
                        color));
            return;
        }

        var (prefix, suffix) =
            style switch
            {
                EquationReferenceStyle.Parenthesized =>
                    ("(", ")"),
                EquationReferenceStyle.EquationPrefix =>
                    (OfficePluginLanguage.Text(
                        "式(",
                        "Eq. ("),
                     ")"),
                _ => (string.Empty, string.Empty),
            };

        if (target.Source ==
            EquationReferenceSource.WordOmml)
        {
            WordFormulaMutationTransaction.Execute(
                _application,
                document,
                "VisualTeX Insert Native Word Equation Reference",
                () =>
                {
                    // (the target re-read in its own paragraph, as the whole-document listing reads it)
                    var current =
                        GetTargetsAround(
                                document,
                                target)
                            .SingleOrDefault(candidate =>
                                candidate.Source ==
                                    EquationReferenceSource.WordOmml
                                && candidate.Position ==
                                    target.Position
                                && candidate.EndPosition ==
                                    target.EndPosition)
                        ?? throw new InvalidDataException(
                            "The selected native Word equation moved before reference insertion.");

                    Range? currentRange = null;
                    try
                    {
                        currentRange = document.Range(
                            current.Position,
                            current.EndPosition);
                        var currentHost =
                            WordFormulaHostLookup.ResolveLocal(
                                document,
                                currentRange,
                                WordFormulaHostKind.Omml)
                            ?? throw new InvalidDataException(
                                "The selected native Word equation no longer exists at its captured range.");
                        if (currentHost.Range.Start != current.Position
                            || currentHost.Range.End != current.EndPosition)
                            throw new InvalidDataException(
                                "The selected native Word equation changed range before reference insertion.");

                        // Word's InsertCrossReference can bind an OMML Equation
                        // item to the complete OMath host, causing the REF result
                        // to reproduce the formula body. Pure OMML already has a
                        // canonical native number range, so reference that exact
                        // range through Word's ordinary hidden _Ref bookmark.
                        // This adds no VisualTeX identity/numbering metadata and
                        // keeps the displayed REF strictly number-only.
                        var nativeBookmark =
                            WordEquationReferenceFields
                                .CreateNativeOmmlNumberReferenceBookmark(
                                    document,
                                    currentHost);
                        WordEquationReferenceFields
                            .InsertDirectReference(
                                document,
                                selection,
                                nativeBookmark,
                                prefix,
                                suffix,
                                color);
                    }
                    finally
                    {
                        Release(currentRange);
                    }
                    return true;
                });
            return;
        }

        var bookmarkName =
            WordFormulaNumberingKernel.ReferenceBookmarkName(
                target.FormulaId);
        WordFormulaMutationTransaction.Execute(
            _application,
            document,
            "VisualTeX Insert Equation Reference",
            () =>
            {
                using (VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("Reference.Insert.EnsureTarget"))
                    EnsureVisualTeXReferenceTargetExists(
                        document,
                        target,
                        bookmarkName);
                using (VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("Reference.Insert.DirectReference"))
                    WordEquationReferenceFields
                        .InsertDirectReference(
                            document,
                            selection,
                            bookmarkName,
                            prefix,
                            suffix,
                            color);
                return true;
            });
    }

    /// One Word undo record around a local reference edit. On failure every entry the edit left in
    /// Word's undo list is undone (Field.Update inside the record can make Word split it into
    /// several entries) and the story length is checked; no copy of the document is taken first.
    private void ExecuteLocalReferenceEdit(
        Document document,
        string name,
        System.Action edit)
    {
        Range? content = null;
        int storyEnd;
        try
        {
            content = document.Content;
            storyEnd = content.End;
        }
        finally { Release(content); }
        WordUndoHistorySnapshot history;
        using (VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("Reference.Insert.UndoHistory"))
            history = new WordUndoHistorySnapshot(document);
        var undo = BeginUndoRecord(name)
            ?? throw new InvalidOperationException($"Word could not establish an independent transaction for '{name}'.");
        var ended = false;
        try
        {
            edit();
            EndUndoRecord(undo);
            ended = true;
        }
        catch (Exception editError)
        {
            try
            {
                if (!ended)
                {
                    EndUndoRecord(undo);
                    ended = true;
                }
                history.UndoChanges(document);
                Range? restored = null;
                try
                {
                    restored = document.Content;
                    if (restored.End != storyEnd)
                        throw new InvalidDataException("Word undo did not restore the document's length.");
                }
                finally { Release(restored); }
            }
            catch (Exception recoveryError)
            {
                throw new AggregateException($"'{name}' failed and recovery could not be verified.", editError, recoveryError);
            }
            throw;
        }
        finally
        {
            if (!ended) EndUndoRecord(undo);
            Release(undo);
        }
    }

    private static void EnsureVisualTeXReferenceTargetExists(
        Document document,
        EquationReferenceTarget target,
        string bookmarkName)
    {
        if (target.Source !=
            EquationReferenceSource.VisualTeX)
            throw new InvalidDataException(
                "Only VisualTeX OLE references use a VTEqNum target.");

        Bookmarks? bookmarks = null;
        try
        {
            bookmarks =
                document.Bookmarks;
            if (bookmarks.Exists(
                    bookmarkName))
                return;
        }
        finally
        {
            Release(bookmarks);
        }

        // New VisualTeX paragraph numbering deliberately creates no permanent
        // bookmark. Match MathType's behavior: materialize VTEqNum_<FormulaId>
        // only when the user actually inserts a body reference.
        using var captureMetric =
            VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure(
                "Reference.Insert.CaptureIndexForBookmark");
        // (the host indexed in its own paragraph: the same identity anchors and table boundaries
        // as a whole-document index there, without indexing every other formula)
        Range? scope = null;
        WordFormulaHostDescriptor host;
        try
        {
            scope = TargetParagraphs(document, target);
            host =
            WordFormulaHostResolver
                .CaptureScopeIndex(
                    document,
                    scope)
                .VisualTeX
                .SingleOrDefault(candidate =>
                    string.Equals(
                        candidate.FormulaId,
                        target.FormulaId,
                        StringComparison.OrdinalIgnoreCase)
                    && candidate.Range.Start ==
                        target.Position
                    && candidate.Range.End ==
                        target.EndPosition)
            ?? throw new InvalidDataException(
                "The selected VisualTeX equation reference target moved before bookmark creation.");
        }
        finally
        {
            Release(scope);
        }

        var numbering =
            WordFormulaNumberingResolver
                .ResolveLocal(
                    document,
                    host);
        if (numbering.ContainerKind !=
                WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
            || !numbering.Numbered)
            throw new InvalidDataException(
                "The selected VisualTeX equation is not a canonical self-contained numbered paragraph.");

        WordVisualTeXParagraphNumbering
            .EnsureReferenceBookmark(
                document,
                host,
                bookmarkName);
    }

    /// The paragraphs holding a captured reference target (the main story).
    private static Range TargetParagraphs(
        Document document,
        EquationReferenceTarget target)
    {
        var scope =
            document.Range(
                target.Position,
                Math.Max(
                    target.Position,
                    target.EndPosition));
        scope.Expand(
            WdUnits.wdParagraph);
        return scope;
    }

    private static IReadOnlyList<EquationReferenceTarget>
        GetTargetsAround(
            Document document,
            EquationReferenceTarget target)
    {
        Range? scope = null;
        try
        {
            scope = TargetParagraphs(document, target);
            // (positions only: the insertion reads no preview)
            return WordFormulaNumberingKernel
                .GetCanonicalReferenceTargets(
                    document,
                    scope,
                    deferPreviews: true);
        }
        finally
        {
            Release(scope);
        }
    }

    /// Every reference target of the document: the VisualTeX and Word OMML ones, then MathType's. A
    /// MathType number counts only beside a MathType object, so the MathType pass over every field of
    /// the document runs only when the index met an inline OLE object that is not VisualTeX's.
    internal IReadOnlyList<EquationReferenceTarget>
        GetEquationReferenceTargets(
            Document document,
            bool deferPreviews = false)
    {
        var canonical =
            WordFormulaNumberingKernel
                .GetCanonicalReferenceTargets(
                    document,
                    scope: null,
                    deferPreviews,
                    out var hasForeignOle);
        return hasForeignOle
            ? canonical
                .Concat(MathTypeEquationReferences.GetTargets(document))
                .ToArray()
            : canonical;
    }
}
