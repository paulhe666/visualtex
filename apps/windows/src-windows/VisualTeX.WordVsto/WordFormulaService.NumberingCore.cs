using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;

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
                    count +=
                        WordFormulaNumberingKernel.RefreshCanonicalNumbers(
                            document);
                    return count;
                });
        }
        finally
        {
            Release(document);
        }
    }

    internal int SetEquationNumberFormatCore(
        string formatId)
    {
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
            var mathTypePlan =
                MathTypeEquationNumbering
                    .PrepareEquationNumberFormat(
                        document,
                        resolvedFormat);

            selection = _application.Selection;
            originalSelection =
                selection.Range.Duplicate;
            var formatVariable =
                WordEquationNumbering
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
                                                currentParagraphStarts));
                        }

                        // Track the user's Selection only after MathType has
                        // finished rewriting its field trees. This bookmark is
                        // transient UI state and is always deleted below.
                        bookmarks = document.Bookmarks;
                        tracking = bookmarks.Add(
                            trackingName,
                            originalSelection);

                        count +=
                            WordFormulaNumberingKernel
                                .SetCanonicalNumberFormat(
                                    document,
                                    resolvedFormat);

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
                        WordEquationNumbering
                            .RestoreEquationNumberFormatVariable(
                                document,
                                formatVariable));

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
            catch { }

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
        if (target.Source ==
            EquationReferenceSource.MathType)
        {
            ExecuteDocumentEdit(
                document,
                "VisualTeX Insert MathType Equation Reference",
                () =>
                {
                    MathTypeEquationReferences.InsertReference(
                        document,
                        selection,
                        target,
                        color);
                    return 0;
                });
            return;
        }

        var (prefix, suffix) =
            style switch
            {
                EquationReferenceStyle.Parenthesized =>
                    ("(", ")"),
                EquationReferenceStyle.EquationPrefix =>
                    (OfficePluginLanguage.Text(
                        "式（",
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
                    var current =
                        WordFormulaNumberingKernel
                            .GetCanonicalReferenceTargets(
                                document)
                            .SingleOrDefault(candidate =>
                                candidate.Source ==
                                    EquationReferenceSource.WordOmml
                                && candidate.Position ==
                                    target.Position
                                && candidate.EndPosition ==
                                    target.EndPosition)
                        ?? throw new InvalidDataException(
                            "The selected native Word equation moved before reference insertion.");

                    WordEquationReferenceFields
                        .InsertNativeWordEquationReference(
                            document,
                            selection,
                            current.NativeReferenceItem,
                            prefix,
                            suffix,
                            color);
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
                EnsureVisualTeXReferenceTargetExists(
                    document,
                    target,
                    bookmarkName);
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
        var host =
            WordFormulaHostResolver
                .CaptureDocumentIndex(
                    document)
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

    internal IReadOnlyList<EquationReferenceTarget>
        GetCanonicalEquationReferenceTargets(
            Document document) =>
        WordFormulaNumberingKernel
            .GetCanonicalReferenceTargets(
                document);
}
