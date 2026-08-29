using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlNativeHashCopyPasteAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-native-hash-copy-paste.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? sourceRange = null;
        Word.Range? pastedRange = null;
        var originalFormulaId = Guid.NewGuid().ToString("D");
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            var service = new WordFormulaService(application);
            Word.Range? insertion = null;
            try
            {
                insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
                application.Selection.SetRange(insertion.Start, insertion.End);
                var session = CreateNumberedOmmlTabSession(
                    originalFormulaId,
                    document.FullName,
                    insertion.Start,
                    insertion.End,
                    @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                    originalMetadata: null);
                service.InsertOmml(session, QuadraticFormulaMathMl());
            }
            finally { Release(insertion); }

            UpdateNativeHashProductionFields(document, new[] { originalFormulaId });
            var originalReference = InsertNativeHashProductionReferences(
                document,
                new[] { originalFormulaId });
            AssertNativeHashProductionReferences(
                document,
                originalReference,
                new[] { originalFormulaId },
                new[] { "1" },
                "copy/paste source REF before copy");

            var sourceMetadata = WordOmmlFormulaStore.TryRead(
                    document,
                    originalFormulaId)
                ?? throw new InvalidDataException(
                    "Copy/paste source numbered OMML lost its metadata.");
            sourceRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                originalFormulaId,
                sourceMetadata);
            sourceRange.Copy();
            // Use the real Word UI paste semantics. Range.Paste and
            // Selection.Paste are not equivalent for professional OMath: on some
            // Word builds Range.Paste drops the copied mathematical bookmark
            // aliases while Ctrl+V/Selection.Paste preserves the native #(SEQ)
            // structure long enough for VisualTeX adoption to rekey it.
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            var pastedStart = application.Selection.Start;
            application.Selection.Paste();
            Release(sourceRange);
            sourceRange = null;

            // FormulaId de-duplication belongs to VisualTeX's OMML adoption path:
            // after a user paste, selecting/opening the copied OMath makes
            // ReadSelection resolve the duplicated native #(SEQ) host and rekey it.
            // Reconcile maintains numbering structure after that identity decision;
            // it intentionally does not invent a new logical FormulaId on its own.
            // A body REF targeting VTEqNum can itself appear through
            // document.OMaths, so do not use the global OMath count to locate the
            // copy. Probe only the tail range beginning at the exact paste point;
            // that range contains the pasted formula and no earlier body REF.
            Word.Range? pastedProbe = null;
            Word.OMaths? pastedMaths = null;
            Word.OMath? pastedMath = null;
            try
            {
                pastedProbe = document.Range(pastedStart, document.Content.End);
                pastedMaths = pastedProbe.OMaths;
                AssertTrue(pastedMaths.Count >= 1,
                    "Copy/paste did not expose a native OMath at the paste point.");
                pastedMath = pastedMaths[1];
                pastedRange = pastedMath.Range.Duplicate;
            }
            finally
            {
                Release(pastedMath);
                Release(pastedMaths);
                Release(pastedProbe);
            }
            AssertEqual(
                1,
                WordOmmlFormulaStore.BookmarkedFormulaIds(document)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count(),
                "The pasted OMath unexpectedly acquired a FormulaId before structural reconciliation.");
            Release(pastedRange);
            pastedRange = null;

            // Do not select or open the pasted formula first. The document-level
            // reconciliation path must prove that this is a physical surplus copy of
            // an existing managed native #(SEQ) formula, allocate a fresh FormulaId,
            // and rebuild only the copy's aliases without user interaction.
            WordEquationNumbering.Reconcile(document);

            var formulaIds = WordOmmlFormulaStore.BookmarkedFormulaIds(document)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            AssertEqual(2, formulaIds.Length,
                "Copy/paste reconciliation did not produce two unique VTOMML FormulaIds.");
            AssertTrue(formulaIds.Contains(
                    originalFormulaId,
                    StringComparer.OrdinalIgnoreCase),
                "Copy/paste reconciliation changed the source FormulaId.");
            var copiedFormulaId = formulaIds.Single(id =>
                !string.Equals(id, originalFormulaId, StringComparison.OrdinalIgnoreCase));
            AssertTrue(!string.Equals(
                    copiedFormulaId,
                    originalFormulaId,
                    StringComparison.OrdinalIgnoreCase),
                "The pasted numbered OMML retained the source FormulaId.");

            // Selecting the already-adopted copy afterwards must resolve the same
            // identity rather than assigning a second FormulaId.
            var copiedMetadata = WordOmmlFormulaStore.TryRead(document, copiedFormulaId)
                ?? throw new InvalidDataException(
                    "The automatically adopted native #(SEQ) copy lost metadata.");
            pastedRange = WordOmmlFormulaStore
                .GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    copiedFormulaId,
                    copiedMetadata);
            pastedRange.Select();
            var copiedSelection = service.ReadSelection();
            AssertTrue(copiedSelection.Metadata is not null,
                "Selecting the automatically adopted copy lost its VisualTeX metadata.");
            AssertTrue(copiedSelection.Metadata!.Numbered,
                "Selecting the automatically adopted copy changed it to unnumbered.");
            AssertEqual(copiedFormulaId, copiedSelection.FormulaId,
                "Selecting the automatically adopted copy assigned a second FormulaId.");

            var orderedFormulaIds = formulaIds
                .Select(id =>
                {
                    Word.Range? owner = null;
                    try
                    {
                        owner = WordEquationNumbering.FindNumberingOwnerRange(document, id)
                            ?? throw new InvalidDataException(
                                $"Copy/paste formula {id} has no numbering owner.");
                        return (FormulaId: id, Position: owner.Start);
                    }
                    finally { Release(owner); }
                })
                .OrderBy(item => item.Position)
                .Select(item => item.FormulaId)
                .ToArray();
            UpdateNativeHashProductionFields(document, orderedFormulaIds);
            AssertNativeHashProductionNumbers(
                document,
                orderedFormulaIds,
                new[] { "1", "2" },
                "copy/paste FormulaId de-duplication");
            AssertNativeHashProductionReferences(
                document,
                originalReference,
                new[] { originalFormulaId },
                new[] { "1" },
                "copy/paste source REF after de-duplication");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            ForceComCleanup();

            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            UpdateNativeHashProductionFields(document, orderedFormulaIds);
            AssertNativeHashProductionNumbers(
                document,
                orderedFormulaIds,
                new[] { "1", "2" },
                "copy/paste save/reopen identities");
            AssertNativeHashProductionReferences(
                document,
                originalReference,
                new[] { originalFormulaId },
                new[] { "1" },
                "copy/paste save/reopen source REF");

            Console.WriteLine(
                $"Production OMML native #(SEQ) copy/paste acceptance passed: source={originalFormulaId}, copied={copiedFormulaId}; FormulaIds/bookmarks were unique, numbering was 1/2, and the source body REF stayed attached across save/reopen.");
        }
        finally
        {
            Release(pastedRange);
            Release(sourceRange);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }
}
