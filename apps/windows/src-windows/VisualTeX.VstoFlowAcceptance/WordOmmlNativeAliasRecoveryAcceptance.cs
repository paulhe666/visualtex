using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlNativeAliasRecoveryAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-native-alias-recovery.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.Field? externalReference = null;
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
            var formulaId = Guid.NewGuid().ToString("D");
            insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            application.Selection.SetRange(insertion.Start, insertion.End);
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertion.Start,
                insertion.End,
                latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);
            service.InsertOmml(session, QuadraticFormulaMathMl());
            Release(insertion); insertion = null;
            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                "native alias-recovery initial host",
                updateReference: true);

            externalReference = InsertExternalEquationReference(document, formulaId);
            var expectedNumber = NormalizeEquationNumberText(
                externalReference.Result.Text);

            DeleteEquationNumberAlias(
                document,
                WordEquationNumbering.EquationBookmarkName(formulaId));
            DeleteEquationNumberAlias(
                document,
                WordEquationNumbering.NativeCaptionBookmarkName(formulaId));
            DeleteEquationNumberAlias(
                document,
                WordEquationNumbering.NativeNumberBookmarkName(formulaId));

            AssertTrue(
                WordOmmlFormulaStore.BookmarkedFormulaIds(document).Contains(
                    formulaId,
                    StringComparer.OrdinalIgnoreCase),
                "Alias-loss injection unexpectedly removed the VTOMML formula identity.");
            var retainedMetadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "Alias-loss injection unexpectedly removed OMML metadata.");
            AssertTrue(retainedMetadata.Numbered,
                "Alias-loss injection unexpectedly changed Numbered metadata.");
            AssertTrue(
                WordEquationNumbering.NeedsLegacyManagedNumberingMigration(document),
                "A numbered VTOMML formula with all three number aliases missing was not detected for recovery.");

            AssertEqual(
                1,
                WordEquationNumbering.RefreshNumberedOmmlTabLayouts(document),
                "Alias recovery did not rebuild exactly one managed numbered OMML formula.");
            WordEquationNumbering.UpdateEquationNumbers(document);
            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                "native alias-recovery rebuilt host",
                updateReference: true);

            externalReference.Update();
            AssertEqual(
                expectedNumber,
                NormalizeEquationNumberText(externalReference.Result.Text),
                "The existing body REF did not reconnect after VTEqNum was rebuilt.");
            AssertExternalEquationReferenceHyperlink(
                externalReference,
                WordEquationNumbering.NativeNumberBookmarkName(formulaId));
            AssertTrue(
                !WordEquationNumbering.NeedsLegacyManagedNumberingMigration(document),
                "The rebuilt native alias host is still classified as legacy or malformed.");

            document.Save();
            Release(externalReference); externalReference = null;
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                "native alias-recovery save/reopened host",
                updateReference: true);
            externalReference = FindExternalEquationReference(document, formulaId)
                ?? throw new InvalidDataException(
                    "Alias recovery save/reopen lost the existing body REF field.");
            externalReference.Update();
            AssertEqual(
                expectedNumber,
                NormalizeEquationNumberText(externalReference.Result.Text),
                "Alias recovery save/reopen disconnected the body REF field.");

            Console.WriteLine(
                "Native OMML alias-recovery acceptance passed: losing VTEq/VTEqCap/VTEqNum triggered managed atomic reconstruction, the pre-existing REF reconnected, and save/reopen retained zero Shape/Table artifacts.");
        }
        finally
        {
            Release(externalReference);
            Release(insertion);
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

    private static void DeleteEquationNumberAlias(
        Word.Document document,
        string bookmarkName)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            AssertTrue(bookmarks.Exists(bookmarkName),
                $"Alias-loss injection could not find {bookmarkName}.");
            bookmark = bookmarks[bookmarkName];
            bookmark.Delete();
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }
}
