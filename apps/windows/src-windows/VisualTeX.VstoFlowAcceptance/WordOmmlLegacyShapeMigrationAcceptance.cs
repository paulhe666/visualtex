using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlLegacyShapeMigrationAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-legacy-shape-migration.docx");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Bookmark? formulaBookmark = null;
        Word.Range? formulaRange = null;
        Word.Field? externalReference = null;
        Word.Range? content = null;
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
            var insertion = document.Content.End - 1;
            application.Selection.SetRange(insertion, insertion);
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertion,
                insertion,
                @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);

            // Create only the managed display OMath and its VTOMML/CustomXML
            // identity. The acceptance-only helper below then materializes the exact
            // Shape/TextBox + hidden-caption host used by the retired producer.
            service.InsertOmml(
                session,
                QuadraticFormulaMathMl(),
                deferNumberingLayout: true);
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "Legacy Shape migration fixture lost its OMML metadata before setup.");
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(
                    "Legacy Shape migration fixture lost its VTOMML bookmark before setup.");
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            WordEquationNumbering.CreateLegacyNumberedNativeOmmlShapeFixtureForAcceptance(
                document,
                formulaRange,
                metadata);
            Release(formulaRange); formulaRange = null;
            Release(formulaBookmark); formulaBookmark = null;

            externalReference = InsertExternalEquationReference(document, formulaId);
            externalReference.Update();
            AssertEqual("1", NormalizeLegacyShapeMigrationNumber(externalReference.Result.Text),
                "Legacy Shape fixture body REF did not read its hidden SEQ target.");
            Release(externalReference); externalReference = null;

            AssertLegacyNumberedOmmlShapeFixture(document, formulaId, "before save/reopen");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            AssertLegacyNumberedOmmlShapeFixture(document, formulaId, "after save/reopen");
            var legacyParagraphCount = document.Paragraphs.Count;

            // The OpenXML health probe must classify a structurally valid
            // VTEqShape_ host as migration input, not as a healthy fast-path host.
            // A single normal update therefore performs the one-time conversion.
            var updated = WordEquationNumbering.UpdateEquationNumbers(document);
            AssertEqual(1, updated,
                "The legacy VTEqShape_ document did not reconcile exactly one numbered formula.");
            FinalizeNumberedOmmlShapesAcrossOfficeTurns(
                document,
                expectedFormulaCount: 1,
                context: "legacy VTEqShape_ to native #SEQ migration");
            AssertNumberedOmmlTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "legacy VTEqShape_ after one-time migration");

            AssertEqual(0, document.Shapes.Count,
                "Legacy VTEqShape_ migration left a floating Shape behind.");
            AssertEqual(1, document.Tables.Count,
                "Legacy VTEqShape_ migration did not converge to exactly one minimal 1x3 numbering table.");
            AssertEqual(0, document.Frames.Count,
                "Legacy VTEqShape_ migration left the hidden caption Frame behind.");
            var normalizedId = Guid.Parse(formulaId).ToString("N");
            AssertTrue(!document.Bookmarks.Exists("VTEqAnc_" + normalizedId),
                "Legacy VTEqShape_ migration left its anchor bookmark behind.");
            AssertTrue(!document.Bookmarks.Exists("VTAncR_" + normalizedId),
                "Legacy VTEqShape_ migration left its anchor commit marker behind.");
            content = document.Content;
            var migratedXml = content.WordOpenXML ?? string.Empty;
            AssertTrue(migratedXml.IndexOf(
                    "VTEqShape_" + normalizedId,
                    StringComparison.OrdinalIgnoreCase) < 0,
                "Legacy VTEqShape_ identity remained in document OpenXML.");
            AssertTrue(migratedXml.IndexOf(
                    "VisualTeX numbered OMML " + normalizedId,
                    StringComparison.OrdinalIgnoreCase) < 0,
                "Legacy numbered-OMML Shape alternative text remained in OpenXML.");
            Release(content); content = null;
            Console.WriteLine(
                $"  legacy Shape migration paragraph count: before={legacyParagraphCount}, after={document.Paragraphs.Count}.");
            TraceOmmlOleRoundtripParagraphs(
                document,
                "legacy Shape after native #SEQ migration");
            AssertTrue(document.Paragraphs.Count != legacyParagraphCount,
                "Legacy Shape migration did not structurally replace the dedicated anchor/hidden-caption host.");

            externalReference = FindExternalEquationReference(document, formulaId)
                ?? throw new InvalidDataException(
                    "Legacy Shape migration lost the ordinary body REF field.");
            externalReference.Update();
            AssertEqual("1", NormalizeLegacyShapeMigrationNumber(externalReference.Result.Text),
                "Body REF did not follow VTEqNum after legacy Shape migration.");
            Release(externalReference); externalReference = null;

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            AssertNumberedOmmlTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "legacy VTEqShape_ migration save/reopen");
            AssertEqual(0, document.Shapes.Count,
                "Save/reopen recreated the retired VTEqShape_ host.");
            AssertEqual(1, document.Tables.Count,
                "Save/reopen lost or duplicated the minimal 1x3 numbering table after Shape migration.");
            AssertTrue(!document.Bookmarks.Exists("VTEqAnc_" + normalizedId),
                "Save/reopen recreated the retired Shape anchor bookmark.");
            externalReference = FindExternalEquationReference(document, formulaId)
                ?? throw new InvalidDataException(
                    "Save/reopen lost the body REF preserved by Shape migration.");
            externalReference.Update();
            AssertEqual("1", NormalizeLegacyShapeMigrationNumber(externalReference.Result.Text),
                "Save/reopen disconnected the body REF from the migrated #SEQ number.");

            Console.WriteLine(
                "Word legacy numbered-OMML Shape migration acceptance passed: a real VTEqShape_/TextBox + VTEqAnc_ + hidden SEQ caption fixture was classified as legacy, converted once to the minimal direct-SEQ 1x3 host, and left zero Shape/Frame artifacts while preserving FormulaId, VTEqNum, body REF and save/reopen.");
        }
        finally
        {
            Release(content);
            Release(externalReference);
            Release(formulaRange);
            Release(formulaBookmark);
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

    private static void AssertLegacyNumberedOmmlShapeFixture(
        Word.Document document,
        string formulaId,
        string context)
    {
        var normalizedId = Guid.Parse(formulaId).ToString("N");
        Word.Range? content = null;
        try
        {
            AssertEqual(1, document.Shapes.Count,
                context + ": the acceptance fixture does not contain exactly one legacy floating Shape.");
            AssertEqual(0, document.Tables.Count,
                context + ": the legacy Shape fixture unexpectedly contains a numbering table.");
            AssertTrue(document.Bookmarks.Exists("VTEqAnc_" + normalizedId),
                context + ": the legacy Shape anchor bookmark is missing.");
            AssertTrue(document.Bookmarks.Exists(
                    WordEquationNumbering.NativeCaptionBookmarkName(formulaId)),
                context + ": the legacy hidden caption bookmark is missing.");
            AssertTrue(document.Bookmarks.Exists(
                    WordEquationNumbering.NativeNumberBookmarkName(formulaId)),
                context + ": the legacy hidden number bookmark is missing.");
            AssertTrue(document.Bookmarks.Exists(
                    WordEquationNumbering.EquationBookmarkName(formulaId)),
                context + ": the legacy visible-number bookmark is missing.");
            content = document.Content;
            var xml = content.WordOpenXML ?? string.Empty;
            AssertTrue(xml.IndexOf(
                    "VTEqShape_" + normalizedId,
                    StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": the deterministic legacy Shape name is absent from OpenXML.");
            AssertTrue(xml.IndexOf("<w:txbxContent", StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": the legacy visible REF is not serialized in a TextBox story.");
            AssertTrue(xml.IndexOf(
                    "REF " + WordEquationNumbering.NativeNumberBookmarkName(formulaId),
                    StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": the legacy TextBox does not REF its hidden number target.");
        }
        finally { Release(content); }
    }

    private static string NormalizeLegacyShapeMigrationNumber(string? value) =>
        (value ?? string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace("\a", string.Empty)
            .Trim()
            .Trim('(', ')');
}
