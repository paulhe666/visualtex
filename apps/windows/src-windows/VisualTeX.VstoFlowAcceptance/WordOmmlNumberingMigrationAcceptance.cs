using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlNumberingMigrationAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-numbering-table-migration.docx");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? tableInsertion = null;
        Word.Table? legacyTable = null;
        Word.Cell? centerCell = null;
        Word.Range? centerRange = null;
        Word.Range? equationRange = null;
        Word.Row? emptyRow = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);

            tableInsertion = document.Range(0, 0);
            // Construct the historical benign 2x3 host up front. Current numbering
            // reconciliation migrates native OMML tables immediately, so adding an
            // empty row after calling the scaffold builder would access a Table RCW
            // that Word has already converted to text.
            legacyTable = document.Tables.Add(tableInsertion, 2, 3);
            centerCell = legacyTable.Cell(1, 2);
            centerRange = centerCell.Range.Duplicate;
            centerRange.End = Math.Max(centerRange.Start, centerRange.End - 1);
            application.Selection.SetRange(centerRange.Start, centerRange.End);

            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var session = CreateOmmlNumberingSession(
                document,
                formulaId,
                mode: "create",
                sourceObjectId: WordRangeReference(centerRange.Start, centerRange.End),
                latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);
            service.InsertOmml(
                session,
                QuadraticMathMl("x"),
                deferNumberingLayout: true);

            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "The legacy OMML migration fixture lost its metadata.");
            equationRange = ResolveOmmlRange(document, formulaId, metadata);
            AssertTrue(
                (bool)equationRange.get_Information(Word.WdInformation.wdWithInTable),
                "The legacy OMML migration fixture did not place the OMath inside its table.");

            AssertEqual(2, legacyTable.Rows.Count,
                "The legacy OMML migration fixture did not start as a benign 2x3 table.");
            AssertEqual(3, legacyTable.Columns.Count,
                "The legacy OMML migration fixture changed the table column count.");

            WordEquationNumbering.ReconcileFormula(
                document,
                equationRange,
                WordOmmlFormulaStore.EstimateHeightPoints(equationRange),
                metadata,
                numberingOrderMayHaveChanged: true,
                reuseExistingNumberedTableFormatting: true,
                knownNumberedTable: legacyTable);
            Release(equationRange); equationRange = null;
            Release(emptyRow); emptyRow = null;
            Release(centerRange); centerRange = null;
            Release(centerCell); centerCell = null;
            Release(legacyTable); legacyTable = null;

            FinalizeNumberedOmmlShapesAcrossOfficeTurns(
                document,
                expectedFormulaCount: 1,
                context: "legacy 2x3 OMML native #SEQ finalization");
            TraceOmmlMigrationIdentity(document, formulaId);
            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                context: "legacy 2x3 OMML first reconcile",
                updateReference: true);
            AssertEqual(0, document.Tables.Count,
                "Legacy OMML migration left a numbering table in the document.");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                context: "legacy OMML migration save/reopen",
                updateReference: true);
            AssertEqual(0, document.Tables.Count,
                "Saved/reopened migrated OMML unexpectedly regained a numbering table.");

            Console.WriteLine(
                "Word OMML numbering migration acceptance passed: one reconcile migrated the benign legacy 2x3 host to a pure wdOMathDisplay/m:oMathPara formula with Word-native #(SEQ), preserved FormulaId/VTEqNum identity and external body REF behavior, and save/reopen remained Shape/Table-free.");
        }
        finally
        {
            Release(emptyRow);
            Release(equationRange);
            Release(centerRange);
            Release(centerCell);
            Release(legacyTable);
            Release(tableInsertion);
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

    private static void TraceOmmlMigrationIdentity(
        Word.Document document,
        string formulaId)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? bookmarkRange = null;
        Word.OMaths? maths = null;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            bookmarkRange = bookmark?.Range;
            Console.WriteLine(
                $"  [migration identity] bookmark={bookmarkRange?.Start}:{bookmarkRange?.End} storedFingerprint={metadata?.NativeOmmlFingerprint ?? "<null>"}.");
            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                Word.OMath? math = null;
                Word.Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    string fingerprint;
                    try
                    {
                        fingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                            range.WordOpenXML);
                    }
                    catch (Exception error)
                    {
                        fingerprint = "<error:" + error.Message + ">";
                    }
                    Console.WriteLine(
                        $"  [migration identity] OMath#{index} type={math.Type} range={range.Start}:{range.End} fingerprint={fingerprint}.");
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
        }
        finally
        {
            Release(maths);
            Release(bookmarkRange);
            Release(bookmark);
        }
    }
}
