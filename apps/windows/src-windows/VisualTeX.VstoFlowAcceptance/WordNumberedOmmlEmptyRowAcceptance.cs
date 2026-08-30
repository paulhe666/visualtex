using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordNumberedOmmlEmptyRowAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-numbered-omml-empty-row.docx");
        const string latex = @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}";
        const string initialMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo>"
            + "<msqrt><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></msqrt>"
            + "</mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>";
        const string editedMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo>"
            + "<msqrt><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></msqrt>"
            + "</mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac><mo>+</mo><mn>0</mn></math>";

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.Range? sourceRange = null;
        Word.Table? legacyTable = null;
        Word.Row? emptyRow = null;
        Word.Table? repairedTable = null;
        Word.Rows? repairedRows = null;
        Word.Columns? repairedColumns = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);

            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var insertionPosition = Math.Max(
                document.Content.Start,
                document.Content.End - 1);
            insertion = document.Range(insertionPosition, insertionPosition);
            application.Selection.SetRange(insertion.Start, insertion.End);
            var createSession = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertion.Start,
                insertion.End,
                latex,
                originalMetadata: null);
            service.InsertOmml(createSession, initialMathMl);
            Release(insertion);
            insertion = null;

            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "empty-row fixture healthy 1x3 source");
            AssertEqual(1, document.Tables.Count,
                "The empty-row fixture did not start with exactly one current 1x3 table.");

            var sourceMetadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "The numbered OMML source lost its metadata before the empty-row fixture was created.");
            sourceRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                formulaId,
                sourceMetadata);
            legacyTable = WordEquationNumbering.FindNumberedEquationTable(
                    document,
                    formulaId)
                ?? throw new InvalidDataException(
                    "The empty-row regression could not resolve the current managed 1x3 table.");
            emptyRow = legacyTable.Rows.Add();
            AssertEqual(2, legacyTable.Rows.Count,
                "The empty-row regression fixture did not produce a 2x3 managed table.");
            AssertEqual(3, legacyTable.Columns.Count,
                "The empty-row regression fixture changed the managed table's column count.");
            AssertEqual(1, document.Tables.Count,
                "Adding one benign row duplicated the numbered OMML table.");
            AssertEqual(1, document.OMaths.Count,
                "Adding one benign row changed the numbered OMML formula count.");

            // Re-resolve the formula after Rows.Add, because Word can shift every
            // range endpoint in the table even though the center OMath is unchanged.
            Release(sourceRange);
            sourceRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                formulaId,
                sourceMetadata);
            var replaceSession = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                sourceRange.Start,
                sourceRange.End,
                latex,
                sourceMetadata);

            // The replacement owns the structural migration. Release every RCW that
            // points into the malformed 2x3 table before the transaction mutates it.
            Release(emptyRow);
            emptyRow = null;
            Release(legacyTable);
            legacyTable = null;
            Release(sourceRange);
            sourceRange = null;

            service.ReplaceOmml(replaceSession, editedMathMl);

            AssertEqual(1, document.Tables.Count,
                "Editing a benign 2x3 numbered OMML host did not converge to one table.");
            AssertEqual(1, document.OMaths.Count,
                "Editing the benign 2x3 host changed the formula count.");
            AssertEqual(0, document.Shapes.Count,
                "Editing the benign 2x3 host recreated a retired Shape/TextBox number.");
            AssertEqual(0, document.Frames.Count,
                "Editing the benign 2x3 host recreated a retired caption Frame.");
            repairedTable = WordEquationNumbering.FindNumberedEquationTable(
                    document,
                    formulaId)
                ?? throw new InvalidDataException(
                    "Editing the benign 2x3 host lost the current numbered table.");
            repairedRows = repairedTable.Rows;
            repairedColumns = repairedTable.Columns;
            AssertEqual(1, repairedRows.Count,
                "Editing the benign 2x3 host left an extra table row.");
            AssertEqual(3, repairedColumns.Count,
                "Editing the benign 2x3 host changed the table column count.");
            Release(repairedColumns);
            repairedColumns = null;
            Release(repairedRows);
            repairedRows = null;
            Release(repairedTable);
            repairedTable = null;
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "benign 2x3 host after OMML edit repair");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            AssertEqual(1, document.Tables.Count,
                "Save/reopen after empty-row repair changed the table count.");
            AssertEqual(1, document.OMaths.Count,
                "Save/reopen after empty-row repair changed the formula count.");
            AssertEqual(0, document.Shapes.Count,
                "Save/reopen after empty-row repair recreated a Shape/TextBox.");
            AssertEqual(0, document.Frames.Count,
                "Save/reopen after empty-row repair recreated a caption Frame.");
            repairedTable = WordEquationNumbering.FindNumberedEquationTable(
                    document,
                    formulaId)
                ?? throw new InvalidDataException(
                    "Save/reopen after empty-row repair lost the 1x3 table.");
            repairedRows = repairedTable.Rows;
            repairedColumns = repairedTable.Columns;
            AssertEqual(1, repairedRows.Count,
                "Save/reopen after empty-row repair restored the extra row.");
            AssertEqual(3, repairedColumns.Count,
                "Save/reopen after empty-row repair changed the column count.");
            Release(repairedColumns);
            repairedColumns = null;
            Release(repairedRows);
            repairedRows = null;
            Release(repairedTable);
            repairedTable = null;
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "benign 2x3 repair save/reopen");

            Console.WriteLine(
                "Numbered OMML empty-row acceptance passed: a current direct-SEQ 1x3 host was deliberately expanded to a benign 2x3 fixture, the next VisualTeX edit converged it back to exactly one 1x3 table without Shape/Frame artifacts, and save/reopen preserved the repaired structure.");
        }
        finally
        {
            Release(repairedColumns);
            Release(repairedRows);
            Release(repairedTable);
            Release(emptyRow);
            Release(legacyTable);
            Release(sourceRange);
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
}
