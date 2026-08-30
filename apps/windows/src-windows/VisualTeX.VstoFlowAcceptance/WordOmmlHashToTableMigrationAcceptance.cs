using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlHashToTableMigrationAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The #(SEQ)->1x3 migration acceptance refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-hash-to-1x3-migration.docx");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Bookmark? formulaBookmark = null;
        Word.Range? equationRange = null;
        Word.Range? hashRange = null;
        Word.Bookmark? repairedBookmark = null;
        Word.Field? externalReference = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            ConfigureOmmlTableNumberPage(document);
            document.SaveAs2(
                documentPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            // First create a completely ordinary managed display OMML. We then use
            // the retired production builder to turn that exact semantic equation
            // into the previous m:eqArr + #(SEQ VisualTeXEquation) host.
            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var insertion = document.Content.End - 1;
            application.Selection.SetRange(insertion, insertion);
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertion,
                insertion,
                latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);
            session.Numbered = false;
            service.InsertOmml(session, QuadraticFormulaMathMl());

            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException("Hash-migration fixture lost VisualTeX metadata.");
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException("Hash-migration fixture lost VTOMML identity.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            var semanticOmml = equationRange.WordOpenXML ?? string.Empty;
            var retiredHashOmml = WordOmmlConverter.BuildImmutableHashSequenceNumberedOmml(
                semanticOmml,
                "VisualTeXEquation",
                WordEquationNumbering.NativeNumberBookmarkName(formulaId),
                WordEquationNumbering.EquationBookmarkName(formulaId),
                WordEquationNumbering.NativeCaptionBookmarkName(formulaId),
                prefix: string.Empty,
                restartHeadingLevel: 0,
                initialSequenceResult: "1");

            Release(formulaBookmark); formulaBookmark = null;
            application = document.Application;
            hashRange = WordOmmlConverter.ReplaceWithPreparedOmml(
                application,
                document,
                equationRange,
                retiredHashOmml,
                display: true,
                mathFontName: document.OMathFontName);
            Release(equationRange); equationRange = null;
            equationRange = hashRange.Duplicate;

            metadata.Numbered = true;
            metadata.DisplayMode = "block";
            repairedBookmark = WordOmmlFormulaStore.Wrap(
                document,
                equationRange,
                metadata,
                replaceExisting: true);
            WordOmmlNativeSource.StampFingerprintFromResolvedRange(metadata, equationRange);
            WordOmmlFormulaStore.Save(document, metadata);

            AssertRetiredHashHostBeforeMigration(document, formulaId, equationRange);
            externalReference = InsertExternalEquationReference(document, formulaId);
            externalReference.Update();
            AssertEqual(
                "1",
                NormalizeEquationNumberText(externalReference.Result.Text),
                "The body REF could not read VTEqNum from the retired #(SEQ) fixture.");

            // This is the actual production migration under test. The old sequence
            // and eqArr must be stripped atomically before the semantic OMath is
            // placed in cell (1,2); only a new ordinary Word SEQ may survive in
            // cell (1,3).
            WordEquationNumbering.ReconcileFormula(
                document,
                equationRange,
                WordOmmlFormulaStore.EstimateHeightPoints(equationRange),
                metadata,
                numberingOrderMayHaveChanged: true);
            Release(hashRange); hashRange = null;
            Release(equationRange); equationRange = null;
            Release(repairedBookmark); repairedBookmark = null;

            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "01-after-hash-to-1x3-migration");
            AssertNoRetiredHashNumberInsideCenterMath(document, formulaId, "after migration");
            externalReference.Update();
            AssertEqual(
                ReadVisibleEquationNumber(document, formulaId).Trim('(', ')'),
                NormalizeEquationNumberText(externalReference.Result.Text),
                "The external REF disconnected while migrating #(SEQ) to 1x3.");

            document.Fields.Update();
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "02-after-migration-f9");
            AssertNoRetiredHashNumberInsideCenterMath(document, formulaId, "after F9");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            Release(externalReference); externalReference = null;
            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "03-after-migration-reopen");
            AssertNoRetiredHashNumberInsideCenterMath(document, formulaId, "after save/reopen");
            externalReference = FindExternalEquationReference(document, formulaId)
                ?? throw new InvalidDataException("Save/reopen lost the body REF after hash migration.");
            externalReference.Update();
            AssertEqual(
                ReadVisibleEquationNumber(document, formulaId).Trim('(', ')'),
                NormalizeEquationNumberText(externalReference.Result.Text),
                "Save/reopen changed the external REF after hash migration.");

            Console.WriteLine(
                "Word retired #(SEQ)->1x3 migration acceptance passed: the old m:eqArr/hash sequence field and all mathematical number aliases were removed from m:oMath, exactly one ordinary SEQ remained in right cell (1,3), the center stayed genuine wdOMathDisplay, and body REF/F9/save-reopen all survived without duplicate numbering.");
        }
        finally
        {
            Release(externalReference);
            Release(repairedBookmark);
            Release(hashRange);
            Release(equationRange);
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

    private static void AssertRetiredHashHostBeforeMigration(
        Word.Document document,
        string formulaId,
        Word.Range equationRange)
    {
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            AssertEqual(0, document.Tables.Count,
                "The retired #(SEQ) fixture unexpectedly started in a table.");
            maths = equationRange.OMaths;
            AssertEqual(1, maths.Count,
                "The retired #(SEQ) fixture does not contain exactly one OMath.");
            math = maths[1];
            AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
                "The retired #(SEQ) fixture is not wdOMathDisplay.");
            fields = equationRange.Fields;
            AssertEqual(1, fields.Count,
                "The retired #(SEQ) fixture must contain exactly one mathematical SEQ field.");
            field = fields[1];
            code = field.Code;
            AssertTrue(WordEquationNumbering.IsVisualTeXSequenceFieldCode(code.Text),
                "The retired #(SEQ) fixture mathematical field is not SEQ VisualTeXEquation.");
            AssertTrue(WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                    equationRange.WordOpenXML,
                    formulaId),
                "The retired fixture is missing the managed #(SEQ) wrapper/aliases.");
            Console.WriteLine(
                $"  retired #(SEQ) fixture: range={equationRange.Start}:{equationRange.End}, fieldsInMath={fields.Count}, tables={document.Tables.Count}.");
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(math);
            Release(maths);
        }
    }

    private static void AssertNoRetiredHashNumberInsideCenterMath(
        Word.Document document,
        string formulaId,
        string context)
    {
        Word.Range? equationRange = null;
        Word.Tables? tables = null;
        Word.Table? table = null;
        Word.Cell? numberCell = null;
        Word.Range? numberCellRange = null;
        Word.Fields? numberFields = null;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(context + ": metadata is missing.");
            equationRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                formulaId,
                metadata);
            AssertEqual(0, equationRange.Fields.Count,
                context + ": a retired mathematical SEQ still remains inside m:oMath.");
            AssertTrue(!WordOmmlConverter.HasVisualTeXNativeEquationNumber(
                    equationRange.WordOpenXML ?? string.Empty),
                context + ": the retired m:eqArr/#(...) wrapper still remains in m:oMath.");
            AssertTrue((equationRange.WordOpenXML ?? string.Empty).IndexOf(
                    "VTEqNum_",
                    StringComparison.OrdinalIgnoreCase) < 0,
                context + ": a VTEq number alias still remains inside the center mathematical XML.");

            tables = equationRange.Tables;
            AssertEqual(1, tables.Count,
                context + ": the migrated equation no longer owns exactly one 1x3 table.");
            table = tables[1];
            numberCell = table.Cell(1, 3);
            numberCellRange = numberCell.Range;
            numberFields = numberCellRange.Fields;
            AssertEqual(1, numberFields.Count,
                context + ": the right number cell does not own exactly one replacement SEQ field.");
            AssertEqual(1, document.Fields.Cast<Word.Field>().Count(field =>
            {
                Word.Range? fieldCode = null;
                try
                {
                    fieldCode = field.Code;
                    return WordEquationNumbering.IsVisualTeXSequenceFieldCode(fieldCode.Text);
                }
                finally { Release(fieldCode); }
            }), context + ": more than one VisualTeXEquation SEQ survives after migration.");
        }
        finally
        {
            Release(numberFields);
            Release(numberCellRange);
            Release(numberCell);
            Release(table);
            Release(tables);
            Release(equationRange);
        }
    }
}
