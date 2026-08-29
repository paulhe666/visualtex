using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMixedLegacyNumberingMigrationAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-mixed-legacy-numbering-migration.docx");
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"mixed-legacy-numbering-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);
        var svgPath = Path.Combine(assetRoot, "mixed-legacy-numbering.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"76\" viewBox=\"0 0 240 76\"><text x=\"4\" y=\"50\" font-size=\"34\">x = 1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 76);
        var pngDataUrl = CreatePngDataUrl("mixed-legacy-numbering", 240, 76);
        var pngPath = Path.Combine(assetRoot, "mixed-legacy-numbering.png");
        File.WriteAllBytes(
            pngPath,
            Convert.FromBase64String(
                pngDataUrl.Substring(pngDataUrl.IndexOf(',') + 1)));

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Field? oleExternalReference = null;
        Word.Field? tableOmmlExternalReference = null;
        Word.Field? badEqArrExternalReference = null;
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

            var oleFormulaId = Guid.NewGuid().ToString("D");
            InsertHealthyMixedLegacyOle(
                application,
                document,
                service,
                oleFormulaId,
                pngPath,
                emfPath);
            oleExternalReference = InsertExternalEquationReference(document, oleFormulaId);

            var tableOmmlFormulaId = Guid.NewGuid().ToString("D");
            InsertHealthyMixedLegacyOmml(
                application,
                document,
                service,
                tableOmmlFormulaId,
                @"u=\frac{p+q}{r}",
                "u");
            tableOmmlExternalReference = InsertExternalEquationReference(
                document,
                tableOmmlFormulaId);

            var badEqArrFormulaId = Guid.NewGuid().ToString("D");
            InsertHealthyMixedLegacyOmml(
                application,
                document,
                service,
                badEqArrFormulaId,
                @"v=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                "v");
            badEqArrExternalReference = InsertExternalEquationReference(
                document,
                badEqArrFormulaId);

            // Degrade from the end of the document toward the start so the fixture
            // preserves every production-created FormulaId/SEQ/REF while materializing
            // the three historical hosts. These table-creation calls exist only in
            // acceptance code and are never reachable from production insertion.
            DegradeNumberedOmmlToBadEqArr(
                application,
                document,
                badEqArrFormulaId);
            DegradeNumberedOmmlToLegacyNumberingTable(
                document,
                tableOmmlFormulaId);
            DegradeVisualTeXOleToLegacyNumberingTable(document, oleFormulaId);

            AssertEqual(2, document.Tables.Count,
                "The mixed legacy fixture does not contain exactly two historical numbering tables before migration.");
            AssertTrue(
                WordEquationNumbering.NeedsLegacyManagedNumberingMigration(document),
                "The document-level legacy-numbering pre-pass did not recognize the mixed historical fixture.");

            var migrated = ThisAddIn.MigrateManagedOmmlDisplayTypes(document);
            AssertTrue(migrated > 0,
                "One document-level migration pass reported no repaired formulas for the mixed historical fixture.");
            FinalizeNumberedOmmlShapesAcrossOfficeTurns(
                document,
                expectedFormulaCount: 2,
                context: "mixed legacy numbered OMML native #SEQ finalization");

            AssertEqual(0, document.Tables.Count,
                "One document-level migration pass left a legacy numbering table behind.");
            AssertVisualTeXNumberedTabHost(
                document,
                oleFormulaId,
                updateReference: true,
                context: "mixed migration VisualTeX OLE");
            AssertOmmlTabNumberingHost(
                document,
                tableOmmlFormulaId,
                context: "mixed migration legacy-table OMML",
                updateReference: true);
            AssertOmmlTabNumberingHost(
                document,
                badEqArrFormulaId,
                context: "mixed migration bad-eqArr OMML",
                updateReference: true);
            AssertExternalEquationReference(
                document,
                oleFormulaId,
                "mixed migration VisualTeX OLE external REF");
            AssertExternalEquationReference(
                document,
                tableOmmlFormulaId,
                "mixed migration legacy-table OMML external REF");
            AssertExternalEquationReference(
                document,
                badEqArrFormulaId,
                "mixed migration bad-eqArr OMML external REF");
            AssertTrue(
                !WordEquationNumbering.NeedsLegacyManagedNumberingMigration(document),
                "The mixed document is still classified as a legacy numbering document after one migration pass.");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();

            AssertEqual(0, document.Tables.Count,
                "The mixed migrated document regained legacy numbering tables after save/reopen.");
            AssertVisualTeXNumberedTabHost(
                document,
                oleFormulaId,
                updateReference: true,
                context: "mixed migration VisualTeX OLE save/reopen");
            AssertOmmlTabNumberingHost(
                document,
                tableOmmlFormulaId,
                context: "mixed migration legacy-table OMML save/reopen",
                updateReference: true);
            AssertOmmlTabNumberingHost(
                document,
                badEqArrFormulaId,
                context: "mixed migration bad-eqArr OMML save/reopen",
                updateReference: true);
            AssertExternalEquationReference(
                document,
                oleFormulaId,
                "mixed migration VisualTeX OLE external REF save/reopen");
            AssertExternalEquationReference(
                document,
                tableOmmlFormulaId,
                "mixed migration legacy-table OMML external REF save/reopen");
            AssertExternalEquationReference(
                document,
                badEqArrFormulaId,
                "mixed migration bad-eqArr OMML external REF save/reopen");

            Console.WriteLine(
                "Word mixed legacy numbering migration acceptance passed: one document-level migration upgraded a legacy VisualTeX OLE table, a legacy OMML table and a broken m:eqArr/# formula to the final table-free hosts while preserving FormulaIds, dynamic SEQ/REF targets, external references and save/reopen stability.");
        }
        finally
        {
            Release(badEqArrExternalReference);
            Release(tableOmmlExternalReference);
            Release(oleExternalReference);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            try { Directory.Delete(assetRoot, recursive: true); } catch { }
            ForceComCleanup();
        }
    }

    private static void InsertHealthyMixedLegacyOle(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        string formulaId,
        string pngPath,
        string emfPath)
    {
        Word.Range? insertion = null;
        try
        {
            insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            application.Selection.SetRange(insertion.Start, insertion.End);
            var session = CreateNumberedPerformanceSession(
                "create",
                formulaId,
                document.FullName,
                WordRangeReference(insertion.Start, insertion.End),
                originalMetadata: null,
                latex: @"x=1");
            session.ExportResult = new OfficeExportDocument
            {
                Width = 240,
                Height = 76,
                Baseline = 57,
                FormulaLetterFont = "katex",
                FormulaChineseFont = "system",
            };
            service.InsertOle(session, pngPath, emfPath);
            AssertVisualTeXNumberedTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "mixed fixture healthy VisualTeX OLE");
        }
        finally { Release(insertion); }
    }

    private static void InsertHealthyMixedLegacyOmml(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        string formulaId,
        string latex,
        string variable)
    {
        Word.Range? insertion = null;
        try
        {
            insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            insertion.InsertParagraphAfter();
            Release(insertion);
            insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            application.Selection.SetRange(insertion.Start, insertion.End);
            var session = CreateOmmlNumberingSession(
                document,
                formulaId,
                mode: "create",
                sourceObjectId: WordRangeReference(insertion.Start, insertion.End),
                latex,
                originalMetadata: null);
            service.InsertOmml(session, QuadraticMathMl(variable));
            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                context: "mixed fixture healthy numbered OMML " + variable,
                updateReference: true,
                requireDocumentTableFree: false);
        }
        finally { Release(insertion); }
    }

    private static void DegradeVisualTeXOleToLegacyNumberingTable(
        Word.Document document,
        string formulaId)
    {
        Word.Range? owner = null;
        Word.Table? table = null;
        Word.Row? emptyRow = null;
        try
        {
            owner = WordEquationNumbering.FindNumberingOwnerRange(document, formulaId)
                ?? throw new InvalidDataException(
                    "The mixed fixture could not resolve the healthy VisualTeX OLE numbering paragraph before legacy degradation.");
            object separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
            object rows = 1;
            object columns = 3;
            table = owner.ConvertToTable(
                ref separator,
                ref rows,
                ref columns);
            AssertEqual(1, table.Rows.Count,
                "Converting the healthy OLE tab host did not create the expected historical 1x3 table.");
            AssertEqual(3, table.Columns.Count,
                "Converting the healthy OLE tab host did not create exactly three historical columns.");
            emptyRow = table.Rows.Add();
            AssertEqual(2, table.Rows.Count,
                "The mixed fixture did not create a benign empty second OLE numbering row.");
        }
        finally
        {
            Release(emptyRow);
            Release(table);
            Release(owner);
        }
    }

    private static void DegradeNumberedOmmlToLegacyNumberingTable(
        Word.Document document,
        string formulaId)
    {
        Word.Application? application = null;
        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
        Word.Range? semanticRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Range? before = null;
        Word.Range? after = null;
        Word.Table? table = null;
        Word.Row? emptyRow = null;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "The mixed OMML table fixture lost metadata before legacy degradation.");
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(
                    "The mixed OMML table fixture lost its formula bookmark before legacy degradation.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var semanticOmml = WordOmmlConverter.StripVisualTeXNativeEquationNumber(
                equationRange.WordOpenXML);

            // Current numbered OMML owns a live mathematical SEQ field, so a legacy
            // fixture cannot be manufactured by writing TABs into that OMath. First
            // replace the complete #(SEQ) equation atomically with its pure semantic
            // OMath, then build the historical 1x3/2x3 table around that plain range.
            WordEquationNumbering.RemoveNativeOmmlHashSequenceAliasesForReplacement(
                document,
                formulaId);
            bookmark.Delete();
            Release(bookmark);
            bookmark = null;
            application = document.Application;
            string mathFontName;
            try { mathFontName = document.OMathFontName ?? string.Empty; }
            catch { mathFontName = string.Empty; }
            semanticRange = WordOmmlConverter.ReplaceWithPreparedOmml(
                application,
                document,
                equationRange,
                semanticOmml,
                display: false,
                mathFontName);
            Release(equationRange);
            equationRange = semanticRange;
            semanticRange = null;

            maths = equationRange.OMaths;
            AssertEqual(1, maths.Count,
                "The mixed OMML table fixture does not contain exactly one semantic OMath.");
            math = maths[1];
            if (math.Type != Word.WdOMathType.wdOMathInline)
                math.Type = Word.WdOMathType.wdOMathInline;
            Release(equationRange); equationRange = math.Range.Duplicate;

            before = document.Range(equationRange.Start, equationRange.Start);
            before.Text = "\t";
            Release(equationRange); equationRange = math.Range.Duplicate;
            after = document.Range(equationRange.End, equationRange.End);
            after.Text = "\t";

            paragraphs = equationRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            object separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
            object rows = 1;
            object columns = 3;
            table = paragraphRange.ConvertToTable(
                ref separator,
                ref rows,
                ref columns);
            AssertEqual(1, table.Rows.Count,
                "Converting the numbered OMML paragraph did not create the expected historical 1x3 table.");
            AssertEqual(3, table.Columns.Count,
                "Converting the numbered OMML paragraph did not create exactly three historical columns.");
            emptyRow = table.Rows.Add();

            var center = table.Cell(1, 2);
            Word.Range? centerRange = null;
            Word.OMaths? centerMaths = null;
            Word.OMath? centerMath = null;
            Word.Range? centerMathRange = null;
            Word.Bookmark? repaired = null;
            try
            {
                centerRange = center.Range;
                centerMaths = centerRange.OMaths;
                AssertEqual(1, centerMaths.Count,
                    "The historical OMML table lost its native equation while being materialized.");
                centerMath = centerMaths[1];
                centerMathRange = centerMath.Range;
                WordOmmlNativeSource.StampFingerprintFromResolvedRange(metadata, centerMathRange);
                repaired = WordOmmlFormulaStore.Wrap(
                    document,
                    centerMathRange,
                    metadata,
                    replaceExisting: true);
                WordOmmlFormulaStore.Save(document, metadata);
            }
            finally
            {
                Release(repaired);
                Release(centerMathRange);
                Release(centerMath);
                Release(centerMaths);
                Release(centerRange);
                Release(center);
            }
        }
        finally
        {
            Release(emptyRow);
            Release(table);
            Release(after);
            Release(before);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(math);
            Release(maths);
            Release(semanticRange);
            Release(equationRange);
            Release(bookmark);
            Release(application);
        }
    }

    private static void DegradeNumberedOmmlToBadEqArr(
        Word.Application application,
        Word.Document document,
        string formulaId)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
        Word.Range? badRange = null;
        Word.Bookmark? repaired = null;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "The mixed bad-eqArr fixture lost metadata before legacy degradation.");
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(
                    "The mixed bad-eqArr fixture lost its formula bookmark before legacy degradation.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var semanticOmml = WordOmmlConverter.ExtractSingleOMath(
                equationRange.WordOpenXML);
            WordEquationNumbering.RemoveVisibleEquationNumberForFormula(
                document,
                formulaId);
            var badOmml = BuildLegacyBadEqArrOmml(
                semanticOmml,
                formulaId);
            string mathFontName;
            try { mathFontName = document.OMathFontName ?? string.Empty; }
            catch { mathFontName = string.Empty; }
            badRange = WordOmmlConverter.ReplaceWithPreparedOmml(
                application,
                document,
                equationRange,
                badOmml,
                display: true,
                mathFontName);
            repaired = WordOmmlFormulaStore.Wrap(
                document,
                badRange,
                metadata,
                replaceExisting: true);
            WordOmmlNativeSource.StampFingerprintFromResolvedRange(metadata, badRange);
            WordOmmlFormulaStore.Save(document, metadata);
            AssertTrue(
                WordOmmlConverter.HasVisualTeXNativeEquationNumber(badRange.WordOpenXML),
                "The mixed fixture did not materialize the historical m:eqArr/# numbering wrapper.");
        }
        finally
        {
            Release(repaired);
            Release(badRange);
            Release(equationRange);
            Release(bookmark);
        }
    }
}
