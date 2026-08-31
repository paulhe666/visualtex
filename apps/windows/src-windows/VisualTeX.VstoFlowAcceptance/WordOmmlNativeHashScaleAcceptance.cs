using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlNativeHashScaleAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-native-hash-20-cross-page.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        var formulaIds = new List<string>();
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

            for (var index = 1; index <= 20; index++)
            {
                formulaIds.Add(InsertNativeHashProductionFormula(
                    application,
                    document,
                    service,
                    document.Content.End - 1,
                    @"x_{" + index + @"}=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                    QuadraticFormulaMathMl()));
                if (index < 20 && index % 5 == 0)
                    AppendNativeHashAcceptancePageBreak(document);
            }

            UpdateDirectTableScaleNumbers(service, document, formulaIds);
            var expected = Enumerable.Range(1, 20)
                .Select(value => value.ToString(
                    System.Globalization.CultureInfo.InvariantCulture))
                .ToArray();
            AssertDirectTableScaleNumbers(
                document,
                formulaIds,
                expected,
                "20-formula cross-page initial numbering");
            try { document.Repaginate(); } catch { }
            var pages = document.ComputeStatistics(
                Word.WdStatistic.wdStatisticPages,
                IncludeFootnotesAndEndnotes: false);
            AssertTrue(pages >= 4,
                $"20-formula acceptance did not span the forced four page regions; pages={pages}.");

            var referencedIds = new[]
            {
                formulaIds[0],
                formulaIds[9],
                formulaIds[19],
            };
            var references = InsertNativeHashProductionReferences(
                document,
                referencedIds);
            AssertNativeHashProductionReferences(
                document,
                references,
                referencedIds,
                new[] { "1", "10", "20" },
                "20-formula first/middle/last body REF");
            AssertDirectTableScaleNumbers(
                document,
                formulaIds,
                expected,
                "20-formula numbering after body REF insertion");
            Console.WriteLine(
                $"  direct-SEQ 1x3 scale stage passed: 20 formulas, pages={pages}, numbers 1..20, body REF 1/10/20");

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
            UpdateDirectTableScaleNumbers(service, document, formulaIds);
            AssertDirectTableScaleNumbers(
                document,
                formulaIds,
                expected,
                "20-formula cross-page save/reopen numbering");
            AssertNativeHashProductionReferences(
                document,
                references,
                referencedIds,
                new[] { "1", "10", "20" },
                "20-formula save/reopen body REF");
            AssertEqual(0, document.Shapes.Count,
                "20-formula save/reopen created a floating Shape.");
            AssertEqual(20, document.Tables.Count,
                "20-formula save/reopen did not retain one direct 1x3 host per formula.");
            AssertEqual(20, document.OMaths.Count,
                "20-formula save/reopen changed the OMML formula count.");

            Console.WriteLine(
                "Production OMML direct-SEQ 1x3 scale acceptance passed: 20 formulas remained 1..20 across forced page breaks and save/reopen, with body REF 1/10/20, no Shape/TextBox and exactly one 1x3 host per formula.");
        }
        finally
        {
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

    private static void UpdateDirectTableScaleNumbers(
        WordFormulaService service,
        Word.Document document,
        IReadOnlyList<string> formulaIds)
    {
        var updated = service.UpdateEquationNumbers();
        AssertTrue(
            updated >= formulaIds.Count,
            $"Direct-SEQ scale update returned {updated} formulas; expected at least {formulaIds.Count}.");
        Word.Fields? fields = null;
        try
        {
            fields = document.Fields;
            if (fields.Count > 0) fields.Update();
        }
        finally { Release(fields); }
    }

    private static void AssertDirectTableScaleNumbers(
        Word.Document document,
        IReadOnlyList<string> formulaIds,
        IReadOnlyList<string> expectedNumbers,
        string context)
    {
        AssertEqual(formulaIds.Count, expectedNumbers.Count,
            context + ": formula/expectation count mismatch.");
        AssertEqual(formulaIds.Count, document.Tables.Count,
            context + ": each numbered OMML must own exactly one 1x3 table.");
        AssertEqual(formulaIds.Count, document.OMaths.Count,
            context + ": OMML formula count changed.");
        AssertEqual(0, document.Shapes.Count,
            context + ": a floating Shape/TextBox exists.");

        Word.Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 0; index < formulaIds.Count; index++)
            {
                var formulaId = formulaIds[index];
                var numberName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
                AssertTrue(bookmarks.Exists(numberName),
                    context + $": {numberName} is missing.");
                Word.Bookmark? bookmark = null;
                Word.Range? range = null;
                Word.Fields? fields = null;
                Word.OMaths? maths = null;
                try
                {
                    bookmark = bookmarks[numberName];
                    range = bookmark.Range;
                    fields = range.Fields;
                    maths = range.OMaths;
                    AssertEqual(1, fields.Count,
                        context + $": {numberName} does not contain exactly one direct SEQ field.");
                    AssertEqual(0, maths.Count,
                        context + $": {numberName} leaked into the mathematical OMath.");
                    AssertEqual(
                        expectedNumbers[index],
                        NormalizeNativeHashProductionNumber(range.Text),
                        context + $": {numberName} result mismatch.");
                    AssertTrue(
                        (bool)range.get_Information(Word.WdInformation.wdWithInTable),
                        context + $": {numberName} is outside the right table cell.");
                }
                finally
                {
                    Release(maths);
                    Release(fields);
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }

        foreach (var representative in new[] { 0, formulaIds.Count / 2, formulaIds.Count - 1 })
        {
            AssertOmmlTabNumberingHost(
                document,
                formulaIds[representative],
                context + $" representative {representative + 1}",
                updateReference: false,
                requireDocumentTableFree: false);
        }
    }

    private static void AppendNativeHashAcceptancePageBreak(Word.Document document)
    {
        Word.Paragraph? paragraph = null;
        Word.Range? range = null;
        Word.OMaths? maths = null;
        try
        {
            paragraph = document.Paragraphs.Add();
            range = paragraph.Range.Duplicate;
            range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            maths = range.OMaths;
            AssertEqual(0, maths.Count,
                "The forced page-break insertion point was absorbed into OMath.");
            range.InsertBreak(Word.WdBreakType.wdPageBreak);
        }
        finally
        {
            Release(maths);
            Release(range);
            Release(paragraph);
        }
    }
}
