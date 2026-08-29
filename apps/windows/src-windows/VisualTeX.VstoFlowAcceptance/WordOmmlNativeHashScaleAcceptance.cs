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

            UpdateNativeHashProductionFields(document, formulaIds);
            var expected = Enumerable.Range(1, 20)
                .Select(value => value.ToString(
                    System.Globalization.CultureInfo.InvariantCulture))
                .ToArray();
            AssertNativeHashProductionNumbers(
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
            AssertNativeHashProductionNumbers(
                document,
                formulaIds,
                expected,
                "20-formula numbering after body REF insertion");
            Console.WriteLine(
                $"  native-hash scale stage passed: 20 formulas, pages={pages}, numbers 1..20, body REF 1/10/20");

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
            UpdateNativeHashProductionFields(document, formulaIds);
            AssertNativeHashProductionNumbers(
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
            AssertEqual(0, document.Tables.Count,
                "20-formula save/reopen created a Word table.");

            Console.WriteLine(
                "Production OMML native #(SEQ) scale acceptance passed: 20 formulas remained 1..20 across forced page breaks and save/reopen, with body REF 1/10/20 and no Shape/Table.");
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
