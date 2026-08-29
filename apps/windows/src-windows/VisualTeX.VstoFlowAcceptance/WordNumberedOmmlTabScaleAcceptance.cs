using System.Diagnostics;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordNumberedOmmlTabScaleAcceptance(string artifactRoot)
    {
        var formulaCount = int.TryParse(
                Environment.GetEnvironmentVariable("VISUALTEX_OMML_SCALE_COUNT"),
                out var requestedCount)
            ? Math.Max(1, Math.Min(100, requestedCount))
            : 20;
        var verbose = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_OMML_SCALE_VERBOSE"),
            "1",
            StringComparison.Ordinal);
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(artifactRoot, "numbered-omml-tab-scale.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
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
            var formulaIds = new List<string>(formulaCount);
            var watch = Stopwatch.StartNew();
            var previousPerfSuppression = Environment.GetEnvironmentVariable(
                "VISUALTEX_VSTO_SUPPRESS_ACCEPTANCE_PERF");
            try
            {
                // Twenty sequential native-OMML inserts intentionally exercise the
                // document-scale numbering path. Suppress only the repetitive stage
                // timings while retaining normal Console/Error handling and one
                // bounded heartbeat per formula. Replacing global Console.Out made
                // failures impossible to diagnose in remote Windows job hosts.
                if (!verbose)
                {
                    Environment.SetEnvironmentVariable(
                        "VISUALTEX_VSTO_SUPPRESS_ACCEPTANCE_PERF",
                        "1");
                }
                for (var index = 1; index <= formulaCount; index++)
                {
                    var formulaId = Guid.NewGuid().ToString("D");
                    formulaIds.Add(formulaId);
                    var insertion = document.Content.End - 1;
                    application.Selection.SetRange(insertion, insertion);
                    var session = CreateNumberedOmmlTabSession(
                        formulaId,
                        document.FullName,
                        insertion,
                        insertion,
                        $@"x_{{{index}}}=\frac{{a_{{{index}}}+b_{{{index}}}}}{{{index + 1}}}",
                        originalMetadata: null);
                    service.InsertOmml(
                        session,
                        $"<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><msub><mi>x</mi><mn>{index}</mn></msub><mo>=</mo><mfrac><mrow><msub><mi>a</mi><mn>{index}</mn></msub><mo>+</mo><msub><mi>b</mi><mn>{index}</mn></msub></mrow><mn>{index + 1}</mn></mfrac></mrow></math>");
                    // Emit one bounded heartbeat per formula so remote acceptance
                    // runners do not mistake the intentionally quiet COM operation
                    // for an idle/hung child and tear down its Word job object.
                    Console.WriteLine(
                        $"  numbered OMML scale insertion {index}/{formulaCount} completed.");
                }
            }
            finally
            {
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_VSTO_SUPPRESS_ACCEPTANCE_PERF",
                    previousPerfSuppression);
                watch.Stop();
            }
            AssertEqual(0, document.Tables.Count,
                "Numbered OMML scale insertion created one or more legacy tables.");
            AssertEqual(formulaCount, WordOmmlFormulaStore.FormulaIds(document).Count,
                "Numbered OMML scale insertion lost one or more managed formulas.");
            foreach (var index in new[] { 0, formulaCount / 2, formulaCount - 1 })
            {
                AssertNumberedOmmlTabHost(
                    document,
                    formulaIds[index],
                    updateReference: true,
                    context: $"numbered OMML scale formula #{index + 1}",
                    expectedDigit: (index + 1).ToString(
                        System.Globalization.CultureInfo.InvariantCulture));
            }

            var middleId = formulaIds[formulaCount / 2];
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, middleId)
                ?? throw new InvalidDataException("Scale edit target bookmark is missing.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var middleMetadata = WordOmmlFormulaStore.TryRead(document, middleId)
                ?? throw new InvalidDataException("Scale edit target metadata is missing.");
            var editSession = CreateNumberedOmmlTabSession(
                middleId,
                document.FullName,
                equationRange.Start,
                equationRange.End,
                @"y=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                middleMetadata);
            var editWatch = Stopwatch.StartNew();
            service.ReplaceOmml(editSession, QuadraticFormulaMathMl().Replace("<mi>x</mi>", "<mi>y</mi>"));
            editWatch.Stop();
            Release(equationRange); equationRange = null;
            Release(bookmark); bookmark = null;
            AssertNumberedOmmlTabHost(
                document,
                middleId,
                updateReference: true,
                context: "numbered OMML scale middle edit",
                expectedLatinVariable: "y",
                expectedDigit: "2");
            AssertEqual(0, document.Tables.Count,
                "Editing one formula in the OMML scale document recreated a legacy table.");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            AssertEqual(
                formulaCount,
                WordEquationNumbering.RefreshNumberedOmmlTabLayouts(document),
                "Document-open OMML tab refresh did not process every numbered native equation.");
            AssertEqual(0, document.Tables.Count,
                "Numbered OMML scale save/reopen recreated a legacy table.");
            foreach (var index in new[] { 0, formulaCount / 2, formulaCount - 1 })
            {
                AssertNumberedOmmlTabHost(
                    document,
                    formulaIds[index],
                    updateReference: true,
                    context: $"numbered OMML scale save/reopen #{index + 1}",
                    expectedLatinVariable:
                        index == formulaCount / 2 ? "y" : "x",
                    expectedDigit:
                        index == formulaCount / 2
                            ? "2"
                            : (index + 1).ToString(
                                System.Globalization.CultureInfo.InvariantCulture));
            }

            Console.WriteLine(
                $"Numbered OMML tab scale acceptance passed: formulas={formulaCount}, insertMs={watch.ElapsedMilliseconds}, middleEditMs={editWatch.ElapsedMilliseconds}, tables=0.");
        }
        finally
        {
            Release(equationRange);
            Release(bookmark);
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
