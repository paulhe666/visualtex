using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
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
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"numbered-omml-empty-row-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);
        var svgPath = Path.Combine(assetRoot, "numbered-omml-empty-row.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"76\" viewBox=\"0 0 240 76\"><text x=\"4\" y=\"50\" font-size=\"34\">x = (-b ± √D) / 2a</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 76);
        var pngDataUrl = CreatePngDataUrl(
            "word-numbered-omml-empty-row",
            240,
            76);
        var pngPath = Path.Combine(assetRoot, "numbered-omml-empty-row.png");
        File.WriteAllBytes(
            pngPath,
            Convert.FromBase64String(
                pngDataUrl.Substring(pngDataUrl.IndexOf(',') + 1)));

        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? sourceShape = null;
        Word.Range? sourceRange = null;
        Word.Table? legacyTable = null;
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

            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var insertion = document.Range(
                document.Content.End - 1,
                document.Content.End - 1);
            try
            {
                application.Selection.SetRange(insertion.Start, insertion.End);
                var createSession = CreateNumberedPerformanceSession(
                    "create",
                    formulaId,
                    document.FullName,
                    WordRangeReference(insertion.Start, insertion.End),
                    originalMetadata: null,
                    latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}");
                createSession.ExportResult = new OfficeExportDocument
                {
                    Width = 240,
                    Height = 76,
                    Baseline = 57,
                    FormulaLetterFont = "katex",
                    FormulaChineseFont = "system",
                };
                service.InsertOle(createSession, pngPath, emfPath);
            }
            finally
            {
                Release(insertion);
            }

            sourceShape = FindVisualTeXOleByFormulaIdForNumberToggle(
                document,
                formulaId);
            var sourceMetadata = WordFormulaMetadataReader.TryRead(sourceShape)
                ?? throw new InvalidDataException(
                    "The numbered VisualTeX OLE source lost its metadata before OMML replacement.");
            sourceRange = sourceShape.Range;
            legacyTable = WordEquationNumbering.FindNumberedEquationTable(
                    document,
                    formulaId)
                ?? throw new InvalidDataException(
                    "The empty-row regression fixture did not start from a managed VisualTeX numbering table.");
            emptyRow = legacyTable.Rows.Add();
            AssertEqual(2, legacyTable.Rows.Count,
                "The empty-row regression fixture did not produce a 2x3 managed numbering table.");
            AssertEqual(3, legacyTable.Columns.Count,
                "The empty-row regression fixture changed the managed numbering-table column count.");

            var replaceSession = CreateNumberedPerformanceSession(
                "edit",
                formulaId,
                document.FullName,
                WordRangeReference(sourceRange.Start, sourceRange.End),
                sourceMetadata,
                latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}");
            replaceSession.ObjectMode = FormulaOleContract.WordOmmlMode;
            replaceSession.ExportResult = new OfficeExportDocument
            {
                FormulaLetterFont = "katex",
                FormulaChineseFont = "system",
            };
            service.ReplaceOmml(
                replaceSession,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
                + "<mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo>"
                + "<msqrt><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></msqrt>"
                + "</mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>");

            // ReplaceOmml owns the structural migration. The old table RCW is
            // intentionally stale once Word converts/deletes that host; never call
            // Delete on it after the replacement has committed.
            Release(emptyRow); emptyRow = null;
            Release(legacyTable); legacyTable = null;
            Release(sourceRange); sourceRange = null;
            Release(sourceShape); sourceShape = null;

            AssertNumberedOmmlTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "OLE-to-OMML replacement with benign empty legacy row");
            AssertTrue(
                WordEquationNumbering.FindNumberedEquationTable(document, formulaId) is null,
                "OLE-to-OMML replacement left a legacy numbering table after tab migration.");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            AssertNumberedOmmlTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "saved/reopened OLE-to-OMML empty-row replacement");

            Console.WriteLine(
                "Numbered OLE-to-OMML empty-row acceptance passed: the first replacement removed the benign extra legacy row, migrated the managed 1x3 host to a table-free center/right-tab OMath paragraph, and save/reopen preserved it.");
        }
        finally
        {
            Release(emptyRow);
            Release(legacyTable);
            Release(sourceRange);
            Release(sourceShape);
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
}
