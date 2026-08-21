using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordVisualTeXNumberParenthesisAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(artifactRoot, "visualtex-number-parenthesis.docx");
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"number-parenthesis-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);
        var svgPath = Path.Combine(assetRoot, "visualtex-number-parenthesis.svg");
        File.WriteAllText(svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"220\" height=\"70\" viewBox=\"0 0 220 70\"><text x=\"4\" y=\"48\" font-size=\"36\">x = 1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 220, 70);
        var pngDataUrl = CreatePngDataUrl("number-parenthesis", 220, 70);
        var pngPath = Path.Combine(assetRoot, "visualtex-number-parenthesis.png");
        File.WriteAllBytes(
            pngPath,
            Convert.FromBase64String(pngDataUrl.Substring(pngDataUrl.IndexOf(',') + 1)));

        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
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
            var insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            try
            {
                application.Selection.SetRange(insertion.Start, insertion.End);
                var createSession = CreateNumberedPerformanceSession(
                    "create",
                    formulaId,
                    document.FullName,
                    WordRangeReference(insertion.Start, insertion.End),
                    originalMetadata: null,
                    latex: @"x=1");
                createSession.ExportResult = new OfficeExportDocument
                {
                    Width = 220,
                    Height = 70,
                    Baseline = 52.5f,
                };
                service.InsertOle(createSession, pngPath, emfPath);
            }
            finally { Release(insertion); }

            AssertVisualTeXNumberParenthesisBoundary(
                document,
                formulaId,
                updateReference: true,
                context: "fresh numbered VisualTeX insertion");

            shape = FindNumberedOleByFormulaId(document, formulaId);
            var originalMetadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException("Fresh numbered VisualTeX OLE lost metadata before edit.");
            shapeRange = shape.Range;
            var editSession = CreateNumberedPerformanceSession(
                "edit",
                formulaId,
                document.FullName,
                WordRangeReference(shapeRange.Start, shapeRange.End),
                originalMetadata,
                latex: @"x=2");
            editSession.ExportResult = new OfficeExportDocument
            {
                Width = 220,
                Height = 70,
                Baseline = 52.5f,
            };
            service.ReplaceOle(editSession, pngPath, emfPath);
            Release(shapeRange); shapeRange = null;
            Release(shape); shape = null;

            AssertVisualTeXNumberParenthesisBoundary(
                document,
                formulaId,
                updateReference: true,
                context: "numbered VisualTeX edit/reconcile");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(documentPath, ReadOnly: false, Visible: false);
            AssertVisualTeXNumberParenthesisBoundary(
                document,
                formulaId,
                updateReference: true,
                context: "numbered VisualTeX save/reopen");

            Console.WriteLine(
                "VisualTeX number-parenthesis acceptance passed: fresh insert, REF update, edit/reconcile and save/reopen all kept ')' outside REF.Result.");
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
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

    private static void AssertVisualTeXNumberParenthesisBoundary(
        Word.Document document,
        string formulaId,
        bool updateReference,
        string context)
    {
        Word.Table? table = null;
        Word.Cell? cell = null;
        Word.Range? cellRange = null;
        Word.Fields? fields = null;
        Word.Field? reference = null;
        Word.Range? code = null;
        Word.Range? result = null;
        try
        {
            table = WordEquationNumbering.FindNumberedEquationTable(document, formulaId)
                ?? throw new InvalidDataException(context + ": VisualTeX numbered table is missing.");
            cell = table.Cell(1, 3);
            cellRange = cell.Range;
            var visible = (cellRange.Text ?? string.Empty).TrimEnd('\r', '\a');
            AssertTrue(
                visible.StartsWith("(", StringComparison.Ordinal)
                && visible.EndsWith(")", StringComparison.Ordinal),
                context + ": visible equation number is not enclosed by both parentheses: '" + visible + "'.");

            fields = cellRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? candidate = null;
                Word.Range? candidateCode = null;
                try
                {
                    candidate = fields[index];
                    candidateCode = candidate.Code;
                    if ((candidateCode.Text ?? string.Empty).IndexOf(
                            "REF " + WordEquationNumbering.NativeNumberBookmarkName(formulaId),
                            StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    reference = candidate;
                    candidate = null;
                    break;
                }
                finally
                {
                    Release(candidateCode);
                    Release(candidate);
                }
            }
            if (reference is null)
                throw new InvalidDataException(context + ": visible number REF field is missing.");

            if (updateReference)
                reference.Update();
            Release(cellRange);
            cellRange = cell.Range;
            visible = (cellRange.Text ?? string.Empty).TrimEnd('\r', '\a');
            AssertTrue(
                visible.StartsWith("(", StringComparison.Ordinal)
                && visible.EndsWith(")", StringComparison.Ordinal),
                context + ": REF update removed a parenthesis: '" + visible + "'.");

            code = reference.Code;
            result = reference.Result;
            var resultText = result.Text ?? string.Empty;
            AssertTrue(
                resultText.IndexOf('(') < 0 && resultText.IndexOf(')') < 0,
                context + ": parenthesis leaked inside REF.Result: '" + resultText + "'.");
            AssertTrue(
                result.End < cellRange.End - 1,
                context + ": REF.Result reaches the table-cell end; ')' is not safely outside the field result.");
            Console.WriteLine(
                $"  {context}: cell='{visible}', REF.Result='{resultText}', result=[{result.Start},{result.End}], cell=[{cellRange.Start},{cellRange.End}]");
        }
        finally
        {
            Release(result);
            Release(code);
            Release(reference);
            Release(fields);
            Release(cellRange);
            Release(cell);
            Release(table);
        }
    }
}
