using System.Text;
using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordBulkImportMathTypeAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourcePath = Path.Combine(artifactRoot, "word-bulk-import-mathtype.tex");
        var logPath = Path.Combine(artifactRoot, "word-bulk-import-mathtype.log");
        var documentPath = Path.Combine(artifactRoot, "word-bulk-import-mathtype.docx");
        const string source = """
            Before MathType bulk formula.

            \[
            x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}
            \]

            Between MathType bulk formulas.

            \[
            e^{i\pi}+1=0
            \]

            After MathType bulk formula.
            """;
        File.WriteAllText(sourcePath, source, new UTF8Encoding(false));
        DeleteBulkPerformanceArtifact(logPath);
        DeleteBulkPerformanceArtifact(documentPath);

        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", sourcePath);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", "latex");
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", "mathtype");
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", logPath);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);

            addIn.OnBulkImport(new object());
            WaitForBulkImportCompletion(logPath, TimeSpan.FromMinutes(3));
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));

            var service = new WordFormulaService(application);
            AssertEqual(0, document.OMaths.Count,
                "MathType bulk import unexpectedly inserted native Word OMath objects.");
            var imported = ReadBulkMathTypeMetadata(service, document, "fresh MathType bulk import");
            AssertEqual(2, imported.Count,
                "MathType bulk import did not create exactly two editable MathType OLE formulas.");
            AssertTrue(imported.Any(item => item.Latex.IndexOf("\\frac", StringComparison.Ordinal) >= 0),
                "MathType bulk import lost the quadratic formula metadata.");
            AssertTrue(imported.Any(item => item.Latex.IndexOf("i\\pi", StringComparison.Ordinal) >= 0),
                "MathType bulk import lost the Euler formula metadata.");
            AssertBulkMathTypeProseOrder(document);

            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            reopened = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            reopened.Activate();
            var reopenedService = new WordFormulaService(application);
            AssertEqual(0, reopened.OMaths.Count,
                "Save/reopen converted one or more MathType bulk formulas into Word OMath.");
            var reopenedMetadata = ReadBulkMathTypeMetadata(
                reopenedService,
                reopened,
                "save/reopened MathType bulk import");
            AssertEqual(2, reopenedMetadata.Count,
                "Save/reopen lost one or more MathType bulk formulas.");
            AssertBulkMathTypeProseOrder(reopened);

            Console.WriteLine(
                "Word MathType bulk-import acceptance passed: two formulas were inserted as native MathType OLE, remained editable through VisualTeX routing, adjacent prose stayed intact, and save/reopen preserved MathType objects.");
            Console.WriteLine($"Artifact: {documentPath}");
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", null);
            if (addIn is not null)
            {
                try
                {
                    addIn.OnDisconnection(
                        ext_DisconnectMode.ext_dm_UserClosed,
                        ref custom);
                }
                catch { }
            }
            try { reopened?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(reopened);
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static List<FormulaMetadata> ReadBulkMathTypeMetadata(
        WordFormulaService service,
        Word.Document document,
        string context)
    {
        var metadata = new List<FormulaMetadata>();
        Word.InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                Word.Range? range = null;
                try
                {
                    shape = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                    range = shape.Range;
                    range.Select();
                    var selection = service.ReadSelection();
                    AssertEqual(
                        FormulaOleContract.MathTypeOleMode,
                        selection.ObjectMode,
                        $"{context}: MathType OLE #{index} reports the wrong object mode.");
                    var item = selection.Metadata
                        ?? throw new InvalidDataException(
                            $"{context}: MathType OLE #{index} has no editable VisualTeX metadata.");
                    AssertTrue(!string.IsNullOrWhiteSpace(item.Latex),
                        $"{context}: MathType OLE #{index} has empty LaTeX metadata.");
                    metadata.Add(item);
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }
        }
        finally { Release(shapes); }
        return metadata;
    }

    private static void AssertBulkMathTypeProseOrder(Word.Document document)
    {
        Word.Range? content = null;
        try
        {
            content = document.Content;
            var text = content.Text ?? string.Empty;
            var before = text.IndexOf("Before MathType bulk formula.", StringComparison.Ordinal);
            var between = text.IndexOf("Between MathType bulk formulas.", StringComparison.Ordinal);
            var after = text.IndexOf("After MathType bulk formula.", StringComparison.Ordinal);
            AssertTrue(before >= 0 && between > before && after > between,
                "MathType bulk import damaged or reordered adjacent prose.");
        }
        finally { Release(content); }
    }
}
