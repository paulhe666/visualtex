using System.Text;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordConvertedVisualTeXNumberingScaffoldAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var tempRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(tempRoot);
        var svgPath = Path.Combine(tempRoot, $"converted-numbering-{Guid.NewGuid():N}.svg");
        var pngPath = Path.Combine(tempRoot, $"converted-numbering-{Guid.NewGuid():N}.png");
        var documentPath = Path.Combine(artifactRoot, "converted-visualtex-numbering-scaffold.docx");
        var xmlPath = Path.Combine(artifactRoot, "converted-visualtex-numbering-scaffold.xml");

        string? emfPath = null;
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            File.WriteAllText(
                svgPath,
                CreateFontAcceptanceSvg("Times New Roman", "SimSun"),
                new UTF8Encoding(false));
            WriteAcceptancePng(pngPath, "x=1", 240, 72);
            emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 72);

            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            document.SaveAs2(
                documentPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);

            var service = new WordFormulaService(application);
            var formulaIds = new List<string>();
            for (var index = 1; index <= 2; index++)
            {
                InsertNumberingHeading(
                    application,
                    document,
                    level: 1,
                    text: "Converted Chapter " + index);
                application.Selection.EndKey(Word.WdUnits.wdStory);
                var insertion = application.Selection.Range;
                try
                {
                    var formulaId = Guid.NewGuid().ToString("D");
                    var session = CreateNumberedPerformanceSession(
                        "create",
                        formulaId,
                        document.FullName,
                        WordRangeReference(insertion.Start, insertion.End),
                        originalMetadata: null,
                        latex: index == 1 ? @"\hbar\omega+1" : @"\int_0^1 x^2\,dx");
                    service.InsertOle(
                        session,
                        pngPath,
                        emfPath,
                        deferNumberingLayout: true,
                        preserveExistingDisplayParagraphBoundary: true);
                    formulaIds.Add(formulaId);
                }
                finally { Release(insertion); }
            }

            AssertEqual(2, CountInstalledVisualTeXOleShapes(document),
                "Deferred converted-VisualTeX fixture did not contain two VisualTeX OLE hosts.");
            AssertEqual(0, document.Tables.Count,
                "Deferred converted-VisualTeX fixture unexpectedly created a table before numbering scaffold finalization.");

            foreach (var formulaId in formulaIds.AsEnumerable().Reverse())
            {
                Word.InlineShape? shape = null;
                Word.Range? range = null;
                try
                {
                    shape = FindNumberedOleByFormulaId(document, formulaId);
                    var metadata = WordFormulaMetadataReader.TryRead(shape)
                        ?? throw new InvalidDataException(
                            $"Deferred converted VisualTeX formula {formulaId} has no metadata.");
                    AssertTrue(metadata.Numbered,
                        $"Deferred converted VisualTeX formula {formulaId} lost Numbered=true.");
                    range = shape.Range;
                    WordEquationNumbering.BuildFormulaNumberingScaffoldForConversion(
                        document,
                        range,
                        shape.Height,
                        metadata);
                    Console.WriteLine(
                        $"  built conversion numbering scaffold formulaId={formulaId} range={range.Start}:{range.End}.");
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }

            Word.Range? content = null;
            try
            {
                content = document.Content;
                File.WriteAllText(xmlPath, content.WordOpenXML ?? string.Empty, new UTF8Encoding(false));
            }
            finally { Release(content); }

            var finalized = WordEquationNumbering.TryFinalizeHealthyConversionNumbering(
                document,
                out var updated);
            Console.WriteLine(
                $"  conversion numbering finalizer result={finalized}, updated={updated}, tables={document.Tables.Count}, shapes={document.InlineShapes.Count}.");
            AssertTrue(finalized,
                "Converted VisualTeX numbering scaffolds failed the shared conversion finalizer.");
            AssertEqual(2, updated,
                "Converted VisualTeX numbering finalizer did not update two formulas.");
            AssertEqual(2, CountInstalledVisualTeXNumberedFormulaHosts(document),
                "Converted VisualTeX numbering finalizer did not leave two numbered OLE hosts.");

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
            AssertEqual(2, CountInstalledVisualTeXNumberedFormulaHosts(document),
                "Save/reopen changed the two converted VisualTeX numbered hosts.");
            AssertTrue(
                WordEquationNumbering.TryFinalizeHealthyConversionNumbering(document, out updated),
                "Save/reopen converted VisualTeX numbering failed the shared finalizer.");
            AssertEqual(2, updated,
                "Save/reopen converted VisualTeX numbering finalizer did not update two formulas.");

            Console.WriteLine(
                "Converted VisualTeX numbering scaffold acceptance passed: two deferred numbered OLE targets built conversion scaffolds, finalized through the shared fast path, and survived save/reopen.");
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
}
