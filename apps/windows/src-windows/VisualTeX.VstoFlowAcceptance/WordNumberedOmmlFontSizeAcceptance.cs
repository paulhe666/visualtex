using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordNumberedOmmlFontSizeAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(artifactRoot, "numbered-omml-font-size.docx");
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
            var formulaId = Guid.NewGuid().ToString("D");
            var insertion = document.Content.End - 1;
            application.Selection.SetRange(insertion, insertion);
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertion,
                insertion,
                @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);
            service.InsertOmml(session, QuadraticFormulaMathMl());

            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                context: "numbered OMML before font-size changes",
                updateReference: true);

            foreach (var target in new[] { 12f, 16f })
            {
                Release(equationRange); equationRange = null;
                Release(bookmark); bookmark = null;
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                    ?? throw new InvalidDataException("Numbered OMML font-size target bookmark is missing.");
                equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                Word.Application? liveApplication = null;
                Word.Selection? liveSelection = null;
                try
                {
                    document.Activate();
                    liveApplication = document.Application;
                    liveSelection = liveApplication.Selection
                        ?? throw new InvalidOperationException(
                            "Word did not expose a live selection for the numbered OMML font-size acceptance.");
                    liveSelection.SetRange(equationRange.Start, equationRange.End);
                }
                finally
                {
                    Release(liveSelection);
                    Release(liveApplication);
                }
                var applied = service.SetSelectedFormulaFontSize(target);
                AssertNear(target, applied, 0.01f,
                    $"Numbered OMML font-size command did not return {target:0.##} pt.");

                Release(equationRange); equationRange = null;
                Release(bookmark); bookmark = null;
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                    ?? throw new InvalidDataException("Numbered OMML metadata was lost after a font-size change.");
                AssertNear(
                    target,
                    (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                    0.01f,
                    $"Numbered OMML semantic font size did not persist {target:0.##} pt.");
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                    ?? throw new InvalidDataException("Numbered OMML bookmark was lost after a font-size change.");
                equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                AssertOmmlTabNumberingHost(
                    document,
                    formulaId,
                    context: $"numbered OMML after {target:0.##} pt font-size change",
                    updateReference: true);
                var metric = ReadNumberedOmmlBodyLayoutMetric(
                    application,
                    document,
                    equationRange,
                    $"numbered OMML {target:0.##} pt body");
                AssertTrue(metric.HeightPx > 0 && metric.WidthPx > 0,
                    $"Numbered OMML {target:0.##} pt geometry is empty.");
                AssertNear(target, metric.FontSizePt, 0.35f,
                    $"Numbered OMML native display runs did not retain {target:0.##} pt.");
            }

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            var reopenedMetadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException("Save/reopen lost numbered OMML font-size metadata.");
            AssertNear(
                16f,
                (float)FormulaFontSize.ResolveSemanticFontSize(reopenedMetadata),
                0.01f,
                "Save/reopen changed the final numbered OMML font size.");
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException("Save/reopen lost numbered OMML bookmark.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                context: "numbered OMML 16 pt after save/reopen",
                updateReference: true);
            var reopenedMetric = ReadNumberedOmmlBodyLayoutMetric(
                application,
                document,
                equationRange,
                "numbered OMML 16 pt reopened body");
            AssertNear(16f, reopenedMetric.FontSizePt, 0.35f,
                "Save/reopen changed the final native display run size.");

            Console.WriteLine(
                "Word numbered OMML font-size acceptance passed: 14→12→16 pt retained genuine center-cell wdOMathDisplay/m:oMathPara geometry while the ordinary right-cell number kept paragraph-inherited Word typography, with external REF numbering and the minimal 1x3 host stable through save/reopen.");
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
