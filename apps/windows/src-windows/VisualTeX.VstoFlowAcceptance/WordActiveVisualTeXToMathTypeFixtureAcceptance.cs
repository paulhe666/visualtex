using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordActiveVisualTeXToMathTypeFixtureAcceptance(string artifactRoot)
    {
        AssertTrue(AttachActiveWord,
            "Active VisualTeX→MathType fixture acceptance requires VISUALTEX_VSTO_ACCEPTANCE_ATTACH_WORD=1.");
        Directory.CreateDirectory(artifactRoot);
        var olePreviewRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX", "office", "temp", "active-vt-to-mt-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(olePreviewRoot);
        var svgPath = Path.Combine(olePreviewRoot, "target.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"96\" viewBox=\"0 0 260 96\"><text x=\"6\" y=\"66\" font-family=\"Cambria Math\" font-size=\"48\">x+1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 260, 96);
        const string mathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mn>1</mn></math>";

        Word.Application? application = null;
        Word.Document? sourceDocument = null;
        Word.Document? document = null;
        Word.InlineShape? sourceShape = null;
        Word.Range? ownerRange = null;
        Word.Bookmark? captionBookmark = null;
        Word.Range? captionRange = null;
        Word.Range? sourceHost = null;
        Word.Range? insertion = null;
        Word.InlineShape? pastedShape = null;
        try
        {
            application = CreateWordApplication(visible: false);
            sourceDocument = application.ActiveDocument
                ?? throw new InvalidOperationException("Active Word has no source document.");
            string? sourceFormulaId = null;
            for (var index = 1; index <= sourceDocument.InlineShapes.Count; index++)
            {
                Release(sourceShape);
                sourceShape = sourceDocument.InlineShapes[index];
                if (!WordFormulaMetadataReader.IsNativeOle(sourceShape)) continue;
                var metadata = WordFormulaMetadataReader.TryRead(sourceShape);
                if (metadata?.Numbered != true
                    || !string.Equals(metadata.DisplayMode, "block", StringComparison.OrdinalIgnoreCase))
                    continue;
                sourceFormulaId = metadata.FormulaId;
                ownerRange = WordEquationNumbering.FindNumberingOwnerRange(
                    sourceDocument, sourceFormulaId)
                    ?? throw new InvalidDataException("Active VisualTeX formula has no numbered owner range.");
                var hostEnd = ownerRange.End;
                var captionName = WordEquationNumbering.NativeCaptionBookmarkName(sourceFormulaId);
                if (sourceDocument.Bookmarks.Exists(captionName))
                {
                    captionBookmark = sourceDocument.Bookmarks[captionName];
                    captionRange = captionBookmark.Range;
                    hostEnd = Math.Max(hostEnd, captionRange.End);
                }
                sourceHost = sourceDocument.Range(ownerRange.Start, hostEnd);
                break;
            }
            if (sourceHost is null || string.IsNullOrWhiteSpace(sourceFormulaId))
                throw new InvalidOperationException("Active Word document has no complete numbered VisualTeX OLE fixture.");

            document = application.Documents.Add();
            WordEquationNumbering.SetEquationNumberFormat(
                document,
                WordEquationNumbering.GetEquationNumberFormatId(sourceDocument));
            insertion = document.Range(document.Content.Start, document.Content.Start);
            sourceHost.Copy();
            insertion.Paste();
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(150);
            AssertEqual(1, document.InlineShapes.Count,
                "Active VisualTeX fixture copy did not produce exactly one OLE source.");

            pastedShape = document.InlineShapes[1];
            pastedShape.Range.Select();
            var service = new WordFormulaService(application);
            _ = service.ReadSelection();
            Release(pastedShape);
            pastedShape = document.InlineShapes[1];
            var repairedMetadata = WordFormulaMetadataReader.TryRead(pastedShape)
                ?? throw new InvalidDataException("Copied VisualTeX source lost metadata after identity repair.");
            Console.WriteLine(
                $"[ACTIVE VT→MT FIXTURE] formulaId={repairedMetadata.FormulaId} owner={WordEquationNumbering.FindNumberingOwnerRange(document, repairedMetadata.FormulaId)?.Start}:{WordEquationNumbering.FindNumberingOwnerRange(document, repairedMetadata.FormulaId)?.End} fields={document.Fields.Count} bookmarks={document.Bookmarks.Count}.");

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.NativeOleMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(1, plan.Targets.Count,
                "Active VisualTeX fixture conversion did not capture exactly one source.");
            var target = plan.Targets[0];
            var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal)
            {
                [target.Id] = new PreparedWordBulkFormula
                {
                    Run = new WordBulkRun
                    {
                        Id = target.Id,
                        IsFormula = true,
                        Latex = target.Latex,
                        DisplayMode = target.DisplayMode,
                    },
                    Session = CreateSimpleMathTypeTargetSession(target, mathMl),
                    MathMl = mathMl,
                    EmfPath = emfPath,
                },
            };
            var result = service.ApplyFormulaFormatConversionPlan(plan, prepared);
            Console.WriteLine(
                $"[ACTIVE VT→MT FIXTURE] converted={result.FormulaCount} failed={result.FailedFormulaCount} failures={string.Join(" | ", result.Failures)}");
            AssertEqual(1, result.FormulaCount,
                "Active numbered VisualTeX source did not convert to MathType.");
            AssertEqual(0, result.FailedFormulaCount,
                $"Active numbered VisualTeX source conversion failed: {string.Join(" | ", result.Failures)}");
            AssertEqual(0, CountVisualTeXNativeOleShapes(document),
                "Active fixture conversion left the VisualTeX OLE source behind.");
            AssertEqual(1, CountMathTypeOleShapes(document),
                "Active fixture conversion did not create exactly one MathType OLE.");

            var outputPath = Path.Combine(artifactRoot, "Active-VisualTeX-To-MathType-Fixed.docx");
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine("[ACTIVE VT→MT FIXTURE] Passed: " + outputPath);
        }
        finally
        {
            Release(pastedShape);
            Release(insertion);
            Release(sourceHost);
            Release(captionRange);
            Release(captionBookmark);
            Release(ownerRange);
            Release(sourceShape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            if (sourceDocument is not null)
            {
                try { sourceDocument.Activate(); } catch { }
            }
            Release(sourceDocument);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            try { Directory.Delete(olePreviewRoot, recursive: true); } catch { }
        }
    }
}
