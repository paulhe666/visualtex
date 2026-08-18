using System.Runtime.InteropServices;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordFormatConversionRollbackResidualFixtureAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? source = null;
        Word.Document? temporary = null;
        Word.InlineShape? sourceShape = null;
        Word.Range? sourceShapeRange = null;
        Word.Paragraphs? sourceParagraphs = null;
        Word.Paragraph? sourceParagraph = null;
        Word.Range? sourceParagraphRange = null;
        Word.Range? probe = null;
        Word.Range? copy = null;
        Word.Range? destination = null;
        Word.InlineShape? temporaryShape = null;
        try
        {
            application = (Word.Application)Marshal.GetActiveObject("Word.Application");
            source = application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document is available for rollback-residual fixture acceptance.");

            FormulaMetadata? sourceMetadata = null;
            string? bridge = null;
            var copyStart = -1;
            var copyEnd = -1;
            for (var index = 1; index <= source.InlineShapes.Count; index++)
            {
                Release(sourceShape);
                sourceShape = source.InlineShapes[index];
                if (!WordFormulaMetadataReader.IsNativeOle(sourceShape)) continue;
                var metadata = WordFormulaMetadataReader.TryRead(sourceShape);
                if (metadata is null || string.IsNullOrWhiteSpace(metadata.Latex)) continue;
                var candidateBridge = string.Equals(
                        metadata.DisplayMode,
                        "block",
                        StringComparison.OrdinalIgnoreCase)
                    ? "$$" + metadata.Latex.Trim() + "$$"
                    : "$" + metadata.Latex.Trim().Replace('\r', ' ').Replace('\n', ' ') + "$";

                Release(sourceShapeRange);
                sourceShapeRange = sourceShape.Range.Duplicate;
                Release(sourceParagraphs);
                sourceParagraphs = sourceShapeRange.Paragraphs;
                if (sourceParagraphs.Count == 0) continue;
                Release(sourceParagraph);
                sourceParagraph = sourceParagraphs[1];
                Release(sourceParagraphRange);
                sourceParagraphRange = sourceParagraph.Range.Duplicate;
                var probeStart = sourceParagraphRange.End;
                var probeEnd = Math.Min(
                    source.Content.End,
                    probeStart + candidateBridge.Length + 4);
                Release(probe);
                probe = source.Range(probeStart, probeEnd);
                var probeText = probe.Text ?? string.Empty;
                var bridgeOffset = probeText.IndexOf(candidateBridge, StringComparison.Ordinal);
                if (bridgeOffset < 0 || bridgeOffset > 4) continue;

                sourceMetadata = metadata;
                bridge = candidateBridge;
                copyStart = sourceParagraphRange.Start;
                copyEnd = probeStart + bridgeOffset + candidateBridge.Length;
                break;
            }

            if (sourceShape is null || sourceMetadata is null || bridge is null || copyStart < 0 || copyEnd <= copyStart)
                throw new InvalidDataException(
                    "The active Word fixture has no VisualTeX source formula immediately followed by its temporary LaTeX rollback bridge.");

            copy = source.Range(copyStart, copyEnd);
            temporary = application.Documents.Add();
            destination = temporary.Range(0, 0);
            destination.FormattedText = copy.FormattedText;
            AssertEqual(1, CountVisualTeXNativeOleShapes(temporary),
                "Rollback-residual fixture copy did not preserve exactly one VisualTeX OLE source.");
            AssertTrue(
                (temporary.Content.Text ?? string.Empty).IndexOf(bridge, StringComparison.Ordinal) >= 0,
                "Rollback-residual fixture copy did not preserve the adjacent temporary LaTeX bridge.");

            temporaryShape = temporary.InlineShapes[1];
            var temporaryMetadata = WordFormulaMetadataReader.TryRead(temporaryShape)
                ?? throw new InvalidDataException("Copied VisualTeX source lost metadata in rollback-residual fixture.");
            var target = new WordFormulaFormatConversionTarget
            {
                SourceFormulaId = temporaryMetadata.FormulaId,
                SourceObjectId = string.Empty,
                SourceStart = temporaryShape.Range.Start,
                Latex = temporaryMetadata.Latex,
                DisplayMode = temporaryMetadata.DisplayMode,
                Numbered = temporaryMetadata.Numbered,
                Metadata = temporaryMetadata,
            };

            Word.Range? temporaryShapeRange = null;
            Word.Paragraphs? temporaryParagraphs = null;
            Word.Paragraph? temporaryParagraph = null;
            Word.Range? temporaryParagraphRange = null;
            try
            {
                temporaryShapeRange = temporaryShape.Range.Duplicate;
                temporaryParagraphs = temporaryShapeRange.Paragraphs;
                temporaryParagraph = temporaryParagraphs[1];
                temporaryParagraphRange = temporaryParagraph.Range.Duplicate;
                var text = temporary.Content.Text ?? string.Empty;
                Console.WriteLine(
                    $"[ROLLBACK RESIDUAL FIXTURE BEFORE] shape={temporaryShapeRange.Start}:{temporaryShapeRange.End} paragraph={temporaryParagraphRange.Start}:{temporaryParagraphRange.End} bridgeIndex={text.IndexOf(bridge, StringComparison.Ordinal)} textLength={text.Length}");
            }
            finally
            {
                Release(temporaryParagraphRange);
                Release(temporaryParagraph);
                Release(temporaryParagraphs);
                Release(temporaryShapeRange);
            }

            WordFormulaService.RemoveResidualFormatConversionBridgeAfterRollback(
                temporary,
                FormulaOleContract.NativeOleMode,
                target);

            AssertEqual(1, CountVisualTeXNativeOleShapes(temporary),
                "Rollback bridge cleanup removed or duplicated the restored VisualTeX OLE source.");
            AssertTrue(
                (temporary.Content.Text ?? string.Empty).IndexOf(bridge, StringComparison.Ordinal) < 0,
                "Rollback bridge cleanup left the transaction-local LaTeX bridge beside the restored source formula.");

            var output = Path.Combine(
                artifactRoot,
                "Document1-Rollback-Residual-Cleanup.docx");
            temporary.SaveAs2(output, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                $"[ROLLBACK RESIDUAL FIXTURE] Copied the real residual from '{source.Name}', removed only the adjacent bridge, and preserved one VisualTeX OLE. Artifact={output}");
        }
        finally
        {
            Release(temporaryShape);
            Release(destination);
            Release(copy);
            Release(probe);
            Release(sourceParagraphRange);
            Release(sourceParagraph);
            Release(sourceParagraphs);
            Release(sourceShapeRange);
            Release(sourceShape);
            if (temporary is not null)
            {
                try { temporary.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(temporary);
            Release(source);
            Release(application);
            ForceComCleanup();
        }
    }
}
