using System.Runtime.InteropServices;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOleCopyEditAcceptance(string artifactRoot)
    {
        Word.Application? application = null;
        Word.Document? sourceDocument = null;
        Word.Document? temporaryDocument = null;
        Word.InlineShapes? sourceShapes = null;
        Word.InlineShape? sourceShape = null;
        Word.Range? sourceRange = null;
        Word.InlineShape? pastedShape = null;
        Word.Range? pastedRange = null;
        try
        {
            application = Marshal.GetActiveObject("Word.Application") as Word.Application
                ?? throw new InvalidOperationException(
                    "The copy-edit acceptance requires an active Word instance containing a VisualTeX OLE formula.");
            sourceDocument = application.ActiveDocument
                ?? throw new InvalidOperationException("Word has no active document.");
            sourceShapes = sourceDocument.InlineShapes;
            FormulaMetadata? sourceMetadata = null;
            for (var index = 1; index <= sourceShapes.Count; index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate = sourceShapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(candidate)) continue;
                    var metadata = WordFormulaMetadataReader.TryRead(candidate);
                    if (metadata is null
                        || metadata.Numbered
                        || !string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
                        continue;
                    sourceMetadata = metadata;
                    sourceShape = candidate;
                    candidate = null;
                    break;
                }
                finally { Release(candidate); }
            }
            if (sourceShape is null || sourceMetadata is null)
                throw new InvalidOperationException(
                    "The active Word document does not contain a readable VisualTeX native OLE formula.");

            sourceRange = sourceShape.Range;
            sourceRange.Copy();

            temporaryDocument = application.Documents.Add();
            temporaryDocument.Activate();
            application.Selection.Paste();
            if (temporaryDocument.InlineShapes.Count != 1)
                throw new InvalidOperationException(
                    $"Word copy/paste created {temporaryDocument.InlineShapes.Count} inline shapes instead of one.");

            pastedShape = temporaryDocument.InlineShapes[1];
            if (!WordFormulaMetadataReader.IsNativeOle(pastedShape))
                throw new InvalidOperationException(
                    "The pasted formula no longer has the VisualTeX native OLE ProgID.");
            pastedRange = pastedShape.Range;
            pastedRange.Select();

            var service = new WordFormulaService(application);
            var selected = service.ReadSelection();
            if (selected.Metadata is null
                || string.IsNullOrWhiteSpace(selected.FormulaId)
                || !string.Equals(
                    selected.ObjectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The pasted VisualTeX OLE object was not recognized as an editable VisualTeX formula.");
            }
            if (!string.Equals(
                    selected.FormulaId,
                    sourceMetadata.FormulaId,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "Word copy/paste unexpectedly changed the embedded VisualTeX formula metadata.");
            }
            if (!WordDoubleClickRouting.ShouldOpenVisualTeX(selected))
                throw new InvalidOperationException(
                    "The pasted VisualTeX OLE selection would not be routed to the VisualTeX editor.");

            var firstFormulaId = selected.FormulaId!;
            pastedRange.Copy();
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            application.Selection.Paste();
            if (temporaryDocument.InlineShapes.Count != 2)
                throw new InvalidOperationException(
                    $"The duplicate identity probe expected two pasted formulas, actual count: {temporaryDocument.InlineShapes.Count}.");

            Release(pastedRange);
            pastedRange = null;
            Release(pastedShape);
            pastedShape = temporaryDocument.InlineShapes[2];
            pastedRange = pastedShape.Range;
            pastedRange.Select();
            var copiedSelection = service.ReadSelection();
            if (copiedSelection.Metadata is null
                || string.IsNullOrWhiteSpace(copiedSelection.FormulaId)
                || !string.Equals(
                    copiedSelection.ObjectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The second pasted VisualTeX OLE object was not recognized as editable.");
            }
            if (string.Equals(
                    copiedSelection.FormulaId,
                    firstFormulaId,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The second pasted VisualTeX OLE formula kept the first formula's FormulaId.");
            }
            var copiedMetadata = WordFormulaMetadataReader.TryRead(pastedShape)
                ?? throw new InvalidOperationException(
                    "The re-keyed pasted VisualTeX OLE formula no longer exposes metadata.");
            if (!string.Equals(
                    copiedMetadata.FormulaId,
                    copiedSelection.FormulaId,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The copied OLE object's effective FormulaId was not preserved after identity repair.");
            }
            if (!WordDoubleClickRouting.ShouldOpenVisualTeX(copiedSelection))
                throw new InvalidOperationException(
                    "The re-keyed pasted VisualTeX OLE selection would not be routed to the VisualTeX editor.");

            Word.InlineShape? firstShapeAfterRekey = null;
            try
            {
                firstShapeAfterRekey = temporaryDocument.InlineShapes[1];
                var firstMetadataAfterRekey = WordFormulaMetadataReader.TryRead(firstShapeAfterRekey)
                    ?? throw new InvalidOperationException(
                        "The first pasted VisualTeX OLE formula became unreadable after re-keying its copy.");
                if (!string.Equals(
                        firstMetadataAfterRekey.FormulaId,
                        firstFormulaId,
                        StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "Re-keying the copied formula changed the first formula's FormulaId.");
                }

                var firstWidthBefore = firstShapeAfterRekey.Width;
                var secondWidthBefore = pastedShape.Width;
                var sourceFontSize = FormulaFontSize.ResolveSemanticFontSize(copiedSelection.Metadata);
                var requestedFontSize = FormulaFontSize.Normalize(
                    sourceFontSize >= 70 ? sourceFontSize - 2 : sourceFontSize + 2);
                service.SetSelectedFormulaFontSize(requestedFontSize);

                Release(firstShapeAfterRekey);
                firstShapeAfterRekey = temporaryDocument.InlineShapes[1];
                Release(pastedShape);
                pastedShape = temporaryDocument.InlineShapes[2];
                if (Math.Abs(firstShapeAfterRekey.Width - firstWidthBefore) > 0.1f)
                    throw new InvalidOperationException(
                        "Changing the copied formula's font size modified the first formula instead.");
                if (Math.Abs(pastedShape.Width - secondWidthBefore) <= 0.1f)
                    throw new InvalidOperationException(
                        "Changing the copied formula's font size did not target the copied formula.");
            }
            finally { Release(firstShapeAfterRekey); }

            Console.WriteLine(
                "Word VisualTeX OLE copy/edit acceptance passed: freshly pasted dormant OLE objects "
                + "are activated on first read, a same-document copy receives its own durable FormulaId, "
                + "and direct formatting targets only the copy.");

            RunNumberedOleCopyIdentityProbe(application, sourceDocument);
            RunOmmlCopyIdentityProbe(application, sourceDocument);
        }
        finally
        {
            Release(pastedRange);
            Release(pastedShape);
            if (temporaryDocument is not null)
            {
                try { temporaryDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(temporaryDocument);
            if (sourceDocument is not null)
            {
                try { sourceDocument.Activate(); } catch { }
            }
            Release(sourceRange);
            Release(sourceShape);
            Release(sourceShapes);
            Release(sourceDocument);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunNumberedOleCopyIdentityProbe(
        Word.Application application,
        Word.Document sourceDocument)
    {
        Word.InlineShapes? sourceShapes = null;
        Word.InlineShape? sourceShape = null;
        Word.Range? sourceRange = null;
        Word.Range? sourceOwnerRange = null;
        Word.Document? temporaryDocument = null;
        Word.InlineShape? firstShape = null;
        Word.InlineShape? secondShape = null;
        Word.Range? firstRange = null;
        Word.Range? secondRange = null;
        Word.Range? firstOwnerRange = null;
        try
        {
            sourceDocument.Activate();
            sourceShapes = sourceDocument.InlineShapes;
            FormulaMetadata? sourceMetadata = null;
            for (var index = 1; index <= sourceShapes.Count; index++)
            {
                Word.InlineShape? candidate = null;
                Word.Range? candidateRange = null;
                try
                {
                    candidate = sourceShapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(candidate)) continue;
                    var metadata = WordFormulaMetadataReader.TryRead(candidate);
                    if (metadata is null
                        || !metadata.Numbered
                        || !string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
                        continue;
                    candidateRange = candidate.Range;
                    sourceMetadata = metadata;
                    sourceShape = candidate;
                    candidate = null;
                    break;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }
            if (sourceShape is null || sourceMetadata is null)
            {
                Console.WriteLine(
                    "Word numbered OLE copy identity probe skipped: the active source document has no numbered VisualTeX OLE formula.");
                return;
            }

            sourceRange = sourceShape.Range;
            sourceOwnerRange = WordEquationNumbering.FindNumberingOwnerRange(
                    sourceDocument,
                    sourceMetadata.FormulaId)
                ?? throw new InvalidOperationException(
                    "The numbered VisualTeX source has no table-or-tab numbering owner.");
            sourceOwnerRange.Copy();

            temporaryDocument = application.Documents.Add();
            temporaryDocument.Activate();
            application.Selection.Paste();
            if (temporaryDocument.InlineShapes.Count != 1)
                throw new InvalidOperationException(
                    $"The numbered-copy probe expected one pasted formula, actual count: {temporaryDocument.InlineShapes.Count}.");
            if (temporaryDocument.Tables.Count != 0)
                throw new InvalidOperationException(
                    "The pasted numbered VisualTeX OLE unexpectedly reverted to a table host.");

            firstShape = temporaryDocument.InlineShapes[1];
            firstRange = firstShape.Range;
            firstRange.Select();
            var service = new WordFormulaService(application);
            var firstSelection = service.ReadSelection();
            if (firstSelection.Metadata is null
                || !firstSelection.Metadata.Numbered
                || string.IsNullOrWhiteSpace(firstSelection.FormulaId))
                throw new InvalidOperationException(
                    "The first pasted numbered VisualTeX OLE formula was not recognized.");
            var firstFormulaId = firstSelection.FormulaId!;

            firstOwnerRange = WordEquationNumbering.FindNumberingOwnerRange(
                    temporaryDocument,
                    firstFormulaId)
                ?? throw new InvalidOperationException(
                    "The first pasted numbered formula has no tab-paragraph owner.");
            firstOwnerRange.Copy();
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            application.Selection.Paste();
            if (temporaryDocument.InlineShapes.Count != 2)
                throw new InvalidOperationException(
                    $"The numbered-copy probe expected two formulas after duplication, actual count: {temporaryDocument.InlineShapes.Count}.");
            if (temporaryDocument.Tables.Count != 0)
                throw new InvalidOperationException(
                    "Duplicating the numbered VisualTeX formula created a legacy table host.");

            secondShape = temporaryDocument.InlineShapes[2];
            secondRange = secondShape.Range;
            secondRange.Select();
            var secondSelection = service.ReadSelection();
            if (secondSelection.Metadata is null
                || !secondSelection.Metadata.Numbered
                || string.IsNullOrWhiteSpace(secondSelection.FormulaId))
                throw new InvalidOperationException(
                    "The copied numbered VisualTeX OLE formula was not recognized.");
            if (string.Equals(
                    firstFormulaId,
                    secondSelection.FormulaId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "The copied numbered VisualTeX OLE formula kept the original FormulaId.");
            var secondFormulaId = secondSelection.FormulaId!;

            if (!WordEquationNumbering.HasCompleteFormulaNumberingArtifacts(
                    temporaryDocument,
                    firstFormulaId))
                throw new InvalidOperationException(
                    "The first numbered formula lost its numbering bookmarks after copying.");
            if (!WordEquationNumbering.HasCompleteFormulaNumberingArtifacts(
                    temporaryDocument,
                    secondSelection.FormulaId!))
                throw new InvalidOperationException(
                    "The copied numbered formula did not receive complete numbering bookmarks.");
            Release(firstRange);
            firstRange = null;
            Release(firstShape);
            firstShape = FindVisualTeXOleByFormulaIdForNumberToggle(
                temporaryDocument,
                firstFormulaId);
            firstRange = firstShape.Range;
            Release(secondRange);
            secondRange = null;
            Release(secondShape);
            secondShape = FindVisualTeXOleByFormulaIdForNumberToggle(
                temporaryDocument,
                secondFormulaId);
            secondRange = secondShape.Range;

            if (!WordEquationNumbering.FormulaRangeOwnsNumberingArtifacts(
                    temporaryDocument,
                    firstRange,
                    firstFormulaId))
                throw new InvalidOperationException(
                    "The original numbered formula no longer owns its numbering artifacts.");
            if (!WordEquationNumbering.FormulaRangeOwnsNumberingArtifacts(
                    temporaryDocument,
                    secondRange,
                    secondFormulaId))
                throw new InvalidOperationException(
                    "The copied numbered formula does not own its new numbering artifacts.");

            AssertVisualTeXNumberedTabHost(
                temporaryDocument,
                firstFormulaId,
                updateReference: true,
                context: "original numbered OLE after copy repair");
            AssertVisualTeXNumberedTabHost(
                temporaryDocument,
                secondFormulaId,
                updateReference: true,
                context: "duplicated numbered OLE after copy repair");

            Console.WriteLine(
                "Word numbered VisualTeX OLE copy identity probe passed: the duplicated tab-paragraph formula "
                + "received an independent FormulaId and complete numbering bookmarks without stealing the original anchors.");
        }
        finally
        {
            Release(firstOwnerRange);
            Release(secondRange);
            Release(firstRange);
            Release(secondShape);
            Release(firstShape);
            if (temporaryDocument is not null)
            {
                try { temporaryDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(temporaryDocument);
            Release(sourceOwnerRange);
            Release(sourceRange);
            Release(sourceShape);
            Release(sourceShapes);
            try { sourceDocument.Activate(); } catch { }
        }
    }

    private static void RunOmmlCopyIdentityProbe(
        Word.Application application,
        Word.Document sourceDocument)
    {
        Word.Bookmark? sourceBookmark = null;
        Word.Range? sourceRange = null;
        Word.Document? temporaryDocument = null;
        Word.Range? firstRange = null;
        Word.Bookmark? firstBookmark = null;
        Word.Range? firstManagedRange = null;
        Word.Range? secondRange = null;
        Word.Bookmark? secondBookmark = null;
        try
        {
            sourceDocument.Activate();
            var sourceFormulaId = WordOmmlFormulaStore
                .BookmarkedFormulaIds(sourceDocument)
                .FirstOrDefault();
            if (string.IsNullOrWhiteSpace(sourceFormulaId))
            {
                Console.WriteLine(
                    "Word OMML copy identity probe skipped: the active source document has no VisualTeX-managed OMML formula.");
                return;
            }

            sourceBookmark = WordOmmlFormulaStore.FindByFormulaId(
                sourceDocument,
                sourceFormulaId!);
            if (sourceBookmark is null)
                throw new InvalidOperationException(
                    "The source VisualTeX OMML bookmark disappeared before the copy probe.");
            sourceRange = WordOmmlFormulaStore.GetEquationRange(sourceBookmark);
            sourceRange.Copy();

            temporaryDocument = application.Documents.Add();
            temporaryDocument.Activate();
            application.Selection.Paste();
            if (temporaryDocument.OMaths.Count != 1)
                throw new InvalidOperationException(
                    $"The OMML copy probe expected one pasted OMath, actual count: {temporaryDocument.OMaths.Count}.");

            firstRange = temporaryDocument.OMaths[1].Range;
            firstRange.Select();
            var service = new WordFormulaService(application);
            var firstSelection = service.ReadSelection();
            if (firstSelection.Metadata is null
                || string.IsNullOrWhiteSpace(firstSelection.FormulaId)
                || !string.Equals(
                    firstSelection.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "The copied VisualTeX OMML formula was not recognized as editable.");
            if (string.Equals(
                    firstSelection.FormulaId,
                    sourceFormulaId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "The copied OMML formula reused the source document FormulaId.");

            var firstFormulaId = firstSelection.FormulaId!;
            firstBookmark = WordOmmlFormulaStore.FindByFormulaId(
                temporaryDocument,
                firstFormulaId);
            if (firstBookmark is null)
                throw new InvalidOperationException(
                    "The copied OMML formula was recognized but did not receive a durable VisualTeX bookmark.");

            var secondRead = service.ReadSelection();
            if (!string.Equals(
                    secondRead.FormulaId,
                    firstFormulaId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "The adopted OMML formula changed FormulaId on its second read.");

            var sourceFontSize = FormulaFontSize.ResolveSemanticFontSize(firstSelection.Metadata);
            var requestedFontSize = FormulaFontSize.Normalize(
                sourceFontSize >= 70 ? sourceFontSize - 2 : sourceFontSize + 2);
            service.SetSelectedFormulaFontSize(requestedFontSize);
            var afterFontSize = service.ReadSelection();
            if (!string.Equals(
                    afterFontSize.FormulaId,
                    firstFormulaId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "Formatting the adopted OMML formula changed its FormulaId.");

            Release(firstManagedRange);
            firstManagedRange = WordOmmlFormulaStore.GetEquationRange(firstBookmark);
            firstManagedRange.Copy();
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            application.Selection.Paste();
            if (temporaryDocument.OMaths.Count != 2)
                throw new InvalidOperationException(
                    $"The OMML duplicate probe expected two equations, actual count: {temporaryDocument.OMaths.Count}.");

            secondRange = temporaryDocument.OMaths[2].Range;
            secondRange.Select();
            var copiedAgain = service.ReadSelection();
            if (copiedAgain.Metadata is null
                || string.IsNullOrWhiteSpace(copiedAgain.FormulaId))
                throw new InvalidOperationException(
                    "The second OMML copy was not recognized as editable.");
            if (string.Equals(
                    copiedAgain.FormulaId,
                    firstFormulaId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "The second OMML copy kept the first copied formula's FormulaId.");
            secondBookmark = WordOmmlFormulaStore.FindByFormulaId(
                temporaryDocument,
                copiedAgain.FormulaId!);
            if (secondBookmark is null)
                throw new InvalidOperationException(
                    "The second OMML copy did not receive a durable VisualTeX bookmark.");

            Console.WriteLine(
                "Word VisualTeX OMML copy identity probe passed: a bookmark-free pasted OMath "
                + "was adopted with a new durable FormulaId, stayed stable across reread/formatting, "
                + "and a second copy received another independent identity.");
        }
        finally
        {
            Release(secondBookmark);
            Release(secondRange);
            Release(firstManagedRange);
            Release(firstBookmark);
            Release(firstRange);
            if (temporaryDocument is not null)
            {
                try { temporaryDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(temporaryDocument);
            Release(sourceRange);
            Release(sourceBookmark);
            try { sourceDocument.Activate(); } catch { }
        }
    }
}
