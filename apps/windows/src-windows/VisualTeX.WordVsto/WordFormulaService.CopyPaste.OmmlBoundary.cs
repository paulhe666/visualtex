using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    private bool TryRepairMergedCopiedOmml(Document document, WordFormulaCopySnapshot snapshot)
    {
        var copiedLength = snapshot.SourceEnd - snapshot.SourceStart;
        if (copiedLength <= 0 || string.IsNullOrEmpty(snapshot.Metadata.NativeOmmlFingerprint)) return false;
        Selection? selection = null;
        Range? selected = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? range = null;
        try
        {
            selection = _application.Selection;
            selected = selection.Range;
            if (selected.Start != selected.End || selected.StoryType != WdStoryType.wdMainTextStory) return false;
            paragraphs = selected.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            maths = paragraphRange.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                Release(range); range = null;
                Release(math); math = maths[index];
                if (math.Type != WdOMathType.wdOMathInline) continue;
                range = math.Range;
                if (range.End - range.Start <= copiedLength || selected.Start < range.Start || selected.Start > range.End) continue;
                // Only the start/end of a whole copied equation qualifies. A paste
                // into the middle of an existing equation remains native Word input.
                if ((selected.Start == range.Start || selected.Start == range.Start + copiedLength)
                    && TrySplitCopiedOmmlBoundary(document, range, snapshot, copiedLength, copyFirst: true)) return true;
                if ((selected.Start == range.End || selected.Start == range.End - copiedLength)
                    && TrySplitCopiedOmmlBoundary(document, range, snapshot, copiedLength, copyFirst: false)) return true;
            }
            return false;
        }
        finally
        {
            Release(range); Release(math); Release(maths); Release(paragraphRange);
            Release(paragraph); Release(paragraphs); Release(selected); Release(selection);
        }
    }

    private bool TrySplitCopiedOmmlBoundary(
        Document document, Range merged, WordFormulaCopySnapshot snapshot, int copiedLength, bool copyFirst)
    {
        Bookmark? existing = null;
        Range? first = null;
        Range? second = null;
        Range? inserted = null;
        OMaths? insertedMaths = null;
        OMath? firstMath = null;
        OMath? secondMath = null;
        Bookmark? rebound = null;
        Bookmark? copiedBookmark = null;
        WordLocalEditSnapshot? rollback = null;
        try
        {
            existing = WordOmmlFormulaStore.FindAtRange(document, merged);
            if (existing is null) return false;
            var owner = WordOmmlFormulaStore.TryRead(document, existing);
            if (owner is null || owner.DisplayMode != "inline" || owner.Numbered
                || string.IsNullOrEmpty(owner.NativeOmmlFingerprint)) return false;
            var start = merged.Start;
            var end = merged.End;
            var boundary = copyFirst ? start + copiedLength : end - copiedLength;
            first = document.Range(start, boundary);
            second = document.Range(boundary, end);
            var firstXml = first.WordOpenXML;
            var secondXml = second.WordOpenXML;
            var firstFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(firstXml);
            var secondFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(secondXml);
            var copiedFingerprint = copyFirst ? firstFingerprint : secondFingerprint;
            var ownerFingerprint = copyFirst ? secondFingerprint : firstFingerprint;
            if (!string.Equals(copiedFingerprint, snapshot.Metadata.NativeOmmlFingerprint, StringComparison.Ordinal)
                || !string.Equals(ownerFingerprint, owner.NativeOmmlFingerprint, StringComparison.Ordinal)) return false;

            XNamespace p = "http://schemas.microsoft.com/office/2006/xmlPackage";
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            XElement? Equation(XDocument package)
            {
                var part = package.Root?.Elements(p + "part")
                    .SingleOrDefault(item => (string?)item.Attribute(p + "name") == "/word/document.xml");
                var equations = part?.Descendants(m + "oMath").ToList();
                return equations?.Count == 1 ? equations[0] : null;
            }
            var package = XDocument.Parse(merged.WordOpenXML, LoadOptions.PreserveWhitespace);
            var currentEquation = Equation(package);
            var firstEquation = Equation(XDocument.Parse(firstXml, LoadOptions.PreserveWhitespace));
            var secondEquation = Equation(XDocument.Parse(secondXml, LoadOptions.PreserveWhitespace));
            if (currentEquation is null || firstEquation is null || secondEquation is null
                || currentEquation.Ancestors(m + "oMathPara").Any()) return false;
            // An ordinary zero-width Word run keeps the two OMaths independent
            // without a visible gap, paragraph break, or a change to either formula.
            currentEquation.ReplaceWith(
                new XElement(firstEquation),
                new XElement(w + "r", new XElement(w + "t", "\u200C")),
                new XElement(secondEquation));
            var copied = CloneForPastedFormula(snapshot.Metadata);
            var beforeMathCount = snapshot.KnownOmmlCount;
            var beforeParagraphCount = ReadDocumentParagraphCount(document);
            rollback = new WordLocalEditSnapshot(document, merged, owner.FormulaId);
            UndoRecord? undoRecord = null;
            var undoEnded = false;
            try
            {
                undoRecord = BeginUndoRecord("VisualTeX Separate Copied Inline OMML");
                merged.InsertXML(package.ToString(SaveOptions.DisableFormatting));
                inserted = document.Range(start, end + 1);
                insertedMaths = inserted.OMaths;
                if (insertedMaths.Count != 2 || ReadDocumentParagraphCount(document) != beforeParagraphCount)
                    throw new InvalidDataException("Word did not preserve two adjacent inline equations in the original paragraph.");
                firstMath = insertedMaths[1]; secondMath = insertedMaths[2];
                Release(first); first = firstMath.Range;
                Release(second); second = secondMath.Range;
                if (first.Start != start || second.End != end + 1
                    || firstMath.Type != WdOMathType.wdOMathInline || secondMath.Type != WdOMathType.wdOMathInline
                    || WordOmmlConverter.ComputeOmmlFingerprint(first.WordOpenXML) != firstFingerprint
                    || WordOmmlConverter.ComputeOmmlFingerprint(second.WordOpenXML) != secondFingerprint)
                    throw new InvalidDataException("The adjacent copy boundary changed equation content or layout.");
                var ownerRange = copyFirst ? second : first;
                var copyRange = copyFirst ? first : second;
                rebound = WordOmmlFormulaStore.Wrap(document, ownerRange, owner, replaceExisting: true);
                WordOmmlNativeSource.StampFingerprintFromResolvedRange(copied, copyRange);
                copiedBookmark = WordOmmlFormulaStore.Wrap(document, copyRange, copied, replaceExisting: true);
                WordOmmlFormulaStore.Save(document, copied);
                ReadFormulaObjectCounts(document, out _, out var afterMathCount);
                if (afterMathCount != beforeMathCount + 1)
                    throw new InvalidDataException("Separating an adjacent copied equation changed an unrelated OMath count.");
                MoveCaretOutsideInlineOmml(copyRange);
                WordDoubleClickHook.TraceMessage(
                    $"copy-paste-repaired mode=wordOmml adjacentSplit=True formulaId={copied.FormulaId} "
                    + $"ownerId={owner.FormulaId} range={copyRange.Start}:{copyRange.End}");
                return true;
            }
            catch
            {
                try { copiedBookmark?.Delete(); } catch { }
                try { WordOmmlFormulaStore.Delete(document, copied.FormulaId); } catch { }
                if (!undoEnded)
                {
                    EndUndoRecord(undoRecord);
                    undoEnded = true;
                }
                rollback.RestoreOmml(document, owner.FormulaId, owner);
                throw;
            }
            finally
            {
                if (!undoEnded) EndUndoRecord(undoRecord);
                Release(undoRecord);
            }
        }
        finally
        {
            Release(copiedBookmark); Release(rebound); Release(firstMath); Release(secondMath);
            Release(insertedMaths); Release(inserted); Release(second); Release(first); Release(existing);
        }
    }
}
