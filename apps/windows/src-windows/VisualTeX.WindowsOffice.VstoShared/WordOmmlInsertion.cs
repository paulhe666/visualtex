using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using Application = Microsoft.Office.Interop.Word.Application;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordOmmlConverter;
using static VisualTeX.WordVsto.WordOmmlBatchSources;
using static VisualTeX.WordVsto.WordOmmlDocuments;
using static VisualTeX.WordVsto.WordOmmlFingerprints;

namespace VisualTeX.WordVsto;

/// <summary>
/// Putting OMML into a Word document: inserting and replacing formulas in their paragraph, paragraph-mark guards,
/// the inline scratch document, bookmarked files and XML, and finding the math just inserted.
/// </summary>
internal static class WordOmmlInsertion
{
    private const string InlineScratchPlaceholder = "\uE001";
    private static readonly object InlineScratchLock = new();
    private static Document? _inlineScratchDocument;
    private static bool _inlineScratchPrimed;

    internal static Range Insert(
        Application application,
        Document targetDocument,
        Range insertionRange,
        string mathMl,
        bool display,
        out string sourceFingerprint,
        bool includeLeadingTab = false,
        bool replaceTarget = false,
        Func<string, string>? transformOmml = null,
        string? mathFontName = null)
    {
        var omml = TransformMathMlToOmml(mathMl);
        if (transformOmml is not null)
            omml = transformOmml(omml);
        sourceFingerprint = ComputeOmmlFingerprint(omml);
        // OMML Font choice: the chosen font becomes the document math font first.
        mathFontName = WordOmmlMathFontPolicy.ForWrite(targetDocument, mathFontName);
        var normalizedMathFontName = NormalizeMathFontName(mathFontName);
        var tempPath = CreateTemporaryDocx(
            omml,
            includeLeadingTab: display && includeLeadingTab,
            forceInline: !display,
            mathFontName: normalizedMathFontName);
        // Import the bookmarked formula range directly from the DOCX. This keeps
        // the native OMML fidelity of the old path without opening and closing a
        // second hidden Word document for every edit.
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_FAST_OMML_INSERT"),
                "0",
                StringComparison.Ordinal))
        {
            try
            {
                var scratchInlineReplacement = !display
                    && replaceTarget
                    && ShouldUseScratchDocumentForInlineReplacement(
                        targetDocument,
                        insertionRange);
                Range imported;
                if (scratchInlineReplacement)
                {
                    imported = InsertBookmarkedFileThroughScratchDocument(
                        application,
                        targetDocument,
                        insertionRange,
                        tempPath,
                        normalizedMathFontName,
                        omml,
                        out var materializedFingerprint);
                    sourceFingerprint = materializedFingerprint;
                }
                else
                {
                    imported = InsertBookmarkedFile(
                        targetDocument,
                        insertionRange,
                        tempPath,
                        display,
                        replaceTarget);
                }
                try { File.Delete(tempPath); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                return imported;
            }
            catch (Exception error)
            {
                if (string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                    Console.WriteLine(
                        $"    [perf] WordOmmlConverter InsertFile fallback: {error.Message}");
                // Fall through to the proven formatted-text transfer path.
            }
        }
        Document? sourceDocument = null;
        OMaths? sourceMaths = null;
        OMath? sourceMath = null;
        Range? sourceRange = null;
        OMath? insertedMath = null;
        Range? result = null;
        try
        {
            sourceDocument = application.Documents.Open(
                FileName: tempPath,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            sourceMaths = sourceDocument.OMaths;
            if (sourceMaths.Count != 1)
                throw new InvalidDataException(
                    "The temporary OMML document did not contain exactly one equation.");
            sourceMath = sourceMaths[1];
            if (display && includeLeadingTab)
            {
                var paragraph = sourceMath.Range.Paragraphs[1];
                try
                {
                    sourceRange = paragraph.Range.Duplicate;
                    sourceRange.End = Math.Max(sourceRange.Start, sourceRange.End - 1);
                }
                finally { Release(paragraph); }
            }
            else
            {
                sourceRange = sourceMath.Range;
            }

            if (!replaceTarget)
                insertionRange.Collapse(WdCollapseDirection.wdCollapseStart);
            var insertionStart = insertionRange.Start;
            insertionRange.FormattedText = sourceRange.FormattedText;
            insertedMath = FindMathAtPosition(
                    targetDocument,
                    insertionStart,
                    insertionRange.End)
                ?? throw new InvalidOperationException(
                    "Word did not materialize the inserted OMML equation.");
            insertedMath.Type = display
                ? WdOMathType.wdOMathDisplay
                : WdOMathType.wdOMathInline;
            // sourceRange already comes from a fully built native OMML tree.
            // Re-running BuildUp asks Word to parse professional OMML as if it
            // were linear equation text and can introduce dotted placeholder
            // slots, especially around matrices and nested scripts.
            result = insertedMath.Range.Duplicate;
            RemoveImportedFormulaBookmark(targetDocument, result);
            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(result);
            Release(insertedMath);
            Release(sourceRange);
            Release(sourceMath);
            Release(sourceMaths);
            if (sourceDocument is not null)
            {
                try { sourceDocument.Close(WdSaveOptions.wdDoNotSaveChanges); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            }
            Release(sourceDocument);
            try { File.Delete(tempPath); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
        }
    }

    private static bool ShouldUseScratchDocumentForInlineReplacement(
        Document targetDocument,
        Range insertionRange)
    {
        var configured = Environment.GetEnvironmentVariable(
            "VISUALTEX_DIRECT_INLINE_OMML_INSERT");
        if (string.Equals(configured, "0", StringComparison.Ordinal))
            return true;
        if (string.Equals(configured, "1", StringComparison.Ordinal))
            return false;

        Range? content = null;
        try
        {
            content = targetDocument.Content;
            // InsertFile is fastest for a small local document, but Word walks
            // and repaginates the complete tail when importing near the front of
            // a formula-heavy document. Materialize the OMML in the cached tiny
            // scratch document and transfer its already-built FormattedText when
            // the target has a substantial tail. This keeps the same professional
            // OMath tree while making the edit cost local instead of proportional
            // to every later equation.
            return content.End - insertionRange.End >= 256;
        }
        catch
        {
            return false;
        }
        finally { Release(content); }
    }

    internal static bool TryPrimeInlineScratchDocumentForLargeTail(
        Application application,
        Document targetDocument,
        Range equationRange,
        string capturedWordOpenXml,
        string? mathFontName)
    {
        if (!ShouldUseScratchDocumentForInlineReplacement(
                targetDocument,
                equationRange))
            return false;

        lock (InlineScratchLock)
        {
            Range? scratchContent = null;
            Range? scratchTarget = null;
            OMaths? maths = null;
            OMath? scratchMath = null;
            Range? scratchFormula = null;
            Document? transferWarmDocument = null;
            Range? transferWarmContent = null;
            Range? transferWarmTarget = null;
            OMaths? transferWarmMaths = null;
            var tempPath = string.Empty;
            try
            {
                var scratch = GetOrCreateInlineScratchDocument(
                    application,
                    NormalizeMathFontName(mathFontName));
                if (_inlineScratchPrimed) return true;

                var omml = ExtractSingleOMath(capturedWordOpenXml);
                tempPath = CreateTemporaryDocx(
                    omml,
                    includeLeadingTab: false,
                    forceInline: true,
                    mathFontName: NormalizeMathFontName(mathFontName));
                scratchContent = scratch.Content;
                var scratchStart = scratchContent.Start;
                scratchContent.Text = "L" + InlineScratchPlaceholder + "R";
                Release(scratchContent);
                scratchContent = null;
                scratchTarget = scratch.Range(
                    scratchStart + 1,
                    scratchStart + 1 + InlineScratchPlaceholder.Length);
                scratchTarget.InsertFile(
                    FileName: tempPath,
                    Range: FormulaBookmarkName,
                    ConfirmConversions: false,
                    Link: false,
                    Attachment: false);
                maths = scratch.OMaths;
                if (maths.Count != 1)
                    throw new InvalidDataException(
                        "The warmed OMML scratch document did not contain exactly one equation.");
                scratchMath = maths[1];
                scratchFormula = scratchMath.Range;

                // The first cross-document FormattedText assignment in a Word
                // process pays an additional OMath/style initialization cost.
                // Pay it while the editor is opening (where the health/session
                // work already has latency headroom), not after the user presses
                // Apply. This transfer is isolated to disposable hidden documents
                // and never mutates the user's target.
                transferWarmDocument = application.Documents.Add(Visible: false);
                ApplyDocumentMathFont(
                    transferWarmDocument,
                    NormalizeMathFontName(mathFontName));
                transferWarmContent = transferWarmDocument.Content;
                var transferStart = transferWarmContent.Start;
                transferWarmContent.Text = "L" + InlineScratchPlaceholder + "R";
                Release(transferWarmContent);
                transferWarmContent = null;
                transferWarmTarget = transferWarmDocument.Range(
                    transferStart + 1,
                    transferStart + 1 + InlineScratchPlaceholder.Length);
                transferWarmTarget.FormattedText = scratchFormula.FormattedText;
                transferWarmMaths = transferWarmDocument.OMaths;
                if (transferWarmMaths.Count != 1)
                    throw new InvalidDataException(
                        "The warmed OMML FormattedText transfer did not contain exactly one equation.");
                _inlineScratchPrimed = true;
                try { scratch.Saved = true; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                try { transferWarmDocument.Saved = true; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                return true;
            }
            catch
            {
                _inlineScratchPrimed = false;
                return false;
            }
            finally
            {
                Release(transferWarmMaths);
                Release(transferWarmTarget);
                Release(transferWarmContent);
                Release(scratchFormula);
                Release(scratchMath);
                Release(maths);
                Release(scratchTarget);
                Release(scratchContent);
                if (transferWarmDocument is not null)
                {
                    try { transferWarmDocument.Saved = true; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                    try
                    {
                        transferWarmDocument.Close(
                            WdSaveOptions.wdDoNotSaveChanges);
                    }
                    catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                }
                Release(transferWarmDocument);
                try { targetDocument.Activate(); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                if (!string.IsNullOrWhiteSpace(tempPath))
                {
                    try { File.Delete(tempPath); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                }
            }
        }
    }

    private static Range ReplaceExistingOmmlNodesInParagraph(
        Document targetDocument,
        Range targetRange,
        string omml,
        bool display,
        string? mathFontName)
    {
        XNamespace math =
            MathNamespace;

        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        OMaths? paragraphMaths = null;
        OMath? paragraphMath = null;
        Range? exact = null;
        OMath? insertedMath = null;
        Range? result = null;
        string? paragraphMarkGuardText = null;
        var guardSearchStart = targetRange.Start;
        try
        {
            // Same paragraph-mark guard as ReplaceWithPreparedOmmlInParagraph:
            // paragraph-level InsertXML can drop this paragraph's mark and join
            // it to the next paragraph or a directly following table.
            var targetStart = targetRange.Start;
            var targetEnd = targetRange.End;
            paragraphMarkGuardText =
                TryCreateParagraphMarkGuard(
                    targetDocument,
                    targetRange);
            targetRange.SetRange(
                targetStart,
                targetEnd);

            paragraphs =
                targetRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "The adjacent inline OMath merge spans more than one Word paragraph.");
            paragraph =
                paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            paragraphMaths =
                paragraphRange.OMaths;

            var targetOrdinals =
                new List<int>();
            for (var index = 1;
                 index <= paragraphMaths.Count;
                 index++)
            {
                Release(exact);
                exact = null;
                Release(paragraphMath);
                paragraphMath =
                    paragraphMaths[index];
                exact =
                    paragraphMath.Range.Duplicate;
                if (exact.Start >= targetRange.Start
                    && exact.End <= targetRange.End)
                {
                    targetOrdinals.Add(
                        index - 1);
                }
            }

            if (targetOrdinals.Count == 0)
                throw new InvalidDataException(
                    "The requested OMath replacement range contains no complete Word OMath.");
            for (var index = 1;
                 index < targetOrdinals.Count;
                 index++)
            {
                if (targetOrdinals[index]
                    != targetOrdinals[index - 1] + 1)
                    throw new InvalidDataException(
                        "The requested OMath replacement range contains non-contiguous Word OMath hosts.");
            }

            var paragraphXml =
                XDocument.Parse(
                    paragraphRange.WordOpenXML,
                    LoadOptions.PreserveWhitespace);
            var xmlMaths =
                paragraphXml
                    .Descendants(
                        math + "oMath")
                    .Where(element =>
                        !element.Ancestors(
                            math + "oMath")
                            .Any())
                    .ToArray();
            if (xmlMaths.Length != paragraphMaths.Count)
            {
                // Track Changes keeps deleted or moved-away copies in the XML;
                // the ordinal mapping holds once those are left out.
                XNamespace word = WordNamespace;
                xmlMaths = xmlMaths
                    .Where(element => !element.Ancestors().Any(ancestor =>
                        ancestor.Name == word + "del" || ancestor.Name == word + "moveFrom"))
                    .ToArray();
            }
            if (xmlMaths.Length !=
                paragraphMaths.Count)
                throw new InvalidDataException(
                    "Word paragraph OMath COM/XML inventories disagree during native inline merge.");

            var replacement =
                XElement.Parse(
                    ApplyExplicitTransferMathFont(
                        ExtractSingleOMath(
                            omml),
                        NormalizeMathFontName(
                            mathFontName)),
                    LoadOptions.PreserveWhitespace);

            var firstOrdinal =
                targetOrdinals[0];
            xmlMaths[firstOrdinal]
                .ReplaceWith(
                    replacement);
            for (var index =
                     targetOrdinals.Count - 1;
                 index >= 1;
                 index--)
            {
                xmlMaths[targetOrdinals[index]]
                    .Remove();
            }

            var insertionStart =
                targetRange.Start;
            WpsXmlInsertion.Insert(
                paragraphRange,
                paragraphXml.ToString(
                    SaveOptions.DisableFormatting));

            insertedMath =
                FindMathAtPosition(
                    targetDocument,
                    insertionStart,
                    ResolveCurrentParagraphEnd(
                        targetDocument,
                        insertionStart))
                ?? throw new InvalidOperationException(
                    "Word did not materialize the paragraph-local OMath node replacement.");

            // Restore the paragraph's own mark before Word applies the type.
            if (paragraphMarkGuardText is not null)
            {
                RemoveParagraphMarkGuard(
                    targetDocument,
                    insertedMath,
                    paragraphMarkGuardText);
                paragraphMarkGuardText = null;
                Release(insertedMath);
                insertedMath =
                    FindMathAtPosition(
                        targetDocument,
                        insertionStart,
                        ResolveCurrentParagraphEnd(
                            targetDocument,
                            insertionStart))
                    ?? throw new InvalidOperationException(
                        "Word lost the OMath while restoring its paragraph mark.");
            }

            var targetType =
                display
                    ? WdOMathType.wdOMathDisplay
                    : WdOMathType.wdOMathInline;
            if (insertedMath.Type !=
                targetType)
                insertedMath.Type =
                    targetType;
            // The imported oMath is already structured. Do not reparse its runs.

            result =
                insertedMath.Range.Duplicate;
            var returned =
                result;
            result = null;
            return returned;
        }
        finally
        {
            if (paragraphMarkGuardText is not null)
                TryDeleteParagraphMarkGuardBestEffort(
                    targetDocument,
                    guardSearchStart,
                    paragraphMarkGuardText);
            Release(result);
            Release(insertedMath);
            Release(exact);
            Release(paragraphMath);
            Release(paragraphMaths);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool IsOnlyParagraphInTableCell(
        Range paragraphRange)
    {
        Cells? cells = null;
        Cell? cell = null;
        Range? cellRange = null;
        Paragraphs? paragraphs = null;
        try
        {
            if (!Convert.ToBoolean(
                    paragraphRange.get_Information(
                        WdInformation.wdWithInTable)))
                return false;

            cells =
                paragraphRange.Cells;
            if (cells.Count != 1)
                return false;

            cell =
                cells[1];
            cellRange =
                cell.Range.Duplicate;
            paragraphs =
                cellRange.Paragraphs;
            return paragraphs.Count == 1;
        }
        finally
        {
            Release(paragraphs);
            Release(cellRange);
            Release(cell);
            Release(cells);
        }
    }

    private static bool TryRemoveRedundantEmptyTableParagraph(
        Document document,
        OMath insertedMath)
    {
        Range? mathRange = null;
        Cells? cells = null;
        Cell? cell = null;
        Range? cellRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? firstParagraph = null;
        Paragraph? secondParagraph = null;
        Range? firstRange = null;
        Range? secondRange = null;
        Tables? secondTables = null;
        InlineShapes? secondShapes = null;
        OMaths? secondMaths = null;
        Fields? secondFields = null;
        Frames? secondFrames = null;
        Range? redundantBreak = null;
        try
        {
            mathRange =
                insertedMath.Range.Duplicate;
            if (!Convert.ToBoolean(
                    mathRange.get_Information(
                        WdInformation.wdWithInTable)))
                return false;

            cells =
                mathRange.Cells;
            if (cells.Count != 1)
                return false;

            cell =
                cells[1];
            cellRange =
                cell.Range.Duplicate;
            paragraphs =
                cellRange.Paragraphs;
            if (paragraphs.Count != 2)
                return false;

            firstParagraph = paragraphs[1];
            secondParagraph = paragraphs[2];
            firstRange =
                firstParagraph.Range.Duplicate;
            secondRange =
                secondParagraph.Range.Duplicate;

            // The OMath must remain in the original first paragraph. The second
            // paragraph must be completely empty except for Word's structural
            // paragraph/cell terminators; otherwise it is user-authored content
            // and must never be removed.
            if (mathRange.Start < firstRange.Start
                || mathRange.End > firstRange.End)
                return false;

            secondTables =
                secondRange.Tables;
            secondShapes =
                secondRange.InlineShapes;
            secondMaths =
                secondRange.OMaths;
            // Including the final cell atom makes Word.Range.Fields expose
            // the whole cell's fields, even when this tail paragraph is empty.
            // Use the established content-only local range, not that aggregate.
            secondFields = WordFormulaHost.GetLocalFields(secondRange);
            secondFrames =
                secondRange.Frames;
            // Range.Tables includes the owning user table for any range
            // inside a cell. Only an additional nested table makes this tail
            // paragraph non-empty.
            if (secondTables.Count > 1
                || secondShapes.Count != 0
                || secondMaths.Count != 0
                || secondFields.Count != 0
                || secondFrames.Count != 0)
                return false;

            var secondText =
                secondRange.Text
                ?? string.Empty;
            var hasParagraphTerminator = false;
            foreach (var character in secondText)
            {
                if (character == (char)13)
                {
                    hasParagraphTerminator = true;
                    continue;
                }

                if (character == (char)7
                    || character == (char)10
                    || character == (char)9
                    || character == (char)11
                    || char.IsWhiteSpace(character))
                    continue;

                return false;
            }
            if (!hasParagraphTerminator)
                return false;

            // Delete only the paragraph mark separating the original paragraph
            // from Word's newly-created empty tail paragraph. The cell's own
            // terminal paragraph/cell mark lives in the second paragraph and is
            // intentionally preserved.
            if (firstRange.End <= firstRange.Start)
                return false;
            redundantBreak =
                document.Range(
                    firstRange.End - 1,
                    firstRange.End);
            var breakText =
                redundantBreak.Text
                ?? string.Empty;
            if (breakText.Length != 1
                || breakText[0] != (char)13)
                return false;

            redundantBreak.Delete();
            return true;
        }
        finally
        {
            Release(redundantBreak);
            Release(secondFrames);
            Release(secondFields);
            Release(secondMaths);
            Release(secondShapes);
            Release(secondTables);
            Release(secondRange);
            Release(firstRange);
            Release(secondParagraph);
            Release(firstParagraph);
            Release(paragraphs);
            Release(cellRange);
            Release(cell);
            Release(cells);
            Release(mathRange);
        }
    }

    internal static Range ReplaceWithPreparedOmmlInParagraph(
        Document targetDocument,
        Range targetRange,
        string omml,
        bool display,
        string? mathFontName = null,
        bool allowWordNativeInlinePromotion = false)
    {
        if (targetDocument is null)
            throw new ArgumentNullException(nameof(targetDocument));
        if (targetRange is null)
            throw new ArgumentNullException(nameof(targetRange));
        if (string.IsNullOrWhiteSpace(omml))
            throw new InvalidDataException(
                "The prepared paragraph-local OMML replacement is empty.");
        // OMML Font choice: the chosen font becomes the document math font first.
        mathFontName = WordOmmlMathFontPolicy.ForWrite(targetDocument, mathFontName);

        OMaths? existingTargetMaths = null;
        OMath? existingTargetMath = null;
        Range? existingTargetExact = null;
        try
        {
            existingTargetMaths =
                targetRange.OMaths;
            if (existingTargetMaths.Count == 1)
            {
                existingTargetMath =
                    existingTargetMaths[1];
                existingTargetExact =
                    existingTargetMath.Range.Duplicate;
                if (existingTargetExact.StoryType ==
                        targetRange.StoryType
                    && existingTargetExact.Start ==
                        targetRange.Start
                    && existingTargetExact.End ==
                        targetRange.End)
                {
                    return ReplaceExistingOmmlNodesInParagraph(
                        targetDocument,
                        targetRange,
                        omml,
                        display,
                        mathFontName);
                }
            }
        }
        finally
        {
            Release(existingTargetExact);
            Release(existingTargetMath);
            Release(existingTargetMaths);
        }

        XNamespace word =
            WordNamespace;
        XNamespace math =
            MathNamespace;
        var markerText =
            "VTX"
            + Guid.NewGuid().ToString("N");
        var insertionStart =
            targetRange.Start;

        Range? markerRange = null;
        Range? followingEmptyParagraphGuard = null;
        string? followingEmptyParagraphGuardText = null;
        string? paragraphMarkGuardText = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        OMath? insertedMath = null;
        Range? result = null;
        try
        {
            // Word's paragraph-level InsertXML can absorb an immediately
            // following empty paragraph even when the inserted OMath is inline.
            // Preserve that user-authored paragraph with a temporary ordinary
            // text witness for every paragraph-local OMML replacement, then
            // remove only the witness after Word has materialized the equation.
            followingEmptyParagraphGuard =
                TryCreateFollowingEmptyParagraphGuard(
                    targetDocument,
                    targetRange,
                    out followingEmptyParagraphGuardText);

            markerRange =
                targetRange.Duplicate;
            markerRange.Text =
                markerText;
            markerRange.SetRange(
                insertionStart,
                insertionStart + markerText.Length);

            // Paragraph-level InsertXML can drop the replaced paragraph's mark:
            // it then joins the next paragraph (a Normal style without paragraph
            // properties) or the first cell of a directly following table.
            // Separate them with a temporary non-empty paragraph that owns the
            // original paragraph mark.
            paragraphMarkGuardText =
                TryCreateParagraphMarkGuard(
                    targetDocument,
                    markerRange);
            markerRange.SetRange(
                insertionStart,
                insertionStart + markerText.Length);

            paragraphs =
                markerRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "The OMML insertion marker does not belong to exactly one Word paragraph.");
            paragraph =
                paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            var singleParagraphTableCellBeforeInsert =
                IsOnlyParagraphInTableCell(
                    paragraphRange);

            var paragraphXml =
                XDocument.Parse(
                    paragraphRange.WordOpenXML,
                    LoadOptions.PreserveWhitespace);
            var ordinaryMarkerNodes =
                paragraphXml
                    .Descendants(
                        word + "t")
                    .Where(node =>
                        node.Value.IndexOf(
                            markerText,
                            StringComparison.Ordinal)
                        >= 0)
                    .ToArray();
            var mathMarkerNodes =
                paragraphXml
                    .Descendants(
                        math + "t")
                    .Where(node =>
                        node.Value.IndexOf(
                            markerText,
                            StringComparison.Ordinal)
                        >= 0)
                    .ToArray();
            if (ordinaryMarkerNodes.Length
                    + mathMarkerNodes.Length
                != 1)
                throw new InvalidDataException(
                    $"Word paragraph XML contains {ordinaryMarkerNodes.Length} ordinary and {mathMarkerNodes.Length} mathematical copies of the OMML insertion marker.");

            var markerNodeForParagraph =
                ordinaryMarkerNodes.Length == 1
                    ? ordinaryMarkerNodes[0]
                    : mathMarkerNodes[0];
            var targetParagraphXml =
                markerNodeForParagraph
                    .Ancestors(word + "p")
                    .SingleOrDefault()
                ?? throw new InvalidDataException(
                    "The OMML insertion marker has no unique containing Word paragraph.");

            var equation =
                XElement.Parse(
                    ApplyExplicitTransferMathFont(
                        ExtractSingleOMath(
                            omml),
                        NormalizeMathFontName(
                            mathFontName)),
                    LoadOptions.PreserveWhitespace);

            if (ordinaryMarkerNodes.Length == 1)
            {
                // Word owns any run-level wrapper around the insertion site.
                // With Track Changes enabled the marker is normally emitted as
                // w:p/w:ins/w:r/w:t, not as a direct paragraph child. Replace the
                // marker run in place so revision/content-control/hyperlink
                // provenance remains exactly where Word put it.
                WordOmmlInsertionXml.ReplaceOrdinaryMarkerRun(
                    ordinaryMarkerNodes[0],
                    markerText,
                    equation,
                    targetParagraphXml);
            }
            else
            {
                // Word may absorb a collapsed insertion directly beside an
                // inline OMath into that OMath. In that native case the marker is
                // an m:t node, not ordinary w:t. Splice the new equation's
                // children into the existing m:oMath exactly where Word placed
                // the marker, preserving Word's single merged OMath topology.
                var markerNode =
                    mathMarkerNodes[0];
                var sourceMathRun =
                    markerNode.Parent;
                if (sourceMathRun is null
                    || sourceMathRun.Name !=
                        math + "r"
                    || sourceMathRun.Parent is null
                    || sourceMathRun.Parent.Name !=
                        math + "oMath")
                    throw new InvalidDataException(
                        "The absorbed OMML insertion marker is not one direct math run inside an inline OMath.");

                var markerIndex =
                    markerNode.Value.IndexOf(
                        markerText,
                        StringComparison.Ordinal);
                var prefix =
                    markerNode.Value.Substring(
                        0,
                        markerIndex);
                var suffix =
                    markerNode.Value.Substring(
                        markerIndex + markerText.Length);

                XElement CreateMathSiblingRun(
                    string text)
                {
                    var run =
                        new XElement(
                            math + "r");
                    var runProperties =
                        sourceMathRun.Element(
                            math + "rPr");
                    if (runProperties is not null)
                        run.Add(
                            new XElement(
                                runProperties));
                    run.Add(
                        new XElement(
                            math + "t",
                            text));
                    return run;
                }

                var replacements =
                    new List<object>();
                if (!string.IsNullOrEmpty(
                        prefix))
                    replacements.Add(
                        CreateMathSiblingRun(
                            prefix));
                foreach (var child in
                         equation.Nodes())
                {
                    if (child is XElement element)
                        replacements.Add(
                            new XElement(
                                element));
                    else
                        replacements.Add(
                            child.ToString());
                }
                if (!string.IsNullOrEmpty(
                        suffix))
                    replacements.Add(
                        CreateMathSiblingRun(
                            suffix));

                sourceMathRun.ReplaceWith(
                    replacements.ToArray());
            }

            if (Convert.ToBoolean(
                    paragraphRange.get_Information(
                        WdInformation.wdWithInTable)))
            {
                // A paragraph range inside a table serializes with its containing
                // w:tbl/w:tr/w:tc context. Re-inserting that complete package into
                // the same cell manufactures a nested table. Keep all package
                // parts, but reduce /word/document.xml to the exact edited w:p so
                // Word replaces only the existing paragraph in-place.
                var mainBody =
                    WordDocumentXml.GetRequiredMainDocumentBody(
                        paragraphXml);
                var standaloneParagraph =
                    new XElement(
                        targetParagraphXml);
                var sectionProperties =
                    mainBody
                        .Elements(word + "sectPr")
                        .LastOrDefault();
                var standaloneSectionProperties =
                    sectionProperties is null
                        ? null
                        : new XElement(sectionProperties);
                mainBody.RemoveNodes();
                mainBody.Add(standaloneParagraph);
                if (standaloneSectionProperties is not null)
                    mainBody.Add(
                        standaloneSectionProperties);
            }

            WpsXmlInsertion.Insert(
                paragraphRange,
                paragraphXml.ToString(
                    SaveOptions.DisableFormatting));

            insertedMath =
                FindMathAtPosition(
                    targetDocument,
                    insertionStart,
                    ResolveCurrentParagraphEnd(
                        targetDocument,
                        insertionStart))
                ?? throw new InvalidOperationException(
                    "Word did not materialize the paragraph-local OMML replacement.");

            // Restore the paragraph's own mark before Word applies the display
            // type: if Word joined the import into the guard paragraph, the
            // equation still shares a paragraph with the guard token here.
            // The mark guard also sits between the formula paragraph and any
            // empty-paragraph witness, so it is removed before that witness.
            if (paragraphMarkGuardText is not null)
            {
                RemoveParagraphMarkGuard(
                    targetDocument,
                    insertedMath,
                    paragraphMarkGuardText);
                paragraphMarkGuardText = null;
                Release(insertedMath);
                insertedMath =
                    FindMathAtPosition(
                        targetDocument,
                        insertionStart,
                        ResolveCurrentParagraphEnd(
                            targetDocument,
                            insertionStart))
                    ?? throw new InvalidOperationException(
                        "Word lost the OMath while restoring its paragraph mark.");
            }

            if (display)
            {
                if (insertedMath.Type !=
                    WdOMathType.wdOMathDisplay)
                    insertedMath.Type =
                        WdOMathType.wdOMathDisplay;
            }
            else if (!allowWordNativeInlinePromotion
                     && insertedMath.Type !=
                         WdOMathType.wdOMathInline)
            {
                insertedMath.Type =
                    WdOMathType.wdOMathInline;
            }

            // A freshly inserted inline-requested OMML equation follows Word's
            // own final OMath type. In particular, a lone equation at the start
            // of a paragraph may be promoted to wdOMathDisplay by Word itself.
            // Do not add a trailing sentinel or force that native promotion back
            // to inline. Explicit display requests remain authoritative above.
            // Preserve the parsed OMML tree: BuildUp can reinterpret fixed runs.

            if (singleParagraphTableCellBeforeInsert
                && TryRemoveRedundantEmptyTableParagraph(
                    targetDocument,
                    insertedMath))
            {
                Release(insertedMath);
                insertedMath =
                    FindMathAtPosition(
                        targetDocument,
                        insertionStart,
                        ResolveCurrentParagraphEnd(
                            targetDocument,
                            insertionStart))
                    ?? throw new InvalidOperationException(
                        "Word lost the table-cell OMath after removing its redundant empty paragraph.");
            }

            if (followingEmptyParagraphGuard is not null)
            {
                RemoveFollowingEmptyParagraphGuard(
                    targetDocument,
                    insertedMath,
                    followingEmptyParagraphGuardText
                    ?? throw new InvalidDataException(
                        "The OMML paragraph-boundary guard lost its expected text."));
                Release(followingEmptyParagraphGuard);
                followingEmptyParagraphGuard = null;
                followingEmptyParagraphGuardText = null;
            }

            result =
                insertedMath.Range.Duplicate;
            var returned =
                result;
            result = null;
            return returned;
        }
        finally
        {
            if (followingEmptyParagraphGuard is not null
                && !string.IsNullOrEmpty(
                    followingEmptyParagraphGuardText))
            {
                try
                {
                    if (string.Equals(
                            followingEmptyParagraphGuard.Text,
                            followingEmptyParagraphGuardText,
                            StringComparison.Ordinal))
                        followingEmptyParagraphGuard.Text = string.Empty;
                }
                catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            }
            if (paragraphMarkGuardText is not null)
                TryDeleteParagraphMarkGuardBestEffort(
                    targetDocument,
                    insertionStart,
                    paragraphMarkGuardText);
            Release(followingEmptyParagraphGuard);
            Release(result);
            Release(insertedMath);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(markerRange);
        }
    }

    // Splits "\r" + token in front of the target paragraph's mark. Paragraph-level
    // InsertXML does not reliably keep the replaced paragraph's own mark: when
    // the document's Normal style carries no paragraph properties (English Word
    // templates, Pandoc/python-docx documents) Word drops the imported mark and
    // joins the paragraph to the next one, and a directly following table
    // absorbs the paragraph into its first cell. The token paragraph keeps the
    // original mark between the replaced paragraph and whatever follows it.
    // Returns the token, or null for a table-cell or final paragraph.
    private static string? TryCreateParagraphMarkGuard(
        Document document,
        Range targetRange)
    {
        Paragraphs? targetParagraphs = null;
        Paragraph? targetParagraph = null;
        Range? targetParagraphRange = null;
        Tables? ownTables = null;
        Range? insertion = null;
        Range? guardRange = null;
        try
        {
            targetParagraphs = targetRange.Paragraphs;
            if (targetParagraphs.Count != 1)
                return null;
            targetParagraph = targetParagraphs[1];
            targetParagraphRange =
                targetParagraph.Range.Duplicate;
            // Paragraphs inside a cell are re-inserted as a standalone w:p.
            ownTables = targetParagraphRange.Tables;
            if (ownTables.Count != 0)
                return null;
            if (targetParagraphRange.End >=
                document.Content.End)
                return null;

            var token =
                "VTOT"
                + Guid.NewGuid().ToString("N");
            var markStart = targetParagraphRange.End - 1;
            insertion = document.Range(
                markStart,
                markStart);
            insertion.Text = "\r" + token;
            guardRange = document.Range(
                markStart,
                markStart + 1 + token.Length);
            if (!string.Equals(
                    guardRange.Text,
                    "\r" + token,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Word did not retain the temporary paragraph that keeps the OMML paragraph's own mark.");
            return token;
        }
        finally
        {
            Release(guardRange);
            Release(insertion);
            Release(ownTables);
            Release(targetParagraphRange);
            Release(targetParagraph);
            Release(targetParagraphs);
        }
    }

    // Restores the replaced paragraph's original mark. If Word kept the imported
    // mark, the token paragraph follows the formula paragraph: delete that mark
    // together with the token. If Word joined the import into the token
    // paragraph, the token ends the formula paragraph: delete the token only.
    // Either way the content ends on the original paragraph mark.
    private static void RemoveParagraphMarkGuard(
        Document document,
        OMath insertedMath,
        string token)
    {
        Range? mathRange = null;
        Paragraphs? mathParagraphs = null;
        Paragraph? mathParagraph = null;
        Range? mathParagraphRange = null;
        Range? guard = null;
        Find? guardFind = null;
        Range? removal = null;
        try
        {
            mathRange = insertedMath.Range.Duplicate;
            mathParagraphs = mathRange.Paragraphs;
            if (mathParagraphs.Count != 1)
                throw new InvalidDataException(
                    "The inserted OMML no longer belongs to exactly one Word paragraph while restoring its paragraph mark.");
            mathParagraph = mathParagraphs[1];
            mathParagraphRange =
                mathParagraph.Range.Duplicate;

            guard = document.Range(
                mathParagraphRange.Start,
                Math.Min(
                    document.Content.End,
                    mathParagraphRange.End + token.Length + 1));
            guardFind = guard.Find;
            guardFind.ClearFormatting();
            guardFind.Text = token;
            guardFind.Forward = true;
            guardFind.Wrap = WdFindWrap.wdFindStop;
            if (!guardFind.Execute()
                || !string.Equals(
                    guard.Text,
                    token,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Word moved or changed the temporary paragraph that keeps the OMML paragraph's own mark.");

            if (guard.Start == mathParagraphRange.End)
            {
                // Word kept the imported mark.
                removal = document.Range(
                    mathParagraphRange.End - 1,
                    guard.End);
                if (!string.Equals(
                        removal.Text,
                        "\r" + token,
                        StringComparison.Ordinal))
                    throw new InvalidDataException(
                        "The OMML paragraph does not end directly before its paragraph-mark guard.");
                removal.Text = string.Empty;
            }
            else if (guard.End == mathParagraphRange.End - 1)
            {
                // Word joined the import into the guard paragraph.
                guard.Text = string.Empty;
            }
            else
            {
                throw new InvalidDataException(
                    "Word moved the OMML paragraph-mark guard away from the paragraph boundary.");
            }
        }
        finally
        {
            Release(removal);
            Release(guardFind);
            Release(guard);
            Release(mathParagraphRange);
            Release(mathParagraph);
            Release(mathParagraphs);
            Release(mathRange);
        }
    }

    private static void TryDeleteParagraphMarkGuardBestEffort(
        Document document,
        int searchStart,
        string token)
    {
        Range? guard = null;
        Find? guardFind = null;
        Range? removal = null;
        try
        {
            guard = document.Range(
                Math.Min(searchStart, document.Content.End),
                document.Content.End);
            guardFind = guard.Find;
            guardFind.ClearFormatting();
            guardFind.Text = token;
            guardFind.Forward = true;
            guardFind.Wrap = WdFindWrap.wdFindStop;
            if (!guardFind.Execute()
                || guard.Start < 1)
                return;
            removal = document.Range(
                guard.Start - 1,
                guard.End);
            if (string.Equals(
                    removal.Text,
                    "\r" + token,
                    StringComparison.Ordinal))
                removal.Text = string.Empty;
            else
                guard.Text = string.Empty;
        }
        catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
        finally
        {
            Release(removal);
            Release(guardFind);
            Release(guard);
        }
    }

    private static Range? TryCreateFollowingEmptyParagraphGuard(
        Document document,
        Range targetRange,
        out string? guardText)
    {
        guardText = null;
        Paragraphs? targetParagraphs = null;
        Paragraph? targetParagraph = null;
        Range? targetParagraphRange = null;
        Range? nextProbe = null;
        Paragraphs? nextParagraphs = null;
        Paragraph? nextParagraph = null;
        Range? nextRange = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Frames? frames = null;
        Tables? tables = null;
        Range? insertion = null;
        Range? guardRange = null;
        try
        {
            targetParagraphs = targetRange.Paragraphs;
            if (targetParagraphs.Count != 1)
                return null;
            targetParagraph = targetParagraphs[1];
            targetParagraphRange =
                targetParagraph.Range.Duplicate;
            if (targetParagraphRange.End >=
                document.Content.End)
                return null;

            nextProbe = document.Range(
                targetParagraphRange.End,
                Math.Min(
                    document.Content.End,
                    targetParagraphRange.End + 1));
            if (Convert.ToBoolean(
                    nextProbe.get_Information(
                        WdInformation.wdWithInTable)))
                return null;
            nextParagraphs = nextProbe.Paragraphs;
            if (nextParagraphs.Count != 1)
                return null;
            nextParagraph = nextParagraphs[1];
            nextRange = nextParagraph.Range.Duplicate;
            if (nextRange.Start !=
                targetParagraphRange.End)
                return null;

            tables = nextRange.Tables;
            shapes = nextRange.InlineShapes;
            maths = nextRange.OMaths;
            fields = nextRange.Fields;
            frames = nextRange.Frames;
            if (tables.Count != 0
                || shapes.Count != 0
                || maths.Count != 0
                || fields.Count != 0
                || frames.Count != 0)
                return null;

            var text =
                nextRange.Text
                ?? string.Empty;
            foreach (var character in text)
            {
                if (character is '\r' or '\n'
                    or '\t' or '\v' or '\a'
                    || char.IsWhiteSpace(character))
                    continue;
                return null;
            }

            var token =
                "VTOB"
                + Guid.NewGuid().ToString("N");
            insertion = document.Range(
                nextRange.Start,
                nextRange.Start);
            insertion.Text = token;
            guardRange = document.Range(
                nextRange.Start,
                nextRange.Start + token.Length);
            if (!string.Equals(
                    guardRange.Text,
                    token,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Word did not retain the temporary OMML paragraph-boundary guard.");

            guardText = token;
            var result = guardRange;
            guardRange = null;
            return result;
        }
        finally
        {
            Release(guardRange);
            Release(insertion);
            Release(tables);
            Release(frames);
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(nextRange);
            Release(nextParagraph);
            Release(nextParagraphs);
            Release(nextProbe);
            Release(targetParagraphRange);
            Release(targetParagraph);
            Release(targetParagraphs);
        }
    }

    private static void RemoveFollowingEmptyParagraphGuard(
        Document document,
        OMath insertedMath,
        string expectedText)
    {
        Range? mathRange = null;
        Paragraphs? mathParagraphs = null;
        Paragraph? mathParagraph = null;
        Range? mathParagraphRange = null;
        Range? nextProbe = null;
        Paragraphs? nextParagraphs = null;
        Paragraph? nextParagraph = null;
        Range? nextRange = null;
        Range? guard = null;
        Find? guardFind = null;
        try
        {
            mathRange = insertedMath.Range.Duplicate;
            mathParagraphs = mathRange.Paragraphs;
            if (mathParagraphs.Count != 1)
                throw new InvalidDataException(
                    "The inserted OMML no longer belongs to exactly one Word paragraph while restoring its boundary.");
            mathParagraph = mathParagraphs[1];
            mathParagraphRange =
                mathParagraph.Range.Duplicate;
            if (mathParagraphRange.End >=
                document.Content.End)
                throw new InvalidDataException(
                    "The following OMML paragraph-boundary guard disappeared at the end of the document.");

            nextProbe = document.Range(
                mathParagraphRange.End,
                Math.Min(
                    document.Content.End,
                    mathParagraphRange.End + 1));
            nextParagraphs = nextProbe.Paragraphs;
            if (nextParagraphs.Count != 1)
                throw new InvalidDataException(
                    "The preserved paragraph after OMML has ambiguous Word paragraph ownership.");
            nextParagraph = nextParagraphs[1];
            nextRange = nextParagraph.Range.Duplicate;
            if (nextRange.Start !=
                mathParagraphRange.End)
                throw new InvalidDataException(
                    "Word moved the preserved paragraph away from the OMML boundary.");
            var nextText = nextRange.Text
                ?? string.Empty;
            if (!nextText.StartsWith(
                    expectedText,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Word moved or changed the temporary OMML paragraph-boundary guard while replacing the equation paragraph.");

            // Track Changes inserts a non-text revision boundary at the start of
            // the paragraph. Paragraph.Text still begins with the guard, while a
            // freshly reconstructed document.Range(start,start+length) begins one
            // story position too early and truncates the final character. Search
            // only inside the already-proven following paragraph and let Word map
            // its own revision boundary to the visible token.
            guard = nextRange.Duplicate;
            guardFind = guard.Find;
            guardFind.ClearFormatting();
            guardFind.Text = expectedText;
            guardFind.Forward = true;
            guardFind.Wrap = WdFindWrap.wdFindStop;
            if (!guardFind.Execute()
                || !string.Equals(
                    guard.Text,
                    expectedText,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The OMML paragraph-boundary guard cannot be resolved inside the preserved paragraph.");
            guard.Text = string.Empty;
        }
        finally
        {
            Release(guardFind);
            Release(guard);
            Release(nextRange);
            Release(nextParagraph);
            Release(nextParagraphs);
            Release(nextProbe);
            Release(mathParagraphRange);
            Release(mathParagraph);
            Release(mathParagraphs);
            Release(mathRange);
        }
    }

    internal static Range ReplaceWithPreparedOmmlDirect(
        Document targetDocument,
        Range targetRange,
        string omml,
        bool display,
        string? mathFontName = null,
        bool allowWordNativeInlinePromotion = false) =>
        ReplaceWithPreparedOmmlDirect(
            targetDocument,
            targetRange,
            omml,
            display,
            mathFontName,
            allowWordNativeInlinePromotion,
            paragraphMarkRunProperties: null);

    internal static Range ReplaceWithPreparedOmmlDirect(
        Document targetDocument,
        Range targetRange,
        string omml,
        bool display,
        string? mathFontName,
        bool allowWordNativeInlinePromotion,
        string? paragraphMarkRunProperties)
    {
        if (targetDocument is null)
            throw new ArgumentNullException(nameof(targetDocument));
        if (targetRange is null)
            throw new ArgumentNullException(nameof(targetRange));
        if (string.IsNullOrWhiteSpace(omml))
            throw new InvalidDataException(
                "The prepared direct OMML replacement is empty.");
        // OMML Font choice: the chosen font becomes the document math font first.
        mathFontName = WordOmmlMathFontPolicy.ForWrite(targetDocument, mathFontName);

        var bookmarkName =
            "VisualTeXDirect"
            + Guid.NewGuid().ToString("N").Substring(0, 16);
        var entry = new BatchEntry(
            bookmarkName,
            ComputeOmmlFingerprint(omml),
            omml,
            31902);
        var wpsInlineSentinel = OfficeHostEnvironment.IsWps && !display;
        string flatOpc;
        using (VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("FormatConversion.OmmlDirect.BuildPackage"))
            flatOpc =
                CreateSingleEquationFlatOpc(
                    entry,
                    NormalizeMathFontName(mathFontName),
                    paragraphMarkRunProperties,
                    trailingWpsSentinel: wpsInlineSentinel);

        Range? target = null;
        Range? content = null;
        Range? rightBefore = null;
        Range? transportBreak = null;
        Range? rightAfterBreak = null;
        OMath? insertedMath = null;
        Range? result = null;
        Bookmarks? bookmarks = null;
        Bookmark? imported = null;
        Range? importedRange = null;
        try
        {
            target = targetRange.Duplicate;
            var insertionStart = target.Start;

            content = targetDocument.Content;
            var rightBoundaryLength = Math.Min(
                32,
                Math.Max(
                    0,
                    content.End - target.End));
            var rightBoundaryBefore = string.Empty;
            if (rightBoundaryLength > 0)
            {
                rightBefore = targetDocument.Range(
                    target.End,
                    target.End + rightBoundaryLength);
                rightBoundaryBefore =
                    rightBefore.Text
                    ?? string.Empty;
            }
            var hasImmediateRightUserText =
                rightBoundaryBefore.Length > 0
                && rightBoundaryBefore[0] != '\r'
                && rightBoundaryBefore[0] != '\u0007';

            using (VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("FormatConversion.OmmlDirect.InsertXml"))
                WpsXmlInsertion.Insert(target, flatOpc);

            using (VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("FormatConversion.OmmlDirect.FindMath"))
                insertedMath = FindMathAtPosition(
                        targetDocument,
                        insertionStart,
                        target.End)
                    ?? throw new InvalidOperationException(
                        "Word did not materialize the direct OMML replacement.");
            if (wpsInlineSentinel)
            {
                RemoveWpsInlineSentinel(insertedMath);
                Release(insertedMath);
                insertedMath = null;
                insertedMath = FindMathAtPosition(
                        targetDocument,
                        insertionStart,
                        Math.Max(insertionStart, target.End))
                    ?? throw new InvalidOperationException(
                        "WPS lost the direct OMML replacement after removing its import separator.");
            }
            var targetType = display
                ? WdOMathType.wdOMathDisplay
                : WdOMathType.wdOMathInline;
            // WPS keeps an equation that is alone in its paragraph display (it adds a
            // space when forced inline); leave that to WPS, as Word's own promotion.
            if (insertedMath.Type != targetType
                && (display || (!allowWordNativeInlinePromotion && !OfficeHostEnvironment.IsWps)))
                insertedMath.Type = targetType;
            // Native inline promotion controls host type, not equation parsing.
            // InsertXML has already materialized the prepared OMML structure.

            result = insertedMath.Range.Duplicate;

            if (hasImmediateRightUserText)
            {
                Release(content);
                content = targetDocument.Content;
                if (result.End < content.End)
                {
                    transportBreak = targetDocument.Range(
                        result.End,
                        Math.Min(
                            content.End,
                            result.End + 1));
                    if (string.Equals(
                            transportBreak.Text,
                            "\r",
                            StringComparison.Ordinal))
                    {
                        var comparableLength = Math.Min(
                            rightBoundaryBefore.Length,
                            Math.Max(
                                0,
                                content.End - (result.End + 1)));
                        if (comparableLength > 0)
                        {
                            rightAfterBreak = targetDocument.Range(
                                result.End + 1,
                                result.End + 1 + comparableLength);
                            var afterText =
                                rightAfterBreak.Text
                                ?? string.Empty;
                            var expected =
                                rightBoundaryBefore.Substring(
                                    0,
                                    comparableLength);
                            if (string.Equals(
                                    afterText,
                                    expected,
                                    StringComparison.Ordinal))
                            {
                                transportBreak.Delete();

                                Release(result);
                                result = null;
                                Release(insertedMath);
                                insertedMath = null;
                                Release(content);
                                content = targetDocument.Content;
                                insertedMath = FindMathAtPosition(
                                        targetDocument,
                                        insertionStart,
                                        Math.Min(
                                            content.End,
                                            insertionStart + 256))
                                    ?? throw new InvalidOperationException(
                                        "Word lost the direct OMML replacement after transport paragraph cleanup.");
                                result =
                                    insertedMath.Range.Duplicate;
                            }
                        }
                    }
                }
            }

            bookmarks = targetDocument.Bookmarks;
            if (bookmarks.Exists(bookmarkName))
            {
                imported = bookmarks[bookmarkName];
                importedRange = imported.Range.Duplicate;
                if (importedRange.StoryType == result.StoryType
                    && importedRange.Start <= result.End
                    && importedRange.End >= result.Start)
                    imported.Delete();
            }

            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(importedRange);
            Release(imported);
            Release(bookmarks);
            Release(result);
            Release(insertedMath);
            Release(rightAfterBreak);
            Release(transportBreak);
            Release(rightBefore);
            Release(content);
            Release(target);
        }
    }

    internal static Range ReplaceWithPreparedOmml(
        Application application,
        Document targetDocument,
        Range targetRange,
        string omml,
        bool display,
        string? mathFontName = null)
    {
        if (application is null) throw new ArgumentNullException(nameof(application));
        if (targetDocument is null) throw new ArgumentNullException(nameof(targetDocument));
        if (targetRange is null) throw new ArgumentNullException(nameof(targetRange));
        if (string.IsNullOrWhiteSpace(omml))
            throw new InvalidDataException("The prepared OMML replacement is empty.");

        // OMML Font choice: the chosen font becomes the document math font first.
        mathFontName = WordOmmlMathFontPolicy.ForWrite(targetDocument, mathFontName);
        var normalizedMathFontName = NormalizeMathFontName(mathFontName);
        var tempPath = CreateTemporaryDocx(
            omml,
            includeLeadingTab: false,
            forceInline: !display,
            mathFontName: normalizedMathFontName);
        Document? sourceDocument = null;
        OMaths? sourceMaths = null;
        OMath? sourceMath = null;
        Range? sourceRange = null;
        Range? target = null;
        OMath? insertedMath = null;
        Range? result = null;
        try
        {
            sourceDocument = application.Documents.Open(
                FileName: tempPath,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            sourceMaths = sourceDocument.OMaths;
            if (sourceMaths.Count != 1)
                throw new InvalidDataException(
                    "The prepared OMML replacement document did not contain exactly one equation.");
            sourceMath = sourceMaths[1];
            sourceRange = sourceMath.Range;

            target = targetRange.Duplicate;
            var insertionStart = target.Start;
            // Word can replace one professional OMath with another through
            // FormattedText without linearizing either equation. This preserves
            // radicals, fractions, matrices and the display-math separators while
            // avoiding the placeholder/BuildUp corruption seen in older builds.
            target.FormattedText = sourceRange.FormattedText;
            insertedMath = FindMathAtPosition(
                    targetDocument,
                    insertionStart,
                    target.End)
                ?? throw new InvalidOperationException(
                    "Word did not materialize the prepared OMML replacement.");
            var targetType = display
                ? WdOMathType.wdOMathDisplay
                : WdOMathType.wdOMathInline;
            if (insertedMath.Type != targetType)
                insertedMath.Type = targetType;
            result = insertedMath.Range.Duplicate;
            RemoveImportedFormulaBookmark(targetDocument, result);
            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(result);
            Release(insertedMath);
            Release(target);
            Release(sourceRange);
            Release(sourceMath);
            Release(sourceMaths);
            if (sourceDocument is not null)
            {
                try { sourceDocument.Close(WdSaveOptions.wdDoNotSaveChanges); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            }
            Release(sourceDocument);
            // Opening the temporary prepared-OMML document can make it Word's
            // ActiveDocument even when Visible=false. Once that document closes,
            // some Word builds expose no ActiveDocument until the caller manually
            // activates the real document; the next editor Apply then fails before
            // it can resolve its source. Restore the supplied target explicitly.
            try { targetDocument.Activate(); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            try { File.Delete(tempPath); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
        }
    }

    private static Range InsertBookmarkedFileThroughScratchDocument(
        Application application,
        Document targetDocument,
        Range insertionRange,
        string filePath,
        string mathFontName,
        string preparedOmml,
        out string materializedFingerprint)
    {
        lock (InlineScratchLock)
            return InsertBookmarkedFileThroughScratchDocumentCore(
                application,
                targetDocument,
                insertionRange,
                filePath,
                mathFontName,
                preparedOmml,
                out materializedFingerprint);
    }

    private static Range InsertBookmarkedFileThroughScratchDocumentCore(
        Application application,
        Document targetDocument,
        Range insertionRange,
        string filePath,
        string mathFontName,
        string preparedOmml,
        out string materializedFingerprint)
    {
        Document? scratchDocument = null;
        Range? scratchInsertion = null;
        OMaths? scratchMaths = null;
        OMath? scratchMath = null;
        Range? scratchFormula = null;
        Range? target = null;
        OMath? insertedMath = null;
        Range? result = null;
        materializedFingerprint = string.Empty;
        try
        {
            scratchDocument = GetOrCreateInlineScratchDocument(
                application,
                mathFontName);
            scratchInsertion = scratchDocument.Content;
            var scratchStart = scratchInsertion.Start;
            scratchInsertion.Text = "L" + InlineScratchPlaceholder + "R";
            Release(scratchInsertion);
            scratchInsertion = scratchDocument.Range(
                scratchStart + 1,
                scratchStart + 1 + InlineScratchPlaceholder.Length);
            // The scratch document itself supplies stable ordinary L/R context.
            // Import only the bookmarked equation, avoiding a full DOCX insertion
            // while still forcing Word to materialize a genuine inline OMath.
            scratchInsertion.InsertFile(
                FileName: filePath,
                Range: FormulaBookmarkName,
                ConfirmConversions: false,
                Link: false,
                Attachment: false);
            scratchMaths = scratchDocument.OMaths;
            if (scratchMaths.Count != 1)
                throw new InvalidDataException(
                    "The OMML scratch document did not contain exactly one equation.");
            scratchMath = scratchMaths[1];
            scratchFormula = scratchMath.Range;
            materializedFingerprint =
                ComputeOmmlFingerprint(scratchFormula.WordOpenXML);

            target = insertionRange.Duplicate;
            if (target.Start >= target.End)
                throw new InvalidOperationException(
                    "Inline OMML replacement requires a non-collapsed placeholder range.");
            var insertionStart = target.Start;
            target.FormattedText = scratchFormula.FormattedText;
            insertedMath = FindMathAtPosition(
                    targetDocument,
                    insertionStart,
                    target.End)
                ?? throw new InvalidOperationException(
                    "Word did not materialize the inline OMML replacement.");
            if (insertedMath.Type != WdOMathType.wdOMathInline)
                insertedMath.Type = WdOMathType.wdOMathInline;
            result = insertedMath.Range.Duplicate;
            RemoveImportedFormulaBookmark(targetDocument, result);
            try { scratchDocument.Saved = true; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(result);
            Release(insertedMath);
            Release(target);
            Release(scratchFormula);
            Release(scratchMath);
            Release(scratchMaths);
            Release(scratchInsertion);
            // The hidden scratch document is intentionally cached for the life
            // of this Word process. Reusing it removes Documents.Add/Close from
            // every inline edit; it is marked Saved so Word can quit silently.
            try { targetDocument.Activate(); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
        }
    }

    private static Document GetOrCreateInlineScratchDocument(
        Application application,
        string mathFontName)
    {
        if (_inlineScratchDocument is not null)
        {
            Application? scratchApplication = null;
            try
            {
                scratchApplication = _inlineScratchDocument.Application;
                if (IsSameComObject(scratchApplication, application))
                {
                    ApplyDocumentMathFont(_inlineScratchDocument, mathFontName);
                    return _inlineScratchDocument;
                }
            }
            catch (Exception ignored)
            {
                // The previous Word process has exited or invalidated its RCW.
                IgnoredErrors.Ignore(ignored);
            }
            finally { Release(scratchApplication); }

            try
            {
                _inlineScratchDocument.Saved = true;
                _inlineScratchDocument.Close(WdSaveOptions.wdDoNotSaveChanges);
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            Release(_inlineScratchDocument);
            _inlineScratchDocument = null;
            _inlineScratchPrimed = false;
        }

        _inlineScratchDocument = application.Documents.Add(Visible: false);
        _inlineScratchPrimed = false;
        ApplyDocumentMathFont(_inlineScratchDocument, mathFontName);
        try { _inlineScratchDocument.Saved = true; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
        return _inlineScratchDocument;
    }

    private static Range InsertBookmarkedFile(
        Document targetDocument,
        Range insertionRange,
        string filePath,
        bool display,
        bool replaceTarget)
    {
        Range? target = null;
        OMath? insertedMath = null;
        Range? result = null;
        try
        {
            target = insertionRange.Duplicate;
            var insertionStart = target.Start;
            if (!replaceTarget)
                target.Collapse(WdCollapseDirection.wdCollapseStart);
            // InsertFile replaces a non-collapsed Range in place. Do not delete
            // the placeholder first: collapsing at its former boundary lets Word
            // move the imported OMath behind the adjacent typing sentinel.
            // In WPS the import blanks the OLE formulas already in the document
            // until they are laid out again.
            WpsOleDisplayRefresh.Mark(targetDocument);
            target.InsertFile(
                FileName: filePath,
                Range: FormulaBookmarkName,
                ConfirmConversions: false,
                Link: false,
                Attachment: false);
            insertedMath = FindMathAtPosition(
                    targetDocument,
                    insertionStart,
                    target.End)
                ?? throw new InvalidOperationException(
                    "Word did not materialize the bookmarked OMML equation.");
            var targetType = display
                ? WdOMathType.wdOMathDisplay
                : WdOMathType.wdOMathInline;
            if (insertedMath.Type != targetType)
                insertedMath.Type = targetType;
            // The imported DOCX already contains professional OMML. Calling
            // BuildUp again reparses an already-built tree and adds substantial
            // latency without changing the equation structure.
            result = insertedMath.Range.Duplicate;
            RemoveImportedFormulaBookmark(targetDocument, result);
            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(result);
            Release(insertedMath);
            Release(target);
        }
    }

    private static void RemoveImportedFormulaBookmark(
        Document targetDocument,
        Range insertedRange)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        try
        {
            bookmarks = targetDocument.Bookmarks;
            if (!bookmarks.Exists(FormulaBookmarkName)) return;
            bookmark = bookmarks[FormulaBookmarkName];
            bookmarkRange = bookmark.Range;
            var overlapsInsertedEquation = bookmarkRange.Start <= insertedRange.End
                && bookmarkRange.End >= insertedRange.Start;
            if (overlapsInsertedEquation) bookmark.Delete();
        }
        catch (Exception ignored)
        {
            // This bookmark exists only to locate the formula inside the temporary
            // DOCX. The native equation remains valid if cleanup is rejected by
            // an older Word build.
            IgnoredErrors.Ignore(ignored);
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static Range InsertXmlDirect(
        Document targetDocument,
        Range insertionRange,
        string omml,
        bool display,
        bool replaceTarget)
    {
        Range? target = null;
        Range? content = null;
        Range? probe = null;
        OMaths? maths = null;
        OMath? selected = null;
        Range? selectedRange = null;
        try
        {
            target = insertionRange.Duplicate;
            if (!replaceTarget)
                target.Collapse(WdCollapseDirection.wdCollapseStart);
            var insertionStart = target.Start;
            WpsXmlInsertion.Insert(target, omml);

            content = targetDocument.Content;
            object probeStart = Math.Max(content.Start, insertionStart - 1);
            object probeEnd = Math.Min(
                content.End,
                Math.Max(target.End + 2, insertionStart + 256));
            probe = targetDocument.Range(ref probeStart, ref probeEnd);
            maths = probe.OMaths;
            var bestDistance = int.MaxValue;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? candidate = null;
                Range? candidateRange = null;
                try
                {
                    candidate = maths[index];
                    candidateRange = candidate.Range;
                    var distance = Math.Abs(candidateRange.Start - insertionStart);
                    if (distance >= bestDistance) continue;
                    Release(selectedRange);
                    Release(selected);
                    selected = candidate;
                    candidate = null;
                    selectedRange = candidateRange;
                    candidateRange = null;
                    bestDistance = distance;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }
            if (selected is null || selectedRange is null)
                throw new InvalidOperationException(
                    "Word did not materialize the directly inserted OMML equation.");
            selected.Type = display
                ? WdOMathType.wdOMathDisplay
                : WdOMathType.wdOMathInline;
            // This is parsed OMML, not linear equation text requiring BuildUp.
            var result = selected.Range.Duplicate;
            return result;
        }
        finally
        {
            Release(selectedRange);
            Release(selected);
            Release(maths);
            Release(probe);
            Release(content);
            Release(target);
        }
    }

    internal static bool ShouldPreferInsertedMathCandidate(
        int distance,
        int span,
        int bestDistance,
        int bestSpan)
    {
        if (distance > 16) return false;
        if (distance < bestDistance) return true;
        if (distance > bestDistance) return false;
        return span > bestSpan;
    }

    private static int ResolveCurrentParagraphEnd(
        Document document,
        int position)
    {
        Range? content = null;
        Range? point = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            content =
                document.Content;
            var safePosition =
                Math.Max(
                    content.Start,
                    Math.Min(
                        position,
                        Math.Max(
                            content.Start,
                            content.End - 1)));
            point =
                document.Range(
                    safePosition,
                    safePosition);
            paragraphs =
                point.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "The rewritten OMML insertion point no longer belongs to exactly one Word paragraph.");

            paragraph =
                paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            return paragraphRange.End;
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(point);
            Release(content);
        }
    }

    internal static OMath? FindMathAtPosition(
        Document document,
        int position,
        int preferredEnd)
    {
        const int rpcCallRejected = unchecked((int)0x80010001);
        const int rpcServerCallRetryLater = unchecked((int)0x8001010A);
        const int maximumAttempts = 40;
        for (var attempt = 0; ; attempt++)
        {
            try
            {
                return FindMathAtPositionCore(document, position, preferredEnd);
            }
            catch (COMException error)
                when ((error.HResult == rpcCallRejected
                        || error.HResult == rpcServerCallRetryLater)
                    && attempt < maximumAttempts - 1)
            {
                // InsertFile/FormattedText can return just before Word has committed
                // the new native equation tree. Retrying the read-only locator is
                // safer than replaying the insertion and cannot duplicate content.
                System.Threading.Thread.Sleep(50);
            }
        }
    }

    private static OMath? FindMathAtPositionCore(
        Document document,
        int position,
        int preferredEnd)
    {
        Range? content = null;
        Range? probe = null;
        OMaths? maths = null;
        OMath? best = null;
        var bestSpan = -1;
        var bestDistance = int.MaxValue;
        try
        {
            content = document.Content;
            object probeStart = Math.Max(content.Start, position - 1);
            object probeEnd = Math.Min(
                content.End,
                Math.Max(preferredEnd + 2, position + 1024));
            probe = document.Range(ref probeStart, ref probeEnd);
            maths = probe.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    // Word exposes nested matrix rows as OMath objects in a local
                    // range. The physical equation nearest the insertion point must
                    // win first; only candidates at the same start distance use the
                    // larger span to prefer the top-level matrix over an inner row.
                    // Prioritizing span before distance can select the next, longer
                    // display equation in an adjacent paragraph.
                    if (range.End <= position || range.Start > preferredEnd + 1) continue;
                    var span = range.End - range.Start;
                    var distance = Math.Abs(range.Start - position);
                    if (!ShouldPreferInsertedMathCandidate(
                            distance,
                            span,
                            bestDistance,
                            bestSpan))
                        continue;
                    Release(best);
                    best = math;
                    math = null;
                    bestSpan = span;
                    bestDistance = distance;
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
            if (best is not null) return best;

            // Defensive fallback for an unusual Word range expansion. Normal
            // insertions resolve through the local probe above, avoiding an
            // O(n) enumeration of every equation in a 100-formula document.
            Release(maths);
            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    if (range.End <= position || range.Start > preferredEnd + 1) continue;
                    var span = range.End - range.Start;
                    var distance = Math.Abs(range.Start - position);
                    if (!ShouldPreferInsertedMathCandidate(
                            distance,
                            span,
                            bestDistance,
                            bestSpan))
                        continue;
                    Release(best);
                    best = math;
                    math = null;
                    bestSpan = span;
                    bestDistance = distance;
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
            return best;
        }
        finally
        {
            Release(maths);
            Release(probe);
            Release(content);
        }
    }

    private static bool IsSameComObject(object left, object right)
    {
        IntPtr leftIdentity = IntPtr.Zero;
        IntPtr rightIdentity = IntPtr.Zero;
        try
        {
            leftIdentity = Marshal.GetIUnknownForObject(left);
            rightIdentity = Marshal.GetIUnknownForObject(right);
            return leftIdentity == rightIdentity;
        }
        catch { return false; }
        finally
        {
            if (rightIdentity != IntPtr.Zero) Marshal.Release(rightIdentity);
            if (leftIdentity != IntPtr.Zero) Marshal.Release(leftIdentity);
        }
    }
}
