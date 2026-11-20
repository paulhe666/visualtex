using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordEquationNumbering;
using static VisualTeX.WordVsto.WordNativeCaptions;
using static VisualTeX.WordVsto.WordNumberedDisplayLayout;
using static VisualTeX.WordVsto.WordNumberedTableDisplay;
using static VisualTeX.WordVsto.WordNumberedTrueDisplay;
using static VisualTeX.WordVsto.WordNumberingFields;
using static VisualTeX.WordVsto.WordNumberingReconcile;

namespace VisualTeX.WordVsto;

/// <summary>
/// The paragraphs around a numbered display formula: leaked caption frames, the typing paragraph after the formula
/// and its spacing.
/// </summary>
internal static class WordNumberedTypingParagraphs
{
    internal static int RepairLeakedNativeCaptionFrames(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));

        var captionNames = new List<string>();
        Bookmarks? inventory = null;
        try
        {
            inventory = document.Bookmarks;
            for (var index = 1; index <= inventory.Count; index++)
            {
                Bookmark? bookmark = null;
                try
                {
                    bookmark = inventory[index];
                    var name = bookmark.Name ?? string.Empty;
                    if (name.StartsWith(NativeCaptionBookmarkPrefix, StringComparison.Ordinal)
                        && name.Length == NativeCaptionBookmarkPrefix.Length + 32)
                        captionNames.Add(name);
                }
                finally { Release(bookmark); }
            }
        }
        finally { Release(inventory); }

        var repaired = 0;
        foreach (var captionName in captionNames)
        {
            Bookmarks? bookmarks = null;
            Bookmark? captionBookmark = null;
            Bookmark? numberBookmark = null;
            Range? captionRange = null;
            Range? numberRange = null;
            Range? trailingRange = null;
            Frames? frames = null;
            Frame? leakedFrame = null;
            Range? splitRange = null;
            Range? repairedCaptionRange = null;
            try
            {
                bookmarks = document.Bookmarks;
                if (!bookmarks.Exists(captionName)) continue;
                var suffix = captionName.Substring(NativeCaptionBookmarkPrefix.Length);
                var numberName = NativeNumberBookmarkPrefix + suffix;
                if (!bookmarks.Exists(numberName)) continue;

                captionBookmark = bookmarks[captionName];
                numberBookmark = bookmarks[numberName];
                captionRange = captionBookmark.Range;
                numberRange = numberBookmark.Range;
                if (numberRange.Start < captionRange.Start
                    || numberRange.End <= numberRange.Start
                    || numberRange.End >= captionRange.End - 1)
                    continue;

                trailingRange = document.Range(numberRange.End, captionRange.End);
                var trailingText = (trailingRange.Text ?? string.Empty)
                    .Replace("\r", string.Empty)
                    .Replace("\a", string.Empty)
                    .Trim();
                if (trailingText.Length == 0 && trailingRange.InlineShapes.Count == 0)
                    continue;

                frames = captionRange.Frames;
                for (var frameIndex = 1; frameIndex <= frames.Count; frameIndex++)
                {
                    Frame? candidate = null;
                    try
                    {
                        candidate = frames[frameIndex];
                        Range? frameRange = null;
                        try
                        {
                            frameRange = candidate.Range;
                            var isVisualTeXClip =
                                candidate.WidthRule == WdFrameSizeRule.wdFrameExact
                                && candidate.HeightRule == WdFrameSizeRule.wdFrameExact
                                && candidate.Width <= 0.5f
                                && candidate.Height <= 0.5f
                                && !candidate.TextWrap
                                && candidate.LockAnchor
                                && frameRange.Start <= numberRange.Start
                                // Word may exclude the paragraph mark itself from a
                                // Frame.Range even when the Frame visually owns the
                                // whole paragraph. Allow exactly that one-character
                                // boundary difference, but no broader overlap.
                                && frameRange.End >= trailingRange.End - 1;
                            if (!isVisualTeXClip) continue;
                            leakedFrame = candidate;
                            candidate = null;
                            break;
                        }
                        finally { Release(frameRange); }
                    }
                    finally { Release(candidate); }
                }
                if (leakedFrame is null) continue;

                var captionStart = captionRange.Start;
                var numberEnd = numberRange.End;
                leakedFrame.Delete();
                Release(leakedFrame);
                leakedFrame = null;

                // The buggy table path inserted SEQ at the first character of the
                // following body paragraph. Split exactly after VTEqNum_ so the
                // original body/formula moves into its own visible paragraph while
                // the native SEQ target remains an independent caption paragraph.
                captionBookmark.Delete();
                Release(captionBookmark);
                captionBookmark = null;
                splitRange = document.Range(numberEnd, numberEnd);
                // InsertParagraphBefore on a collapsed range in the middle of a
                // paragraph inserts a paragraph before the *owner paragraph* on
                // some Office 2019/2021 builds; it does not reliably split at the
                // collapsed coordinate. That leaves the leaked body text inside
                // VTEqCap or moves it ahead of the caption. Insert the paragraph
                // mark as exact story text at VTEqNum.End so the SEQ remains in
                // the first paragraph and every original trailing character moves
                // intact into the new visible body paragraph.
                splitRange.Text = "\r";
                repairedCaptionRange = document.Range(captionStart, numberEnd + 1);
                captionBookmark = bookmarks.Add(captionName, repairedCaptionRange);

                Release(numberRange);
                numberRange = numberBookmark.Range;
                StyleNativeCaption(
                    repairedCaptionRange,
                    numberRange,
                    cleanupLegacyFrames: false);
                repaired++;
            }
            catch (COMException ignored)
            {
                // A malformed legacy bookmark should not make an otherwise valid
                // document unusable. The normal conversion safety checks will still
                // stop before deleting a source if this one could not be repaired.
                IgnoredErrors.Ignore(ignored);
            }
            finally
            {
                Release(repairedCaptionRange);
                Release(splitRange);
                Release(leakedFrame);
                Release(frames);
                Release(trailingRange);
                Release(numberRange);
                Release(captionRange);
                Release(numberBookmark);
                Release(captionBookmark);
                Release(bookmarks);
            }
        }
        return repaired;
    }

    internal static WordCharacterFormatting CaptureNumberedBodyCharacterFormatting(Document document, string formulaId)
    {
        Range? number = null;
        try
        {
            // The number host's paragraph mark retains the source body format.
            // Neither OMath/OLE typography nor the clipped SEQ caption is a body source.
            number = FindVisibleEquationNumberTextRange(document, formulaId)
                ?? throw new InvalidDataException("The numbered formula has no visible number host for body continuation.");
            return WordCharacterFormatting.CaptureParagraphMark(number);
        }
        finally { Release(number); }
    }

    private static bool TryTrimExpandedNativeCaptionBookmarksAfterStyle(
        Document document,
        string formulaId,
        Range captionParagraphRange)
    {
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Bookmark? numberBookmark = null;
        Range? captionBookmarkRange = null;
        Range? numberBookmarkRange = null;
        Range? trailing = null;
        Range? canonicalCaption = null;
        Range? canonicalNumber = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        Tables? tables = null;
        ContentControls? controls = null;
        Frames? frames = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            var numberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName) || !bookmarks.Exists(numberName))
                return false;

            captionBookmark = bookmarks[captionName];
            numberBookmark = bookmarks[numberName];
            captionBookmarkRange = captionBookmark.Range;
            numberBookmarkRange = numberBookmark.Range;
            var overflowEnd = Math.Max(captionBookmarkRange.End, numberBookmarkRange.End);
            if (overflowEnd <= captionParagraphRange.End)
                return false;

            trailing = document.Range(captionParagraphRange.End, overflowEnd);
            var trailingText = trailing.Text ?? string.Empty;
            if (!trailingText.All(character =>
                    char.IsWhiteSpace(character) || character == '\a'))
                return false;
            maths = trailing.OMaths;
            shapes = trailing.InlineShapes;
            fields = trailing.Fields;
            tables = trailing.Tables;
            controls = trailing.ContentControls;
            frames = trailing.Frames;
            if (maths.Count != 0
                || shapes.Count != 0
                || fields.Count != 0
                || tables.Count != 0
                || controls.Count != 0
                || frames.Count != 0)
                return false;

            var canonicalNumberStart = Math.Max(
                captionParagraphRange.Start,
                numberBookmarkRange.Start);
            var canonicalNumberEnd = Math.Min(
                captionParagraphRange.End,
                numberBookmarkRange.End);
            if (canonicalNumberEnd <= canonicalNumberStart)
                throw new InvalidDataException(
                    "The native equation number bookmark no longer intersects its caption paragraph.");

            captionBookmark.Delete();
            numberBookmark.Delete();
            Release(captionBookmark);
            captionBookmark = null;
            Release(numberBookmark);
            numberBookmark = null;

            canonicalCaption = document.Range(
                captionBookmarkRange.Start,
                captionParagraphRange.End);
            canonicalNumber = document.Range(
                canonicalNumberStart,
                canonicalNumberEnd);
            captionBookmark = bookmarks.Add(captionName, canonicalCaption);
            numberBookmark = bookmarks.Add(numberName, canonicalNumber);

            Release(captionBookmarkRange);
            captionBookmarkRange = captionBookmark.Range;
            Release(numberBookmarkRange);
            numberBookmarkRange = numberBookmark.Range;
            if (captionBookmarkRange.End > captionParagraphRange.End
                || numberBookmarkRange.End > captionParagraphRange.End)
                throw new InvalidDataException(
                    "Word re-expanded the native equation caption bookmarks after their final binding.");

            return true;
        }
        finally
        {
            Release(frames);
            Release(controls);
            Release(tables);
            Release(fields);
            Release(shapes);
            Release(maths);
            Release(canonicalNumber);
            Release(canonicalCaption);
            Release(trailing);
            Release(numberBookmarkRange);
            Release(captionBookmarkRange);
            Release(numberBookmark);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    private static bool TryNormalizeExpandedNativeCaptionBeforeTyping(
        Document document,
        string formulaId,
        Range captionRange,
        Range captionParagraphRange)
    {
        if (captionRange.End <= captionParagraphRange.End)
            return false;

        Range? trailing = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        Tables? tables = null;
        ContentControls? controls = null;
        Frames? frames = null;
        Frame? frame = null;
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Range? canonicalCaption = null;
        Range? refreshedCaption = null;
        Range? numberRange = null;
        try
        {
            trailing = document.Range(captionParagraphRange.End, captionRange.End);
            var trailingText = trailing.Text ?? string.Empty;
            if (!trailingText.All(character =>
                    char.IsWhiteSpace(character) || character == '\a'))
                return false;

            maths = trailing.OMaths;
            shapes = trailing.InlineShapes;
            fields = trailing.Fields;
            tables = trailing.Tables;
            controls = trailing.ContentControls;
            if (maths.Count != 0
                || shapes.Count != 0
                || fields.Count != 0
                || tables.Count != 0
                || controls.Count != 0)
                return false;

            // Compatibility-mode Word can expand the hidden VTEqCap_<id>
            // bookmark and its clipping Frame across the pre-existing blank
            // paragraph at an insertion point before EnsureNormalTypingParagraph...
            // gets a chance to create/reuse the continuation. Canonicalize that
            // ownership first: the hidden caption is exactly its first paragraph;
            // the following structurally-empty paragraph belongs to body typing.
            frames = captionRange.Frames;
            if (frames.Count > 0)
            {
                frame = frames[1];
                frame.Delete();
                Release(frame);
                frame = null;
            }

            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName))
                return false;
            captionBookmark = bookmarks[captionName];
            captionBookmark.Delete();
            Release(captionBookmark);
            captionBookmark = null;

            canonicalCaption = document.Range(
                captionRange.Start,
                captionParagraphRange.End);
            captionBookmark = bookmarks.Add(captionName, canonicalCaption);

            if (!TryGetNativeCaptionRanges(
                    document,
                    formulaId,
                    GetNativeEquationSequenceName(document),
                    out refreshedCaption,
                    out numberRange)
                || refreshedCaption is null
                || numberRange is null)
                throw new InvalidDataException(
                    "VisualTeX could not rebind an expanded native equation caption to its own paragraph.");

            StyleNativeCaption(
                refreshedCaption,
                numberRange,
                cleanupLegacyFrames: false);
            TryTrimExpandedNativeCaptionBookmarksAfterStyle(
                document,
                formulaId,
                captionParagraphRange);
            return true;
        }
        finally
        {
            Release(numberRange);
            Release(refreshedCaption);
            Release(canonicalCaption);
            Release(captionBookmark);
            Release(bookmarks);
            Release(frame);
            Release(frames);
            Release(controls);
            Release(tables);
            Release(fields);
            Release(shapes);
            Release(maths);
            Release(trailing);
        }
    }

    internal static Range? EnsureNormalTypingParagraphAfterNumberedDisplay(
        Document document,
        string formulaId,
        Table? knownDirectTable = null)
    {
        if (knownDirectTable is not null
            && IsManagedNativeOmmlDirectTable(knownDirectTable))
        {
            Cell? knownNumberCell = null;
            Range? knownNumberRange = null;
            try
            {
                knownNumberCell = knownDirectTable.Cell(1, 3);
                knownNumberRange = knownNumberCell.Range.Duplicate;
                var knownBodyFormatting =
                    WordCharacterFormatting.CaptureParagraphMark(knownNumberRange);
                return EnsureDeletableTypingParagraphAfterNativeOmmlTable(
                    document,
                    knownDirectTable,
                    knownBodyFormatting);
            }
            finally
            {
                Release(knownNumberRange);
                Release(knownNumberCell);
            }
        }

        var bodyFormatting = CaptureNumberedBodyCharacterFormatting(document, formulaId);
        var nativeTableTypingRange =
            EnsureNormalTypingParagraphAfterNativeOmmlTable(
                document,
                formulaId,
                bodyFormatting,
                out var nativeDirectTableMatched);
        if (nativeDirectTableMatched)
            return nativeTableTypingRange;

        // Retired native #(SEQ) OMML kept SEQ and all number aliases inside the
        // display OMath itself. It has no hidden caption paragraph or Shape anchor,
        // so the legacy caption/frame routine below must not touch it. Create or
        // reuse the ordinary paragraph immediately after that mathematical paragraph
        // through the native-hash compatibility helper instead.
        var nativeHashTypingRange =
            EnsureNormalTypingParagraphAfterNativeOmmlHashSequence(
                document,
                formulaId);
        if (nativeHashTypingRange is not null)
            return nativeHashTypingRange;

        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Range? captionRange = null;
        Range? content = null;
        Range? numberRange = null;
        Paragraphs? captionParagraphs = null;
        Paragraph? captionParagraph = null;
        Range? captionParagraphRange = null;
        Range? restoredCaptionBookmarkRange = null;
        Frames? captionFrames = null;
        Frame? captionFrame = null;
        Paragraphs? typingParagraphs = null;
        Paragraph? typingParagraph = null;
        Range? typingRange = null;
        Range? trimmedCaptionRange = null;
        Frames? typingFrames = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName)) return null;

            captionBookmark = bookmarks[captionName];
            captionRange = captionBookmark.Range;
            content = document.Content;
            captionParagraphs = captionRange.Paragraphs;
            captionParagraph = captionParagraphs[1];
            captionParagraphRange = captionParagraph.Range;
            if (TryNormalizeExpandedNativeCaptionBeforeTyping(
                    document,
                    formulaId,
                    captionRange,
                    captionParagraphRange))
            {
                Release(captionParagraphRange);
                captionParagraphRange = null;
                Release(captionParagraph);
                captionParagraph = null;
                Release(captionParagraphs);
                captionParagraphs = null;
                Release(captionRange);
                captionRange = null;
                Release(captionBookmark);
                captionBookmark = null;

                if (!bookmarks.Exists(captionName))
                    throw new InvalidDataException(
                        "VisualTeX lost the normalized native equation caption bookmark.");
                captionBookmark = bookmarks[captionName];
                captionRange = captionBookmark.Range;
                captionParagraphs = captionRange.Paragraphs;
                captionParagraph = captionParagraphs[1];
                captionParagraphRange = captionParagraph.Range;
            }

            // Reuse only the paragraph immediately after this formula's native
            // caption. Never inspect document.Paragraphs[last]: doing so redirects
            // a create anchor from an earlier numbered formula to the document
            // tail whenever any later formula/body content exists.
            var nextParagraphStart = captionParagraphRange.End;
            if (nextParagraphStart < content.End)
            {
                typingRange = document.Range(
                    nextParagraphStart,
                    Math.Min(content.End, nextParagraphStart + 1));
                var inTable = (bool)typingRange.get_Information(WdInformation.wdWithInTable);
                typingFrames = typingRange.Frames;
                if (!inTable && typingFrames.Count == 0)
                {
                    typingParagraphs = typingRange.Paragraphs;
                    if (typingParagraphs.Count > 0)
                    {
                        typingParagraph = typingParagraphs[1];
                        Release(typingRange);
                        typingRange = typingParagraph.Range.Duplicate;
                        var paragraphText = typingRange.Text ?? string.Empty;
                        var isEmptyTypingParagraph = paragraphText.All(character =>
                            char.IsWhiteSpace(character) || character == '\a');
                        if (isEmptyTypingParagraph)
                        {
                            // Whole-paragraph Word paste can expand VTEqCap_<id>
                            // across the newly pasted empty continuation paragraph.
                            // The caption itself is still the first hidden SEQ
                            // paragraph; the trailing paragraph is ordinary body
                            // space and must not become caption-owned. Trim only
                            // this exact one-paragraph overflow before validating
                            // the continuation so formula+number and full-paragraph
                            // paste share the same stable typing behavior.
                            if (captionRange.End > captionParagraphRange.End
                                && typingRange.Start == captionParagraphRange.End
                                && typingRange.End <= captionRange.End)
                            {
                                trimmedCaptionRange = document.Range(
                                    captionRange.Start,
                                    captionParagraphRange.End);
                                captionBookmark.Delete();
                                Release(captionBookmark);
                                captionBookmark = bookmarks.Add(
                                    captionName,
                                    trimmedCaptionRange);
                            }
                            return NormalizeBodyTypingParagraph(document, typingRange, bodyFormatting);
                        }
                    }
                }
                Release(typingFrames);
                typingFrames = null;
                Release(typingRange);
                typingRange = null;
                Release(typingParagraph);
                typingParagraph = null;
                Release(typingParagraphs);
                typingParagraphs = null;
            }

            // No reusable local typing paragraph exists. Create one immediately
            // after this caption, even in the middle of the document. This keeps
            // a new numbered formula local while pushing the following formula or
            // body paragraph forward instead of nesting the new table in either.
            // Capture the bookmark coordinates first: Word automatically expands
            // a bookmark ending at this paragraph to include the new paragraph.
            var originalCaptionStart = captionRange.Start;
            var originalCaptionEnd = captionRange.End;
            captionParagraphRange.InsertParagraphAfter();

            // Word 2021 inherits a Frame when a paragraph is inserted after the
            // clipped native SEQ caption. If that inherited frame is left in
            // place, the "typing paragraph" is still hidden and Word can later
            // absorb user text into the neighboring numbered table. Remove the
            // shared frame first, then recreate the clipping frame around only
            // the original caption paragraph.
            captionFrames = captionRange.Frames;
            if (captionFrames.Count > 0)
            {
                captionFrame = captionFrames[1];
                captionFrame.Delete();
                Release(captionFrame);
                captionFrame = null;
                Release(captionFrames);
                captionFrames = null;
            }

            // InsertParagraphAfter expands VTEqCap_<id> across the inherited
            // blank paragraph. If StyleNativeCaption sees that expanded bookmark,
            // it frames the blank paragraph again and the returned insertion point
            // collapses onto the following table boundary. Restore the exact old
            // bookmark extent before rebuilding the clipped caption frame.
            captionBookmark.Delete();
            Release(captionBookmark);
            captionBookmark = null;
            restoredCaptionBookmarkRange = document.Range(
                originalCaptionStart,
                originalCaptionEnd);
            captionBookmark = bookmarks.Add(
                captionName,
                restoredCaptionBookmarkRange);

            Release(captionParagraphRange);
            captionParagraphRange = null;
            Release(captionParagraph);
            captionParagraph = null;
            Release(captionParagraphs);
            captionParagraphs = null;
            Release(captionRange);
            captionRange = null;
            if (!TryGetNativeCaptionRanges(
                    document,
                    formulaId,
                    GetNativeEquationSequenceName(document),
                    out captionRange,
                    out numberRange)
                || captionRange is null
                || numberRange is null)
                return null;
            StyleNativeCaption(
                captionRange,
                numberRange,
                cleanupLegacyFrames: false);

            captionParagraphs = captionRange.Paragraphs;
            captionParagraph = captionParagraphs[1];
            captionParagraphRange = captionParagraph.Range;
            // Compatibility-mode Word can expand both VTEqCap_<id> and
            // VTEqNum_<id> when the clipping Frame is rebuilt. Do not restyle the
            // caption again here: that would repeat the same expansion. Finalize
            // both bookmark extents only after all paragraph/Frame mutations are
            // complete, then re-resolve every live RCW before using the body tail.
            if (TryTrimExpandedNativeCaptionBookmarksAfterStyle(
                    document,
                    formulaId,
                    captionParagraphRange))
            {
                Release(captionParagraphRange);
                captionParagraphRange = null;
                Release(captionParagraph);
                captionParagraph = null;
                Release(captionParagraphs);
                captionParagraphs = null;
                Release(captionRange);
                captionRange = null;
                Release(captionBookmark);
                captionBookmark = null;

                if (!bookmarks.Exists(captionName))
                    throw new InvalidDataException(
                        "VisualTeX lost the finalized native equation caption bookmark after rebuilding its frame.");
                captionBookmark = bookmarks[captionName];
                captionRange = captionBookmark.Range;
                captionParagraphs = captionRange.Paragraphs;
                captionParagraph = captionParagraphs[1];
                captionParagraphRange = captionParagraph.Range;
            }

            var currentContentEnd = document.Content.End;
            var typingStart = captionParagraphRange.End;
            if (typingStart >= currentContentEnd) return null;
            typingRange = document.Range(
                typingStart,
                Math.Min(currentContentEnd, typingStart + 1));
            if ((bool)typingRange.get_Information(WdInformation.wdWithInTable))
                return null;
            typingFrames = typingRange.Frames;
            if (typingFrames.Count > 0) return null;

            return NormalizeBodyTypingParagraph(document, typingRange, bodyFormatting);
        }
        finally
        {
            Release(typingFrames);
            Release(trimmedCaptionRange);
            Release(typingRange);
            Release(typingParagraph);
            Release(typingParagraphs);
            Release(captionFrame);
            Release(captionFrames);
            Release(restoredCaptionBookmarkRange);
            Release(captionParagraphRange);
            Release(captionParagraph);
            Release(captionParagraphs);
            Release(numberRange);
            Release(content);
            Release(captionRange);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    internal static void CleanupNumberedDisplayInsertionSpacing(
        Document document,
        string formulaId,
        Table? knownDirectTable = null,
        bool cleanupFreshOmmlSource = false)
    {
        // The caller that just created a direct-SEQ host already owns its live Table
        // RCW. Reuse it instead of resolving VTEq_ -> Range -> Tables twice through
        // Word COM; on a 100-OMML document those two redundant lookups dominated the
        // otherwise trivial source-paragraph cleanup.
        if (knownDirectTable is not null)
        {
            try
            {
                // This overload is used only by the transaction that has just built
                // and validated the exact 1x3 host. Re-validating every cell, OMath
                // and field here costs hundreds of milliseconds, and a standalone
                // OMML→numbered edit creates no post-table typing tail: the only
                // generated residue is its now-empty source paragraph immediately
                // before this table.
                RemoveEmptyBodyParagraphImmediatelyBeforeTable(
                    document,
                    knownDirectTable,
                    formulaId,
                    cleanupFreshOmmlSource);
                return;
            }
            catch
            {
                // The direct table is already durable. A stale cleanup RCW must not
                // route a successful edit into retired caption/Shape logic.
                return;
            }
        }

        // Current numbered OMML is a direct-SEQ 1x3 table. Its source paragraph is
        // intentionally retained until formula, number and FormulaId ownership are
        // all durable, then removed here. Do not let it fall through to the retired
        // hidden-caption/Shape spacing path.
        if (CleanupNativeOmmlTablePrecedingParagraph(document, formulaId))
        {
            CleanupGeneratedNativeOmmlTypingTailBeforeFollowingTable(
                document,
                formulaId);
            return;
        }

        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Range? captionRange = null;
        Range? content = null;
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Frames? frames = null;
        ShapeRange? anchoredShapes = null;
        Bookmark? nativeDisplayAnchorBookmark = null;
        Range? nativeDisplayAnchorRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName)) return;
            captionBookmark = bookmarks[captionName];
            captionRange = captionBookmark.Range;
            content = document.Content;
            if (captionRange.End >= content.End) return;

            probe = document.Range(
                captionRange.End,
                Math.Min(content.End, captionRange.End + 1));
            if ((bool)probe.get_Information(WdInformation.wdWithInTable)) return;
            frames = probe.Frames;
            if (frames.Count > 0) return;
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            var nativeDisplayAnchorName = NativeDisplayAnchorBookmarkName(formulaId);
            if (bookmarks.Exists(nativeDisplayAnchorName))
            {
                nativeDisplayAnchorBookmark = bookmarks[nativeDisplayAnchorName];
                nativeDisplayAnchorRange = nativeDisplayAnchorBookmark.Range;
                if (nativeDisplayAnchorRange.Start >= paragraphRange.Start
                    && nativeDisplayAnchorRange.Start < paragraphRange.End)
                    return;
            }
            try
            {
                anchoredShapes = paragraphRange.ShapeRange;
                if (anchoredShapes.Count > 0) return;
            }
            catch
            {
                Release(anchoredShapes);
                anchoredShapes = null;
            }
            if (paragraphRange.Start != captionRange.End
                || !IsNumberingParagraphAdornment(paragraphRange.Text))
                return;

            // A temporary paragraph is needed while Word creates a numbered table
            // between two existing structures. Once the new formula is stable, that
            // paragraph is redundant and can be removed safely; the hidden native
            // caption paragraph itself keeps the neighboring formula tables separate.
            if (paragraphRange.End < content.End)
            {
                paragraphRange.Delete();
                return;
            }

            // Word always recreates the final document paragraph when it is deleted.
            // Keep that structural terminator as an ordinary typing paragraph. A
            // former implementation collapsed it to 1 pt; besides exposing a
            // surprising 1 pt caret to users, Word could retain the source list and
            // outline state on that paragraph. Heading-aware numbering then mistook
            // the generated tail for (for example) chapter 10 and rewrote later
            // equations as 10.0.x.
            NormalizeTrailingTypingParagraph(document, paragraphRange);
        }
        finally
        {
            Release(nativeDisplayAnchorRange);
            Release(nativeDisplayAnchorBookmark);
            Release(anchoredShapes);
            Release(frames);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
            Release(content);
            Release(captionRange);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    internal static bool TryExpandCompactTrailingTypingParagraphAfterFormula(
        Document document,
        string formulaId,
        Selection selection)
    {
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Bookmark? tailBookmark = null;
        Range? captionRange = null;
        Range? tailRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName)
                || !bookmarks.Exists(CompactTypingTailBookmarkName))
                return false;

            captionBookmark = bookmarks[captionName];
            tailBookmark = bookmarks[CompactTypingTailBookmarkName];
            captionRange = captionBookmark.Range;
            tailRange = tailBookmark.Range;
            if (tailRange.Start != captionRange.End)
                return false;

            selection.SetRange(tailRange.Start, tailRange.Start);
            return ExpandCompactTrailingTypingParagraph(selection);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(tailRange);
            Release(captionRange);
            Release(tailBookmark);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    internal static bool ExpandCompactTrailingTypingParagraph(Selection selection)
    {
        if (selection is null || selection.Start != selection.End) return false;
        Document? document = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        ParagraphFormat? format = null;
        ListFormat? listFormat = null;
        Microsoft.Office.Interop.Word.Font? selectionFont = null;
        try
        {
            document = selection.Document;
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(CompactTypingTailBookmarkName)) return false;
            bookmark = bookmarks[CompactTypingTailBookmarkName];
            bookmarkRange = bookmark.Range;
            if (selection.Start < bookmarkRange.Start
                || selection.Start > bookmarkRange.End)
                return false;
            paragraphs = bookmarkRange.Paragraphs;
            if (paragraphs.Count == 0) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (!IsNumberingParagraphAdornment(paragraphRange.Text)) return false;

            bookmark.Delete();
            try
            {
                object normalStyle = WdBuiltinStyle.wdStyleNormal;
                paragraphRange.set_Style(ref normalStyle);
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            try
            {
                listFormat = paragraphRange.ListFormat;
                listFormat.RemoveNumbers(WdNumberType.wdNumberParagraph);
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            font = paragraphRange.Font;
            font.Reset();
            font.Hidden = 0;
            font.Position = 0;
            font.Color = WdColor.wdColorAutomatic;
            format = paragraphRange.ParagraphFormat;
            format.Reset();
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            try { format.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText; }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }

            selection.SetRange(paragraphRange.Start, paragraphRange.Start);
            selectionFont = selection.Font;
            selectionFont.Reset();
            selectionFont.Hidden = 0;
            selectionFont.Position = 0;
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(selectionFont);
            Release(listFormat);
            Release(format);
            Release(font);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(document);
        }
    }

    internal static void RepairNumberedDisplaySpacing(Document document)
    {
        // Do not probe every bookmark/table through COM here. On a 100-formula
        // document that costs several seconds even when nothing is wrong. One
        // WordOpenXML snapshot tells us which formulas can actually own legacy
        // spacing damage; only those FormulaIds are touched through COM below.
        var repairFormulaIds = ReadNumberedDisplaySpacingRepairPlan(document);
        foreach (var formulaId in repairFormulaIds)
        {
            Table? table = null;
            try
            {
                table = FindNumberedEquationTable(document, formulaId);
                if (table is not null)
                    RemoveEmptyTrailingNumberedTableRows(table);
            }
            finally { Release(table); }
            CleanupNumberedDisplayInsertionSpacing(document, formulaId);
        }
    }

    private static IReadOnlyList<string> ReadNumberedDisplaySpacingRepairPlan(
        Document document)
    {
        Range? content = null;
        try
        {
            content = document.Content;
            var xml = WordDocumentXml.Read(document) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(xml)) return Array.Empty<string>();

            var planned = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var captionMatches = Regex.Matches(
                xml,
                @"<w:bookmarkStart\b(?=[^>]*\bw:name=""VTEqCap_(?<guid>[0-9A-F]{32})"")[^>]*/>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            // The final Word paragraph cannot be deleted; probing just the final
            // numbered formula lets CleanupNumberedDisplayInsertionSpacing restore
            // that terminator to a clean, ordinary typing paragraph when it directly
            // follows the last caption. This also upgrades documents created by the
            // historical 1 pt-tail implementation.
            if (captionMatches.Count > 0 && Guid.TryParseExact(
                    captionMatches[captionMatches.Count - 1].Groups["guid"].Value,
                    "N",
                    out var tailGuid))
                planned.Add(tailGuid.ToString("D"));

            // Legacy insertion builds could leave a second, completely empty row
            // in an otherwise standard three-column equation table. Detect those
            // candidates in the same WordOpenXML snapshot instead of reading
            // document.Tables / Rows.Count through COM for every formula. The latter
            // costs several seconds at 100+ formulas even when every table is healthy.
            var visibleMatches = Regex.Matches(
                xml,
                @"<w:bookmarkStart\b(?=[^>]*\bw:name=""VTEq_(?<guid>[0-9A-F]{32})"")[^>]*/>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            foreach (Match visibleMatch in visibleMatches)
            {
                if (!Guid.TryParseExact(
                        visibleMatch.Groups["guid"].Value,
                        "N",
                        out var visibleGuid))
                    continue;
                var tableStart = xml.LastIndexOf(
                    "<w:tbl",
                    visibleMatch.Index,
                    StringComparison.OrdinalIgnoreCase);
                var tableEnd = xml.IndexOf(
                    "</w:tbl>",
                    visibleMatch.Index,
                    StringComparison.OrdinalIgnoreCase);
                if (tableStart < 0 || tableEnd <= visibleMatch.Index) continue;
                tableEnd += "</w:tbl>".Length;
                if (tableEnd - tableStart > 262144) continue;
                var tableXml = xml.Substring(tableStart, tableEnd - tableStart);
                var rowCount = Regex.Matches(
                    tableXml,
                    @"<w:tr\b",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                if (rowCount > 1)
                    planned.Add(visibleGuid.ToString("D"));
            }

            // Old insertion builds could leave one ordinary empty paragraph
            // immediately after VTEqCap_<id> and immediately before the next table.
            // The paragraph is safe to delete, but only schedule it when the XML
            // shows that exact VisualTeX structural adjacency; ordinary user blank
            // paragraphs elsewhere are not touched.
            foreach (Match captionMatch in captionMatches)
            {
                if (!Guid.TryParseExact(
                        captionMatch.Groups["guid"].Value,
                        "N",
                        out var captionGuid))
                    continue;
                var captionParagraphEnd = xml.IndexOf(
                    "</w:p>",
                    captionMatch.Index,
                    StringComparison.OrdinalIgnoreCase);
                if (captionParagraphEnd < 0) continue;
                captionParagraphEnd += "</w:p>".Length;
                var probeLength = Math.Min(32768, xml.Length - captionParagraphEnd);
                if (probeLength <= 0) continue;
                var probe = xml.Substring(captionParagraphEnd, probeLength);
                var spacerMatch = Regex.Match(
                    probe,
                    @"^\s*(?:<w:bookmarkEnd\b[^>]*/>\s*)*"
                    + @"(?<paragraph><w:p\b[^>]*>.*?</w:p>)\s*<w:tbl\b",
                    RegexOptions.IgnoreCase
                    | RegexOptions.CultureInvariant
                    | RegexOptions.Singleline,
                    TimeSpan.FromSeconds(1));
                if (!spacerMatch.Success) continue;
                if (!IsStructurallyEmptyOpenXmlParagraph(
                        spacerMatch.Groups["paragraph"].Value))
                    continue;
                planned.Add(captionGuid.ToString("D"));
            }

            // Process later formulas first because deleting an obsolete spacer or
            // table row shifts story coordinates before earlier formulas.
            return planned
                .Select(formulaId => new
                {
                    FormulaId = formulaId!,
                    Position = FindOpenXmlCaptionPosition(xml, formulaId),
                })
                .OrderByDescending(item => item.Position)
                .Select(item => item.FormulaId)
                .ToArray();
        }
        catch
        {
            // Spacing cleanup is a compatibility repair layered on top of equation
            // numbering. If Word refuses to expose/parse the snapshot, keep the
            // existing numbering update path rather than turning a cosmetic repair
            // into a hard failure.
            return Array.Empty<string>();
        }
        finally { Release(content); }
    }

    private static int FindOpenXmlCaptionPosition(string xml, string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var formulaGuid)) return -1;
        return xml.IndexOf(
            $"VTEqCap_{formulaGuid:N}",
            StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsStructurallyEmptyOpenXmlParagraph(string paragraphXml)
    {
        if (string.IsNullOrWhiteSpace(paragraphXml)) return false;
        if (Regex.IsMatch(
                paragraphXml,
                @"<(?:w:t|w:instrText)\b[^>]*>\s*[^<\s]",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            return false;
        return paragraphXml.IndexOf("<w:drawing", StringComparison.OrdinalIgnoreCase) < 0
            && paragraphXml.IndexOf("<w:object", StringComparison.OrdinalIgnoreCase) < 0
            && paragraphXml.IndexOf("<w:pict", StringComparison.OrdinalIgnoreCase) < 0
            && paragraphXml.IndexOf("<w:fldChar", StringComparison.OrdinalIgnoreCase) < 0
            && paragraphXml.IndexOf("<m:oMath", StringComparison.OrdinalIgnoreCase) < 0;
    }

    private static void NormalizeTrailingTypingParagraph(
        Document document,
        Range paragraphRange)
    {
        Bookmarks? bookmarks = null;
        Bookmark? previousBookmark = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        ParagraphFormat? format = null;
        ListFormat? listFormat = null;
        try
        {
            try
            {
                object normalStyle = WdBuiltinStyle.wdStyleNormal;
                paragraphRange.set_Style(ref normalStyle);
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            try
            {
                listFormat = paragraphRange.ListFormat;
                listFormat.RemoveNumbers(WdNumberType.wdNumberParagraph);
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            font = paragraphRange.Font;
            font.Reset();
            font.Hidden = 0;
            font.Position = 0;
            font.Color = WdColor.wdColorAutomatic;
            format = paragraphRange.ParagraphFormat;
            format.Reset();
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            try { format.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText; }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }

            bookmarks = document.Bookmarks;
            if (bookmarks.Exists(CompactTypingTailBookmarkName))
            {
                previousBookmark = bookmarks[CompactTypingTailBookmarkName];
                previousBookmark.Delete();
                Release(previousBookmark);
                previousBookmark = null;
            }
        }
        finally
        {
            Release(previousBookmark);
            Release(bookmarks);
            Release(listFormat);
            Release(format);
            Release(font);
        }
    }

    internal static void RemoveEmptyTrailingNumberedTableRows(Table table)
    {
        Rows? rows = null;
        Row? row = null;
        Range? rowRange = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        try
        {
            rows = table.Rows;
            for (var index = rows.Count; index >= 2; index--)
            {
                Release(fields);
                fields = null;
                Release(maths);
                maths = null;
                Release(shapes);
                shapes = null;
                Release(rowRange);
                rowRange = null;
                Release(row);
                row = null;
                row = rows[index];
                rowRange = row.Range;
                shapes = rowRange.InlineShapes;
                maths = rowRange.OMaths;
                fields = rowRange.Fields;
                if (shapes.Count > 0 || maths.Count > 0 || fields.Count > 0)
                    continue;
                var rowText = (rowRange.Text ?? string.Empty)
                    .Replace("\a", string.Empty);
                if (!IsNumberingParagraphAdornment(rowText)) continue;
                row.Delete();
            }
        }
        finally
        {
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(rowRange);
            Release(row);
            Release(rows);
        }
    }
}
