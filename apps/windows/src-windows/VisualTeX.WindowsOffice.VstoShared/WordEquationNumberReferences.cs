using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordEquationNumbering;
using static VisualTeX.WordVsto.WordEquationParagraphLayout;
using static VisualTeX.WordVsto.WordEquationSequenceFields;
using static VisualTeX.WordVsto.WordNativeCaptions;
using static VisualTeX.WordVsto.WordNumberedDisplayLayout;
using static VisualTeX.WordVsto.WordNumberedTableDisplay;
using static VisualTeX.WordVsto.WordNumberedTrueDisplay;
using static VisualTeX.WordVsto.WordNumberingFields;
using static VisualTeX.WordVsto.WordNumberingReconcile;
using static VisualTeX.WordVsto.WordNumberingState;

namespace VisualTeX.WordVsto;

/// <summary>
/// Caption numbers and cross-references to equation numbers: reference targets, inserting and updating references,
/// their font size.
/// </summary>
internal static class WordEquationNumberReferences
{
    internal static void UpdateNativeEquationCaptionNumber(
        Document document,
        string formulaId,
        string nativeSequenceName,
        int ordinal,
        string prefix,
        bool formatOnly,
        bool cleanupLegacyFrames = true,
        Bookmarks? knownBookmarks = null,
        bool trustedHealthyDirectTable = false)
    {
        if (TryUpdateDirectTableSequenceNumber(
                document,
                formulaId,
                nativeSequenceName,
                ordinal,
                prefix,
                formatOnly,
                knownBookmarks,
                trustedHealthyDirectTable))
            return;
        if (TryRefreshOrAtomicallyRebuildNativeHashSequenceV2(
                document,
                formulaId,
                ordinal,
                prefix))
            return;

        // A SEQ field inside professional OMML is immutable as field code. If the
        // requested number format changes its heading reset/prefix, rebuild the one
        // OMath wrapper atomically; otherwise this path performs only Field.Update.
        // Never assign Field.Code.Text for the native #() host.
        if (IsNumberedNativeOmmlHashSequenceFormula(document, formulaId))
        {
            if (!TryRefreshNumberedNativeOmmlHashSequence(
                    document,
                    formulaId,
                    prefix))
                throw new InvalidOperationException(
                    $"Native #(SEQ) OMML {formulaId} could not update without rewriting Field.Code.Text.");
            return;
        }

        Bookmarks? bookmarks = null;
        Bookmark? numberBookmark = null;
        Bookmark? captionBookmark = null;
        Range? numberRange = null;
        Range? captionRange = null;
        Field? field = null;
        Range? code = null;
        Range? fieldResult = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefixRange = null;
        Range? refreshedNumberRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var numberName = NativeNumberBookmarkName(formulaId);
            var captionName = NativeCaptionBookmarkName(formulaId);
            if (!bookmarks.Exists(numberName) || !bookmarks.Exists(captionName)) return;
            numberBookmark = bookmarks[numberName];
            captionBookmark = bookmarks[captionName];
            numberRange = numberBookmark.Range;
            captionRange = captionBookmark.Range;
            field = FindNativeEquationFieldInRange(captionRange, nativeSequenceName);
            if (field is null)
                field = FindNativeEquationFieldAtRange(document, numberRange, nativeSequenceName);
            if (field is null) return;

            code = field.Code;
            EnsureSequenceFieldCodeCanBeRewritten(field);
            code.Text = $" SEQ {LegacyEquationSequenceName} \\r {ordinal} \\* ARABIC ";
            field.Update();
            fieldResult = field.Result;
            paragraphs = fieldResult.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            numberBookmark.Delete();
            Release(numberBookmark);
            numberBookmark = null;
            Release(code);
            code = field.Code;
            var fieldStart = Math.Max(paragraphRange.Start, code.Start - 1);
            prefixRange = document.Range(paragraphRange.Start, fieldStart);
            prefixRange.Text = prefix;

            Release(fieldResult);
            fieldResult = field.Result;
            refreshedNumberRange = document.Range(paragraphRange.Start, fieldResult.End);
            numberBookmark = bookmarks.Add(numberName, refreshedNumberRange);

            if (!formatOnly)
            {
                captionBookmark.Delete();
                Release(captionBookmark);
                captionBookmark = null;
                Release(captionRange);
                captionRange = paragraph.Range;
                captionBookmark = bookmarks.Add(captionName, captionRange);
                // The strict XML inventory already proved this existing host.
                // Updating only its ordinal/prefix must not reapply Frame/font/
                // paragraph styles for every formula; that layout work dominated
                // a 50-number refresh. Malformed/legacy paths retain full restyling.
                if (!trustedHealthyDirectTable)
                    StyleNativeCaption(
                        captionRange,
                        refreshedNumberRange,
                        cleanupLegacyFrames);
            }
        }
        finally
        {
            Release(refreshedNumberRange);
            Release(prefixRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(fieldResult);
            Release(code);
            Release(field);
            Release(captionRange);
            Release(numberRange);
            Release(captionBookmark);
            Release(numberBookmark);
            Release(bookmarks);
        }
    }

    internal static void UpdateFieldInBookmark(
        Document document,
        string bookmarkName,
        Func<string?, bool> predicate)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Fields? fields = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(bookmarkName)) return;
            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range;
            fields = range.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    if (IsFieldInsideOmml(field)) continue;
                    code = field.Code;
                    if (predicate(code.Text))
                    {
                        field.Update();
                        NormalizeReferenceResult(
                            field,
                            knownFontSize: null);
                    }
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            ApplyParagraphEquationNumberFont(range, fallbackSize: 0f, position: 0);
        }
        finally
        {
            Release(fields);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static void AlignEquationNumberVertically(
        Range numberRange,
        float formulaHeightPoints,
        float formulaFontSizePoints)
    {
        var numberFontSize = FormulaFontSize.Normalize(formulaFontSizePoints);
        ApplyParagraphEquationNumberFont(
            numberRange,
            numberFontSize,
            CalculateEquationNumberFontPosition(formulaHeightPoints, numberFontSize));
    }

    internal static void ApplyParagraphEquationNumberFont(
        Range targetRange,
        float fallbackSize,
        int position)
    {
        // Word can transiently expand a just-created REF/bookmark range across
        // the following paragraph while a batch of numbered formulas is being
        // materialized. Formatting that expanded range both throws in the strict
        // paragraph-mark helper and can leak the number style into following body
        // text. Resolve the paragraph that owns the label's first character and
        // clamp the formatting write to that paragraph only.
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? effectiveRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            paragraphs = targetRange.Paragraphs;
            if (paragraphs.Count == 0)
                throw new InvalidDataException("The equation number has no paragraph owner.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var bodyFormatting = WordCharacterFormatting.CaptureParagraphMark(paragraphRange);

            effectiveRange = targetRange.Duplicate;
            var start = Math.Max(effectiveRange.Start, paragraphRange.Start);
            var end = Math.Min(effectiveRange.End, paragraphRange.End);
            if (end <= start)
                throw new InvalidDataException("The equation number range does not intersect its paragraph owner.");
            effectiveRange.SetRange(start, end);

            font = effectiveRange.Font;
            font.Reset();
            if (fallbackSize > 0f)
                font.Size = FormulaFontSize.Normalize(fallbackSize);
            bodyFormatting.Apply(font);
            font.Position = position;
        }
        finally
        {
            Release(font);
            Release(effectiveRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool TryReadReferenceTargetsFromOpenXml(
        Document document,
        out IReadOnlyList<NativeEquationCaptionEntry> entries,
        out bool sawLegacyNumberingArtifact)
    {
        entries = Array.Empty<NativeEquationCaptionEntry>();
        sawLegacyNumberingArtifact = false;
        Range? content = null;
        try
        {
            content = document.Content;
            var xml = WordDocumentXml.Read(document) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(xml)) return false;

            var bookmarkNames = Regex.Matches(
                    xml,
                    @"\bw:name=""(?<name>VTEq(?:Cap|Num)?_[0-9A-F]{32})""",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                .Cast<Match>()
                .Select(match => match.Groups["name"].Value)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            sawLegacyNumberingArtifact = bookmarkNames.Any(name =>
                name.StartsWith(EquationBookmarkPrefix, StringComparison.OrdinalIgnoreCase)
                || name.StartsWith(NativeCaptionBookmarkPrefix, StringComparison.OrdinalIgnoreCase));

            var startMatches = Regex.Matches(
                xml,
                @"<w:bookmarkStart\b(?=[^>]*\bw:id=""(?<id>-?\d+)"")(?=[^>]*\bw:name=""VTEqNum_(?<guid>[0-9A-F]{32})"")[^>]*/>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (startMatches.Count == 0) return false;

            var result = new List<NativeEquationCaptionEntry>(startMatches.Count);
            foreach (Match startMatch in startMatches)
            {
                if (!Guid.TryParseExact(
                        startMatch.Groups["guid"].Value,
                        "N",
                        out var formulaGuid))
                    continue;
                var formulaId = formulaGuid.ToString("D");
                if (!bookmarkNames.Contains(NativeCaptionBookmarkName(formulaId))
                    || !bookmarkNames.Contains(EquationBookmarkName(formulaId)))
                    continue;

                var bookmarkId = Regex.Escape(startMatch.Groups["id"].Value);
                var endMatch = Regex.Match(
                    xml,
                    $@"<w:bookmarkEnd\b[^>]*\bw:id=""{bookmarkId}""[^>]*/>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                    TimeSpan.FromSeconds(1));
                if (!endMatch.Success || endMatch.Index <= startMatch.Index)
                    continue;
                if (endMatch.Index - startMatch.Index > 16384)
                    continue;

                var segment = xml.Substring(
                    startMatch.Index + startMatch.Length,
                    endMatch.Index - startMatch.Index - startMatch.Length);
                var numberText = NormalizeNativeEquationNumberText(
                    ExtractVisibleEquationNumberTextFromOpenXmlSegment(segment));
                if (string.IsNullOrWhiteSpace(numberText)) continue;
                result.Add(new NativeEquationCaptionEntry(
                    formulaId,
                    startMatch.Index,
                    numberText));
            }

            if (result.Count == 0) return false;
            entries = result.OrderBy(entry => entry.Position).ToArray();
            return true;
        }
        catch
        {
            entries = Array.Empty<NativeEquationCaptionEntry>();
            return false;
        }
        finally { Release(content); }
    }

    internal static IReadOnlyList<EquationReferenceTarget> GetEquationReferenceTargets(
        Document document) =>
        GetEquationReferenceTargets(document, allowLegacyReconcile: true);

    private static IReadOnlyList<EquationReferenceTarget> GetEquationReferenceTargets(
        Document document,
        bool allowLegacyReconcile)
    {
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        var watch = tracePerformance
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        long checkpoint = 0;
        void TraceStage(string stage)
        {
            if (watch is null) return;
            var elapsed = watch.ElapsedMilliseconds;
            Console.WriteLine(
                $"    [perf] reference-targets.{stage}: +{elapsed - checkpoint}ms ({elapsed}ms)");
            checkpoint = elapsed;
        }

        var entries = new List<NativeEquationCaptionEntry>();
        var previews = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var sawLegacyNumberingArtifact = false;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        InlineShapes? inlineShapes = null;
        try
        {
            // WordOpenXML gives all bookmark names, rendered number text and
            // document order in one COM call. This avoids materializing hundreds
            // of individual Bookmark COM objects merely to open the picker.
            if (TryReadReferenceTargetsFromOpenXml(
                    document,
                    out var xmlEntries,
                    out sawLegacyNumberingArtifact))
            {
                entries.AddRange(xmlEntries);
                TraceStage("number-bookmarks-xml");
            }
            else
            {
                // Conservative compatibility fallback for unusual Word XML or
                // older builds: enumerate bookmarks exactly as before.
                var bookmarkNames = new HashSet<string>(StringComparer.Ordinal);
                bookmarks = document.Bookmarks;
                for (var index = 1; index <= bookmarks.Count; index++)
                {
                    Release(range);
                    range = null;
                    Release(bookmark);
                    bookmark = bookmarks[index];
                    var name = bookmark.Name;
                    bookmarkNames.Add(name);
                    if (TryFormulaIdFromBookmark(
                            name,
                            NativeNumberBookmarkPrefix,
                            out var formulaId))
                    {
                        range = bookmark.Range;
                        var numberText = NormalizeNativeEquationNumberText(range.Text);
                        if (string.IsNullOrWhiteSpace(numberText)) continue;
                        entries.Add(new NativeEquationCaptionEntry(
                            formulaId,
                            range.Start,
                            numberText));
                        continue;
                    }
                    if (name.StartsWith(EquationBookmarkPrefix, StringComparison.Ordinal)
                        || name.StartsWith(NativeCaptionBookmarkPrefix, StringComparison.Ordinal))
                        sawLegacyNumberingArtifact = true;
                }
                entries.RemoveAll(entry =>
                    !bookmarkNames.Contains(NativeCaptionBookmarkName(entry.FormulaId))
                    || !bookmarkNames.Contains(EquationBookmarkName(entry.FormulaId)));
                TraceStage("number-bookmarks-com");
            }

            // LaTeX is only a convenience preview/search key. New and edited OLE
            // formulas carry a cheap Word-side metadata cache; old OLEs simply
            // show their number instead of launching 100 OLE servers. Functional
            // reference identity always comes from VTEqNum_* above.
            if (entries.Count > 0)
            {
                inlineShapes = document.InlineShapes;
                for (var index = 1; index <= inlineShapes.Count; index++)
                {
                    InlineShape? shape = null;
                    try
                    {
                        shape = inlineShapes[index];
                        var metadata = WordFormulaMetadataReader.TryReadCachedPreview(shape);
                        if (metadata is null || !metadata.Numbered) continue;
                        var latex = string.Join(" ", metadata.Lines.Select(line => line.Latex))
                            .Replace("\r", " ")
                            .Replace("\n", " ")
                            .Trim();
                        if (latex.Length > 90) latex = latex.Substring(0, 87) + "…";
                        if (!string.IsNullOrWhiteSpace(latex))
                            previews[metadata.FormulaId] = latex;
                    }
                    finally { Release(shape); }
                }
            }
            TraceStage("cached-previews");
        }
        finally
        {
            Release(inlineShapes);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }

        if (entries.Count == 0 && allowLegacyReconcile && sawLegacyNumberingArtifact)
        {
            // Old documents from before VTEqNum_* can repair themselves once.
            // Healthy current documents never enter this structural path.
            Reconcile(document);
            TraceStage("legacy-reconcile");
            return GetEquationReferenceTargets(document, allowLegacyReconcile: false);
        }

        var ordered = entries
            .OrderBy(entry => entry.Position)
            .ToArray();
        var result = ordered
            .Select(entry => new EquationReferenceTarget(
                entry.FormulaId,
                entry.NumberText,
                previews.TryGetValue(entry.FormulaId, out var preview)
                    ? preview
                    : string.Empty,
                entry.Position))
            .ToArray();
        TraceStage("sort");
        return result;
    }

    internal static void InsertEquationReference(
        Document document,
        Selection selection,
        EquationReferenceTarget target,
        EquationReferenceStyle style,
        WdColor? preferredInsertionColor = null)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (selection is null) throw new ArgumentNullException(nameof(selection));
        if (target is null) throw new ArgumentNullException(nameof(target));
        if (target.Source != EquationReferenceSource.VisualTeX)
            throw new InvalidOperationException(
                "The selected equation is not a VisualTeX/OMML reference target.");

        var prefix = style switch
        {
            EquationReferenceStyle.EquationPrefix => OfficePluginLanguage.Text("式(", "Eq. ("),
            EquationReferenceStyle.Parenthesized => "(",
            _ => string.Empty,
        };
        var suffix = style switch
        {
            EquationReferenceStyle.EquationPrefix => ")",
            EquationReferenceStyle.Parenthesized => ")",
            _ => string.Empty,
        };

        // A direct REF field updates its text, but Word does not treat it like the
        // native MathType equation reference that users can double-click. Use the
        // same real Word field tree for every VisualTeX/OMML reference: an outer
        // GOTOBUTTON routes navigation and one nested REF keeps the number live.
        WordEquationReferenceFields.InsertNavigableReference(
            document,
            selection,
            NativeNumberBookmarkName(target.FormulaId),
            prefix,
            suffix,
            preferredInsertionColor);
    }

    internal static int UpdateHealthyNativeCrossReferencesAfterRenumbering(
        Document document,
        IReadOnlyDictionary<string, string> changedFormulaNumbers,
        int? searchEndExclusive = null,
        IReadOnlyDictionary<string, int>? knownReferenceCounts = null)
    {
        if (changedFormulaNumbers.Count == 0) return 0;

        // The explicit update command already parsed WordOpenXML while validating
        // the numbering scaffold. Reuse its exact REF counts to avoid touching the
        // document-wide Fields COM collection when a changed formula has only its
        // generated right-side number REF (the overwhelmingly common case).
        if (!searchEndExclusive.HasValue && knownReferenceCounts is not null)
        {
            Bookmarks? bookmarks = null;
            var fastUpdated = 0;
            var hasExternalReferences = false;
            var canUseTargetedPath = true;
            try
            {
                bookmarks = document.Bookmarks;
                foreach (var changed in changedFormulaNumbers)
                {
                    if (!knownReferenceCounts.TryGetValue(changed.Key, out var referenceCount)
                        || referenceCount < 1)
                    {
                        canUseTargetedPath = false;
                        break;
                    }
                    // TryPatchVisibleEquationNumberResult already distinguishes the
                    // current direct-SEQ 1x3 host from retired mathematical #(SEQ)
                    // hosts. Calling IsNumberedNativeOmmlHashSequenceFormula first
                    // re-opened CustomXML metadata and OMath WordOpenXML for every
                    // changed formula; five ordinary table numbers alone cost close
                    // to a second in a 100-OMML document.
                    if (!TryPatchVisibleEquationNumberResult(
                            bookmarks,
                            changed.Key,
                            changed.Value))
                    {
                        canUseTargetedPath = false;
                        break;
                    }
                    fastUpdated++;
                    if (referenceCount > 1)
                        hasExternalReferences = true;
                }
            }
            catch
            {
                canUseTargetedPath = false;
            }
            finally { Release(bookmarks); }

            if (canUseTargetedPath)
            {
                if (!hasExternalReferences)
                    return fastUpdated;

                // Extra body references really exist. Let Word update them in one
                // native batch call rather than enumerating every Field through COM.
                // The generated visible REF results were already patched locally.
                UpdateMainStoryFields(document);
                foreach (var changed in changedFormulaNumbers)
                    NormalizeGeneratedTabLayoutEquationNumberLabel(
                        document,
                        changed.Key);
                return fastUpdated;
            }
        }

        Range? content = null;
        Range? searchRange = null;
        Fields? fields = null;
        var updated = 0;
        var foundTargetReference = false;
        try
        {
            if (searchEndExclusive.HasValue)
            {
                content = document.Content;
                var end = Math.Max(
                    content.Start,
                    Math.Min(searchEndExclusive.Value, content.End));
                searchRange = document.Range(content.Start, end);
                fields = searchRange.Fields;
            }
            else
            {
                fields = document.Fields;
            }
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? result = null;
                var formulaId = string.Empty;
                var generatedTabLayout = false;
                try
                {
                    field = fields[index];
                    // Mathematical REF fields belong only to retired bad #(...)
                    // wrappers. Never update, rewrite, patch or unlink them in the
                    // cross-reference layer; structural migration replaces the
                    // complete OMath atomically.
                    if (IsFieldInsideOmml(field)) continue;
                    // Most large numbered documents also contain one EMBED field
                    // per OLE formula. Reading Field.Code for every EMBED field is
                    // surprisingly expensive in Word, so reject non-REF fields by
                    // their native type before touching the code range.
                    if (field.Type != WdFieldType.wdFieldRef) continue;
                    code = field.Code;
                    var codeText = code.Text ?? string.Empty;
                    if (!IsReferenceFieldCode(codeText)) continue;
                    if (!TryResolveVisualTeXReferenceBookmark(
                            document,
                            codeText,
                            out var visualTeXBookmark)
                        || !TryFormulaIdFromBookmark(
                            visualTeXBookmark,
                            NativeNumberBookmarkPrefix,
                            out formulaId)
                        || !changedFormulaNumbers.TryGetValue(
                            formulaId,
                            out var expectedNumber))
                        continue;
                    foundTargetReference = true;

                    // Body references (including nested REF entries exposed by
                    // Document.Fields) belong to the common formatting-preserving
                    // refresher. Only a physical generated number is patched here.
                    if (!IsGeneratedEquationNumberReference(document, field, formulaId))
                        continue;

                    var alreadyCanonical = IsReferenceToBookmark(
                            codeText,
                            visualTeXBookmark)
                        && Regex.IsMatch(
                            codeText,
                            @"\\\*\s+CHARFORMAT\b",
                            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                        && !Regex.IsMatch(
                            codeText,
                            @"\\\*\s+MERGEFORMAT\b",
                            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    generatedTabLayout =
                        IsGeneratedTabLayoutEquationNumberReference(
                            document,
                            field,
                            formulaId);
                    if (!alreadyCanonical)
                    {
                        NormalizeReferenceResult(
                            field,
                            knownFontSize: null);
                        if (generatedTabLayout)
                            NormalizeGeneratedTabLayoutEquationNumberLabel(
                                document,
                                formulaId);
                        updated++;
                        continue;
                    }

                    result = field.Result;
                    var currentText = NormalizeNativeEquationNumberText(result.Text);
                    if (!string.Equals(currentText, expectedNumber, StringComparison.Ordinal))
                    {
                        // The target bookmark and field code are unchanged; only
                        // the numeric result shifted. Replacing the existing field
                        // result preserves the REF field and its CHARFORMAT while
                        // avoiding a full Word field recalculation for every one of
                        // the 80+ suffix formulas in a large document.
                        result.Text = expectedNumber;
                    }
                    if (generatedTabLayout)
                        NormalizeGeneratedTabLayoutEquationNumberLabel(
                            document,
                            formulaId);
                    updated++;
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
        }
        finally
        {
            Release(fields);
            Release(searchRange);
            Release(content);
        }
        // In the scoped middle-insertion path the suffix's generated labels were
        // already patched directly. If the prefix contains no REF to a shifted
        // formula there is nothing left for the document-wide reference walker to
        // do; avoiding that scan is material in documents with hundreds of EMBEDs.
        if (searchEndExclusive.HasValue && !foundTargetReference)
            return updated;
        return updated + WordEquationReferenceFields.UpdateReferences(document,
            changedFormulaNumbers.Keys.Select(NativeNumberBookmarkName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase));
    }

    internal static int UpdateNativeCrossReferences(Document document) =>
        UpdateNativeCrossReferences(document, targetFormulaIds: null);

    internal static int UpdateNativeCrossReferences(
        Document document,
        ISet<string>? targetFormulaIds)
    {
        Fields? fields = null;
        var updated = 0;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsReferenceFieldCode(code.Text)) continue;
                    if (!TryResolveVisualTeXReferenceBookmark(
                            document,
                            code.Text,
                            out var visualTeXBookmark))
                        continue;

                    if (!TryFormulaIdFromBookmark(
                            visualTeXBookmark,
                            NativeNumberBookmarkPrefix,
                            out var formulaId))
                        continue;
                    if (targetFormulaIds is not null
                        && !targetFormulaIds.Contains(formulaId))
                        continue;

                    // A body REF keeps its original bookmark and delimiter contract.
                    // Resolving an alias to an equation identifies its owner; it does
                    // not authorize rewriting ZEqnNum (visible label) to VTEqNum
                    // (digits only). Canonicalize only the generated number itself.
                    if (!IsGeneratedEquationNumberReference(document, field, formulaId))
                        continue;

                    if (!IsReferenceToBookmark(code.Text, visualTeXBookmark))
                    {
                        EnsureReferenceFieldCodeCanBeRewritten(field);
                        code.Text = $" REF {visualTeXBookmark} \\h ";
                    }

                    var generatedTabLayout =
                        IsGeneratedTabLayoutEquationNumberReference(
                            document,
                            field,
                            formulaId);
                    NormalizeReferenceResult(
                        field,
                        knownFontSize: null);
                    if (generatedTabLayout)
                        NormalizeGeneratedTabLayoutEquationNumberLabel(
                            document,
                            formulaId);
                    updated++;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }

        // One common body-reference pass handles both direct and nested REF.
        ISet<string>? targetBookmarkNames = targetFormulaIds is null
            ? null
            : targetFormulaIds
                .Select(NativeNumberBookmarkName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        updated += WordEquationReferenceFields.UpdateReferences(
            document,
            targetBookmarkNames);
        return updated;
    }

    private static void NormalizeGeneratedTabLayoutEquationNumberLabel(
        Document document,
        string formulaId,
        float fallbackSize = 0f)
    {
        Range? labelRange = null;
        try
        {
            labelRange = FindVisibleEquationNumberTextRange(document, formulaId);
            if (labelRange is null
                || labelRange.StoryType != WdStoryType.wdMainTextStory
                || (bool)labelRange.get_Information(WdInformation.wdWithInTable))
                return;
            var labelText = labelRange.Text ?? string.Empty;
            if (!labelText.StartsWith("(", StringComparison.Ordinal)
                || !labelText.EndsWith(")", StringComparison.Ordinal))
                return;
            ApplyParagraphEquationNumberFont(
                labelRange,
                fallbackSize,
                position: 0);
        }
        finally { Release(labelRange); }
    }

    internal static bool IsGeneratedEquationNumberReference(
        Document document, Field field, string formulaId)
    {
        Range? label = null;
        Range? owner = null;
        Range? code = null;
        Range? result = null;
        try
        {
            if (field.Type != WdFieldType.wdFieldRef)
                return false;

            label = FindVisibleEquationNumberRange(document, formulaId);
            if (label is null) return false;
            owner = FindNumberingOwnerRange(document, formulaId);
            if (owner is null) return false;

            code = field.Code;
            if (!IsReferenceToBookmark(
                    code.Text,
                    NativeNumberBookmarkName(formulaId)))
                return false;

            result = field.Result;

            // Word 2021 may report Result.End at or beyond the paragraph
            // structure position for a REF whose visible digits are still wholly
            // inside VTEq_*. End is therefore not a stable identity boundary.
            // Prove generation from the local owner and the result's start
            // instead: this exact REF targets this formula's private VTEqNum_*,
            // both code and result live in the numbering owner, and the visible
            // result starts inside the VTEq_* label. A user/body REF to the same
            // bookmark lives outside that private label and does not satisfy this.
            return code.StoryType == owner.StoryType
                && result.StoryType == owner.StoryType
                && label.StoryType == owner.StoryType
                && code.Start >= owner.Start
                && code.Start < owner.End
                && result.Start >= label.Start
                && result.Start <= label.End
                && result.Start >= owner.Start
                && result.Start < owner.End;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(result);
            Release(code);
            Release(owner);
            Release(label);
        }
    }

    private static bool IsGeneratedTabLayoutEquationNumberReference(
        Document document,
        Field field,
        string formulaId)
    {
        Range? visibleRange = null;
        Range? result = null;
        try
        {
            visibleRange = FindVisibleEquationNumberRange(document, formulaId);
            if (visibleRange is null || IsNumberedEquationTable(visibleRange)) return false;
            if (visibleRange.StoryType != WdStoryType.wdMainTextStory)
                return false;
            result = field.Result;
            return result.StoryType == WdStoryType.wdMainTextStory
                && result.Start >= visibleRange.Start
                && result.End <= visibleRange.End;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(result);
            Release(visibleRange);
        }
    }

    internal static void NormalizeReferenceResult(
        Field field,
        float? knownFontSize = null)
    {
        if (IsFieldInsideOmml(field))
            throw new InvalidOperationException(
                "VisualTeX refused to normalize or rewrite a REF field inside OMML. Legacy #(REF) numbering must be migrated by atomically replacing the complete equation.");

        Range? code = null;
        Range? result = null;
        try
        {
            var size = knownFontSize.HasValue && knownFontSize.Value > 0f
                ? FormulaFontSize.Normalize(knownFontSize.Value)
                : ResolveReferenceFontSize(field);
            code = field.Code;
            var codeText = code.Text ?? string.Empty;
            var normalizedCode = Regex.Replace(
                codeText,
                @"\\\*\s+MERGEFORMAT\b",
                string.Empty,
                RegexOptions.IgnoreCase);
            if (!Regex.IsMatch(
                    normalizedCode,
                    @"\\\*\s+CHARFORMAT\b",
                    RegexOptions.IgnoreCase))
            {
                normalizedCode = normalizedCode.TrimEnd() + " \\* CHARFORMAT ";
            }
            if (!string.Equals(codeText, normalizedCode, StringComparison.Ordinal))
            {
                EnsureReferenceFieldCodeCanBeRewritten(field);
                code.Text = normalizedCode;
            }

            // CHARFORMAT protects the visible number from its clipped SEQ
            // target. Creation and refresh share the body's paragraph-mark format
            // for table, tab and text-frame hosts alike.
            ApplyParagraphEquationNumberFont(code, size, position: 0);
            field.Update();
            result = field.Result;
            ApplyParagraphEquationNumberFont(result, size, position: 0);
        }
        finally
        {
            Release(result);
            Release(code);
        }
    }

    private static float ResolveReferenceFontSize(Field field)
    {
        Range? result = null;
        Range? code = null;
        Range? paragraphRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Document? document = null;
        try
        {
            result = field.Result;
            var size = ReadUsableFontSize(result);
            if (size.HasValue) return size.Value;

            code = field.Code;
            size = ReadUsableFontSize(code);
            if (size.HasValue) return size.Value;

            paragraphs = result.Paragraphs;
            if (paragraphs.Count == 0) return 11f;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            document = result.Document;

            // Result.Start sits inside the field, so the immediately adjacent
            // character can still be a hidden field separator. Probe outside the
            // complete field code/result boundary, then the paragraph mark. This
            // recovers the surrounding body size even after Word copied the 1 pt
            // SEQ target appearance into both field code and result.
            var candidatePositions = new[]
            {
                code.Start - 2,
                code.Start - 1,
                result.End + 1,
                result.End + 2,
                paragraphRange.End - 1,
                paragraphRange.Start,
            };
            foreach (var position in candidatePositions.Distinct())
            {
                size = ReadUsableFontSizeAt(
                    document,
                    paragraphRange,
                    position);
                if (size.HasValue) return size.Value;
            }
            return 11f;
        }
        finally
        {
            Release(document);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(code);
            Release(result);
        }
    }

    private static float? ReadUsableFontSizeAt(
        Document document,
        Range paragraphRange,
        int position)
    {
        if (position < paragraphRange.Start || position >= paragraphRange.End)
            return null;
        Range? range = null;
        try
        {
            object start = position;
            object end = Math.Min(paragraphRange.End, position + 1);
            range = document.Range(ref start, ref end);
            return ReadUsableFontSize(range);
        }
        catch
        {
            return null;
        }
        finally { Release(range); }
    }

    private static float? ReadUsableFontSize(Range range)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            var size = font.Size;
            return float.IsNaN(size)
                || float.IsInfinity(size)
                || size <= 2f
                || size > 256f
                ? null
                : size;
        }
        catch
        {
            return null;
        }
        finally { Release(font); }
    }
}
