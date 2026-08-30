using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static partial class WordEquationNumbering
{
    private const string NativeHashSequenceMigrationPlaceholder = "\uE00A";

    private static bool ConfigureNumberedNativeOmmlDisplayHashRetired(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId,
        bool reuseExistingScaffold,
        FormulaMetadata? metadata,
        Action<string> traceStage,
        int? plannedOrdinal = null,
        string? plannedPrefix = null,
        bool deferFieldUpdate = false,
        bool deferExternalShapeCreation = false)
    {
        _ = formulaHeightPoints;
        _ = formulaFontSizePoints;
        _ = reuseExistingScaffold;
        _ = plannedOrdinal;
        _ = deferExternalShapeCreation;

        Range? activeRange = null;
        Range? migratedRange = null;
        Range? replacementRange = null;
        Range? cleanReplacementTarget = null;
        Bookmark? repairedBookmark = null;
        Microsoft.Office.Interop.Word.Application? application = null;
        try
        {
            metadata ??= WordOmmlFormulaStore.TryRead(document, formulaId);
            activeRange = ResolveSingleNativeOmmlRange(formulaRange);
            if (TryRemoveLegacyNativeDisplayShapeBeforeHashMigrationV2(
                    document,
                    formulaId))
            {
                Release(activeRange);
                activeRange = null;
                try
                {
                    activeRange = ResolveSingleNativeOmmlRange(formulaRange);
                }
                catch
                {
                    if (metadata is null) throw;
                    activeRange = WordOmmlFormulaStore
                        .GetEquationRangeVerifiedForStructuralEdit(
                            document,
                            formulaId,
                            metadata);
                }
                traceStage("migrate-shape-host-to-native-hash");
            }
            EnsureNumberedOmmlIsDisplay(activeRange);
            ConfigureEquationParagraph(activeRange, numbered: false);
            var format = ReadEquationNumberFormat(document);
            var resolvedPrefix = plannedPrefix;
            if (resolvedPrefix is null)
            {
                var anchors = format.UsesHeading
                    ? GetHeadingNumberAnchors(document, format.HeadingLevel)
                    : Array.Empty<HeadingNumberAnchor>();
                resolvedPrefix = ResolveEquationNumberScope(
                    activeRange.Start,
                    format,
                    anchors).Prefix;
            }

            // The production numbered-OMML host is Word's own display-equation
            // #(...) structure. Its number is the result of a SEQ field inside the
            // mathematical delimiter itself; VTEqNum_<FormulaId> bookmarks that
            // result so ordinary body REF fields can target it. No visible REF,
            // floating Shape/TextBox, anchor paragraph, or hidden SEQ paragraph is
            // part of this host.
            if (IsHealthyNumberedNativeOmmlHashSequenceHost(
                    document,
                    activeRange,
                    formulaId)
                && NativeOmmlHashSequenceFormatMatches(
                    document,
                    activeRange,
                    formulaId,
                    format,
                    resolvedPrefix ?? string.Empty))
            {
                if (!deferFieldUpdate)
                    UpdateNativeOmmlHashSequenceField(activeRange);
                EnsureNumberedOmmlIsDisplay(activeRange);
                ConfigureEquationParagraph(activeRange, numbered: false);
                formulaRange.SetRange(activeRange.Start, activeRange.End);
                traceStage("native-hash-seq-preserved");
                return false;
            }

            var semanticOmml = metadata is not null
                ? WordOmmlConverter.StripVisualTeXNativeEquationNumberForManagedRepair(
                    activeRange.WordOpenXML)
                : WordOmmlConverter.StripVisualTeXNativeEquationNumber(
                    activeRange.WordOpenXML);
            if (WordOmmlConverter.HasVisualTeXNativeEquationNumber(activeRange.WordOpenXML))
                traceStage("strip-legacy-or-stale-number-wrapper");

            if (IsNumberedEquationTable(activeRange))
            {
                TrimBenignEmptyRowsFromNumberedTable(document, activeRange, formulaId);
                migratedRange = TryConvertStandardNumberedOmmlTableToStandaloneDisplayParagraph(
                    document,
                    activeRange,
                    formulaId,
                    metadata);
                if (migratedRange is null)
                    throw new InvalidOperationException(
                        "VisualTeX could not safely migrate this managed numbered OMML table to Word's native #(SEQ) display host.");
                Release(activeRange);
                activeRange = ResolveSingleNativeOmmlRange(migratedRange);
                semanticOmml = WordOmmlConverter.StripVisualTeXNativeEquationNumber(
                    activeRange.WordOpenXML);
                traceStage("migrate-omml-table-to-hash-seq");
            }

            // Shape-era documents are migration input only. Remove their drawing
            // and external caption before replacing the mathematical paragraph.
            // If VTEqCap/VTEqNum already live inside this OMath, they are aliases of
            // a stale mathematical wrapper and are deleted as bookmarks only; the
            // OMath itself is replaced atomically below.
            RemoveLegacyNativeOmmlNumberingScaffold(
                document,
                activeRange,
                formulaId);
            traceStage("remove-legacy-number-scaffold");

            var preparedOmml = BuildNativeOmmlHashSequenceNumber(
                semanticOmml,
                formulaId,
                format,
                resolvedPrefix ?? string.Empty);

            // OLE→OMML and a few legacy tab hosts materialize a temporary plain
            // OMath between layout TAB/manual-break characters. Replacing only that
            // OMath would leave those old layout characters outside the new
            // m:oMathPara, violating the invariant that OMath.End is immediately
            // followed by ¶. When—and only when—the surrounding paragraph contains
            // no user text, second object or field, collapse its complete editable
            // content to one ordinary placeholder before importing #(SEQ).
            cleanReplacementTarget =
                PrepareCleanNativeHashSequenceMigrationTarget(
                    document,
                    activeRange);
            if (cleanReplacementTarget is not null)
            {
                Release(activeRange);
                activeRange = cleanReplacementTarget;
                cleanReplacementTarget = null;
                traceStage("remove-legacy-tab-adornments");
            }

            application = document.Application;
            replacementRange = WordOmmlConverter.ReplaceWithPreparedOmml(
                application,
                document,
                activeRange,
                preparedOmml,
                display: true,
                mathFontName: document.OMathFontName);
            Release(activeRange);
            activeRange = replacementRange;
            replacementRange = null;
            EnsureNumberedOmmlIsDisplay(activeRange);
            ConfigureEquationParagraph(activeRange, numbered: false);
            traceStage("materialize-native-hash-seq");

            if (!deferFieldUpdate)
                UpdateNativeOmmlHashSequenceField(activeRange);
            traceStage(deferFieldUpdate ? "defer-native-seq-update" : "update-native-seq");

            if (!IsHealthyNumberedNativeOmmlHashSequenceHost(
                    document,
                    activeRange,
                    formulaId,
                    requireCurrentFieldResult: !deferFieldUpdate))
            {
                TraceNativeOmmlHashSequenceDiagnostics(
                    document,
                    activeRange,
                    formulaId,
                    "post-materialize-health-failure");
                throw new InvalidOperationException(
                    $"Numbered OMML {formulaId} did not materialize as one healthy Word #(SEQ) display equation.");
            }

            if (metadata is not null)
            {
                repairedBookmark = WordOmmlFormulaStore.Wrap(
                    document,
                    activeRange,
                    metadata,
                    replaceExisting: true);
                WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                    metadata,
                    activeRange);
                WordOmmlFormulaStore.Save(document, metadata);
            }

            formulaRange.SetRange(activeRange.Start, activeRange.End);
            return true;
        }
        finally
        {
            Release(application);
            Release(repairedBookmark);
            Release(cleanReplacementTarget);
            Release(replacementRange);
            Release(migratedRange);
            Release(activeRange);
        }
    }

    private static Range? PrepareCleanNativeHashSequenceMigrationTarget(
        Document document,
        Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? before = null;
        Range? after = null;
        Range? editableRange = null;
        Range? placeholderRange = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The temporary numbered OMML migration host spans multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable))
                return null;

            var editableEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            var alreadyPure = formulaRange.Start == paragraphRange.Start
                && formulaRange.End == editableEnd;
            if (alreadyPure) return null;

            if (formulaRange.Start < paragraphRange.Start
                || formulaRange.End > editableEnd)
                throw new InvalidOperationException(
                    "The temporary numbered OMML migration range escaped its paragraph.");

            maths = paragraphRange.OMaths;
            shapes = paragraphRange.InlineShapes;
            fields = paragraphRange.Fields;
            if (maths.Count != 1 || shapes.Count != 0 || fields.Count != 0)
            {
                if (string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                {
                    var fieldCodes = new List<string>();
                    for (var fieldIndex = 1; fieldIndex <= fields.Count; fieldIndex++)
                    {
                        Field? field = null;
                        Range? code = null;
                        try
                        {
                            field = fields[fieldIndex];
                            code = field.Code;
                            fieldCodes.Add((code.Text ?? string.Empty)
                                .Replace("\r", "\\r")
                                .Replace("\n", "\\n"));
                        }
                        finally
                        {
                            Release(code);
                            Release(field);
                        }
                    }
                    var paragraphText = (paragraphRange.Text ?? string.Empty)
                        .Replace("\r", "\\r")
                        .Replace("\n", "\\n")
                        .Replace("\t", "\\t")
                        .Replace("\v", "\\v")
                        .Replace("\u0013", "<FIELD_BEGIN>")
                        .Replace("\u0014", "<FIELD_SEPARATE>")
                        .Replace("\u0015", "<FIELD_END>");
                    Console.WriteLine(
                        $"    native-hash-migration-refused formula={formulaRange.Start}:{formulaRange.End} paragraph={paragraphRange.Start}:{paragraphRange.End} text='{paragraphText}' maths={maths.Count} shapes={shapes.Count} fields={fields.Count} fieldCodes=[{string.Join(" | ", fieldCodes)}]");
                }
                throw new InvalidOperationException(
                    "VisualTeX refused to remove legacy OMML layout characters from a paragraph containing another equation, object, or field.");
            }

            before = document.Range(paragraphRange.Start, formulaRange.Start);
            after = document.Range(formulaRange.End, editableEnd);
            if (!ContainsOnlyNativeHashSequenceMigrationAdornment(before.Text)
                || !ContainsOnlyNativeHashSequenceMigrationAdornment(after.Text))
                throw new InvalidOperationException(
                    "VisualTeX refused to remove ordinary user text surrounding a numbered OMML migration.");

            editableRange = document.Range(paragraphRange.Start, editableEnd);
            editableRange.Text = NativeHashSequenceMigrationPlaceholder;
            placeholderRange = document.Range(
                paragraphRange.Start,
                paragraphRange.Start + NativeHashSequenceMigrationPlaceholder.Length);
            var result = placeholderRange;
            placeholderRange = null;
            return result;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(maths);
            Release(placeholderRange);
            Release(editableRange);
            Release(after);
            Release(before);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool ContainsOnlyNativeHashSequenceMigrationAdornment(
        string? text)
    {
        foreach (var character in text ?? string.Empty)
        {
            if (character is '\t' or '\v' or '\r' or '\n' or '\a'
                or '\u0001' or '\u200B' or '\u200C' or '\u200D'
                or '\u2060' or '\uFEFF'
                || char.IsWhiteSpace(character))
                continue;
            return false;
        }
        return true;
    }

    private static string BuildNativeOmmlHashSequenceNumber(
        string semanticOmml,
        string formulaId,
        EquationNumberFormat format,
        string prefix)
    {
        return WordOmmlConverter.BuildImmutableHashSequenceNumberedOmml(
            semanticOmml,
            LegacyEquationSequenceName,
            NativeNumberBookmarkName(formulaId),
            EquationBookmarkName(formulaId),
            NativeCaptionBookmarkName(formulaId),
            prefix,
            restartHeadingLevel: format.UsesHeading ? format.HeadingLevel : 0,
            initialSequenceResult: "1");
    }

    internal static bool IsSafeManagedNativeOmmlHashSequenceOwnerForReplacement(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        if (document is null || formulaRange is null) return false;
        if (!HasManagedNativeOmmlHashSequenceHost(document, formulaId))
            return false;

        Range? activeRange = null;
        OMaths? formulaMaths = null;
        OMath? formulaMath = null;
        Fields? formulaFields = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        OMaths? paragraphMaths = null;
        InlineShapes? inlineShapes = null;
        ShapeRange? shapeRange = null;
        Range? before = null;
        Range? after = null;
        try
        {
            activeRange = ResolveSingleNativeOmmlRange(formulaRange);
            formulaMaths = activeRange.OMaths;
            formulaFields = activeRange.Fields;
            if (formulaMaths.Count != 1
                || formulaFields.Count != 1
                || !ContainsVisualTeXSequenceInsideOmml(activeRange))
                return false;
            formulaMath = formulaMaths[1];
            if (formulaMath.Type != WdOMathType.wdOMathDisplay)
                return false;

            paragraphs = activeRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable))
                return false;
            paragraphMaths = paragraphRange.OMaths;
            inlineShapes = paragraphRange.InlineShapes;
            if (paragraphMaths.Count != 1 || inlineShapes.Count != 0)
                return false;

            // Current native numbering owns the complete paragraph: the OMath
            // begins at the paragraph start and its Range is followed immediately
            // by the normal paragraph mark. These exact boundaries prove that an
            // atomic owner replacement cannot delete adjacent user prose.
            if (activeRange.Start != paragraphRange.Start
                || paragraphRange.End != activeRange.End + 1)
                return false;
            before = document.Range(paragraphRange.Start, activeRange.Start);
            after = document.Range(activeRange.End, paragraphRange.End);
            if (!IsNumberingParagraphAdornment(before.Text)
                || !IsNumberingParagraphAdornment(after.Text))
                return false;

            try
            {
                shapeRange = paragraphRange.ShapeRange;
                if (shapeRange.Count > 0) return false;
            }
            catch (COMException)
            {
                // Word throws when no Shape is anchored to the paragraph, which is
                // the expected healthy case. Any other structural proof above still
                // remains mandatory.
            }
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(after);
            Release(before);
            Release(shapeRange);
            Release(inlineShapes);
            Release(paragraphMaths);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(formulaFields);
            Release(formulaMath);
            Release(formulaMaths);
            Release(activeRange);
        }
    }

    internal static bool HasReusableNumberedNativeOmmlHashSequenceHost(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        if (!HasReusableNumberedNativeOmmlHashSequenceHostCore(
                document,
                formulaRange,
                formulaId))
            return false;
        Bookmarks? ownershipBookmarks = null;
        try
        {
            ownershipBookmarks = document.Bookmarks;
            if (ownershipBookmarks.Exists(
                    NativeDisplayAnchorBookmarkName(formulaId))
                || HasNativeDisplayAnchorCommitMarker(document, formulaId))
                return false;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(ownershipBookmarks);
        }

        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Bookmarks? bookmarks = null;
        Bookmark? numberBookmark = null;
        Range? numberRange = null;
        try
        {
            fields = formulaRange.Fields;
            if (fields.Count != 1) return false;
            field = fields[1];
            code = field.Code;
            var instruction = code.Text ?? string.Empty;
            if (!IsVisualTeXSequenceFieldCode(instruction)
                || System.Text.RegularExpressions.Regex.IsMatch(
                    instruction,
                    @"\\r\s+\d+",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                    | System.Text.RegularExpressions.RegexOptions.CultureInvariant))
                return false;

            var format = ReadEquationNumberFormat(document);
            var anyHeadingReset = System.Text.RegularExpressions.Regex.IsMatch(
                instruction,
                @"\\s\s+\d+",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
                | System.Text.RegularExpressions.RegexOptions.CultureInvariant);
            if (!format.UsesHeading)
            {
                if (anyHeadingReset) return false;
            }
            else if (!System.Text.RegularExpressions.Regex.IsMatch(
                         instruction,
                         $@"\\s\s+{format.HeadingLevel}\b",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase
                         | System.Text.RegularExpressions.RegexOptions.CultureInvariant))
            {
                return false;
            }

            bookmarks = document.Bookmarks;
            var numberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(numberName)) return false;
            numberBookmark = bookmarks[numberName];
            numberRange = numberBookmark.Range;
            if (!format.UsesHeading) return true;

            var anchors = GetHeadingNumberAnchorsForFormatBatch(
                document,
                format.HeadingLevel,
                new[]
                {
                    new NativeEquationCaptionEntry(
                        formulaId,
                        formulaRange.Start,
                        string.Empty),
                });
            var scope = ResolveEquationNumberScope(
                formulaRange.Start,
                format,
                anchors);
            var renderedNumber = NormalizeNativeEquationNumberText(numberRange.Text);
            return renderedNumber.StartsWith(scope.Prefix, StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static bool HasReusableNumberedNativeOmmlHashSequenceHostCore(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (formulaRange is null) throw new ArgumentNullException(nameof(formulaRange));
        _ = NormalizeFormulaIdForBookmark(formulaId);
        var reusable = IsHealthyNumberedNativeOmmlHashSequenceHost(
            document,
            formulaRange,
            formulaId,
            requireCurrentFieldResult: false);
        if (!reusable)
        {
            TraceNativeOmmlHashSequenceDiagnostics(
                document,
                formulaRange,
                formulaId,
                "atomic-replacement-probe-failure");
        }
        return reusable;
    }

    internal static string PrepareNumberedNativeOmmlHashSequenceForReplacement(
        Document document,
        string semanticOmml,
        string formulaId,
        int formulaPosition)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(semanticOmml))
            throw new InvalidDataException(
                "The replacement numbered OMML payload is empty.");
        _ = NormalizeFormulaIdForBookmark(formulaId);

        var format = ReadEquationNumberFormat(document);
        var anchors = format.UsesHeading
            ? GetHeadingNumberAnchors(document, format.HeadingLevel)
            : Array.Empty<HeadingNumberAnchor>();
        var prefix = ResolveEquationNumberScope(
            formulaPosition,
            format,
            anchors).Prefix;
        return BuildNativeOmmlHashSequenceNumber(
            semanticOmml,
            formulaId,
            format,
            prefix);
    }

    internal static void RemoveNativeOmmlHashSequenceAliasesForReplacement(
        Document document,
        string formulaId)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        _ = NormalizeFormulaIdForBookmark(formulaId);

        // These three bookmarks are nested inside the OMath that is about to be
        // atomically replaced. Delete only the bookmark identities, never their
        // ranges: deleting the range would damage the live mathematical SEQ field.
        // Removing the names first also prevents Word from renaming or discarding
        // the same FormulaId aliases imported with the replacement OMath.
        DeleteBookmarkOnly(document, EquationBookmarkName(formulaId));
        DeleteBookmarkOnly(document, NativeCaptionBookmarkName(formulaId));
        DeleteBookmarkOnly(document, NativeNumberBookmarkName(formulaId));
    }

    private static bool NativeOmmlHashSequenceFormatMatches(
        Document document,
        Range formulaRange,
        string formulaId,
        EquationNumberFormat format,
        string expectedPrefix)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Bookmarks? bookmarks = null;
        Bookmark? numberBookmark = null;
        Range? numberRange = null;
        try
        {
            fields = formulaRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? candidate = null;
                Range? candidateCode = null;
                try
                {
                    candidate = fields[index];
                    candidateCode = candidate.Code;
                    if (!IsVisualTeXSequenceFieldCode(candidateCode.Text))
                        continue;
                    field = candidate;
                    candidate = null;
                    code = candidateCode;
                    candidateCode = null;
                    break;
                }
                finally
                {
                    Release(candidateCode);
                    Release(candidate);
                }
            }
            if (field is null || code is null) return false;
            var codeText = code.Text ?? string.Empty;
            if (Regex.IsMatch(
                    codeText,
                    @"\\r\s+\d+\b",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                return false;
            var resetMatch = Regex.Match(
                codeText,
                @"\\s\s+(?<level>\d+)\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (format.UsesHeading)
            {
                if (!resetMatch.Success
                    || !int.TryParse(
                        resetMatch.Groups["level"].Value,
                        out var resetLevel)
                    || resetLevel != format.HeadingLevel)
                    return false;
            }
            else if (resetMatch.Success)
            {
                return false;
            }

            bookmarks = document.Bookmarks;
            var numberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(numberName)) return false;
            numberBookmark = bookmarks[numberName];
            numberRange = numberBookmark.Range;
            var currentNumber = NormalizeNativeEquationNumberText(numberRange.Text);
            return string.IsNullOrEmpty(expectedPrefix)
                || currentNumber.StartsWith(expectedPrefix, StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static bool TryRefreshNumberedNativeOmmlHashSequence(
        Document document,
        string formulaId,
        string? plannedPrefix = null)
    {
        FormulaMetadata? metadata = null;
        Range? formulaRange = null;
        try
        {
            metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
            if (metadata is null
                || !metadata.Numbered
                || !string.Equals(
                    metadata.DisplayMode,
                    "block",
                    StringComparison.OrdinalIgnoreCase))
                return false;
            formulaRange = WordOmmlFormulaStore
                .GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    formulaId,
                    metadata);
            if (!WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                    formulaRange.WordOpenXML))
                return false;

            ConfigureNumberedNativeOmmlDisplay(
                document,
                formulaRange,
                WordOmmlFormulaStore.EstimateHeightPoints(formulaRange),
                (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                formulaId,
                reuseExistingScaffold: true,
                metadata,
                _ => { },
                plannedPrefix: plannedPrefix,
                deferFieldUpdate: false,
                deferExternalShapeCreation: false);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(formulaRange);
        }
    }

    private static bool TryRefreshOrAtomicallyRebuildNativeHashSequenceV2(
        Document document,
        string formulaId,
        int ordinal,
        string prefix)
    {
        _ = ordinal;
        if (!IsNumberedNativeOmmlHashSequenceFormula(document, formulaId))
            return false;

        // A native #(SEQ) formula must never rewrite Field.Code.Text in-place.
        // ConfigureNumberedNativeOmmlDisplay preserves a matching host and performs
        // an ordinary Field.Update, or atomically rebuilds the complete OMath when
        // the current heading reset/prefix no longer matches the requested format.
        if (!TryRefreshNumberedNativeOmmlHashSequence(
                document,
                formulaId,
                prefix))
            throw new InvalidOperationException(
                $"Native #(SEQ) OMML {formulaId} could not refresh or rebuild safely.");
        return true;
    }

    private static bool IsNumberedNativeOmmlHashSequenceFormula(
        Document document,
        string formulaId)
    {
        Range? formulaRange = null;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
            if (metadata is null
                || !metadata.Numbered
                || !string.Equals(
                    metadata.DisplayMode,
                    "block",
                    StringComparison.OrdinalIgnoreCase))
                return false;
            formulaRange = WordOmmlFormulaStore
                .GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    formulaId,
                    metadata);
            return WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                formulaRange.WordOpenXML);
        }
        catch
        {
            return false;
        }
        finally { Release(formulaRange); }
    }

    private static int AdoptUnownedCopiesOfManagedNativeHashSequenceHosts(
        Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (document.ReadOnly) return 0;

        var storedMetadata = new Dictionary<string, FormulaMetadata>(
            StringComparer.OrdinalIgnoreCase);
        var storedFingerprintCounts = new Dictionary<string, int>(
            StringComparer.OrdinalIgnoreCase);
        foreach (var formulaId in WordOmmlFormulaStore.StoredFormulaIds(document))
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
            if (metadata is null
                || !metadata.Numbered
                || !string.Equals(
                    metadata.DisplayMode,
                    "block",
                    StringComparison.OrdinalIgnoreCase))
                continue;
            if (string.IsNullOrWhiteSpace(metadata.NativeOmmlFingerprint))
            {
                // Without every existing logical formula's semantic fingerprint,
                // physical-host surplus cannot prove that an unowned OMath is a
                // newly pasted copy rather than an older formula whose bookmark was
                // lost. Defer to explicit selection adoption in that ambiguous case.
                return 0;
            }
            storedMetadata[formulaId] = metadata;
            storedFingerprintCounts.TryGetValue(
                metadata.NativeOmmlFingerprint!,
                out var count);
            storedFingerprintCounts[metadata.NativeOmmlFingerprint!] = count + 1;
        }
        if (storedMetadata.Count == 0) return 0;

        var managedAnchorPositions = new HashSet<int>();
        foreach (var formulaId in WordOmmlFormulaStore.BookmarkedFormulaIds(document))
        {
            Bookmark? bookmark = null;
            Range? bookmarkRange = null;
            try
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                if (bookmark is null) continue;
                bookmarkRange = bookmark.Range;
                managedAnchorPositions.Add(bookmarkRange.Start);
            }
            catch { }
            finally
            {
                Release(bookmarkRange);
                Release(bookmark);
            }
        }

        var managedRanges = new List<(int Start, int End)>();
        var managedFingerprints = new HashSet<string>(
            storedFingerprintCounts.Keys,
            StringComparer.OrdinalIgnoreCase);
        foreach (var formulaId in WordOmmlFormulaStore.FormulaIds(document))
        {
            Range? formulaRange = null;
            try
            {
                if (!storedMetadata.TryGetValue(formulaId, out var metadata))
                    continue;
                formulaRange =
                    TryResolveManagedNativeHashSequenceRangeForAliasRepair(
                        document,
                        formulaId,
                        metadata);
                if (formulaRange is null) continue;
                var formulaXml = formulaRange.WordOpenXML ?? string.Empty;
                if (!WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                        formulaXml))
                    continue;
                managedRanges.Add((formulaRange.Start, formulaRange.End));
                try
                {
                    var semantic = WordOmmlConverter
                        .StripManagedVisualTeXNativeEquationNumber(formulaXml);
                    managedFingerprints.Add(
                        WordOmmlConverter.ComputeOmmlFingerprint(semantic));
                }
                catch
                {
                    if (!string.IsNullOrWhiteSpace(metadata.NativeOmmlFingerprint))
                        managedFingerprints.Add(metadata.NativeOmmlFingerprint!);
                }
            }
            catch
            {
                // A damaged managed formula remains the responsibility of the
                // normal identity/migration repair path. It is not evidence that an
                // unrelated unowned OMath should be adopted as a pasted copy.
            }
            finally { Release(formulaRange); }
        }
        if (managedRanges.Count == 0 || managedFingerprints.Count == 0)
            return 0;

        var physicalFingerprintCounts = new Dictionary<string, int>(
            StringComparer.OrdinalIgnoreCase);
        var candidates = new List<(int Start, int End, string Fingerprint)>();
        OMaths? documentMaths = null;
        try
        {
            documentMaths = document.OMaths;
            for (var index = 1; index <= documentMaths.Count; index++)
            {
                OMath? math = null;
                Range? mathRange = null;
                Paragraphs? paragraphs = null;
                Paragraph? paragraph = null;
                Range? paragraphRange = null;
                try
                {
                    math = documentMaths[index];
                    if (math.Type != WdOMathType.wdOMathDisplay) continue;
                    mathRange = math.Range.Duplicate;
                    if (mathRange.StoryType != WdStoryType.wdMainTextStory
                        || (bool)mathRange.get_Information(
                            WdInformation.wdWithInTable))
                        continue;
                    var mathXml = mathRange.WordOpenXML ?? string.Empty;
                    if (!WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                            mathXml))
                        continue;
                    var semantic = WordOmmlConverter
                        .StripManagedVisualTeXNativeEquationNumber(mathXml);
                    var fingerprint =
                        WordOmmlConverter.ComputeOmmlFingerprint(semantic);
                    physicalFingerprintCounts.TryGetValue(
                        fingerprint,
                        out var physicalCount);
                    physicalFingerprintCounts[fingerprint] = physicalCount + 1;
                    if (!managedFingerprints.Contains(fingerprint))
                        continue;
                    if (managedRanges.Any(owner =>
                            mathRange.Start < owner.End
                            && mathRange.End > owner.Start))
                        continue;
                    if (managedAnchorPositions.Any(anchor =>
                            anchor == mathRange.Start
                            || anchor == mathRange.Start - 1
                            || anchor >= mathRange.Start
                                && anchor <= mathRange.End))
                        continue;

                    // A pasted VisualTeX display formula must still be one pure
                    // OMath paragraph. Reject inline mixtures, adjacent user text and
                    // tables before considering the semantic fingerprint match.
                    paragraphs = mathRange.Paragraphs;
                    if (paragraphs.Count != 1) continue;
                    paragraph = paragraphs[1];
                    paragraphRange = paragraph.Range;
                    if (mathRange.End + 1 != paragraphRange.End)
                        continue;

                    candidates.Add((
                        mathRange.Start,
                        mathRange.End,
                        fingerprint));
                }
                catch
                {
                    // Skip an inaccessible/transient OMath. No identity is safer
                    // than guessing while Word is still committing a paste.
                }
                finally
                {
                    Release(paragraphRange);
                    Release(paragraph);
                    Release(paragraphs);
                    Release(mathRange);
                    Release(math);
                }
            }
        }
        finally { Release(documentMaths); }
        if (candidates.Count == 0) return 0;

        var eligibleCandidates = new List<(
            int Start,
            int End,
            string Fingerprint)>();
        foreach (var group in candidates.GroupBy(
                     candidate => candidate.Fingerprint,
                     StringComparer.OrdinalIgnoreCase))
        {
            physicalFingerprintCounts.TryGetValue(group.Key, out var physicalCount);
            storedFingerprintCounts.TryGetValue(group.Key, out var logicalCount);
            var provenCopySurplus = Math.Max(0, physicalCount - logicalCount);
            var groupedCandidates = group.ToArray();
            if (provenCopySurplus <= 0) continue;
            if (groupedCandidates.Length != provenCopySurplus)
            {
                if (string.Equals(
                        Environment.GetEnvironmentVariable(
                            "VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                    Console.WriteLine(
                        $"    native-hash-copy-adoption-ambiguous fingerprint={group.Key} physical={physicalCount} metadata={logicalCount} unowned={groupedCandidates.Length}");
                continue;
            }
            eligibleCandidates.AddRange(groupedCandidates);
        }
        if (eligibleCandidates.Count == 0) return 0;

        var adopted = 0;
        foreach (var candidate in eligibleCandidates
                     .OrderByDescending(item => item.Start))
        {
            Range? probe = null;
            OMaths? maths = null;
            OMath? math = null;
            Range? formulaRange = null;
            try
            {
                probe = document.Range(candidate.Start, candidate.End);
                maths = probe.OMaths;
                if (maths.Count != 1) continue;
                math = maths[1];
                if (math.Type != WdOMathType.wdOMathDisplay) continue;
                formulaRange = math.Range.Duplicate;
                var liveXml = formulaRange.WordOpenXML ?? string.Empty;
                if (!WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                        liveXml))
                    continue;
                var liveSemantic = WordOmmlConverter
                    .StripManagedVisualTeXNativeEquationNumber(liveXml);
                var liveFingerprint =
                    WordOmmlConverter.ComputeOmmlFingerprint(liveSemantic);
                if (!string.Equals(
                        liveFingerprint,
                        candidate.Fingerprint,
                        StringComparison.OrdinalIgnoreCase)
                    || !managedFingerprints.Contains(liveFingerprint))
                    continue;

                var metadata = WordOmmlNativeSource.CreateForNative(
                    document,
                    formulaRange);
                if (!metadata.Numbered
                    || !string.Equals(
                        metadata.DisplayMode,
                        "block",
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                ConfigureNumberedNativeOmmlDisplay(
                    document,
                    formulaRange,
                    WordOmmlFormulaStore.EstimateHeightPoints(formulaRange),
                    (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                    metadata.FormulaId,
                    reuseExistingScaffold: false,
                    metadata,
                    _ => { });
                adopted++;
                if (string.Equals(
                        Environment.GetEnvironmentVariable(
                            "VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                    Console.WriteLine(
                        $"    native-hash-copy-adopted formulaId={metadata.FormulaId} range={candidate.Start}:{candidate.End}");
            }
            catch (Exception error)
            {
                if (string.Equals(
                        Environment.GetEnvironmentVariable(
                            "VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                    Console.WriteLine(
                        $"    native-hash-copy-adoption-skipped range={candidate.Start}:{candidate.End} error={error.GetType().Name}:{error.Message}");
            }
            finally
            {
                Release(formulaRange);
                Release(math);
                Release(maths);
                Release(probe);
            }
        }

        if (adopted > 0)
        {
            // Copy/paste can also drag one half of a source VTEq* bookmark into the
            // copied OMath. Revalidate every managed native host after rebuilding
            // the copies so the original FormulaId/body REF target is repaired
            // before document-wide sequence reconciliation begins.
            TryRepairDriftedManagedNativeHashSequenceHosts(document, out _);
        }
        return adopted;
    }

    private static Range? TryResolveManagedNativeHashSequenceRangeForAliasRepair(
        Document document,
        string formulaId,
        FormulaMetadata metadata)
    {
        Bookmark? formulaBookmark = null;
        Range? bookmarkRange = null;
        OMaths? documentMaths = null;
        OMath? candidateMath = null;
        Range? candidateRange = null;
        Range? adjacentMatch = null;
        Range? fingerprintMatch = null;
        var fingerprintMatchCount = 0;
        try
        {
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            if (formulaBookmark is null) return null;
            bookmarkRange = formulaBookmark.Range;
            var anchor = bookmarkRange.Start;
            var expectedFingerprint = metadata.NativeOmmlFingerprint ?? string.Empty;

            documentMaths = document.OMaths;
            for (var index = 1; index <= documentMaths.Count; index++)
            {
                Release(candidateRange); candidateRange = null;
                Release(candidateMath); candidateMath = documentMaths[index];
                if (candidateMath.Type != WdOMathType.wdOMathDisplay)
                    continue;
                candidateRange = candidateMath.Range.Duplicate;

                string candidateXml;
                try { candidateXml = candidateRange.WordOpenXML ?? string.Empty; }
                catch { continue; }
                if (!WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                        candidateXml))
                    continue;

                var fingerprintMatches = false;
                if (!string.IsNullOrWhiteSpace(expectedFingerprint))
                {
                    try
                    {
                        var semanticOmml = WordOmmlConverter
                            .StripManagedVisualTeXNativeEquationNumber(candidateXml);
                        var candidateFingerprint =
                            WordOmmlConverter.ComputeOmmlFingerprint(semanticOmml);
                        fingerprintMatches = string.Equals(
                            candidateFingerprint,
                            expectedFingerprint,
                            StringComparison.OrdinalIgnoreCase);
                    }
                    catch
                    {
                        fingerprintMatches = false;
                    }
                }

                var bookmarkOverlaps = bookmarkRange.End > bookmarkRange.Start
                    && bookmarkRange.Start < candidateRange.End
                    && bookmarkRange.End > candidateRange.Start;
                var isCanonicalAdjacent = anchor == candidateRange.Start
                    || anchor == candidateRange.Start - 1
                    || bookmarkOverlaps;
                if (isCanonicalAdjacent
                    && (string.IsNullOrWhiteSpace(expectedFingerprint)
                        || fingerprintMatches))
                {
                    if (adjacentMatch is not null)
                        return null;
                    adjacentMatch = candidateRange.Duplicate;
                }

                if (!fingerprintMatches) continue;
                fingerprintMatchCount++;
                Release(fingerprintMatch);
                fingerprintMatch = candidateRange.Duplicate;
            }

            if (adjacentMatch is not null)
            {
                var result = adjacentMatch;
                adjacentMatch = null;
                return result;
            }
            if (fingerprintMatchCount == 1 && fingerprintMatch is not null)
            {
                var result = fingerprintMatch;
                fingerprintMatch = null;
                return result;
            }
            return null;
        }
        finally
        {
            Release(fingerprintMatch);
            Release(adjacentMatch);
            Release(candidateRange);
            Release(candidateMath);
            Release(documentMaths);
            Release(bookmarkRange);
            Release(formulaBookmark);
        }
    }

    private static bool TryRepairDriftedManagedNativeHashSequenceHosts(
        Document document,
        out int repaired)
    {
        repaired = 0;
        var hosts = new List<(string FormulaId, int Start, int End)>();
        var repairRequired = false;
        try
        {
            // Freeze every managed native #(SEQ) OMath before touching a single
            // VTEq* alias. Word can cross-wire one formula's bookmark start/end
            // into a later OMath after an insertion. Repairing only the formula
            // first observed as unhealthy can therefore invalidate a second host
            // that currently contains the other half of the crossing bookmark.
            foreach (var formulaId in WordOmmlFormulaStore.BookmarkedFormulaIds(document))
            {
                Range? formulaRange = null;
                try
                {
                    var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                    if (metadata is null
                        || !metadata.Numbered
                        || !string.Equals(
                            metadata.DisplayMode,
                            "block",
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    // Complete VTEq* alias loss makes the normal structural range
                    // resolver intentionally reject the OMath identity. Recover the
                    // physical host from the durable VTOMML anchor plus semantic
                    // fingerprint, without requiring the missing number aliases.
                    formulaRange =
                        TryResolveManagedNativeHashSequenceRangeForAliasRepair(
                            document,
                            formulaId,
                            metadata);
                    if (formulaRange is null)
                        continue;

                    // Formula metadata makes this a trusted managed host. The
                    // structure-only probe intentionally does not require VTEqNum
                    // to still be inside the delimiter, because that missing alias
                    // is exactly what this repair path exists to recover.
                    if (!WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                            formulaRange.WordOpenXML))
                        continue;
                    hosts.Add((formulaId, formulaRange.Start, formulaRange.End));
                    if (!IsHealthyNumberedNativeOmmlHashSequenceHost(
                            document,
                            formulaRange,
                            formulaId,
                            requireCurrentFieldResult: false))
                        repairRequired = true;
                }
                finally { Release(formulaRange); }
            }

            if (!repairRequired || hosts.Count == 0)
                return false;

            var traceRepair = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal);
            if (traceRepair)
                Console.WriteLine(
                    $"    native-hash-identity-repair inventory={hosts.Count} ids=[{string.Join(",", hosts.OrderBy(item => item.Start).Select(item => item.FormulaId + "@" + item.Start + ":" + item.End))}]");

            // Remove aliases for the entire managed native-hash set before any
            // OMath is rebuilt. Bookmark.Delete removes only the identity, never
            // its mathematical contents, and prevents crossing ranges from being
            // re-associated with another formula while the first host is replaced.
            foreach (var host in hosts)
            {
                DeleteBookmarkOnly(document, EquationBookmarkName(host.FormulaId));
                DeleteBookmarkOnly(document, NativeCaptionBookmarkName(host.FormulaId));
                DeleteBookmarkOnly(document, NativeNumberBookmarkName(host.FormulaId));
            }

            // Rebuild end-to-start using the physical OMath ranges frozen above.
            // Replacing a later equation cannot move an earlier Start/End, so this
            // avoids consulting temporarily missing VTEqNum aliases or guessing by
            // bookmark proximity during the repair transaction.
            foreach (var host in hosts.OrderByDescending(item => item.Start))
            {
                Range? formulaRange = null;
                Range? repairedRange = null;
                try
                {
                    var metadata = WordOmmlFormulaStore.TryRead(
                            document,
                            host.FormulaId)
                        ?? throw new InvalidDataException(
                            $"Managed native #(SEQ) formula {host.FormulaId} lost metadata during identity repair.");
                    formulaRange = document.Range(host.Start, host.End);
                    ConfigureNumberedNativeOmmlDisplay(
                        document,
                        formulaRange,
                        WordOmmlFormulaStore.EstimateHeightPoints(formulaRange),
                        (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                        host.FormulaId,
                        reuseExistingScaffold: true,
                        metadata,
                        _ => { });

                    repairedRange = WordOmmlFormulaStore
                        .GetEquationRangeVerifiedForStructuralEdit(
                            document,
                            host.FormulaId,
                            metadata);
                    if (!IsHealthyNumberedNativeOmmlHashSequenceHost(
                            document,
                            repairedRange,
                            host.FormulaId))
                        throw new InvalidDataException(
                            $"Managed native #(SEQ) formula {host.FormulaId} remained unhealthy after atomic identity repair.");
                    repaired++;
                    if (traceRepair)
                    {
                        Bookmarks? repairBookmarks = null;
                        try
                        {
                            repairBookmarks = document.Bookmarks;
                            Console.WriteLine(
                                $"    native-hash-identity-repair rebuilt={host.FormulaId} range={repairedRange.Start}:{repairedRange.End} aliases=visible:{repairBookmarks.Exists(EquationBookmarkName(host.FormulaId))},caption:{repairBookmarks.Exists(NativeCaptionBookmarkName(host.FormulaId))},number:{repairBookmarks.Exists(NativeNumberBookmarkName(host.FormulaId))}");
                        }
                        finally { Release(repairBookmarks); }
                    }
                }
                finally
                {
                    Release(repairedRange);
                    Release(formulaRange);
                }
            }

            if (repaired != hosts.Count)
                throw new InvalidDataException(
                    $"Managed native #(SEQ) identity repair completed {repaired}/{hosts.Count} hosts.");

            // Ordinary body REF / GOTOBUTTON+nested-REF fields kept their bookmark
            // names while the aliases were absent. Refresh them only after every
            // VTEqNum target has been recreated.
            UpdateNativeCrossReferences(document);
            return true;
        }
        catch
        {
            repaired = 0;
            return false;
        }
    }

    private static Range? EnsureNormalTypingParagraphAfterNativeOmmlHashSequence(
        Document document,
        string formulaId)
    {
        Range? ownerRange = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        Paragraphs? ownerParagraphs = null;
        Paragraph? ownerParagraph = null;
        Range? ownerParagraphRange = null;
        Range? content = null;
        Range? probe = null;
        Paragraphs? typingParagraphs = null;
        Paragraph? typingParagraph = null;
        Range? typingParagraphRange = null;
        Tables? typingTables = null;
        Frames? typingFrames = null;
        OMaths? typingMaths = null;
        InlineShapes? typingShapes = null;
        Fields? typingFields = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        ParagraphFormat? format = null;
        try
        {
            ownerRange = FindNumberingOwnerRange(document, formulaId);
            if (ownerRange is null
                || (bool)ownerRange.get_Information(WdInformation.wdWithInTable))
                return null;
            maths = ownerRange.OMaths;
            if (maths.Count != 1) return null;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay) return null;
            mathRange = math.Range.Duplicate;
            if (!WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                    mathRange.WordOpenXML,
                    formulaId))
                return null;

            ownerParagraphs = mathRange.Paragraphs;
            if (ownerParagraphs.Count != 1) return null;
            ownerParagraph = ownerParagraphs[1];
            ownerParagraphRange = ownerParagraph.Range.Duplicate;
            if (mathRange.End + 1 != ownerParagraphRange.End)
                return null;

            content = document.Content;
            var typingStart = ownerParagraphRange.End;
            var createParagraph = typingStart >= content.End;
            if (!createParagraph)
            {
                probe = document.Range(
                    typingStart,
                    Math.Min(content.End, typingStart + 1));
                if ((bool)probe.get_Information(WdInformation.wdWithInTable))
                    createParagraph = true;
                else
                {
                    typingFrames = probe.Frames;
                    if (typingFrames.Count > 0)
                        createParagraph = true;
                    else
                    {
                        typingParagraphs = probe.Paragraphs;
                        if (typingParagraphs.Count != 1)
                            createParagraph = true;
                        else
                        {
                            typingParagraph = typingParagraphs[1];
                            typingParagraphRange = typingParagraph.Range.Duplicate;
                            typingTables = typingParagraphRange.Tables;
                            typingMaths = typingParagraphRange.OMaths;
                            typingShapes = typingParagraphRange.InlineShapes;
                            typingFields = typingParagraphRange.Fields;
                            createParagraph = typingTables.Count != 0
                                || typingMaths.Count != 0
                                || typingShapes.Count != 0
                                || typingFields.Count != 0
                                || !IsNumberingParagraphAdornment(
                                    typingParagraphRange.Text);
                        }
                    }
                }
            }

            if (createParagraph)
            {
                Release(typingFields); typingFields = null;
                Release(typingShapes); typingShapes = null;
                Release(typingMaths); typingMaths = null;
                Release(typingTables); typingTables = null;
                Release(typingParagraphRange); typingParagraphRange = null;
                Release(typingParagraph); typingParagraph = null;
                Release(typingParagraphs); typingParagraphs = null;
                Release(typingFrames); typingFrames = null;
                Release(probe); probe = null;

                // Freeze the old paragraph end before Word expands the Range used
                // by InsertParagraphAfter. The new ordinary paragraph starts at this
                // exact position, immediately after the numbered OMath's own ¶.
                typingStart = ownerParagraphRange.End;
                ownerParagraphRange.InsertParagraphAfter();
                Release(content);
                content = document.Content;
                if (typingStart >= content.End) return null;
                probe = document.Range(
                    typingStart,
                    Math.Min(content.End, typingStart + 1));
                if ((bool)probe.get_Information(WdInformation.wdWithInTable))
                    return null;
                typingFrames = probe.Frames;
                if (typingFrames.Count > 0) return null;
                typingParagraphs = probe.Paragraphs;
                if (typingParagraphs.Count != 1) return null;
                typingParagraph = typingParagraphs[1];
                typingParagraphRange = typingParagraph.Range.Duplicate;
                typingTables = typingParagraphRange.Tables;
                typingMaths = typingParagraphRange.OMaths;
                typingShapes = typingParagraphRange.InlineShapes;
                typingFields = typingParagraphRange.Fields;
                if (typingTables.Count != 0
                    || typingMaths.Count != 0
                    || typingShapes.Count != 0
                    || typingFields.Count != 0
                    || !IsNumberingParagraphAdornment(typingParagraphRange.Text))
                    return null;
            }

            try
            {
                object normalStyle = WdBuiltinStyle.wdStyleNormal;
                typingParagraphRange!.set_Style(ref normalStyle);
            }
            catch
            {
                // A locked/custom style collection may reject Normal. Direct
                // formatting reset below still produces an ordinary typing row.
            }
            font = typingParagraphRange!.Font;
            font.Reset();
            font.Hidden = 0;
            font.Position = 0;
            font.Color = WdColor.wdColorAutomatic;
            format = typingParagraphRange.ParagraphFormat;
            format.Reset();
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            typingParagraphRange.Collapse(WdCollapseDirection.wdCollapseStart);
            var result = typingParagraphRange;
            typingParagraphRange = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(format);
            Release(font);
            Release(typingFields);
            Release(typingShapes);
            Release(typingMaths);
            Release(typingFrames);
            Release(typingTables);
            Release(typingParagraphRange);
            Release(typingParagraph);
            Release(typingParagraphs);
            Release(probe);
            Release(content);
            Release(ownerParagraphRange);
            Release(ownerParagraph);
            Release(ownerParagraphs);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(ownerRange);
        }
    }

    private static void UpdateNativeOmmlHashSequenceField(Range formulaRange)
    {
        Fields? fields = null;
        Field? sequenceField = null;
        Range? code = null;
        try
        {
            fields = formulaRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? candidate = null;
                Range? candidateCode = null;
                try
                {
                    candidate = fields[index];
                    candidateCode = candidate.Code;
                    if (!IsVisualTeXSequenceFieldCode(candidateCode.Text))
                        continue;
                    if (Regex.IsMatch(
                            candidateCode.Text ?? string.Empty,
                            @"\bREF\s+VTEqNum_",
                            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                        continue;
                    sequenceField = candidate;
                    candidate = null;
                    break;
                }
                finally
                {
                    Release(candidateCode);
                    Release(candidate);
                }
            }
            if (sequenceField is null)
                throw new InvalidDataException(
                    "The native #(SEQ) OMML host lost its VisualTeXEquation field.");

            // Critical invariant: never assign Field.Code.Text for a field hosted
            // inside professional OMML. Word can split the instruction into m:r/m:t
            // fragments and corrupt switches such as \\* ARABIC. Number refresh is
            // strictly a normal Word field update on the field code created with the
            // OMath wrapper.
            sequenceField.Update();
        }
        finally
        {
            Release(code);
            Release(sequenceField);
            Release(fields);
        }
    }

    private static bool IsHealthyNumberedNativeOmmlHashSequenceHost(
        Document document,
        Range formulaRange,
        string formulaId,
        bool requireCurrentFieldResult = true)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        Field? sequenceField = null;
        Range? sequenceCode = null;
        Range? sequenceResult = null;
        Bookmarks? bookmarks = null;
        Bookmark? numberBookmark = null;
        Bookmark? visibleBookmark = null;
        Bookmark? captionBookmark = null;
        Range? numberRange = null;
        Range? visibleRange = null;
        Range? captionRange = null;
        Shape? legacyShape = null;
        try
        {
            maths = formulaRange.OMaths;
            if (maths.Count != 1) return false;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay) return false;
            mathRange = math.Range.Duplicate;
            paragraphs = mathRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable))
                return false;
            if (mathRange.End + 1 != paragraphRange.End)
                return false;
            if (!WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                    mathRange.WordOpenXML))
                return false;

            fields = mathRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? candidate = null;
                Range? candidateCode = null;
                try
                {
                    candidate = fields[index];
                    candidateCode = candidate.Code;
                    if (!IsVisualTeXSequenceFieldCode(candidateCode.Text))
                        continue;
                    if (sequenceField is not null) return false;
                    sequenceField = candidate;
                    candidate = null;
                    sequenceCode = candidateCode;
                    candidateCode = null;
                }
                finally
                {
                    Release(candidateCode);
                    Release(candidate);
                }
            }
            if (sequenceField is null || fields.Count != 1)
                return false;
            if ((sequenceCode?.Text ?? string.Empty).IndexOf(
                    "REF ",
                    StringComparison.OrdinalIgnoreCase) >= 0)
                return false;
            sequenceResult = sequenceField.Result;
            if (requireCurrentFieldResult
                && string.IsNullOrWhiteSpace(
                    NormalizeNativeEquationNumberText(sequenceResult.Text)))
                return false;

            bookmarks = document.Bookmarks;
            var numberName = NativeNumberBookmarkName(formulaId);
            var visibleName = EquationBookmarkName(formulaId);
            var captionName = NativeCaptionBookmarkName(formulaId);
            if (!bookmarks.Exists(numberName)
                || !bookmarks.Exists(visibleName)
                || !bookmarks.Exists(captionName))
                return false;
            numberBookmark = bookmarks[numberName];
            visibleBookmark = bookmarks[visibleName];
            captionBookmark = bookmarks[captionName];
            numberRange = numberBookmark.Range;
            visibleRange = visibleBookmark.Range;
            captionRange = captionBookmark.Range;
            if (numberRange.StoryType != WdStoryType.wdMainTextStory
                || numberRange.Start < mathRange.Start
                || numberRange.End > mathRange.End
                || visibleRange.Start < mathRange.Start
                || visibleRange.End > mathRange.End
                || captionRange.Start < mathRange.Start
                || captionRange.End > mathRange.End)
                return false;

            legacyShape = FindNativeDisplayNumberShape(document, formulaId);
            if (legacyShape is not null) return false;

            var xml = paragraphRange.WordOpenXML ?? string.Empty;
            return xml.IndexOf("<m:oMathPara", StringComparison.OrdinalIgnoreCase) >= 0
                && xml.IndexOf("<m:eqArr", StringComparison.OrdinalIgnoreCase) >= 0
                && xml.IndexOf("<w:fldChar", StringComparison.OrdinalIgnoreCase) >= 0
                && xml.IndexOf(
                    "SEQ " + LegacyEquationSequenceName,
                    StringComparison.OrdinalIgnoreCase) >= 0
                && xml.IndexOf(
                    "REF " + NativeNumberBookmarkPrefix,
                    StringComparison.OrdinalIgnoreCase) < 0
                && xml.IndexOf("<w:txbxContent", StringComparison.OrdinalIgnoreCase) < 0;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(legacyShape);
            Release(captionRange);
            Release(visibleRange);
            Release(numberRange);
            Release(captionBookmark);
            Release(visibleBookmark);
            Release(numberBookmark);
            Release(bookmarks);
            Release(sequenceResult);
            Release(sequenceCode);
            Release(sequenceField);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(mathRange);
            Release(math);
            Release(maths);
        }
    }

    private static void TraceNativeOmmlHashSequenceDiagnostics(
        Document document,
        Range formulaRange,
        string formulaId,
        string context)
    {
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            return;
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? trailingRange = null;
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        try
        {
            maths = formulaRange.OMaths;
            math = maths.Count > 0 ? maths[1] : null;
            mathRange = math?.Range?.Duplicate;
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs.Count > 0 ? paragraphs[1] : null;
            paragraphRange = paragraph?.Range;
            fields = formulaRange.Fields;
            bookmarks = document.Bookmarks;
            var codes = new List<string>();
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    codes.Add((code.Text ?? string.Empty).Replace("\r", "\\r"));
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            var numberName = NativeNumberBookmarkName(formulaId);
            var visibleName = EquationBookmarkName(formulaId);
            var captionName = NativeCaptionBookmarkName(formulaId);
            var direct = false;
            var xml = formulaRange.WordOpenXML ?? string.Empty;
            var paragraphText = (paragraphRange?.Text ?? string.Empty)
                .Replace("\r", "\\r")
                .Replace("\n", "\\n")
                .Replace("\t", "\\t")
                .Replace("\v", "\\v")
                .Replace("\u0013", "<FIELD_BEGIN>")
                .Replace("\u0014", "<FIELD_SEPARATE>")
                .Replace("\u0015", "<FIELD_END>");
            var trailingText = string.Empty;
            if (mathRange is not null && paragraphRange is not null)
            {
                trailingRange = document.Range(
                    Math.Min(mathRange.End, paragraphRange.End),
                    paragraphRange.End);
                trailingText = (trailingRange.Text ?? string.Empty)
                    .Replace("\r", "\\r")
                    .Replace("\n", "\\n")
                    .Replace("\t", "\\t")
                    .Replace("\v", "\\v")
                    .Replace("\u0013", "<FIELD_BEGIN>")
                    .Replace("\u0014", "<FIELD_SEPARATE>")
                    .Replace("\u0015", "<FIELD_END>");
            }
            try
            {
                direct = WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                    xml,
                    formulaId);
            }
            catch { }
            Console.WriteLine(
                $"    native-hash-diagnostic context={context} formulaId={formulaId} maths={maths.Count} type={(math is null ? "none" : math.Type.ToString())} range={formulaRange.Start}:{formulaRange.End} mathRange={(mathRange is null ? "none" : mathRange.Start + ":" + mathRange.End)} paragraph={(paragraphRange is null ? "none" : paragraphRange.Start + ":" + paragraphRange.End)} paragraphText='{paragraphText}' trailing='{trailingText}' fields={fields.Count} codes=[{string.Join(" | ", codes)}] direct={direct} bookmarks=num:{bookmarks.Exists(numberName)},visible:{bookmarks.Exists(visibleName)},caption:{bookmarks.Exists(captionName)} xml={xml}");
        }
        catch (Exception error)
        {
            Console.WriteLine(
                $"    native-hash-diagnostic context={context} formulaId={formulaId} failed={error.GetType().Name}:{error.Message}");
        }
        finally
        {
            Release(bookmarks);
            Release(fields);
            Release(trailingRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(mathRange);
            Release(math);
            Release(maths);
        }
    }

    private static void RemoveLegacyNativeOmmlNumberingScaffold(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Range? captionRange = null;
        try
        {
            // This removes only the old VTEqShape_/VTEqAnc_ host when present. It
            // is deliberately retained for migration and is never called after the
            // new mathematical host has passed its health check above.
            RemoveNativeDisplayNumberShapeAndAnchor(document, formulaId);

            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            if (bookmarks.Exists(captionName))
            {
                captionBookmark = bookmarks[captionName];
                captionRange = captionBookmark.Range;
                var captionIsInsideMath =
                    captionRange.StoryType == formulaRange.StoryType
                    && captionRange.Start >= formulaRange.Start
                    && captionRange.End <= formulaRange.End;
                if (!captionIsInsideMath)
                    RemoveNativeCaption(document, formulaId);
            }

            // Any remaining aliases are inside the mathematical wrapper that will
            // be replaced atomically. Delete only the bookmarks, never their text.
            DeleteBookmarkOnly(document, EquationBookmarkName(formulaId));
            DeleteBookmarkOnly(document, NativeCaptionBookmarkName(formulaId));
            DeleteBookmarkOnly(document, NativeNumberBookmarkName(formulaId));
            // Do not clear VTAncR_ here. When Word delays removal of the retired
            // one-point Shape anchor paragraph, that marker is the durable evidence
            // used by the next finalization turn. The anchor cleanup path clears it
            // only after the physical paragraph has actually disappeared.
        }
        finally
        {
            Release(captionRange);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    internal static void CreateLegacyNumberedNativeOmmlShapeFixtureForAcceptance(
        Document document,
        Range formulaRange,
        FormulaMetadata metadata)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (formulaRange is null) throw new ArgumentNullException(nameof(formulaRange));
        if (metadata is null) throw new ArgumentNullException(nameof(metadata));
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            throw new InvalidOperationException(
                "The legacy numbered-OMML Shape producer is available only to the isolated VSTO acceptance process.");

        Shape? shape = null;
        Range? activeRange = null;
        Range? anchoredRange = null;
        Bookmark? repairedBookmark = null;
        try
        {
            // Construct the retired host explicitly instead of routing through the
            // production numbered-OMML producer. Production now permanently refuses
            // to create an external caption or Shape for native OMML; this bypass is
            // acceptance-only so migration remains testable without weakening that
            // invariant.
            activeRange = ResolveSingleNativeOmmlRange(formulaRange);
            anchoredRange = InsertNativeDisplayShapeAnchorBeforeExistingFormula(
                document,
                activeRange,
                metadata.FormulaId);
            Release(activeRange);
            activeRange = anchoredRange;
            anchoredRange = null;

            CreateNativeCaption(
                document,
                activeRange,
                metadata.FormulaId,
                GetNativeEquationSequenceName(document),
                knownNumberedTable: null,
                deferFieldUpdate: false,
                plannedOrdinal: null,
                plannedPrefix: null,
                allowNativeOmmlAcceptanceFixture: true);
            EnsureNativeDisplayNumberShape(
                document,
                activeRange,
                WordOmmlFormulaStore.EstimateHeightPoints(activeRange),
                (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                metadata.FormulaId,
                deferFieldUpdate: false);

            repairedBookmark = WordOmmlFormulaStore.Wrap(
                document,
                activeRange,
                metadata,
                replaceExisting: true);
            WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                metadata,
                activeRange);
            WordOmmlFormulaStore.Save(document, metadata);
            formulaRange.SetRange(activeRange.Start, activeRange.End);

            shape = FindNativeDisplayNumberShape(document, metadata.FormulaId);
            if (shape is null)
                throw new InvalidOperationException(
                    "The acceptance-only legacy producer did not create VTEqShape_ numbering.");
        }
        finally
        {
            Release(repairedBookmark);
            Release(anchoredRange);
            Release(activeRange);
            Release(shape);
        }
    }

    // Retained only as a migration implementation reference while old documents
    // carrying VTEqShape_/TextBox numbering are accepted and dismantled. Production
    // numbered OMML never calls this Shape host.
    private static bool ConfigureNumberedNativeOmmlDisplayLegacyShape(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId,
        bool reuseExistingScaffold,
        FormulaMetadata? metadata,
        Action<string> traceStage,
        int? plannedOrdinal = null,
        string? plannedPrefix = null,
        bool deferFieldUpdate = false,
        bool deferExternalShapeCreation = false)
    {
        // A mathematical SEQ must remain order-driven. Batch planning may know the
        // current ordinal, but encoding that value as \\r N would reset the sequence
        // at every formula and make later F9/middle insertion unable to renumber the
        // suffix. Use the planned value only as transient display knowledge outside
        // this native path; the OMML field instruction is always created without
        // an ordinal restart switch.
        plannedOrdinal = null;

        Range? activeRange = null;
        Range? migratedRange = null;
        Range? replacementRange = null;
        Bookmark? repairedBookmark = null;
        Microsoft.Office.Interop.Word.Application? application = null;
        try
        {
            metadata ??= WordOmmlFormulaStore.TryRead(document, formulaId);
            activeRange = ResolveSingleNativeOmmlRange(formulaRange);

            // Shape/TextBox numbered OMML is migration input only. Even a complete
            // legacy host must never short-circuit reconciliation, otherwise old
            // documents would preserve draggable numbering forever. Record the
            // detection here; the cleanup/replacement path below removes the Shape
            // and anchor before atomically materializing native #(SEQ).
            var legacyShapeHostDetected =
                HasStructurallyReusableNumberedNativeOmmlDisplayHost(
                    document,
                    activeRange,
                    formulaId)
                || IsHealthyNumberedNativeOmmlTrueDisplayHost(
                    document,
                    activeRange,
                    formulaId);
            if (legacyShapeHostDetected)
                traceStage("migrate-legacy-true-display-shape");

            var semanticOmml = WordOmmlConverter
                .StripManagedVisualTeXNativeEquationNumber(
                    activeRange.WordOpenXML);
            if (WordOmmlConverter.HasVisualTeXNativeEquationNumber(activeRange.WordOpenXML))
                traceStage("strip-native-number-wrapper");

            if (IsNumberedEquationTable(activeRange))
            {
                TrimBenignEmptyRowsFromNumberedTable(document, activeRange, formulaId);
                migratedRange = TryConvertStandardNumberedOmmlTableToStandaloneDisplayParagraph(
                    document,
                    activeRange,
                    formulaId,
                    metadata);
                if (migratedRange is null)
                    throw new InvalidOperationException(
                        "VisualTeX could not safely migrate this managed numbered OMML table to the native display host.");
                Release(activeRange);
                activeRange = ResolveSingleNativeOmmlRange(migratedRange);
                traceStage("migrate-omml-table-to-true-display-host");
            }

            var preserveExistingTrueDisplayBody =
                IsPureTrueDisplayFormulaParagraph(activeRange);
            var preparedTrueDisplayAnchor = preserveExistingTrueDisplayBody
                && !legacyShapeHostDetected
                && HasPreparedNativeDisplayShapeAnchor(
                    document,
                    activeRange,
                    formulaId);
            var visibleNumberAlreadyExisted = HasVisibleEquationNumberBookmark(
                document,
                formulaId);
            if (!preparedTrueDisplayAnchor)
            {
                RemoveNativeDisplayNumberShapeAndAnchor(document, formulaId);
                if (HasVisibleEquationNumberBookmark(document, formulaId))
                    RemoveVisibleEquationNumberLegacy(document, formulaId);
            }

            if (preserveExistingTrueDisplayBody)
            {
                if (!preparedTrueDisplayAnchor)
                {
                    // Fresh block insertion already materializes the mathematical
                    // body as genuine m:oMathPara/wdOMathDisplay. Creating the
                    // external number host must not replace that OMath a second time
                    // merely to obtain its preceding Shape-anchor paragraph.
                    replacementRange = InsertNativeDisplayShapeAnchorBeforeExistingFormula(
                        document,
                        activeRange,
                        formulaId);
                    Release(activeRange);
                    activeRange = replacementRange;
                    replacementRange = null;
                    traceStage("preserve-true-display-body");
                }
                else
                {
                    traceStage("reuse-prepared-true-display-host");
                }
                EnsureNumberedOmmlIsDisplay(activeRange);
                ConfigureEquationParagraph(activeRange, numbered: false);
            }
            else
            {
                application = document.Application;
                replacementRange = ReplaceNumberedOmmlWithTrueDisplayHost(
                    application,
                    document,
                    activeRange,
                    semanticOmml,
                    formulaId);
                Release(activeRange);
                activeRange = replacementRange;
                replacementRange = null;
                EnsureNumberedOmmlIsDisplay(activeRange);
                ConfigureEquationParagraph(activeRange, numbered: false);
                traceStage("materialize-true-display-host");
            }

            var nativeSequenceName = GetNativeEquationSequenceName(document);
            if (plannedOrdinal.HasValue || !string.IsNullOrEmpty(plannedPrefix))
            {
                RemoveNativeCaption(document, formulaId);
                CreateNativeCaption(
                    document,
                    activeRange,
                    formulaId,
                    nativeSequenceName,
                    knownNumberedTable: null,
                    deferFieldUpdate: deferFieldUpdate,
                    plannedOrdinal: plannedOrdinal,
                    plannedPrefix: plannedPrefix);
            }
            else
            {
                EnsureNativeCaption(
                    document,
                    activeRange,
                    formulaId,
                    nativeSequenceName,
                    restyleExisting: !reuseExistingScaffold,
                    knownNumberedTable: null);
            }
            traceStage("native-caption");

            if (metadata is not null)
            {
                repairedBookmark = WordOmmlFormulaStore.Wrap(
                    document,
                    activeRange,
                    metadata,
                    replaceExisting: true);
                WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                    metadata,
                    activeRange);
                WordOmmlFormulaStore.Save(document, metadata);
            }

            // Compatibility parameters remain in the private signature while
            // older callers are migrated, but numbered native OMML no longer has an
            // external visible REF host under any mode. The SEQ result inside #()
            // is already the Word-native visible number.
            _ = deferExternalShapeCreation;
            traceStage("native-hash-sequence-complete");
            formulaRange.SetRange(activeRange.Start, activeRange.End);
            return !visibleNumberAlreadyExisted;
        }
        finally
        {
            Release(application);
            Release(repairedBookmark);
            Release(replacementRange);
            Release(migratedRange);
            Release(activeRange);
        }
    }

    internal static Range PrepareNumberedNativeOmmlInsertionHost(
        Document document,
        Range displayInsertion,
        string formulaId)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (displayInsertion is null) throw new ArgumentNullException(nameof(displayInsertion));
        _ = NormalizeFormulaIdForBookmark(formulaId);
        if ((bool)displayInsertion.get_Information(WdInformation.wdWithInTable))
            throw new InvalidOperationException(
                "A numbered OMML display formula cannot be inserted inside a Word table.");
        if (displayInsertion.StoryType != WdStoryType.wdMainTextStory)
            throw new InvalidOperationException(
                "A numbered OMML display formula requires the main Word document story.");

        // Retained only as a source-compatible entry for older callers. The former
        // implementation inserted a dedicated one-point Shape-anchor paragraph and
        // a private-use replacement placeholder. Native #() needs neither: return a
        // collapsed ordinary insertion Range and let Word materialize m:oMathPara
        // directly at the requested location.
        var result = displayInsertion.Duplicate;
        result.Collapse(WdCollapseDirection.wdCollapseStart);
        return result;
    }

    private static bool HasPreparedNativeDisplayShapeAnchor(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? anchorBookmark = null;
        Range? anchorBookmarkRange = null;
        Paragraphs? anchorParagraphs = null;
        Paragraph? anchorParagraph = null;
        Range? anchorParagraphRange = null;
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        OMaths? anchorMaths = null;
        Fields? anchorFields = null;
        InlineShapes? anchorInlineShapes = null;
        try
        {
            if (!IsPureTrueDisplayFormulaParagraph(formulaRange)) return false;
            bookmarks = document.Bookmarks;
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            if (!bookmarks.Exists(anchorName)
                || bookmarks.Exists(EquationBookmarkName(formulaId))
                || bookmarks.Exists(NativeCaptionBookmarkName(formulaId))
                || bookmarks.Exists(NativeNumberBookmarkName(formulaId)))
                return false;

            anchorBookmark = bookmarks[anchorName];
            anchorBookmarkRange = anchorBookmark.Range;
            anchorParagraphs = anchorBookmarkRange.Paragraphs;
            formulaParagraphs = formulaRange.Paragraphs;
            if (anchorParagraphs.Count != 1 || formulaParagraphs.Count != 1)
                return false;
            anchorParagraph = anchorParagraphs[1];
            formulaParagraph = formulaParagraphs[1];
            anchorParagraphRange = anchorParagraph.Range;
            formulaParagraphRange = formulaParagraph.Range;
            if ((bool)anchorParagraphRange.get_Information(WdInformation.wdWithInTable)
                || anchorParagraphRange.End != formulaParagraphRange.Start
                || anchorParagraphRange.Start >= formulaParagraphRange.Start
                || !IsNumberingParagraphAdornment(anchorParagraphRange.Text))
                return false;
            anchorMaths = anchorParagraphRange.OMaths;
            anchorFields = anchorParagraphRange.Fields;
            anchorInlineShapes = anchorParagraphRange.InlineShapes;
            return anchorMaths.Count == 0
                && anchorFields.Count == 0
                && anchorInlineShapes.Count == 0;
        }
        catch { return false; }
        finally
        {
            Release(anchorInlineShapes);
            Release(anchorFields);
            Release(anchorMaths);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(anchorParagraphRange);
            Release(anchorParagraph);
            Release(anchorParagraphs);
            Release(anchorBookmarkRange);
            Release(anchorBookmark);
            Release(bookmarks);
        }
    }

    private static Range InsertNativeDisplayShapeAnchorBeforeExistingFormula(
        Document document,
        Range sourceFormulaRange,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? temporaryBookmark = null;
        Range? temporaryBookmarkRange = null;
        OMaths? temporaryMaths = null;
        OMath? temporaryMath = null;
        Range? formulaRange = null;
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Range? anchorProbe = null;
        Paragraphs? anchorParagraphs = null;
        Paragraph? anchorParagraph = null;
        Range? anchorParagraphRange = null;
        Range? anchorBookmarkRange = null;
        OMaths? anchorMaths = null;
        Fields? anchorFields = null;
        InlineShapes? anchorInlineShapes = null;
        var temporaryBookmarkName = "VTTmp_" + Guid.NewGuid().ToString("N");
        try
        {
            formulaRange = sourceFormulaRange.Duplicate;
            if (!IsPureTrueDisplayFormulaParagraph(formulaRange))
                throw new InvalidOperationException(
                    $"Formula {formulaId} is not a pure m:oMathPara before Shape-anchor creation.");
            formulaParagraphs = formulaRange.Paragraphs;
            if (formulaParagraphs.Count != 1)
                throw new InvalidOperationException(
                    $"Formula {formulaId} does not occupy one display paragraph.");
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range.Duplicate;
            if ((bool)formulaParagraphRange.get_Information(
                    WdInformation.wdWithInTable))
                throw new InvalidOperationException(
                    $"Formula {formulaId} cannot receive a Shape anchor inside a table.");

            // The old-table migration has intentionally removed the stale VTOMML
            // bookmark before reaching this point. Track the already-materialized
            // pure Display OMath only with a short-lived bookmark while Word inserts
            // the preceding anchor paragraph; the durable FormulaId bookmark is
            // written later after the complete physical host is established.
            bookmarks = document.Bookmarks;
            temporaryBookmark = bookmarks.Add(
                temporaryBookmarkName,
                formulaRange);

            formulaParagraphRange.InsertParagraphBefore();

            Release(formulaParagraphRange); formulaParagraphRange = null;
            Release(formulaParagraph); formulaParagraph = null;
            Release(formulaParagraphs); formulaParagraphs = null;
            Release(formulaRange); formulaRange = null;
            Release(temporaryBookmark); temporaryBookmark = null;
            Release(bookmarks); bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(temporaryBookmarkName))
                throw new InvalidOperationException(
                    $"Word lost the temporary genuine-display anchor for formula {formulaId}.");
            temporaryBookmark = bookmarks[temporaryBookmarkName];
            temporaryBookmarkRange = temporaryBookmark.Range;
            temporaryMaths = temporaryBookmarkRange.OMaths;
            if (temporaryMaths.Count != 1)
                throw new InvalidOperationException(
                    $"Word did not preserve exactly one OMath for formula {formulaId} while adding its Shape anchor paragraph.");
            temporaryMath = temporaryMaths[1];
            formulaRange = temporaryMath.Range.Duplicate;
            if (!IsPureTrueDisplayFormulaParagraph(formulaRange))
                throw new InvalidOperationException(
                    $"Word downgraded genuine-display formula {formulaId} while adding its Shape anchor paragraph.");
            formulaParagraphs = formulaRange.Paragraphs;
            if (formulaParagraphs.Count != 1)
                throw new InvalidOperationException(
                    $"Formula {formulaId} no longer occupies one display paragraph after anchor creation.");
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range;
            if (formulaParagraphRange.Start <= 0)
                throw new InvalidOperationException(
                    $"Word did not create a preceding paragraph for formula {formulaId}.");

            anchorProbe = document.Range(
                formulaParagraphRange.Start - 1,
                formulaParagraphRange.Start);
            anchorParagraphs = anchorProbe.Paragraphs;
            if (anchorParagraphs.Count != 1)
                throw new InvalidOperationException(
                    $"The Shape anchor for formula {formulaId} is not one Word paragraph.");
            anchorParagraph = anchorParagraphs[1];
            anchorParagraphRange = anchorParagraph.Range;
            if ((bool)anchorParagraphRange.get_Information(
                    WdInformation.wdWithInTable)
                || anchorParagraphRange.End != formulaParagraphRange.Start)
                throw new InvalidOperationException(
                    $"The Shape anchor for formula {formulaId} is not the adjacent table-free paragraph.");
            anchorMaths = anchorParagraphRange.OMaths;
            anchorFields = anchorParagraphRange.Fields;
            anchorInlineShapes = anchorParagraphRange.InlineShapes;
            var anchorText = (anchorParagraphRange.Text ?? string.Empty)
                .Trim('\r', '\n', '\t', '\v', '\u0001', ' ');
            if (anchorText.Length != 0
                || anchorMaths.Count != 0
                || anchorFields.Count != 0
                || anchorInlineShapes.Count != 0)
                throw new InvalidOperationException(
                    $"The new Shape anchor paragraph for formula {formulaId} contains unexpected content.");

            StyleNativeDisplayAnchorParagraph(anchorParagraphRange);
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            DeleteBookmarkOnly(document, anchorName);
            anchorBookmarkRange = document.Range(
                anchorParagraphRange.Start,
                anchorParagraphRange.Start);
            bookmarks.Add(anchorName, anchorBookmarkRange);

            temporaryBookmark.Delete();
            temporaryBookmark = null;
            var result = formulaRange;
            formulaRange = null;
            return result;
        }
        finally
        {
            try { DeleteBookmarkOnly(document, temporaryBookmarkName); } catch { }
            Release(anchorInlineShapes);
            Release(anchorFields);
            Release(anchorMaths);
            Release(anchorBookmarkRange);
            Release(anchorParagraphRange);
            Release(anchorParagraph);
            Release(anchorParagraphs);
            Release(anchorProbe);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(formulaRange);
            Release(temporaryMath);
            Release(temporaryMaths);
            Release(temporaryBookmarkRange);
            Release(temporaryBookmark);
            Release(bookmarks);
        }
    }

    private static Range InsertNativeDisplayShapeAnchorBeforeExistingFormula(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? formulaBookmark = null;
        Range? formulaRange = null;
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Range? anchorProbe = null;
        Paragraphs? anchorParagraphs = null;
        Paragraph? anchorParagraph = null;
        Range? anchorParagraphRange = null;
        Range? anchorBookmarkRange = null;
        OMaths? anchorMaths = null;
        Fields? anchorFields = null;
        InlineShapes? anchorInlineShapes = null;
        try
        {
            bookmarks = document.Bookmarks;
            var formulaName = WordOmmlFormulaStore.BookmarkName(formulaId);
            if (!bookmarks.Exists(formulaName))
                throw new InvalidOperationException(
                    $"The genuine-display formula bookmark for {formulaId} is missing.");
            formulaBookmark = bookmarks[formulaName];
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            if (!IsPureTrueDisplayFormulaParagraph(formulaRange))
                throw new InvalidOperationException(
                    $"Formula {formulaId} is not a pure m:oMathPara before Shape-anchor creation.");
            formulaParagraphs = formulaRange.Paragraphs;
            if (formulaParagraphs.Count != 1)
                throw new InvalidOperationException(
                    $"Formula {formulaId} does not occupy one display paragraph.");
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range.Duplicate;
            if ((bool)formulaParagraphRange.get_Information(
                    WdInformation.wdWithInTable))
                throw new InvalidOperationException(
                    $"Formula {formulaId} cannot receive a Shape anchor inside a table.");

            // Build the final physical host before any numbering drawing/field is
            // materialized. The existing m:oMathPara is not deleted or copied.
            formulaParagraphRange.InsertParagraphBefore();

            Release(formulaParagraphRange);
            formulaParagraphRange = null;
            Release(formulaParagraph);
            formulaParagraph = null;
            Release(formulaParagraphs);
            formulaParagraphs = null;
            Release(formulaRange);
            formulaRange = null;
            Release(formulaBookmark);
            formulaBookmark = null;
            Release(bookmarks);
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(formulaName))
                throw new InvalidOperationException(
                    $"Word lost genuine-display formula {formulaId} while adding its Shape anchor paragraph.");
            formulaBookmark = bookmarks[formulaName];
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            if (!IsPureTrueDisplayFormulaParagraph(formulaRange))
                throw new InvalidOperationException(
                    $"Word downgraded genuine-display formula {formulaId} while adding its Shape anchor paragraph.");
            formulaParagraphs = formulaRange.Paragraphs;
            if (formulaParagraphs.Count != 1)
                throw new InvalidOperationException(
                    $"Formula {formulaId} no longer occupies one display paragraph after anchor creation.");
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range;
            if (formulaParagraphRange.Start <= 0)
                throw new InvalidOperationException(
                    $"Word did not create a preceding paragraph for formula {formulaId}.");

            anchorProbe = document.Range(
                formulaParagraphRange.Start - 1,
                formulaParagraphRange.Start);
            anchorParagraphs = anchorProbe.Paragraphs;
            if (anchorParagraphs.Count != 1)
                throw new InvalidOperationException(
                    $"The Shape anchor for formula {formulaId} is not one Word paragraph.");
            anchorParagraph = anchorParagraphs[1];
            anchorParagraphRange = anchorParagraph.Range;
            if ((bool)anchorParagraphRange.get_Information(
                    WdInformation.wdWithInTable)
                || anchorParagraphRange.End != formulaParagraphRange.Start)
                throw new InvalidOperationException(
                    $"The Shape anchor for formula {formulaId} is not the adjacent table-free paragraph.");
            anchorMaths = anchorParagraphRange.OMaths;
            anchorFields = anchorParagraphRange.Fields;
            anchorInlineShapes = anchorParagraphRange.InlineShapes;
            var anchorText = (anchorParagraphRange.Text ?? string.Empty)
                .Trim('\r', '\n', '\t', '\v', '\u0001', ' ');
            if (anchorText.Length != 0
                || anchorMaths.Count != 0
                || anchorFields.Count != 0
                || anchorInlineShapes.Count != 0)
                throw new InvalidOperationException(
                    $"The new Shape anchor paragraph for formula {formulaId} contains unexpected content.");

            StyleNativeDisplayAnchorParagraph(anchorParagraphRange);
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            DeleteBookmarkOnly(document, anchorName);
            anchorBookmarkRange = document.Range(
                anchorParagraphRange.Start,
                anchorParagraphRange.Start);
            bookmarks.Add(anchorName, anchorBookmarkRange);

            var result = formulaRange;
            formulaRange = null;
            return result;
        }
        finally
        {
            Release(anchorInlineShapes);
            Release(anchorFields);
            Release(anchorMaths);
            Release(anchorBookmarkRange);
            Release(anchorParagraphRange);
            Release(anchorParagraph);
            Release(anchorParagraphs);
            Release(anchorProbe);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(formulaRange);
            Release(formulaBookmark);
            Release(bookmarks);
        }
    }

    private static Range ReplaceNumberedOmmlWithTrueDisplayHost(
        Microsoft.Office.Interop.Word.Application application,
        Document document,
        Range sourceMathRange,
        string semanticOmml,
        string formulaId)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? editableRange = null;
        Range? placeholderRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? temporaryBookmark = null;
        Range? temporaryRange = null;
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Range? anchorProbe = null;
        Paragraphs? anchorParagraphs = null;
        Paragraph? anchorParagraph = null;
        Range? anchorParagraphRange = null;
        Range? anchorBookmarkRange = null;
        Range? replacement = null;
        OMaths? paragraphMaths = null;
        InlineShapes? inlineShapes = null;
        try
        {
            paragraphs = sourceMathRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "A numbered OMML formula must occupy exactly one source paragraph before native-display migration.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable))
                throw new InvalidOperationException(
                    "The native display host cannot be materialized inside a Word table.");
            paragraphMaths = paragraphRange.OMaths;
            inlineShapes = paragraphRange.InlineShapes;
            if (paragraphMaths.Count != 1 || inlineShapes.Count != 0)
                throw new InvalidOperationException(
                    "VisualTeX refused to rebuild a numbered OMML paragraph containing another equation or inline object.");

            var paragraphStart = paragraphRange.Start;
            var editableEnd = Math.Max(paragraphStart, paragraphRange.End - 1);
            editableRange = document.Range(paragraphStart, editableEnd);
            const string placeholderText = "\uE000";
            editableRange.Text = placeholderText;
            placeholderRange = document.Range(paragraphStart, paragraphStart + 1);

            var temporaryBookmarkName = "VTTmp_" + Guid.NewGuid().ToString("N");
            bookmarks = document.Bookmarks;
            temporaryBookmark = bookmarks.Add(temporaryBookmarkName, placeholderRange);
            Release(temporaryBookmark);
            temporaryBookmark = null;

            Release(paragraphRange);
            paragraphRange = paragraph.Range.Duplicate;
            paragraphRange.InsertParagraphBefore();

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(temporaryBookmarkName))
                throw new InvalidOperationException(
                    "Word lost the numbered OMML placeholder while creating its dedicated Shape anchor paragraph.");
            temporaryBookmark = bookmarks[temporaryBookmarkName];
            temporaryRange = temporaryBookmark.Range;
            var temporaryText = temporaryRange.Text ?? string.Empty;
            var placeholderOffset = temporaryText.IndexOf(
                placeholderText,
                StringComparison.Ordinal);
            if (placeholderOffset < 0)
                throw new InvalidOperationException(
                    "Word lost the numbered OMML placeholder character while inserting its Shape anchor paragraph.");
            temporaryRange.SetRange(
                temporaryRange.Start + placeholderOffset,
                temporaryRange.Start + placeholderOffset + placeholderText.Length);
            formulaParagraphs = temporaryRange.Paragraphs;
            if (formulaParagraphs.Count != 1)
                throw new InvalidOperationException(
                    "The numbered OMML placeholder no longer belongs to one formula paragraph.");
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range;
            if (formulaParagraphRange.Start <= 0)
                throw new InvalidOperationException(
                    "Word did not create a paragraph before the numbered OMML formula.");

            anchorProbe = document.Range(
                formulaParagraphRange.Start - 1,
                formulaParagraphRange.Start);
            anchorParagraphs = anchorProbe.Paragraphs;
            if (anchorParagraphs.Count != 1)
                throw new InvalidOperationException(
                    "The numbered OMML Shape anchor is not a single Word paragraph.");
            anchorParagraph = anchorParagraphs[1];
            anchorParagraphRange = anchorParagraph.Range;
            StyleNativeDisplayAnchorParagraph(anchorParagraphRange);

            DeleteBookmarkOnly(document, NativeDisplayAnchorBookmarkName(formulaId));
            anchorBookmarkRange = document.Range(
                anchorParagraphRange.Start,
                anchorParagraphRange.Start);
            bookmarks.Add(
                NativeDisplayAnchorBookmarkName(formulaId),
                anchorBookmarkRange);

            string mathFontName;
            try { mathFontName = document.OMathFontName ?? string.Empty; }
            catch { mathFontName = string.Empty; }
            replacement = WordOmmlConverter.ReplaceWithPreparedOmml(
                application,
                document,
                temporaryRange,
                semanticOmml,
                display: true,
                mathFontName);
            temporaryBookmark.Delete();
            temporaryBookmark = null;
            EnsureNumberedOmmlIsDisplay(replacement);
            AssertPureTrueDisplayFormulaParagraph(replacement, formulaId);

            var result = replacement;
            replacement = null;
            return result;
        }
        finally
        {
            Release(replacement);
            Release(anchorBookmarkRange);
            Release(anchorParagraphRange);
            Release(anchorParagraph);
            Release(anchorParagraphs);
            Release(anchorProbe);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(temporaryRange);
            Release(temporaryBookmark);
            Release(bookmarks);
            Release(placeholderRange);
            Release(editableRange);
            Release(inlineShapes);
            Release(paragraphMaths);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void StyleNativeDisplayAnchorParagraph(Range anchorParagraphRange)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        ParagraphFormat? paragraphFormat = null;
        ListFormat? listFormat = null;
        Borders? borders = null;
        TabStops? tabStops = null;
        try
        {
            font = anchorParagraphRange.Font;
            font.Hidden = 0;
            font.Size = NativeDisplayAnchorLineSpacingPoints;
            font.Color = WdColor.wdColorAutomatic;
            font.Position = 0;

            paragraphFormat = anchorParagraphRange.ParagraphFormat;
            paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            paragraphFormat.SpaceBefore = 0f;
            paragraphFormat.SpaceAfter = 0f;
            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            paragraphFormat.LineSpacing = NativeDisplayAnchorLineSpacingPoints;
            paragraphFormat.KeepTogether = -1;
            paragraphFormat.KeepWithNext = -1;
            paragraphFormat.WidowControl = 0;
            paragraphFormat.PageBreakBefore = 0;
            try
            {
                listFormat = anchorParagraphRange.ListFormat;
                listFormat.RemoveNumbers(WdNumberType.wdNumberParagraph);
            }
            catch { }
            tabStops = paragraphFormat.TabStops;
            tabStops.ClearAll();
            borders = anchorParagraphRange.Borders;
            borders.Enable = 0;
        }
        finally
        {
            Release(tabStops);
            Release(borders);
            Release(listFormat);
            Release(paragraphFormat);
            Release(font);
        }
    }

    private static bool EnsureNativeDisplayNumberShape(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId,
        bool deferFieldUpdate)
    {
        var visibleBookmarkName = EquationBookmarkName(formulaId);
        var created = !HasVisibleEquationNumberBookmark(document, formulaId);
        Shape? existingShape = null;
        try
        {
            existingShape = FindNativeDisplayNumberShape(document, formulaId);
            if (existingShape is not null
                && IsNativeDisplayNumberShapeFieldHealthy(
                    document,
                    existingShape,
                    formulaId))
            {
                StyleNativeDisplayNumberShape(
                    document,
                    existingShape,
                    formulaRange,
                    formulaHeightPoints,
                    formulaFontSizePoints,
                    formulaId,
                    updateField: !deferFieldUpdate);
                return created;
            }
        }
        finally { Release(existingShape); }

        RemoveNativeDisplayNumberShapeAndAnchor(
            document,
            formulaId,
            removeAnchorParagraph: false);
        DeleteBookmarkOnly(document, visibleBookmarkName);

        Bookmarks? bookmarks = null;
        Bookmark? anchorBookmark = null;
        Range? anchorRange = null;
        Paragraphs? anchorParagraphs = null;
        Paragraph? anchorParagraph = null;
        Range? anchorParagraphRange = null;
        Shapes? shapes = null;
        Shape? shape = null;
        TextFrame? textFrame = null;
        Range? textRange = null;
        Range? insertionRange = null;
        Range? visibleRange = null;
        Fields? fields = null;
        Field? referenceField = null;
        ParagraphFormat? paragraphFormat = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        WrapFormat? wrapFormat = null;
        LineFormat? line = null;
        FillFormat? fill = null;
        object? anchor = null;
        try
        {
            bookmarks = document.Bookmarks;
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            if (!bookmarks.Exists(anchorName))
                throw new InvalidOperationException(
                    "The numbered OMML Shape anchor bookmark is missing.");
            anchorBookmark = bookmarks[anchorName];
            anchorRange = anchorBookmark.Range;
            anchorParagraphs = anchorRange.Paragraphs;
            if (anchorParagraphs.Count != 1)
                throw new InvalidOperationException(
                    "The numbered OMML Shape anchor is not in exactly one paragraph.");
            anchorParagraph = anchorParagraphs[1];
            anchorParagraphRange = anchorParagraph.Range;
            StyleNativeDisplayAnchorParagraph(anchorParagraphRange);

            shapes = document.Shapes;
            anchor = anchorParagraphRange;
            shape = shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                0f,
                0f,
                NativeDisplayNumberShapeWidthPoints,
                ResolveNativeDisplayNumberShapeHeight(
                    formulaRange,
                    formulaHeightPoints,
                    formulaFontSizePoints),
                ref anchor);
            shape.Name = NativeDisplayNumberShapeName(formulaId);
            shape.AlternativeText = NativeDisplayNumberShapeAlternativeText(formulaId);

            textFrame = shape.TextFrame;
            textRange = textFrame.TextRange;
            textRange.Text = "()";
            insertionRange = textRange.Duplicate;
            insertionRange.SetRange(textRange.Start + 1, textRange.Start + 1);
            fields = textRange.Fields;
            referenceField = fields.Add(
                insertionRange,
                WdFieldType.wdFieldRef,
                NativeNumberBookmarkName(formulaId) + " \\h \\* CHARFORMAT",
                PreserveFormatting: true);
            if (!deferFieldUpdate)
                referenceField.Update();

            Release(textRange);
            textRange = textFrame.TextRange;
            visibleRange = textRange.Duplicate;
            var visibleText = visibleRange.Text ?? string.Empty;
            if (visibleText.EndsWith("\r", StringComparison.Ordinal)
                && visibleRange.End > visibleRange.Start)
                visibleRange.MoveEnd(WdUnits.wdCharacter, -1);
            bookmarks.Add(visibleBookmarkName, visibleRange);

            paragraphFormat = textRange.ParagraphFormat;
            paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            paragraphFormat.SpaceBefore = 0f;
            paragraphFormat.SpaceAfter = 0f;
            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            paragraphFormat.KeepTogether = -1;
            paragraphFormat.KeepWithNext = 0;
            paragraphFormat.WidowControl = 0;
            font = visibleRange.Font;
            ApplyEquationNumberFont(
                font,
                FormulaFontSize.Normalize(formulaFontSizePoints),
                position: 0);

            shape.RelativeHorizontalPosition =
                WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
            shape.Left = (float)WdShapePosition.wdShapeRight;
            shape.RelativeVerticalPosition =
                WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph;
            shape.Top = 0f;
            shape.LockAnchor = -1;
            wrapFormat = shape.WrapFormat;
            wrapFormat.Type = WdWrapType.wdWrapNone;
            wrapFormat.AllowOverlap = -1;
            wrapFormat.DistanceLeft = 0f;
            wrapFormat.DistanceRight = 0f;
            wrapFormat.DistanceTop = 0f;
            wrapFormat.DistanceBottom = 0f;
            line = shape.Line;
            line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            line.Transparency = 1f;
            fill = shape.Fill;
            fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            fill.Transparency = 1f;
            textFrame.MarginLeft = 0f;
            textFrame.MarginRight = 0f;
            textFrame.MarginTop = 0f;
            textFrame.MarginBottom = 0f;
            textFrame.VerticalAnchor =
                Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            ApplyNativeDisplayNumberTextLayout(
                textFrame,
                paragraphFormat,
                shape.Height,
                FormulaFontSize.Normalize(formulaFontSizePoints));
            // Geometry is finalized once, after SEQ/REF updates and all structural
            // Range mutations have completed. Measuring here caused the same Word
            // document to repaginate several times per insertion.
            return created;
        }
        finally
        {
            Release(fill);
            Release(line);
            Release(wrapFormat);
            Release(font);
            Release(paragraphFormat);
            Release(referenceField);
            Release(fields);
            Release(visibleRange);
            Release(insertionRange);
            Release(textRange);
            Release(textFrame);
            Release(shape);
            Release(shapes);
            Release(anchorParagraphRange);
            Release(anchorParagraph);
            Release(anchorParagraphs);
            Release(anchorRange);
            Release(anchorBookmark);
            Release(bookmarks);
        }
    }

    private static void StyleNativeDisplayNumberShape(
        Document document,
        Shape shape,
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId,
        bool updateField)
    {
        TextFrame? textFrame = null;
        Range? textRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? visibleBookmark = null;
        Range? visibleRange = null;
        Fields? fields = null;
        Field? field = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        ParagraphFormat? paragraphFormat = null;
        WrapFormat? wrapFormat = null;
        LineFormat? line = null;
        FillFormat? fill = null;
        try
        {
            shape.Name = NativeDisplayNumberShapeName(formulaId);
            shape.AlternativeText = NativeDisplayNumberShapeAlternativeText(formulaId);
            shape.RelativeHorizontalPosition =
                WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
            shape.Left = (float)WdShapePosition.wdShapeRight;
            shape.RelativeVerticalPosition =
                WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph;
            shape.Top = 0f;
            shape.Width = NativeDisplayNumberShapeWidthPoints;
            shape.Height = ResolveNativeDisplayNumberShapeHeight(
                formulaRange,
                formulaHeightPoints,
                formulaFontSizePoints);
            shape.LockAnchor = -1;
            wrapFormat = shape.WrapFormat;
            wrapFormat.Type = WdWrapType.wdWrapNone;
            wrapFormat.AllowOverlap = -1;
            wrapFormat.DistanceLeft = 0f;
            wrapFormat.DistanceRight = 0f;
            wrapFormat.DistanceTop = 0f;
            wrapFormat.DistanceBottom = 0f;
            line = shape.Line;
            line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            line.Transparency = 1f;
            fill = shape.Fill;
            fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            fill.Transparency = 1f;

            textFrame = shape.TextFrame;
            textFrame.MarginLeft = 0f;
            textFrame.MarginRight = 0f;
            textFrame.MarginTop = 0f;
            textFrame.MarginBottom = 0f;
            textFrame.VerticalAnchor =
                Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            textRange = textFrame.TextRange;
            paragraphFormat = textRange.ParagraphFormat;
            paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            paragraphFormat.SpaceBefore = 0f;
            paragraphFormat.SpaceAfter = 0f;
            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;

            bookmarks = document.Bookmarks;
            var visibleName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(visibleName))
                throw new InvalidOperationException(
                    "The numbered OMML Shape is missing its visible-number bookmark.");
            visibleBookmark = bookmarks[visibleName];
            visibleRange = visibleBookmark.Range;
            fields = visibleRange.Fields;
            if (fields.Count != 1)
                throw new InvalidOperationException(
                    "The numbered OMML Shape does not contain exactly one REF field.");
            field = fields[1];
            if (updateField) field.Update();
            font = visibleRange.Font;
            var normalizedFontSize = FormulaFontSize.Normalize(formulaFontSizePoints);
            ApplyEquationNumberFont(
                font,
                normalizedFontSize,
                position: 0);
            ApplyNativeDisplayNumberTextLayout(
                textFrame,
                paragraphFormat,
                shape.Height,
                normalizedFontSize);
            // ReconcileFormula performs the one post-field geometry/decorations
            // finalization after this style update.
        }
        finally
        {
            Release(fill);
            Release(line);
            Release(wrapFormat);
            Release(paragraphFormat);
            Release(font);
            Release(field);
            Release(fields);
            Release(visibleRange);
            Release(visibleBookmark);
            Release(bookmarks);
            Release(textRange);
            Release(textFrame);
        }
    }

    private static string ResolveNativeDisplayGeometryCacheKey(Range formulaRange)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? paragraphFormat = null;
        try
        {
            var fontSize = 0f;
            try
            {
                font = formulaRange.Font;
                fontSize = font.Size;
            }
            catch { }
            if (float.IsNaN(fontSize)
                || float.IsInfinity(fontSize)
                || fontSize <= 0f
                || fontSize > 256f)
                fontSize = 0f;

            var lineSpacingRule = 0;
            var lineSpacing = 0f;
            try
            {
                paragraphs = formulaRange.Paragraphs;
                if (paragraphs.Count > 0)
                {
                    paragraph = paragraphs[1];
                    paragraphFormat = paragraph.Format;
                    lineSpacingRule = (int)paragraphFormat.LineSpacingRule;
                    lineSpacing = paragraphFormat.LineSpacing;
                }
            }
            catch { }

            var xml = XDocument.Parse(
                formulaRange.WordOpenXML ?? string.Empty,
                LoadOptions.None);
            var math = (XNamespace)
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var structuralNames = xml
                .Descendants()
                .Where(element => element.Name.Namespace == math)
                .Select(element => element.Name.LocalName)
                .Where(name => name is not "r"
                    and not "t"
                    and not "rPr"
                    and not "ctrlPr"
                    and not "sty"
                    and not "nor")
                .ToArray();
            return string.Concat(
                fontSize.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture),
                "|",
                lineSpacingRule.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "|",
                lineSpacing.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture),
                "|",
                string.Join("/", structuralNames));
        }
        catch
        {
            return string.Empty;
        }
        finally
        {
            Release(paragraphFormat);
            Release(paragraph);
            Release(paragraphs);
            Release(font);
        }
    }

    private static bool AlignNativeDisplayNumberShapeToFormula(
        Document document,
        Shape shape,
        Range formulaRange,
        string formulaId)
    {
        Microsoft.Office.Interop.Word.Window? window = null;
        Microsoft.Office.Interop.Word.View? view = null;
        Bookmarks? bookmarks = null;
        Bookmark? visibleBookmark = null;
        Range? visibleRange = null;
        try
        {
            window = document.ActiveWindow;
            view = window.View;
            var zoomPercentage = Math.Max(1, view.Zoom.Percentage);
            var dpi = Math.Max(
                1u,
                GetDpiForWindow(new IntPtr(window.Hwnd)));
            bookmarks = document.Bookmarks;
            var visibleName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(visibleName)) return false;
            visibleBookmark = bookmarks[visibleName];
            visibleRange = visibleBookmark.Range;

            var geometryCacheKey = ResolveNativeDisplayGeometryCacheKey(formulaRange);
            if (!string.IsNullOrEmpty(geometryCacheKey))
            {
                lock (NativeDisplayGeometryCacheSync)
                {
                    if (NativeDisplayGeometryTopCache.TryGetValue(
                            geometryCacheKey,
                            out var cachedTop))
                    {
                        shape.Top = cachedTop;
                        if (string.Equals(
                                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                                "1",
                                StringComparison.Ordinal))
                            Console.WriteLine(
                                $"    true-display-number-shape geometry-cache formulaId={formulaId} top={cachedTop:0.###}pt");
                        return true;
                    }
                }
            }

            object scrollStart = true;
            window.ScrollIntoView(formulaRange, ref scrollStart);
            var finalDeltaPixels = 0.0;
            var formulaLeft = 0;
            var formulaTop = 0;
            var formulaWidth = 0;
            var formulaHeight = 0;
            var numberLeft = 0;
            var numberTop = 0;
            var numberWidth = 0;
            var numberHeight = 0;
            var appliedCorrectionPoints = 0f;
            // One measured correction is sufficient: Shape.Top is linear in Word's
            // page geometry. Re-measuring the same hidden Word window several times
            // per formula eventually destabilizes Office 2021 during 20+ rapid
            // numbered insertions, so verification belongs to the outer acceptance/
            // save boundary instead of an inner calibration loop.
            for (var iteration = 0; iteration < 1; iteration++)
            {
                document.Repaginate();
                System.Threading.Thread.Sleep(40);
                window.GetPoint(
                    out formulaLeft,
                    out formulaTop,
                    out formulaWidth,
                    out formulaHeight,
                    formulaRange);
                window.GetPoint(
                    out numberLeft,
                    out numberTop,
                    out numberWidth,
                    out numberHeight,
                    visibleRange);
                if (formulaHeight <= 0 || numberHeight <= 0) return false;

                var formulaCenterPixels = formulaTop + formulaHeight / 2.0;
                var numberCenterPixels = numberTop + numberHeight / 2.0;
                finalDeltaPixels = formulaCenterPixels - numberCenterPixels;
                if (Math.Abs(finalDeltaPixels) <= 1.0) break;
                var correctionPoints = (float)(
                    finalDeltaPixels
                    * 72.0
                    * 100.0
                    / dpi
                    / zoomPercentage);
                if (float.IsNaN(correctionPoints)
                    || float.IsInfinity(correctionPoints)
                    || Math.Abs(correctionPoints) > 120f)
                    return false;
                shape.Top = Math.Max(
                    -120f,
                    Math.Min(240f, shape.Top + correctionPoints));
                appliedCorrectionPoints += correctionPoints;
            }

            if (!string.IsNullOrEmpty(geometryCacheKey))
            {
                lock (NativeDisplayGeometryCacheSync)
                    NativeDisplayGeometryTopCache[geometryCacheKey] = shape.Top;
            }
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                Console.WriteLine(
                    $"    true-display-number-shape formulaId={formulaId} formulaPx={formulaLeft},{formulaTop},{formulaWidth},{formulaHeight} numberPx={numberLeft},{numberTop},{numberWidth},{numberHeight} centerDelta={finalDeltaPixels:0.###}px zoom={zoomPercentage}% dpi={dpi} correction={appliedCorrectionPoints:0.###}pt finalTop={shape.Top:0.###}pt");
            }
            return true;
        }
        catch (COMException error)
        {
            // A newly created or just-resized floating text box can temporarily
            // reject Window.GetPoint while Word commits its DrawingML. Numbering is
            // already structurally durable; the post-repagination finalizer retries
            // exact geometry without turning this transient layout state into an
            // insertion/edit failure.
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    true-display-number-shape geometry deferred formulaId={formulaId} error={error.Message}");
            return false;
        }
        finally
        {
            Release(visibleRange);
            Release(visibleBookmark);
            Release(bookmarks);
            Release(view);
            Release(window);
        }
    }

    private static void HideNativeDisplayNumberShapeDecoration(Shape shape)
    {
        Exception? firstError = null;
        LineFormat? line = null;
        try
        {
            line = shape.Line;
            line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            line.Transparency = 1f;
        }
        catch (Exception error)
        {
            // Word can expose Line as an E_FAIL proxy while Fill remains writable.
            // Never let the first decoration failure prevent the other independent
            // property from being persisted into DrawingML/VML.
            firstError = error;
        }
        finally { Release(line); }

        FillFormat? fill = null;
        try
        {
            fill = shape.Fill;
            fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            fill.Transparency = 1f;
        }
        catch (Exception error)
        {
            firstError ??= error;
        }
        finally { Release(fill); }

        if (firstError is not null)
            throw new InvalidOperationException(
                "Word deferred one or more numbered-OMML Shape decoration properties.",
                firstError);
    }

    private static void ApplyNativeDisplayNumberTextLayout(
        TextFrame textFrame,
        ParagraphFormat paragraphFormat,
        float shapeHeightPoints,
        float formulaFontSizePoints)
    {
        var normalizedFontSize = FormulaFontSize.Normalize(formulaFontSizePoints);
        var lineHeight = Math.Max(
            normalizedFontSize + 1f,
            Math.Min(
                Math.Max(normalizedFontSize + 1f, shapeHeightPoints),
                normalizedFontSize * 1.35f));
        textFrame.VerticalAnchor =
            Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
        paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
        paragraphFormat.SpaceBefore = 0f;
        paragraphFormat.SpaceAfter = 0f;
        paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
        paragraphFormat.LineSpacing = lineHeight;
        paragraphFormat.KeepTogether = -1;
        paragraphFormat.KeepWithNext = 0;
        paragraphFormat.WidowControl = 0;
    }

    private static void AlignNativeDisplayNumberShapeToRightBoundary(
        Shape shape,
        Range formulaRange)
    {
        TextFrame? textFrame = null;
        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        try
        {
            textFrame = shape.TextFrame;
            sections = formulaRange.Sections;
            if (sections.Count == 0) return;
            section = sections[1];
            pageSetup = section.PageSetup;
            var writableWidth = pageSetup.PageWidth
                - pageSetup.LeftMargin
                - pageSetup.RightMargin;
            var rightInset = Math.Max(0f, textFrame.MarginRight);
            shape.RelativeHorizontalPosition =
                WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
            shape.Left = Math.Max(
                0f,
                writableWidth - shape.Width + rightInset);
        }
        finally
        {
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(textFrame);
        }
    }

    private static float ResolveNativeDisplayNumberShapeHeight(
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints)
    {
        var fontSize = formulaFontSizePoints;
        if (float.IsNaN(fontSize)
            || float.IsInfinity(fontSize)
            || fontSize <= 0f
            || fontSize > 256f)
            fontSize = 11f;
        var height = NativeDisplayNumberShapeDefaultHeightPoints;
        if (!float.IsNaN(formulaHeightPoints)
            && !float.IsInfinity(formulaHeightPoints)
            && formulaHeightPoints > fontSize * 1.8f
            && formulaHeightPoints < 512f)
            height = Math.Max(height, formulaHeightPoints + 3f);

        try
        {
            var xml = XDocument.Parse(
                formulaRange.WordOpenXML ?? string.Empty,
                LoadOptions.None);
            var mathNamespace = (XNamespace)
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var maxFractionDepth = 0;
            foreach (var fraction in xml.Descendants(mathNamespace + "f"))
            {
                var depth = fraction.Ancestors(mathNamespace + "f").Count() + 1;
                maxFractionDepth = Math.Max(maxFractionDepth, depth);
            }
            if (maxFractionDepth > 0)
                height = Math.Max(
                    height,
                    fontSize * (2.35f + 1.25f * (maxFractionDepth - 1)) + 3f);

            var maximumMatrixRows = xml
                .Descendants(mathNamespace + "m")
                .Select(matrix => matrix.Elements(mathNamespace + "mr").Count())
                .DefaultIfEmpty(0)
                .Max();
            if (maximumMatrixRows > 0)
                height = Math.Max(
                    height,
                    maximumMatrixRows * fontSize * 1.35f + 7f);

            if (xml.Descendants(mathNamespace + "nary").Any())
                height = Math.Max(height, fontSize * 3.1f + 3f);
            if (xml.Descendants(mathNamespace + "rad").Any())
                height = Math.Max(height, fontSize * 2.35f + 3f);
        }
        catch
        {
            // Keep the conservative default when Word returns a story fragment
            // that cannot be parsed as standalone OpenXML.
        }
        return Math.Max(
            NativeDisplayNumberShapeDefaultHeightPoints,
            Math.Min(240f, height));
    }

    internal static bool HasStructurallyReusableNumberedNativeOmmlDisplayHost(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? anchorBookmark = null;
        Range? anchorBookmarkRange = null;
        Paragraphs? anchorParagraphs = null;
        Paragraph? anchorParagraph = null;
        Range? anchorParagraphRange = null;
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        OMaths? anchorMaths = null;
        Fields? anchorFields = null;
        InlineShapes? anchorInlineShapes = null;
        bool Fail(string reason)
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    true-display-structural-reuse formulaId={formulaId} result=false reason={reason}");
            return false;
        }
        try
        {
            if (!IsPureTrueDisplayFormulaParagraph(formulaRange))
                return Fail("formula-not-pure-display");

            bookmarks = document.Bookmarks;
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            var visibleName = EquationBookmarkName(formulaId);
            var captionName = NativeCaptionBookmarkName(formulaId);
            var numberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(anchorName)
                || !bookmarks.Exists(visibleName)
                || !bookmarks.Exists(captionName)
                || !bookmarks.Exists(numberName))
                return Fail(
                    $"bookmark-missing anchor={bookmarks.Exists(anchorName)} visible={bookmarks.Exists(visibleName)} caption={bookmarks.Exists(captionName)} number={bookmarks.Exists(numberName)}");

            anchorBookmark = bookmarks[anchorName];
            anchorBookmarkRange = anchorBookmark.Range;
            anchorParagraphs = anchorBookmarkRange.Paragraphs;
            formulaParagraphs = formulaRange.Paragraphs;
            if (anchorParagraphs.Count != 1 || formulaParagraphs.Count != 1)
                return Fail(
                    $"paragraph-count anchor={anchorParagraphs.Count} formula={formulaParagraphs.Count}");
            anchorParagraph = anchorParagraphs[1];
            formulaParagraph = formulaParagraphs[1];
            anchorParagraphRange = anchorParagraph.Range;
            formulaParagraphRange = formulaParagraph.Range;
            var anchorInTable = (bool)anchorParagraphRange.get_Information(
                WdInformation.wdWithInTable);
            if (anchorInTable
                || anchorParagraphRange.End != formulaParagraphRange.Start
                || anchorParagraphRange.Start >= formulaParagraphRange.Start)
                return Fail(
                    $"anchor-adjacency table={anchorInTable} anchor={anchorParagraphRange.Start}:{anchorParagraphRange.End} formula={formulaParagraphRange.Start}:{formulaParagraphRange.End}");

            anchorMaths = anchorParagraphRange.OMaths;
            anchorFields = anchorParagraphRange.Fields;
            anchorInlineShapes = anchorParagraphRange.InlineShapes;
            var anchorText = (anchorParagraphRange.Text ?? string.Empty)
                .Trim('\r', '\n', '\t', '\v', '\u0001', ' ');
            if (anchorText.Length != 0
                || anchorMaths.Count != 0
                || anchorFields.Count != 0
                || anchorInlineShapes.Count != 0)
                return Fail(
                    $"anchor-content text='{anchorText}' maths={anchorMaths.Count} fields={anchorFields.Count} inlineShapes={anchorInlineShapes.Count}");

            // Validate the actual serialized drawing/REF host without touching the
            // Shape RCW. Word can return E_FAIL from Shape.TextFrame/Line for the
            // duration of an OMath edit even though this exact OOXML remains valid.
            var anchorXml = anchorParagraphRange.WordOpenXML ?? string.Empty;
            var normalizedId = NormalizeFormulaIdForBookmark(formulaId);
            var shapeName = NativeDisplayNumberShapeName(formulaId);
            var shapeAlternativeText = NativeDisplayNumberShapeAlternativeText(formulaId);
            var hasShapeIdentity = anchorXml.IndexOf(
                    shapeName,
                    StringComparison.OrdinalIgnoreCase) >= 0
                || anchorXml.IndexOf(
                    shapeAlternativeText,
                    StringComparison.Ordinal) >= 0;
            var hasTextBox = anchorXml.IndexOf(
                "<w:txbxContent",
                StringComparison.OrdinalIgnoreCase) >= 0;
            var hasVisibleBookmark = anchorXml.IndexOf(
                EquationBookmarkPrefix + normalizedId,
                StringComparison.OrdinalIgnoreCase) >= 0;
            var hasReference = anchorXml.IndexOf(
                "REF " + NativeNumberBookmarkPrefix + normalizedId,
                StringComparison.OrdinalIgnoreCase) >= 0;
            var noMath = anchorXml.IndexOf(
                "<m:oMath",
                StringComparison.OrdinalIgnoreCase) < 0;
            if (!hasShapeIdentity || !hasTextBox || !hasVisibleBookmark
                || !hasReference || !noMath)
                return Fail(
                    $"anchor-openxml shape={hasShapeIdentity} txbx={hasTextBox} visible={hasVisibleBookmark} ref={hasReference} noMath={noMath} length={anchorXml.Length}");

            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    true-display-structural-reuse formulaId={formulaId} result=true");
            return true;
        }
        catch (Exception error)
        {
            return Fail($"exception={error.GetType().Name}:{error.Message}");
        }
        finally
        {
            Release(anchorInlineShapes);
            Release(anchorFields);
            Release(anchorMaths);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(anchorParagraphRange);
            Release(anchorParagraph);
            Release(anchorParagraphs);
            Release(anchorBookmarkRange);
            Release(anchorBookmark);
            Release(bookmarks);
        }
    }

    private static bool IsHealthyNumberedNativeOmmlTrueDisplayHost(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        OMaths? maths = null;
        OMath? math = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Shape? shape = null;
        try
        {
            maths = formulaRange.OMaths;
            if (maths.Count != 1) return false;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay) return false;
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable))
                return false;
            var xml = paragraphRange.WordOpenXML ?? string.Empty;
            if (xml.IndexOf("<m:oMathPara", StringComparison.OrdinalIgnoreCase) < 0
                || xml.IndexOf("<m:eqArr", StringComparison.OrdinalIgnoreCase) >= 0
                || xml.IndexOf("<w:fldChar", StringComparison.OrdinalIgnoreCase) >= 0
                || xml.IndexOf(" REF ", StringComparison.OrdinalIgnoreCase) >= 0)
                return false;
            shape = FindNativeDisplayNumberShape(document, formulaId);
            return shape is not null
                && IsNativeDisplayNumberShapeFieldHealthy(document, shape, formulaId)
                && NativeDisplayNumberShapeBelongsToFormula(
                    document,
                    formulaRange,
                    formulaId);
        }
        catch { return false; }
        finally
        {
            Release(shape);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(math);
            Release(maths);
        }
    }

    private static void AssertPureTrueDisplayFormulaParagraph(
        Range formulaRange,
        string formulaId)
    {
        if (!IsPureTrueDisplayFormulaParagraph(formulaRange))
            throw new InvalidOperationException(
                $"Numbered OMML {formulaId} did not materialize as one pure m:oMathPara paragraph.");
    }

    private static bool IsPureTrueDisplayFormulaParagraph(Range formulaRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            maths = formulaRange.OMaths;
            if (maths.Count != 1) return false;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay) return false;
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable))
                return false;
            var xml = paragraphRange.WordOpenXML ?? string.Empty;
            return xml.IndexOf("<m:oMathPara", StringComparison.OrdinalIgnoreCase) >= 0
                && xml.IndexOf("<m:eqArr", StringComparison.OrdinalIgnoreCase) < 0
                && xml.IndexOf("<w:fldChar", StringComparison.OrdinalIgnoreCase) < 0
                && xml.IndexOf(" REF ", StringComparison.OrdinalIgnoreCase) < 0;
        }
        catch { return false; }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(math);
            Release(maths);
        }
    }

    private static bool NativeDisplayNumberShapeBelongsToFormula(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Shape? shape = null;
        Range? anchor = null;
        Bookmarks? bookmarks = null;
        Bookmark? anchorBookmark = null;
        Range? anchorBookmarkRange = null;
        Paragraphs? anchorParagraphs = null;
        Paragraph? anchorParagraph = null;
        Range? anchorParagraphRange = null;
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        try
        {
            shape = FindNativeDisplayNumberShape(document, formulaId);
            if (shape is null) return false;
            anchor = shape.Anchor;
            bookmarks = document.Bookmarks;
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            if (!bookmarks.Exists(anchorName)) return false;
            anchorBookmark = bookmarks[anchorName];
            anchorBookmarkRange = anchorBookmark.Range;
            anchorParagraphs = anchorBookmarkRange.Paragraphs;
            formulaParagraphs = formulaRange.Paragraphs;
            if (anchorParagraphs.Count != 1 || formulaParagraphs.Count != 1)
                return false;
            anchorParagraph = anchorParagraphs[1];
            formulaParagraph = formulaParagraphs[1];
            anchorParagraphRange = anchorParagraph.Range;
            formulaParagraphRange = formulaParagraph.Range;
            if (anchor.Start < anchorParagraphRange.Start
                || anchor.Start > anchorParagraphRange.End)
                return false;
            return anchorParagraphRange.End == formulaParagraphRange.Start
                && anchorParagraphRange.Start < formulaParagraphRange.Start;
        }
        catch { return false; }
        finally
        {
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(anchorParagraphRange);
            Release(anchorParagraph);
            Release(anchorParagraphs);
            Release(anchorBookmarkRange);
            Release(anchorBookmark);
            Release(bookmarks);
            Release(anchor);
            Release(shape);
        }
    }

    private static bool IsNativeDisplayNumberShapeFieldHealthy(
        Document document,
        Shape shape,
        string formulaId)
    {
        TextFrame? textFrame = null;
        Range? textRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Bookmarks? bookmarks = null;
        Bookmark? visibleBookmark = null;
        Range? visibleRange = null;
        WrapFormat? wrapFormat = null;
        bool Fail(string reason)
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    true-display-number-shape health-fail formulaId={formulaId} reason={reason}");
            return false;
        }
        try
        {
            if (!string.Equals(
                    shape.AlternativeText ?? string.Empty,
                    NativeDisplayNumberShapeAlternativeText(formulaId),
                    StringComparison.Ordinal)
                && !string.Equals(
                    shape.Name ?? string.Empty,
                    NativeDisplayNumberShapeName(formulaId),
                    StringComparison.OrdinalIgnoreCase))
                return Fail("identity");
            if (shape.RelativeHorizontalPosition
                    != WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
                || shape.RelativeVerticalPosition
                    != WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph)
                return Fail($"relative-position horizontal={shape.RelativeHorizontalPosition} vertical={shape.RelativeVerticalPosition}");
            wrapFormat = shape.WrapFormat;
            if (wrapFormat.Type != WdWrapType.wdWrapNone)
                return Fail($"wrap={wrapFormat.Type}");

            textFrame = shape.TextFrame;
            textRange = textFrame.TextRange;
            if (textRange.StoryType != WdStoryType.wdTextFrameStory)
                return Fail($"story={textRange.StoryType}");
            fields = textRange.Fields;
            if (fields.Count != 1) return Fail($"field-count={fields.Count}");
            field = fields[1];
            code = field.Code;
            if (!IsReferenceToBookmark(
                    code.Text,
                    NativeNumberBookmarkName(formulaId)))
                return Fail($"ref-code='{code.Text}'");
            bookmarks = document.Bookmarks;
            var visibleName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(visibleName)) return Fail("visible-bookmark-missing");
            visibleBookmark = bookmarks[visibleName];
            visibleRange = visibleBookmark.Range;
            if (visibleRange.StoryType != WdStoryType.wdTextFrameStory)
                return Fail($"visible-story={visibleRange.StoryType}");
            if (visibleRange.Start < textRange.Start || visibleRange.End > textRange.End)
                return Fail($"visible-range={visibleRange.Start}:{visibleRange.End} text-range={textRange.Start}:{textRange.End}");
            return true;
        }
        catch (Exception error)
        {
            return Fail($"exception={error.GetType().Name}:{error.Message}");
        }
        finally
        {
            Release(wrapFormat);
            Release(visibleRange);
            Release(visibleBookmark);
            Release(bookmarks);
            Release(code);
            Release(field);
            Release(fields);
            Release(textRange);
            Release(textFrame);
        }
    }

    internal static bool IsSerializedNativeDisplayNumberShapeHealthy(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? anchorBookmark = null;
        Range? anchorBookmarkRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        bool Fail(string reason)
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    true-display-number-shape serialized-health-fail formulaId={formulaId} reason={reason}");
            return false;
        }
        try
        {
            bookmarks = document.Bookmarks;
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            if (!bookmarks.Exists(anchorName)) return Fail("anchor-bookmark-missing");
            anchorBookmark = bookmarks[anchorName];
            anchorBookmarkRange = anchorBookmark.Range;
            paragraphs = anchorBookmarkRange.Paragraphs;
            if (paragraphs.Count != 1) return Fail($"anchor-paragraph-count={paragraphs.Count}");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var xmlText = paragraphRange.WordOpenXML ?? string.Empty;
            if (string.IsNullOrWhiteSpace(xmlText)) return Fail("empty-openxml");
            var xml = XDocument.Parse(xmlText, LoadOptions.PreserveWhitespace);

            var word = (XNamespace)
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var drawing = (XNamespace)
                "http://schemas.openxmlformats.org/drawingml/2006/main";
            var wordDrawing = (XNamespace)
                "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            var wordShape = (XNamespace)
                "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
            var vml = (XNamespace)"urn:schemas-microsoft-com:vml";
            var math = (XNamespace)
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var normalizedId = NormalizeFormulaIdForBookmark(formulaId);
            var expectedName = NativeDisplayNumberShapeName(formulaId);
            var expectedDescription =
                NativeDisplayNumberShapeAlternativeText(formulaId);
            var expectedVisibleBookmark = EquationBookmarkPrefix + normalizedId;
            var expectedReference = NativeNumberBookmarkPrefix + normalizedId;

            var docProperties = xml
                .Descendants(wordDrawing + "docPr")
                .FirstOrDefault(element =>
                    string.Equals(
                        (string?)element.Attribute("name"),
                        expectedName,
                        StringComparison.OrdinalIgnoreCase)
                    || string.Equals(
                        (string?)element.Attribute("descr"),
                        expectedDescription,
                        StringComparison.Ordinal));
            var drawingAnchor = docProperties?
                .Ancestors(wordDrawing + "anchor")
                .FirstOrDefault();
            if (drawingAnchor is null) return Fail("drawing-anchor-missing");
            var positionHorizontal = drawingAnchor.Element(wordDrawing + "positionH");
            var positionVertical = drawingAnchor.Element(wordDrawing + "positionV");
            var horizontalOffsetText =
                positionHorizontal?.Element(wordDrawing + "posOffset")?.Value;
            if (!long.TryParse(
                    horizontalOffsetText,
                    System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var horizontalOffsetEmu)
                || horizontalOffsetEmu < 0
                || !string.Equals(
                    (string?)positionHorizontal?.Attribute("relativeFrom"),
                    "margin",
                    StringComparison.OrdinalIgnoreCase)
                || !string.Equals(
                    (string?)positionVertical?.Attribute("relativeFrom"),
                    "paragraph",
                    StringComparison.OrdinalIgnoreCase)
                || drawingAnchor.Element(wordDrawing + "wrapNone") is null)
                return Fail("drawing-position-or-wrap");

            var shape = drawingAnchor.Descendants(wordShape + "wsp").FirstOrDefault();
            var shapeProperties = shape?.Element(wordShape + "spPr");
            var line = shapeProperties?.Element(drawing + "ln");
            var drawingNoFill = shapeProperties?.Element(drawing + "noFill") is not null;
            var drawingNoStroke = line?.Element(drawing + "noFill") is not null;
            var textBox = shape?.Descendants(word + "txbxContent").FirstOrDefault();
            if (!drawingNoFill || !drawingNoStroke || textBox is null)
                return Fail(
                    $"drawing-decoration-or-textbox fill={drawingNoFill} stroke={drawingNoStroke} txbx={textBox is not null}");
            var hasVisibleBookmark = textBox
                .Descendants(word + "bookmarkStart")
                .Any(element => string.Equals(
                    (string?)element.Attribute(word + "name"),
                    expectedVisibleBookmark,
                    StringComparison.OrdinalIgnoreCase));
            var hasReference = textBox
                .Descendants(word + "instrText")
                .Any(element => element.Value.IndexOf(
                    "REF " + expectedReference,
                    StringComparison.OrdinalIgnoreCase) >= 0);
            if (!hasVisibleBookmark || !hasReference)
                return Fail(
                    $"drawing-visible-ref visible={hasVisibleBookmark} ref={hasReference}");

            var vmlShape = xml
                .Descendants(vml + "shape")
                .FirstOrDefault(element =>
                    string.Equals(
                        (string?)element.Attribute("id"),
                        expectedName,
                        StringComparison.OrdinalIgnoreCase)
                    || string.Equals(
                        (string?)element.Attribute("alt"),
                        expectedDescription,
                        StringComparison.Ordinal));
            if (vmlShape is null) return Fail("vml-fallback-missing");
            var vmlNoFill = string.Equals(
                    (string?)vmlShape.Attribute("filled"),
                    "f",
                    StringComparison.OrdinalIgnoreCase)
                || vmlShape.Descendants(vml + "fill").Any(element =>
                    string.Equals(
                        (string?)element.Attribute("opacity"),
                        "0",
                        StringComparison.OrdinalIgnoreCase));
            var vmlNoStroke = string.Equals(
                    (string?)vmlShape.Attribute("stroked"),
                    "f",
                    StringComparison.OrdinalIgnoreCase)
                || vmlShape.Descendants(vml + "stroke").Any(element =>
                    string.Equals(
                        (string?)element.Attribute("opacity"),
                        "0",
                        StringComparison.OrdinalIgnoreCase));
            if (!vmlNoFill || !vmlNoStroke)
                return Fail(
                    $"vml-decoration fill={vmlNoFill} stroke={vmlNoStroke}");
            if (xml.Descendants(math + "oMath").Any()
                || xml.Descendants(word + "tbl").Any())
                return Fail("math-or-table-in-anchor");
            return true;
        }
        catch (Exception error)
        {
            return Fail($"exception={error.GetType().Name}:{error.Message}");
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(anchorBookmarkRange);
            Release(anchorBookmark);
            Release(bookmarks);
        }
    }

    internal static bool IsNativeDisplayNumberShapeGeometryCommitted(
        Document document,
        string formulaId)
    {
        Shape? shape = null;
        Bookmark? formulaBookmark = null;
        Range? formulaRange = null;
        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        TextFrame? textFrame = null;
        try
        {
            shape = FindNativeDisplayNumberShape(document, formulaId);
            if (shape is null) return false;
            var left = shape.Left;
            var top = shape.Top;
            var width = shape.Width;
            if (float.IsNaN(left)
                || float.IsInfinity(left)
                || left < 0f
                || float.IsNaN(top)
                || float.IsInfinity(top)
                || top < -120f
                || top > 240f
                || float.IsNaN(width)
                || float.IsInfinity(width)
                || width <= 0f)
                return false;

            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(
                document,
                formulaId);
            if (formulaBookmark is null) return false;
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            sections = formulaRange.Sections;
            if (sections.Count == 0) return false;
            section = sections[1];
            pageSetup = section.PageSetup;
            textFrame = shape.TextFrame;
            var writableWidth = pageSetup.PageWidth
                - pageSetup.LeftMargin
                - pageSetup.RightMargin;
            var expectedLeft = Math.Max(
                0f,
                writableWidth - width + Math.Max(0f, textFrame.MarginRight));
            return Math.Abs(left - expectedLeft) <= 1.5f;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(textFrame);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(formulaRange);
            Release(formulaBookmark);
            Release(shape);
        }
    }

    internal static bool TryNormalizeSerializedNativeDisplayNumberShape(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? anchorBookmark = null;
        Range? anchorBookmarkRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        try
        {
            bookmarks = document.Bookmarks;
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            if (!bookmarks.Exists(anchorName)) return false;
            anchorBookmark = bookmarks[anchorName];
            anchorBookmarkRange = anchorBookmark.Range;
            paragraphs = anchorBookmarkRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            var xmlText = paragraphRange.WordOpenXML ?? string.Empty;
            if (string.IsNullOrWhiteSpace(xmlText)) return false;
            var xml = XDocument.Parse(xmlText, LoadOptions.PreserveWhitespace);

            var drawing = (XNamespace)
                "http://schemas.openxmlformats.org/drawingml/2006/main";
            var wordDrawing = (XNamespace)
                "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            var wordShape = (XNamespace)
                "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
            var vml = (XNamespace)"urn:schemas-microsoft-com:vml";
            var expectedName = NativeDisplayNumberShapeName(formulaId);
            var expectedDescription =
                NativeDisplayNumberShapeAlternativeText(formulaId);
            var docProperties = xml
                .Descendants(wordDrawing + "docPr")
                .FirstOrDefault(element =>
                    string.Equals(
                        (string?)element.Attribute("name"),
                        expectedName,
                        StringComparison.OrdinalIgnoreCase)
                    || string.Equals(
                        (string?)element.Attribute("descr"),
                        expectedDescription,
                        StringComparison.Ordinal));
            var drawingAnchor = docProperties?
                .Ancestors(wordDrawing + "anchor")
                .FirstOrDefault();
            var shapeProperties = drawingAnchor?
                .Descendants(wordShape + "spPr")
                .FirstOrDefault();
            var vmlShape = xml
                .Descendants(vml + "shape")
                .FirstOrDefault(element =>
                    string.Equals(
                        (string?)element.Attribute("id"),
                        expectedName,
                        StringComparison.OrdinalIgnoreCase)
                    || string.Equals(
                        (string?)element.Attribute("alt"),
                        expectedDescription,
                        StringComparison.Ordinal));
            if (drawingAnchor is null || shapeProperties is null || vmlShape is null)
                return false;

            // Even when Word exposes the newly created Shape only as an E_FAIL
            // proxy, its serialized DrawingML remains editable. Normalize the
            // right-margin location to an explicit offset so the package never
            // persists Word's wdShapeRight/alignment sentinel as a half-finished
            // numbered-OMML host.
            sections = paragraphRange.Sections;
            if (sections.Count == 0) return false;
            section = sections[1];
            pageSetup = section.PageSetup;
            var writableWidthPoints = pageSetup.PageWidth
                - pageSetup.LeftMargin
                - pageSetup.RightMargin;
            var explicitLeftPoints = Math.Max(
                0f,
                writableWidthPoints - NativeDisplayNumberShapeWidthPoints);
            var explicitLeftEmu = (long)Math.Round(
                explicitLeftPoints * 12700.0,
                MidpointRounding.AwayFromZero);
            var positionHorizontal = drawingAnchor.Element(
                wordDrawing + "positionH");
            if (positionHorizontal is null) return false;
            positionHorizontal.SetAttributeValue("relativeFrom", "margin");
            positionHorizontal.RemoveNodes();
            positionHorizontal.Add(
                new XElement(
                    wordDrawing + "posOffset",
                    explicitLeftEmu.ToString(
                        System.Globalization.CultureInfo.InvariantCulture)));

            static string UpsertVmlStyle(
                string? style,
                string property,
                string value)
            {
                var parts = (style ?? string.Empty)
                    .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(part => part.Trim())
                    .Where(part => part.Length > 0)
                    .ToList();
                var prefix = property + ":";
                var replaced = false;
                for (var index = 0; index < parts.Count; index++)
                {
                    if (!parts[index].StartsWith(
                            prefix,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    parts[index] = prefix + value;
                    replaced = true;
                    break;
                }
                if (!replaced) parts.Add(prefix + value);
                return string.Join(";", parts) + ";";
            }
            var vmlStyle = (string?)vmlShape.Attribute("style");
            vmlStyle = UpsertVmlStyle(
                vmlStyle,
                "margin-left",
                explicitLeftPoints.ToString(
                    "0.###",
                    System.Globalization.CultureInfo.InvariantCulture) + "pt");
            vmlStyle = UpsertVmlStyle(
                vmlStyle,
                "mso-position-horizontal",
                "absolute");
            vmlStyle = UpsertVmlStyle(
                vmlStyle,
                "mso-position-horizontal-relative",
                "margin");
            vmlStyle = UpsertVmlStyle(
                vmlStyle,
                "v-text-anchor",
                "middle");
            vmlShape.SetAttributeValue("style", vmlStyle);

            // Word's fresh text boxes default to 7.2/3.6pt insets. Those insets
            // shift a migrated REF label left even when the Shape itself is pinned
            // to the writable right boundary. Normalize both the DrawingML Choice
            // and VML Fallback to the same zero-inset representation used by a
            // fully committed healthy numbered-OMML Shape.
            var bodyProperties = drawingAnchor
                .Descendants(wordShape + "bodyPr")
                .FirstOrDefault();
            if (bodyProperties is not null)
            {
                bodyProperties.SetAttributeValue("lIns", "0");
                bodyProperties.SetAttributeValue("tIns", "0");
                bodyProperties.SetAttributeValue("rIns", "0");
                bodyProperties.SetAttributeValue("bIns", "0");
                bodyProperties.SetAttributeValue("anchor", "ctr");
                bodyProperties.SetAttributeValue("anchorCtr", "0");
            }
            var vmlTextBox = vmlShape
                .Descendants(vml + "textbox")
                .FirstOrDefault();
            if (vmlTextBox is not null)
                vmlTextBox.SetAttributeValue("inset", "0,0,0,0");

            static bool IsDrawingFillElement(XElement element) =>
                element.Name.LocalName is "noFill" or "solidFill" or "gradFill"
                    or "blipFill" or "pattFill" or "grpFill";
            foreach (var fillElement in shapeProperties
                         .Elements()
                         .Where(IsDrawingFillElement)
                         .ToArray())
                fillElement.Remove();
            var geometry = shapeProperties.Elements().FirstOrDefault(element =>
                element.Name.LocalName is "prstGeom" or "custGeom");
            var noFill = new XElement(drawing + "noFill");
            if (geometry is not null) geometry.AddAfterSelf(noFill);
            else shapeProperties.AddFirst(noFill);

            var line = shapeProperties.Element(drawing + "ln");
            if (line is null)
            {
                line = new XElement(
                    drawing + "ln",
                    new XAttribute("w", "6350"),
                    new XElement(drawing + "noFill"));
                shapeProperties.Add(line);
            }
            else
            {
                foreach (var fillElement in line
                             .Elements()
                             .Where(IsDrawingFillElement)
                             .ToArray())
                    fillElement.Remove();
                line.AddFirst(new XElement(drawing + "noFill"));
            }

            vmlShape.SetAttributeValue("filled", "f");
            vmlShape.SetAttributeValue("stroked", "f");
            var vmlFill = vmlShape.Element(vml + "fill");
            if (vmlFill is null)
            {
                vmlFill = new XElement(vml + "fill");
                vmlShape.Add(vmlFill);
            }
            vmlFill.SetAttributeValue("opacity", "0");
            var vmlStroke = vmlShape.Element(vml + "stroke");
            if (vmlStroke is null)
            {
                vmlStroke = new XElement(vml + "stroke");
                vmlShape.Add(vmlStroke);
            }
            vmlStroke.SetAttributeValue("opacity", "0");

            paragraphRange.InsertXML(
                xml.ToString(SaveOptions.DisableFormatting));
            try { document.Repaginate(); } catch { }
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    true-display-number-shape serialized-decoration-normalized formulaId={formulaId}");
            return true;
        }
        catch (Exception error)
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    true-display-number-shape serialized-decoration-normalization-failed formulaId={formulaId} error={error.GetType().Name}:{error.Message}");
            return false;
        }
        finally
        {
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(anchorBookmarkRange);
            Release(anchorBookmark);
            Release(bookmarks);
        }
    }

    internal static bool TryHideNativeDisplayNumberShapeDecoration(
        Document document,
        string formulaId)
    {
        Shape? shape = null;
        try
        {
            shape = FindNativeDisplayNumberShape(document, formulaId);
            if (shape is null) return false;
            HideNativeDisplayNumberShapeDecoration(shape);
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    true-display-number-shape decoration-finalized formulaId={formulaId}");
            return true;
        }
        catch (Exception error)
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    true-display-number-shape decoration-deferred formulaId={formulaId} error={error.GetType().Name}:{error.Message}");
            return false;
        }
        finally { Release(shape); }
    }

    internal static void TryFinalizeNativeDisplayNumberShapeLayout(
        Document document,
        string formulaId)
    {
        Bookmark? formulaBookmark = null;
        Range? formulaRange = null;
        Shape? shape = null;
        var traceAcceptance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
            "1",
            StringComparison.Ordinal);
        void TraceFinalize(string stage)
        {
            if (traceAcceptance)
                Console.WriteLine(
                    $"    true-display-number-shape finalize formulaId={formulaId} stage={stage}");
        }
        try
        {
            TraceFinalize("enter");
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(
                document,
                formulaId);
            if (formulaBookmark is null)
            {
                TraceFinalize("no-formula-bookmark");
                return;
            }
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            if (!IsPureTrueDisplayFormulaParagraph(formulaRange))
            {
                TraceFinalize("formula-not-pure-display");
                return;
            }
            TraceFinalize("formula-ready");

            // Geometry is independent from Line/Fill persistence. Measure it once
            // per outer finalization call; decoration retries must never repeat
            // Window.GetPoint because rapid hidden-Word runs can terminate Office
            // after dozens of screen-geometry probes.
            shape = FindNativeDisplayNumberShape(document, formulaId);
            if (shape is null)
            {
                TraceFinalize("no-shape");
                return;
            }
            try
            {
                AlignNativeDisplayNumberShapeToRightBoundary(
                    shape,
                    formulaRange);
            }
            catch (Exception horizontalError)
            {
                if (traceAcceptance)
                {
                    Console.WriteLine(
                        $"    true-display-number-shape finalize formulaId={formulaId} stage=horizontal-deferred error={horizontalError.GetType().Name}:{horizontalError.Message}");
                }
            }
            _ = AlignNativeDisplayNumberShapeToFormula(
                document,
                shape,
                formulaRange,
                formulaId);
            Release(shape);
            shape = null;

            Exception? lastShapeError = null;
            for (var attempt = 0; attempt < 3; attempt++)
            {
                if (attempt > 0)
                {
                    try { document.Repaginate(); } catch { }
                    System.Threading.Thread.Sleep(40 * (attempt + 1));
                }
                Release(shape);
                shape = FindNativeDisplayNumberShape(document, formulaId);
                if (shape is null)
                {
                    TraceFinalize("no-shape");
                    return;
                }
                try
                {
                    // Line/Fill are independent from geometry and can be retried
                    // safely without another screen-layout query.
                    HideNativeDisplayNumberShapeDecoration(shape);
                    TraceFinalize($"aligned-attempt-{attempt + 1}");
                    return;
                }
                catch (Exception error)
                {
                    lastShapeError = error;
                    if (traceAcceptance)
                    {
                        var inner = error.InnerException;
                        Console.WriteLine(
                            $"    true-display-number-shape finalize formulaId={formulaId} stage=attempt-{attempt + 1}-failed error={error.GetType().Name}:{error.Message} inner={(inner is null ? "none" : inner.GetType().Name + ":" + inner.Message)}");
                    }
                }
            }
            if (lastShapeError is not null)
            {
                // Shape/Line/TextFrame COM access can be transiently unavailable
                // while Word commits the surrounding professional OMath. The
                // serialized anchor, external REF and previously applied geometry
                // remain valid. Never delete/recreate the Shape here: doing so is a
                // destructive response to a cosmetic retry failure and restores
                // Word's default visible border after legacy-document migration.
                TraceFinalize("deferred");
            }
        }
        catch
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                throw;
        }
        finally
        {
            Release(shape);
            Release(formulaRange);
            Release(formulaBookmark);
        }
    }

    private static Range RecreateNativeDisplayShapeAnchorParagraph(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? formulaBookmark = null;
        Bookmark? anchorBookmark = null;
        Range? formulaRange = null;
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Range? anchorBookmarkRange = null;
        Paragraphs? anchorParagraphs = null;
        Paragraph? anchorParagraph = null;
        Range? anchorParagraphRange = null;
        OMaths? anchorMaths = null;
        Fields? anchorFields = null;
        InlineShapes? anchorInlineShapes = null;
        Range? anchorProbe = null;
        Range? replacementAnchorRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var formulaName = WordOmmlFormulaStore.BookmarkName(formulaId);
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            if (!bookmarks.Exists(formulaName) || !bookmarks.Exists(anchorName))
                throw new InvalidOperationException(
                    $"The numbered-OMML formula or Shape anchor bookmark for {formulaId} is missing.");

            formulaBookmark = bookmarks[formulaName];
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            formulaParagraphs = formulaRange.Paragraphs;
            if (formulaParagraphs.Count != 1)
                throw new InvalidOperationException(
                    $"The numbered-OMML formula {formulaId} does not occupy one paragraph while repairing its Shape anchor.");
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range;

            anchorBookmark = bookmarks[anchorName];
            anchorBookmarkRange = anchorBookmark.Range;
            anchorParagraphs = anchorBookmarkRange.Paragraphs;
            if (anchorParagraphs.Count != 1)
                throw new InvalidOperationException(
                    $"The numbered-OMML Shape anchor for {formulaId} does not occupy one paragraph.");
            anchorParagraph = anchorParagraphs[1];
            anchorParagraphRange = anchorParagraph.Range.Duplicate;
            if (anchorParagraphRange.End != formulaParagraphRange.Start)
                throw new InvalidOperationException(
                    $"The numbered-OMML Shape anchor for {formulaId} is not immediately before its formula.");
            anchorMaths = anchorParagraphRange.OMaths;
            anchorFields = anchorParagraphRange.Fields;
            anchorInlineShapes = anchorParagraphRange.InlineShapes;
            var anchorText = (anchorParagraphRange.Text ?? string.Empty)
                .Trim('\r', '\n', '\t', '\v', '\u0001', ' ');
            if (anchorText.Length != 0
                || anchorMaths.Count != 0
                || anchorFields.Count != 0
                || anchorInlineShapes.Count != 0)
                throw new InvalidOperationException(
                    $"The numbered-OMML Shape anchor paragraph for {formulaId} contains user content.");

            DeleteBookmarkOnly(document, EquationBookmarkName(formulaId));
            anchorBookmark.Delete();
            anchorBookmark = null;
            anchorParagraphRange.Delete();
            document.Repaginate();

            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(formulaRange);
            Release(formulaBookmark);
            Release(bookmarks);
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(formulaName))
                throw new InvalidOperationException(
                    $"Word lost numbered-OMML formula {formulaId} while deleting its unusable Shape anchor.");
            formulaBookmark = bookmarks[formulaName];
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            formulaParagraphs = formulaRange.Paragraphs;
            if (formulaParagraphs.Count != 1)
                throw new InvalidOperationException(
                    $"The numbered-OMML formula {formulaId} no longer occupies one paragraph after Shape-anchor deletion.");
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range.Duplicate;
            formulaParagraphRange.InsertParagraphBefore();

            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(formulaRange);
            Release(formulaBookmark);
            Release(bookmarks);
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(formulaName))
                throw new InvalidOperationException(
                    $"Word lost numbered-OMML formula {formulaId} while recreating its Shape anchor.");
            formulaBookmark = bookmarks[formulaName];
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            formulaParagraphs = formulaRange.Paragraphs;
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range;
            if (formulaParagraphRange.Start <= 0)
                throw new InvalidOperationException(
                    $"Word did not create a paragraph before numbered-OMML formula {formulaId}.");
            anchorProbe = document.Range(
                formulaParagraphRange.Start - 1,
                formulaParagraphRange.Start);
            anchorParagraphs = anchorProbe.Paragraphs;
            if (anchorParagraphs.Count != 1)
                throw new InvalidOperationException(
                    $"The recreated Shape anchor for numbered-OMML formula {formulaId} is invalid.");
            anchorParagraph = anchorParagraphs[1];
            Release(anchorParagraphRange);
            anchorParagraphRange = anchorParagraph.Range;
            StyleNativeDisplayAnchorParagraph(anchorParagraphRange);
            replacementAnchorRange = document.Range(
                anchorParagraphRange.Start,
                anchorParagraphRange.Start);
            bookmarks.Add(anchorName, replacementAnchorRange);

            var result = formulaRange;
            formulaRange = null;
            return result;
        }
        finally
        {
            Release(replacementAnchorRange);
            Release(anchorProbe);
            Release(anchorInlineShapes);
            Release(anchorFields);
            Release(anchorMaths);
            Release(anchorParagraphRange);
            Release(anchorParagraph);
            Release(anchorParagraphs);
            Release(anchorBookmarkRange);
            Release(anchorBookmark);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(formulaRange);
            Release(formulaBookmark);
            Release(bookmarks);
        }
    }

    private static Shape? FindNativeDisplayNumberShape(
        Document document,
        string formulaId)
    {
        Shapes? shapes = null;
        Shape? candidate = null;
        Shape? match = null;
        try
        {
            shapes = document.Shapes;
            var expectedName = NativeDisplayNumberShapeName(formulaId);
            var expectedAlternativeText = NativeDisplayNumberShapeAlternativeText(formulaId);

            // Word guarantees Shape.Name uniqueness. Current VisualTeX hosts carry
            // a deterministic name, so use the collection's native name indexer
            // instead of enumerating every Shape for every insert/edit/layout pass.
            // The latter made the Nth numbered OMML operation linear in N and was
            // responsible for multi-second latency by only twenty formulas.
            object directIndex = expectedName;
            try
            {
                candidate = shapes[ref directIndex];
                if (candidate is not null)
                {
                    var directMatch = string.Equals(
                            candidate.Name ?? string.Empty,
                            expectedName,
                            StringComparison.OrdinalIgnoreCase)
                        || string.Equals(
                            candidate.AlternativeText ?? string.Empty,
                            expectedAlternativeText,
                            StringComparison.Ordinal);
                    if (directMatch)
                    {
                        match = candidate;
                        candidate = null;
                        var directResult = match;
                        match = null;
                        return directResult;
                    }
                }
            }
            catch (Exception error) when (error is COMException or ArgumentException)
            {
                // A missing name is surfaced as ArgumentException on some Word
                // builds and COMException on others. Legacy/copied Shapes can also
                // have an Office-generated name; fall back to deterministic
                // AlternativeText identity below.
            }
            finally
            {
                Release(candidate);
                candidate = null;
            }

            // A freshly materialized display formula has its anchor paragraph but
            // no visible VTEq bookmark yet, so there cannot be a legacy Shape to
            // recover. Avoid an O(N) enumeration in that dominant insertion path.
            if (!HasVisibleEquationNumberBookmark(document, formulaId))
                return null;

            for (var index = 1; index <= shapes.Count; index++)
            {
                candidate = shapes[index];
                bool isMatch;
                try
                {
                    isMatch = string.Equals(
                            candidate.Name ?? string.Empty,
                            expectedName,
                            StringComparison.OrdinalIgnoreCase)
                        || string.Equals(
                            candidate.AlternativeText ?? string.Empty,
                            expectedAlternativeText,
                            StringComparison.Ordinal);
                }
                catch { isMatch = false; }
                if (!isMatch)
                {
                    Release(candidate);
                    candidate = null;
                    continue;
                }
                if (match is not null)
                    throw new InvalidOperationException(
                        $"More than one numbered OMML Shape is associated with formula {formulaId}.");
                match = candidate;
                candidate = null;
            }
            var result = match;
            match = null;
            return result;
        }
        finally
        {
            Release(match);
            Release(candidate);
            Release(shapes);
        }
    }

    internal static bool TryCaptureNativeDisplayAnchorParagraphBounds(
        Document document,
        string formulaId,
        out int paragraphStart,
        out int paragraphEnd)
    {
        paragraphStart = -1;
        paragraphEnd = -1;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        Shape? shape = null;
        Range? shapeAnchor = null;
        Fields? fields = null;
        OMaths? maths = null;
        InlineShapes? inlineShapes = null;
        try
        {
            bookmarks = document.Bookmarks;
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            if (!bookmarks.Exists(anchorName)) return false;
            bookmark = bookmarks[anchorName];
            bookmarkRange = bookmark.Range;
            paragraphs = bookmarkRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable))
                return false;
            var text = (paragraphRange.Text ?? string.Empty)
                .Trim('\r', '\n', '\t', '\v', ' ');
            if (text.Length != 0) return false;
            fields = paragraphRange.Fields;
            maths = paragraphRange.OMaths;
            inlineShapes = paragraphRange.InlineShapes;
            if (fields.Count != 0 || maths.Count != 0 || inlineShapes.Count != 0)
                return false;
            format = paragraphRange.ParagraphFormat;
            if (format.LineSpacingRule != WdLineSpacing.wdLineSpaceExactly
                || Math.Abs(format.LineSpacing - NativeDisplayAnchorLineSpacingPoints) > 0.1f)
                return false;
            font = paragraphRange.Font;
            var fontSize = font.Size;
            if (fontSize > 0f
                && fontSize < 999999f
                && Math.Abs(fontSize - NativeDisplayAnchorLineSpacingPoints) > 0.1f)
                return false;
            shape = FindNativeDisplayNumberShape(document, formulaId);
            if (shape is null) return false;
            shapeAnchor = shape.Anchor;
            if (shapeAnchor.Start < paragraphRange.Start
                || shapeAnchor.Start > paragraphRange.End)
                return false;
            paragraphStart = paragraphRange.Start;
            paragraphEnd = paragraphRange.End;
            return true;
        }
        catch
        {
            paragraphStart = -1;
            paragraphEnd = -1;
            return false;
        }
        finally
        {
            Release(inlineShapes);
            Release(maths);
            Release(fields);
            Release(shapeAnchor);
            Release(shape);
            Release(font);
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static bool RemoveCapturedNativeDisplayAnchorParagraphBeforeOle(
        Document document,
        int capturedParagraphStart,
        Range replacementOleRange)
    {
        if (capturedParagraphStart < 0) return false;
        Range? content = null;
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Paragraphs? replacementParagraphs = null;
        Paragraph? replacementParagraph = null;
        Range? replacementParagraphRange = null;
        ParagraphFormat? format = null;
        Fields? fields = null;
        OMaths? maths = null;
        InlineShapes? inlineShapes = null;
        Bookmarks? bookmarks = null;
        Frames? frames = null;
        ShapeRange? anchoredShapes = null;
        try
        {
            content = document.Content;
            if (capturedParagraphStart < content.Start
                || capturedParagraphStart >= content.End)
                return false;
            replacementParagraphs = replacementOleRange.Paragraphs;
            if (replacementParagraphs.Count != 1) return false;
            replacementParagraph = replacementParagraphs[1];
            replacementParagraphRange = replacementParagraph.Range;
            if (replacementParagraphRange.Start <= capturedParagraphStart)
                return false;
            probe = document.Range(
                capturedParagraphStart,
                Math.Min(content.End, capturedParagraphStart + 1));
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.Start != capturedParagraphStart
                || paragraphRange.End != replacementParagraphRange.Start
                || (bool)paragraphRange.get_Information(WdInformation.wdWithInTable))
                return false;
            var text = (paragraphRange.Text ?? string.Empty)
                .Trim('\r', '\n', '\t', '\v', '\u0001', ' ');
            if (text.Length != 0) return false;
            fields = paragraphRange.Fields;
            maths = paragraphRange.OMaths;
            inlineShapes = paragraphRange.InlineShapes;
            bookmarks = paragraphRange.Bookmarks;
            frames = paragraphRange.Frames;
            if (fields.Count != 0
                || maths.Count != 0
                || inlineShapes.Count != 0
                || bookmarks.Count != 0
                || frames.Count != 0)
                return false;
            try
            {
                anchoredShapes = paragraphRange.ShapeRange;
                if (anchoredShapes.Count != 0) return false;
            }
            catch (COMException)
            {
                // Word throws when the paragraph has no anchored Shape.
            }
            format = paragraphRange.ParagraphFormat;
            if (format.LineSpacingRule != WdLineSpacing.wdLineSpaceExactly
                || Math.Abs(format.LineSpacing - NativeDisplayAnchorLineSpacingPoints) > 0.1f
                || Math.Abs(format.SpaceBefore) > 0.1f
                || Math.Abs(format.SpaceAfter) > 0.1f)
                return false;

            var deleteStart = paragraphRange.Start;
            var deleteEnd = paragraphRange.End;
            var paragraphCountBefore = document.Paragraphs.Count;

            // Release every RCW that still points into the obsolete anchor before
            // deleting it. Word 2021 can otherwise report success while retaining
            // the empty one-point paragraph after its floating Shape was removed.
            Release(anchoredShapes); anchoredShapes = null;
            Release(frames); frames = null;
            Release(bookmarks); bookmarks = null;
            Release(inlineShapes); inlineShapes = null;
            Release(maths); maths = null;
            Release(fields); fields = null;
            Release(format); format = null;
            Release(replacementParagraphRange); replacementParagraphRange = null;
            Release(replacementParagraph); replacementParagraph = null;
            Release(replacementParagraphs); replacementParagraphs = null;
            Release(paragraphRange); paragraphRange = null;
            Release(paragraph); paragraph = null;
            Release(paragraphs); paragraphs = null;
            Release(probe); probe = null;
            Release(content); content = null;

            paragraphRange = document.Range(deleteStart, deleteEnd);
            paragraphRange.Delete();
            var deleted = document.Paragraphs.Count < paragraphCountBefore;
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    former-true-display-anchor-after-ole removed={deleted} range={deleteStart}:{deleteEnd}");
            return deleted;
        }
        finally
        {
            Release(anchoredShapes);
            Release(frames);
            Release(bookmarks);
            Release(inlineShapes);
            Release(maths);
            Release(fields);
            Release(format);
            Release(replacementParagraphRange);
            Release(replacementParagraph);
            Release(replacementParagraphs);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
            Release(content);
        }
    }

    private static bool RemoveNativeDisplayNumberShapeAndAnchor(
        Document document,
        string formulaId,
        bool removeAnchorParagraph = true)
    {
        var removed = false;
        var capturedAnchorStart = -1;
        var capturedAnchorEnd = -1;
        Bookmarks? preDeleteBookmarks = null;
        Bookmark? preDeleteAnchorBookmark = null;
        Range? preDeleteAnchorRange = null;
        Paragraphs? preDeleteParagraphs = null;
        Paragraph? preDeleteParagraph = null;
        Range? preDeleteParagraphRange = null;
        try
        {
            if (removeAnchorParagraph)
            {
                preDeleteBookmarks = document.Bookmarks;
                var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
                if (preDeleteBookmarks.Exists(anchorName))
                {
                    preDeleteAnchorBookmark = preDeleteBookmarks[anchorName];
                    preDeleteAnchorRange = preDeleteAnchorBookmark.Range;
                    preDeleteParagraphs = preDeleteAnchorRange.Paragraphs;
                    if (preDeleteParagraphs.Count == 1)
                    {
                        preDeleteParagraph = preDeleteParagraphs[1];
                        preDeleteParagraphRange = preDeleteParagraph.Range;
                        capturedAnchorStart = preDeleteParagraphRange.Start;
                        capturedAnchorEnd = preDeleteParagraphRange.End;
                    }
                }
            }
        }
        catch
        {
            capturedAnchorStart = -1;
            capturedAnchorEnd = -1;
        }
        finally
        {
            Release(preDeleteParagraphRange);
            Release(preDeleteParagraph);
            Release(preDeleteParagraphs);
            Release(preDeleteAnchorRange);
            Release(preDeleteAnchorBookmark);
            Release(preDeleteBookmarks);
        }

        Shape? shape = null;
        try
        {
            shape = FindNativeDisplayNumberShape(document, formulaId);
            if (shape is not null)
            {
                DeleteBookmarkOnly(document, EquationBookmarkName(formulaId));
                shape.Delete();
                removed = true;
            }
        }
        catch (COMException)
        {
            // Continue with bookmark/anchor cleanup. A protected Shape can still
            // leave enough deterministic identity for the next reconciliation.
        }
        finally { Release(shape); }

        if (!removeAnchorParagraph)
            return removed;

        if (TryDeleteNativeDisplayAnchorParagraphFromFrozenBounds(
                document,
                formulaId,
                capturedAnchorStart,
                capturedAnchorEnd))
        {
            ClearNativeDisplayAnchorCommitMarker(document, formulaId);
            return true;
        }

        Bookmarks? bookmarks = null;
        Bookmark? anchorBookmark = null;
        Range? anchorRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        OMaths? maths = null;
        InlineShapes? inlineShapes = null;
        Range? content = null;
        try
        {
            bookmarks = document.Bookmarks;
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            if (bookmarks.Exists(anchorName))
            {
                anchorBookmark = bookmarks[anchorName];
                anchorRange = anchorBookmark.Range;
            }
            else if (capturedAnchorStart >= 0)
            {
                content = document.Content;
                if (capturedAnchorStart < content.Start
                    || capturedAnchorStart >= content.End)
                    return removed;
                var probeEnd = Math.Min(
                    content.End,
                    Math.Max(capturedAnchorStart + 1, capturedAnchorEnd));
                anchorRange = document.Range(capturedAnchorStart, probeEnd);
            }
            else
            {
                return removed;
            }

            paragraphs = anchorRange.Paragraphs;
            if (paragraphs.Count != 1)
            {
                try { anchorBookmark?.Delete(); } catch { }
                return removed;
            }
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (capturedAnchorStart >= 0
                && paragraphRange.Start != capturedAnchorStart)
                return removed;
            fields = paragraphRange.Fields;
            maths = paragraphRange.OMaths;
            inlineShapes = paragraphRange.InlineShapes;
            var text = (paragraphRange.Text ?? string.Empty)
                .Trim('\r', '\n', '\t', '\v', ' ');
            if (text.Length == 0
                && fields.Count == 0
                && maths.Count == 0
                && inlineShapes.Count == 0)
            {
                var paragraphCountBeforeDelete = document.Paragraphs.Count;
                paragraphRange.Delete();
                var paragraphDeleted = document.Paragraphs.Count < paragraphCountBeforeDelete;
                if (paragraphDeleted)
                {
                    DeleteBookmarkOnly(document, NativeDisplayAnchorBookmarkName(formulaId));
                    ClearNativeDisplayAnchorCommitMarker(document, formulaId);
                    removed = true;
                }
                else
                {
                    RestoreNativeDisplayAnchorBookmarkForDeferredCleanup(
                        document,
                        formulaId,
                        capturedAnchorStart);
                }
            }
            return removed;
        }
        catch
        {
            try { DeleteBookmarkOnly(document, NativeDisplayAnchorBookmarkName(formulaId)); }
            catch { }
            return removed;
        }
        finally
        {
            Release(content);
            Release(inlineShapes);
            Release(maths);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(anchorRange);
            Release(anchorBookmark);
            Release(bookmarks);
        }
    }

    private static bool TryDeleteNativeDisplayAnchorParagraphFromFrozenBounds(
        Document document,
        string formulaId,
        int capturedStart,
        int capturedEnd)
    {
        Bookmarks? bookmarks = null;
        Bookmark? anchorBookmark = null;
        Range? anchorRange = null;
        Range? content = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        OMaths? maths = null;
        InlineShapes? inlineShapes = null;
        Shapes? shapes = null;
        Shape? shape = null;
        Range? shapeAnchor = null;
        Bookmark? formulaBookmark = null;
        Range? formulaRange = null;
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Range? deleteRange = null;
        var deleteStart = -1;
        var deleteEnd = -1;
        var paragraphCountBefore = -1;
        try
        {
            bookmarks = document.Bookmarks;
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            if (bookmarks.Exists(anchorName))
            {
                anchorBookmark = bookmarks[anchorName];
                anchorRange = anchorBookmark.Range;
            }
            else
            {
                if (capturedStart < 0 || capturedEnd <= capturedStart)
                    return false;
                content = document.Content;
                if (capturedStart < content.Start || capturedStart >= content.End)
                    return false;
                anchorRange = document.Range(
                    capturedStart,
                    Math.Min(capturedEnd, content.End));
            }

            paragraphs = anchorRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (capturedStart >= 0 && paragraphRange.Start != capturedStart)
                return false;
            if ((bool)paragraphRange.get_Information(
                    WdInformation.wdWithInTable))
                return false;

            var formulaBookmarkName = WordOmmlFormulaStore.BookmarkName(formulaId);
            if (bookmarks.Exists(formulaBookmarkName))
            {
                formulaBookmark = bookmarks[formulaBookmarkName];
                formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
                formulaParagraphs = formulaRange.Paragraphs;
                if (formulaParagraphs.Count != 1) return false;
                formulaParagraph = formulaParagraphs[1];
                formulaParagraphRange = formulaParagraph.Range;
                if (paragraphRange.End != formulaParagraphRange.Start)
                    return false;
            }

            fields = paragraphRange.Fields;
            maths = paragraphRange.OMaths;
            inlineShapes = paragraphRange.InlineShapes;
            var text = (paragraphRange.Text ?? string.Empty)
                .Trim('\r', '\n', '\t', '\v', '\u0001', ' ');
            if (text.Length != 0
                || fields.Count != 0
                || maths.Count != 0
                || inlineShapes.Count != 0)
                return false;

            shapes = document.Shapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shapeAnchor); shapeAnchor = null;
                Release(shape); shape = shapes[index];
                try
                {
                    shapeAnchor = shape.Anchor;
                    if (shapeAnchor.Start >= paragraphRange.Start
                        && shapeAnchor.Start < paragraphRange.End)
                        return false;
                }
                catch { return false; }
            }

            deleteStart = paragraphRange.Start;
            deleteEnd = paragraphRange.End;
            paragraphCountBefore = document.Paragraphs.Count;

            // Keep VTEqAnc_ alive until the physical paragraph is actually gone.
            // Word can expose a just-deleted floating Shape for one dispatcher turn;
            // paragraph deletion may then be deferred/no-op. Deleting the bookmark
            // first would lose the only durable locator for the next cleanup turn.
            // A successful paragraph deletion removes or relocates the bookmark;
            // we explicitly clear any survivor only after success.

            // Release all RCWs that still point into the anchor paragraph. On Word
            // 2021, deleting the Shape and then deleting its paragraph while a stale
            // Shape/Paragraph collection is alive can return success but retain the
            // empty one-point paragraph.
            Release(shapeAnchor); shapeAnchor = null;
            Release(shape); shape = null;
            Release(shapes); shapes = null;
            Release(inlineShapes); inlineShapes = null;
            Release(maths); maths = null;
            Release(fields); fields = null;
            Release(paragraphRange); paragraphRange = null;
            Release(paragraph); paragraph = null;
            Release(paragraphs); paragraphs = null;
            Release(anchorRange); anchorRange = null;
            Release(anchorBookmark); anchorBookmark = null;
            Release(content); content = null;

            deleteRange = document.Range(deleteStart, deleteEnd);
            deleteRange.Delete();
            Release(deleteRange); deleteRange = null;
            var deleted = document.Paragraphs.Count < paragraphCountBefore;
            if (!deleted)
            {
                content = document.Content;
                if (deleteStart < content.End)
                {
                    anchorRange = document.Range(
                        deleteStart,
                        Math.Min(Math.Max(deleteStart + 1, deleteEnd), content.End));
                    paragraphs = anchorRange.Paragraphs;
                    if (paragraphs.Count == 1)
                    {
                        paragraph = paragraphs[1];
                        paragraphRange = paragraph.Range.Duplicate;
                        var retryText = (paragraphRange.Text ?? string.Empty)
                            .Trim('\r', '\n', '\t', '\v', '\u0001', ' ');
                        if (paragraphRange.Start == deleteStart
                            && retryText.Length == 0
                            && paragraphRange.Fields.Count == 0
                            && paragraphRange.OMaths.Count == 0
                            && paragraphRange.InlineShapes.Count == 0)
                            paragraphRange.Delete();
                    }
                }
                deleted = document.Paragraphs.Count < paragraphCountBefore;
            }
            if (deleted)
            {
                DeleteBookmarkOnly(document, NativeDisplayAnchorBookmarkName(formulaId));
                ClearNativeDisplayAnchorCommitMarker(document, formulaId);
            }
            else
            {
                RestoreNativeDisplayAnchorBookmarkForDeferredCleanup(
                    document,
                    formulaId,
                    deleteStart);
            }
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    true-display-anchor-paragraph-removed formulaId={formulaId} deleted={deleted} range={deleteStart}:{deleteEnd}");
            return deleted;
        }
        catch (Exception error)
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    true-display-anchor-paragraph-remove-failed formulaId={formulaId} error={error.GetType().Name}:{error.Message}");
            return false;
        }
        finally
        {
            Release(deleteRange);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(formulaRange);
            Release(formulaBookmark);
            Release(shapeAnchor);
            Release(shape);
            Release(shapes);
            Release(inlineShapes);
            Release(maths);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(content);
            Release(anchorRange);
            Release(anchorBookmark);
            Release(bookmarks);
        }
    }

    private static bool EnsureRetiredNativeDisplayAnchorParagraphRemoved(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Shape? legacyShape = null;
        Range? activeFormulaRange = null;
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Range? content = null;
        Range? probe = null;
        Paragraphs? candidateParagraphs = null;
        Paragraph? candidateParagraph = null;
        Range? candidateRange = null;
        Fields? candidateFields = null;
        OMaths? candidateMaths = null;
        InlineShapes? candidateInlineShapes = null;
        Frames? candidateFrames = null;
        Bookmarks? candidateBookmarks = null;
        ParagraphFormat? candidateFormat = null;
        Microsoft.Office.Interop.Word.Font? candidateFont = null;
        TabStops? candidateTabStops = null;
        Borders? candidateBorders = null;
        ListFormat? candidateListFormat = null;
        ShapeRange? anchoredShapes = null;
        Range? deleteRange = null;
        var deleteStart = -1;
        var deleteEnd = -1;
        var paragraphCountBefore = -1;
        try
        {
            legacyShape = FindNativeDisplayNumberShape(document, formulaId);
            if (legacyShape is not null)
            {
                // The drawing deletion has not committed yet. Let the next Office
                // turn reacquire both the Shape and its one-point paragraph.
                return false;
            }

            var hasAnchorEvidence = HasNativeDisplayAnchorBookmark(document, formulaId)
                || HasNativeDisplayAnchorCommitMarker(document, formulaId);

            // Word can discard both retired anchor bookmarks while atomically
            // replacing the adjacent OMath, yet leave the physical one-point empty
            // paragraph behind. Continue to the strict structural/style probe even
            // without a surviving marker; an ordinary preceding paragraph will be
            // rejected and treated as already clean.
            activeFormulaRange = ResolveSingleNativeOmmlRange(formulaRange);
            formulaParagraphs = activeFormulaRange.Paragraphs;
            if (formulaParagraphs.Count != 1) return false;
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range.Duplicate;
            if ((bool)formulaParagraphRange.get_Information(
                    WdInformation.wdWithInTable))
                return false;

            content = document.Content;
            if (formulaParagraphRange.Start <= content.Start)
            {
                DeleteBookmarkOnly(document, NativeDisplayAnchorBookmarkName(formulaId));
                ClearNativeDisplayAnchorCommitMarker(document, formulaId);
                return true;
            }

            probe = document.Range(
                formulaParagraphRange.Start - 1,
                formulaParagraphRange.Start);
            candidateParagraphs = probe.Paragraphs;
            if (candidateParagraphs.Count != 1) return !hasAnchorEvidence;
            candidateParagraph = candidateParagraphs[1];
            candidateRange = candidateParagraph.Range.Duplicate;
            if (candidateRange.End != formulaParagraphRange.Start
                || (bool)candidateRange.get_Information(
                    WdInformation.wdWithInTable))
                return !hasAnchorEvidence;

            var candidateText = (candidateRange.Text ?? string.Empty)
                .Trim('\r', '\n', '\t', '\v', '\u0001', ' ');
            candidateFields = candidateRange.Fields;
            candidateMaths = candidateRange.OMaths;
            candidateInlineShapes = candidateRange.InlineShapes;
            candidateFrames = candidateRange.Frames;
            if (candidateText.Length != 0
                || candidateFields.Count != 0
                || candidateMaths.Count != 0
                || candidateInlineShapes.Count != 0
                || candidateFrames.Count != 0)
                return !hasAnchorEvidence;

            try
            {
                anchoredShapes = candidateRange.ShapeRange;
                if (anchoredShapes.Count != 0) return !hasAnchorEvidence;
            }
            catch (COMException)
            {
                // Word throws when a paragraph has no anchored Shape.
            }

            candidateFormat = candidateRange.ParagraphFormat;
            candidateFont = candidateRange.Font;
            candidateTabStops = candidateFormat.TabStops;
            candidateBorders = candidateRange.Borders;
            candidateListFormat = candidateRange.ListFormat;
            var candidateFontSize = candidateFont.Size;
            var hasRetiredAnchorStyle =
                candidateFormat.Alignment == WdParagraphAlignment.wdAlignParagraphLeft
                && candidateFormat.LineSpacingRule == WdLineSpacing.wdLineSpaceExactly
                && Math.Abs(
                    candidateFormat.LineSpacing
                    - NativeDisplayAnchorLineSpacingPoints) <= 0.1f
                && Math.Abs(candidateFormat.SpaceBefore) <= 0.1f
                && Math.Abs(candidateFormat.SpaceAfter) <= 0.1f
                && candidateFormat.KeepTogether == -1
                && candidateFormat.KeepWithNext == -1
                && candidateFormat.WidowControl == 0
                && candidateFormat.PageBreakBefore == 0
                && candidateFontSize > 0f
                && candidateFontSize < 999999f
                && Math.Abs(
                    candidateFontSize
                    - NativeDisplayAnchorLineSpacingPoints) <= 0.1f
                && candidateFont.Hidden == 0
                && Math.Abs(candidateFont.Position) <= 0.1f
                // Word exposes document-default tab stops through this collection
                // even after ClearAll(); their count is locale/page-width dependent
                // and is therefore not an anchor-identity signal.
                && candidateBorders.Enable == 0
                && candidateListFormat.ListType == WdListType.wdListNoNumbering;
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    retired-native-display-anchor-probe formulaId={formulaId} evidence={hasAnchorEvidence} range={candidateRange.Start}:{candidateRange.End} alignment={candidateFormat.Alignment} lineRule={candidateFormat.LineSpacingRule} line={candidateFormat.LineSpacing:0.###} before={candidateFormat.SpaceBefore:0.###} after={candidateFormat.SpaceAfter:0.###} keep={candidateFormat.KeepTogether}/{candidateFormat.KeepWithNext} widow={candidateFormat.WidowControl} pageBreak={candidateFormat.PageBreakBefore} font={candidateFontSize:0.###} hidden={candidateFont.Hidden} position={candidateFont.Position:0.###} tabs={candidateTabStops.Count} borders={candidateBorders.Enable} list={candidateListFormat.ListType} matched={hasRetiredAnchorStyle}");
            if (!hasRetiredAnchorStyle)
                return !hasAnchorEvidence;

            // At this point the retained migration markers plus the exact old
            // one-point anchor fingerprint prove ownership. Remove only those two
            // markers, then require the paragraph to contain no unrelated bookmark
            // before deleting its paragraph mark.
            DeleteBookmarkOnly(document, NativeDisplayAnchorBookmarkName(formulaId));
            ClearNativeDisplayAnchorCommitMarker(document, formulaId);
            Release(candidateBookmarks);
            candidateBookmarks = candidateRange.Bookmarks;
            if (candidateBookmarks.Count != 0)
            {
                if (hasAnchorEvidence)
                {
                    RestoreNativeDisplayAnchorBookmarkForDeferredCleanup(
                        document,
                        formulaId,
                        candidateRange.Start);
                    return false;
                }
                return true;
            }

            deleteStart = candidateRange.Start;
            deleteEnd = candidateRange.End;
            paragraphCountBefore = document.Paragraphs.Count;

            Release(anchoredShapes); anchoredShapes = null;
            Release(candidateListFormat); candidateListFormat = null;
            Release(candidateBorders); candidateBorders = null;
            Release(candidateTabStops); candidateTabStops = null;
            Release(candidateFont); candidateFont = null;
            Release(candidateFormat); candidateFormat = null;
            Release(candidateBookmarks); candidateBookmarks = null;
            Release(candidateFrames); candidateFrames = null;
            Release(candidateInlineShapes); candidateInlineShapes = null;
            Release(candidateMaths); candidateMaths = null;
            Release(candidateFields); candidateFields = null;
            Release(candidateRange); candidateRange = null;
            Release(candidateParagraph); candidateParagraph = null;
            Release(candidateParagraphs); candidateParagraphs = null;
            Release(probe); probe = null;
            Release(content); content = null;
            Release(formulaParagraphRange); formulaParagraphRange = null;
            Release(formulaParagraph); formulaParagraph = null;
            Release(formulaParagraphs); formulaParagraphs = null;
            Release(activeFormulaRange); activeFormulaRange = null;

            deleteRange = document.Range(deleteStart, deleteEnd);
            deleteRange.Delete();
            Release(deleteRange); deleteRange = null;
            var deleted = document.Paragraphs.Count < paragraphCountBefore;
            if (!deleted)
            {
                RestoreNativeDisplayAnchorBookmarkForDeferredCleanup(
                    document,
                    formulaId,
                    deleteStart);
            }
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    retired-native-display-anchor-cleanup formulaId={formulaId} deleted={deleted} range={deleteStart}:{deleteEnd}");
            return deleted;
        }
        catch (Exception error)
        {
            if (deleteStart >= 0)
                RestoreNativeDisplayAnchorBookmarkForDeferredCleanup(
                    document,
                    formulaId,
                    deleteStart);
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"    retired-native-display-anchor-cleanup-failed formulaId={formulaId} error={error.GetType().Name}:{error.Message}");
            return false;
        }
        finally
        {
            Release(deleteRange);
            Release(anchoredShapes);
            Release(candidateListFormat);
            Release(candidateBorders);
            Release(candidateTabStops);
            Release(candidateFont);
            Release(candidateFormat);
            Release(candidateBookmarks);
            Release(candidateFrames);
            Release(candidateInlineShapes);
            Release(candidateMaths);
            Release(candidateFields);
            Release(candidateRange);
            Release(candidateParagraph);
            Release(candidateParagraphs);
            Release(probe);
            Release(content);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(activeFormulaRange);
            Release(legacyShape);
        }
    }

    private static void RestoreNativeDisplayAnchorBookmarkForDeferredCleanup(
        Document document,
        string formulaId,
        int paragraphStart)
    {
        if (paragraphStart < 0) return;
        Bookmarks? bookmarks = null;
        Range? content = null;
        Range? anchor = null;
        try
        {
            bookmarks = document.Bookmarks;
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            if (!bookmarks.Exists(anchorName))
            {
                content = document.Content;
                if (paragraphStart < content.Start || paragraphStart >= content.End)
                    return;
                anchor = document.Range(paragraphStart, paragraphStart);
                bookmarks.Add(anchorName, anchor);
            }
            if (!HasNativeDisplayAnchorCommitMarker(document, formulaId))
                MarkNativeDisplayAnchorCommitted(document, formulaId);
        }
        catch
        {
            // The next document-open migration can still identify a surviving
            // Shape-era paragraph from its retained one-point formatting. This
            // best-effort marker restoration primarily covers Word's one-turn
            // delayed paragraph deletion after removing a floating Shape.
        }
        finally
        {
            Release(anchor);
            Release(content);
            Release(bookmarks);
        }
    }

    private static bool HasNativeDisplayAnchorBookmark(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            return bookmarks.Exists(NativeDisplayAnchorBookmarkName(formulaId));
        }
        catch
        {
            return true;
        }
        finally { Release(bookmarks); }
    }

    private static Range? TryResolveNativeDisplayFormulaOwnerRange(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? formulaBookmark = null;
        Range? formulaRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? ownerRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var formulaBookmarkName = WordOmmlFormulaStore.BookmarkName(formulaId);
            if (!bookmarks.Exists(formulaBookmarkName)) return null;
            formulaBookmark = bookmarks[formulaBookmarkName];
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            if (!IsPureTrueDisplayFormulaParagraph(formulaRange)) return null;

            // The numbering owner is identified by the durable VTOMML FormulaId,
            // not by the state of its separate visible Shape or by transient
            // Shape-anchor coordinates. During a multi-formula
            // conversion Word can normalize later floating anchors before the next
            // COM turn; requiring adjacency here incorrectly falls through to the
            // TextFrame paragraph, which contains no OMath. Anchor adjacency remains
            // a separate health invariant and is repaired/validated independently.
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            ownerRange = paragraph.Range.Duplicate;
            var result = ownerRange;
            ownerRange = null;
            return result;
        }
        catch { return null; }
        finally
        {
            Release(ownerRange);
            Release(paragraph);
            Release(paragraphs);
            Release(formulaRange);
            Release(formulaBookmark);
            Release(bookmarks);
        }
    }

    private static bool HasVisibleEquationNumberBookmark(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            return bookmarks.Exists(EquationBookmarkName(formulaId));
        }
        catch { return false; }
        finally { Release(bookmarks); }
    }

    private static bool HasNativeDisplayAnchorCommitMarker(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            return bookmarks.Exists(
                NativeDisplayAnchorCommitBookmarkName(formulaId));
        }
        catch { return false; }
        finally { Release(bookmarks); }
    }

    private static void MarkNativeDisplayAnchorCommitted(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? anchorBookmark = null;
        Bookmark? markerBookmark = null;
        Range? markerRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            var markerName = NativeDisplayAnchorCommitBookmarkName(formulaId);
            if (!bookmarks.Exists(anchorName))
                throw new InvalidOperationException(
                    "The numbered OMML Shape anchor bookmark is missing while marking its committed Office turn.");
            if (bookmarks.Exists(markerName)) return;
            anchorBookmark = bookmarks[anchorName];
            markerRange = anchorBookmark.Range.Duplicate;
            markerRange.Collapse(WdCollapseDirection.wdCollapseStart);
            markerBookmark = bookmarks.Add(markerName, markerRange);
        }
        finally
        {
            Release(markerBookmark);
            Release(markerRange);
            Release(anchorBookmark);
            Release(bookmarks);
        }
    }

    private static void ClearNativeDisplayAnchorCommitMarker(
        Document document,
        string formulaId) =>
        DeleteBookmarkOnly(
            document,
            NativeDisplayAnchorCommitBookmarkName(formulaId));

    private static string NativeDisplayAnchorBookmarkName(string formulaId) =>
        NativeDisplayAnchorBookmarkPrefix + NormalizeFormulaIdForBookmark(formulaId);

    private static string NativeDisplayAnchorCommitBookmarkName(string formulaId) =>
        NativeDisplayAnchorCommitBookmarkPrefix
        + NormalizeFormulaIdForBookmark(formulaId);

    private static string NativeDisplayNumberShapeName(string formulaId) =>
        NativeDisplayNumberShapeNamePrefix + NormalizeFormulaIdForBookmark(formulaId);

    private static string NativeDisplayNumberShapeAlternativeText(string formulaId) =>
        NativeDisplayNumberShapeAlternativeTextPrefix
        + NormalizeFormulaIdForBookmark(formulaId);

    private static string NormalizeFormulaIdForBookmark(string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var parsed))
            throw new InvalidDataException(
                "VisualTeX numbered OMML formulaId must be a UUID.");
        return parsed.ToString("N");
    }
    private static Range? TryEnsureNativeHashSequenceTypingParagraph(
        Document document,
        string formulaId)
    {
        Bookmark? formulaBookmark = null;
        Range? formulaRange = null;
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Range? content = null;
        Range? probe = null;
        Paragraphs? nextParagraphs = null;
        Paragraph? nextParagraph = null;
        Range? nextParagraphRange = null;
        Paragraph? appendedParagraph = null;
        OMaths? nextMaths = null;
        try
        {
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            if (formulaBookmark is null) return null;
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            var openXml = formulaRange.WordOpenXML ?? string.Empty;
            if (!WordOmmlConverter.HasVisualTeXNativeEquationNumber(openXml)
                || openXml.IndexOf("SEQ VisualTeXEquation", StringComparison.OrdinalIgnoreCase) < 0
                || openXml.IndexOf("<m:eqArr", StringComparison.OrdinalIgnoreCase) < 0)
                return null;

            formulaParagraphs = formulaRange.Paragraphs;
            if (formulaParagraphs.Count != 1)
                throw new InvalidDataException(
                    $"Numbered OMML {formulaId} no longer occupies one display paragraph.");
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range.Duplicate;
            if ((bool)formulaParagraphRange.get_Information(WdInformation.wdWithInTable))
                throw new InvalidDataException(
                    $"Numbered OMML {formulaId} unexpectedly moved into a Word table.");

            content = document.Content;
            var nextStart = formulaParagraphRange.End;
            if (nextStart < content.End)
            {
                probe = document.Range(nextStart, Math.Min(content.End, nextStart + 1));
                nextParagraphs = probe.Paragraphs;
                if (nextParagraphs.Count == 1)
                {
                    nextParagraph = nextParagraphs[1];
                    nextParagraphRange = nextParagraph.Range.Duplicate;
                    nextMaths = nextParagraphRange.OMaths;
                    if (nextParagraphRange.Start >= nextStart && nextMaths.Count == 0)
                    {
                        var existing = document.Range(
                            nextParagraphRange.Start,
                            nextParagraphRange.Start);
                        return existing;
                    }
                }
            }

            // The display equation is the final paragraph. Paragraphs.Add creates
            // a new main-story paragraph without deriving an insertion point from
            // the OMath boundary, which avoids Word 0x800A1831.
            appendedParagraph = document.Paragraphs.Add();
            Release(nextParagraphRange);
            nextParagraphRange = appendedParagraph.Range.Duplicate;
            Release(nextMaths);
            nextMaths = nextParagraphRange.OMaths;
            if (nextMaths.Count != 0)
                throw new InvalidDataException(
                    "Word absorbed the numbered-OMML typing paragraph into OMath.");
            return document.Range(nextParagraphRange.Start, nextParagraphRange.Start);
        }
        finally
        {
            Release(nextMaths);
            Release(appendedParagraph);
            Release(nextParagraphRange);
            Release(nextParagraph);
            Release(nextParagraphs);
            Release(probe);
            Release(content);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(formulaRange);
            Release(formulaBookmark);
        }
    }

    private static bool TryRemoveNativeHashSequenceCaptionBookmarks(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? numberBookmark = null;
        Range? numberRange = null;
        OMaths? maths = null;
        try
        {
            bookmarks = document.Bookmarks;
            var numberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(numberName)) return false;
            numberBookmark = bookmarks[numberName];
            numberRange = numberBookmark.Range;
            maths = numberRange.OMaths;
            if (maths.Count == 0) return false;

            // In the native #() host VTEqNum_/VTEqCap_ are aliases around the
            // mathematical SEQ result. Removing a caption means removing only the
            // aliases; deleting either bookmark Range would delete the OMath field
            // and potentially the complete equation.
            DeleteBookmarkOnly(document, NativeCaptionBookmarkName(formulaId));
            DeleteBookmarkOnly(document, numberName);
            return true;
        }
        finally
        {
            Release(maths);
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
        }
    }

    private static bool TryFinalizeRetiredNativeDisplayAnchorCleanup(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Shape? shape = null;
        try
        {
            shape = FindNativeDisplayNumberShape(document, formulaId);
            if (shape is not null)
                return false;

            bookmarks = document.Bookmarks;
            var anchorName = NativeDisplayAnchorBookmarkName(formulaId);
            if (!bookmarks.Exists(anchorName))
                return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(shape);
            Release(bookmarks);
        }

        // The drawing has disappeared from the live Shapes collection, so the
        // bookmark now identifies only the retired empty anchor paragraph. Delete
        // it through the same conservative structural guard used by migration.
        if (!TryDeleteNativeDisplayAnchorParagraphFromFrozenBounds(
                document,
                formulaId,
                capturedStart: -1,
                capturedEnd: -1))
            return false;

        try
        {
            bookmarks = document.Bookmarks;
            return !bookmarks.Exists(NativeDisplayAnchorBookmarkName(formulaId));
        }
        catch
        {
            return false;
        }
        finally { Release(bookmarks); }
    }

    private static bool TryRemoveLegacyNativeDisplayShapeBeforeHashMigrationV2(
        Document document,
        string formulaId)
    {
        Shape? shape = null;
        try
        {
            shape = FindNativeDisplayNumberShape(document, formulaId);
            if (shape is null) return false;
        }
        finally { Release(shape); }

        // A Shape-era numbered OMML is migration input only. Preserve a second
        // collapsed marker before deleting the floating object: Word can discard
        // VTEqAnc_ in the same COM turn even when it retains the physical one-point
        // paragraph. The finalizer clears both markers only after that paragraph is
        // verifiably gone.
        if (!HasNativeDisplayAnchorCommitMarker(document, formulaId))
            MarkNativeDisplayAnchorCommitted(document, formulaId);

        // Remove the floating visible REF and its dedicated anchor paragraph first,
        // then remove the hidden SEQ caption. The mathematical formula itself is
        // recovered by its VTOMML metadata/fingerprint and rebuilt as native #(...).
        RemoveNativeDisplayNumberShapeAndAnchor(document, formulaId);
        RemoveVisibleEquationNumberLegacy(document, formulaId);
        RemoveNativeCaption(document, formulaId);
        return true;
    }

    private static string ResolveNativeHashSequencePrefixV3(
        Document document,
        int formulaPosition,
        EquationNumberFormat format)
    {
        var anchors = format.UsesHeading
            ? GetHeadingNumberAnchors(document, format.HeadingLevel)
            : Array.Empty<HeadingNumberAnchor>();
        return ResolveEquationNumberScope(
            formulaPosition,
            format,
            anchors).Prefix;
    }

    private static bool IsNativeHashSequencePlanHealthyV3(
        Document document,
        Range formulaRange,
        string formulaId,
        EquationNumberFormat format,
        string prefix,
        bool updateField)
    {
        if (!IsHealthyNumberedNativeOmmlHashSequenceHost(
                document,
                formulaRange,
                formulaId,
                requireCurrentFieldResult: !updateField))
            return false;
        if (!NativeOmmlHashSequenceFormatMatches(
                document,
                formulaRange,
                formulaId,
                format,
                prefix))
            return false;
        if (!updateField) return true;
        UpdateNativeOmmlHashSequenceField(formulaRange);
        return IsHealthyNumberedNativeOmmlHashSequenceHost(
            document,
            formulaRange,
            formulaId,
            requireCurrentFieldResult: true);
    }

    private static bool IsNativeHashSequenceHostForFinalizationV2(
        Document document,
        Range formulaRange,
        string formulaId,
        bool updateField)
    {
        try
        {
            var format = ReadEquationNumberFormat(document);
            var prefix = ResolveNativeHashSequencePrefixV3(
                document,
                formulaRange.Start,
                format);
            return IsNativeHashSequencePlanHealthyV3(
                document,
                formulaRange,
                formulaId,
                format,
                prefix,
                updateField);
        }
        catch
        {
            return false;
        }
    }

    internal static bool IsSafeNativeHashSequenceOmmlForConversion(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        try
        {
            var format = ReadEquationNumberFormat(document);
            var prefix = ResolveNativeHashSequencePrefixV3(
                document,
                formulaRange.Start,
                format);
            return IsNativeHashSequencePlanHealthyV3(
                document,
                formulaRange,
                formulaId,
                format,
                prefix,
                updateField: false);
        }
        catch
        {
            return false;
        }
    }

}
