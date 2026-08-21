using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal enum EquationReferenceSource
{
    VisualTeX,
    MathType,
}

internal sealed class EquationReferenceTarget
{
    public EquationReferenceTarget(
        string formulaId,
        int nativeReferenceItem,
        string numberText,
        string latexPreview,
        int position,
        EquationReferenceSource source = EquationReferenceSource.VisualTeX)
    {
        FormulaId = formulaId;
        NativeReferenceItem = nativeReferenceItem;
        NumberText = numberText;
        LatexPreview = latexPreview;
        Position = position;
        Source = source;
    }

    public string FormulaId { get; }
    public int NativeReferenceItem { get; }
    public string NumberText { get; }
    public string LatexPreview { get; }
    public int Position { get; }
    public EquationReferenceSource Source { get; }

    public override string ToString() => Source == EquationReferenceSource.MathType
        ? $"{NumberText}    {LatexPreview}"
        : $"({NumberText})    {LatexPreview}";
}

internal enum EquationReferenceStyle
{
    Parenthesized,
    EquationPrefix,
    NumberOnly,
}

internal sealed class EquationNumberFormat
{
    public const string ContinuousId = "continuous";
    public const string Heading1DotId = "heading1-dot";
    public const string Heading1DashId = "heading1-dash";
    public const string Heading2DotId = "heading2-dot";
    public const string Heading2DashId = "heading2-dash";

    private EquationNumberFormat(
        string id,
        string displayName,
        int headingLevel,
        string separator)
    {
        Id = id;
        DisplayName = displayName;
        HeadingLevel = headingLevel;
        Separator = separator;
    }

    public string Id { get; }
    public string DisplayName { get; }
    public int HeadingLevel { get; }
    public string Separator { get; }
    public bool UsesHeading => HeadingLevel > 0;

    public static EquationNumberFormat Resolve(string? id) => id switch
    {
        Heading1DotId => new EquationNumberFormat(Heading1DotId, "按章编号（1.1）", 1, "."),
        Heading1DashId => new EquationNumberFormat(Heading1DashId, "按章编号（1-1）", 1, "-"),
        Heading2DotId => new EquationNumberFormat(Heading2DotId, "按节编号（1.1.1）", 2, "."),
        Heading2DashId => new EquationNumberFormat(Heading2DashId, "按节编号（1.1-1）", 2, "-"),
        _ => new EquationNumberFormat(ContinuousId, "全文连续编号（1）", 0, string.Empty),
    };
}

internal sealed class ResolvedEquationHeadingScope
{
    internal ResolvedEquationHeadingScope(
        int scopeStart,
        int scopeEnd,
        string numberText)
    {
        ScopeStart = scopeStart;
        ScopeEnd = scopeEnd;
        NumberText = numberText;
    }

    internal int ScopeStart { get; }
    internal int ScopeEnd { get; }
    internal string NumberText { get; }
}

internal static class WordEquationNumbering
{
    private const int WdTabAlignmentCenter = 1;
    private const int WdTabAlignmentRight = 2;
    private const int WdTabLeaderSpaces = 0;
    private const int WdFieldEmpty = -1;
    private const string EquationNumberFontName = "Cambria Math";
    private const string LegacyEquationSequenceName = "VisualTeXEquation";
    private const string EquationBookmarkPrefix = "VTEq_";
    private const string NativeCaptionBookmarkPrefix = "VTEqCap_";
    private const string NativeNumberBookmarkPrefix = "VTEqNum_";
    private const string EquationNumberFormatVariableName = "VisualTeXEquationNumberFormat";
    private const string EquationNumberTailFormulaVariableName = "VisualTeXEquationNumberTailFormulaId";
    private const string CompactTypingTailBookmarkName = "VTNumberedTypingTail";
    private const float CompactTypingTailFontSizePoints = 1f;
    private const float CompactTypingTailLineSpacingPoints = 1f;
    private const string UserPreferenceRegistryPath = @"Software\VisualTeX\Word";
    private const string DefaultNumberedPreferenceName = "DefaultDisplayEquationNumbered";
    private const string DefaultNumberFormatPreferenceName = "DefaultEquationNumberFormat";
    private const string DefaultCreateObjectModePreferenceName = "DefaultCreateFormulaObjectMode";
    private const string DefaultMathTypeNumberPositionPreferenceName = "DefaultMathTypeNumberPosition";
    private static readonly object NumberingPerformanceTraceSync = new();

    private static void TraceNumberingPerformance(string message)
    {
        Console.WriteLine(message);
        var tracePath = Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE_PATH");
        if (string.IsNullOrWhiteSpace(tracePath)) return;
        try
        {
            lock (NumberingPerformanceTraceSync)
            {
                System.IO.File.AppendAllText(
                    tracePath,
                    $"{DateTimeOffset.Now:O} pid={System.Diagnostics.Process.GetCurrentProcess().Id} {message}{Environment.NewLine}");
            }
        }
        catch
        {
            // Diagnostic tracing must never affect numbering behavior.
        }
    }

    internal static bool GetDefaultDisplayEquationNumbered()
    {
        var value = ReadUserPreference(DefaultNumberedPreferenceName);
        if (value is int integer) return integer != 0;
        if (value is string text)
        {
            if (bool.TryParse(text, out var boolean)) return boolean;
            if (int.TryParse(text, out var numeric)) return numeric != 0;
        }
        return false;
    }

    internal static void SetDefaultDisplayEquationNumbered(bool numbered) =>
        WriteUserPreference(
            DefaultNumberedPreferenceName,
            numbered ? 1 : 0,
            RegistryValueKind.DWord);

    internal static string GetDefaultCreateObjectMode()
    {
        var value = ReadUserPreference(DefaultCreateObjectModePreferenceName) as string;
        return string.Equals(value, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal)
            ? FormulaOleContract.MathTypeOleMode
            : FormulaOleContract.NativeOleMode;
    }

    internal static void SetDefaultCreateObjectMode(string? objectMode) =>
        WriteUserPreference(
            DefaultCreateObjectModePreferenceName,
            string.Equals(objectMode, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal)
                ? FormulaOleContract.MathTypeOleMode
                : FormulaOleContract.NativeOleMode,
            RegistryValueKind.String);

    internal static string? TryGetDefaultMathTypeNumberPosition()
    {
        var value = ReadUserPreference(DefaultMathTypeNumberPositionPreferenceName) as string;
        if (string.Equals(value, "left", StringComparison.OrdinalIgnoreCase)) return "left";
        if (string.Equals(value, "right", StringComparison.OrdinalIgnoreCase)) return "right";
        return null;
    }

    internal static string GetDefaultMathTypeNumberPosition() =>
        TryGetDefaultMathTypeNumberPosition() ?? "right";

    internal static void SetDefaultMathTypeNumberPosition(string? position) =>
        WriteUserPreference(
            DefaultMathTypeNumberPositionPreferenceName,
            string.Equals(position, "left", StringComparison.OrdinalIgnoreCase)
                ? "left"
                : "right",
            RegistryValueKind.String);

    internal static string GetDefaultEquationNumberFormatId() =>
        EquationNumberFormat.Resolve(
            ReadUserPreference(DefaultNumberFormatPreferenceName) as string).Id;

    internal static void SetDefaultEquationNumberFormatPreference(string formatId)
    {
        var format = EquationNumberFormat.Resolve(formatId);
        WriteUserPreference(
            DefaultNumberFormatPreferenceName,
            format.Id,
            RegistryValueKind.String);
    }

    internal static string GetEquationNumberFormatId(Document document) =>
        ReadEquationNumberFormat(document).Id;

    internal static string GetEquationNumberFormatDisplayName(Document document) =>
        ReadEquationNumberFormat(document).DisplayName;

    internal static int UpdateEquationNumbers(Document document)
    {
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal)
            || string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal);
        var watch = tracePerformance
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        RepairNumberedDisplaySpacing(document);
        var fastPath = TryRefreshHealthyEquationNumbersInPlace(document, out var updated);
        var result = fastPath ? updated : Reconcile(document);
        if (watch is not null)
        {
            watch.Stop();
            Console.WriteLine(
                $"    [perf] UpdateEquationNumbers.{(fastPath ? "fast" : "reconcile")}: {result} formulas in {watch.ElapsedMilliseconds}ms");
        }
        return result;
    }

    private static bool TryRefreshHealthyEquationNumbersInPlace(
        Document document,
        out int updated)
    {
        updated = 0;
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal)
            || string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
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
                $"      [perf] update-numbers.{stage}: +{elapsed - checkpoint}ms ({elapsed}ms)");
            checkpoint = elapsed;
        }

        // The explicit command should still repair malformed/legacy numbering.
        // Only take the fast path when WordOpenXML shows one complete, unique
        // VTEq_/VTEqCap_/VTEqNum_ triplet for every currently numbered formula.
        // This check never activates an OLE object or reads OMML metadata.
        if (!TryReadHealthyEquationNumberArtifactsFromOpenXml(
                document,
                out var openXmlCaptions,
                out var referenceCounts))
            return false;
        TraceStage("artifact-inventory");

        var nativeSequenceName = GetNativeEquationSequenceName(document);
        var format = ReadEquationNumberFormat(document);
        // Continuous numbering only needs formula order, which WordOpenXML already
        // gives us. Heading-aware numbering additionally needs Word Range positions
        // so headings can be compared with formulas; resolve only the known
        // VTEqNum_* bookmarks instead of enumerating all 3N+ document bookmarks.
        var captions = format.UsesHeading
            ? ResolveNativeEquationCaptionPositions(document, openXmlCaptions)
            : openXmlCaptions;
        if (captions.Count != openXmlCaptions.Count)
            return false;
        TraceStage("caption-inventory");

        var changedFormulaNumbers = UpdateNativeEquationSequenceFieldsIncremental(
            document,
            nativeSequenceName,
            captions,
            format);
        TraceStage("sequence-refresh");
        if (changedFormulaNumbers.Count > 0)
        {
            UpdateHealthyNativeCrossReferencesAfterRenumbering(
                document,
                changedFormulaNumbers,
                knownReferenceCounts: referenceCounts);
            TraceStage("reference-refresh");
        }

        updated = captions.Count;
        return true;
    }

    internal static bool TryFinalizeHealthyConversionNumbering(
        Document document,
        out int updated)
    {
        updated = 0;
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
            TraceNumberingPerformance(
                $"[perf] ConversionFinalize.{stage}: +{elapsed - checkpoint}ms ({elapsed}ms)");
            checkpoint = elapsed;
        }
        try
        {
            // A conversion scaffold is intentionally only locally complete while
            // the item loop is running: its SEQ value and visible REF can still
            // reflect the pre-batch document order. Requiring the strict healthy
            // OpenXML inventory before the first refresh therefore creates a
            // circular dependency and unnecessarily falls back to Reconcile().
            // Prime all already-created captions once in document order, refresh
            // the existing REF fields, and only then run the strict health check.
            var nativeSequenceName = GetNativeEquationSequenceName(document);
            var captions = GetNativeEquationCaptionEntries(document, nativeSequenceName);
            TraceStage("caption-inventory");
            if (captions.Count == 0)
                return true;

            // Use the dedicated batch field path: it plans every final ordinal and
            // heading prefix first, mutates caption fields from the end of the
            // document toward the start, and patches generated REF results locally.
            // This is specifically designed to avoid the repeated global field
            // maintenance performed by the ordinary incremental updater.
            if (!TryApplyEquationNumberFormatByFieldBatch(
                    document,
                    nativeSequenceName,
                    captions,
                    out updated,
                    nativeTargetsAlreadyPlanned: true))
                return false;
            TraceStage("field-batch");

            if (!TryRefreshHealthyEquationNumbersInPlace(document, out updated))
                return false;
            TraceStage("healthy-refresh");
            return true;
        }
        catch
        {
            updated = 0;
            return false;
        }
    }

    internal static int SetEquationNumberFormat(Document document, string formatId)
    {
        var acceptanceTiming = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
            "1",
            StringComparison.Ordinal)
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        SetEquationNumberFormatPreference(document, formatId);
        var fastPath = TryApplyEquationNumberFormatInPlace(document, out var updated);
        var result = fastPath ? updated : Reconcile(document);
        if (acceptanceTiming is not null)
        {
            acceptanceTiming.Stop();
            Console.WriteLine(
                $"    [perf] SetEquationNumberFormat.{(fastPath ? "fast" : "reconcile")}: {result} formulas in {acceptanceTiming.ElapsedMilliseconds}ms");
        }
        return result;
    }

    private static bool TryApplyEquationNumberFormatInPlace(
        Document document,
        out int updated)
    {
        updated = 0;
        var acceptanceTiming = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
            "1",
            StringComparison.Ordinal)
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        var checkpoint = 0L;
        void TraceStage(string stage)
        {
            if (acceptanceTiming is null) return;
            var elapsed = acceptanceTiming.ElapsedMilliseconds;
            Console.WriteLine($"      [perf] number-format.{stage}: +{elapsed - checkpoint}ms ({elapsed}ms)");
            checkpoint = elapsed;
        }
        var nativeSequenceName = GetNativeEquationSequenceName(document);
        var captions = GetNativeEquationCaptionEntries(document, nativeSequenceName);
        var activeFormulaIds = captions
            .Select(item => item.FormulaId)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (activeFormulaIds.Count != captions.Count)
            return false;
        TraceStage("caption-inventory");

        if (TryApplyEquationNumberFormatByFieldBatch(
                document,
                nativeSequenceName,
                captions,
                out updated))
        {
            TraceStage("field-batch");
            return true;
        }

        // Changing only the display format must not rebuild otherwise healthy
        // three-column equation structures. Validate the lightweight numbering
        // bookmarks first; malformed/legacy documents still fall back to the
        // full reconciliation path below.
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            var artifactFormulaIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmark);
                bookmark = bookmarks[index];
                if (TryFormulaIdFromBookmark(
                        bookmark.Name,
                        EquationBookmarkPrefix,
                        out var formulaId)
                    || TryFormulaIdFromBookmark(
                        bookmark.Name,
                        NativeCaptionBookmarkPrefix,
                        out formulaId)
                    || TryFormulaIdFromBookmark(
                        bookmark.Name,
                        NativeNumberBookmarkPrefix,
                        out formulaId))
                {
                    artifactFormulaIds.Add(formulaId);
                }
            }
            if (!artifactFormulaIds.SetEquals(activeFormulaIds))
                return false;

            foreach (var formulaId in activeFormulaIds)
            {
                if (!HasCompleteFormulaNumberingArtifacts(document, formulaId))
                    return false;
                var visibleName = EquationBookmarkName(formulaId);
                if (!bookmarks.Exists(visibleName)) return false;
                Release(bookmark);
                bookmark = bookmarks[visibleName];
                Release(range);
                range = bookmark.Range;
                if (!IsNumberedEquationTable(range))
                    return false;
            }
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }

        TraceStage("validate-artifacts");
        if (activeFormulaIds.Count == 0)
            return true;

        // The hidden SEQ captions carry the chapter/section prefix and local
        // ordinal. Updating them in document order is sufficient for a format
        // switch; every table, formula object, metadata record and bookmark
        // scaffold already exists and must remain untouched.
        UpdateNativeEquationSequenceFields(
            document,
            nativeSequenceName,
            captions,
            formatOnly: true);
        TraceStage("update-sequence");
        // Refresh only REF fields. This updates both the visible number cells and
        // user cross-references without the expensive document-wide Fields.Update
        // plus structural OLE/OMML scans performed by Reconcile().
        UpdateNativeCrossReferences(document);
        TraceStage("update-references");
        updated = activeFormulaIds.Count;
        return true;
    }

    private static bool TryApplyEquationNumberFormatByFieldBatch(
        Document document,
        string nativeSequenceName,
        IReadOnlyList<NativeEquationCaptionEntry> captions,
        out int updated,
        bool nativeTargetsAlreadyPlanned = false)
    {
        updated = 0;
        if (captions.Count == 0) return true;
        var tracePerformance = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal)
            || string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                "1",
                StringComparison.Ordinal);
        var batchWatch = tracePerformance
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        long batchCheckpoint = 0;
        void TraceBatch(string stage)
        {
            if (batchWatch is null) return;
            var elapsed = batchWatch.ElapsedMilliseconds;
            TraceNumberingPerformance(
                $"[perf] number-format-batch.{stage}: +{elapsed - batchCheckpoint}ms ({elapsed}ms)");
            batchCheckpoint = elapsed;
        }

        var format = ReadEquationNumberFormat(document);
        var headingAnchors = format.UsesHeading
            ? GetHeadingNumberAnchorsForFormatBatch(
                document,
                format.HeadingLevel,
                captions)
            : Array.Empty<HeadingNumberAnchor>();
        var ordinalByScope = new Dictionary<int, int>();
        var plan = new List<(
            string FormulaId,
            int Ordinal,
            string Prefix,
            string ExpectedNumber)>(captions.Count);
        foreach (var caption in captions)
        {
            var scope = ResolveEquationNumberScope(
                caption.Position,
                format,
                headingAnchors);
            ordinalByScope.TryGetValue(scope.ScopePosition, out var localOrdinal);
            localOrdinal++;
            ordinalByScope[scope.ScopePosition] = localOrdinal;
            plan.Add((
                caption.FormulaId,
                localOrdinal,
                scope.Prefix,
                scope.Prefix + localOrdinal.ToString(
                    System.Globalization.CultureInfo.InvariantCulture)));
        }
        TraceBatch("plan");

        Bookmarks? bookmarks = null;
        var mutationStarted = false;
        try
        {
            bookmarks = document.Bookmarks;
            mutationStarted = true;

            // Resolve each native caption through its durable FormulaId. This
            // avoids enumerating the document's OLE EMBED and REF fields entirely.
            // Work backwards so prefix length changes cannot disturb captions that
            // have not yet been touched.
            long nativeLookupMs = 0;
            long nativeCodeMs = 0;
            long nativeUpdateMs = 0;
            long nativePrefixMs = 0;
            long nativeBookmarkMs = 0;
            for (var planIndex = plan.Count - 1; planIndex >= 0; planIndex--)
            {
                var item = plan[planIndex];
                Bookmark? captionBookmark = null;
                Bookmark? numberBookmark = null;
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
                    var captionName = NativeCaptionBookmarkName(item.FormulaId);
                    var numberName = NativeNumberBookmarkName(item.FormulaId);
                    if (!bookmarks.Exists(captionName) || !bookmarks.Exists(numberName))
                        throw new InvalidDataException(
                            "A VisualTeX equation-number bookmark is missing.");
                    var nativeStageWatch = batchWatch is null
                        ? null
                        : System.Diagnostics.Stopwatch.StartNew();
                    captionBookmark = bookmarks[captionName];
                    numberBookmark = bookmarks[numberName];
                    captionRange = captionBookmark.Range;
                    field = FindNativeEquationFieldInRange(captionRange, nativeSequenceName);
                    if (field is null)
                        throw new InvalidDataException(
                            "A VisualTeX equation caption lost its native SEQ field.");
                    if (nativeStageWatch is not null)
                    {
                        nativeLookupMs += nativeStageWatch.ElapsedMilliseconds;
                        nativeStageWatch.Restart();
                    }

                    code = field.Code;
                    if (!nativeTargetsAlreadyPlanned)
                        code.Text = $" SEQ {nativeSequenceName} \\r {item.Ordinal} \\* ARABIC ";
                    if (nativeStageWatch is not null)
                    {
                        nativeCodeMs += nativeStageWatch.ElapsedMilliseconds;
                        nativeStageWatch.Restart();
                    }
                    field.Update();
                    if (nativeStageWatch is not null)
                    {
                        nativeUpdateMs += nativeStageWatch.ElapsedMilliseconds;
                        nativeStageWatch.Restart();
                    }
                    fieldResult = field.Result;
                    paragraphs = fieldResult.Paragraphs;
                    if (paragraphs.Count != 1)
                        throw new InvalidDataException(
                            "A VisualTeX equation caption no longer occupies one paragraph.");
                    paragraph = paragraphs[1];
                    paragraphRange = paragraph.Range;

                    numberBookmark.Delete();
                    Release(numberBookmark);
                    numberBookmark = null;
                    if (!nativeTargetsAlreadyPlanned)
                    {
                        Release(code);
                        code = field.Code;
                        var fieldStart = Math.Max(paragraphRange.Start, code.Start - 1);
                        prefixRange = document.Range(paragraphRange.Start, fieldStart);
                        prefixRange.Text = item.Prefix;
                    }
                    if (nativeStageWatch is not null)
                    {
                        nativePrefixMs += nativeStageWatch.ElapsedMilliseconds;
                        nativeStageWatch.Restart();
                    }

                    Release(fieldResult);
                    fieldResult = field.Result;
                    refreshedNumberRange = document.Range(
                        paragraphRange.Start,
                        fieldResult.End);
                    numberBookmark = bookmarks.Add(numberName, refreshedNumberRange);
                    if (nativeStageWatch is not null)
                        nativeBookmarkMs += nativeStageWatch.ElapsedMilliseconds;
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
                    Release(numberBookmark);
                    Release(captionBookmark);
                }
            }
            if (batchWatch is not null)
            {
                TraceNumberingPerformance(
                    $"[perf] number-format-batch.native-breakdown lookup={nativeLookupMs}ms code={nativeCodeMs}ms update={nativeUpdateMs}ms prefix={nativePrefixMs}ms bookmark={nativeBookmarkMs}ms");
            }
            TraceBatch("native-targets");

            // Generated right-side numbers have their own durable bookmarks, so
            // patch those REF results locally instead of scanning document.Fields.
            var generatedReferencesPatched = true;
            foreach (var item in plan)
            {
                if (TryPatchVisibleEquationNumberResult(
                        bookmarks,
                        item.FormulaId,
                        item.ExpectedNumber))
                    continue;
                generatedReferencesPatched = false;
                break;
            }
            TraceBatch("visible-references");

            // If the document contains only the one generated REF per numbered
            // formula, no global reference pass is needed. An exact OpenXML count
            // is a safe fast check: any ambiguity or extra user reference falls
            // back to the normal comprehensive REF updater.
            if (!generatedReferencesPatched
                || !HasOnlyGeneratedEquationReferences(document, captions.Count))
                UpdateNativeCrossReferences(document);
            TraceBatch("extra-references");

            WriteNativeEquationTailFormulaId(document, captions.LastOrDefault()?.FormulaId);
            updated = captions.Count;
            return true;
        }
        catch
        {
            if (!mutationStarted)
            {
                updated = 0;
                return false;
            }
            try
            {
                updated = Reconcile(document);
                return true;
            }
            catch
            {
                updated = 0;
                return false;
            }
        }
        finally { Release(bookmarks); }
    }

    private static bool TryReadHealthyEquationNumberArtifactsFromOpenXml(
        Document document,
        out IReadOnlyList<NativeEquationCaptionEntry> entries,
        out IReadOnlyDictionary<string, int> referenceCounts)
    {
        entries = Array.Empty<NativeEquationCaptionEntry>();
        referenceCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        Range? content = null;
        try
        {
            content = document.Content;
            var xml = content.WordOpenXML ?? string.Empty;
            if (string.IsNullOrWhiteSpace(xml)) return false;

            var bookmarkMatches = Regex.Matches(
                xml,
                @"<w:bookmarkStart\b[^>]*\bw:name=""(?<name>VTEq(?:Cap|Num)?_[^""]+)""[^>]*/?>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (bookmarkMatches.Count == 0) return false;

            var visibleIds = new List<string>();
            var captionIds = new List<string>();
            var numberIds = new List<string>();
            foreach (Match match in bookmarkMatches)
            {
                var name = match.Groups["name"].Value;
                if (TryFormulaIdFromBookmark(
                        name,
                        NativeCaptionBookmarkPrefix,
                        out var formulaId))
                {
                    captionIds.Add(formulaId);
                    continue;
                }
                if (TryFormulaIdFromBookmark(
                        name,
                        NativeNumberBookmarkPrefix,
                        out formulaId))
                {
                    numberIds.Add(formulaId);
                    continue;
                }
                if (TryFormulaIdFromBookmark(
                        name,
                        EquationBookmarkPrefix,
                        out formulaId))
                {
                    visibleIds.Add(formulaId);
                    continue;
                }
                return false;
            }

            if (visibleIds.Count == 0
                || visibleIds.Count != captionIds.Count
                || visibleIds.Count != numberIds.Count)
                return false;

            var visibleSet = visibleIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var captionSet = captionIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var numberSet = numberIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (visibleSet.Count != visibleIds.Count
                || captionSet.Count != captionIds.Count
                || numberSet.Count != numberIds.Count
                || !visibleSet.SetEquals(captionSet)
                || !visibleSet.SetEquals(numberSet))
                return false;

            // A complete numbering scaffold whose three-column table no longer
            // contains a VisualTeX formula is an orphan (for example after a user
            // manually deletes only the OLE/OMML object). OLE formulas in older
            // documents do not necessarily carry VTO_* identity bookmarks, so
            // validate the actual table payload instead of requiring VTO_*.
            var visibleStartMatches = Regex.Matches(
                xml,
                @"<w:bookmarkStart\b(?=[^>]*\bw:name=""VTEq_(?<guid>[0-9A-F]{32})"")[^>]*/>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (visibleStartMatches.Count != visibleSet.Count) return false;
            foreach (Match visibleStart in visibleStartMatches)
            {
                var formulaId = Guid.ParseExact(
                    visibleStart.Groups["guid"].Value,
                    "N").ToString("D");
                var tableStart = xml.LastIndexOf(
                    "<w:tbl",
                    visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                var tableEnd = xml.IndexOf(
                    "</w:tbl>",
                    visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                if (tableStart < 0 || tableEnd <= visibleStart.Index) return false;
                tableEnd += "</w:tbl>".Length;
                if (tableEnd - tableStart > 262144) return false;
                var tableXml = xml.Substring(tableStart, tableEnd - tableStart);
                var hasOleFormula = tableXml.IndexOf(
                    "ProgID=\"VisualTeX.Formula.1\"",
                    StringComparison.OrdinalIgnoreCase) >= 0;
                var hasOmmlFormula = tableXml.IndexOf(
                    $"w:name=\"{WordOmmlFormulaStore.BookmarkName(formulaId)}\"",
                    StringComparison.OrdinalIgnoreCase) >= 0;
                if (!hasOleFormula && !hasOmmlFormula) return false;
            }

            var startMatches = Regex.Matches(
                xml,
                @"<w:bookmarkStart\b(?=[^>]*\bw:id=""(?<id>-?\d+)"")(?=[^>]*\bw:name=""VTEqNum_(?<guid>[0-9A-F]{32})"")[^>]*/>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (startMatches.Count != numberSet.Count) return false;

            var result = new List<NativeEquationCaptionEntry>(startMatches.Count);
            var seenNumbers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match startMatch in startMatches)
            {
                if (!Guid.TryParseExact(
                        startMatch.Groups["guid"].Value,
                        "N",
                        out var formulaGuid))
                    return false;
                var formulaId = formulaGuid.ToString("D");
                if (!numberSet.Contains(formulaId) || !seenNumbers.Add(formulaId))
                    return false;

                var bookmarkId = Regex.Escape(startMatch.Groups["id"].Value);
                var endMatch = Regex.Match(
                    xml,
                    $@"<w:bookmarkEnd\b[^>]*\bw:id=""{bookmarkId}""[^>]*/>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                    TimeSpan.FromSeconds(1));
                if (!endMatch.Success || endMatch.Index <= startMatch.Index)
                    return false;
                if (endMatch.Index - startMatch.Index > 16384)
                    return false;

                var segment = xml.Substring(
                    startMatch.Index + startMatch.Length,
                    endMatch.Index - startMatch.Index - startMatch.Length);
                var textMatches = Regex.Matches(
                    segment,
                    @"<w:t(?:\s[^>]*)?>(?<text>.*?)</w:t>",
                    RegexOptions.IgnoreCase
                    | RegexOptions.CultureInvariant
                    | RegexOptions.Singleline);
                var numberText = NormalizeNativeEquationNumberText(string.Concat(
                    textMatches.Cast<Match>()
                        .Select(match => System.Net.WebUtility.HtmlDecode(
                            match.Groups["text"].Value))));
                if (string.IsNullOrWhiteSpace(numberText)) return false;
                result.Add(new NativeEquationCaptionEntry(
                    formulaId,
                    startMatch.Index,
                    numberText));
            }

            if (result.Count != visibleSet.Count) return false;

            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var referenceMatches = Regex.Matches(
                xml,
                @"\bREF\s+VTEqNum_(?<guid>[0-9A-F]{32})\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            foreach (Match referenceMatch in referenceMatches)
            {
                if (!Guid.TryParseExact(
                        referenceMatch.Groups["guid"].Value,
                        "N",
                        out var referenceGuid))
                    continue;
                var formulaId = referenceGuid.ToString("D");
                if (!visibleSet.Contains(formulaId)) continue;
                counts.TryGetValue(formulaId, out var currentCount);
                counts[formulaId] = currentCount + 1;
            }
            // Every healthy numbered formula owns one generated REF in its right
            // number cell. Missing that REF is a structural anomaly that belongs
            // on the conservative full reconciliation path.
            if (visibleSet.Any(formulaId =>
                    !counts.TryGetValue(formulaId, out var count) || count < 1))
                return false;

            entries = result.OrderBy(entry => entry.Position).ToArray();
            referenceCounts = counts;
            return true;
        }
        catch
        {
            entries = Array.Empty<NativeEquationCaptionEntry>();
            referenceCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            return false;
        }
        finally { Release(content); }
    }

    internal static bool TryGetHealthyEquationReferenceCounts(
        Document document,
        out IReadOnlyDictionary<string, int> referenceCounts)
    {
        // Reuse the same one-pass WordOpenXML validation used by the fast
        // numbering updater. A healthy numbered formula has exactly one generated
        // REF in its own number cell; counts > 1 mean genuine external references
        // also exist and must be frozen before that VisualTeX number is removed.
        return TryReadHealthyEquationNumberArtifactsFromOpenXml(
            document,
            out _,
            out referenceCounts);
    }

    private static IReadOnlyList<NativeEquationCaptionEntry> ResolveNativeEquationCaptionPositions(
        Document document,
        IReadOnlyList<NativeEquationCaptionEntry> openXmlEntries)
    {
        var result = new List<NativeEquationCaptionEntry>(openXmlEntries.Count);
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var entry in openXmlEntries)
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    bookmark = bookmarks[NativeNumberBookmarkName(entry.FormulaId)];
                    range = bookmark.Range;
                    result.Add(new NativeEquationCaptionEntry(
                        entry.FormulaId,
                        range.Start,
                        entry.NumberText));
                }
                catch
                {
                    return Array.Empty<NativeEquationCaptionEntry>();
                }
                finally
                {
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }
        return result.OrderBy(item => item.Position).ToArray();
    }

    private static bool HasOnlyGeneratedEquationReferences(
        Document document,
        int numberedFormulaCount)
    {
        Range? content = null;
        try
        {
            content = document.Content;
            var xml = content.WordOpenXML ?? string.Empty;
            var referenceCount = Regex.Matches(
                xml,
                @"\bREF\s+VTEqNum_[0-9A-F]{32}\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                .Count;
            return referenceCount == numberedFormulaCount;
        }
        catch
        {
            return false;
        }
        finally { Release(content); }
    }

    internal static void SetEquationNumberFormatPreference(
        Document document,
        string formatId)
    {
        var format = EquationNumberFormat.Resolve(formatId);
        SetDefaultEquationNumberFormatPreference(format.Id);
        WriteEquationNumberFormat(document, format.Id);
    }

    private static EquationNumberFormat ReadEquationNumberFormat(Document document)
    {
        return TryReadDocumentEquationNumberFormatId(document, out var formatId)
            ? EquationNumberFormat.Resolve(formatId)
            : EquationNumberFormat.Resolve(GetDefaultEquationNumberFormatId());
    }

    private static bool TryReadDocumentEquationNumberFormatId(
        Document document,
        out string formatId)
    {
        formatId = EquationNumberFormat.ContinuousId;
        Variables? variables = null;
        Variable? variable = null;
        try
        {
            variables = document.Variables;
            object index = EquationNumberFormatVariableName;
            try
            {
                variable = variables.get_Item(ref index);
                formatId = EquationNumberFormat.Resolve(variable.Value).Id;
                return true;
            }
            catch (COMException)
            {
                return false;
            }
        }
        finally
        {
            Release(variable);
            Release(variables);
        }
    }

    private static void EnsureDocumentEquationNumberFormatPreference(Document document)
    {
        if (TryReadDocumentEquationNumberFormatId(document, out _)) return;
        WriteEquationNumberFormat(document, GetDefaultEquationNumberFormatId());
    }

    private static RegistryView[] UserPreferenceRegistryViews() =>
        Environment.Is64BitOperatingSystem
            ? new[] { RegistryView.Registry64, RegistryView.Registry32 }
            : new[] { RegistryView.Default };

    private static object? ReadUserPreference(string name)
    {
        foreach (var view in UserPreferenceRegistryViews())
        {
            try
            {
                using var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
                using var key = root.OpenSubKey(UserPreferenceRegistryPath, writable: false);
                var value = key?.GetValue(name);
                if (value is not null) return value;
            }
            catch
            {
                // Preferences are optional. A locked or unavailable registry
                // must never block formula insertion or numbering.
            }
        }
        return null;
    }

    private static void WriteUserPreference(
        string name,
        object value,
        RegistryValueKind valueKind)
    {
        foreach (var view in UserPreferenceRegistryViews())
        {
            try
            {
                using var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
                using var key = root.CreateSubKey(UserPreferenceRegistryPath, writable: true);
                key?.SetValue(name, value, valueKind);
            }
            catch
            {
                // Keep document operations functional even if persistence is
                // unavailable on a restricted machine.
            }
        }
    }

    private static void WriteEquationNumberFormat(Document document, string formatId)
    {
        Variables? variables = null;
        Variable? variable = null;
        try
        {
            variables = document.Variables;
            object index = EquationNumberFormatVariableName;
            try
            {
                variable = variables.get_Item(ref index);
                variable.Value = formatId;
            }
            catch (COMException)
            {
                object value = formatId;
                variable = variables.Add(EquationNumberFormatVariableName, ref value);
            }
        }
        finally
        {
            Release(variable);
            Release(variables);
        }
    }

    private static bool TryReadNativeEquationTailFormulaId(
        Document document,
        out string formulaId)
    {
        formulaId = string.Empty;
        Variables? variables = null;
        Variable? variable = null;
        try
        {
            variables = document.Variables;
            object index = EquationNumberTailFormulaVariableName;
            try
            {
                variable = variables.get_Item(ref index);
                formulaId = (variable.Value ?? string.Empty).Trim();
                return !string.IsNullOrWhiteSpace(formulaId);
            }
            catch (COMException)
            {
                return false;
            }
        }
        finally
        {
            Release(variable);
            Release(variables);
        }
    }

    private static void WriteNativeEquationTailFormulaId(
        Document document,
        string? formulaId)
    {
        Variables? variables = null;
        Variable? variable = null;
        try
        {
            variables = document.Variables;
            object index = EquationNumberTailFormulaVariableName;
            try
            {
                variable = variables.get_Item(ref index);
                if (string.IsNullOrWhiteSpace(formulaId))
                {
                    variable.Delete();
                    return;
                }
                variable.Value = formulaId;
            }
            catch (COMException)
            {
                if (string.IsNullOrWhiteSpace(formulaId)) return;
                object value = formulaId;
                variable = variables.Add(EquationNumberTailFormulaVariableName, ref value);
            }
        }
        finally
        {
            Release(variable);
            Release(variables);
        }
    }

    public static void TryReconcile(Document document)
    {
        try { Reconcile(document); }
        catch
        {
            // Formula insertion/update is already durable. The user can retry
            // only the numbering command without duplicating or losing it.
        }
    }

    public static void TryReconcileFormula(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        FormulaMetadata metadata,
        bool numberingOrderMayHaveChanged = true,
        bool reuseExistingNumberedTableFormatting = false,
        Table? knownNumberedTable = null)
    {
        try
        {
            ReconcileFormula(
                document,
                formulaRange,
                formulaHeightPoints,
                metadata,
                numberingOrderMayHaveChanged,
                reuseExistingNumberedTableFormatting,
                knownNumberedTable);
        }
        catch
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                throw;
            // The inserted or edited formula is already durable. The explicit
            // update-number command still performs a complete reconciliation.
        }
    }

    internal static void BuildFormulaNumberingScaffoldForConversion(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        FormulaMetadata metadata,
        Table? knownNumberedTable = null,
        int? plannedOrdinal = null,
        string? plannedPrefix = null)
    {
        if (!metadata.Numbered
            || !string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
            return;
        EnsureDocumentEquationNumberFormatPreference(document);
        var formulaFontSizePoints = (float)FormulaFontSize.ResolveSemanticFontSize(metadata);
        if (knownNumberedTable is not null)
        {
            var traceFreshScaffold = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                "1",
                StringComparison.Ordinal);
            var freshWatch = traceFreshScaffold
                ? System.Diagnostics.Stopwatch.StartNew()
                : null;
            long freshCheckpoint = 0;
            void TraceFresh(string stage)
            {
                if (freshWatch is null) return;
                var elapsed = freshWatch.ElapsedMilliseconds;
                TraceNumberingPerformance(
                    $"[perf] ConversionScaffold.{stage}: +{elapsed - freshCheckpoint}ms ({elapsed}ms)");
                freshCheckpoint = elapsed;
            }

            // Format conversion passes a freshly-created 1x3 table and a fresh
            // FormulaId. None of the generic idempotent probes can succeed here;
            // they only enumerate existing bookmarks/fields/tables as the document
            // grows. Build the known-new local scaffold directly, then let the
            // batch finalizer perform the one sequence/reference refresh and the
            // strict document-wide health validation.
            EnsureNumberedOmmlIsDisplay(formulaRange);
            TraceFresh("ensure-display");
            var sequenceName = GetNativeEquationSequenceName(document);
            TraceFresh("sequence-name");
            CreateNativeCaption(
                document,
                formulaRange,
                metadata.FormulaId,
                sequenceName,
                knownNumberedTable,
                deferFieldUpdate: true,
                plannedOrdinal: plannedOrdinal,
                plannedPrefix: plannedPrefix);
            TraceFresh("native-caption");
            InsertVisibleEquationNumber(
                document,
                formulaRange,
                formulaHeightPoints,
                formulaFontSizePoints,
                metadata.FormulaId,
                NativeNumberBookmarkName(metadata.FormulaId),
                knownNumberedTable,
                useConversionSafeVisibleNumber: true,
                deferFieldUpdate: true);
            TraceFresh("visible-ref");
            return;
        }

        ConfigureNumberedDisplayFormula(
            document,
            formulaRange,
            formulaHeightPoints,
            formulaFontSizePoints,
            metadata.FormulaId,
            reuseExistingScaffold: false,
            knownNumberedTable: null,
            useConversionSafeVisibleNumber: true);
    }

    internal static bool TryBuildConvertedOmmlNumberingBatch(
        Document document,
        IReadOnlyList<FormulaMetadata> metadataItems,
        out int built)
    {
        built = 0;
        var entries = new List<(string FormulaId, FormulaMetadata Metadata, int Position)>();
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
            TraceNumberingPerformance(
                $"[perf] ConvertedOmmlBatch.{stage}: +{elapsed - checkpoint}ms ({elapsed}ms)");
            checkpoint = elapsed;
        }
        try
        {
            foreach (var metadata in metadataItems
                         .Where(item => item is not null
                             && item.Numbered
                             && string.Equals(item.DisplayMode, "block", StringComparison.Ordinal)
                             && !string.IsNullOrWhiteSpace(item.FormulaId))
                         .GroupBy(item => item.FormulaId, StringComparer.OrdinalIgnoreCase)
                         .Select(group => group.Last()))
            {
                Range? range = null;
                try
                {
                    range = GetFreshConvertedOmmlRange(document, metadata);
                    entries.Add((metadata.FormulaId, metadata, range.Start));
                }
                finally { Release(range); }
            }

            TraceStage("inventory");

            // Plan the final number for every converted formula before any table,
            // caption or REF structure is inserted. This is the same old fast-path
            // principle used for large numbered documents: discover heading scope
            // once, then carry known local ordinals through the mutation phase.
            var orderedEntries = entries.OrderBy(item => item.Position).ToArray();
            var format = ReadEquationNumberFormat(document);
            var syntheticCaptions = orderedEntries
                .Select(item => new NativeEquationCaptionEntry(
                    item.FormulaId,
                    item.Position,
                    string.Empty))
                .ToArray();
            var headingAnchors = format.UsesHeading
                ? GetHeadingNumberAnchorsForFormatBatch(
                    document,
                    format.HeadingLevel,
                    syntheticCaptions)
                : Array.Empty<HeadingNumberAnchor>();
            var ordinalByScope = new Dictionary<int, int>();
            var plannedNumbers = new Dictionary<string, (int Ordinal, string Prefix)>(
                StringComparer.OrdinalIgnoreCase);
            foreach (var item in orderedEntries)
            {
                var scope = ResolveEquationNumberScope(
                    item.Position,
                    format,
                    headingAnchors);
                ordinalByScope.TryGetValue(scope.ScopePosition, out var ordinal);
                ordinal++;
                ordinalByScope[scope.ScopePosition] = ordinal;
                plannedNumbers[item.FormulaId] = (ordinal, scope.Prefix);
            }
            TraceStage("number-plan");

            // Structural table migration is deliberately end-to-start. This is
            // the same safety rule used by the long-standing reconcile path: a
            // table inserted for a later formula cannot drag the bookmark/range
            // of an earlier formula that has not yet been processed.
            foreach (var entry in entries.OrderByDescending(item => item.Position))
            {
                Range? range = null;
                Tables? tables = null;
                Table? table = null;
                try
                {
                    var metadata = entry.Metadata;
                    range = GetFreshConvertedOmmlRange(document, metadata);
                    TraceStage("locate-target");

                    EnsureNumberedOmmlIsDisplay(range);
                    EnsureStandardNumberedEquationTable(
                        document,
                        range,
                        entry.FormulaId);
                    TraceStage("ensure-table");
                    ConfigureEquationParagraph(range, numbered: false);
                    TraceStage("format-paragraph");
                    ConfigureFreshConversionNumberedEquationTable(range);
                    TraceStage("format-table");

                    tables = range.Tables;
                    if (tables.Count == 0)
                        throw new InvalidOperationException(
                            $"Converted OMML formula {entry.FormulaId} did not enter its numbered table.");
                    table = tables[1];
                    var plannedNumber = plannedNumbers[entry.FormulaId];
                    BuildFormulaNumberingScaffoldForConversion(
                        document,
                        range,
                        WordOmmlFormulaStore.EstimateHeightPoints(range),
                        metadata,
                        table,
                        plannedNumber.Ordinal,
                        plannedNumber.Prefix);
                    TraceStage("scaffold");
                    TrimBenignEmptyRowsFromNumberedTable(
                        document,
                        range,
                        entry.FormulaId);
                    TraceStage("trim");
                    built++;
                }
                finally
                {
                    Release(table);
                    Release(tables);
                    Release(range);
                }
            }
            return true;
        }
        catch (Exception error)
        {
            TraceNumberingPerformance(
                $"[perf] converted-omml-numbering-batch.fallback built={built}/{entries.Count} error={error.Message}");
            built = 0;
            return false;
        }
    }

    private static Range GetFreshConvertedOmmlRange(
        Document document,
        FormulaMetadata metadata)
    {
        Bookmark? bookmark = null;
        Range? localRange = null;
        try
        {
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, metadata.FormulaId);
            if (bookmark is not null)
            {
                try
                {
                    // These bookmarks were created moments earlier by the same
                    // conversion transaction. Resolve the adjacent OMath locally
                    // first; do not re-enumerate document.OMaths or revalidate the
                    // CustomXMLPart for every formula in the batch.
                    localRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                    if (WordOmmlFormulaStore.IsCanonicalAnchor(bookmark, localRange))
                    {
                        var result = localRange;
                        localRange = null;
                        return result;
                    }
                }
                catch
                {
                    // Fall through to the durable fingerprint recovery path below.
                }
            }
        }
        finally
        {
            Release(localRange);
            Release(bookmark);
        }

        // A non-canonical fresh bookmark is unexpected, but correctness still
        // wins over speed. Re-enter the mature metadata/fingerprint recovery path
        // only for that exceptional formula.
        var persistedMetadata = WordOmmlFormulaStore.TryRead(
            document,
            metadata.FormulaId) ?? metadata;
        return WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
            document,
            metadata.FormulaId,
            persistedMetadata);
    }

    internal static void ReconcileFormula(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        FormulaMetadata metadata,
        bool numberingOrderMayHaveChanged = true,
        bool reuseExistingNumberedTableFormatting = false,
        Table? knownNumberedTable = null)
    {
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        var performanceWatch = tracePerformance
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        var performanceCheckpoint = 0L;
        void TraceStage(string stage)
        {
            if (performanceWatch is null) return;
            var elapsed = performanceWatch.ElapsedMilliseconds;
            Console.WriteLine(
                $"      [perf] ReconcileFormula.{stage}: +{elapsed - performanceCheckpoint}ms ({elapsed}ms)");
            performanceCheckpoint = elapsed;
        }

        var hadNumberingArtifacts = HasFormulaNumberingArtifacts(
            document,
            metadata.FormulaId);
        TraceStage("artifact-probe");
        if (metadata.DisplayMode != "block")
        {
            if (!hadNumberingArtifacts) return;
            RemoveFormulaNumberingArtifacts(document, metadata.FormulaId);
            UpdateMainStoryFields(document);
            UpdateNativeCrossReferences(document);
            return;
        }

        var formulaFontSizePoints = (float)FormulaFontSize.ResolveSemanticFontSize(metadata);
        if (metadata.Numbered)
        {
            var hadCompleteOwnedArtifacts =
                HasCompleteFormulaNumberingArtifacts(document, metadata.FormulaId)
                && FormulaRangeOwnsNumberingArtifacts(
                    document,
                    formulaRange,
                    metadata.FormulaId);
            TraceStage("owned-artifacts");

            EnsureDocumentEquationNumberFormatPreference(document);
            TraceStage("format-preference");
            var rebuiltStableNumberingSlot =
                !hadCompleteOwnedArtifacts
                && knownNumberedTable is not null
                && reuseExistingNumberedTableFormatting
                && !numberingOrderMayHaveChanged;
            var visibleNumberCreated = ConfigureNumberedDisplayFormula(
                document,
                formulaRange,
                formulaHeightPoints,
                formulaFontSizePoints,
                metadata.FormulaId,
                reuseExistingScaffold:
                    knownNumberedTable is not null
                    || ((hadCompleteOwnedArtifacts || reuseExistingNumberedTableFormatting)
                        && IsNumberedEquationTable(formulaRange)),
                knownNumberedTable: knownNumberedTable);
            TraceStage("configure-scaffold");

            // OLE -> OMML replacement deliberately removes this formula's three
            // numbering bookmarks before inserting the adjacent native equation,
            // then rebuilds them in the same known table. The formula did not move,
            // so the newly updated local SEQ/REF is already the correct ordinal;
            // treating the missing pre-rebuild bookmarks as an insertion would
            // renumber the entire suffix of a large document for no reason.
            if (rebuiltStableNumberingSlot)
                TraceStage("stable-slot-rebuild");

            // Editing a healthy numbered formula does not change its document
            // position or sequence ordinal. The previous path still enumerated
            // every SEQ/REF/field in the document for every Apply, which made a
            // six-formula document visibly block Word for one or two seconds.
            // Only insert/copy/numbering-structure changes need an order refresh.
            var currentVisibleNumberRefreshed = false;
            if (!rebuiltStableNumberingSlot
                && (numberingOrderMayHaveChanged || !hadCompleteOwnedArtifacts))
            {
                if (TryUpdateAppendedNativeEquationSequenceField(
                        document,
                        metadata.FormulaId,
                        out var appendedNumberChanged))
                {
                    TraceStage("append-sequence-fast");
                    // A newly appended FormulaId cannot have pre-existing body
                    // cross-references. If its heading-aware caption changed,
                    // refresh only this table's visible REF instead of scanning
                    // every Field in the document.
                    if (appendedNumberChanged)
                    {
                        UpdateEquationNumberFields(
                            document,
                            formulaHeightPoints,
                            formulaFontSizePoints,
                            metadata.FormulaId);
                        currentVisibleNumberRefreshed = true;
                    }
                }
                else if (TryUpdateInsertedContinuousEquationFieldRanges(
                             document,
                             metadata.FormulaId))
                {
                    TraceStage("continuous-range-fast");
                }
                else if (TryUpdateInsertedContinuousEquationSequenceSuffix(
                             document,
                             metadata.FormulaId,
                             out var changedFormulaNumbers,
                             out var referencesAlreadyUpdatedFrom))
                {
                    TraceStage("continuous-suffix-fast");
                    if (changedFormulaNumbers.Count > 0)
                    {
                        UpdateHealthyNativeCrossReferencesAfterRenumbering(
                            document,
                            changedFormulaNumbers,
                            referencesAlreadyUpdatedFrom >= 0
                                ? referencesAlreadyUpdatedFrom
                                : null);
                        TraceStage("continuous-suffix-references");
                    }
                }
                else
                {
                    var fallbackChangedFormulaNumbers =
                        UpdateNativeEquationSequenceFieldsIncremental(document);
                    TraceStage("sequence-fallback");
                    if (fallbackChangedFormulaNumbers.Count > 0)
                    {
                        UpdateHealthyNativeCrossReferencesAfterRenumbering(
                            document,
                            fallbackChangedFormulaNumbers);
                        TraceStage("cross-reference-fallback");
                    }
                }
            }

            // The current formula can change size/font even when its ordinal is
            // stable, so keep the local visible-number formatting synchronized.
            if (!visibleNumberCreated && !currentVisibleNumberRefreshed)
            {
                UpdateEquationNumberFields(
                    document,
                    formulaHeightPoints,
                    formulaFontSizePoints,
                    metadata.FormulaId);
            }
            TraceStage("visible-number");
            return;
        }

        if (!hadNumberingArtifacts)
        {
            // Ordinary unnumbered display formulas have no numbering artifacts
            // to remove. Configure only the local paragraph; scanning bookmarks,
            // fields and cross-references here made an unrelated 100-formula
            // document pay a document-wide cost for every single edit.
            ConfigureEquationParagraph(formulaRange, numbered: false);
            return;
        }

        ConfigureUnnumberedDisplayFormula(
            document,
            formulaRange,
            metadata.FormulaId);

        // Removing numbering changes the ordinals of later formulas. Keep the
        // conservative full field refresh for this less-common structural path.
        UpdateNativeEquationSequenceFields(document);
        UpdateMainStoryFields(document);
        UpdateNativeCrossReferences(document);
    }

    private static bool HasFormulaNumberingArtifacts(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            return bookmarks.Exists(EquationBookmarkName(formulaId))
                || bookmarks.Exists(NativeCaptionBookmarkName(formulaId))
                || bookmarks.Exists(NativeNumberBookmarkName(formulaId));
        }
        catch { return false; }
        finally { Release(bookmarks); }
    }

    internal static bool HasCompleteFormulaNumberingArtifacts(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            return bookmarks.Exists(EquationBookmarkName(formulaId))
                && bookmarks.Exists(NativeCaptionBookmarkName(formulaId))
                && bookmarks.Exists(NativeNumberBookmarkName(formulaId));
        }
        catch { return false; }
        finally { Release(bookmarks); }
    }

    internal static Table? FindNumberedEquationTable(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Tables? tables = null;
        Table? table = null;
        Columns? columns = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return null;
            bookmark = bookmarks[name];
            range = bookmark.Range;
            tables = range.Tables;
            if (tables.Count == 0) return null;
            table = tables[1];
            columns = table.Columns;
            if (columns.Count < 3) return null;
            var result = table;
            table = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(columns);
            Release(table);
            Release(tables);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static bool FormulaRangeOwnsNumberingArtifacts(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? numberBookmark = null;
        Range? numberRange = null;
        Tables? formulaTables = null;
        Tables? numberTables = null;
        Table? formulaTable = null;
        Table? numberTable = null;
        Range? formulaTableRange = null;
        Range? numberTableRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var numberName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(numberName)) return false;
            numberBookmark = bookmarks[numberName];
            numberRange = numberBookmark.Range;

            formulaTables = formulaRange.Tables;
            numberTables = numberRange.Tables;
            if (formulaTables.Count == 0 || numberTables.Count == 0)
                return formulaRange.Start <= numberRange.Start
                    && formulaRange.End >= numberRange.End;

            formulaTable = formulaTables[1];
            numberTable = numberTables[1];
            formulaTableRange = formulaTable.Range;
            numberTableRange = numberTable.Range;
            return formulaTableRange.Start == numberTableRange.Start;
        }
        catch { return false; }
        finally
        {
            Release(numberTableRange);
            Release(formulaTableRange);
            Release(numberTable);
            Release(formulaTable);
            Release(numberTables);
            Release(formulaTables);
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
        }
    }

    private static void UpdateMainStoryFields(Document document)
    {
        Fields? fields = null;
        try
        {
            fields = document.Fields;
            if (fields.Count > 0) fields.Update();
        }
        finally { Release(fields); }
    }

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
                splitRange.InsertParagraphBefore();
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
            catch (COMException)
            {
                // A malformed legacy bookmark should not make an otherwise valid
                // document unusable. The normal conversion safety checks will still
                // stop before deleting a source if this one could not be repaired.
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

    internal static Range? EnsureNormalTypingParagraphAfterNumberedDisplay(
        Document document,
        string formulaId)
    {
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
        Frames? typingFrames = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        ParagraphFormat? format = null;
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
                            try
                            {
                                object normalStyle = WdBuiltinStyle.wdStyleNormal;
                                typingRange.set_Style(ref normalStyle);
                            }
                            catch { }
                            font = typingRange.Font;
                            font.Reset();
                            font.Hidden = 0;
                            font.Position = 0;
                            font.Color = WdColor.wdColorAutomatic;
                            format = typingRange.ParagraphFormat;
                            format.Reset();
                            format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                            format.SpaceBefore = 0f;
                            format.SpaceAfter = 0f;
                            typingRange.Collapse(WdCollapseDirection.wdCollapseStart);
                            var existingResult = typingRange;
                            typingRange = null;
                            return existingResult;
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

            try
            {
                object normalStyle = WdBuiltinStyle.wdStyleNormal;
                typingRange.set_Style(ref normalStyle);
            }
            catch
            {
                // Documents with locked/custom style collections can reject the
                // built-in Normal style. Direct-format reset below still removes
                // the hidden caption's one-point formatting.
            }

            font = typingRange.Font;
            font.Reset();
            font.Hidden = 0;
            font.Position = 0;
            font.Color = WdColor.wdColorAutomatic;
            format = typingRange.ParagraphFormat;
            format.Reset();
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            typingRange.Collapse(WdCollapseDirection.wdCollapseStart);
            var result = typingRange;
            typingRange = null;
            return result;
        }
        finally
        {
            Release(format);
            Release(font);
            Release(typingFrames);
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
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Range? captionRange = null;
        Range? content = null;
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Frames? frames = null;
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
            // Keep that structural terminator, but collapse it to one point so it no
            // longer appears as a full blank line below the last numbered formula.
            CompactTrailingTypingParagraph(document, paragraphRange);
        }
        finally
        {
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
            catch { }
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
            var xml = content.WordOpenXML ?? string.Empty;
            if (string.IsNullOrWhiteSpace(xml)) return Array.Empty<string>();

            var planned = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var captionMatches = Regex.Matches(
                xml,
                @"<w:bookmarkStart\b(?=[^>]*\bw:name=""VTEqCap_(?<guid>[0-9A-F]{32})"")[^>]*/>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (captionMatches.Count == 0) return Array.Empty<string>();

            // The final Word paragraph cannot be deleted; probing just the final
            // numbered formula lets CleanupNumberedDisplayInsertionSpacing compact
            // that terminator to 1 pt when it directly follows the last caption.
            if (Guid.TryParseExact(
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
                    FormulaId = formulaId,
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

    private static void CompactTrailingTypingParagraph(
        Document document,
        Range paragraphRange)
    {
        Bookmarks? bookmarks = null;
        Bookmark? previousBookmark = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        ParagraphFormat? format = null;
        try
        {
            try
            {
                object normalStyle = WdBuiltinStyle.wdStyleNormal;
                paragraphRange.set_Style(ref normalStyle);
            }
            catch { }
            font = paragraphRange.Font;
            font.Reset();
            font.Size = CompactTypingTailFontSizePoints;
            font.Hidden = 0;
            font.Position = 0;
            font.Color = WdColor.wdColorAutomatic;
            format = paragraphRange.ParagraphFormat;
            format.Reset();
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            format.LineSpacing = CompactTypingTailLineSpacingPoints;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;

            bookmarks = document.Bookmarks;
            if (bookmarks.Exists(CompactTypingTailBookmarkName))
            {
                previousBookmark = bookmarks[CompactTypingTailBookmarkName];
                previousBookmark.Delete();
                Release(previousBookmark);
                previousBookmark = null;
            }
            previousBookmark = bookmarks.Add(
                CompactTypingTailBookmarkName,
                paragraphRange);
        }
        finally
        {
            Release(previousBookmark);
            Release(bookmarks);
            Release(format);
            Release(font);
        }
    }

    private static void RemoveEmptyTrailingNumberedTableRows(Table table)
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

    internal static void RemoveFormulaNumberingArtifacts(
        Document document,
        string formulaId,
        bool preserveNativeCaptionParagraph = false)
    {
        RemoveVisibleEquationNumber(document, formulaId);
        RemoveNativeCaption(
            document,
            formulaId,
            preserveParagraphSeparator: preserveNativeCaptionParagraph);
    }

    public static int Reconcile(Document document)
    {
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        var performanceWatch = tracePerformance
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        var performanceCheckpoint = 0L;
        void TraceReconcileStage(string stage)
        {
            if (performanceWatch is null) return;
            var elapsed = performanceWatch.ElapsedMilliseconds;
            TraceNumberingPerformance(
                $"[perf] Reconcile.{stage}: +{elapsed - performanceCheckpoint}ms ({elapsed}ms)");
            performanceCheckpoint = elapsed;
        }

        var numberedFormulaIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // Freeze and verify every display-OMML identity before the first table
        // insertion, then migrate from the end of the document toward the start.
        // Later insertions therefore cannot drag the collapsed bookmark of an
        // equation that has not yet been processed.
        var ommlFormulaIds = GetOmmlDisplayFormulaIdsForStructuralEdit(document);
        TraceReconcileStage("capture-omml-ids");
        InlineShapes? inlineShapes = null;
        try
        {
            inlineShapes = document.InlineShapes;
            var inlineCount = inlineShapes.Count;
            var numberedFormulaLocations = CaptureNumberedFormulaLocations(
                document,
                inlineShapes,
                ommlFormulaIds);
            TraceReconcileStage("capture-numbered-locations");
            RepairSharedNativeCaptionArtifacts(document, numberedFormulaLocations);
            TraceReconcileStage("repair-shared-captions");

            for (var index = 1; index <= inlineCount; index++)
            {
                InlineShape? shape = null;
                Range? formulaRange = null;
                try
                {
                    shape = inlineShapes[index];
                    var metadata = ReadMetadata(shape);
                    if (metadata is null || metadata.DisplayMode != "block") continue;
                    formulaRange = shape.Range;
                    if (metadata.Numbered)
                    {
                        ConfigureNumberedDisplayFormula(
                            document,
                            formulaRange,
                            shape.Height,
                            (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                            metadata.FormulaId);
                        numberedFormulaIds.Add(metadata.FormulaId);
                    }
                    else
                    {
                        ConfigureUnnumberedDisplayFormula(
                            document,
                            formulaRange,
                            metadata.FormulaId);
                    }
                }
                finally
                {
                    Release(formulaRange);
                    Release(shape);
                }
            }
            TraceReconcileStage("configure-inline-formulas");

            foreach (var formulaId in ommlFormulaIds)
            {
                Range? formulaRange = null;
                try
                {
                    var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                    if (metadata is null || metadata.DisplayMode != "block") continue;
                    formulaRange =
                        WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                            document,
                            formulaId,
                            metadata);
                    var formulaHeightPoints =
                        WordOmmlFormulaStore.EstimateHeightPoints(formulaRange);
                    if (metadata.Numbered)
                    {
                        ConfigureNumberedDisplayFormula(
                            document,
                            formulaRange,
                            formulaHeightPoints,
                            (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                            formulaId);
                        numberedFormulaIds.Add(formulaId);
                    }
                    else
                    {
                        ConfigureUnnumberedDisplayFormula(document, formulaRange, formulaId);
                    }
                }
                finally { Release(formulaRange); }
            }
            TraceReconcileStage("configure-omml-formulas");

            if (numberedFormulaIds.Count > 0)
                EnsureDocumentEquationNumberFormatPreference(document);
            TraceReconcileStage("number-format-preference");

            RemoveOrphanEquationArtifacts(document, numberedFormulaIds);
            TraceReconcileStage("remove-orphan-artifacts");
            RebuildNativeNumberBookmarksFromCaptions(
                document,
                numberedFormulaIds);
            TraceReconcileStage("rebuild-number-bookmarks-1");

            // Word caches SEQ results independently from REF results. After a
            // numbered formula is deleted, refresh every native Equation SEQ
            // field in document order before updating any visible or body REF
            // field. Otherwise a REF can continue displaying the removed
            // formula's old ordinal until Word performs a later global update.
            UpdateNativeEquationSequenceFields(document);
            TraceReconcileStage("update-native-sequences");
            // Writing chapter/section prefixes mutates the hidden caption
            // paragraphs. Word can invalidate a number bookmark whose range was
            // created before that mutation even though the SEQ field and caption
            // bookmark survive. Re-wrap every final SEQ result after all caption
            // text is stable so each formula keeps an independent REF target.
            RebuildNativeNumberBookmarksFromCaptions(
                document,
                numberedFormulaIds);
            TraceReconcileStage("rebuild-number-bookmarks-2");
            // Newly batch-numbered formulas create their visible REF fields
            // before the native number bookmarks are rewritten with the final
            // chapter/section prefix. Word can temporarily cache "reference
            // source not found" for those fresh fields. Refresh the complete
            // main story once after all target bookmarks are stable, then apply
            // formula-specific alignment below.
            UpdateMainStoryFields(document);
            TraceReconcileStage("update-main-story-fields");

            for (var index = 1; index <= inlineCount; index++)
            {
                InlineShape? shape = null;
                try
                {
                    shape = inlineShapes[index];
                    var metadata = ReadMetadata(shape);
                    if (metadata?.DisplayMode == "block" && metadata.Numbered)
                        UpdateEquationNumberFields(
                            document,
                            shape.Height,
                            (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                            metadata.FormulaId);
                }
                finally { Release(shape); }
            }
            TraceReconcileStage("update-inline-number-fields");
            foreach (var formulaId in ommlFormulaIds)
            {
                Bookmark? bookmark = null;
                try
                {
                    bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                    if (bookmark is null) continue;
                    var metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
                    if (metadata?.DisplayMode == "block" && metadata.Numbered)
                        UpdateEquationNumberFields(
                            document,
                            WordOmmlFormulaStore.EstimateHeightPoints(bookmark),
                            (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                            formulaId);
                }
                finally { Release(bookmark); }
            }
            TraceReconcileStage("update-omml-number-fields");
            UpdateNativeCrossReferences(document);
            TraceReconcileStage("update-cross-references");
        }
        finally { Release(inlineShapes); }

        return numberedFormulaIds.Count;
    }

    private sealed class OmmlStructuralFormulaEntry
    {
        public OmmlStructuralFormulaEntry(string formulaId, int position)
        {
            FormulaId = formulaId;
            Position = position;
        }

        public string FormulaId { get; }
        public int Position { get; }
    }

    private static IReadOnlyList<string> GetOmmlDisplayFormulaIdsForStructuralEdit(
        Document document)
    {
        var entries = new List<OmmlStructuralFormulaEntry>();
        var formulaIds = WordOmmlFormulaStore.BookmarkedFormulaIds(document);
        foreach (var formulaId in formulaIds)
        {
            Range? range = null;
            try
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                if (metadata?.DisplayMode != "block") continue;
                range = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    formulaId,
                    metadata);
                entries.Add(new OmmlStructuralFormulaEntry(formulaId, range.Start));
            }
            finally { Release(range); }
        }
        var ordered = entries
            .OrderByDescending(entry => entry.Position)
            .ToArray();
        return ordered
            .Select(entry => entry.FormulaId)
            .ToArray();
    }

    private static FormulaMetadata? ReadMetadata(InlineShape shape) =>
        WordFormulaMetadataReader.TryRead(shape);

    private static bool ConfigureNumberedDisplayFormula(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId,
        bool reuseExistingScaffold = false,
        Table? knownNumberedTable = null,
        bool useConversionSafeVisibleNumber = false)
    {
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        var performanceWatch = tracePerformance
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        var performanceCheckpoint = 0L;
        void TraceStage(string stage)
        {
            if (performanceWatch is null) return;
            var elapsed = performanceWatch.ElapsedMilliseconds;
            TraceNumberingPerformance(
                $"[perf] ConfigureNumbered.{stage}: +{elapsed - performanceCheckpoint}ms ({elapsed}ms)");
            performanceCheckpoint = elapsed;
        }

        EnsureNumberedOmmlIsDisplay(formulaRange);
        TraceStage("ensure-display");
        EnsureStandardNumberedEquationTable(
            document,
            formulaRange,
            formulaId,
            knownNumberedTable);
        TraceStage("ensure-table");
        if (!reuseExistingScaffold)
        {
            ConfigureEquationParagraph(formulaRange, numbered: false);
            TraceStage("paragraph");
            ConfigureNumberedEquationTable(formulaRange);
            TraceStage("table-format");
        }
        var sequenceName = GetNativeEquationSequenceName(document);
        EnsureNativeCaption(
            document,
            formulaRange,
            formulaId,
            sequenceName,
            restyleExisting: !reuseExistingScaffold,
            knownNumberedTable);
        TraceStage("native-caption");
        var visibleNumberCreated = EnsureVisibleEquationNumber(
            document,
            formulaRange,
            formulaHeightPoints,
            formulaFontSizePoints,
            formulaId,
            adoptExistingTableReference:
                reuseExistingScaffold && knownNumberedTable is null,
            knownNumberedTable,
            useConversionSafeVisibleNumber);
        TraceStage("visible-ref");
        TrimBenignEmptyRowsFromNumberedTable(document, formulaRange, formulaId);
        TraceStage("trim-empty-rows");
        return visibleNumberCreated;
    }

    private static void TrimBenignEmptyRowsFromNumberedTable(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Table? table = null;
        Rows? rows = null;
        Columns? columns = null;
        Row? row = null;
        Range? rowRange = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        try
        {
            table = FindNumberedEquationTable(document, formulaId);
            if (table is null) return;
            columns = table.Columns;
            rows = table.Rows;
            if (columns.Count != 3 || rows.Count <= 1) return;

            var formulaRow = 0;
            for (var index = 1; index <= rows.Count; index++)
            {
                Release(rowRange);
                rowRange = null;
                Release(row);
                row = rows[index];
                rowRange = row.Range;
                if (formulaRange.Start >= rowRange.Start
                    && formulaRange.Start < rowRange.End)
                {
                    if (formulaRow != 0) return;
                    formulaRow = index;
                    continue;
                }

                Release(shapes);
                shapes = rowRange.InlineShapes;
                if (shapes.Count != 0) return;
                Release(maths);
                maths = rowRange.OMaths;
                if (maths.Count != 0) return;
                Release(fields);
                fields = rowRange.Fields;
                if (fields.Count != 0) return;
                Release(bookmarks);
                bookmarks = rowRange.Bookmarks;
                if (bookmarks.Count != 0) return;
                foreach (var character in rowRange.Text ?? string.Empty)
                {
                    if (character is '\r' or '\a' or '\n' or '\v'
                        or '\u200b' or '\u200c' or '\u200d' or '\ufeff'
                        || char.IsWhiteSpace(character))
                        continue;
                    return;
                }
            }
            if (formulaRow == 0) return;

            for (var index = rows.Count; index >= 1; index--)
            {
                if (index == formulaRow) continue;
                Release(row);
                row = rows[index];
                row.Delete();
            }
        }
        catch
        {
            // This is a defensive cleanup for a Word table-conversion quirk.
            // Numbering itself is already durable, so never make formula creation
            // fail solely because an extra empty row could not be trimmed.
        }
        finally
        {
            Release(bookmarks);
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(rowRange);
            Release(row);
            Release(columns);
            Release(rows);
            Release(table);
        }
    }

    private static void EnsureStandardNumberedEquationTable(
        Document document,
        Range formulaRange,
        string formulaId,
        Table? knownNumberedTable = null)
    {
        if (knownNumberedTable is not null || IsNumberedEquationTable(formulaRange)) return;
        RemoveVisibleEquationNumber(document, formulaId);

        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefixRange = null;
        Range? suffixRange = null;
        Range? sourceContent = null;
        Range? formattedContent = null;
        Range? tableAnchor = null;
        Range? documentContent = null;
        Range? sourceDeleteRange = null;
        Cell? centerCell = null;
        Range? centerCellRange = null;
        Range? centerInsertion = null;
        Table? table = null;
        Columns? columns = null;
        Column? addedColumn = null;
        var sourceDeleted = false;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "VisualTeX cannot safely number a display formula spanning multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable))
                throw new InvalidOperationException(
                    "VisualTeX cannot safely migrate this nonstandard table formula to the numbered layout.");

            object prefixStart = paragraphRange.Start;
            object prefixEnd = Math.Max(paragraphRange.Start, formulaRange.Start);
            prefixRange = document.Range(ref prefixStart, ref prefixEnd);
            var suffixStartPosition = Math.Min(formulaRange.End, paragraphRange.End);
            object suffixStart = suffixStartPosition;
            object suffixEnd = Math.Max(suffixStartPosition, paragraphRange.End - 1);
            suffixRange = document.Range(ref suffixStart, ref suffixEnd);
            if (!IsNumberingParagraphAdornment(prefixRange.Text)
                || !IsNumberingParagraphAdornment(suffixRange.Text))
                throw new InvalidOperationException(
                    "VisualTeX only batch-numbers display formulas that occupy their own paragraph.");

            var sourceStart = paragraphRange.Start;
            var sourceIsOmml = formulaRange.OMaths.Count > 0;
            if (!sourceIsOmml)
            {
                ConvertStandaloneOleParagraphToNumberedTable(
                    document,
                    paragraphRange,
                    formulaId,
                    formulaRange);
                return;
            }
            sourceContent = paragraphRange.Duplicate;
            sourceContent.End = Math.Max(sourceContent.Start, sourceContent.End - 1);
            formattedContent = sourceContent.FormattedText;

            // Insert one ordinary paragraph after the source formula. Word expands
            // paragraphRange to include that new paragraph, so paragraphRange.End
            // becomes the start of the following paragraph (which may itself begin
            // with another OMath). Keep the pre-insertion end as the table anchor;
            // using the expanded End would ask Tables.Add to operate inside the next
            // mathematical formula.
            var insertedParagraphStart = paragraphRange.End;
            paragraphRange.InsertParagraphAfter();
            documentContent = document.Content;
            var tableAnchorPosition = Math.Max(
                documentContent.Start,
                Math.Min(insertedParagraphStart, documentContent.End - 1));
            object anchorStart = tableAnchorPosition;
            object anchorEnd = tableAnchorPosition;
            tableAnchor = document.Range(ref anchorStart, ref anchorEnd);
            try
            {
                table = document.Tables.Add(tableAnchor, 1, 3);
            }
            catch (Exception error)
            {
                throw new InvalidOperationException(
                    $"Word could not create the numbered OMML table for {formulaId} "
                    + $"at {sourceStart} (paragraph {paragraphRange.Start}-{paragraphRange.End}, "
                    + $"story {(int)paragraphRange.StoryType}).",
                    error);
            }
            columns = table.Columns;
            while (columns.Count < 3)
            {
                object appendAtRight = Type.Missing;
                addedColumn = columns.Add(ref appendAtRight);
                Release(addedColumn);
                addedColumn = null;
            }
            centerCell = table.Cell(1, 2);
            centerCellRange = centerCell.Range;
            centerInsertion = centerCellRange.Duplicate;
            centerInsertion.End = Math.Max(
                centerInsertion.Start,
                centerInsertion.End - 1);
            centerInsertion.Collapse(WdCollapseDirection.wdCollapseStart);
            centerInsertion.FormattedText = formattedContent;

            RefreshFormulaRangeInNumberedTable(
                document,
                table,
                formulaId,
                formulaRange,
                allowUnanchoredOmml: true);

            DeleteOriginalStandaloneFormulaContent(
                document,
                table,
                formulaId,
                sourceIsOmml);
            sourceDeleted = true;

            RefreshFormulaRangeInNumberedTable(
                document,
                table,
                formulaId,
                formulaRange,
                allowUnanchoredOmml: true);
            RemoveNumberingTableCenterDecorations(document, table, formulaRange);
            RefreshFormulaRangeInNumberedTable(
                document,
                table,
                formulaId,
                formulaRange,
                allowUnanchoredOmml: true);

            var ommlMetadata = WordOmmlFormulaStore.TryRead(document, formulaId);
            if (ommlMetadata is not null && formulaRange.OMaths.Count > 0)
            {
                Bookmark? migratedBookmark = null;
                try
                {
                    migratedBookmark = WordOmmlFormulaStore.Wrap(
                        document,
                        formulaRange,
                        ommlMetadata,
                        replaceExisting: true);
                }
                finally { Release(migratedBookmark); }
            }
        }
        catch
        {
            if (!sourceDeleted && table is not null)
            {
                Range? rollbackRange = null;
                try
                {
                    rollbackRange = table.Range;
                    rollbackRange.Delete();
                }
                catch { }
                finally { Release(rollbackRange); }
            }
            throw;
        }
        finally
        {
            Release(addedColumn);
            Release(columns);
            Release(table);
            Release(centerInsertion);
            Release(centerCellRange);
            Release(centerCell);
            Release(sourceDeleteRange);
            Release(documentContent);
            Release(tableAnchor);
            Release(formattedContent);
            Release(sourceContent);
            Release(suffixRange);
            Release(prefixRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void ConvertStandaloneOleParagraphToNumberedTable(
        Document document,
        Range paragraphRange,
        string formulaId,
        Range formulaRange)
    {
        Range? conversionRange = null;
        Table? table = null;
        Columns? columns = null;
        Column? originalColumn = null;
        Column? addedColumn = null;
        Column? leftColumn = null;
        Column? centerColumn = null;
        Column? rightColumn = null;
        try
        {
            conversionRange = paragraphRange.Duplicate;
            object separator = WdTableFieldSeparator.wdSeparateByParagraphs;
            object numRows = 1;
            object numColumns = 1;
            object initialColumnWidth = Type.Missing;
            object format = Type.Missing;
            object applyBorders = Type.Missing;
            object applyShading = Type.Missing;
            object applyFont = Type.Missing;
            object applyColor = Type.Missing;
            object applyHeadingRows = Type.Missing;
            object applyLastRow = Type.Missing;
            object applyFirstColumn = Type.Missing;
            object applyLastColumn = Type.Missing;
            object autoFit = false;
            object autoFitBehavior = WdAutoFitBehavior.wdAutoFitFixed;
            object defaultTableBehavior = WdDefaultTableBehavior.wdWord9TableBehavior;
            table = conversionRange.ConvertToTable(
                ref separator,
                ref numRows,
                ref numColumns,
                ref initialColumnWidth,
                ref format,
                ref applyBorders,
                ref applyShading,
                ref applyFont,
                ref applyColor,
                ref applyHeadingRows,
                ref applyLastRow,
                ref applyFirstColumn,
                ref applyLastColumn,
                ref autoFit,
                ref autoFitBehavior,
                ref defaultTableBehavior);

            columns = table.Columns;
            if (columns.Count != 1)
                throw new InvalidOperationException(
                    "Word did not create the expected one-column OLE migration table.");
            originalColumn = columns[1];
            var originalTableWidth = originalColumn.Width;
            object beforeOriginal = originalColumn;
            addedColumn = columns.Add(ref beforeOriginal);
            Release(addedColumn);
            addedColumn = null;
            object appendAtRight = Type.Missing;
            addedColumn = columns.Add(ref appendAtRight);
            Release(addedColumn);
            addedColumn = null;
            Release(originalColumn);
            originalColumn = null;
            Release(columns);
            columns = table.Columns;
            if (columns.Count != 3)
                throw new InvalidOperationException(
                    "Word did not preserve the expected three-column OLE migration table.");

            // ConvertToTable starts with one column that already spans the full
            // text width. Word clones that physical width when Columns.Add is
            // used, so the temporary three-column table becomes 300% wide.
            // PreferredWidth percentages written later do not reliably shrink a
            // fixed-layout converted table, leaving both the OLE and its number
            // off the visible page. Normalize the physical widths immediately
            // while the original one-column width is still known.
            leftColumn = columns[1];
            centerColumn = columns[2];
            rightColumn = columns[3];
            leftColumn.SetWidth(
                originalTableWidth * 0.2f,
                WdRulerStyle.wdAdjustNone);
            centerColumn.SetWidth(
                originalTableWidth * 0.6f,
                WdRulerStyle.wdAdjustNone);
            rightColumn.SetWidth(
                originalTableWidth * 0.2f,
                WdRulerStyle.wdAdjustNone);

            RefreshFormulaRangeInNumberedTable(
                document,
                table,
                formulaId,
                formulaRange);
            RemoveNumberingTableCenterDecorations(document, table, formulaRange);
            RefreshFormulaRangeInNumberedTable(
                document,
                table,
                formulaId,
                formulaRange);
        }
        finally
        {
            Release(rightColumn);
            Release(centerColumn);
            Release(leftColumn);
            Release(addedColumn);
            Release(originalColumn);
            Release(columns);
            Release(table);
            Release(conversionRange);
        }
    }

    private static void DeleteOriginalStandaloneFormulaContent(
        Document document,
        Table migratedTable,
        string formulaId,
        bool sourceIsOmml)
    {
        Bookmark? bookmark = null;
        Range? originalFormulaRange = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? editableRange = null;
        Range? migratedTableRange = null;
        try
        {
            migratedTableRange = migratedTable.Range;
            var migratedTableStart = migratedTableRange.Start;
            if (sourceIsOmml)
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                    ?? throw new InvalidOperationException(
                        "Word lost the original VisualTeX OMML anchor during numbered-layout migration.");
                originalFormulaRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            }
            else
            {
                shapes = document.InlineShapes;
                for (var index = 1; index <= shapes.Count; index++)
                {
                    Release(shape);
                    shape = shapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata is null
                        || !string.Equals(
                            metadata.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    Range? candidate = null;
                    try
                    {
                        candidate = shape.Range;
                        var isMigratedCopy =
                            (bool)candidate.get_Information(WdInformation.wdWithInTable)
                            && candidate.Tables.Count > 0
                            && candidate.Tables[1].Range.Start == migratedTableStart;
                        if (isMigratedCopy) continue;
                        originalFormulaRange = candidate;
                        candidate = null;
                        break;
                    }
                    finally { Release(candidate); }
                }
                if (originalFormulaRange is null)
                    throw new InvalidOperationException(
                        "Word lost the original VisualTeX OLE object during numbered-layout migration.");
            }

            if ((bool)originalFormulaRange.get_Information(WdInformation.wdWithInTable)
                && originalFormulaRange.Tables.Count > 0
                && originalFormulaRange.Tables[1].Range.Start == migratedTableStart)
                throw new InvalidOperationException(
                    "Word could not distinguish the original formula from its numbered-layout copy.");

            paragraphs = originalFormulaRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The original VisualTeX display formula no longer occupies one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            editableRange = paragraphRange.Duplicate;
            editableRange.End = Math.Max(
                editableRange.Start,
                editableRange.End - 1);
            editableRange.Delete();
        }
        finally
        {
            Release(migratedTableRange);
            Release(editableRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shape);
            Release(shapes);
            Release(originalFormulaRange);
            Release(bookmark);
        }
    }

    private static bool IsNumberingParagraphAdornment(string? text)
    {
        foreach (var character in text ?? string.Empty)
        {
            if (character is '\t' or '\r' or '\n' or '\v'
                or '\u200b' or '\u200c' or '\u200d' or '\ufeff'
                || char.IsWhiteSpace(character))
                continue;
            return false;
        }
        return true;
    }

    private static void RefreshFormulaRangeInNumberedTable(
        Document document,
        Table table,
        string formulaId,
        Range formulaRange,
        bool allowUnanchoredOmml = false)
    {
        Cell? centerCell = null;
        Range? centerRange = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        OMaths? maths = null;
        OMath? math = null;
        Bookmark? bookmark = null;
        Range? refreshed = null;
        try
        {
            centerCell = table.Cell(1, 2);
            centerRange = centerCell.Range;
            shapes = centerRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                var metadata = WordFormulaMetadataReader.TryRead(shape);
                if (metadata is null
                    || !string.Equals(
                        metadata.FormulaId,
                        formulaId,
                        StringComparison.OrdinalIgnoreCase))
                    continue;
                refreshed = shape.Range;
                formulaRange.SetRange(refreshed.Start, refreshed.End);
                return;
            }

            if (allowUnanchoredOmml)
            {
                maths = centerRange.OMaths;
                if (maths.Count == 1)
                {
                    math = maths[1];
                    refreshed = math.Range;
                    formulaRange.SetRange(refreshed.Start, refreshed.End);
                    return;
                }
            }

            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            if (bookmark is not null)
            {
                refreshed = WordOmmlFormulaStore.GetEquationRange(bookmark);
                if ((bool)refreshed.get_Information(WdInformation.wdWithInTable)
                    && refreshed.Tables.Count > 0
                    && refreshed.Tables[1].Range.Start == table.Range.Start)
                {
                    formulaRange.SetRange(refreshed.Start, refreshed.End);
                    return;
                }
            }

            throw new InvalidOperationException(
                "Word did not preserve the VisualTeX formula while creating its numbered layout.");
        }
        finally
        {
            Release(refreshed);
            Release(bookmark);
            Release(math);
            Release(maths);
            Release(shape);
            Release(shapes);
            Release(centerRange);
            Release(centerCell);
        }
    }

    private static void RemoveNumberingTableCenterDecorations(
        Document document,
        Table table,
        Range formulaRange)
    {
        Cell? centerCell = null;
        Range? centerRange = null;
        Range? characterRange = null;
        try
        {
            centerCell = table.Cell(1, 2);
            centerRange = centerCell.Range;
            for (var position = centerRange.End - 2;
                 position >= centerRange.Start;
                 position--)
            {
                if (position >= formulaRange.Start && position < formulaRange.End)
                    continue;
                object characterStart = position;
                object characterEnd = position + 1;
                characterRange = document.Range(ref characterStart, ref characterEnd);
                if (string.Equals(characterRange.Text, "\t", StringComparison.Ordinal)
                    || string.Equals(characterRange.Text, "\v", StringComparison.Ordinal))
                    characterRange.Delete();
                Release(characterRange);
                characterRange = null;
            }
        }
        finally
        {
            Release(characterRange);
            Release(centerRange);
            Release(centerCell);
        }
    }

    private static void EnsureNumberedOmmlIsDisplay(Range formulaRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? refreshed = null;
        try
        {
            maths = formulaRange.OMaths;
            if (maths.Count == 0) return;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay)
            {
                math.Type = WdOMathType.wdOMathDisplay;
                math.BuildUp();
            }
            refreshed = math.Range;
            formulaRange.SetRange(refreshed.Start, refreshed.End);
        }
        finally
        {
            Release(refreshed);
            Release(math);
            Release(maths);
        }
    }

    private static bool IsNumberedEquationTable(Range formulaRange)
    {
        try
        {
            return (bool)formulaRange.get_Information(WdInformation.wdWithInTable)
                && formulaRange.Tables.Count > 0
                && formulaRange.Tables[1].Columns.Count >= 3;
        }
        catch { return false; }
    }

    private static void ConfigureFreshConversionNumberedEquationTable(Range formulaRange)
    {
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        var watch = tracePerformance
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        var checkpoint = 0L;
        void TraceStage(string stage)
        {
            if (watch is null) return;
            var elapsed = watch.ElapsedMilliseconds;
            TraceNumberingPerformance(
                $"[perf] FreshConversionTable.{stage}: +{elapsed - checkpoint}ms ({elapsed}ms)");
            checkpoint = elapsed;
        }

        Table? table = null;
        Columns? columns = null;
        Column? leftColumn = null;
        Column? centerColumn = null;
        Cell? centerCell = null;
        Cell? numberCell = null;
        Range? centerRange = null;
        Range? numberRange = null;
        ParagraphFormat? centerFormat = null;
        ParagraphFormat? numberFormat = null;
        Borders? borders = null;
        try
        {
            table = formulaRange.Tables[1];

            // Tables.Add already creates a fixed table spanning the current text
            // width. Rewriting PreferredWidth=100% and AutoFitFixed here forces
            // an unnecessary Word layout pass for every converted formula. Keep
            // only the VisualTeX-specific cell margins and 20/60/20 geometry.
            table.LeftPadding = 0f;
            table.RightPadding = 0f;
            table.TopPadding = 0f;
            table.BottomPadding = 0f;
            TraceStage("padding");

            columns = table.Columns;
            if (columns.Count < 3)
                throw new InvalidOperationException(
                    "A numbered formula table must contain three columns.");
            leftColumn = columns[1];
            centerColumn = columns[2];

            // On a fresh full-width 1x3 table, setting the left and center
            // preferred percentages makes Word assign the remaining 20% to the
            // right column. Avoid a third width mutation and its extra reflow.
            leftColumn.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            leftColumn.PreferredWidth = 20f;
            centerColumn.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            centerColumn.PreferredWidth = 60f;
            TraceStage("widths");

            borders = table.Borders;
            borders.Enable = 0;
            TraceStage("borders");

            centerCell = table.Cell(1, 2);
            numberCell = table.Cell(1, 3);
            centerCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            numberCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            TraceStage("vertical-align");

            centerRange = centerCell.Range;
            numberRange = numberCell.Range;
            centerFormat = centerRange.ParagraphFormat;
            numberFormat = numberRange.ParagraphFormat;
            centerFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            numberFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            centerFormat.LeftIndent = centerFormat.RightIndent = 0f;
            centerFormat.FirstLineIndent = 0f;
            numberFormat.LeftIndent = numberFormat.RightIndent = 0f;
            numberFormat.FirstLineIndent = 0f;
            centerFormat.SpaceBefore = centerFormat.SpaceAfter = 0f;
            numberFormat.SpaceBefore = numberFormat.SpaceAfter = 0f;
            centerFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            numberFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            try { centerFormat.DisableLineHeightGrid = -1; } catch { }
            try { numberFormat.DisableLineHeightGrid = -1; } catch { }
            TraceStage("paragraphs");
        }
        finally
        {
            Release(borders);
            Release(numberFormat);
            Release(centerFormat);
            Release(numberRange);
            Release(centerRange);
            Release(numberCell);
            Release(centerCell);
            Release(centerColumn);
            Release(leftColumn);
            Release(columns);
            Release(table);
        }
    }

    private static void ConfigureNumberedEquationTable(Range formulaRange)
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
            TraceNumberingPerformance(
                $"[perf] NumberedTable.{stage}: +{elapsed - checkpoint}ms ({elapsed}ms)");
            checkpoint = elapsed;
        }

        Table? table = null;
        Columns? columns = null;
        Column? leftColumn = null;
        Column? centerColumn = null;
        Column? rightColumn = null;
        Cell? centerCell = null;
        Cell? numberCell = null;
        Range? centerRange = null;
        Range? numberRange = null;
        ParagraphFormat? centerFormat = null;
        ParagraphFormat? numberFormat = null;
        Borders? borders = null;
        try
        {
            table = formulaRange.Tables[1];
            table.AllowAutoFit = false;
            table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            table.PreferredWidth = 100f;
            table.LeftPadding = 0f;
            table.RightPadding = 0f;
            table.TopPadding = 0f;
            table.BottomPadding = 0f;
            try { table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed); } catch { }
            TraceStage("table-geometry");
            columns = table.Columns;
            leftColumn = columns[1];
            centerColumn = columns[2];
            rightColumn = columns[3];
            leftColumn.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            centerColumn.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            rightColumn.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            leftColumn.PreferredWidth = 20f;
            centerColumn.PreferredWidth = 60f;
            rightColumn.PreferredWidth = 20f;
            TraceStage("column-widths");
            borders = table.Borders;
            borders.Enable = 0;
            TraceStage("borders");
            centerCell = table.Cell(1, 2);
            numberCell = table.Cell(1, 3);
            centerCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            numberCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            TraceStage("cell-vertical");
            centerRange = centerCell.Range;
            numberRange = numberCell.Range;
            centerFormat = centerRange.ParagraphFormat;
            numberFormat = numberRange.ParagraphFormat;
            centerFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            numberFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            centerFormat.LeftIndent = centerFormat.RightIndent = 0f;
            centerFormat.FirstLineIndent = 0f;
            numberFormat.LeftIndent = numberFormat.RightIndent = 0f;
            numberFormat.FirstLineIndent = 0f;
            centerFormat.SpaceBefore = centerFormat.SpaceAfter = 0f;
            numberFormat.SpaceBefore = numberFormat.SpaceAfter = 0f;
            centerFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            numberFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            TraceStage("paragraph-format");
            try { centerFormat.DisableLineHeightGrid = -1; } catch { }
            try { numberFormat.DisableLineHeightGrid = -1; } catch { }
            TraceStage("line-grid");
        }
        finally
        {
            Release(rightColumn);
            Release(centerColumn);
            Release(leftColumn);
            Release(columns);
            Release(borders);
            Release(numberFormat);
            Release(centerFormat);
            Release(numberRange);
            Release(centerRange);
            Release(numberCell);
            Release(centerCell);
            Release(table);
        }
    }

    private static void ConfigureUnnumberedDisplayFormula(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        RemoveVisibleEquationNumber(document, formulaId);
        RemoveNativeCaption(document, formulaId);
        RemoveLeadingEquationTab(document, formulaRange);
        ConfigureEquationParagraph(formulaRange, numbered: false);
    }

    private static void ConfigureEquationParagraph(Range formulaRange, bool numbered)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        ListFormat? listFormat = null;
        OMaths? maths = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            format = paragraph.Format;
            var nativeOmmlParagraph = false;
            try
            {
                maths = formulaRange.OMaths;
                nativeOmmlParagraph = maths.Count > 0
                    && !(bool)formulaRange.get_Information(WdInformation.wdWithInTable);
            }
            catch { }
            format.LeftIndent = 0f;
            format.RightIndent = 0f;
            format.FirstLineIndent = 0f;
            if (!nativeOmmlParagraph)
            {
                // OLE/picture formulas carry their own transparent preview
                // margins, so their host paragraph must stay compact. Native
                // OMML formulas have no image padding and should retain the
                // document paragraph style exactly as a formula inserted by
                // Word itself would.
                format.SpaceBefore = 0f;
                format.SpaceAfter = 0f;
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            }
            format.KeepTogether = 0;
            format.KeepWithNext = 0;
            format.PageBreakBefore = 0;
            format.WidowControl = 0;
            try
            {
                listFormat = paragraphRange.ListFormat;
                listFormat.RemoveNumbers(WdNumberType.wdNumberParagraph);
            }
            catch
            {
                // Protected/custom stories can reject list normalization. The
                // page-break flags above still remove Word's black paragraph
                // marker when formatting marks are shown.
            }
            tabStops = format.TabStops;
            tabStops.ClearAll();

            if (!numbered)
            {
                format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                return;
            }

            var pageWidth = 612f;
            var leftMargin = 72f;
            var rightMargin = 72f;
            try
            {
                sections = paragraphRange.Sections;
                if (sections.Count > 0)
                {
                    section = sections[1];
                    pageSetup = section.PageSetup;
                    pageWidth = pageSetup.PageWidth;
                    leftMargin = pageSetup.LeftMargin;
                    rightMargin = pageSetup.RightMargin;
                }
            }
            catch
            {
                // Standard US Letter and one-inch margins are a safe fallback
                // for protected/custom stories without an exposed PageSetup.
            }

            var positions = CalculateEquationTabStops(pageWidth, leftMargin, rightMargin, 0, 0);
            format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            tabStops.Add(
                positions.Center,
                (WdTabAlignment)WdTabAlignmentCenter,
                (WdTabLeader)WdTabLeaderSpaces);
            tabStops.Add(
                positions.Right,
                (WdTabAlignment)WdTabAlignmentRight,
                (WdTabLeader)WdTabLeaderSpaces);
        }
        finally
        {
            Release(maths);
            Release(listFormat);
            Release(tabStops);
            Release(format);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static (float Center, float Right) CalculateEquationTabStops(
        float pageWidth,
        float leftMargin,
        float rightMargin,
        float leftIndent,
        float rightIndent)
    {
        var availableWidth = Math.Max(
            72f,
            pageWidth
                - Math.Max(0f, leftMargin)
                - Math.Max(0f, rightMargin)
                - Math.Max(0f, leftIndent)
                - Math.Max(0f, rightIndent));
        return (availableWidth / 2f, availableWidth);
    }

    internal static int CalculateEquationNumberFontPosition(
        float formulaHeightPoints,
        float numberFontSizePoints)
    {
        if (float.IsNaN(formulaHeightPoints)
            || float.IsInfinity(formulaHeightPoints)
            || formulaHeightPoints <= 0)
            return 0;
        if (float.IsNaN(numberFontSizePoints)
            || float.IsInfinity(numberFontSizePoints)
            || numberFontSizePoints <= 0
            || numberFontSizePoints > 256)
            numberFontSizePoints = 11f;
        return Math.Max(
            0,
            (int)Math.Round(
                (formulaHeightPoints - numberFontSizePoints) / 2f,
                MidpointRounding.AwayFromZero));
    }

    internal static (string Text, int FieldOffset) EquationNumberScaffold() => ("\t()", 2);

    private static void EnsureLeadingEquationTab(Document document, Range formulaRange)
    {
        Range? preceding = null;
        Range? insertion = null;
        try
        {
            var insertionPosition = formulaRange.Start;
            if (formulaRange.Start > 0)
            {
                object precedingStart = formulaRange.Start - 1;
                object precedingEnd = formulaRange.Start;
                preceding = document.Range(ref precedingStart, ref precedingEnd);
                if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal)) return;

                // A display OMath is preceded by Word's vertical-tab math
                // separator (0x0B). Its layout tab therefore sits one character
                // earlier and must be inspected/inserted outside the OMath edge.
                if (string.Equals(preceding.Text, "\v", StringComparison.Ordinal))
                {
                    insertionPosition = formulaRange.Start - 1;
                    if (formulaRange.Start > 1)
                    {
                        preceding.SetRange(formulaRange.Start - 2, formulaRange.Start - 1);
                        if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal)) return;
                    }
                }
            }

            object start = insertionPosition;
            object end = insertionPosition;
            insertion = document.Range(ref start, ref end);
            insertion.Text = "\t";
        }
        finally
        {
            Release(insertion);
            Release(preceding);
        }
    }

    private static void RemoveLeadingEquationTab(Document document, Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? preceding = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (formulaRange.Start <= paragraphRange.Start) return;
            object start = formulaRange.Start - 1;
            object end = formulaRange.Start;
            preceding = document.Range(ref start, ref end);
            if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal))
            {
                preceding.Delete();
                return;
            }
            if (string.Equals(preceding.Text, "\v", StringComparison.Ordinal)
                && formulaRange.Start - 2 >= paragraphRange.Start)
            {
                preceding.SetRange(formulaRange.Start - 2, formulaRange.Start - 1);
                if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal))
                    preceding.Delete();
            }
        }
        finally
        {
            Release(preceding);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void EnsureNativeCaption(
        Document document,
        Range formulaRange,
        string formulaId,
        string nativeSequenceName,
        bool restyleExisting = true,
        Table? knownNumberedTable = null)
    {
        if (TryGetNativeCaptionRanges(
                document,
                formulaId,
                nativeSequenceName,
                out var captionRange,
                out var numberRange)
            && captionRange is not null
            && numberRange is not null)
        {
            try
            {
                if (restyleExisting)
                    StyleNativeCaption(captionRange, numberRange);
            }
            finally
            {
                Release(numberRange);
                Release(captionRange);
            }
            return;
        }
        Release(numberRange);
        Release(captionRange);

        RemoveNativeCaption(document, formulaId);
        CreateNativeCaption(
            document,
            formulaRange,
            formulaId,
            nativeSequenceName,
            knownNumberedTable);
    }

    private static bool CanReuseEmptyNativeCaptionParagraph(
        Document document,
        int position)
    {
        Range? content = null;
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Frames? frames = null;
        Tables? tables = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        try
        {
            content = document.Content;
            if (position < content.Start || position >= content.End) return false;
            probe = document.Range(position, Math.Min(content.End, position + 1));
            if ((bool)probe.get_Information(WdInformation.wdWithInTable)) return false;
            frames = probe.Frames;
            if (frames.Count > 0) return false;
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.Start != position) return false;
            tables = paragraphRange.Tables;
            shapes = paragraphRange.InlineShapes;
            maths = paragraphRange.OMaths;
            fields = paragraphRange.Fields;
            if (tables.Count > 0 || shapes.Count > 0 || maths.Count > 0 || fields.Count > 0)
                return false;
            var text = (paragraphRange.Text ?? string.Empty)
                .Trim('\r', '\n', '\t', '\v', ' ');
            return text.Length == 0;
        }
        finally
        {
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(tables);
            Release(frames);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
            Release(content);
        }
    }

    private static void CreateNativeCaption(
        Document document,
        Range formulaRange,
        string formulaId,
        string nativeSequenceName,
        Table? knownNumberedTable = null,
        bool deferFieldUpdate = false,
        int? plannedOrdinal = null,
        string? plannedPrefix = null)
    {
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        var performanceWatch = tracePerformance
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        var performanceCheckpoint = 0L;
        void TraceStage(string stage)
        {
            if (performanceWatch is null) return;
            var elapsed = performanceWatch.ElapsedMilliseconds;
            TraceNumberingPerformance(
                $"[perf] NativeCaption.{stage}: +{elapsed - performanceCheckpoint}ms ({elapsed}ms)");
            performanceCheckpoint = elapsed;
        }
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Table? formulaTable = null;
        Range? formulaTableRange = null;
        Range? fieldInsertion = null;
        Fields? fields = null;
        Field? captionField = null;
        Range? numberRange = null;
        Range? captionRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Bookmarks? bookmarks = null;
        try
        {
            // Word's Range.InsertCaption mutates the equation paragraph. For a
            // trailing inline OMath it moves the ordinary run after the equation
            // into a new paragraph, so the visible REF field is subsequently
            // absorbed into m:oMath. Build the native SEQ caption in a dedicated
            // hidden paragraph instead and leave the equation paragraph intact.
            int captionStart;
            if (knownNumberedTable is not null)
            {
                formulaTableRange = knownNumberedTable.Range;
                // A Word table's Range.InsertParagraphAfter() expands the table
                // range itself; its new End is then the start of the following
                // body paragraph. Adding the native SEQ field at that position
                // merges the caption with the user's next paragraph, and the
                // 0.1 pt clipping Frame subsequently hides that paragraph and any
                // inline formula in it. Capture the original table boundary and
                // create a dedicated paragraph there instead.
                captionStart = formulaTableRange.End;
                if (!CanReuseEmptyNativeCaptionParagraph(document, captionStart))
                {
                    fieldInsertion = document.Range(captionStart, captionStart);
                    fieldInsertion.InsertParagraphBefore();
                    Release(fieldInsertion);
                    fieldInsertion = null;
                }
            }
            else if (IsNumberedEquationTable(formulaRange))
            {
                formulaTable = formulaRange.Tables[1];
                formulaTableRange = formulaTable.Range;
                captionStart = formulaTableRange.End;
                if (!CanReuseEmptyNativeCaptionParagraph(document, captionStart))
                {
                    fieldInsertion = document.Range(captionStart, captionStart);
                    fieldInsertion.InsertParagraphBefore();
                    Release(fieldInsertion);
                    fieldInsertion = null;
                }
            }
            else
            {
                formulaParagraphs = formulaRange.Paragraphs;
                formulaParagraph = formulaParagraphs[1];
                formulaParagraphRange = formulaParagraph.Range;
                captionStart = formulaParagraphRange.End;
                formulaParagraphRange.InsertParagraphAfter();
            }

            object insertionStart = captionStart;
            object insertionEnd = captionStart;
            fieldInsertion = document.Range(ref insertionStart, ref insertionEnd);
            if (!string.IsNullOrEmpty(plannedPrefix))
            {
                fieldInsertion.Text = plannedPrefix;
                var fieldPosition = fieldInsertion.End;
                Release(fieldInsertion);
                fieldInsertion = document.Range(fieldPosition, fieldPosition);
            }
            TraceStage("prepare-paragraph");
            // Add through the tiny target Range rather than document.Fields.
            // Materializing the full document field collection on every appended
            // equation makes Word's field bookkeeping grow with formula count.
            fields = fieldInsertion.Fields;
            object fieldType = WdFieldEmpty;
            object fieldCode = plannedOrdinal.HasValue
                ? $"SEQ {nativeSequenceName} \\r {plannedOrdinal.Value} \\* ARABIC"
                : $"SEQ {nativeSequenceName} \\* ARABIC";
            object preserveFormatting = true;
            captionField = fields.Add(
                fieldInsertion,
                ref fieldType,
                ref fieldCode,
                ref preserveFormatting);
            TraceStage("add-seq");
            if (!deferFieldUpdate)
                captionField.Update();
            TraceStage(deferFieldUpdate ? "defer-update-seq" : "update-seq");
            numberRange = captionField.Result;
            paragraphs = numberRange.Paragraphs;
            paragraph = paragraphs[1];
            captionRange = paragraph.Range;
            try
            {
                object captionStyle = WdBuiltinStyle.wdStyleCaption;
                captionRange.set_Style(ref captionStyle);
            }
            catch
            {
                // Some locked/custom documents reject assigning the built-in
                // caption style. The SEQ field and bookmarks remain valid.
            }

            bookmarks = document.Bookmarks;
            if (!string.IsNullOrEmpty(plannedPrefix))
            {
                Range? plannedNumberRange = null;
                try
                {
                    plannedNumberRange = document.Range(captionRange.Start, numberRange.End);
                    bookmarks.Add(NativeNumberBookmarkName(formulaId), plannedNumberRange);
                }
                finally { Release(plannedNumberRange); }
            }
            else
            {
                bookmarks.Add(NativeNumberBookmarkName(formulaId), numberRange);
            }
            bookmarks.Add(NativeCaptionBookmarkName(formulaId), captionRange);
            TraceStage("bookmarks");
            if (deferFieldUpdate && knownNumberedTable is not null)
                StyleFreshConversionNativeCaption(captionRange, numberRange);
            else
                StyleNativeCaption(
                    captionRange,
                    numberRange,
                    cleanupLegacyFrames: false);
            TraceStage("style-frame");
        }
        finally
        {
            Release(bookmarks);
            Release(paragraph);
            Release(paragraphs);
            Release(captionRange);
            Release(numberRange);
            Release(captionField);
            Release(fields);
            Release(fieldInsertion);
            Release(formulaTableRange);
            Release(formulaTable);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
        }
    }

    private static Field? FindNewNativeEquationField(
        Document document,
        string nativeSequenceName,
        ISet<int> existingPositions,
        int formulaPosition)
    {
        Fields? fields = null;
        Field? result = null;
        var bestDistance = int.MaxValue;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? fieldResult = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName)) continue;
                    fieldResult = field.Result;
                    if (existingPositions.Contains(fieldResult.Start)) continue;
                    var distance = Math.Abs(fieldResult.Start - formulaPosition);
                    if (distance >= bestDistance) continue;
                    Release(result);
                    result = field;
                    field = null;
                    bestDistance = distance;
                }
                finally
                {
                    Release(fieldResult);
                    Release(code);
                    Release(field);
                }
            }
            return result;
        }
        finally
        {
            Release(fields);
        }
    }

    private static void StyleFreshConversionNativeCaption(
        Range captionRange,
        Range numberRange)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        ParagraphFormat? paragraph = null;
        ListFormat? listFormat = null;
        Frames? frames = null;
        Frame? frame = null;
        Borders? borders = null;
        try
        {
            // Fresh conversion captions have no legacy frames/list state to scan.
            // Keep normal visible formatting so native REF fields inherit black
            // 11pt text, but clip the SEQ paragraph into a 0.1pt frame using Word's
            // built-in page-edge positions. Avoid querying Section.PageSetup for
            // every formula: that forces pagination and becomes O(N^2) as fields
            // accumulate in large format-conversion batches.
            font = captionRange.Font;
            font.Hidden = 0;
            font.Size = 11f;
            font.Color = WdColor.wdColorAutomatic;
            font.Position = 0;
            numberFont = numberRange.Font;
            numberFont.Hidden = 0;
            numberFont.Size = 11f;
            numberFont.Color = WdColor.wdColorAutomatic;
            numberFont.Position = 0;

            paragraph = captionRange.ParagraphFormat;
            paragraph.SpaceBefore = 0f;
            paragraph.SpaceAfter = 0f;
            paragraph.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            paragraph.KeepTogether = 0;
            paragraph.KeepWithNext = 0;
            paragraph.PageBreakBefore = 0;
            paragraph.WidowControl = 0;
            try
            {
                listFormat = captionRange.ListFormat;
                listFormat.RemoveNumbers(WdNumberType.wdNumberParagraph);
            }
            catch { }

            frames = captionRange.Frames;
            frame = frames.Count > 0 ? frames[1] : frames.Add(captionRange);
            const float clippedFrameSize = 0.1f;
            frame.WidthRule = WdFrameSizeRule.wdFrameExact;
            frame.HeightRule = WdFrameSizeRule.wdFrameExact;
            frame.Width = clippedFrameSize;
            frame.Height = clippedFrameSize;
            frame.RelativeHorizontalPosition =
                WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            frame.RelativeVerticalPosition =
                WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            frame.HorizontalPosition = (float)WdFramePosition.wdFrameRight;
            frame.VerticalPosition = (float)WdFramePosition.wdFrameBottom;
            frame.TextWrap = false;
            frame.LockAnchor = true;
            borders = captionRange.Borders;
            borders.Enable = 0;
        }
        finally
        {
            Release(borders);
            Release(frame);
            Release(frames);
            Release(listFormat);
            Release(paragraph);
            Release(numberFont);
            Release(font);
        }
    }

    private static void StyleNativeCaption(
        Range captionRange,
        Range numberRange,
        bool cleanupLegacyFrames = true)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        ParagraphFormat? paragraph = null;
        ListFormat? listFormat = null;
        Frames? frames = null;
        Frame? frame = null;
        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        Borders? borders = null;
        try
        {
            // Keep the native SEQ target in ordinary visible formatting. Word's
            // own InsertCrossReference creates a plain REF ... \\h field, which
            // inherits the target formatting on every F9 update. Hiding this
            // paragraph with white 1 pt text therefore made native references
            // white and one point as well.
            font = captionRange.Font;
            font.Hidden = 0;
            font.Size = 11f;
            font.Color = WdColor.wdColorAutomatic;
            font.Position = 0;
            numberFont = numberRange.Font;
            numberFont.Hidden = 0;
            numberFont.Size = 11f;
            numberFont.Color = WdColor.wdColorAutomatic;
            numberFont.Position = 0;

            paragraph = captionRange.ParagraphFormat;
            paragraph.SpaceBefore = 0f;
            paragraph.SpaceAfter = 0f;
            paragraph.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            paragraph.KeepTogether = 0;
            paragraph.KeepWithNext = 0;
            paragraph.PageBreakBefore = 0;
            paragraph.WidowControl = 0;
            try
            {
                listFormat = captionRange.ListFormat;
                listFormat.RemoveNumbers(WdNumberType.wdNumberParagraph);
            }
            catch { }

            // A legacy Word Frame remains in the main document story, so Word's
            // native Cross-reference dialog and plain REF fields still recognize
            // the SEQ target. Negative frame coordinates are clamped by Word to
            // the top-left of the page, which exposed the number and frame edge.
            // Use an exact 0.1 pt clipping frame at the bottom-right page boundary
            // instead. The SEQ retains normal black 11 pt formatting for native
            // REF inheritance, while its rendered content lies beyond the page.
            var document = captionRange.Document;
            if (cleanupLegacyFrames)
                RemoveLegacyEmptyCaptionFrames(document, captionRange);
            frames = captionRange.Frames;
            frame = frames.Count > 0 ? frames[1] : frames.Add(captionRange);
            sections = captionRange.Sections;
            section = sections[1];
            pageSetup = section.PageSetup;
            const float clippedFrameSize = 0.1f;
            frame.WidthRule = WdFrameSizeRule.wdFrameExact;
            frame.HeightRule = WdFrameSizeRule.wdFrameExact;
            frame.Width = clippedFrameSize;
            frame.Height = clippedFrameSize;
            frame.RelativeHorizontalPosition =
                WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            frame.RelativeVerticalPosition =
                WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            frame.HorizontalPosition = Math.Max(0f, pageSetup.PageWidth - clippedFrameSize);
            frame.VerticalPosition = Math.Max(0f, pageSetup.PageHeight - clippedFrameSize);
            frame.TextWrap = false;
            frame.LockAnchor = true;
            borders = captionRange.Borders;
            borders.Enable = 0;
        }
        finally
        {
            Release(borders);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(frame);
            Release(frames);
            Release(listFormat);
            Release(paragraph);
            Release(numberFont);
            Release(font);
        }
    }

    private static void RemoveLegacyEmptyCaptionFrames(
        Document document,
        Range keepRange)
    {
        Frames? frames = null;
        try
        {
            frames = document.Frames;
            for (var index = frames.Count; index >= 1; index--)
            {
                Frame? candidate = null;
                Range? range = null;
                Fields? fields = null;
                try
                {
                    candidate = frames[index];
                    range = candidate.Range;
                    if (range.Start <= keepRange.End && range.End >= keepRange.Start)
                        continue;
                    fields = range.Fields;
                    var text = (range.Text ?? string.Empty)
                        .Trim('\r', '\n', '\t', '\v', ' ');
                    var oldVisualTeXFrame =
                        candidate.HorizontalPosition <= -999f
                        && candidate.VerticalPosition <= -999f
                        && candidate.Width >= 70f
                        && candidate.Height >= 17f;
                    if (oldVisualTeXFrame
                        && fields.Count == 0
                        && string.IsNullOrEmpty(text))
                        candidate.Delete();
                }
                finally
                {
                    Release(fields);
                    Release(range);
                    Release(candidate);
                }
            }
        }
        finally { Release(frames); }
    }

    private static bool TryGetNativeCaptionRanges(
        Document document,
        string formulaId,
        string nativeSequenceName,
        out Range? captionRange,
        out Range? numberRange)
    {
        captionRange = null;
        numberRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Bookmark? numberBookmark = null;
        Field? nativeField = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            var numberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName) || !bookmarks.Exists(numberName)) return false;
            captionBookmark = bookmarks[captionName];
            numberBookmark = bookmarks[numberName];
            captionRange = captionBookmark.Range;
            numberRange = numberBookmark.Range;
            nativeField = FindNativeEquationFieldInRange(
                captionRange,
                nativeSequenceName);
            if (nativeField is null)
                nativeField = FindNativeEquationFieldAtRange(
                    document,
                    numberRange,
                    nativeSequenceName);
            if (nativeField is not null) return true;
            Release(numberRange);
            numberRange = null;
            Release(captionRange);
            captionRange = null;
            return false;
        }
        finally
        {
            Release(nativeField);
            Release(numberBookmark);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    private static bool EnsureVisibleEquationNumber(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId,
        bool adoptExistingTableReference = false,
        Table? knownNumberedTable = null,
        bool useConversionSafeVisibleNumber = false)
    {
        var targetBookmarkName = NativeNumberBookmarkName(formulaId);
        if (HasVisibleEquationNumber(
                document,
                formulaRange,
                formulaId,
                targetBookmarkName,
                knownNumberedTable)) return false;
        if (adoptExistingTableReference
            && TryAdoptExistingTableEquationNumber(
                document,
                formulaRange,
                formulaFontSizePoints,
                formulaId,
                targetBookmarkName))
            return true;
        RemoveVisibleEquationNumber(document, formulaId);
        InsertVisibleEquationNumber(
            document,
            formulaRange,
            formulaHeightPoints,
            formulaFontSizePoints,
            formulaId,
            targetBookmarkName,
            knownNumberedTable,
            useConversionSafeVisibleNumber);
        return true;
    }

    private static bool TryAdoptExistingTableEquationNumber(
        Document document,
        Range formulaRange,
        float formulaFontSizePoints,
        string formulaId,
        string targetBookmarkName)
    {
        if (!IsNumberedEquationTable(formulaRange)) return false;
        Table? table = null;
        Cell? numberCell = null;
        Range? cellRange = null;
        Range? editableRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        Range? labelRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? existingBookmark = null;
        ParagraphFormat? paragraph = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            table = formulaRange.Tables[1];
            numberCell = table.Cell(1, 3);
            cellRange = numberCell.Range;
            editableRange = cellRange.Duplicate;
            editableRange.End = Math.Max(editableRange.Start, editableRange.End - 1);
            var visibleText = editableRange.Text ?? string.Empty;
            if (!visibleText.StartsWith("(", StringComparison.Ordinal)
                || !visibleText.EndsWith(")", StringComparison.Ordinal))
                return false;

            fields = editableRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(field);
                field = fields[index];
                Release(code);
                code = field.Code;
                if (IsReferenceFieldCode(code.Text)) break;
                Release(field);
                field = null;
            }
            if (field is null) return false;

            if (!IsReferenceToBookmark(code?.Text, targetBookmarkName))
                code!.Text = $" REF {targetBookmarkName} \\h ";
            NormalizeReferenceResult(field, formulaFontSizePoints);
            result = field.Result;
            var labelEnd = Math.Min(editableRange.End, result.End + 1);
            if (labelEnd <= editableRange.Start) return false;
            labelRange = document.Range(editableRange.Start, labelEnd);
            var labelText = labelRange.Text ?? string.Empty;
            if (!labelText.StartsWith("(", StringComparison.Ordinal)
                || !labelText.EndsWith(")", StringComparison.Ordinal))
                return false;

            bookmarks = document.Bookmarks;
            var visibleName = EquationBookmarkName(formulaId);
            if (bookmarks.Exists(visibleName))
            {
                existingBookmark = bookmarks[visibleName];
                existingBookmark.Delete();
                Release(existingBookmark);
                existingBookmark = null;
            }
            bookmarks.Add(visibleName, labelRange);
            paragraph = labelRange.ParagraphFormat;
            paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            font = labelRange.Font;
            ApplyEquationNumberFont(font, formulaFontSizePoints, position: 0);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(font);
            Release(paragraph);
            Release(existingBookmark);
            Release(bookmarks);
            Release(labelRange);
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(editableRange);
            Release(cellRange);
            Release(numberCell);
            Release(table);
        }
    }

    private static bool HasVisibleEquationNumber(
        Document document,
        Range formulaRange,
        string formulaId,
        string targetBookmarkName,
        Table? knownNumberedTable = null)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Fields? fields = null;
        OMaths? maths = null;
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Paragraphs? numberParagraphs = null;
        Paragraph? numberParagraph = null;
        Range? numberParagraphRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return false;
            bookmark = bookmarks[name];
            range = bookmark.Range;

            // Older builds inserted the visible REF field at OMath.Range.End.
            // Word still considers that position part of m:oMath, so the tab,
            // parentheses and number became equation content. A valid number is
            // an ordinary Word run after the formula, never a child of OMML.
            if (range.Start < formulaRange.End) return false;
            maths = range.OMaths;
            if (maths.Count > 0) return false;
            var tableLayout = knownNumberedTable is not null
                || IsNumberedEquationTable(formulaRange);
            var visibleText = range.Text ?? string.Empty;
            var expectedPrefix = tableLayout ? "(" : "\t(";
            if (!visibleText.StartsWith(expectedPrefix, StringComparison.Ordinal)
                || !visibleText.EndsWith(")", StringComparison.Ordinal))
                return false;

            if (tableLayout)
            {
                if (!(bool)range.get_Information(WdInformation.wdWithInTable)
                    || range.Tables.Count == 0)
                    return false;
                Range? expectedTableRange = null;
                Range? visibleTableRange = null;
                try
                {
                    expectedTableRange = (knownNumberedTable ?? formulaRange.Tables[1]).Range;
                    visibleTableRange = range.Tables[1].Range;
                    if (visibleTableRange.Start != expectedTableRange.Start)
                        return false;
                }
                finally
                {
                    Release(visibleTableRange);
                    Release(expectedTableRange);
                }
            }
            else
            {
                formulaParagraphs = formulaRange.Paragraphs;
                formulaParagraph = formulaParagraphs[1];
                formulaParagraphRange = formulaParagraph.Range;
                numberParagraphs = range.Paragraphs;
                numberParagraph = numberParagraphs[1];
                numberParagraphRange = numberParagraph.Range;
                if (formulaParagraphRange.Start != numberParagraphRange.Start) return false;
            }

            fields = range.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (IsReferenceToBookmark(code.Text, targetBookmarkName)) return true;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            return false;
        }
        finally
        {
            Release(numberParagraphRange);
            Release(numberParagraph);
            Release(numberParagraphs);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(maths);
            Release(fields);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void InsertVisibleEquationNumber(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId,
        string targetBookmarkName,
        Table? knownNumberedTable = null,
        bool useConversionSafeVisibleNumber = false,
        bool deferFieldUpdate = false)
    {
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        var performanceWatch = tracePerformance
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        var performanceCheckpoint = 0L;
        void TraceStage(string stage)
        {
            if (performanceWatch is null) return;
            var elapsed = performanceWatch.ElapsedMilliseconds;
            TraceNumberingPerformance(
                $"[perf] VisibleRef.{stage}: +{elapsed - performanceCheckpoint}ms ({elapsed}ms)");
            performanceCheckpoint = elapsed;
        }
        Range? scaffoldRange = null;
        Range? fieldRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? fieldResult = null;
        Range? bookmarkRange = null;
        Bookmarks? bookmarks = null;
        try
        {
            var tableLayout = knownNumberedTable is not null
                || IsNumberedEquationTable(formulaRange);
            var suffixStart = PrepareEquationNumberInsertionPosition(
                formulaRange,
                knownNumberedTable);
            // Normal VisualTeX insertion/edit keeps the long-standing stable
            // "()" scaffold. Only MathType/OMML -> VisualTeX conversion uses the
            // isolated REF-first path because migrated MathType paragraphs can
            // make Word consume a pre-seeded parenthesis while the field is born.
            var scaffold = tableLayout
                ? useConversionSafeVisibleNumber
                    ? (Text: string.Empty, FieldOffset: 0)
                    : (Text: "()", FieldOffset: 1)
                : EquationNumberScaffold();
            object suffixStartObject = suffixStart;
            object suffixEndObject = suffixStart;
            scaffoldRange = document.Range(ref suffixStartObject, ref suffixEndObject);
            if (scaffold.Text.Length > 0)
                scaffoldRange.Text = scaffold.Text;
            TraceStage("scaffold");

            var fieldStart = suffixStart + scaffold.FieldOffset;
            object fieldStartObject = fieldStart;
            object fieldEndObject = fieldStart;
            fieldRange = document.Range(ref fieldStartObject, ref fieldEndObject);
            // Keep field creation local to the number cell. document.Fields.Add
            // forces Word to maintain the complete story field collection even
            // though this REF belongs to one collapsed insertion range.
            fields = fieldRange.Fields;
            object fieldType = WdFieldEmpty;
            object fieldCode = $"REF {targetBookmarkName} \\h";
            object preserveFormatting = true;
            field = fields.Add(
                fieldRange,
                ref fieldType,
                ref fieldCode,
                ref preserveFormatting);
            TraceStage("add-ref");
            // Normal single-formula insertion resolves the REF immediately. Batch
            // format conversion keeps ScreenUpdating disabled and defers the field
            // result refresh until all 35 numbered targets exist; updating each REF
            // while the document field collection is still growing is O(N^2).
            if (!deferFieldUpdate)
                NormalizeReferenceResult(field, formulaFontSizePoints);
            TraceStage(deferFieldUpdate ? "defer-update-ref" : "update-ref");
            fieldResult = field.Result;

            // Do not rely on scaffoldRange.End. Word 2013/2016 perpetual,
            // Microsoft 365 and compatibility mode expand a Range around an
            // inserted field differently. Resolve the actual closing parenthesis
            // from the document text so the bookmark always contains the tab,
            // both brackets and the complete REF field.
            bookmarkRange = tableLayout && useConversionSafeVisibleNumber
                ? CompleteTableEquationNumberLabelAfterField(
                    document,
                    formulaRange,
                    field,
                    knownNumberedTable)
                : ResolveEquationNumberLabelRangeFast(
                    document,
                    formulaRange,
                    suffixStart,
                    scaffoldRange,
                    fieldResult,
                    tableLayout,
                    knownNumberedTable);
            bookmarks = document.Bookmarks;
            bookmarks.Add(EquationBookmarkName(formulaId), bookmarkRange);
            if (tableLayout)
            {
                ParagraphFormat? format = null;
                try
                {
                    format = bookmarkRange.ParagraphFormat;
                    if (format.Alignment != WdParagraphAlignment.wdAlignParagraphRight)
                        format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                }
                finally { Release(format); }
            }
            TraceStage("bookmark-format");
            AlignEquationNumberVertically(
                bookmarkRange,
                tableLayout ? 0f : formulaHeightPoints,
                formulaFontSizePoints);
            TraceStage("vertical-align");
        }
        finally
        {
            Release(bookmarks);
            Release(bookmarkRange);
            Release(fieldResult);
            Release(field);
            Release(fields);
            Release(fieldRange);
            Release(scaffoldRange);
        }
    }

    private static Range CompleteTableEquationNumberLabelAfterField(
        Document document,
        Range formulaRange,
        Field field,
        Table? knownNumberedTable)
    {
        Table? discoveredTable = null;
        Cell? cell = null;
        Range? cellRange = null;
        Range? prefix = null;
        Range? suffix = null;
        Range? candidate = null;
        try
        {
            discoveredTable = knownNumberedTable is null
                && formulaRange.Tables.Count > 0
                    ? formulaRange.Tables[1]
                    : null;
            var table = knownNumberedTable ?? discoveredTable
                ?? throw new InvalidOperationException(
                    "VisualTeX could not resolve the numbered equation table after creating its REF field.");
            cell = table.Cell(1, 3);

            // Prefix: insert at the cell's first editable position, which is now
            // guaranteed to be before the already-materialized REF field.
            cellRange = cell.Range;
            prefix = document.Range(cellRange.Start, cellRange.Start);
            prefix.Text = "(";
            Release(cellRange);
            cellRange = null;

            // Prefix insertion shifts the REF field. Re-read its live Result and
            // insert ')' *after the field-end control*, not at Result.End itself.
            // Inserting at Result.End is exactly what made Word absorb ')' into
            // Field.Result; the next field update then erased the parenthesis.
            Release(cellRange);
            cellRange = null;
            Release(candidate);
            candidate = null;
            Range? liveResult = null;
            try
            {
                liveResult = field.Result;
                cellRange = cell.Range;
                var suffixPosition = Math.Min(
                    Math.Max(cellRange.Start + 1, liveResult.End + 1),
                    Math.Max(cellRange.Start + 1, cellRange.End - 1));
                suffix = document.Range(suffixPosition, suffixPosition);
                suffix.Text = ")";

                Release(liveResult);
                liveResult = field.Result;
                if ((liveResult.Text ?? string.Empty).EndsWith(")", StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "Word absorbed the VisualTeX closing parenthesis into the REF field result.");
            }
            finally { Release(liveResult); }

            Release(cellRange);
            cellRange = cell.Range;
            var editableEnd = Math.Max(cellRange.Start, cellRange.End - 1);
            candidate = document.Range(cellRange.Start, editableEnd);
            var text = candidate.Text ?? string.Empty;
            if (!text.StartsWith("(", StringComparison.Ordinal)
                || !text.EndsWith(")", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "Word did not materialize the complete VisualTeX equation-number label after the REF field was finalized.");
            var result = candidate;
            candidate = null;
            return result;
        }
        finally
        {
            Release(candidate);
            Release(suffix);
            Release(prefix);
            Release(cellRange);
            Release(cell);
            Release(discoveredTable);
        }
    }

    private static Range ResolveEquationNumberLabelRangeFast(
        Document document,
        Range formulaRange,
        int labelStart,
        Range scaffoldRange,
        Range fieldResult,
        bool tableLayout,
        Table? knownNumberedTable)
    {
        Range? candidate = null;
        try
        {
            var contentEnd = document.Content.End;
            var expectedEnd = Math.Min(contentEnd, fieldResult.End + 1);
            if (expectedEnd > labelStart)
            {
                candidate = document.Range(labelStart, expectedEnd);
                var text = candidate.Text ?? string.Empty;
                var expectedPrefix = tableLayout ? "(" : "\t(";
                if (text.StartsWith(expectedPrefix, StringComparison.Ordinal)
                    && text.EndsWith(")", StringComparison.Ordinal))
                {
                    var result = candidate;
                    candidate = null;
                    return result;
                }
            }
        }
        finally { Release(candidate); }

        // Compatibility fallback for older Word builds that report a field
        // Result.End outside the scaffold in a different way.
        return ResolveEquationNumberLabelRange(
            document,
            formulaRange,
            labelStart,
            scaffoldRange,
            fieldResult,
            tableLayout,
            knownNumberedTable);
    }

    private static Range ResolveEquationNumberLabelRange(
        Document document,
        Range formulaRange,
        int labelStart,
        Range scaffoldRange,
        Range fieldResult,
        bool tableLayout,
        Table? knownNumberedTable)
    {
        Range? character = null;
        Range? candidate = null;
        try
        {
            var searchEnd = Math.Min(document.Content.End, labelStart + 512);
            for (var position = labelStart + 1; position < searchEnd; position++)
            {
                object characterStart = position;
                object characterEnd = position + 1;
                character = document.Range(ref characterStart, ref characterEnd);
                if (!string.Equals(character.Text, ")", StringComparison.Ordinal))
                {
                    Release(character);
                    character = null;
                    continue;
                }

                object candidateStart = labelStart;
                object candidateEnd = character.End;
                candidate = document.Range(ref candidateStart, ref candidateEnd);
                var text = candidate.Text ?? string.Empty;
                var expectedPrefix = tableLayout ? "(" : "\t(";
                if (text.StartsWith(expectedPrefix, StringComparison.Ordinal)
                    && text.EndsWith(")", StringComparison.Ordinal))
                {
                    var result = candidate;
                    candidate = null;
                    return result;
                }
                Release(candidate);
                candidate = null;
                Release(character);
                character = null;
            }

            object fallbackStart = labelStart;
            object fallbackEnd = scaffoldRange.End;
            candidate = document.Range(ref fallbackStart, ref fallbackEnd);
            var fallbackText = candidate.Text ?? string.Empty;
            var fallbackPrefix = tableLayout ? "(" : "\t(";
            if (fallbackText.StartsWith(fallbackPrefix, StringComparison.Ordinal)
                && fallbackText.EndsWith(")", StringComparison.Ordinal))
            {
                var fallback = candidate;
                candidate = null;
                return fallback;
            }

            // Some Word builds consume the literal ')' that was placed after the
            // zero-width REF insertion point while materializing/updating the
            // field.  A numbered equation table reserves cell (1,3) exclusively
            // for VisualTeX, so repair only that missing suffix locally instead of
            // failing the whole MathType/OMML conversion or running a document-wide
            // rebuild.  The opening parenthesis and field must already be intact.
            Release(candidate);
            candidate = null;
            if (tableLayout)
            {
                var repaired = TryRepairEquationNumberClosingParenthesis(
                    document,
                    formulaRange,
                    labelStart,
                    fieldResult,
                    knownNumberedTable);
                if (repaired is not null) return repaired;
            }

            var diagnostic = string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                ? DescribeEquationNumberLabelFailure(
                    document,
                    formulaRange,
                    labelStart,
                    scaffoldRange,
                    fieldResult,
                    knownNumberedTable)
                : string.Empty;
            throw new InvalidOperationException(
                "Word did not preserve the complete VisualTeX equation-number label."
                + diagnostic);
        }
        finally
        {
            Release(candidate);
            Release(character);
        }
    }

    private static string DescribeEquationNumberLabelFailure(
        Document document,
        Range formulaRange,
        int labelStart,
        Range scaffoldRange,
        Range fieldResult,
        Table? knownNumberedTable)
    {
        Table? table = null;
        Cell? cell = null;
        Range? cellRange = null;
        Range? probe = null;
        try
        {
            table = knownNumberedTable is null
                && formulaRange.Tables.Count > 0
                    ? formulaRange.Tables[1]
                    : null;
            var owner = knownNumberedTable ?? table;
            if (owner is null)
            {
                return $" [labelStart={labelStart}; scaffold=[{scaffoldRange.Start},{scaffoldRange.End}]; fieldResult=[{fieldResult.Start},{fieldResult.End}]; no-number-table]";
            }
            cell = owner.Cell(1, 3);
            cellRange = cell.Range;
            var probeStart = Math.Max(cellRange.Start, Math.Min(labelStart, cellRange.End - 1));
            var probeEnd = Math.Min(cellRange.End, Math.Max(probeStart, cellRange.End));
            probe = document.Range(probeStart, probeEnd);
            var text = (probe.Text ?? string.Empty)
                .Replace("\r", "<P>")
                .Replace("\a", "<CELL>")
                .Replace("\t", "<TAB>")
                .Replace("\v", "<BR>");
            return $" [labelStart={labelStart}; scaffold=[{scaffoldRange.Start},{scaffoldRange.End}]; fieldResult=[{fieldResult.Start},{fieldResult.End}]; cell=[{cellRange.Start},{cellRange.End}]; text='{text}']";
        }
        catch (Exception error)
        {
            return $" [label diagnostic failed: {error.GetType().Name}: {error.Message}]";
        }
        finally
        {
            Release(probe);
            Release(cellRange);
            Release(cell);
            Release(table);
        }
    }

    private static Range? TryRepairEquationNumberClosingParenthesis(
        Document document,
        Range formulaRange,
        int labelStart,
        Range fieldResult,
        Table? knownNumberedTable)
    {
        Tables? tables = null;
        Table? table = null;
        Cell? cell = null;
        Range? cellRange = null;
        Range? prefix = null;
        Range? insertion = null;
        Range? candidate = null;
        try
        {
            if (knownNumberedTable is null)
            {
                tables = formulaRange.Tables;
                if (tables.Count == 0) return null;
                table = tables[1];
            }
            var owner = knownNumberedTable ?? table;
            if (owner is null) return null;
            cell = owner.Cell(1, 3);
            cellRange = cell.Range;
            if (labelStart < cellRange.Start || labelStart >= cellRange.End)
                return null;

            var prefixEnd = Math.Min(cellRange.End, labelStart + 1);
            prefix = document.Range(labelStart, prefixEnd);
            if (!string.Equals(prefix.Text, "(", StringComparison.Ordinal))
                return null;

            // field.Result excludes Word's field-end control character; +1 is the
            // first legal character position after the complete REF field. Clamp
            // it to this reserved number cell so a malformed field can never make
            // the repair touch adjacent document content.
            var editableEnd = Math.Max(cellRange.Start, cellRange.End - 1);
            var insertionPosition = Math.Max(labelStart + 1, fieldResult.End + 1);
            insertionPosition = Math.Min(insertionPosition, editableEnd);
            if (insertionPosition <= labelStart) return null;

            insertion = document.Range(insertionPosition, insertionPosition);
            insertion.Text = ")";
            candidate = document.Range(labelStart, insertionPosition + 1);
            var text = candidate.Text ?? string.Empty;
            if (!text.StartsWith("(", StringComparison.Ordinal)
                || !text.EndsWith(")", StringComparison.Ordinal))
                return null;
            var result = candidate;
            candidate = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(candidate);
            Release(insertion);
            Release(prefix);
            Release(cellRange);
            Release(cell);
            Release(table);
            Release(tables);
        }
    }

    private static int PrepareEquationNumberInsertionPosition(
        Range formulaRange,
        Table? knownNumberedTable = null)
    {
        if (knownNumberedTable is not null || IsNumberedEquationTable(formulaRange))
        {
            Table? table = null;
            Cell? cell = null;
            Range? cellRange = null;
            Range? editableRange = null;
            try
            {
                table = knownNumberedTable is null
                    ? formulaRange.Tables[1]
                    : null;
                cell = (knownNumberedTable ?? table!).Cell(1, 3);
                cellRange = cell.Range;
                // This cell is reserved exclusively for the generated number.
                // Word can leave an empty paragraph behind when a REF field is
                // removed/reconciled. Centering that empty paragraph together
                // with the new number pushes the visible number downward.
                // Clear everything except the structural cell mark, which
                // normalizes the cell to exactly one paragraph, then insert at
                // the beginning of that paragraph.
                editableRange = cellRange.Duplicate;
                editableRange.End = Math.Max(
                    editableRange.Start,
                    editableRange.End - 1);
                editableRange.Text = string.Empty;
                return cellRange.Start;
            }
            finally
            {
                Release(editableRange);
                Release(cellRange);
                Release(cell);
                Release(table);
            }
        }
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return Math.Max(paragraphRange.Start, paragraphRange.End - 1);
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void RemoveVisibleEquationNumber(Document document, string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Range? trailing = null;
        OMaths? maths = null;
        OMath? containingMath = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            range = bookmark.Range;
            var start = range.Start;
            var text = range.Text ?? string.Empty;
            try
            {
                maths = range.OMaths;
                if (maths.Count > 0) containingMath = maths[1];
            }
            catch { }
            range.Delete();

            // Legacy OMML numbering bookmarks stopped immediately before the
            // closing parenthesis. Remove that orphan as part of migration.
            if (!text.EndsWith(")", StringComparison.Ordinal)
                && start < document.Content.End)
            {
                object trailingStart = start;
                object trailingEnd = Math.Min(document.Content.End, start + 1);
                trailing = document.Range(ref trailingStart, ref trailingEnd);
                if (string.Equals(trailing.Text, ")", StringComparison.Ordinal))
                    trailing.Delete();
            }
            try { containingMath?.BuildUp(); } catch { }
        }
        finally
        {
            Release(containingMath);
            Release(maths);
            Release(trailing);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void RemoveNativeCaption(
        Document document,
        string formulaId,
        bool preserveParagraphSeparator = false)
    {
        DeleteBookmarkOnly(document, NativeNumberBookmarkName(formulaId));
        DeleteBookmarkedRangeAndContainingFrame(
            document,
            NativeCaptionBookmarkName(formulaId),
            preserveParagraphSeparator);
    }

    private static void DeleteBookmarkedRangeAndContainingFrame(
        Document document,
        string name,
        bool preserveParagraphSeparator = false)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Frames? frames = null;
        Frame? frame = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            range = bookmark.Range;

            // VisualTeX keeps the native SEQ caption in a 0.1 pt clipping Frame
            // at the bottom-right page boundary so Word's native cross-reference
            // dialog can still discover the target. Deleting only the bookmarked
            // caption text leaves that empty Frame alive. A later numbered-formula
            // -> LaTeX conversion can then insert the restored source at the same
            // document position and Word silently adopts the new source into the
            // surviving clipping Frame, making the LaTeX text effectively invisible.
            // Remove the Frame formatting first; Frame.Delete preserves its text.
            try
            {
                frames = range.Frames;
                if (frames.Count > 0)
                {
                    frame = frames[1];
                    frame.Delete();
                    Release(frame);
                    frame = null;
                    Release(frames);
                    frames = null;
                    // Frame.Delete can update the bookmarked range object. Re-read
                    // it before deleting the native caption contents.
                    Release(range);
                    range = bookmark.Range;
                }
            }
            catch
            {
                // If Word refuses to expose a Frame, the bookmark deletion below
                // still removes the caption. The conversion layer separately guards
                // against any surviving clipping Frame adopting the LaTeX source.
            }

            if (preserveParagraphSeparator
                && range.End > range.Start
                && (range.Text ?? string.Empty).EndsWith("\r", StringComparison.Ordinal))
            {
                // The hidden caption paragraph is also the separator between two
                // adjacent numbered tables. Removing its final paragraph mark makes
                // Word merge those tables immediately. A numbered OLE -> OMML
                // replacement rebuilds the caption in the same table, so remove the
                // old SEQ/content while preserving that separator until the rebuild.
                var contentStart = range.Start;
                var contentEnd = range.End - 1;
                bookmark.Delete();
                if (contentEnd > contentStart)
                {
                    Range? captionContents = null;
                    try
                    {
                        captionContents = document.Range(contentStart, contentEnd);
                        captionContents.Delete();
                    }
                    finally { Release(captionContents); }
                }
            }
            else
            {
                range.Delete();
            }
        }
        finally
        {
            Release(frame);
            Release(frames);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void RemoveOrphanEquationArtifacts(
        Document document,
        ISet<string> numberedFormulaIds)
    {
        RemoveOrphanBookmarks(
            document,
            EquationBookmarkPrefix,
            numberedFormulaIds,
            deleteRange: true);
        RemoveOrphanBookmarks(
            document,
            NativeCaptionBookmarkPrefix,
            numberedFormulaIds,
            deleteRange: true);
        RemoveOrphanBookmarks(
            document,
            NativeNumberBookmarkPrefix,
            numberedFormulaIds,
            deleteRange: false);
    }

    private static void RemoveOrphanBookmarks(
        Document document,
        string prefix,
        ISet<string> activeFormulaIds,
        bool deleteRange)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = bookmarks.Count; index >= 1; index--)
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryFormulaIdFromBookmark(bookmark.Name, prefix, out var formulaId)
                        || activeFormulaIds.Contains(formulaId))
                        continue;
                    if (deleteRange)
                    {
                        range = bookmark.Range;
                        range.Delete();
                    }
                    else
                    {
                        bookmark.Delete();
                    }
                }
                finally
                {
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }
    }

    private static void UpdateEquationNumberFields(
        Document document,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId)
    {
        UpdateFieldInBookmark(
            document,
            EquationBookmarkName(formulaId),
            code => IsReferenceToBookmark(code, NativeNumberBookmarkName(formulaId)));

        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            var visibleName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(visibleName)) return;
            bookmark = bookmarks[visibleName];
            range = bookmark.Range;
            // A numbered table centers both cells vertically. Applying the
            // legacy height-derived baseline shift as well makes OLE and OMML
            // numbers disagree because their measured heights differ.
            AlignEquationNumberVertically(
                range,
                IsNumberedEquationTable(range) ? 0f : formulaHeightPoints,
                formulaFontSizePoints);
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static IReadOnlyDictionary<string, FormulaDocumentLocation>
        CaptureNumberedFormulaLocations(
            Document document,
            InlineShapes inlineShapes,
            IReadOnlyList<string> ommlFormulaIds)
    {
        var result = new Dictionary<string, FormulaDocumentLocation>(
            StringComparer.OrdinalIgnoreCase);
        for (var index = 1; index <= inlineShapes.Count; index++)
        {
            InlineShape? shape = null;
            Range? range = null;
            try
            {
                shape = inlineShapes[index];
                var metadata = ReadMetadata(shape);
                if (metadata?.DisplayMode != "block" || !metadata.Numbered) continue;
                range = shape.Range;
                result[metadata.FormulaId] = new FormulaDocumentLocation(
                    range.Start,
                    range.End);
            }
            finally
            {
                Release(range);
                Release(shape);
            }
        }

        foreach (var formulaId in ommlFormulaIds)
        {
            Bookmark? bookmark = null;
            Range? range = null;
            try
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                if (metadata?.DisplayMode != "block" || !metadata.Numbered) continue;
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                if (bookmark is null) continue;
                range = WordOmmlFormulaStore.GetEquationRange(bookmark);
                result[formulaId] = new FormulaDocumentLocation(
                    range.Start,
                    range.End);
            }
            finally
            {
                Release(range);
                Release(bookmark);
            }
        }
        return result;
    }

    private static void RepairSharedNativeCaptionArtifacts(
        Document document,
        IReadOnlyDictionary<string, FormulaDocumentLocation> formulaLocations)
    {
        var ownersByArtifactRange = new Dictionary<string, HashSet<string>>(
            StringComparer.Ordinal);
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    bookmark = bookmarks[index];
                    string artifactKind;
                    string formulaId;
                    if (TryFormulaIdFromBookmark(
                            bookmark.Name,
                            NativeNumberBookmarkPrefix,
                            out formulaId))
                    {
                        artifactKind = "number";
                    }
                    else if (TryFormulaIdFromBookmark(
                                 bookmark.Name,
                                 NativeCaptionBookmarkPrefix,
                                 out formulaId))
                    {
                        artifactKind = "caption";
                    }
                    else
                    {
                        continue;
                    }
                    if (!formulaLocations.ContainsKey(formulaId)) continue;

                    range = bookmark.Range;
                    var key = artifactKind + ":" + range.Start + ":" + range.End;
                    if (!ownersByArtifactRange.TryGetValue(key, out var owners))
                    {
                        owners = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        ownersByArtifactRange[key] = owners;
                    }
                    owners.Add(formulaId);
                }
                finally
                {
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }

        var conflictedFormulaIds = ownersByArtifactRange.Values
            .Where(owners => owners.Count > 1)
            .SelectMany(owners => owners)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (conflictedFormulaIds.Count == 0) return;

        // A shared caption paragraph cannot safely keep one "owner": deleting or
        // rewriting that paragraph invalidates every overlapping VTEqCap/VTEqNum
        // bookmark. Remove the complete numbering scaffold for every participant
        // first, then the normal reconciliation loop rebuilds one independent
        // visible REF and one independent native SEQ target per formula.
        foreach (var formulaId in conflictedFormulaIds
                     .OrderByDescending(id => formulaLocations[id].Start))
        {
            RemoveVisibleEquationNumber(document, formulaId);
            RemoveNativeCaption(document, formulaId);
        }
    }

    private static void RebuildNativeNumberBookmarksFromCaptions(
        Document document,
        ISet<string> numberedFormulaIds)
    {
        var nativeSequenceName = GetNativeEquationSequenceName(document);
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var formulaId in numberedFormulaIds)
            {
                Bookmark? captionBookmark = null;
                Range? captionRange = null;
                Field? sequenceField = null;
                Range? fieldResult = null;
                Range? completeNumberRange = null;
                Bookmark? rebuiltNumberBookmark = null;
                try
                {
                    var captionName = NativeCaptionBookmarkName(formulaId);
                    if (!bookmarks.Exists(captionName))
                        throw new InvalidOperationException(
                            $"VisualTeX formula {formulaId} has no native caption after reconciliation.");
                    captionBookmark = bookmarks[captionName];
                    captionRange = captionBookmark.Range;
                    sequenceField = FindNativeEquationFieldInCaption(
                        captionRange,
                        nativeSequenceName);
                    if (sequenceField is null)
                        throw new InvalidOperationException(
                            $"VisualTeX formula {formulaId} has no native SEQ field inside its caption.");
                    fieldResult = sequenceField.Result;
                    // After heading-aware renumbering, the chapter/section
                    // prefix is ordinary text immediately before the SEQ field.
                    // Rebuilding from Field.Result alone silently changes a
                    // target such as "1-1" back to "1". Preserve the complete
                    // visible caption number from the paragraph start through
                    // the field result; continuous numbering naturally has an
                    // empty prefix and uses the same range.
                    completeNumberRange = document.Range(
                        captionRange.Start,
                        fieldResult.End);

                    var numberName = NativeNumberBookmarkName(formulaId);
                    if (bookmarks.Exists(numberName))
                    {
                        Bookmark? existing = null;
                        try
                        {
                            existing = bookmarks[numberName];
                            existing.Delete();
                        }
                        finally { Release(existing); }
                    }
                    rebuiltNumberBookmark = bookmarks.Add(
                        numberName,
                        completeNumberRange);
                }
                finally
                {
                    Release(rebuiltNumberBookmark);
                    Release(completeNumberRange);
                    Release(fieldResult);
                    Release(sequenceField);
                    Release(captionRange);
                    Release(captionBookmark);
                }
            }
        }
        finally { Release(bookmarks); }
    }

    private static Field? FindNativeEquationFieldInCaption(
        Range captionRange,
        string nativeSequenceName)
    {
        Fields? fields = null;
        try
        {
            fields = captionRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(
                            code.Text,
                            nativeSequenceName))
                        continue;
                    var found = field;
                    field = null;
                    return found;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            return null;
        }
        finally { Release(fields); }
    }

    private sealed class FormulaDocumentLocation
    {
        public FormulaDocumentLocation(int start, int end)
        {
            Start = start;
            End = end;
        }

        public int Start { get; }
        public int End { get; }
    }

    private sealed class NativeEquationCaptionEntry
    {
        public NativeEquationCaptionEntry(
            string formulaId,
            int position,
            string numberText)
        {
            FormulaId = formulaId;
            Position = position;
            NumberText = numberText;
        }

        public string FormulaId { get; }
        public int Position { get; }
        public string NumberText { get; }
    }

    private sealed class HeadingNumberAnchor
    {
        public HeadingNumberAnchor(int position, string numberText)
        {
            Position = position;
            NumberText = numberText;
        }

        public int Position { get; }
        public string NumberText { get; }
    }

    private sealed class HeadingParagraphCandidate
    {
        public HeadingParagraphCandidate(
            int position,
            int outlineLevel,
            string listNumber,
            string textNumber)
        {
            Position = position;
            OutlineLevel = outlineLevel;
            ListNumber = listNumber;
            TextNumber = textNumber;
        }

        public int Position { get; }
        public int OutlineLevel { get; }
        public string ListNumber { get; }
        public string TextNumber { get; }
        public string ExplicitNumber =>
            !string.IsNullOrWhiteSpace(ListNumber) ? ListNumber : TextNumber;
    }

    private static void UpdateNativeEquationSequenceFields(Document document)
    {
        var nativeSequenceName = GetNativeEquationSequenceName(document);
        var captions = GetNativeEquationCaptionEntries(document, nativeSequenceName);
        UpdateNativeEquationSequenceFields(
            document,
            nativeSequenceName,
            captions,
            formatOnly: false);
        WriteNativeEquationTailFormulaId(document, captions.LastOrDefault()?.FormulaId);
    }

    private static bool TryUpdateAppendedNativeEquationSequenceField(
        Document document,
        string formulaId,
        out bool numberChanged)
    {
        numberChanged = false;
        if (!TryReadNativeEquationTailFormulaId(document, out var tailFormulaId)
            || string.Equals(tailFormulaId, formulaId, StringComparison.OrdinalIgnoreCase))
            return false;

        Bookmarks? bookmarks = null;
        Bookmark? tailBookmark = null;
        Bookmark? currentBookmark = null;
        Range? tailRange = null;
        Range? currentRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var tailNumberName = NativeNumberBookmarkName(tailFormulaId);
            var currentNumberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(tailNumberName)
                || !bookmarks.Exists(currentNumberName)
                || !bookmarks.Exists(NativeCaptionBookmarkName(tailFormulaId))
                || !bookmarks.Exists(NativeCaptionBookmarkName(formulaId)))
                return false;

            tailBookmark = bookmarks[tailNumberName];
            currentBookmark = bookmarks[currentNumberName];
            tailRange = tailBookmark.Range;
            currentRange = currentBookmark.Range;
            if (currentRange.Start <= tailRange.Start)
                return false;

            var format = ReadEquationNumberFormat(document);
            if (!TryParseNativeEquationNumber(
                    tailRange.Text,
                    format,
                    out var prefix,
                    out var previousOrdinal))
                return false;

            var ordinal = previousOrdinal + 1;
            if (format.UsesHeading)
            {
                if (!TryResolveAppendHeadingPrefix(
                        document,
                        tailRange.End,
                        currentRange.Start,
                        format,
                        prefix,
                        out prefix,
                        out var startsNewScope))
                    return false;
                if (startsNewScope) ordinal = 1;
            }

            var expectedNumberText = prefix
                + ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
            var currentNumberText = NormalizeNativeEquationNumberText(currentRange.Text);
            if (!string.Equals(currentNumberText, expectedNumberText, StringComparison.Ordinal))
            {
                UpdateNativeEquationCaptionNumber(
                    document,
                    formulaId,
                    GetNativeEquationSequenceName(document),
                    ordinal,
                    prefix,
                    formatOnly: false,
                    cleanupLegacyFrames: false);
                numberChanged = true;
            }

            WriteNativeEquationTailFormulaId(document, formulaId);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(currentRange);
            Release(tailRange);
            Release(currentBookmark);
            Release(tailBookmark);
            Release(bookmarks);
        }
    }

    private static bool TryParseNativeEquationNumber(
        string? value,
        EquationNumberFormat format,
        out string prefix,
        out int ordinal)
    {
        prefix = string.Empty;
        ordinal = 0;
        var text = NormalizeNativeEquationNumberText(value);
        if (!format.UsesHeading)
            return int.TryParse(
                    text,
                    System.Globalization.NumberStyles.None,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out ordinal)
                && ordinal > 0;

        var separatorIndex = text.LastIndexOf(format.Separator, StringComparison.Ordinal);
        if (separatorIndex <= 0) return false;
        var ordinalText = text.Substring(separatorIndex + format.Separator.Length);
        if (!int.TryParse(
                ordinalText,
                System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture,
                out ordinal)
            || ordinal <= 0)
            return false;
        prefix = text.Substring(0, separatorIndex + format.Separator.Length);
        return true;
    }

    private static bool TryResolveAppendHeadingPrefix(
        Document document,
        int previousNumberEnd,
        int currentNumberStart,
        EquationNumberFormat format,
        string previousPrefix,
        out string prefix,
        out bool startsNewScope)
    {
        prefix = previousPrefix;
        startsNewScope = false;
        if (!format.UsesHeading || currentNumberStart <= previousNumberEnd)
            return true;

        Range? between = null;
        Paragraphs? paragraphs = null;
        try
        {
            between = document.Range(previousNumberEnd, currentNumberStart);
            paragraphs = between.Paragraphs;
            var sawHeadingCandidate = false;
            string? lastExplicitNumber = null;
            var lastExplicitLevel = 0;
            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Paragraph? paragraph = null;
                Range? range = null;
                ListFormat? listFormat = null;
                Frames? frames = null;
                try
                {
                    paragraph = paragraphs[index];
                    var outlineLevel = (int)paragraph.OutlineLevel;
                    if (outlineLevel < 1 || outlineLevel > format.HeadingLevel) continue;
                    range = paragraph.Range;
                    if ((bool)range.get_Information(WdInformation.wdWithInTable)) continue;
                    try
                    {
                        frames = range.Frames;
                        if (frames.Count > 0) continue;
                    }
                    catch { }

                    sawHeadingCandidate = true;
                    var listNumber = string.Empty;
                    try
                    {
                        listFormat = range.ListFormat;
                        listNumber = NormalizeHeadingNumberText(listFormat.ListString);
                    }
                    catch { }
                    var explicitNumber = !string.IsNullOrWhiteSpace(listNumber)
                        ? listNumber
                        : ParseHeadingNumberFromText(range.Text, outlineLevel);
                    if (string.IsNullOrWhiteSpace(explicitNumber)) continue;
                    lastExplicitNumber = explicitNumber;
                    lastExplicitLevel = outlineLevel;
                }
                finally
                {
                    Release(frames);
                    Release(listFormat);
                    Release(range);
                    Release(paragraph);
                }
            }

            if (!sawHeadingCandidate) return true;
            // If the new interval contains only unnumbered headings, the full
            // document pass is required to know whether they are real synthesized
            // chapter counters or titles skipped by an explicitly numbered scheme.
            if (string.IsNullOrWhiteSpace(lastExplicitNumber)) return false;
            if (lastExplicitLevel < format.HeadingLevel)
            {
                lastExplicitNumber += string.Concat(
                    Enumerable.Repeat(".0", format.HeadingLevel - lastExplicitLevel));
            }
            prefix = lastExplicitNumber + format.Separator;
            startsNewScope = true;
            return true;
        }
        finally
        {
            Release(paragraphs);
            Release(between);
        }
    }

    private static bool TryUpdateInsertedContinuousEquationFieldRanges(
        Document document,
        string formulaId)
    {
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        var rangeWatch = tracePerformance
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        long rangeCheckpoint = 0;
        void TraceRangeStage(string stage)
        {
            if (rangeWatch is null) return;
            var elapsed = rangeWatch.ElapsedMilliseconds;
            Console.WriteLine(
                $"      [perf] continuous-range.{stage}: +{elapsed - rangeCheckpoint}ms ({elapsed}ms)");
            rangeCheckpoint = elapsed;
        }

        var format = ReadEquationNumberFormat(document);
        if (format.UsesHeading) return false;

        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Range? captionRange = null;
        Range? content = null;
        Range? suffixRange = null;
        Fields? suffixFields = null;
        Range? prefixRange = null;
        Fields? prefixFields = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            var numberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName) || !bookmarks.Exists(numberName))
                return false;

            captionBookmark = bookmarks[captionName];
            captionRange = captionBookmark.Range;
            content = document.Content;
            if (captionRange.End >= content.End) return false;

            // A middle insertion only changes the sequence state after its native
            // caption. Let Word update that suffix twice: pass one recalculates the
            // later SEQ targets, pass two lets visible/body REF fields observe the
            // new target results. This is the same Fields.Update semantics used by
            // full reconciliation, but scoped to the affected half of the story.
            suffixRange = document.Range(captionRange.End, content.End);
            suffixFields = suffixRange.Fields;
            if (suffixFields.Count == 0) return false;
            TraceRangeStage("prepare-suffix");
            suffixFields.Update();
            TraceRangeStage("suffix-pass-1");
            suffixFields.Update();
            TraceRangeStage("suffix-pass-2");

            // References located before the insertion point can target a formula
            // whose number just shifted. Refresh only that prefix once after the
            // suffix SEQ values are stable. Previous formulas' own SEQ fields are
            // unchanged, and Word simply reproduces the same results for them.
            if (captionRange.Start > content.Start)
            {
                prefixRange = document.Range(content.Start, captionRange.Start);
                prefixFields = prefixRange.Fields;
                if (prefixFields.Count > 0) prefixFields.Update();
            }
            TraceRangeStage("prefix-pass");
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(prefixFields);
            Release(prefixRange);
            Release(suffixFields);
            Release(suffixRange);
            Release(content);
            Release(captionRange);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    private static bool TryUpdateInsertedContinuousEquationSequenceSuffix(
        Document document,
        string formulaId,
        out Dictionary<string, string> changedFormulaNumbers,
        out int referencesAlreadyUpdatedFrom)
    {
        changedFormulaNumbers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        referencesAlreadyUpdatedFrom = -1;
        var format = ReadEquationNumberFormat(document);
        if (format.UsesHeading) return false;

        Bookmarks? bookmarks = null;
        Bookmark? currentBookmark = null;
        Range? currentRange = null;
        Range? content = null;
        Range? suffixRange = null;
        Bookmarks? suffixBookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            var currentNumberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(currentNumberName)
                || !bookmarks.Exists(NativeCaptionBookmarkName(formulaId)))
                return false;
            currentBookmark = bookmarks[currentNumberName];
            currentRange = currentBookmark.Range;
            if (!TryParseNativeEquationNumber(
                    currentRange.Text,
                    format,
                    out _,
                    out var currentOrdinal))
                return false;

            content = document.Content;
            var suffixEntries = new List<NativeEquationCaptionEntry>();
            if (currentRange.End < content.End)
            {
                suffixRange = document.Range(currentRange.End, content.End);
                suffixBookmarks = suffixRange.Bookmarks;
                for (var index = 1; index <= suffixBookmarks.Count; index++)
                {
                    Bookmark? bookmark = null;
                    Range? numberRange = null;
                    try
                    {
                        bookmark = suffixBookmarks[index];
                        if (!TryFormulaIdFromBookmark(
                                bookmark.Name,
                                NativeNumberBookmarkPrefix,
                                out var suffixFormulaId)
                            || string.Equals(
                                suffixFormulaId,
                                formulaId,
                                StringComparison.OrdinalIgnoreCase))
                            continue;
                        if (!bookmarks.Exists(NativeCaptionBookmarkName(suffixFormulaId)))
                            return false;
                        numberRange = bookmark.Range;
                        suffixEntries.Add(new NativeEquationCaptionEntry(
                            suffixFormulaId,
                            numberRange.Start,
                            numberRange.Text ?? string.Empty));
                    }
                    finally
                    {
                        Release(numberRange);
                        Release(bookmark);
                    }
                }
            }

            var orderedSuffix = suffixEntries
                .OrderBy(entry => entry.Position)
                .ToArray();
            for (var index = 0; index < orderedSuffix.Length; index++)
            {
                if (!TryParseNativeEquationNumber(
                        orderedSuffix[index].NumberText,
                        format,
                        out _,
                        out var existingOrdinal))
                    return false;
                var expectedOrdinal = currentOrdinal + index + 1;
                // A healthy continuous sequence after one middle insertion is
                // either still on the old ordinal (expected - 1) or has already
                // been refreshed by Word to the new ordinal. Anything else is a
                // structural anomaly and belongs on the conservative full path.
                if (existingOrdinal != expectedOrdinal
                    && existingOrdinal != expectedOrdinal - 1)
                    return false;
            }

            var nativeSequenceName = GetNativeEquationSequenceName(document);
            for (var index = 0; index < orderedSuffix.Length; index++)
            {
                var expectedOrdinal = currentOrdinal + index + 1;
                if (!TryParseNativeEquationNumber(
                        orderedSuffix[index].NumberText,
                        format,
                        out _,
                        out var existingOrdinal))
                    return false;
                if (existingOrdinal == expectedOrdinal - 1)
                {
                    changedFormulaNumbers[orderedSuffix[index].FormulaId] =
                        expectedOrdinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
            }

            // A SEQ field result must be recalculated by Word; replacing that
            // result text directly can destroy the field/bookmark relationship.
            // Use one native suffix update (with our OLE EMBED fields locked),
            // then repair only VisualTeX REF results directly. This keeps Word's
            // sequence engine authoritative while avoiding the previous second
            // full Fields.Update pass.
            var suffixHasOnlyGeneratedReferences = false;
            var usedSingleBatchUpdate = changedFormulaNumbers.Count > 0
                && suffixRange is not null
                && TryBatchUpdateHealthyVisualTeXSuffixFields(
                    suffixRange,
                    nativeSequenceName,
                    changedFormulaNumbers.Count,
                    out suffixHasOnlyGeneratedReferences);
            if (usedSingleBatchUpdate)
            {
                var patchedGeneratedReferences = true;
                foreach (var changed in changedFormulaNumbers)
                {
                    if (TryPatchVisibleEquationNumberResult(
                            bookmarks,
                            changed.Key,
                            changed.Value))
                        continue;
                    patchedGeneratedReferences = false;
                    break;
                }
                if (patchedGeneratedReferences && suffixHasOnlyGeneratedReferences)
                    referencesAlreadyUpdatedFrom = currentRange.End;
            }
            else
            {
                // Compatibility fallback for an older/protected Word field or a
                // suffix that contains unrelated fields. Only shifted captions are
                // updated individually and cross references use the normal path.
                for (var index = 0; index < orderedSuffix.Length; index++)
                {
                    var expectedOrdinal = currentOrdinal + index + 1;
                    if (!changedFormulaNumbers.ContainsKey(orderedSuffix[index].FormulaId))
                        continue;
                    if (!TryRefreshExistingContinuousCaptionField(
                            document,
                            bookmarks,
                            orderedSuffix[index].FormulaId,
                            nativeSequenceName,
                            expectedOrdinal))
                        return false;
                }
            }

            for (var index = 0; index < orderedSuffix.Length; index++)
            {
                var expectedOrdinal = currentOrdinal + index + 1;
                Bookmark? refreshedNumberBookmark = null;
                Range? refreshedNumberRange = null;
                try
                {
                    var numberName = NativeNumberBookmarkName(orderedSuffix[index].FormulaId);
                    if (!bookmarks.Exists(numberName)) return false;
                    refreshedNumberBookmark = bookmarks[numberName];
                    refreshedNumberRange = refreshedNumberBookmark.Range;
                    if (!TryParseNativeEquationNumber(
                            refreshedNumberRange.Text,
                            format,
                            out _,
                            out var refreshedOrdinal)
                        || refreshedOrdinal != expectedOrdinal)
                        return false;
                }
                finally
                {
                    Release(refreshedNumberRange);
                    Release(refreshedNumberBookmark);
                }
            }

            WriteNativeEquationTailFormulaId(
                document,
                orderedSuffix.LastOrDefault()?.FormulaId ?? formulaId);
            return true;
        }
        catch
        {
            changedFormulaNumbers.Clear();
            referencesAlreadyUpdatedFrom = -1;
            return false;
        }
        finally
        {
            Release(suffixBookmarks);
            Release(suffixRange);
            Release(content);
            Release(currentRange);
            Release(currentBookmark);
            Release(bookmarks);
        }
    }

    private static bool TryPatchVisibleEquationNumberResult(
        Bookmarks bookmarks,
        string formulaId,
        string expectedNumber)
    {
        Bookmark? bookmark = null;
        Range? range = null;
        Fields? fields = null;
        try
        {
            var bookmarkName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(bookmarkName)) return false;
            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range;
            fields = range.Fields;
            var targetBookmarkName = NativeNumberBookmarkName(formulaId);
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsReferenceToBookmark(code.Text, targetBookmarkName))
                        continue;
                    result = field.Result;
                    if (!string.Equals(
                            NormalizeNativeEquationNumberText(result.Text),
                            expectedNumber,
                            StringComparison.Ordinal))
                        result.Text = expectedNumber;
                    return true;
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
            return false;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(fields);
            Release(range);
            Release(bookmark);
        }
    }

    private static bool TryBatchUpdateHealthyVisualTeXSuffixFields(
        Range suffixRange,
        string nativeSequenceName,
        int changedFormulaCount,
        out bool suffixHasOnlyGeneratedReferences)
    {
        suffixHasOnlyGeneratedReferences = false;
        Fields? fields = null;
        var referenceFieldCount = 0;
        try
        {
            fields = suffixRange.Fields;
            if (fields.Count == 0) return false;

            // Full VisualTeX reconciliation has always used document.Fields.Update,
            // so updating ordinary Word fields is existing product semantics. The
            // middle-insert fast path narrows that same operation to the suffix
            // instead of paying hundreds of Lock/Unlock COM writes just to avoid
            // work the full command already performs.
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                try
                {
                    field = fields[index];
                    if (field.Type == WdFieldType.wdFieldRef)
                        referenceFieldCount++;
                }
                finally { Release(field); }
            }

            fields.Update();
            suffixHasOnlyGeneratedReferences = referenceFieldCount == changedFormulaCount;
            return true;
        }
        catch
        {
            suffixHasOnlyGeneratedReferences = false;
            return false;
        }
        finally { Release(fields); }
    }

    private static bool TryRefreshExistingContinuousCaptionField(
        Document document,
        Bookmarks bookmarks,
        string formulaId,
        string nativeSequenceName,
        int expectedOrdinal)
    {
        Bookmark? captionBookmark = null;
        Bookmark? numberBookmark = null;
        Range? captionRange = null;
        Range? numberRange = null;
        Field? field = null;
        Range? code = null;
        try
        {
            var captionName = NativeCaptionBookmarkName(formulaId);
            var numberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName) || !bookmarks.Exists(numberName))
                return false;

            captionBookmark = bookmarks[captionName];
            captionRange = captionBookmark.Range;
            field = FindNativeEquationFieldInRange(captionRange, nativeSequenceName);
            if (field is null) return false;

            code = field.Code;
            var codeText = code.Text ?? string.Empty;
            if (!IsNativeEquationSequenceFieldCode(codeText, nativeSequenceName))
                return false;

            // Older structural passes could freeze a continuous SEQ with \\r N.
            // Restore normal Word sequence semantics before the local update so
            // later insertions can again be recalculated without rewriting the
            // field or its bookmarks.
            if (Regex.IsMatch(
                    codeText,
                    @"\\r\s+\d+",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                code.Text = $" SEQ {nativeSequenceName} \\* ARABIC ";
            }
            field.Update();

            numberBookmark = bookmarks[numberName];
            numberRange = numberBookmark.Range;
            var expectedText = expectedOrdinal.ToString(
                System.Globalization.CultureInfo.InvariantCulture);
            return string.Equals(
                NormalizeNativeEquationNumberText(numberRange.Text),
                expectedText,
                StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(numberRange);
            Release(captionRange);
            Release(numberBookmark);
            Release(captionBookmark);
        }
    }

    private static Dictionary<string, string> UpdateNativeEquationSequenceFieldsIncremental(
        Document document)
    {
        var nativeSequenceName = GetNativeEquationSequenceName(document);
        var captions = GetNativeEquationCaptionEntries(document, nativeSequenceName);
        return UpdateNativeEquationSequenceFieldsIncremental(
            document,
            nativeSequenceName,
            captions);
    }

    private static Dictionary<string, string> UpdateNativeEquationSequenceFieldsIncremental(
        Document document,
        string nativeSequenceName,
        IReadOnlyList<NativeEquationCaptionEntry> captions)
    {
        return UpdateNativeEquationSequenceFieldsIncremental(
            document,
            nativeSequenceName,
            captions,
            ReadEquationNumberFormat(document));
    }

    private static Dictionary<string, string> UpdateNativeEquationSequenceFieldsIncremental(
        Document document,
        string nativeSequenceName,
        IReadOnlyList<NativeEquationCaptionEntry> captions,
        EquationNumberFormat format)
    {
        var headingAnchors = format.UsesHeading
            ? GetHeadingNumberAnchorsForFormatBatch(
                document,
                format.HeadingLevel,
                captions)
            : Array.Empty<HeadingNumberAnchor>();
        var ordinalByScope = new Dictionary<int, int>();
        var changedFormulaNumbers = new Dictionary<string, string>(
            StringComparer.OrdinalIgnoreCase);

        foreach (var caption in captions)
        {
            var scope = ResolveEquationNumberScope(
                caption.Position,
                format,
                headingAnchors);
            ordinalByScope.TryGetValue(scope.ScopePosition, out var localOrdinal);
            localOrdinal++;
            ordinalByScope[scope.ScopePosition] = localOrdinal;

            var expectedNumberText = scope.Prefix
                + localOrdinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (string.Equals(
                    NormalizeNativeEquationNumberText(caption.NumberText),
                    expectedNumberText,
                    StringComparison.Ordinal))
                continue;

            UpdateNativeEquationCaptionNumber(
                document,
                caption.FormulaId,
                nativeSequenceName,
                localOrdinal,
                scope.Prefix,
                formatOnly: false,
                cleanupLegacyFrames: false);
            changedFormulaNumbers[caption.FormulaId] = expectedNumberText;
        }

        WriteNativeEquationTailFormulaId(document, captions.LastOrDefault()?.FormulaId);
        return changedFormulaNumbers;
    }

    private static string NormalizeNativeEquationNumberText(string? value) =>
        (value ?? string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace("\a", string.Empty)
            .Trim();

    private static void UpdateNativeEquationSequenceFields(
        Document document,
        string nativeSequenceName,
        IReadOnlyList<NativeEquationCaptionEntry> captions,
        bool formatOnly)
    {
        var format = ReadEquationNumberFormat(document);
        var headingAnchors = format.UsesHeading
            ? GetHeadingNumberAnchors(document, format.HeadingLevel)
            : Array.Empty<HeadingNumberAnchor>();
        var ordinalByScope = new Dictionary<int, int>();

        foreach (var caption in captions)
        {
            var scope = ResolveEquationNumberScope(
                caption.Position,
                format,
                headingAnchors);
            ordinalByScope.TryGetValue(scope.ScopePosition, out var localOrdinal);
            localOrdinal++;
            ordinalByScope[scope.ScopePosition] = localOrdinal;
            UpdateNativeEquationCaptionNumber(
                document,
                caption.FormulaId,
                nativeSequenceName,
                localOrdinal,
                scope.Prefix,
                formatOnly);
        }
        WriteNativeEquationTailFormulaId(document, captions.LastOrDefault()?.FormulaId);
    }

    private static IReadOnlyList<NativeEquationCaptionEntry> GetNativeEquationCaptionEntries(
        Document document,
        string nativeSequenceName)
    {
        var result = new List<NativeEquationCaptionEntry>();
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? numberRange = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryFormulaIdFromBookmark(
                            bookmark.Name,
                            NativeNumberBookmarkPrefix,
                            out var formulaId))
                        continue;
                    // VTEqNum_* already bookmarks the exact SEQ field result.
                    // Reading its Range gives both the document position and the
                    // current rendered number; opening every companion caption
                    // and Field here made inventory O(n) expensive COM calls.
                    if (!bookmarks.Exists(NativeCaptionBookmarkName(formulaId)))
                        continue;
                    numberRange = bookmark.Range;
                    result.Add(new NativeEquationCaptionEntry(
                        formulaId,
                        numberRange.Start,
                        numberRange.Text ?? string.Empty));
                }
                finally
                {
                    Release(numberRange);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }
        return result.OrderBy(item => item.Position).ToArray();
    }

    private static IReadOnlyList<HeadingNumberAnchor> GetHeadingNumberAnchorsForFormatBatch(
        Document document,
        int targetLevel,
        IReadOnlyList<NativeEquationCaptionEntry> captions)
    {
        var captionPositions = captions
            .Select(caption => caption.Position)
            .ToHashSet();
        var candidates = new List<HeadingParagraphCandidate>();
        Range? content = null;
        try
        {
            content = document.Content;
            for (var outlineLevel = 1; outlineLevel <= targetLevel; outlineLevel++)
            {
                var cursor = content.Start;
                while (cursor < content.End)
                {
                    Range? searchRange = null;
                    Find? find = null;
                    ParagraphFormat? findParagraphFormat = null;
                    Paragraphs? foundParagraphs = null;
                    Paragraph? foundParagraph = null;
                    Range? range = null;
                    ListFormat? listFormat = null;
                    Frames? frames = null;
                    try
                    {
                        searchRange = document.Range(cursor, content.End);
                        find = searchRange.Find;
                        find.ClearFormatting();
                        find.Text = "^p";
                        find.Forward = true;
                        find.Wrap = WdFindWrap.wdFindStop;
                        find.Format = true;
                        findParagraphFormat = find.ParagraphFormat;
                        findParagraphFormat.OutlineLevel = (WdOutlineLevel)outlineLevel;
                        if (!find.Execute()) break;

                        // Find narrows searchRange to the matching paragraph mark.
                        // Resolve only that paragraph; Word's native formatter has
                        // already skipped hundreds of formula/table/caption rows.
                        foundParagraphs = searchRange.Paragraphs;
                        if (foundParagraphs.Count == 0)
                        {
                            cursor = Math.Max(cursor + 1, searchRange.End);
                            continue;
                        }
                        foundParagraph = foundParagraphs[1];
                        range = foundParagraph.Range;
                        cursor = Math.Max(cursor + 1, range.End);

                        if (captionPositions.Contains(range.Start)) continue;
                        if ((bool)range.get_Information(WdInformation.wdWithInTable))
                            continue;
                        try
                        {
                            frames = range.Frames;
                            if (frames.Count > 0) continue;
                        }
                        catch { }

                        var listNumber = string.Empty;
                        try
                        {
                            listFormat = range.ListFormat;
                            listNumber = NormalizeHeadingNumberText(listFormat.ListString);
                        }
                        catch { }
                        candidates.Add(new HeadingParagraphCandidate(
                            range.Start,
                            outlineLevel,
                            listNumber,
                            ParseHeadingNumberFromText(range.Text, outlineLevel)));
                    }
                    finally
                    {
                        Release(frames);
                        Release(listFormat);
                        Release(range);
                        Release(foundParagraph);
                        Release(foundParagraphs);
                        Release(findParagraphFormat);
                        Release(find);
                        Release(searchRange);
                    }
                }
            }
        }
        catch
        {
            // Formatting-only Find is unavailable on a few compatibility-mode
            // builds. Preserve the existing exhaustive paragraph scan there.
            return GetHeadingNumberAnchors(document, targetLevel);
        }
        finally { Release(content); }

        candidates.Sort((left, right) => left.Position.CompareTo(right.Position));
        var result = new List<HeadingNumberAnchor>();
        var usesExplicitNumbering = candidates.Any(candidate =>
            !string.IsNullOrWhiteSpace(candidate.ExplicitNumber));
        var counters = new int[10];
        foreach (var candidate in candidates)
        {
            var numberText = candidate.ExplicitNumber;
            if (string.IsNullOrWhiteSpace(numberText))
            {
                if (usesExplicitNumbering) continue;
                counters[candidate.OutlineLevel]++;
                for (var deeper = candidate.OutlineLevel + 1;
                     deeper < counters.Length;
                     deeper++)
                    counters[deeper] = 0;
                var parts = new List<string>();
                for (var level = 1; level <= candidate.OutlineLevel; level++)
                    parts.Add(Math.Max(1, counters[level]).ToString());
                numberText = string.Join(".", parts);
            }

            if (candidate.OutlineLevel < targetLevel)
            {
                var missingLevels = targetLevel - candidate.OutlineLevel;
                numberText += string.Concat(
                    Enumerable.Repeat(".0", missingLevels));
            }
            result.Add(new HeadingNumberAnchor(candidate.Position, numberText));
        }
        return result;
    }

    private static IReadOnlyList<HeadingNumberAnchor> GetHeadingNumberAnchors(
        Document document,
        int targetLevel)
    {
        var candidates = new List<HeadingParagraphCandidate>();
        Paragraphs? paragraphs = null;
        try
        {
            paragraphs = document.Paragraphs;
            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Paragraph? paragraph = null;
                Range? range = null;
                ListFormat? listFormat = null;
                Frames? frames = null;
                try
                {
                    paragraph = paragraphs[index];
                    var outlineLevel = (int)paragraph.OutlineLevel;
                    if (outlineLevel < 1 || outlineLevel > targetLevel) continue;

                    range = paragraph.Range;
                    // Numbering tables and clipped native-caption paragraphs can
                    // inherit Heading 1 from the source formula paragraph. They
                    // are implementation scaffolds, never chapter boundaries.
                    if ((bool)range.get_Information(WdInformation.wdWithInTable))
                        continue;
                    try
                    {
                        frames = range.Frames;
                        if (frames.Count > 0) continue;
                    }
                    catch { }

                    var listNumber = string.Empty;
                    try
                    {
                        listFormat = range.ListFormat;
                        listNumber = NormalizeHeadingNumberText(listFormat.ListString);
                    }
                    catch { }
                    candidates.Add(new HeadingParagraphCandidate(
                        range.Start,
                        outlineLevel,
                        listNumber,
                        ParseHeadingNumberFromText(range.Text, outlineLevel)));
                }
                finally
                {
                    Release(frames);
                    Release(listFormat);
                    Release(range);
                    Release(paragraph);
                }
            }
        }
        finally { Release(paragraphs); }

        var result = new List<HeadingNumberAnchor>();
        var usesExplicitNumbering = candidates.Any(candidate =>
            !string.IsNullOrWhiteSpace(candidate.ExplicitNumber));
        var counters = new int[10];
        foreach (var candidate in candidates)
        {
            var numberText = candidate.ExplicitNumber;
            if (string.IsNullOrWhiteSpace(numberText))
            {
                // In a manually/automatically numbered document, an unnumbered
                // Heading 1 is usually the document title, not "chapter 1".
                if (usesExplicitNumbering) continue;
                counters[candidate.OutlineLevel]++;
                for (var deeper = candidate.OutlineLevel + 1;
                     deeper < counters.Length;
                     deeper++)
                    counters[deeper] = 0;
                var parts = new List<string>();
                for (var level = 1; level <= candidate.OutlineLevel; level++)
                    parts.Add(Math.Max(1, counters[level]).ToString());
                numberText = string.Join(".", parts);
            }

            if (candidate.OutlineLevel < targetLevel)
            {
                var missingLevels = targetLevel - candidate.OutlineLevel;
                numberText += string.Concat(
                    Enumerable.Repeat(".0", missingLevels));
            }
            result.Add(new HeadingNumberAnchor(candidate.Position, numberText));
        }
        return result;
    }

    private static string NormalizeHeadingNumberText(string? value)
    {
        var text = (value ?? string.Empty)
            .Replace("\t", string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Trim();
        return text.TrimEnd(' ', '.', '-', '–', '—', '、', ')', '）');
    }

    private static string ParseHeadingNumberFromText(string? value, int outlineLevel)
    {
        var text = (value ?? string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace("\a", string.Empty)
            .Trim();
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;

        // Support common manually typed headings such as "8. 多元微积分",
        // "8-1 小节" and "第 8 章" when Word's ListString is empty.
        var match = Regex.Match(
            text,
            @"^(?:第\s*)?(?<number>\d+(?:\s*[.．\-–—]\s*\d+)*)(?:\s*[章节篇部]|\s*[.．、:：)）\-–—]|\s+)",
            RegexOptions.CultureInvariant);
        if (!match.Success) return string.Empty;
        var number = Regex.Replace(
            match.Groups["number"].Value,
            @"\s*[.．\-–—]\s*",
            ".",
            RegexOptions.CultureInvariant);
        var parts = number
            .Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim())
            .Where(part => part.Length > 0)
            .ToArray();
        if (parts.Length == 0 || parts.Length > Math.Max(1, outlineLevel))
            return string.Empty;
        return string.Join(".", parts);
    }

    internal static IReadOnlyDictionary<int, ResolvedEquationHeadingScope>
        CaptureHeadingScopesAtPositions(
            Document document,
            string? formatId,
            IEnumerable<int> formulaPositions)
    {
        var positions = formulaPositions.Distinct().ToArray();
        var result = new Dictionary<int, ResolvedEquationHeadingScope>();
        if (positions.Length == 0) return result;

        var format = EquationNumberFormat.Resolve(formatId);
        if (!format.UsesHeading)
        {
            foreach (var position in positions)
                result[position] = new ResolvedEquationHeadingScope(
                    int.MinValue,
                    int.MinValue,
                    string.Empty);
            return result;
        }

        // Expensive Word paragraph/outline discovery happens exactly once for the
        // whole conversion batch. ResolveEquationNumberScope itself is pure list
        // lookup, and the paragraph end for a real heading is read once per unique
        // heading rather than once per numbered formula.
        var anchors = GetHeadingNumberAnchors(document, format.HeadingLevel);
        var scopeEnds = new Dictionary<int, int>();
        foreach (var formulaPosition in positions)
        {
            var scope = ResolveEquationNumberScope(formulaPosition, format, anchors);
            var numberText = scope.Prefix;
            if (!string.IsNullOrEmpty(format.Separator)
                && numberText.EndsWith(format.Separator, StringComparison.Ordinal))
                numberText = numberText.Substring(
                    0,
                    numberText.Length - format.Separator.Length);

            if (scope.ScopePosition == int.MinValue)
            {
                result[formulaPosition] = new ResolvedEquationHeadingScope(
                    int.MinValue,
                    int.MinValue,
                    numberText);
                continue;
            }

            if (!scopeEnds.TryGetValue(scope.ScopePosition, out var scopeEnd))
            {
                Range? probe = null;
                Paragraphs? paragraphs = null;
                Paragraph? paragraph = null;
                Range? paragraphRange = null;
                try
                {
                    probe = document.Range(scope.ScopePosition, scope.ScopePosition);
                    paragraphs = probe.Paragraphs;
                    if (paragraphs.Count == 0)
                        scopeEnd = scope.ScopePosition;
                    else
                    {
                        paragraph = paragraphs[1];
                        paragraphRange = paragraph.Range;
                        scopeEnd = paragraphRange.End;
                    }
                    scopeEnds[scope.ScopePosition] = scopeEnd;
                }
                finally
                {
                    Release(paragraphRange);
                    Release(paragraph);
                    Release(paragraphs);
                    Release(probe);
                }
            }

            result[formulaPosition] = new ResolvedEquationHeadingScope(
                scope.ScopePosition,
                scopeEnd,
                numberText);
        }
        return result;
    }

    internal static (int ScopeStart, int ScopeEnd, string NumberText) ResolveHeadingScopeAtPosition(
        Document document,
        int formulaPosition,
        string? formatId)
    {
        var format = EquationNumberFormat.Resolve(formatId);
        if (!format.UsesHeading)
            return (int.MinValue, int.MinValue, string.Empty);

        var anchors = GetHeadingNumberAnchors(document, format.HeadingLevel);
        var scope = ResolveEquationNumberScope(formulaPosition, format, anchors);
        var numberText = scope.Prefix;
        if (!string.IsNullOrEmpty(format.Separator)
            && numberText.EndsWith(format.Separator, StringComparison.Ordinal))
            numberText = numberText.Substring(0, numberText.Length - format.Separator.Length);

        if (scope.ScopePosition == int.MinValue)
            return (int.MinValue, int.MinValue, numberText);

        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            probe = document.Range(scope.ScopePosition, scope.ScopePosition);
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count == 0)
                return (scope.ScopePosition, scope.ScopePosition, numberText);
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return (scope.ScopePosition, paragraphRange.End, numberText);
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    private static (int ScopePosition, string Prefix) ResolveEquationNumberScope(
        int formulaPosition,
        EquationNumberFormat format,
        IReadOnlyList<HeadingNumberAnchor> anchors)
    {
        if (!format.UsesHeading) return (0, string.Empty);
        HeadingNumberAnchor? selected = null;
        for (var index = 0; index < anchors.Count; index++)
        {
            var anchor = anchors[index];
            if (anchor.Position > formulaPosition) break;
            selected = anchor;
        }
        var headingText = selected?.NumberText
            ?? string.Join(".", Enumerable.Repeat("0", format.HeadingLevel));
        return (
            selected?.Position ?? int.MinValue,
            headingText + format.Separator);
    }

    private static void UpdateNativeEquationCaptionNumber(
        Document document,
        string formulaId,
        string nativeSequenceName,
        int ordinal,
        string prefix,
        bool formatOnly,
        bool cleanupLegacyFrames = true)
    {
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
            code.Text = $" SEQ {nativeSequenceName} \\r {ordinal} \\* ARABIC ";
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

    private static void UpdateFieldInBookmark(
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
                    code = field.Code;
                    if (predicate(code.Text))
                    {
                        field.Update();
                        NormalizeReferenceResult(field);
                    }
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
        }
        finally
        {
            Release(fields);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void AlignEquationNumberVertically(
        Range numberRange,
        float formulaHeightPoints,
        float formulaFontSizePoints)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = numberRange.Font;
            var numberFontSize = FormulaFontSize.Normalize(formulaFontSizePoints);

            // The native caption target is deliberately white and one point.
            // Word propagates that appearance into REF results unless the
            // visible range is normalized after every field update. Normalize
            // the brackets and field result as one run so locale-specific body
            // fonts cannot put the digits and parentheses on different baselines.
            ApplyEquationNumberFont(
                font,
                numberFontSize,
                CalculateEquationNumberFontPosition(
                    formulaHeightPoints,
                    numberFontSize));
        }
        finally { Release(font); }
    }

    private static void ApplyEquationNumberFont(
        Microsoft.Office.Interop.Word.Font font,
        float size,
        int position)
    {
        font.Hidden = 0;
        font.Color = WdColor.wdColorAutomatic;
        font.Size = size;
        font.Position = position;
        try { font.Name = EquationNumberFontName; } catch { }
        try { font.NameAscii = EquationNumberFontName; } catch { }
        try { font.NameFarEast = EquationNumberFontName; } catch { }
        try { font.NameBi = EquationNumberFontName; } catch { }
        try { font.Bold = 0; } catch { }
        try { font.Italic = 0; } catch { }
        try { font.Superscript = 0; } catch { }
        try { font.Subscript = 0; } catch { }
        try { font.Scaling = 100; } catch { }
        try { font.Spacing = 0f; } catch { }
        try { font.Kerning = 0f; } catch { }
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
            var xml = content.WordOpenXML ?? string.Empty;
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
                var textMatches = Regex.Matches(
                    segment,
                    @"<w:t(?:\s[^>]*)?>(?<text>.*?)</w:t>",
                    RegexOptions.IgnoreCase
                    | RegexOptions.CultureInvariant
                    | RegexOptions.Singleline);
                var numberText = NormalizeNativeEquationNumberText(string.Concat(
                    textMatches.Cast<Match>()
                        .Select(match => System.Net.WebUtility.HtmlDecode(
                            match.Groups["text"].Value))));
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
            .Select((entry, index) => new EquationReferenceTarget(
                entry.FormulaId,
                index + 1,
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
        EquationReferenceStyle style)
    {
        var prefix = style switch
        {
            EquationReferenceStyle.EquationPrefix => "式（",
            EquationReferenceStyle.Parenthesized => "(",
            _ => string.Empty,
        };
        var suffix = style switch
        {
            EquationReferenceStyle.EquationPrefix => "）",
            EquationReferenceStyle.Parenthesized => ")",
            _ => string.Empty,
        };

        Range? insertion = null;
        Fields? fields = null;
        Field? field = null;
        Range? result = null;
        try
        {
            if (!string.IsNullOrEmpty(prefix)) selection.TypeText(prefix);
            insertion = selection.Range.Duplicate;
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            fields = document.Fields;
            object fieldType = WdFieldEmpty;
            object fieldCode = $"REF {NativeNumberBookmarkName(target.FormulaId)} \\h";
            object preserveFormatting = true;
            field = fields.Add(
                insertion,
                ref fieldType,
                ref fieldCode,
                ref preserveFormatting);
            NormalizeReferenceResult(field);
            result = field.Result;
            selection.SetRange(result.End, result.End);
            if (!string.IsNullOrEmpty(suffix)) selection.TypeText(suffix);
        }
        finally
        {
            Release(result);
            Release(field);
            Release(fields);
            Release(insertion);
        }
    }

    internal static int FreezeFormulaCrossReferences(
        Document document,
        string formulaId,
        IReadOnlyDictionary<string, int>? knownReferenceCounts = null)
    {
        // The one generated REF inside the formula's own number cell is about to
        // be removed together with the numbering scaffold, so it does not need to
        // be unlinked. Only external references require document.Fields traversal.
        // This turns the common no-external-reference batch case from O(N^2) COM
        // enumeration into O(1) per numbered formula after one OpenXML parse.
        if (knownReferenceCounts is not null
            && knownReferenceCounts.TryGetValue(formulaId, out var knownCount)
            && knownCount <= 1)
            return 0;

        var targetBookmarkName = NativeNumberBookmarkName(formulaId);
        Fields? fields = null;
        var frozen = 0;
        try
        {
            fields = document.Fields;
            for (var index = fields.Count; index >= 1; index--)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    var fieldCode = code.Text;
                    var matches = IsReferenceToBookmark(
                        fieldCode,
                        targetBookmarkName);
                    if (!matches
                        && TryResolveVisualTeXReferenceBookmark(
                            document,
                            fieldCode,
                            out var resolvedBookmark))
                    {
                        matches = string.Equals(
                            resolvedBookmark,
                            targetBookmarkName,
                            StringComparison.OrdinalIgnoreCase);
                    }
                    if (!matches) continue;
                    NormalizeReferenceResult(field);
                    field.Unlink();
                    frozen++;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }
        return frozen;
    }

    private static int UpdateHealthyNativeCrossReferencesAfterRenumbering(
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
                        || referenceCount < 1
                        || !TryPatchVisibleEquationNumberResult(
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
                return fastUpdated;
            }
        }

        Range? content = null;
        Range? searchRange = null;
        Fields? fields = null;
        var updated = 0;
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
                try
                {
                    field = fields[index];
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
                            out var formulaId)
                        || !changedFormulaNumbers.TryGetValue(
                            formulaId,
                            out var expectedNumber))
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
                    if (!alreadyCanonical)
                    {
                        NormalizeReferenceResult(field);
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
                    updated++;
                }
                catch
                {
                    // A protected/legacy field can reject direct result mutation.
                    // Fall back to Word's normal field update for that one field;
                    // healthy generated references never pay this heavier path.
                    try { NormalizeReferenceResult(field!); updated++; } catch { }
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
        return updated;
    }

    internal static int UpdateNativeCrossReferences(Document document) =>
        UpdateNativeCrossReferences(document, targetFormulaIds: null);

    private static int UpdateNativeCrossReferences(
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

                    if (targetFormulaIds is not null)
                    {
                        if (!TryFormulaIdFromBookmark(
                                visualTeXBookmark,
                                NativeNumberBookmarkPrefix,
                                out var formulaId)
                            || !targetFormulaIds.Contains(formulaId))
                            continue;
                    }

                    if (!IsReferenceToBookmark(code.Text, visualTeXBookmark))
                        code.Text = $" REF {visualTeXBookmark} \\h ";

                    // NormalizeReferenceResult performs the update itself after
                    // applying CHARFORMAT. Restrict this pass to VisualTeX
                    // equation references; touching unrelated REF fields made a
                    // local formula edit scale with the rest of the document.
                    NormalizeReferenceResult(field);
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
        return updated;
    }

    private static void NormalizeReferenceResult(
        Field field,
        float? knownFontSize = null)
    {
        Range? code = null;
        Range? result = null;
        Microsoft.Office.Interop.Word.Font? codeFont = null;
        Microsoft.Office.Interop.Word.Font? resultFont = null;
        try
        {
            var size = knownFontSize.HasValue && knownFontSize.Value > 0f
                ? knownFontSize.Value
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
                code.Text = normalizedCode;

            // CHARFORMAT makes Word use the field-code appearance instead of
            // copying the hidden one-point SEQ target appearance into the REF.
            codeFont = code.Font;
            ApplyEquationNumberFont(codeFont, size, position: 0);
            field.Update();

            result = field.Result;
            resultFont = result.Font;
            ApplyEquationNumberFont(resultFont, size, position: 0);
        }
        finally
        {
            Release(resultFont);
            Release(codeFont);
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

    private static List<int> GetNativeEquationFieldPositions(
        Document document,
        string nativeSequenceName)
    {
        var positions = new List<int>();
        Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName)) continue;
                    result = field.Result;
                    positions.Add(result.Start);
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }
        positions.Sort();
        return positions;
    }

    private static bool TryGetNativeCaptionInfo(
        Document document,
        string formulaId,
        string nativeSequenceName,
        out int position,
        out string numberText)
    {
        position = -1;
        numberText = string.Empty;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Field? field = null;
        Range? result = null;
        try
        {
            bookmarks = document.Bookmarks;
            var bookmarkName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(bookmarkName)) return false;
            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range;
            field = FindNativeEquationFieldAtRange(document, range, nativeSequenceName);
            if (field is null) return false;
            field.Update();
            result = field.Result;
            position = result.Start;
            numberText = (range.Text ?? string.Empty).Trim();
            return !string.IsNullOrWhiteSpace(numberText);
        }
        finally
        {
            Release(result);
            Release(field);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static Field? FindNativeEquationFieldInRange(
        Range scopeRange,
        string nativeSequenceName)
    {
        Fields? fields = null;
        try
        {
            fields = scopeRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName))
                        continue;
                    var found = field;
                    field = null;
                    return found;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            return null;
        }
        finally { Release(fields); }
    }

    private static Field? FindNativeEquationFieldAtRange(
        Document document,
        Range targetRange,
        string nativeSequenceName)
    {
        // Word can invalidate a live Range RCW while the document-wide Fields
        // collection is materialized. Freeze the coordinates before enumerating
        // fields so later comparisons never re-enter a deleted COM proxy.
        var targetStart = targetRange.Start;
        var targetEnd = targetRange.End;
        Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName)) continue;
                    result = field.Result;
                    var overlaps = result.Start < targetEnd
                        && result.End > targetStart;
                    var sameCollapsedPosition = result.Start == targetStart
                        && result.End == targetEnd;
                    if (!overlaps && !sameCollapsedPosition) continue;
                    var found = field;
                    field = null;
                    return found;
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
            return null;
        }
        finally { Release(fields); }
    }

    private static string GetNativeEquationSequenceName(Document document)
    {
        Microsoft.Office.Interop.Word.Application? application = null;
        CaptionLabels? labels = null;
        CaptionLabel? label = null;
        try
        {
            application = document.Application;
            labels = application.CaptionLabels;
            label = labels[WdCaptionLabelID.wdCaptionEquation];
            var name = label.Name;
            if (string.IsNullOrWhiteSpace(name))
                throw new InvalidOperationException("Word built-in Equation caption label is unavailable.");
            return name;
        }
        finally
        {
            Release(label);
            Release(labels);
            Release(application);
        }
    }

    internal static string EquationBookmarkName(string formulaId) =>
        BookmarkName(EquationBookmarkPrefix, formulaId);

    internal static string NativeCaptionBookmarkName(string formulaId) =>
        BookmarkName(NativeCaptionBookmarkPrefix, formulaId);

    internal static string NativeNumberBookmarkName(string formulaId) =>
        BookmarkName(NativeNumberBookmarkPrefix, formulaId);

    private static string BookmarkName(string prefix, string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var value))
            throw new InvalidOperationException("VisualTeX formulaId must be a UUID.");
        return $"{prefix}{value:N}";
    }

    internal static bool TryFormulaIdFromEquationBookmark(
        string? bookmarkName,
        out string formulaId) =>
        TryFormulaIdFromBookmark(bookmarkName, EquationBookmarkPrefix, out formulaId);

    private static bool TryFormulaIdFromBookmark(
        string? bookmarkName,
        string prefix,
        out string formulaId)
    {
        formulaId = string.Empty;
        if (string.IsNullOrWhiteSpace(bookmarkName)) return false;
        var name = bookmarkName!;
        if (!name.StartsWith(prefix, StringComparison.Ordinal)
            || !Guid.TryParseExact(name.Substring(prefix.Length), "N", out var value))
            return false;
        formulaId = value.ToString();
        return true;
    }

    internal static bool IsVisualTeXSequenceFieldCode(string? code) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf(
            $"SEQ {LegacyEquationSequenceName}",
            StringComparison.OrdinalIgnoreCase) >= 0;

    internal static bool IsNativeEquationSequenceFieldCode(
        string? code,
        string nativeSequenceName)
    {
        if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(nativeSequenceName))
            return false;
        return code!.IndexOf(
                   $"SEQ {nativeSequenceName}",
                   StringComparison.OrdinalIgnoreCase) >= 0
            || code.IndexOf(
                   $"SEQ \"{nativeSequenceName}\"",
                   StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static bool TryResolveVisualTeXReferenceBookmark(
        Document document,
        string? fieldCode,
        out string bookmarkName)
    {
        bookmarkName = string.Empty;
        if (string.IsNullOrWhiteSpace(fieldCode)) return false;
        var match = Regex.Match(
            fieldCode!,
            @"^\s*REF\s+(?:""(?<quoted>[^""]+)""|(?<plain>[^\s\\]+))",
            RegexOptions.IgnoreCase);
        if (!match.Success) return false;
        var targetName = match.Groups["quoted"].Success
            ? match.Groups["quoted"].Value
            : match.Groups["plain"].Value;
        if (targetName.StartsWith(NativeNumberBookmarkPrefix, StringComparison.Ordinal))
        {
            bookmarkName = targetName;
            return true;
        }

        Bookmarks? bookmarks = null;
        Bookmark? targetBookmark = null;
        Range? targetRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(targetName)) return false;
            targetBookmark = bookmarks[targetName];
            targetRange = targetBookmark.Range;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? candidateBookmark = null;
                Range? candidateRange = null;
                try
                {
                    candidateBookmark = bookmarks[index];
                    if (!candidateBookmark.Name.StartsWith(
                            NativeNumberBookmarkPrefix,
                            StringComparison.Ordinal))
                        continue;
                    candidateRange = candidateBookmark.Range;
                    var overlaps = candidateRange.Start < targetRange.End
                        && candidateRange.End > targetRange.Start;
                    var sameCollapsedPosition = candidateRange.Start == targetRange.Start
                        && candidateRange.End == targetRange.End;
                    if (!overlaps && !sameCollapsedPosition) continue;
                    bookmarkName = candidateBookmark.Name;
                    return true;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidateBookmark);
                }
            }
            return false;
        }
        finally
        {
            Release(targetRange);
            Release(targetBookmark);
            Release(bookmarks);
        }
    }

    private static bool IsReferenceFieldCode(string? code)
    {
        if (string.IsNullOrWhiteSpace(code)) return false;
        var trimmed = code!.TrimStart();
        return trimmed.StartsWith("REF ", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsReferenceToBookmark(string? code, string bookmarkName) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf(
            $"REF {bookmarkName}",
            StringComparison.OrdinalIgnoreCase) >= 0;

    private static void DeleteBookmarkedRange(Document document, string name)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            range = bookmark.Range;
            range.Delete();
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void DeleteBookmarkOnly(Document document, string name)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            bookmark.Delete();
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
