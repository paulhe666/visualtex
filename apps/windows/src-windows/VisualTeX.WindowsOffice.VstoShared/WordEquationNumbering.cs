using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml.Linq;
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

internal static partial class WordEquationNumbering
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
    private const string NativeDisplayAnchorBookmarkPrefix = "VTEqAnc_";
    // Word bookmark names are limited to 40 characters; this 7-character prefix
    // plus a 32-character UUID marks an anchor that has survived one Office turn.
    private const string NativeDisplayAnchorCommitBookmarkPrefix = "VTAncR_";
    private const string NativeDisplayNumberShapeNamePrefix = "VTEqShape_";
    private const string NativeDisplayNumberShapeAlternativeTextPrefix =
        "VisualTeX numbered OMML ";
    private const float NativeDisplayAnchorLineSpacingPoints = 1f;
    private const float NativeDisplayNumberShapeWidthPoints = 72f;
    private const float NativeDisplayNumberShapeDefaultHeightPoints = 30f;
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
    private static readonly object NativeDisplayGeometryCacheSync = new();
    private static readonly Dictionary<string, float> NativeDisplayGeometryTopCache =
        new(StringComparer.Ordinal);

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr windowHandle);

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

    internal static bool NeedsLegacyManagedNumberingMigration(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (HasLegacyNumberedOmmlShapeArtifact(document))
            return true;
        if (TryReadHealthyEquationNumberArtifactsFromOpenXml(
                document,
                out _,
                out _))
        {
            Range? healthyContent = null;
            try
            {
                healthyContent = document.Content;
                var healthyXml = healthyContent.WordOpenXML ?? string.Empty;
                // A structurally valid Shape-era document is still migration input.
                // Current native #(SEQ) hosts never serialize these identities.
                if (healthyXml.IndexOf(
                        NativeDisplayNumberShapeNamePrefix,
                        StringComparison.OrdinalIgnoreCase) >= 0
                    || healthyXml.IndexOf(
                        NativeDisplayNumberShapeAlternativeTextPrefix,
                        StringComparison.Ordinal) >= 0
                    || healthyXml.IndexOf(
                        NativeDisplayAnchorBookmarkPrefix,
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
                return false;
            }
            finally { Release(healthyContent); }
        }

        Range? content = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            content = document.Content;
            var xml = content.WordOpenXML ?? string.Empty;
            if (Regex.IsMatch(
                    xml,
                    @"<w:bookmarkStart\b(?=[^>]*\bw:name=""(?:VTEq_|VTEqCap_|VTEqNum_)[0-9A-F]{32}""?)",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                return true;
            // Do not classify m:eqArr or #() itself as legacy: Word's native
            // display-number syntax deliberately serializes to m:eqArr. Only the
            // retired VisualTeX wrappers that put REF/placeholder payload inside
            // that array are migration markers.
            if (Regex.IsMatch(
                    xml,
                    @"<m:eqArr\b(?:(?!</m:eqArr>).){0,65536}(?:\bREF\s+VTEqNum_[0-9A-F]{32}\b|981730)",
                    RegexOptions.IgnoreCase
                    | RegexOptions.CultureInvariant
                    | RegexOptions.Singleline))
                return true;

            // A current managed native #(SEQ) formula can lose all three VTEq*
            // aliases while retaining its durable VTOMML FormulaId and Numbered
            // metadata. In that state the generic XML inventory sees no numbering
            // bookmark at all, so explicitly compare each managed numbered OMML
            // identity with its required alias triplet. This does not claim an
            // arbitrary user m:eqArr: only VisualTeX metadata can enter this path.
            foreach (var formulaId in WordOmmlFormulaStore.BookmarkedFormulaIds(document))
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                if (metadata is null
                    || !metadata.Numbered
                    || !string.Equals(
                        metadata.DisplayMode,
                        "block",
                        StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!HasCompleteFormulaNumberingArtifacts(document, formulaId))
                    return true;
            }

            // Very old VisualTeX OLE documents can have lost one or more numbering
            // bookmarks while still retaining Numbered=true metadata in a legacy
            // 1x3/2x3 host. Probe only cached VisualTeX metadata here so document
            // open never activates arbitrary embedded OLE servers.
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shapeRange); shapeRange = null;
                Release(shape); shape = shapes[index];
                var metadata = WordFormulaMetadataReader.TryReadCached(shape);
                if (metadata is null
                    || !metadata.Numbered
                    || !string.Equals(
                        metadata.DisplayMode,
                        "block",
                        StringComparison.OrdinalIgnoreCase))
                    continue;
                shapeRange = shape.Range;
                if (IsNumberedEquationTable(shapeRange)
                    || !HasCompleteFormulaNumberingArtifacts(
                        document,
                        metadata.FormulaId)
                    || !FormulaRangeOwnsNumberingArtifacts(
                        document,
                        shapeRange,
                        metadata.FormulaId))
                    return true;
            }
            return false;
        }
        catch
        {
            // A protected/custom story that cannot be safely inventoried must not
            // make ordinary document opening fail. The explicit update-number
            // command can retry after editing is enabled.
            return false;
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(content);
        }
    }

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
        // Healthy current direct-SEQ documents need only one OpenXML inventory.
        // Running the legacy spacing-repair scan first took a second full snapshot
        // of large documents even when no repair was possible. Keep repair as the
        // compatibility fallback, then retry the strict fast path before entering
        // full structural reconciliation.
        var fastPath = TryRefreshHealthyEquationNumbersInPlace(document, out var updated);
        if (!fastPath)
        {
            RepairNumberedDisplaySpacing(document);
            fastPath = TryRefreshHealthyEquationNumbersInPlace(document, out updated);
        }
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
        {
            if (!TryRepairDriftedManagedNativeHashSequenceHosts(
                    document,
                    out var repairedNativeHash)
                || repairedNativeHash <= 0)
                return false;
            TraceStage("native-hash-identity-repair");
            if (!TryReadHealthyEquationNumberArtifactsFromOpenXml(
                    document,
                    out openXmlCaptions,
                    out referenceCounts))
                return false;
        }
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
            format,
            trustedHealthyDirectTables: true);
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
        var format = ReadEquationNumberFormat(document);
        IReadOnlyDictionary<string, int>? knownReferenceCounts = null;
        IReadOnlyList<NativeEquationCaptionEntry> captions;
        if (TryReadHealthyEquationNumberArtifactsFromOpenXml(
                document,
                out var openXmlCaptions,
                out var referenceCounts))
        {
            captions = format.UsesHeading
                ? ResolveNativeEquationCaptionPositions(document, openXmlCaptions)
                : openXmlCaptions;
            if (captions.Count != openXmlCaptions.Count)
                return false;
            knownReferenceCounts = referenceCounts;
        }
        else
        {
            captions = GetNativeEquationCaptionEntries(document, nativeSequenceName);
        }
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
                out updated,
                knownReferenceCounts: knownReferenceCounts))
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

    private static bool ContainsVisualTeXSequenceInsideOmml(Range formulaRange)
    {
        OMaths? maths = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            maths = formulaRange.OMaths;
            fields = formulaRange.Fields;
            if (maths.Count != 1 || fields.Count != 1) return false;
            field = fields[1];
            code = field.Code;
            return IsVisualTeXSequenceFieldCode(code.Text);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(maths);
        }
    }

    internal static bool HasManagedNativeOmmlHashSequenceHost(
        Document document,
        string formulaId)
    {
        Bookmark? bookmark = null;
        Range? formulaRange = null;
        OMaths? maths = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
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
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            if (bookmark is null) return false;
            formulaRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            maths = formulaRange.OMaths;
            fields = formulaRange.Fields;
            if (maths.Count != 1 || fields.Count != 1) return false;
            field = fields[1];
            code = field.Code;
            // Do not require intact VTEq/VTEqCap/VTEqNum aliases here. This probe
            // exists specifically to keep a partially damaged mathematical SEQ out
            // of legacy caption code; strict wrapper/bookmark validation decides
            // later whether the host can be reused or must be rebuilt atomically.
            return IsVisualTeXSequenceFieldCode(code.Text);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(maths);
            Release(formulaRange);
            Release(bookmark);
        }
    }

    private static bool HasReusableManagedNativeOmmlHashSequenceHost(
        Document document,
        string formulaId)
    {
        Bookmark? bookmark = null;
        Range? formulaRange = null;
        try
        {
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            if (bookmark is null) return false;
            formulaRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            return HasReusableNumberedNativeOmmlHashSequenceHost(
                document,
                formulaRange,
                formulaId);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(formulaRange);
            Release(bookmark);
        }
    }

    private static bool TryApplyEquationNumberFormatByFieldBatch(
        Document document,
        string nativeSequenceName,
        IReadOnlyList<NativeEquationCaptionEntry> captions,
        out int updated,
        bool nativeTargetsAlreadyPlanned = false,
        IReadOnlyDictionary<string, int>? knownReferenceCounts = null)
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

            // Current direct-SEQ 1x3 hosts can be identified and updated from
            // their three right-cell aliases alone. Do this cheap local pass
            // before probing retired mathematical #(SEQ) hosts: the latter reads
            // CustomXML metadata and a complete OMath WordOpenXML payload for each
            // candidate, which made a four-number document take several seconds.
            var directTableFormulaIds = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            for (var planIndex = plan.Count - 1; planIndex >= 0; planIndex--)
            {
                var item = plan[planIndex];
                if (!TryUpdateDirectTableSequenceNumber(
                        document,
                        item.FormulaId,
                        nativeSequenceName,
                        item.Ordinal,
                        item.Prefix,
                        formatOnly: true,
                        knownBookmarks: bookmarks,
                        trustedHealthyDirectTable:
                            knownReferenceCounts is not null))
                    continue;
                directTableFormulaIds.Add(item.FormulaId);
            }
            TraceBatch("direct-table-targets");

            // Only non-table formulas can be retired mathematical #(SEQ) hosts.
            // Cache that small compatibility set after the direct-table pass so
            // healthy current documents never enter the expensive legacy probe.
            var nativeHashFormulaIds = plan
                .Where(item => !directTableFormulaIds.Contains(item.FormulaId)
                    && IsNumberedNativeOmmlHashSequenceFormula(
                        document,
                        item.FormulaId))
                .Select(item => item.FormulaId)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            TraceBatch("native-hash-inventory");

            // Resolve each legacy caption through its durable FormulaId. Work
            // backwards because the old external-prefix path mutates ordinary text;
            // native mathematical SEQ hosts were already handled above in order.
            long nativeLookupMs = 0;
            long nativeCodeMs = 0;
            long nativeUpdateMs = 0;
            long nativePrefixMs = 0;
            long nativeBookmarkMs = 0;
            for (var planIndex = plan.Count - 1; planIndex >= 0; planIndex--)
            {
                var item = plan[planIndex];
                if (directTableFormulaIds.Contains(item.FormulaId)
                    || nativeHashFormulaIds.Contains(item.FormulaId))
                    continue;

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
                    if (TryRefreshOrAtomicallyRebuildNativeHashSequenceV2(
                            document,
                            item.FormulaId,
                            item.Ordinal,
                            item.Prefix))
                        continue;

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
                    {
                        EnsureSequenceFieldCodeCanBeRewritten(field);
                        code.Text = $" SEQ {LegacyEquationSequenceName} \\r {item.Ordinal} \\* ARABIC ";
                    }
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
            TraceBatch("legacy-targets");

            // Legacy OLE captions above use explicit planned \\r ordinals. Commit
            // them first. A later native #(SEQ) field must see those new ordinals,
            // not the previous format's stale restart value (for example an OLE
            // field still at \\r 2 would otherwise make the following OMML field
            // render 2.1.3 instead of 2.1.2). Mathematical fields are then rebuilt
            // when necessary and refreshed in document order; their Field.Code.Text
            // is never modified in place.
            foreach (var item in plan)
            {
                if (!nativeHashFormulaIds.Contains(item.FormulaId))
                    continue;
                if (!TryRefreshNumberedNativeOmmlHashSequence(
                        document,
                        item.FormulaId,
                        item.Prefix))
                    throw new InvalidDataException(
                        "A native #(SEQ) OMML formula could not refresh its number format.");
            }
            TraceBatch("native-hash-targets");

            // Generated right-side numbers have their own durable bookmarks, so
            // patch those REF results locally instead of scanning document.Fields.
            var generatedReferencesPatched = true;
            foreach (var item in plan)
            {
                if (directTableFormulaIds.Contains(item.FormulaId)
                    || nativeHashFormulaIds.Contains(item.FormulaId))
                    continue;
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
            var hasOnlyGeneratedReferences = knownReferenceCounts is not null
                ? knownReferenceCounts.Count == captions.Count
                    && plan.All(item =>
                        knownReferenceCounts.TryGetValue(
                            item.FormulaId,
                            out var count)
                        && count == 1)
                : HasOnlyGeneratedEquationReferences(
                    document,
                    captions.Count);
            if (!generatedReferencesPatched || !hasOnlyGeneratedReferences)
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

    private sealed class OpenXmlFieldCodeState
    {
        internal System.Text.StringBuilder Code { get; } = new();
        internal bool ReadingInstruction { get; set; } = true;
    }

    private static IReadOnlyList<string> ExtractFieldCodesFromWordOpenXml(
        string wordOpenXml)
    {
        if (string.IsNullOrWhiteSpace(wordOpenXml))
            return Array.Empty<string>();
        const string wordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string mathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        try
        {
            var document = System.Xml.Linq.XDocument.Parse(
                wordOpenXml,
                System.Xml.Linq.LoadOptions.PreserveWhitespace);
            var word = (System.Xml.Linq.XNamespace)wordNamespace;
            var math = (System.Xml.Linq.XNamespace)mathNamespace;
            var active = new Stack<OpenXmlFieldCodeState>();
            var result = new List<string>();
            foreach (var element in document.Descendants())
            {
                if (element.Name == word + "fldSimple")
                {
                    var instruction = (string?)element.Attribute(word + "instr");
                    if (!string.IsNullOrWhiteSpace(instruction))
                        result.Add(instruction!);
                    continue;
                }

                if (element.Name == word + "fldChar")
                {
                    var kind = (string?)element.Attribute(word + "fldCharType")
                        ?? string.Empty;
                    if (string.Equals(kind, "begin", StringComparison.OrdinalIgnoreCase))
                    {
                        active.Push(new OpenXmlFieldCodeState());
                    }
                    else if (string.Equals(
                                 kind,
                                 "separate",
                                 StringComparison.OrdinalIgnoreCase))
                    {
                        if (active.Count > 0)
                            active.Peek().ReadingInstruction = false;
                    }
                    else if (string.Equals(kind, "end", StringComparison.OrdinalIgnoreCase))
                    {
                        if (active.Count > 0)
                        {
                            var completed = active.Pop().Code.ToString();
                            if (!string.IsNullOrWhiteSpace(completed))
                                result.Add(completed);
                        }
                    }
                    continue;
                }

                if (active.Count == 0 || !active.Peek().ReadingInstruction)
                    continue;
                if (element.Name == word + "instrText"
                    || element.Name == word + "t"
                    || element.Name == math + "t")
                    active.Peek().Code.Append(element.Value);
            }
            return result;
        }
        catch
        {
            return Array.Empty<string>();
        }
    }

    private static bool TryReadHealthyEquationNumberArtifactsFromOpenXml(
        Document document,
        out IReadOnlyList<NativeEquationCaptionEntry> entries,
        out IReadOnlyDictionary<string, int> referenceCounts)
    {
        entries = Array.Empty<NativeEquationCaptionEntry>();
        referenceCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var xml = string.Empty;
        bool Fail(string reason)
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                    "1",
                    StringComparison.Ordinal)
                || string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                TraceNumberingPerformance($"[perf] update-numbers.health-fail reason={reason}");
            return false;
        }
        int LastElementStart(string qualifiedName, int beforeIndex)
        {
            var bare = xml.LastIndexOf(
                $"<{qualifiedName}>",
                beforeIndex,
                StringComparison.OrdinalIgnoreCase);
            var attributed = xml.LastIndexOf(
                $"<{qualifiedName} ",
                beforeIndex,
                StringComparison.OrdinalIgnoreCase);
            return Math.Max(bare, attributed);
        }
        int FirstElementStart(string qualifiedName, int afterIndex)
        {
            var bare = xml.IndexOf(
                $"<{qualifiedName}>",
                afterIndex,
                StringComparison.OrdinalIgnoreCase);
            var attributed = xml.IndexOf(
                $"<{qualifiedName} ",
                afterIndex,
                StringComparison.OrdinalIgnoreCase);
            if (bare < 0) return attributed;
            if (attributed < 0) return bare;
            return Math.Min(bare, attributed);
        }
        Range? content = null;
        try
        {
            content = document.Content;
            xml = content.WordOpenXML ?? string.Empty;
            if (string.IsNullOrWhiteSpace(xml)) return Fail("empty-openxml");
            var parsedFieldCodes = ExtractFieldCodesFromWordOpenXml(xml);
            if (parsedFieldCodes.Count == 0)
                return Fail("no-parseable-field-codes");

            var bookmarkMatches = Regex.Matches(
                xml,
                @"<w:bookmarkStart\b[^>]*\bw:name=""(?<name>VTEq(?:Cap|Num)?_[^""]+)""[^>]*/?>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (bookmarkMatches.Count == 0) return Fail("no-numbering-bookmarks");
            if (Regex.IsMatch(
                    xml,
                    @"<w:bookmarkStart\b(?=[^>]*\bw:name=""(?:VTEqAnc_|VTAncR_)[0-9A-F]{32}""?)[^>]*/?>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                || Regex.IsMatch(
                    xml,
                    @"\b(?:name|id|descr|alt)=""(?:VTEqShape_[0-9A-F]{32}|VisualTeX numbered OMML [0-9A-F-]{36})""",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                return Fail("legacy-native-display-anchor-or-shape-remains");

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

            var visibleSet = visibleIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var captionSet = captionIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var numberSet = numberIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
            // A floating Word text box is commonly serialized twice inside one
            // mc:AlternateContent block: once in the DrawingML Choice and once in
            // the VML Fallback. Both representations contain the same VTEq_ bookmark
            // and generated REF, but they are one logical Shape in Word COM. Allow
            // only that visible-bookmark duplication; caption/number targets remain
            // unique main-story structures.
            if (visibleSet.Count == 0
                || captionSet.Count != captionIds.Count
                || numberSet.Count != numberIds.Count
                || !visibleSet.SetEquals(captionSet)
                || !visibleSet.SetEquals(numberSet))
                return Fail($"bookmark-id-set-mismatch visible={visibleIds.Count}/{visibleSet.Count}[{string.Join(",", visibleIds)}] caption={captionIds.Count}/{captionSet.Count}[{string.Join(",", captionIds)}] number={numberIds.Count}/{numberSet.Count}[{string.Join(",", numberIds)}]");

            var visualTeXSequenceCodes = parsedFieldCodes
                .Where(IsVisualTeXSequenceFieldCode)
                .ToArray();
            if (visualTeXSequenceCodes.Length != visibleSet.Count)
                return Fail(
                    $"visualtex-sequence-count-mismatch fields={visualTeXSequenceCodes.Length} formulas={visibleSet.Count}");
            // Legacy/current OLE captions may intentionally use an ordinal restart
            // while rebuilding a mixed-format document. That is ordinary Word text
            // and remains editable. Only a mathematical native #(SEQ) field is
            // forbidden from carrying \\r N, because its instruction must stay
            // immutable and dynamically ordered by F9. Enforce that per native
            // OMath below rather than rejecting the entire mixed document here.

            // A complete scaffold whose visible number no longer shares an owner
            // with its formula is an orphan (for example after a user manually
            // deletes only the OLE/OMML object). Current OLE and native OMML both
            // use one tab paragraph; legacy managed OMML tables remain accepted
            // long enough to be migrated by the structural reconciliation path.
            // Older OLE documents do not necessarily carry VTO_* identity
            // bookmarks, so validate the actual owner payload rather than requiring
            // an OLE identity bookmark.
            var allVisibleStartMatches = Regex.Matches(
                    xml,
                    @"<w:bookmarkStart\b(?=[^>]*\bw:name=""VTEq_(?<guid>[0-9A-F]{32})"")[^>]*/>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                .Cast<Match>()
                .ToArray();
            var visibleRepresentationCounts = allVisibleStartMatches
                .GroupBy(
                    match => match.Groups["guid"].Value,
                    StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    group => Guid.ParseExact(group.Key, "N").ToString("D"),
                    group => group.Count(),
                    StringComparer.OrdinalIgnoreCase);
            var visibleStartMatches = allVisibleStartMatches
                .GroupBy(
                    match => match.Groups["guid"].Value,
                    StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToArray();
            if (visibleStartMatches.Length != visibleSet.Count)
                return Fail($"visible-bookmark-start-count-mismatch representations={allVisibleStartMatches.Length} unique={visibleStartMatches.Length} expected={visibleSet.Count}");
            var hashSequenceFormulaIds = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            var directTableFormulaIds = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (var visibleStart in visibleStartMatches)
            {
                var normalizedFormulaId = visibleStart.Groups["guid"].Value;
                var formulaId = Guid.ParseExact(
                    normalizedFormulaId,
                    "N").ToString("D");

                // Current numbered OMML stores all three VisualTeX aliases inside
                // Word's own #(SEQ VisualTeXEquation) delimiter. A native m:eqArr
                // is therefore healthy when the VTEq_ alias is inside the one
                // display OMath, the same wrapper owns VTEqNum_<FormulaId>, and the
                // math field is SEQ rather than REF. Do this check before the legacy
                // Shape/TextBox branch so legal Word #() markup is never classified
                // as a malformed artificial eqArr.
                var mathStart = LastElementStart("m:oMath", visibleStart.Index);
                var precedingMathEnd = xml.LastIndexOf(
                    "</m:oMath>",
                    visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                var visibleNumberIsInMath = mathStart >= 0
                    && mathStart > precedingMathEnd;
                if (visibleNumberIsInMath)
                {
                    var mathEnd = xml.IndexOf(
                        "</m:oMath>",
                        visibleStart.Index,
                        StringComparison.OrdinalIgnoreCase);
                    if (mathEnd <= visibleStart.Index)
                        return Fail($"hash-seq-math-end-missing formulaId={formulaId}");
                    mathEnd += "</m:oMath>".Length;
                    if (mathEnd - mathStart > 262144)
                        return Fail($"hash-seq-math-too-large formulaId={formulaId}");
                    var mathXml = xml.Substring(mathStart, mathEnd - mathStart);
                    var expectedVisibleBookmark = EquationBookmarkPrefix + normalizedFormulaId;
                    var expectedNumberBookmark = NativeNumberBookmarkPrefix + normalizedFormulaId;
                    var expectedCaptionBookmark = NativeCaptionBookmarkPrefix + normalizedFormulaId;
                    bool BookmarkIsFullyInsideMath(string bookmarkName)
                    {
                        var escapedName = Regex.Escape(bookmarkName);
                        var start = Regex.Match(
                            mathXml,
                            $@"<w:bookmarkStart\b(?=[^>]*\bw:id=""(?<id>-?\d+)"")(?=[^>]*\bw:name=""{escapedName}"")[^>]*/?>",
                            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                        if (!start.Success) return false;
                        var escapedId = Regex.Escape(start.Groups["id"].Value);
                        var end = Regex.Match(
                            mathXml,
                            $@"<w:bookmarkEnd\b[^>]*\bw:id=""{escapedId}""[^>]*/?>",
                            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                        return end.Success && end.Index > start.Index;
                    }
                    var directSequence = Regex.IsMatch(
                        mathXml,
                        @"\bSEQ\s+(?:&quot;|"")?VisualTeXEquation(?:&quot;|"")?\b",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var hasVisibleAlias =
                        BookmarkIsFullyInsideMath(expectedVisibleBookmark);
                    var hasNumberAlias =
                        BookmarkIsFullyInsideMath(expectedNumberBookmark);
                    var hasCaptionAlias =
                        BookmarkIsFullyInsideMath(expectedCaptionBookmark);
                    var hasHash = Regex.IsMatch(
                        mathXml,
                        @"<m:t(?:\s[^>]*)?>#</m:t>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var hasEquationArray = mathXml.IndexOf(
                        "<m:eqArr",
                        StringComparison.OrdinalIgnoreCase) >= 0;
                    var hasMathReference = Regex.IsMatch(
                        mathXml,
                        @"\bREF\s+VTEqNum_",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var hasOrdinalRestart = Regex.IsMatch(
                        mathXml,
                        @"\\r\s+\d+\b",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    if (!directSequence || !hasVisibleAlias || !hasNumberAlias
                        || !hasCaptionAlias || !hasHash || !hasEquationArray
                        || hasMathReference || hasOrdinalRestart)
                        return Fail(
                            $"hash-seq-invalid formulaId={formulaId} seq={directSequence} visible={hasVisibleAlias} num={hasNumberAlias} cap={hasCaptionAlias} hash={hasHash} eqArr={hasEquationArray} ref={hasMathReference} restart={hasOrdinalRestart}");

                    var hashParagraphStart = LastElementStart("w:p", mathStart);
                    var hashPrecedingParagraphEnd = xml.LastIndexOf(
                        "</w:p>",
                        mathStart,
                        StringComparison.OrdinalIgnoreCase);
                    var hashParagraphEnd = xml.IndexOf(
                        "</w:p>",
                        mathEnd,
                        StringComparison.OrdinalIgnoreCase);
                    if (hashParagraphStart < 0
                        || hashParagraphStart <= hashPrecedingParagraphEnd
                        || hashParagraphEnd < mathEnd)
                        return Fail($"hash-seq-owner-paragraph-missing formulaId={formulaId}");
                    hashParagraphEnd += "</w:p>".Length;
                    var hashParagraphXml = xml.Substring(
                        hashParagraphStart,
                        hashParagraphEnd - hashParagraphStart);
                    if (hashParagraphXml.IndexOf(
                            "<m:oMathPara",
                            StringComparison.OrdinalIgnoreCase) < 0)
                        return Fail($"hash-seq-not-display formulaId={formulaId}");
                    var formulaBookmarkName =
                        WordOmmlFormulaStore.BookmarkName(formulaId);
                    var formulaBookmarkMatches = Regex.Matches(
                        xml,
                        $@"<w:bookmarkStart\b(?=[^>]*\bw:name=""{Regex.Escape(formulaBookmarkName)}"")[^>]*/?>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    if (formulaBookmarkMatches.Count != 1)
                        return Fail(
                            $"hash-seq-formula-identity-count formulaId={formulaId} count={formulaBookmarkMatches.Count}");
                    var formulaBookmarkIndex = formulaBookmarkMatches[0].Index;
                    var firstMathAfterFormulaBookmark = FirstElementStart(
                        "m:oMath",
                        formulaBookmarkIndex);
                    if (firstMathAfterFormulaBookmark != mathStart)
                        return Fail(
                            $"hash-seq-formula-identity-mismatch formulaId={formulaId} formulaBookmark={formulaBookmarkIndex} expectedMath={mathStart} firstMathAfterIdentity={firstMathAfterFormulaBookmark}");
                    var hashParagraphTableStart = LastElementStart(
                        "w:tbl",
                        hashParagraphStart);
                    var hashPrecedingTableEnd = xml.LastIndexOf(
                        "</w:tbl>",
                        hashParagraphStart,
                        StringComparison.OrdinalIgnoreCase);
                    if (hashParagraphTableStart >= 0
                        && hashParagraphTableStart > hashPrecedingTableEnd)
                        return Fail($"hash-seq-in-table formulaId={formulaId}");
                    if (hashParagraphXml.IndexOf(
                            NativeDisplayNumberShapeName(formulaId),
                            StringComparison.OrdinalIgnoreCase) >= 0
                        || hashParagraphXml.IndexOf(
                            "<w:txbxContent",
                            StringComparison.OrdinalIgnoreCase) >= 0)
                        return Fail($"hash-seq-has-shape formulaId={formulaId}");
                    hashSequenceFormulaIds.Add(formulaId);
                    continue;
                }

                // Shape-era numbered OMML stored the visible VTEq_ bookmark and an
                // ordinary REF inside a floating text-box story. Keep accepting it
                // only as migration input; the current producer never creates it.
                var textBoxStart = xml.LastIndexOf(
                    "<w:txbxContent",
                    visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                var precedingTextBoxEnd = xml.LastIndexOf(
                    "</w:txbxContent>",
                    visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                var visibleNumberIsInTextBox = textBoxStart >= 0
                    && textBoxStart > precedingTextBoxEnd;
                if (visibleNumberIsInTextBox)
                    return Fail($"legacy-numbering-shape formulaId={formulaId}");

                var tableStart = LastElementStart("w:tbl", visibleStart.Index);
                var precedingTableEnd = xml.LastIndexOf(
                    "</w:tbl>",
                    visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                var visibleNumberIsInTable = tableStart >= 0
                    && tableStart > precedingTableEnd;
                if (visibleNumberIsInTable)
                {
                    // Current numbered OMML uses one minimal 1x3 table only as a
                    // layout container. Cell (1,2) owns the one genuine display
                    // OMath and no fields; cell (1,3) owns TAB + direct ordinary
                    // SEQ + paragraph mark. Accept that exact structure as healthy
                    // rather than sending every Update Numbers command through the
                    // legacy-table migration/reconciliation path.
                    var tableEnd = xml.IndexOf(
                        "</w:tbl>",
                        visibleStart.Index,
                        StringComparison.OrdinalIgnoreCase);
                    if (tableEnd <= visibleStart.Index)
                        return Fail($"direct-table-end-missing formulaId={formulaId}");
                    tableEnd += "</w:tbl>".Length;
                    if (tableEnd - tableStart > 524288)
                        return Fail($"direct-table-too-large formulaId={formulaId}");
                    var tableXml = xml.Substring(tableStart, tableEnd - tableStart);
                    var rowCount = Regex.Matches(
                        tableXml,
                        @"<w:tr(?:\s|>)",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                    var cellStarts = Regex.Matches(
                            tableXml,
                            @"<w:tc(?:\s|>)",
                            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                        .Cast<Match>()
                        .ToArray();
                    if (rowCount != 1 || cellStarts.Length != 3)
                        return Fail(
                            $"direct-table-dimensions formulaId={formulaId} rows={rowCount} cells={cellStarts.Length}");
                    string ReadCellXml(int cellIndex)
                    {
                        var cellStart = cellStarts[cellIndex].Index;
                        var cellEnd = tableXml.IndexOf(
                            "</w:tc>",
                            cellStart,
                            StringComparison.OrdinalIgnoreCase);
                        return cellEnd <= cellStart
                            ? string.Empty
                            : tableXml.Substring(
                                cellStart,
                                cellEnd + "</w:tc>".Length - cellStart);
                    }
                    var leftCellXml = ReadCellXml(0);
                    var centerCellXml = ReadCellXml(1);
                    var rightCellXml = ReadCellXml(2);
                    if (string.IsNullOrEmpty(leftCellXml)
                        || string.IsNullOrEmpty(centerCellXml)
                        || string.IsNullOrEmpty(rightCellXml))
                        return Fail($"direct-table-cell-end-missing formulaId={formulaId}");
                    var centerMathCount = Regex.Matches(
                        centerCellXml,
                        @"<m:oMath(?:\s|>)",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                    var centerDisplayCount = Regex.Matches(
                        centerCellXml,
                        @"<m:oMathPara(?:\s|>)",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                    var centerFieldCount = Regex.Matches(
                        centerCellXml,
                        @"<w:fldChar\b",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                    if (centerMathCount != 1
                        || centerDisplayCount != 1
                        || centerFieldCount != 0)
                        return Fail(
                            $"direct-table-center-invalid formulaId={formulaId} math={centerMathCount} display={centerDisplayCount} fields={centerFieldCount}");
                    var formulaBookmarkName =
                        WordOmmlFormulaStore.BookmarkName(formulaId);
                    // VTEqNum_<FormulaId> in cell (1,3) is the durable physical
                    // owner of current numbered OMML: it resolves this exact 1x3
                    // table, whose center cell has already been proven to contain
                    // exactly one Display OMath. Word can normalize the zero-length
                    // VTOMML_* convenience anchor from the start of cell (1,2)'s
                    // paragraph to the row boundary immediately before that cell,
                    // especially after MathType→OMML batch conversion. Accept that
                    // serialization only when the anchor remains uniquely inside
                    // this same table; do not make an unstable collapsed bookmark
                    // carry the formula's primary identity.
                    var formulaBookmarkMatches = Regex.Matches(
                        tableXml,
                        $@"<w:bookmarkStart\b(?=[^>]*\bw:name=""{Regex.Escape(formulaBookmarkName)}"")[^>]*/?>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    if (formulaBookmarkMatches.Count != 1)
                        return Fail(
                            $"direct-table-formula-identity-missing formulaId={formulaId} count={formulaBookmarkMatches.Count}");
                    var expectedVisibleBookmark = EquationBookmarkPrefix + normalizedFormulaId;
                    var expectedNumberBookmark = NativeNumberBookmarkPrefix + normalizedFormulaId;
                    var expectedCaptionBookmark = NativeCaptionBookmarkPrefix + normalizedFormulaId;
                    bool RightCellHasBookmark(string bookmarkName) => Regex.IsMatch(
                        rightCellXml,
                        $@"<w:bookmarkStart\b(?=[^>]*\bw:name=""{Regex.Escape(bookmarkName)}"")[^>]*/?>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var hasDirectSequence = Regex.IsMatch(
                        rightCellXml,
                        @"\bSEQ\s+(?:&quot;|"")?VisualTeXEquation(?:&quot;|"")?\b",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var directTableHasRightTabStop = Regex.IsMatch(
                        rightCellXml,
                        @"<w:tab\b(?=[^>]*\bw:val=""right"")(?=[^>]*\bw:pos=""(?<pos>\d+)"")[^>]*/>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var hasLayoutTab = Regex.IsMatch(
                        rightCellXml,
                        @"<w:tab\s*/>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var hasMathInRightCell = Regex.IsMatch(
                        rightCellXml,
                        @"<m:oMath(?:\s|>)",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var hasRefInRightCell = Regex.IsMatch(
                        rightCellXml,
                        @"\bREF\s+VTEqNum_",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var rightSequenceCount = Regex.Matches(
                        rightCellXml,
                        @"\bSEQ\s+(?:&quot;|"")?VisualTeXEquation(?:&quot;|"")?\b",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                    var gridWidths = Regex.Matches(
                            tableXml,
                            @"<w:gridCol\b(?=[^>]*\bw:w=""(?<width>\d+)"")[^>]*/>",
                            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                        .Cast<Match>()
                        .Select(match => int.TryParse(
                            match.Groups["width"].Value,
                            out var width) ? width : -1)
                        .ToArray();
                    var rightTabMatch = Regex.Match(
                        rightCellXml,
                        @"<w:tab\b(?=[^>]*\bw:val=""right"")(?=[^>]*\bw:pos=""(?<pos>\d+)"")[^>]*/>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var rightTabMatchesColumn = gridWidths.Length == 3
                        && gridWidths[0] > 0
                        && gridWidths[1] > 0
                        && gridWidths[2] == gridWidths[0]
                        && rightTabMatch.Success
                        && int.TryParse(
                            rightTabMatch.Groups["pos"].Value,
                            out var rightTabPosition)
                        && rightTabPosition == gridWidths[2];
                    if (!hasDirectSequence
                        || rightSequenceCount != 1
                        || !RightCellHasBookmark(expectedVisibleBookmark)
                        || !RightCellHasBookmark(expectedNumberBookmark)
                        || !RightCellHasBookmark(expectedCaptionBookmark)
                        || !directTableHasRightTabStop
                        || !hasLayoutTab
                        || hasMathInRightCell
                        || hasRefInRightCell
                        || !rightTabMatchesColumn)
                        return Fail(
                            $"direct-table-right-invalid formulaId={formulaId} seq={hasDirectSequence}/{rightSequenceCount} visible={RightCellHasBookmark(expectedVisibleBookmark)} num={RightCellHasBookmark(expectedNumberBookmark)} cap={RightCellHasBookmark(expectedCaptionBookmark)} rightTab={directTableHasRightTabStop} layoutTab={hasLayoutTab} math={hasMathInRightCell} ref={hasRefInRightCell} grid={string.Join("/", gridWidths)} tabPos={(rightTabMatch.Success ? rightTabMatch.Groups["pos"].Value : "-")}");
                    directTableFormulaIds.Add(formulaId);
                    continue;
                }

                var paragraphStart = LastElementStart("w:p", visibleStart.Index);
                var precedingParagraphEnd = xml.LastIndexOf(
                    "</w:p>",
                    visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                var paragraphEnd = xml.IndexOf(
                    "</w:p>",
                    visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                if (paragraphStart < 0
                    || paragraphStart <= precedingParagraphEnd
                    || paragraphEnd <= visibleStart.Index)
                    return Fail($"visible-bookmark-has-no-owner-paragraph formulaId={formulaId}");
                paragraphEnd += "</w:p>".Length;
                if (paragraphEnd - paragraphStart > 262144)
                    return Fail($"numbering-paragraph-too-large formulaId={formulaId}");
                var paragraphXml = xml.Substring(
                    paragraphStart,
                    paragraphEnd - paragraphStart);
                var oleFormulaCount = Regex.Matches(
                    paragraphXml,
                    Regex.Escape("ProgID=\"VisualTeX.Formula.1\""),
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                var ommlFormulaCount = Regex.Matches(
                    paragraphXml,
                    @"<m:oMath(?:\s|>)",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                var ownsOle = oleFormulaCount == 1 && ommlFormulaCount == 0;
                if (!ownsOle)
                    return Fail($"legacy-numbering-paragraph formulaId={formulaId} ole={oleFormulaCount} omml={ommlFormulaCount}");
                var hasCenterTabStop = Regex.IsMatch(
                    paragraphXml,
                    @"<w:tab\b(?=[^>]*\bw:val=""center"")[^>]*/>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                var hasRightTabStop = Regex.IsMatch(
                    paragraphXml,
                    @"<w:tab\b(?=[^>]*\bw:val=""right"")[^>]*/>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                var layoutTabCount = Regex.Matches(
                    paragraphXml,
                    @"<w:tab\s*/>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                if (!hasCenterTabStop || !hasRightTabStop || layoutTabCount < 2)
                    return Fail($"numbering-paragraph-tab-geometry-invalid formulaId={formulaId} owner=ole center={hasCenterTabStop} right={hasRightTabStop} tabs={layoutTabCount}");
            }

            var startMatches = Regex.Matches(
                xml,
                @"<w:bookmarkStart\b(?=[^>]*\bw:id=""(?<id>-?\d+)"")(?=[^>]*\bw:name=""VTEqNum_(?<guid>[0-9A-F]{32})"")[^>]*/>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (startMatches.Count != numberSet.Count) return Fail($"number-bookmark-start-count-mismatch starts={startMatches.Count} expected={numberSet.Count}");

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
                    return Fail($"duplicate-or-unknown-number-bookmark formulaId={formulaId}");

                var bookmarkId = Regex.Escape(startMatch.Groups["id"].Value);
                var endMatch = Regex.Match(
                    xml,
                    $@"<w:bookmarkEnd\b[^>]*\bw:id=""{bookmarkId}""[^>]*/>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                    TimeSpan.FromSeconds(1));
                if (!endMatch.Success || endMatch.Index <= startMatch.Index)
                    return Fail($"number-bookmark-end-missing formulaId={formulaId}");
                if (endMatch.Index - startMatch.Index > 16384)
                    return Fail($"number-bookmark-range-too-large formulaId={formulaId}");

                var segment = xml.Substring(
                    startMatch.Index + startMatch.Length,
                    endMatch.Index - startMatch.Index - startMatch.Length);
                var numberText = hashSequenceFormulaIds.Contains(formulaId)
                    ? ExtractNativeHashSequenceResultFromBookmarkXml(segment)
                    : NormalizeNativeEquationNumberText(string.Concat(
                        Regex.Matches(
                                segment,
                                @"<(?:w|m):t(?:\s[^>]*)?>(?<text>.*?)</(?:w|m):t>",
                                RegexOptions.IgnoreCase
                                | RegexOptions.CultureInvariant
                                | RegexOptions.Singleline)
                            .Cast<Match>()
                            .Select(match => System.Net.WebUtility.HtmlDecode(
                                match.Groups["text"].Value))));
                if (string.IsNullOrWhiteSpace(numberText)) return Fail($"number-bookmark-empty formulaId={formulaId}");
                result.Add(new NativeEquationCaptionEntry(
                    formulaId,
                    startMatch.Index,
                    numberText));
            }

            if (result.Count != visibleSet.Count) return Fail($"number-entry-count-mismatch entries={result.Count} expected={visibleSet.Count}");

            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var fieldCode in parsedFieldCodes)
            {
                var referenceMatch = Regex.Match(
                    fieldCode,
                    @"\bREF\s+(?:""(?<quoted>VTEqNum_[0-9A-F]{32})""|(?<plain>VTEqNum_[0-9A-F]{32}))(?=\s|\\|$)",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                if (!referenceMatch.Success) continue;
                var bookmarkName = referenceMatch.Groups["quoted"].Success
                    ? referenceMatch.Groups["quoted"].Value
                    : referenceMatch.Groups["plain"].Value;
                var normalizedGuid = bookmarkName.Substring(
                    NativeNumberBookmarkPrefix.Length);
                if (!Guid.TryParseExact(
                        normalizedGuid,
                        "N",
                        out var referenceGuid))
                    continue;
                var formulaId = referenceGuid.ToString("D");
                if (!visibleSet.Contains(formulaId)) continue;
                counts.TryGetValue(formulaId, out var currentCount);
                counts[formulaId] = currentCount + 1;
            }
            foreach (var formulaId in hashSequenceFormulaIds.Concat(directTableFormulaIds))
            {
                // Preserve the historical reference-count contract used by format
                // conversion: one means "no external body reference" and values
                // above one prove external dynamic references exist. Native #SEQ
                // and the current direct-SEQ 1x3 OMML host have no generated visible
                // REF, so account for their own visible number slot virtually.
                counts.TryGetValue(formulaId, out var externalCount);
                counts[formulaId] = externalCount + 1;
            }

            // OLE and legacy Shape hosts still physically own one generated REF;
            // native #SEQ hosts contribute the virtual slot above. Either way the
            // compatibility count must be at least one for every numbered formula.
            if (visibleSet.Any(formulaId =>
                    !counts.TryGetValue(formulaId, out var count) || count < 1))
                return Fail("generated-or-native-number-slot-missing");

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

    private static string ExtractVisibleEquationNumberTextFromOpenXmlSegment(
        string segment)
    {
        if (string.IsNullOrEmpty(segment)) return string.Empty;
        const string wordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string mathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        try
        {
            var wrapped =
                $"<root xmlns:w=\"{wordNamespace}\" xmlns:m=\"{mathNamespace}\">"
                + segment
                + "</root>";
            var document = System.Xml.Linq.XDocument.Parse(
                wrapped,
                System.Xml.Linq.LoadOptions.PreserveWhitespace);
            var word = (System.Xml.Linq.XNamespace)wordNamespace;
            var math = (System.Xml.Linq.XNamespace)mathNamespace;
            var fieldCodeState = new Stack<bool>();
            var visible = new System.Text.StringBuilder();
            foreach (var element in document.Root!.Descendants())
            {
                if (element.Name == word + "fldChar")
                {
                    var kind = (string?)element.Attribute(word + "fldCharType")
                        ?? string.Empty;
                    if (string.Equals(kind, "begin", StringComparison.OrdinalIgnoreCase))
                    {
                        fieldCodeState.Push(true);
                    }
                    else if (string.Equals(
                                 kind,
                                 "separate",
                                 StringComparison.OrdinalIgnoreCase))
                    {
                        if (fieldCodeState.Count > 0)
                        {
                            fieldCodeState.Pop();
                            fieldCodeState.Push(false);
                        }
                    }
                    else if (string.Equals(kind, "end", StringComparison.OrdinalIgnoreCase))
                    {
                        if (fieldCodeState.Count > 0)
                            fieldCodeState.Pop();
                    }
                    continue;
                }

                if (element.Name == word + "instrText")
                    continue;
                if (element.Name != word + "t" && element.Name != math + "t")
                    continue;
                if (fieldCodeState.Any(inCode => inCode))
                    continue;
                visible.Append(element.Value);
            }
            return visible.ToString();
        }
        catch
        {
            // The fast health probe is deliberately conservative. A malformed or
            // namespace-incomplete bookmark segment belongs on full reconciliation;
            // never guess its visible number by concatenating mathematical m:t runs,
            // because those runs can contain the hidden SEQ instruction itself.
            return string.Empty;
        }
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
        if (!TryReadHealthyEquationNumberArtifactsFromOpenXml(
                document,
                out var entries,
                out var referenceCounts))
            return false;
        return entries.Count == numberedFormulaCount
            && referenceCounts.Count == numberedFormulaCount
            && referenceCounts.Values.All(count => count == 1);
    }

    private static bool HasOnlyGeneratedEquationReferencesLegacy(
        Document document,
        int numberedFormulaCount)
    {
        Range? content = null;
        try
        {
            content = document.Content;
            var xmlText = content.WordOpenXML ?? string.Empty;
            var referenceCount = Regex.Matches(
                xmlText,
                @"\bREF\s+VTEqNum_[0-9A-F]{32}\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                .Count;

            // Shape/table/OLE-era layouts own one generated REF per formula. The
            // native OMML #() layout owns none: its visible number is the SEQ result
            // itself. Count those direct mathematical hosts and subtract them from
            // the generated-REF expectation. Any remaining REF is therefore a real
            // body reference and must be refreshed.
            var directHashIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var xml = XDocument.Parse(xmlText, LoadOptions.PreserveWhitespace);
                XNamespace math =
                    "http://schemas.openxmlformats.org/officeDocument/2006/math";
                XNamespace word =
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                foreach (var equation in xml.Descendants(math + "oMath"))
                {
                    var fieldText = string.Concat(
                        equation.Descendants(word + "instrText").Select(node => node.Value)
                        .Concat(equation.Descendants(math + "t").Select(node => node.Value)));
                    if (!IsVisualTeXSequenceFieldCode(fieldText)) continue;
                    foreach (var bookmarkStart in equation.Descendants(word + "bookmarkStart"))
                    {
                        var name = (string?)bookmarkStart.Attribute(word + "name");
                        if (TryFormulaIdFromBookmark(
                                name,
                                NativeNumberBookmarkPrefix,
                                out var formulaId))
                            directHashIds.Add(formulaId);
                    }
                }
            }
            catch
            {
                // Ambiguous XML must take the conservative reference-update path.
                return false;
            }

            var generatedReferenceCount = Math.Max(
                0,
                numberedFormulaCount - directHashIds.Count);
            return referenceCount == generatedReferenceCount;
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

    internal static void RestoreEquationNumberFormatForConversion(
        Document document,
        string formatId)
    {
        var resolved = EquationNumberFormat.Resolve(formatId).Id;
        if (TryReadDocumentEquationNumberFormatId(document, out var current)
            && string.Equals(current, resolved, StringComparison.Ordinal))
            return;
        WriteEquationNumberFormat(document, resolved);
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
        Table? knownNumberedTable = null,
        bool deferNativeOmmlShapeFinalization = false,
        bool deferNativeOmmlShapeCreation = false,
        bool deferNativeOmmlMetadataPersistence = false,
        string? preparedUnnumberedOmml = null)
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
                knownNumberedTable,
                deferNativeOmmlShapeFinalization,
                deferNativeOmmlShapeCreation,
                deferNativeOmmlMetadataPersistence,
                preparedUnnumberedOmml);
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
        if (ContainsNativeOmml(formulaRange))
        {
            ConfigureNumberedNativeOmmlDisplay(
                document,
                formulaRange,
                formulaHeightPoints,
                formulaFontSizePoints,
                metadata.FormulaId,
                reuseExistingScaffold: false,
                metadata,
                _ => { },
                plannedOrdinal,
                plannedPrefix,
                deferFieldUpdate: true);
            return;
        }

        // A conversion may still hand us the RCW for a legacy 1x3/2x3 owner,
        // but it is migration input only. No conversion path is allowed to create,
        // preserve or populate a numbered table; ConfigureNumberedDisplayFormula
        // converts a proven legacy OLE table to the final center/right-tab paragraph
        // before constructing its caption and external REF.

        ConfigureNumberedDisplayFormula(
            document,
            formulaRange,
            formulaHeightPoints,
            formulaFontSizePoints,
            metadata.FormulaId,
            reuseExistingScaffold: false,
            knownNumberedTable: null,
            useConversionSafeVisibleNumber: true,
            metadata: metadata);
    }

    internal static bool TryBuildConvertedOmmlNumberingBatch(
        Document document,
        IReadOnlyList<FormulaMetadata> metadataItems,
        out int built)
    {
        built = 0;
        var entries = new List<(string FormulaId, FormulaMetadata Metadata, int Position, int SourceParagraphStart)>();
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
                Paragraphs? sourceParagraphs = null;
                Paragraph? sourceParagraph = null;
                Range? sourceParagraphRange = null;
                try
                {
                    range = GetFreshConvertedOmmlRange(document, metadata);
                    sourceParagraphs = range.Paragraphs;
                    if (sourceParagraphs.Count != 1)
                        throw new InvalidOperationException(
                            $"Converted OMML formula {metadata.FormulaId} does not occupy one source paragraph before numbering migration.");
                    sourceParagraph = sourceParagraphs[1];
                    sourceParagraphRange = sourceParagraph.Range;
                    entries.Add((
                        metadata.FormulaId,
                        metadata,
                        range.Start,
                        sourceParagraphRange.Start));
                }
                finally
                {
                    Release(sourceParagraphRange);
                    Release(sourceParagraph);
                    Release(sourceParagraphs);
                    Release(range);
                }
            }

            TraceStage("inventory");

            // Plan the final number for every converted formula before any tab,
            // caption or REF structure is inserted. This is the same fast-path
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

            // Materialize every converted numbered OMML through the same genuine
            // display host as a normal insertion. The formula becomes one pure
            // wdOMathDisplay/m:oMathPara paragraph; its ordinary dynamic REF lives
            // in the external right-margin Word text-box story. Processing from the
            // end keeps all frozen source positions stable while anchor paragraphs
            // are inserted.
            foreach (var entry in entries.OrderByDescending(item => item.Position))
            {
                Range? range = null;
                Range? refreshedRange = null;
                Bookmark? repairedBookmark = null;
                try
                {
                    var metadata = entry.Metadata;
                    range = GetFreshConvertedOmmlRange(document, metadata);
                    TraceStage("locate-target");

                    var plannedNumber = plannedNumbers[entry.FormulaId];
                    ConfigureNumberedNativeOmmlDisplay(
                        document,
                        range,
                        WordOmmlFormulaStore.EstimateHeightPoints(range),
                        (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                        entry.FormulaId,
                        reuseExistingScaffold: false,
                        metadata,
                        stage => TraceStage("true-display-" + stage),
                        plannedNumber.Ordinal,
                        plannedNumber.Prefix,
                        deferFieldUpdate: true,
                        deferMetadataPersistence: true);
                    TraceStage("true-display-host");

                    refreshedRange = ResolveSingleNativeOmmlRange(range);
                    repairedBookmark = WordOmmlFormulaStore.Wrap(
                        document,
                        refreshedRange,
                        metadata,
                        replaceExisting: true);
                    WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                        metadata,
                        refreshedRange);
                    WordOmmlFormulaStore.Save(document, metadata);
                    TraceStage("repair-identity");
                    built++;
                }
                finally
                {
                    Release(repairedBookmark);
                    Release(refreshedRange);
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
        Table? knownNumberedTable = null,
        bool deferNativeOmmlShapeFinalization = false,
        bool deferNativeOmmlShapeCreation = false,
        bool deferNativeOmmlMetadataPersistence = false,
        string? preparedUnnumberedOmml = null)
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
            var deferLegacyNativeOmmlShapeCreation =
                deferNativeOmmlShapeCreation;
            if (ContainsNativeOmml(formulaRange))
            {
                try
                {
                    deferLegacyNativeOmmlShapeCreation =
                        deferLegacyNativeOmmlShapeCreation
                        || IsNumberedEquationTable(formulaRange)
                        || WordOmmlConverter.HasVisualTeXNativeEquationNumber(
                            formulaRange.WordOpenXML);
                }
                catch
                {
                    // Any ambiguous legacy native host is handled conservatively by
                    // the normal structural path; fresh table-free insertions keep
                    // their immediate Shape creation behavior.
                }
            }
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
                knownNumberedTable: knownNumberedTable,
                metadata: metadata,
                deferNativeOmmlShapeCreation:
                    deferLegacyNativeOmmlShapeCreation,
                deferNativeOmmlMetadataPersistence:
                    deferNativeOmmlMetadataPersistence);
            if (ContainsNativeOmml(formulaRange))
            {
                // Do not repaginate here. SEQ/REF updates below can still change the
                // Shape label width and Word's display paragraph geometry. One final
                // measurement is performed at the end of this reconciliation.
                // Legacy table/eqArr migration intentionally stops with a committed
                // pure-display formula and anchor; its posted Office turn creates and
                // finalizes the external REF Shape after replacement RCWs unwind.
            }
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
            if (rebuiltStableNumberingSlot)
            {
                // OLE -> OMML keeps the existing 1x3 table but deliberately removes
                // this formula's old VTEq/VTEqCap/VTEqNum artifacts before rebuilding
                // them. CreateNativeCaption therefore starts from a bare SEQ field.
                // The formula has not moved, but its *format* still has to be applied:
                // otherwise a document using chapter/section numbering silently falls
                // back to a continuous number after conversion. Re-plan the existing
                // captions against the document's current format and patch only the
                // formulas whose rendered number actually changes.
                var stableSlotChangedNumbers =
                    UpdateNativeEquationSequenceFieldsIncremental(document);
                TraceStage("stable-slot-format-refresh");
                if (stableSlotChangedNumbers.Count > 0)
                {
                    UpdateHealthyNativeCrossReferencesAfterRenumbering(
                        document,
                        stableSlotChangedNumbers);
                    currentVisibleNumberRefreshed =
                        stableSlotChangedNumbers.ContainsKey(metadata.FormulaId);
                    TraceStage("stable-slot-references");
                }
            }
            else if (numberingOrderMayHaveChanged || !hadCompleteOwnedArtifacts)
            {
                if (TryUpdateAppendedNativeEquationSequenceField(
                        document,
                        metadata.FormulaId,
                        out var appendedNumberChanged))
                {
                    TraceStage("append-sequence-fast");
                    // A newly appended FormulaId cannot have pre-existing body
                    // cross-references. If its heading-aware caption changed,
                    // refresh only this formula owner's visible REF instead of scanning
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
                else if (TryUpdateInsertedEquationSequenceSuffixWithinScope(
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
                    Dictionary<string, string> fallbackChangedFormulaNumbers;
                    IReadOnlyDictionary<string, int>? fallbackReferenceCounts = null;
                    if (TryReadHealthyEquationNumberArtifactsFromOpenXml(
                            document,
                            out var fallbackOpenXmlCaptions,
                            out var knownFallbackReferenceCounts))
                    {
                        var fallbackFormat = ReadEquationNumberFormat(document);
                        var fallbackCaptions = fallbackFormat.UsesHeading
                            ? ResolveNativeEquationCaptionPositions(
                                document,
                                fallbackOpenXmlCaptions)
                            : fallbackOpenXmlCaptions;
                        if (fallbackCaptions.Count == fallbackOpenXmlCaptions.Count)
                        {
                            fallbackChangedFormulaNumbers =
                                UpdateNativeEquationSequenceFieldsIncremental(
                                    document,
                                    GetNativeEquationSequenceName(document),
                                    fallbackCaptions,
                                    fallbackFormat,
                                    trustedHealthyDirectTables: true);
                            fallbackReferenceCounts = knownFallbackReferenceCounts;
                            TraceStage("sequence-fallback-healthy");
                        }
                        else
                        {
                            fallbackChangedFormulaNumbers =
                                UpdateNativeEquationSequenceFieldsIncremental(document);
                            TraceStage("sequence-fallback");
                        }
                    }
                    else
                    {
                        fallbackChangedFormulaNumbers =
                            UpdateNativeEquationSequenceFieldsIncremental(document);
                        TraceStage("sequence-fallback");
                    }
                    if (fallbackChangedFormulaNumbers.Count > 0)
                    {
                        UpdateHealthyNativeCrossReferencesAfterRenumbering(
                            document,
                            fallbackChangedFormulaNumbers,
                            knownReferenceCounts: fallbackReferenceCounts);
                        TraceStage("cross-reference-fallback");
                    }
                }
            }

            var containsNativeOmml = ContainsNativeOmml(formulaRange);
            // Current numbered OMML is the direct-SEQ 1x3 host. Validate that
            // local structure first; only non-table legacy formulas need the much
            // more expensive retired mathematical #(SEQ) metadata/OpenXML probe.
            // On a 100-OMML document this avoids over a second of work after every
            // newly numbered formula.
            var isNativeDirectTable = containsNativeOmml
                && (visibleNumberCreated
                    || IsHealthyNativeOmmlDirectTableHost(
                        document,
                        formulaRange,
                        metadata.FormulaId,
                        updateField: false));
            var isNativeHashSequence = containsNativeOmml
                && !isNativeDirectTable
                && IsNumberedNativeOmmlHashSequenceFormula(
                    document,
                    metadata.FormulaId);
            // Legacy OLE/table/Shape hosts display their number through a generated
            // REF range and therefore still need local REF formatting. Native OMML
            // displays the SEQ result directly inside its managed host; applying
            // the old REF path only enumerates/normalizes the equation again.
            if (!visibleNumberCreated
                && !currentVisibleNumberRefreshed
                && !isNativeHashSequence
                && !isNativeDirectTable)
            {
                UpdateEquationNumberFields(
                    document,
                    formulaHeightPoints,
                    formulaFontSizePoints,
                    metadata.FormulaId);
            }
            if (containsNativeOmml
                && !deferNativeOmmlShapeFinalization
                && !isNativeHashSequence
                && !isNativeDirectTable)
                TryFinalizeNativeDisplayNumberShapeLayout(
                    document,
                    metadata.FormulaId);
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

        IReadOnlyList<NativeEquationCaptionEntry>? removalSnapshot = null;
        HashSet<string>? removalExternalReferenceIds = null;
        var canUseTargetedDirectRemoval = false;
        if (!string.IsNullOrWhiteSpace(preparedUnnumberedOmml)
            && TryReadTargetedRemovalSnapshot(
                document,
                metadata.FormulaId,
                out var targetedRemovalCaptions,
                out var targetedExternalReferenceIds))
        {
            removalSnapshot = targetedRemovalCaptions;
            removalExternalReferenceIds = targetedExternalReferenceIds;
            canUseTargetedDirectRemoval = true;
        }
        TraceStage("unnumber-snapshot");

        ConfigureUnnumberedDisplayFormula(
            document,
            formulaRange,
            metadata.FormulaId,
            metadata,
            preparedUnnumberedOmml,
            deferMetadataPersistence: deferNativeOmmlMetadataPersistence);
        TraceStage("unnumber-local-structure");

        if (canUseTargetedDirectRemoval
            && removalSnapshot is not null
            && removalExternalReferenceIds is not null
            && TryUpdateRemovedEquationSequenceSuffixWithinScope(
                document,
                metadata.FormulaId,
                removalSnapshot,
                out var removalChangedNumbers))
        {
            TraceStage("unnumber-suffix-fast");
            if (removalChangedNumbers.Keys.Any(
                    removalExternalReferenceIds.Contains))
            {
                UpdateHealthyNativeCrossReferencesAfterRenumbering(
                    document,
                    removalChangedNumbers);
                TraceStage("unnumber-references-fast");
            }
            return;
        }

        // Legacy, damaged, externally referenced, or otherwise ambiguous hosts
        // retain the mature comprehensive reconciliation behavior.
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

    internal static Range? FindNativeEquationCaptionRange(
        Document document,
        string formulaId)
    {
        Range? captionRange = null;
        Range? numberRange = null;
        try
        {
            if (!TryGetNativeCaptionRanges(
                    document,
                    formulaId,
                    GetNativeEquationSequenceName(document),
                    out captionRange,
                    out numberRange)
                || captionRange is null)
                return null;
            var result = captionRange.Duplicate;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(numberRange);
            Release(captionRange);
        }
    }

    internal static Range? FindVisibleEquationNumberRange(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Bookmark? formulaBookmark = null;
        Bookmark? repairedBookmark = null;
        Range? formulaRange = null;
        Field? nativeNumberField = null;
        Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = EquationBookmarkName(formulaId);
            if (bookmarks.Exists(name))
            {
                bookmark = bookmarks[name];
                range = bookmark.Range.Duplicate;
                var bookmarkedResult = range;
                range = null;
                return bookmarkedResult;
            }

            // Word may discard a bookmark nested inside a professional OMath when
            // it normalizes or serializes the equation. The generated REF field is
            // still durable, so recover its result from the formula identity rather
            // than treating the missing convenience bookmark as number loss.
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            if (formulaBookmark is null) return null;
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            if (!WordOmmlConverter.HasVisualTeXNativeEquationNumber(
                    formulaRange.WordOpenXML))
                return null;
            nativeNumberField = FindNativeOmmlNumberReferenceField(
                document,
                formulaRange,
                formulaId);
            if (nativeNumberField is null) return null;
            range = nativeNumberField.Result.Duplicate;
            if (!document.ReadOnly)
            {
                try { repairedBookmark = bookmarks.Add(name, range); }
                catch { }
            }
            var recoveredResult = range;
            range = null;
            return recoveredResult;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(range);
            Release(nativeNumberField);
            Release(formulaRange);
            Release(repairedBookmark);
            Release(formulaBookmark);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static Range? FindVisibleEquationNumberTextRange(
        Document document,
        string formulaId)
    {
        Range? range = null;
        try
        {
            range = FindVisibleEquationNumberRange(document, formulaId);
            if (range is null) return null;
            var text = range.Text ?? string.Empty;
            var trimStart = 0;
            while (trimStart < text.Length && text[trimStart] == '\t')
                trimStart++;
            var trimEnd = 0;
            while (trimEnd < text.Length - trimStart)
            {
                var character = text[text.Length - 1 - trimEnd];
                if (character != '\r' && character != '\a') break;
                trimEnd++;
            }
            if (trimStart > 0 || trimEnd > 0)
            {
                var start = Math.Min(range.End, range.Start + trimStart);
                var end = Math.Max(start, range.End - trimEnd);
                range.SetRange(start, end);
            }
            var result = range;
            range = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(range);
        }
    }

    internal static Range? FindNumberingOwnerRange(
        Document document,
        string formulaId)
    {
        Range? visibleRange = null;
        Tables? tables = null;
        Table? table = null;
        Columns? columns = null;
        Range? ownerRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        try
        {
            visibleRange = FindVisibleEquationNumberRange(document, formulaId);
            if (visibleRange is null) return null;
            if (visibleRange.StoryType == WdStoryType.wdTextFrameStory)
            {
                ownerRange = TryResolveNativeDisplayFormulaOwnerRange(
                    document,
                    formulaId);
                if (ownerRange is not null)
                {
                    var trueDisplayResult = ownerRange;
                    ownerRange = null;
                    return trueDisplayResult;
                }
            }
            try
            {
                if ((bool)visibleRange.get_Information(WdInformation.wdWithInTable))
                {
                    tables = visibleRange.Tables;
                    if (tables.Count > 0)
                    {
                        table = tables[1];
                        columns = table.Columns;
                        if (columns.Count >= 3)
                        {
                            ownerRange = table.Range.Duplicate;
                            var tableResult = ownerRange;
                            ownerRange = null;
                            return tableResult;
                        }
                    }
                }
            }
            catch
            {
                // Fall through to the paragraph owner. This is the native
                // MathType-style layout used by numbered VisualTeX OLE formulas.
            }

            paragraphs = visibleRange.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            ownerRange = paragraph.Range.Duplicate;
            var paragraphResult = ownerRange;
            ownerRange = null;
            return paragraphResult;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(paragraph);
            Release(paragraphs);
            Release(ownerRange);
            Release(columns);
            Release(table);
            Release(tables);
            Release(visibleRange);
        }
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

            if (ContainsNativeOmml(formulaRange)
                && NativeDisplayNumberShapeBelongsToFormula(
                    document,
                    formulaRange,
                    formulaId))
                return true;

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
        Range? content = null;
        try
        {
            content = document.Content;
            UpdateVisualTeXOwnedFieldsInRange(content);
        }
        finally { Release(content); }
    }

    private static int UpdateVisualTeXOwnedFieldsInRange(
        Range range,
        out int visualTeXReferenceCount)
    {
        visualTeXReferenceCount = 0;
        if (range is null) return 0;

        Document? document = null;
        Fields? fields = null;
        var updated = 0;
        try
        {
            document = range.Document;
            fields = range.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];

                    // Fast-reject the common OLE object itself before reading its
                    // code. In Word an embedded VisualTeX/MathType object is exposed
                    // as an EMBED field; updating it is unnecessary and can trigger
                    // the "fields may refer to other files" confirmation dialog.
                    if (field.Type == WdFieldType.wdFieldEmbed)
                        continue;

                    code = field.Code;
                    var codeText = code.Text ?? string.Empty;
                    if (IsVisualTeXSequenceFieldCode(codeText))
                    {
                        field.Update();
                        updated++;
                        continue;
                    }

                    if (field.Type != WdFieldType.wdFieldRef
                        || !TryResolveVisualTeXReferenceBookmark(
                            document,
                            codeText,
                            out var bookmarkName)
                        || !TryFormulaIdFromBookmark(
                            bookmarkName,
                            NativeNumberBookmarkPrefix,
                            out _))
                        continue;

                    field.Update();
                    visualTeXReferenceCount++;
                    updated++;
                }
                catch
                {
                    // A protected VisualTeX field can be repaired by the existing
                    // targeted reconciliation path. Never fall back to Fields.Update
                    // because that would also touch unrelated/external fields.
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            return updated;
        }
        finally
        {
            Release(fields);
            Release(document);
        }
    }

    private static int UpdateVisualTeXOwnedFieldsInRange(Range range) =>
        UpdateVisualTeXOwnedFieldsInRange(range, out _);

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
        var nativeTableTypingRange =
            EnsureNormalTypingParagraphAfterNativeOmmlTable(
                document,
                formulaId,
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
        string formulaId,
        Table? knownDirectTable = null)
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
                    formulaId);
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
            // Keep that structural terminator, but collapse it to one point so it no
            // longer appears as a full blank line below the last numbered formula.
            CompactTrailingTypingParagraph(document, paragraphRange);
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

    internal static void RemoveVisibleEquationNumberForFormula(
        Document document,
        string formulaId)
    {
        RemoveVisibleEquationNumber(document, formulaId);
    }

    internal static void RemoveFormulaNumberingArtifacts(
        Document document,
        string formulaId,
        bool preserveNativeCaptionParagraph = false)
    {
        Bookmark? formulaBookmark = null;
        Range? formulaRange = null;
        try
        {
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            if (formulaBookmark is not null)
            {
                formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
                if (WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                        formulaRange.WordOpenXML))
                {
                    // VTEqCap_/VTEqNum_ live inside the professional OMath in the
                    // native #SEQ design. Never delete their bookmarked ranges as
                    // ordinary caption text; strip the complete number wrapper in
                    // one OMML replacement instead.
                    ConfigureUnnumberedDisplayFormula(
                        document,
                        formulaRange,
                        formulaId);
                    return;
                }
            }
        }
        catch
        {
            // Fall back to the legacy artifact cleanup only when this is not a
            // recognizable native #SEQ host. Structural conversion paths perform
            // their own source-preservation checks before calling this method.
        }
        finally
        {
            Release(formulaRange);
            Release(formulaBookmark);
        }

        RemoveVisibleEquationNumber(document, formulaId);
        RemoveNativeCaption(
            document,
            formulaId,
            preserveParagraphSeparator: preserveNativeCaptionParagraph);
    }

    internal static int RefreshNumberedOmmlTabLayouts(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));

        // Complete VTEq_/VTEqCap_/VTEqNum_ loss makes the normal structural
        // resolver reject an otherwise valid managed #(SEQ) OMath. Repair those
        // identities first from VTOMML + semantic fingerprint; the ordinary loop
        // below can then use its strict FormulaId-bound range checks unchanged.
        TryRepairDriftedManagedNativeHashSequenceHosts(document, out _);

        var refreshed = 0;
        var formulaIds = WordOmmlFormulaStore.BookmarkedFormulaIds(document);
        foreach (var formulaId in formulaIds)
        {
            Range? formulaRange = null;
            try
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                if (metadata is null
                    || !string.Equals(
                        metadata.DisplayMode,
                        "block",
                        StringComparison.OrdinalIgnoreCase))
                    continue;
                formulaRange = WordOmmlFormulaStore
                    .GetEquationRangeVerifiedForStructuralEdit(
                        document,
                        formulaId,
                        metadata);
                if (!metadata.Numbered)
                {
                    // Older builds could dismantle a direct-SEQ 1x3 into a valid
                    // standalone Display OMath while accidentally retaining the
                    // adjacent VisualTeX table separator's Exactly-1pt paragraph
                    // metrics. Reuse this existing one-time document-open scan to
                    // self-heal only that unmistakable managed-OMML signature.
                    if (NormalizeCompactStandaloneNativeOmmlParagraph(formulaRange))
                        refreshed++;
                    continue;
                }
                ConfigureNumberedDisplayFormula(
                    document,
                    formulaRange,
                    WordOmmlFormulaStore.EstimateHeightPoints(formulaRange),
                    (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                    formulaId,
                    reuseExistingScaffold: true,
                    metadata: metadata,
                    deferNativeOmmlShapeCreation: true);
                // Repair documents produced by builds that retained the mandatory
                // separator between adjacent 1x3 tables at ordinary 10.5/12pt line
                // metrics. The paragraph cannot be deleted without merging the two
                // tables, but it can be made visually negligible and remains outside
                // every FormulaId/field/bookmark identity.
                CompactManagedNativeOmmlTableSeparatorBefore(document, formulaId);
                CleanupGeneratedNativeOmmlTypingTailBeforeFollowingTable(
                    document,
                    formulaId);
                refreshed++;
            }
            finally { Release(formulaRange); }
        }
        return refreshed;
    }

    internal static int FinalizeNumberedOmmlDisplayShapeLayouts(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        return FinalizeConvertedNumberedOmmlDisplayShapes(
            document,
            WordOmmlFormulaStore.BookmarkedFormulaIds(document));
    }

    internal static int FinalizeConvertedNumberedOmmlDisplayShapes(
        Document document,
        IReadOnlyList<string> formulaIds)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (formulaIds is null) throw new ArgumentNullException(nameof(formulaIds));
        var finalized = 0;
        foreach (var formulaId in formulaIds
                     .Where(id => !string.IsNullOrWhiteSpace(id))
                     .Distinct(StringComparer.OrdinalIgnoreCase))
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
                    continue;
                formulaRange = WordOmmlFormulaStore
                    .GetEquationRangeVerifiedForStructuralEdit(
                        document,
                        formulaId,
                        metadata);

                if (IsHealthyNativeOmmlDirectTableHost(
                        document,
                        formulaRange,
                        formulaId,
                        updateField: true))
                {
                    finalized++;
                    continue;
                }

                if (!IsNativeHashSequenceHostForFinalizationV2(
                        document,
                        formulaRange,
                        formulaId,
                        updateField: true))
                {
                    // This compatibility entry is still called after batch insert,
                    // document-open migration and old add-in dispatcher turns. It
                    // may repair/migrate a legacy host, but it must never create a
                    // floating object. ConfigureNumberedNativeOmmlDisplay now has
                    // only the native #() producer; old Shape/table/REF eqArr forms
                    // are consumed as migration input.
                    ConfigureNumberedNativeOmmlDisplay(
                        document,
                        formulaRange,
                        WordOmmlFormulaStore.EstimateHeightPoints(formulaRange),
                        (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                        formulaId,
                        reuseExistingScaffold: true,
                        metadata,
                        _ => { },
                        deferFieldUpdate: false,
                        deferExternalShapeCreation: true);
                    Release(formulaRange);
                    formulaRange = WordOmmlFormulaStore
                        .GetEquationRangeVerifiedForStructuralEdit(
                            document,
                            formulaId,
                            metadata);
                }

                if (IsHealthyNativeOmmlDirectTableHost(
                        document,
                        formulaRange,
                        formulaId,
                        updateField: true))
                {
                    finalized++;
                    continue;
                }
                if (!IsNativeHashSequenceHostForFinalizationV2(
                        document,
                        formulaRange,
                        formulaId,
                        updateField: true))
                    continue;

                // A just-migrated Shape host can expose a healthy native #(SEQ)
                // equation one Office turn before Word physically removes the old
                // one-point anchor paragraph. Word may also remove VTEqAnc_ while
                // retaining that empty paragraph, so completion checks both the
                // bookmark and the retired paragraph's exact structural/style
                // fingerprint before reporting the formula finalized.
                if (!EnsureRetiredNativeDisplayAnchorParagraphRemoved(
                        document,
                        formulaRange,
                        formulaId))
                    continue;
                ClearNativeDisplayAnchorCommitMarker(document, formulaId);
                finalized++;
            }
            finally
            {
                Release(formulaRange);
            }
        }
        return finalized;
    }

    private static void ScrollRangeIntoLayoutView(Window? window, Range range)
    {
        if (window is null) return;
        try
        {
            object start = true;
            window.ScrollIntoView(range, ref start);
        }
        catch
        {
            // Background/protected windows can reject scrolling. Width/geometry
            // readers retain their conservative fallback behavior in that case.
        }
    }

    private static bool HasNumberedOmmlLeftFormulaTab(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            format = paragraph.Format;
            tabStops = format.TabStops;
            var hasLeft = false;
            var hasRight = false;
            for (var index = 1; index <= tabStops.Count; index++)
            {
                Release(tabStop);
                tabStop = tabStops[index];
                if (tabStop.Alignment == WdTabAlignment.wdAlignTabLeft)
                    hasLeft = true;
                else if (tabStop.Alignment == WdTabAlignment.wdAlignTabRight)
                    hasRight = true;
            }
            return hasLeft && hasRight;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(tabStop);
            Release(tabStops);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
        }
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
        var adoptedNativeHashCopies =
            AdoptUnownedCopiesOfManagedNativeHashSequenceHosts(document);
        if (adoptedNativeHashCopies > 0)
            TraceNumberingPerformance(
                $"[perf] Reconcile.adopt-native-hash-copies count={adoptedNativeHashCopies}");
        TraceReconcileStage("adopt-native-hash-copies");

        // Freeze and verify every display-OMML identity before the first numbering
        // paragraph mutation, then reconcile from the end of the document toward
        // the start. Later tab/caption insertions therefore cannot drag the collapsed
        // bookmark of an equation that has not yet been processed.
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
                            metadata.FormulaId,
                            metadata: metadata);
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
                            formulaId,
                            metadata: metadata);
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
        bool useConversionSafeVisibleNumber = false,
        FormulaMetadata? metadata = null,
        bool deferNativeOmmlShapeCreation = false,
        bool deferNativeOmmlMetadataPersistence = false)
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

        if (ContainsNativeOmml(formulaRange))
        {
            return ConfigureNumberedNativeOmmlDisplay(
                document,
                formulaRange,
                formulaHeightPoints,
                formulaFontSizePoints,
                formulaId,
                reuseExistingScaffold,
                metadata,
                TraceStage,
                deferExternalShapeCreation: deferNativeOmmlShapeCreation,
                deferMetadataPersistence:
                    deferNativeOmmlMetadataPersistence);
        }

        Range? migratedFormulaRange = null;
        try
        {
            // Numbered VisualTeX OLE retains the accepted center/right-tab
            // paragraph contract. Native OMML has already branched to Word's own
            // display-equation number slot above. Migrate only the exact managed
            // legacy OLE table; arbitrary user tables remain untouched.
            if (IsNumberedEquationTable(formulaRange))
            {
                // A legacy 1x3/2x3 table is migration input only. Normalize the
                // known empty-row quirk, then require a complete conversion to the
                // final ordinary center/right-tab paragraph before doing anything
                // else. Retaining or repairing the table is no longer a supported
                // fallback.
                TrimBenignEmptyRowsFromNumberedTable(document, formulaRange, formulaId);
                migratedFormulaRange = TryConvertStandardNumberedOleTableToTabParagraph(
                    document,
                    formulaRange,
                    formulaId);
                if (migratedFormulaRange is null)
                    throw new InvalidOperationException(
                        "VisualTeX could not safely migrate this managed numbered OLE table to the table-free tab layout.");
                knownNumberedTable = null;
                TraceStage("migrate-ole-table-to-tabs");
            }

            var activeFormulaRange = migratedFormulaRange ?? formulaRange;
            if (IsNumberedEquationTable(activeFormulaRange))
                throw new InvalidOperationException(
                    "A numbered VisualTeX OLE formula remained inside a Word table after migration.");

            ConfigureEquationParagraph(activeFormulaRange, numbered: true);
            EnsureLeadingEquationTab(document, activeFormulaRange);
            TraceStage("tab-paragraph");

            var sequenceName = GetNativeEquationSequenceName(document);
            EnsureNativeCaption(
                document,
                activeFormulaRange,
                formulaId,
                sequenceName,
                restyleExisting: !reuseExistingScaffold,
                knownNumberedTable: null);
            TraceStage("native-caption");
            var visibleNumberCreated = EnsureVisibleEquationNumber(
                document,
                activeFormulaRange,
                formulaHeightPoints,
                formulaFontSizePoints,
                formulaId,
                adoptExistingTableReference: false,
                knownNumberedTable: null,
                useConversionSafeVisibleNumber);
            TraceStage("visible-ref");
            if (metadata is not null && !ContainsNativeOmml(activeFormulaRange))
            {
                AlignNumberedOleObjectToParagraphBaseline(
                    document,
                    activeFormulaRange,
                    formulaHeightPoints,
                    metadata);
                TraceStage("ole-baseline");
            }
            return visibleNumberCreated;
        }
        finally { Release(migratedFormulaRange); }
    }

    private static Range ResolveSingleNativeOmmlRange(Range formulaRange)
    {
        Document? document = null;
        Range? firstCharacter = null;
        Range? probe = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? range = null;
        try
        {
            document = formulaRange.Document;
            var start = formulaRange.Start;
            var end = formulaRange.End;

            // In the MathType-compatible numbering host Word tracks a Range across
            // text inserted immediately before it, so an old RCW can become
            // "TAB + OMath" even though the OMath itself still starts one character
            // later. Probe after that ordinary layout tab before asking Word for the
            // equation; otherwise OMath.Range is clipped to the caller's start and
            // VTOMML becomes anchored to the paragraph instead of the formula.
            if (start < end)
            {
                firstCharacter = document.Range(start, start + 1);
                if (string.Equals(firstCharacter.Text, "\t", StringComparison.Ordinal))
                    start++;
            }
            probe = document.Range(start, Math.Max(start, end));
            maths = probe.OMaths;
            if (maths.Count != 1)
                throw new InvalidOperationException(
                    $"The managed OMML host contains {maths.Count} equations instead of exactly one.");
            math = maths[1];
            range = math.Range.Duplicate;
            var result = range;
            range = null;
            return result;
        }
        finally
        {
            Release(range);
            Release(math);
            Release(maths);
            Release(probe);
            Release(firstCharacter);
            Release(document);
        }
    }

    private static Field? FindNativeOmmlNumberReferenceField(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        var targetBookmarkName = NativeNumberBookmarkName(formulaId);
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        Bookmark? visibleBookmark = null;
        Range? visibleRange = null;
        Field? positionalMatch = null;
        try
        {
            bookmarks = document.Bookmarks;
            var visibleBookmarkName = EquationBookmarkName(formulaId);
            if (bookmarks.Exists(visibleBookmarkName))
            {
                visibleBookmark = bookmarks[visibleBookmarkName];
                visibleRange = visibleBookmark.Range;
            }
            fields = formulaRange.Fields;
            var formulaText = formulaRange.Text ?? string.Empty;
            var separatorOffset = formulaText.LastIndexOf('#');
            var separatorPosition = separatorOffset >= 0
                ? formulaRange.Start + separatorOffset
                : formulaRange.Start;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? candidate = null;
                Range? candidateCode = null;
                Range? candidateResult = null;
                try
                {
                    candidate = fields[index];
                    candidateCode = candidate.Code;
                    candidateResult = candidate.Result;
                    if (IsReferenceToBookmark(candidateCode.Text, targetBookmarkName)
                        || (visibleRange is not null
                            && candidateResult.Start <= visibleRange.Start
                            && candidateResult.End >= visibleRange.End))
                    {
                        var result = candidate;
                        candidate = null;
                        return result;
                    }
                    if (candidateResult.Start > separatorPosition)
                    {
                        Release(positionalMatch);
                        positionalMatch = candidate;
                        candidate = null;
                    }
                }
                finally
                {
                    Release(candidateResult);
                    Release(candidateCode);
                    Release(candidate);
                }
            }
            var fallback = positionalMatch;
            positionalMatch = null;
            return fallback;
        }
        finally
        {
            Release(positionalMatch);
            Release(visibleRange);
            Release(visibleBookmark);
            Release(bookmarks);
            Release(fields);
        }
    }

    private static void RemoveOrdinaryTabsAfterNativeOmml(
        Document document,
        Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? probe = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var end = Math.Max(formulaRange.End, paragraphRange.End - 1);
            for (var position = end - 1; position >= formulaRange.End; position--)
            {
                Release(probe);
                probe = document.Range(position, position + 1);
                if (string.Equals(probe.Text, "\t", StringComparison.Ordinal))
                    probe.Delete();
            }
        }
        finally
        {
            Release(probe);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
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
                math.Type = WdOMathType.wdOMathDisplay;
            try { math.Justification = WdOMathJc.wdOMathJcCenterGroup; } catch { }
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

    private static bool ContainsNativeOmml(Range formulaRange)
    {
        OMaths? maths = null;
        try
        {
            maths = formulaRange.OMaths;
            return maths.Count > 0;
        }
        catch { return false; }
        finally { Release(maths); }
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

    private static Range? TryConvertStandardNumberedOmmlTableToStandaloneDisplayParagraph(
        Document document,
        Range formulaRange,
        string formulaId,
        FormulaMetadata? metadata,
        string? preparedStandaloneOmml = null)
    {
        Tables? tables = null;
        Table? table = null;
        Rows? rows = null;
        Columns? columns = null;
        Range? tableRange = null;
        OMaths? tableMaths = null;
        InlineShapes? tableShapes = null;
        Cell? leftCell = null;
        Cell? formulaCell = null;
        Cell? numberCell = null;
        Range? leftRange = null;
        Range? formulaCellRange = null;
        Range? numberCellRange = null;
        OMaths? formulaCellMaths = null;
        OMath? formulaCellMath = null;
        InlineShapes? formulaCellShapes = null;
        Fields? formulaCellFields = null;
        OMaths? numberCellMaths = null;
        InlineShapes? numberCellShapes = null;
        Fields? numberCellFields = null;
        Range? visibleNumberRange = null;
        Fields? visibleNumberFields = null;
        Range? equationRange = null;
        Microsoft.Office.Interop.Word.Application? application = null;
        Document? stagingDocument = null;
        Range? stagingInsertion = null;
        OMaths? stagingMaths = null;
        OMath? stagingMath = null;
        Range? stagingMathRange = null;
        Range? convertedRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? editableRange = null;
        Range? placeholderRange = null;
        OMaths? migratedMaths = null;
        OMath? migratedMath = null;
        Range? migratedMathRange = null;
        Bookmark? oldFormulaBookmark = null;
        var migrationWatch = System.Diagnostics.Stopwatch.StartNew();
        long migrationCheckpoint = 0;
        void TraceMigration(string stage)
        {
            if (!string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                    "1",
                    StringComparison.Ordinal))
                return;
            var elapsed = migrationWatch.ElapsedMilliseconds;
            TraceNumberingPerformance(
                $"      [perf] unnumber-local.{stage}: +{elapsed - migrationCheckpoint}ms ({elapsed}ms)");
            migrationCheckpoint = elapsed;
        }
        Range? Fail(string reason)
        {
            TraceNumberingPerformance(
                $"[migration] legacy-omml-table-rejected formulaId={formulaId} reason={reason}");
            return null;
        }
        try
        {
            if (!IsNumberedEquationTable(formulaRange) || !ContainsNativeOmml(formulaRange))
                return Fail("range-not-numbered-omml-table");
            TraceMigration("initial-host-probe");

            tables = formulaRange.Tables;
            if (tables.Count == 0) return Fail("range-table-collection-empty");
            table = tables[1];
            rows = table.Rows;
            columns = table.Columns;
            // Flatten only the exact VisualTeX-managed legacy host. User-created
            // equation tables, nested tables and malformed structures are never
            // rewritten merely because they happen to contain an OMath. Historical
            // builds may have appended wholly empty rows; those are validated and
            // trimmed below before conversion.
            if (rows.Count < 1 || columns.Count != 3)
                return Fail($"unexpected-dimensions rows={rows.Count} columns={columns.Count}");

            tableRange = table.Range;
            tableMaths = tableRange.OMaths;
            tableShapes = tableRange.InlineShapes;
            if (tableMaths.Count != 1 || tableShapes.Count != 0)
                return Fail($"table-payload maths={tableMaths.Count} shapes={tableShapes.Count}");

            leftCell = table.Cell(1, 1);
            formulaCell = table.Cell(1, 2);
            numberCell = table.Cell(1, 3);
            leftRange = leftCell.Range.Duplicate;
            formulaCellRange = formulaCell.Range.Duplicate;
            numberCellRange = numberCell.Range.Duplicate;
            if (!ContainsOnlyStructuralWordText(leftRange.Text))
                return Fail("left-cell-has-nonstructural-text");
            TraceMigration("table-and-cells");

            formulaCellMaths = formulaCellRange.OMaths;
            formulaCellShapes = formulaCellRange.InlineShapes;
            formulaCellFields = formulaCellRange.Fields;
            if (formulaCellMaths.Count != 1
                || formulaCellShapes.Count != 0
                || formulaCellFields.Count != 0)
                return Fail(
                    $"formula-cell-payload maths={formulaCellMaths.Count} shapes={formulaCellShapes.Count} fields={formulaCellFields.Count}");

            metadata ??= WordOmmlFormulaStore.TryRead(document, formulaId);
            if (metadata is null
                || !string.Equals(
                    metadata.FormulaId,
                    formulaId,
                    StringComparison.OrdinalIgnoreCase))
                return Fail("metadata-missing-or-formulaid-mismatch");
            // The caller already holds the just-created live OMath Range. During
            // OLE→OMML replacement Word can temporarily expand VTOMML across the
            // surrounding legacy table, so resolving through that bookmark again
            // may report a non-unique fingerprint even though this exact center cell
            // contains one and only one native equation. Resolve that unique OMath
            // directly from the fully validated cell and use the caller range only
            // as an ownership cross-check.
            formulaCellMath = formulaCellMaths[1];
            equationRange = formulaCellMath.Range.Duplicate;
            if (equationRange.Start < formulaCellRange.Start
                || equationRange.End > formulaCellRange.End)
                return Fail(
                    $"equation-outside-center-cell equation={equationRange.Start}:{equationRange.End} cell={formulaCellRange.Start}:{formulaCellRange.End}");
            var callerOverlapsEquation = formulaRange.Start < equationRange.End
                && formulaRange.End > equationRange.Start;
            var callerMatchesCollapsedBoundary = formulaRange.Start == equationRange.Start
                && formulaRange.End == equationRange.End;
            if (!callerOverlapsEquation && !callerMatchesCollapsedBoundary)
                return Fail(
                    $"caller-range-does-not-own-center-equation caller={formulaRange.Start}:{formulaRange.End} equation={equationRange.Start}:{equationRange.End}");
            TraceMigration("center-equation");
            // The converter-side fingerprint was captured before Word imported and
            // normalized this OMath, so it can legitimately differ at this exact
            // intermediate point. Identity is already proven by the caller-range
            // overlap plus the one-OMath-only managed center cell; the durable live
            // fingerprint is stamped after the completed #(SEQ) replacement.

            // Some historical VisualTeX documents contain one or more wholly empty
            // trailing rows below the standard first 1x3 formula row. Remove only
            // rows proven empty by the existing defensive row cleaner, then require
            // the exact 1x3 managed host before any table-to-text conversion. A
            // non-empty user row is never deleted and causes migration to stop.
            if (rows.Count > 1)
            {
                RemoveEmptyTrailingNumberedTableRows(table);
                Release(rows);
                rows = table.Rows;
                if (rows.Count != 1)
                    return Fail($"nonempty-extra-rows remaining={rows.Count}");
            }

            // Legacy callers may need to preserve the existing live OMath exactly,
            // so keep the mature hidden-Word FormattedText staging fallback. The
            // interactive ReplaceOmml path already owns the clean target semantic
            // OMML, however; staging that same equation in another document adds a
            // costly Activate/Documents.Add/FormattedText round-trip for no value.
            application = document.Application;
            if (string.IsNullOrWhiteSpace(preparedStandaloneOmml))
            {
                stagingDocument = application.Documents.Add(Visible: false);
                stagingInsertion = stagingDocument.Content;
                stagingInsertion.FormattedText = equationRange.FormattedText;
                stagingMaths = stagingDocument.OMaths;
                if (stagingMaths.Count != 1)
                    throw new InvalidOperationException(
                        "Word could not stage exactly one native OMath before numbering-table migration.");
                stagingMath = stagingMaths[1];
                stagingMathRange = stagingMath.Range.Duplicate;
                try { stagingDocument.Saved = true; } catch { }
                document.Activate();
            }

            numberCellMaths = numberCellRange.OMaths;
            numberCellShapes = numberCellRange.InlineShapes;
            numberCellFields = numberCellRange.Fields;
            if (numberCellMaths.Count != 0 || numberCellShapes.Count != 0)
                return Fail(
                    $"number-cell-object-payload maths={numberCellMaths.Count} shapes={numberCellShapes.Count}");
            visibleNumberRange = FindVisibleEquationNumberRange(document, formulaId);
            if (visibleNumberRange is null)
            {
                if (numberCellFields.Count != 0
                    || !ContainsOnlyStructuralWordText(numberCellRange.Text))
                    return Fail(
                        $"number-cell-without-owned-ref fields={numberCellFields.Count} text='{NormalizeNativeEquationNumberText(numberCellRange.Text)}'");
            }
            else
            {
                if (visibleNumberRange.Start < numberCellRange.Start
                    || visibleNumberRange.End > numberCellRange.End
                    || !ContainsOnlyStructuralWordTextOutsideRange(
                        document,
                        numberCellRange,
                        visibleNumberRange))
                    return Fail(
                        $"owned-ref-outside-number-cell visible={visibleNumberRange.Start}:{visibleNumberRange.End} cell={numberCellRange.Start}:{numberCellRange.End}");
                visibleNumberFields = visibleNumberRange.Fields;
                if (numberCellFields.Count != visibleNumberFields.Count)
                    return Fail(
                        $"number-cell-field-count-mismatch cell={numberCellFields.Count} visible={visibleNumberFields.Count}");
            }
            TraceMigration("number-cell-ownership");

            // Remove only the generated visible REF label before flattening. The
            // hidden VTEqNum_* SEQ target and every external REF remain intact.
            RemoveVisibleEquationNumber(document, formulaId);
            // Table.ConvertToText can expand its returned range across the hidden
            // caption paragraph that immediately follows the table. Reusing that
            // live SEQ/bookmark while replacing the converted range can make the
            // visible REF point to itself. Remove only this formula's hidden caption
            // now; the normal reconciliation path recreates the same VTEqCap/VTEqNum
            // identities before any external reference is updated.
            RemoveNativeCaption(document, formulaId);
            oldFormulaBookmark = WordOmmlFormulaStore.FindByFormulaId(
                document,
                formulaId);
            try { oldFormulaBookmark?.Delete(); } catch { }
            TraceMigration("remove-number-artifacts");

            object separator = WdTableFieldSeparator.wdSeparateByTabs;
            object nestedTables = false;
            convertedRange = table.ConvertToText(ref separator, ref nestedTables);
            TraceMigration("convert-to-text");
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                    "1",
                    StringComparison.Ordinal))
            {
                OMaths? convertedMathsProbe = null;
                OMath? convertedMathProbe = null;
                try
                {
                    convertedMathsProbe = convertedRange.OMaths;
                    var convertedType = -1;
                    if (convertedMathsProbe.Count == 1)
                    {
                        convertedMathProbe = convertedMathsProbe[1];
                        convertedType = (int)convertedMathProbe.Type;
                    }
                    Console.WriteLine(
                        $"      [perf] unnumber-convert-to-text maths={convertedMathsProbe.Count} type={convertedType} range={convertedRange.Start}:{convertedRange.End}");
                }
                finally
                {
                    Release(convertedMathProbe);
                    Release(convertedMathsProbe);
                }
            }

            var paragraphStart = convertedRange.Start;
            if (!string.IsNullOrWhiteSpace(preparedStandaloneOmml))
            {
                // Current direct-SEQ 1x3 hosts are already fully owned and were
                // validated above. Word's ConvertToText preserves the center OMath
                // itself and only demotes it to Inline while inserting the two cell
                // separator TABs. Keep that exact live OMath: remove only structural
                // text around it, then restore Display type. Replacing the whole
                // converted range with a placeholder and importing the same OMath a
                // second time cost ~0.5-0.8s per checkbox toggle.
                Release(migratedMaths); migratedMaths = convertedRange.OMaths;
                if (migratedMaths.Count != 1)
                    throw new InvalidOperationException(
                        "Word did not preserve exactly one OMath while removing its numbering table.");
                Release(migratedMath); migratedMath = migratedMaths[1];
                Release(migratedMathRange); migratedMathRange = migratedMath.Range.Duplicate;

                editableRange = document.Range(
                    migratedMathRange.End,
                    Math.Max(migratedMathRange.End, convertedRange.End - 1));
                if (!ContainsOnlyStructuralWordText(editableRange.Text))
                    throw new InvalidOperationException(
                        "The converted number-table suffix contains user text.");
                editableRange.Delete();
                Release(editableRange); editableRange = null;

                placeholderRange = document.Range(
                    convertedRange.Start,
                    migratedMathRange.Start);
                if (!ContainsOnlyStructuralWordText(placeholderRange.Text))
                    throw new InvalidOperationException(
                        "The converted number-table prefix contains user text.");
                placeholderRange.Delete();
                Release(placeholderRange); placeholderRange = null;
                TraceMigration("strip-cell-separators");

                var directProbe = document.Range(
                    paragraphStart,
                    Math.Min(document.Content.End, paragraphStart + 1));
                try
                {
                    Release(paragraphs); paragraphs = directProbe.Paragraphs;
                    if (paragraphs.Count != 1)
                        throw new InvalidOperationException(
                            "Word did not leave one standalone paragraph after removing the numbering table.");
                    Release(paragraph); paragraph = paragraphs[1];
                    Release(paragraphRange); paragraphRange = paragraph.Range;
                    Release(migratedMaths); migratedMaths = paragraphRange.OMaths;
                    if (migratedMaths.Count != 1)
                        throw new InvalidOperationException(
                            "Word lost the preserved OMath after removing table separators.");
                    Release(migratedMath); migratedMath = migratedMaths[1];
                    if (migratedMath.Type != WdOMathType.wdOMathDisplay)
                        migratedMath.Type = WdOMathType.wdOMathDisplay;
                    Release(migratedMathRange); migratedMathRange = migratedMath.Range.Duplicate;
                    Release(equationRange); equationRange = migratedMathRange.Duplicate;
                    TraceMigration("restore-display-range");
                    if (string.Equals(
                            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                            "1",
                            StringComparison.Ordinal))
                    {
                        ParagraphFormat? probeFormat = null;
                        TabStops? probeTabs = null;
                        try
                        {
                            probeFormat = paragraph.Format;
                            probeTabs = probeFormat.TabStops;
                            TraceNumberingPerformance(
                                $"      [perf] unnumber-direct-paragraph alignment={(int)probeFormat.Alignment} left={probeFormat.LeftIndent:0.###} right={probeFormat.RightIndent:0.###} first={probeFormat.FirstLineIndent:0.###} lineRule={(int)probeFormat.LineSpacingRule} line={probeFormat.LineSpacing:0.###} before={probeFormat.SpaceBefore:0.###} after={probeFormat.SpaceAfter:0.###} tabs={probeTabs.Count} keep={probeFormat.KeepTogether}/{probeFormat.KeepWithNext} pageBreak={probeFormat.PageBreakBefore} widow={probeFormat.WidowControl}");
                        }
                        finally
                        {
                            Release(probeTabs);
                            Release(probeFormat);
                        }
                    }
                    var directResult = equationRange.Duplicate;
                    return directResult;
                }
                finally { Release(directProbe); }
            }

            // Legacy compatibility path: rebuild the complete range returned by
            // Table.ConvertToText from a staged OMath.
            const string placeholderText = "\uE000";
            convertedRange.Text = placeholderText + "\r";
            var paragraphProbe = document.Range(
                paragraphStart,
                Math.Min(document.Content.End, paragraphStart + 1));
            try
            {
                paragraphs = paragraphProbe.Paragraphs;
                if (paragraphs.Count != 1)
                    throw new InvalidOperationException(
                        "Word did not create one stable paragraph while rebuilding the numbered OMML host.");
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
            }
            finally { Release(paragraphProbe); }

            // Restore the staged professional equation directly into one ordinary
            // table-free paragraph. Do not create the retired TAB + inline OMath
            // intermediate host: numbered OMML must remain genuine Word display math
            // throughout migration, before the outer reconciler adds the dedicated
            // external REF Shape/anchor structure.
            var editableEnd = Math.Max(paragraphStart, paragraphRange.End - 1);
            editableRange = document.Range(paragraphStart, editableEnd);
            editableRange.Text = placeholderText;
            placeholderRange = document.Range(
                paragraphStart,
                paragraphStart + placeholderText.Length);
            if (!string.IsNullOrWhiteSpace(preparedStandaloneOmml))
            {
                Range? preparedRange = null;
                try
                {
                    string mathFontName;
                    try { mathFontName = document.OMathFontName ?? string.Empty; }
                    catch { mathFontName = string.Empty; }
                    preparedRange = WordOmmlConverter.ReplaceWithPreparedOmml(
                        application,
                        document,
                        placeholderRange,
                        preparedStandaloneOmml!,
                        display: true,
                        mathFontName);
                }
                finally { Release(preparedRange); }
            }
            else
            {
                placeholderRange.FormattedText = stagingMathRange.FormattedText;
            }

            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            migratedMaths = paragraphRange.OMaths;
            if (migratedMaths.Count != 1)
                throw new InvalidOperationException(
                    "Word did not restore exactly one native OMath while migrating its numbering table.");
            migratedMath = migratedMaths[1];
            if (migratedMath.Type != WdOMathType.wdOMathDisplay)
                migratedMath.Type = WdOMathType.wdOMathDisplay;
            migratedMathRange = migratedMath.Range;
            equationRange = migratedMathRange.Duplicate;

            // Do not bind metadata here. The outer numbered-OMML reconciler still
            // creates the final anchor/Shape/REF scaffold and owns the one durable
            // bookmark/fingerprint save after that physical host is complete.
            var result = equationRange.Duplicate;
            return result;
        }
        finally
        {
            Release(oldFormulaBookmark);
            Release(migratedMathRange);
            Release(migratedMath);
            Release(migratedMaths);
            Release(placeholderRange);
            Release(editableRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(convertedRange);
            Release(equationRange);
            Release(stagingMathRange);
            Release(stagingMath);
            Release(stagingMaths);
            Release(stagingInsertion);
            if (stagingDocument is not null)
            {
                try
                {
                    stagingDocument.Saved = true;
                    stagingDocument.Close(WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(stagingDocument);
            Release(application);
            Release(visibleNumberFields);
            Release(visibleNumberRange);
            Release(numberCellFields);
            Release(numberCellShapes);
            Release(numberCellMaths);
            Release(formulaCellFields);
            Release(formulaCellShapes);
            Release(formulaCellMath);
            Release(formulaCellMaths);
            Release(numberCellRange);
            Release(formulaCellRange);
            Release(leftRange);
            Release(numberCell);
            Release(formulaCell);
            Release(leftCell);
            Release(tableShapes);
            Release(tableMaths);
            Release(tableRange);
            Release(columns);
            Release(rows);
            Release(table);
            Release(tables);
        }
    }

    private static bool ContainsOnlyStructuralWordText(string? text)
    {
        if (string.IsNullOrEmpty(text)) return true;
        foreach (var character in text!)
        {
            if (character is '\r' or '\n' or '\t' or '\v' or '\a'
                or '\u0001' or '\u200B' or '\u200C' or '\u2060' or '\uFEFF')
                continue;
            if (!char.IsWhiteSpace(character)) return false;
        }
        return true;
    }

    private static bool ContainsOnlyStructuralWordTextOutsideRange(
        Document document,
        Range container,
        Range allowed)
    {
        Range? before = null;
        Range? after = null;
        try
        {
            if (allowed.Start < container.Start || allowed.End > container.End)
                return false;
            if (container.Start < allowed.Start)
            {
                before = document.Range(container.Start, allowed.Start);
                if (!ContainsOnlyStructuralWordText(before.Text)) return false;
            }
            if (allowed.End < container.End)
            {
                after = document.Range(allowed.End, container.End);
                if (!ContainsOnlyStructuralWordText(after.Text)) return false;
            }
            return true;
        }
        finally
        {
            Release(after);
            Release(before);
        }
    }

    private static Range? TryConvertStandardNumberedOleTableToTabParagraph(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Tables? tables = null;
        Table? table = null;
        Rows? rows = null;
        Columns? columns = null;
        Range? tableRange = null;
        InlineShapes? tableShapes = null;
        InlineShape? tableShape = null;
        Range? convertedRange = null;
        InlineShapes? convertedShapes = null;
        InlineShape? convertedShape = null;
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? cleanupRange = null;
        try
        {
            if (!IsNumberedEquationTable(formulaRange) || ContainsNativeOmml(formulaRange))
                return null;

            tables = formulaRange.Tables;
            if (tables.Count == 0) return null;
            table = tables[1];
            rows = table.Rows;
            columns = table.Columns;
            // Historical 2x3 hosts can carry one or more entirely empty trailing
            // rows. Normalize only those rows on this exact table before requiring
            // the final 1x3 migration input. This does not rely on numbering
            // bookmarks, which some oldest OLE documents never stored.
            if (columns.Count == 3 && rows.Count > 1)
            {
                RemoveEmptyTrailingNumberedTableRows(table);
                Release(rows);
                rows = table.Rows;
            }
            // Restrict the compatibility migration to the exact VisualTeX scaffold.
            // Custom or malformed user tables are left untouched rather than flattened.
            if (rows.Count != 1 || columns.Count != 3) return null;

            tableRange = table.Range;
            tableShapes = tableRange.InlineShapes;
            if (tableShapes.Count != 1) return null;
            tableShape = tableShapes[1];
            var metadata = ReadMetadata(tableShape);
            if (metadata is null
                || !string.Equals(metadata.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase))
                return null;

            object separator = WdTableFieldSeparator.wdSeparateByTabs;
            object nestedTables = false;
            convertedRange = table.ConvertToText(ref separator, ref nestedTables);
            convertedShapes = convertedRange.InlineShapes;
            for (var index = 1; index <= convertedShapes.Count; index++)
            {
                Release(convertedShape);
                convertedShape = convertedShapes[index];
                var convertedMetadata = ReadMetadata(convertedShape);
                if (convertedMetadata is not null
                    && string.Equals(
                        convertedMetadata.FormulaId,
                        formulaId,
                        StringComparison.OrdinalIgnoreCase))
                    break;
                Release(convertedShape);
                convertedShape = null;
            }
            if (convertedShape is null)
                throw new InvalidOperationException(
                    "Word converted the numbered VisualTeX table, but the OLE formula could not be resolved afterwards.");

            // Discard the table-cell separators and the old visible REF label, then let
            // the normal paragraph path rebuild exactly one leading and one trailing tab.
            RemoveVisibleEquationNumber(document, formulaId);
            shapeRange = convertedShape.Range;
            paragraphs = shapeRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (shapeRange.End < paragraphRange.End - 1)
            {
                cleanupRange = document.Range(shapeRange.End, paragraphRange.End - 1);
                cleanupRange.Delete();
                Release(cleanupRange);
                cleanupRange = null;
            }

            Release(shapeRange);
            shapeRange = convertedShape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            if (paragraphRange.Start < shapeRange.Start)
            {
                cleanupRange = document.Range(paragraphRange.Start, shapeRange.Start);
                cleanupRange.Delete();
                Release(cleanupRange);
                cleanupRange = null;
            }

            Release(shapeRange);
            shapeRange = convertedShape.Range;
            var result = shapeRange.Duplicate;
            return result;
        }
        finally
        {
            Release(cleanupRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
            Release(convertedShape);
            Release(convertedShapes);
            Release(convertedRange);
            Release(tableShape);
            Release(tableShapes);
            Release(tableRange);
            Release(columns);
            Release(rows);
            Release(table);
            Release(tables);
        }
    }

    private static void ConfigureUnnumberedDisplayFormula(
        Document document,
        Range formulaRange,
        string formulaId,
        FormulaMetadata? knownMetadata = null,
        string? preparedStandaloneOmml = null,
        bool deferMetadataPersistence = false)
    {
        Range? activeRange = null;
        Bookmark? repairedBookmark = null;
        var unnumberWatch = System.Diagnostics.Stopwatch.StartNew();
        long unnumberCheckpoint = 0;
        void TraceUnnumber(string stage)
        {
            if (!string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                    "1",
                    StringComparison.Ordinal))
                return;
            var elapsed = unnumberWatch.ElapsedMilliseconds;
            TraceNumberingPerformance(
                $"      [perf] unnumber-post.{stage}: +{elapsed - unnumberCheckpoint}ms ({elapsed}ms)");
            unnumberCheckpoint = elapsed;
        }
        try
        {
            if (ContainsNativeOmml(formulaRange))
            {
                activeRange = ResolveSingleNativeOmmlRange(formulaRange);
                TraceUnnumber("initial-resolve");
                var metadata = knownMetadata
                    ?? WordOmmlFormulaStore.TryRead(document, formulaId);
                if (IsNumberedEquationTable(activeRange))
                {
                    Range? standaloneRange = null;
                    try
                    {
                        standaloneRange = TryConvertStandardNumberedOmmlTableToStandaloneDisplayParagraph(
                            document,
                            activeRange,
                            formulaId,
                            metadata,
                            preparedStandaloneOmml);
                        if (standaloneRange is null)
                            throw new InvalidOperationException(
                                "VisualTeX could not safely remove the managed 1x3 OMML numbering table.");
                        Release(activeRange);
                        // TryConvertStandardNumberedOmmlTableToStandaloneDisplayParagraph
                        // returns the exact preserved live Display OMath Range after
                        // ConvertToText and separator cleanup. Re-resolving the same
                        // OMath through Range.OMaths costs ~100ms in a 100-OMML
                        // document and adds no identity information here.
                        activeRange = standaloneRange;
                        standaloneRange = null;
                        TraceUnnumber("convert-live-range");
                    }
                    finally { Release(standaloneRange); }
                    if (!string.IsNullOrWhiteSpace(preparedStandaloneOmml))
                    {
                        FinalizeDirectUnnumberedOmmlParagraph(activeRange);
                        TraceUnnumber("direct-paragraph-finalize");
                    }
                    else
                    {
                        NormalizeCompactStandaloneNativeOmmlParagraph(activeRange);
                        TraceUnnumber("compact-paragraph");
                        ConfigureEquationParagraph(activeRange, numbered: false);
                        TraceUnnumber("configure-paragraph");
                        EnsureNumberedOmmlIsDisplay(activeRange);
                        TraceUnnumber("ensure-display");
                    }
                    if (metadata is not null && !deferMetadataPersistence)
                    {
                        repairedBookmark = WordOmmlFormulaStore.Wrap(
                            document,
                            activeRange,
                            metadata,
                            replaceExisting: true);
                        TraceUnnumber("wrap-anchor");
                        WordOmmlNativeSource.StampFingerprintFromResolvedRange(metadata, activeRange);
                        WordOmmlFormulaStore.Save(document, metadata);
                    }
                    formulaRange.SetRange(activeRange.Start, activeRange.End);
                    // Table dismantling stages the professional OMath through a
                    // temporary hidden Word document. Re-activate the real target
                    // after that staging document is closed so the next immediate
                    // VisualTeX edit does not observe a transient/no ActiveDocument.
                    try { document.Activate(); } catch { }
                    return;
                }
                if (WordOmmlConverter.HasVisualTeXNativeEquationNumber(
                        activeRange.WordOpenXML)
                    || HasManagedNativeOmmlHashSequenceHost(
                        document,
                        formulaId))
                {
                    RemoveVisibleEquationNumber(document, formulaId);
                    var semanticOmml = WordOmmlConverter
                        .StripManagedVisualTeXNativeEquationNumber(
                            activeRange.WordOpenXML);
                    Microsoft.Office.Interop.Word.Application? application = null;
                    Range? replacementRange = null;
                    try
                    {
                        application = document.Application;
                        string mathFontName;
                        try { mathFontName = document.OMathFontName ?? string.Empty; }
                        catch { mathFontName = string.Empty; }
                        replacementRange = WordOmmlConverter.ReplaceWithPreparedOmml(
                            application,
                            document,
                            activeRange,
                            semanticOmml,
                            display: true,
                            mathFontName);
                        Release(activeRange);
                        activeRange = replacementRange;
                        replacementRange = null;
                    }
                    finally
                    {
                        Release(replacementRange);
                        Release(application);
                    }
                }
                else
                {
                    RemoveVisibleEquationNumber(document, formulaId);
                }

                RemoveNativeCaption(document, formulaId);
                RemoveLeadingEquationTab(document, activeRange);
                RemoveOrdinaryTabsAfterNativeOmml(document, activeRange);
                ConfigureEquationParagraph(activeRange, numbered: false);
                EnsureNumberedOmmlIsDisplay(activeRange);
                if (metadata is not null)
                {
                    repairedBookmark = WordOmmlFormulaStore.Wrap(
                        document,
                        activeRange,
                        metadata,
                        replaceExisting: true);
                    if (!deferMetadataPersistence)
                    {
                        WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                            metadata,
                            activeRange);
                        WordOmmlFormulaStore.Save(document, metadata);
                    }
                }
                formulaRange.SetRange(activeRange.Start, activeRange.End);
                return;
            }

            RemoveVisibleEquationNumber(document, formulaId);
            RemoveNativeCaption(document, formulaId);
            RemoveLeadingEquationTab(document, formulaRange);
            ResetOleObjectBaselinePosition(document, formulaRange);
            ConfigureEquationParagraph(formulaRange, numbered: false);
        }
        finally
        {
            Release(repairedBookmark);
            Release(activeRange);
        }
    }

    private static void AlignNumberedOleObjectToParagraphBaseline(
        Document document,
        Range formulaRange,
        float actualHeightPoints,
        FormulaMetadata metadata)
    {
        Microsoft.Office.Interop.Word.Font? rangeFont = null;
        Range? character = null;
        Microsoft.Office.Interop.Word.Font? characterFont = null;
        try
        {
            // VisualTeX OLE and MathType are different objects, but Word lays both
            // out as an EMBED field whose U+0001 result character owns the visible
            // object box. Keep the field instruction at the ordinary paragraph
            // baseline and move only that object-result character. The offset is
            // calculated from VisualTeX's own exported MathJax baseline, not from
            // MathType's storage or an arbitrary half-height approximation.
            rangeFont = formulaRange.Font;
            rangeFont.Position = 0;

            var semanticFontSize = FormulaFontSize.ResolveSemanticFontSize(metadata);
            var exportedHeight = metadata.RenderHeightPx.HasValue
                ? (float)metadata.RenderHeightPx.Value
                : 0f;
            float? exportedBaseline = metadata.Baseline.HasValue
                ? (float)metadata.Baseline.Value
                : null;
            var position = WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                actualHeightPoints,
                exportedHeight,
                exportedBaseline,
                existingFontPosition: null,
                sourceSemanticFontSizePoints: semanticFontSize,
                targetSemanticFontSizePoints: semanticFontSize);
            var clamped = Math.Max(-256, Math.Min(256, position));

            for (var index = formulaRange.Start; index < formulaRange.End; index++)
            {
                Release(characterFont);
                characterFont = null;
                Release(character);
                character = document.Range(index, index + 1);
                if (!string.Equals(character.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                characterFont = character.Font;
                characterFont.Position = clamped;
                return;
            }
        }
        finally
        {
            Release(characterFont);
            Release(character);
            Release(rangeFont);
        }
    }

    private static void ResetOleObjectBaselinePosition(Document document, Range formulaRange)
    {
        Microsoft.Office.Interop.Word.Font? rangeFont = null;
        Range? character = null;
        Microsoft.Office.Interop.Word.Font? characterFont = null;
        try
        {
            rangeFont = formulaRange.Font;
            rangeFont.Position = 0;
            for (var index = formulaRange.Start; index < formulaRange.End; index++)
            {
                Release(characterFont);
                characterFont = null;
                Release(character);
                character = document.Range(index, index + 1);
                if (!string.Equals(character.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                characterFont = character.Font;
                characterFont.Position = 0;
                return;
            }
        }
        finally
        {
            Release(characterFont);
            Release(character);
            Release(rangeFont);
        }
    }

    private static void FinalizeDirectUnnumberedOmmlParagraph(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The direct-SEQ table did not flatten to one formula paragraph.");
            paragraph = paragraphs[1];
            format = paragraph.Format;

            // The managed 1x3 host already normalized indents, spacing, list state
            // and pagination flags when it was created. ConvertToText preserves
            // those values; the only structural formatting it carries into the
            // standalone paragraph is the table's tab-stop set and left alignment.
            // Clear exactly those two artifacts. Retain the old compact-line guard
            // for historical tables whose mandatory separator paragraph used an
            // Exactly-1pt line box.
            tabStops = format.TabStops;
            if (tabStops.Count > 0) tabStops.ClearAll();
            format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            if (format.LineSpacingRule == WdLineSpacing.wdLineSpaceExactly
                && format.LineSpacing <= 2.01f)
            {
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                format.SpaceBefore = 0f;
                format.SpaceAfter = 0f;
                try { format.DisableLineHeightGrid = -1; } catch { }
            }
        }
        finally
        {
            Release(tabStops);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool NormalizeCompactStandaloneNativeOmmlParagraph(
        Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? format = null;
        try
        {
            if ((bool)formulaRange.get_Information(WdInformation.wdWithInTable))
                return false;
            OMaths? maths = null;
            try
            {
                maths = formulaRange.OMaths;
                if (maths.Count != 1) return false;
            }
            finally { Release(maths); }
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            format = paragraph.Format;
            // A direct-SEQ 1x3 is separated from adjacent Word tables by a mandatory
            // VisualTeX structural paragraph using Exactly 1pt line spacing. When
            // Word dismantles the table it can reuse that paragraph's pPr for the
            // newly restored standalone Display OMath. Preserve ordinary/custom
            // paragraph formatting, but never let this unmistakable compact-table
            // sentinel become the visible formula's line box.
            if (format.LineSpacingRule != WdLineSpacing.wdLineSpaceExactly
                || format.LineSpacing > 2.01f)
                return false;
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            try { format.DisableLineHeightGrid = -1; } catch { }
            return true;
        }
        finally
        {
            Release(format);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void ConfigureEquationParagraph(
        Range formulaRange,
        bool numbered,
        float? nativeOmmlWidthPoints = null)
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
            // Use the same semantic center tab as MathType/VisualTeX OLE. Once
            // Latin Modern runs are represented as native normal-text OMath runs,
            // Word correctly centers the painted equation box and, on save/reopen,
            // may internally materialize an equivalent measured left tab from the
            // final glyph width. Persisting our own pre-save width estimate caused
            // a visible drift after Word normalized the OpenType glyph metrics.
            format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
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

    private static float? MeasureNativeOmmlWidthPoints(Range formulaRange)
    {
        Document? document = null;
        Range? visualRange = null;
        Range? start = null;
        Range? end = null;
        Window? window = null;
        Microsoft.Office.Interop.Word.View? view = null;
        Microsoft.Office.Interop.Word.Zoom? zoom = null;
        try
        {
            document = formulaRange.Document;
            var formulaText = formulaRange.Text ?? string.Empty;
            var visibleStart = formulaRange.Start
                + (formulaText.StartsWith("\t", StringComparison.Ordinal) ? 1 : 0);
            visibleStart = Math.Min(visibleStart, formulaRange.End);
            visualRange = document.Range(visibleStart, formulaRange.End);

            // Prefer Word's range geometry. It is stable across repeated bulk
            // insertions and does not force the hidden Word window to allocate a
            // screen rectangle for every OMath. Repeated Window.GetPoint calls can
            // terminate WINWORD after a few dozen formulas on Office 2021. Any
            // transient sub-point discrepancy is closed by
            // CorrectNumberedOmmlPhysicalCenter immediately after the tab is set.
            try
            {
                start = document.Range(visibleStart, visibleStart);
                end = formulaRange.Duplicate;
                end.Collapse(WdCollapseDirection.wdCollapseEnd);
                var startX = Convert.ToSingle(
                    start.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
                var endX = Convert.ToSingle(
                    end.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
                var rangeWidth = endX - startX;
                if (rangeWidth > 0f
                    && !float.IsNaN(rangeWidth)
                    && !float.IsInfinity(rangeWidth))
                {
                    if (string.Equals(
                            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                            "1",
                            StringComparison.Ordinal))
                    {
                        Console.WriteLine(
                            $"    [layout] native OMML width: range={startX:0.###}-{endX:0.###}, points={rangeWidth:0.###}");
                    }
                    return rangeWidth;
                }
            }
            catch (COMException)
            {
                // Protected/custom stories can reject positional information.
                // Fall through to the screen-box fallback only for that case.
            }

            // Exceptional fallback for stories without usable Range.Information.
            // Keep this out of the normal bulk path because Word's GetPoint API is
            // not stable under many consecutive hidden-window measurements.
            try
            {
                window = document.ActiveWindow;
                window.GetPoint(
                    out _,
                    out _,
                    out var screenWidth,
                    out _,
                    visualRange);
                view = window.View;
                zoom = view.Zoom;
                var zoomPercentage = zoom.Percentage;
                var dpi = 96u;
                try
                {
                    var detectedDpi = GetDpiForWindow(new IntPtr(window.Hwnd));
                    if (detectedDpi > 0) dpi = detectedDpi;
                }
                catch (EntryPointNotFoundException)
                {
                    // Windows versions without per-window DPI APIs use 96 DPI.
                }
                if (screenWidth <= 0 || zoomPercentage <= 0 || dpi == 0)
                    return null;
                var screenWidthPoints = screenWidth
                    * 72f
                    * 100f
                    / dpi
                    / zoomPercentage;
                return screenWidthPoints > 0f
                    && !float.IsNaN(screenWidthPoints)
                    && !float.IsInfinity(screenWidthPoints)
                        ? screenWidthPoints
                        : null;
            }
            catch (COMException)
            {
                return null;
            }
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(zoom);
            Release(view);
            Release(window);
            Release(end);
            Release(start);
            Release(visualRange);
            Release(document);
        }
    }

    private static void CorrectNumberedOmmlPhysicalCenter(Range formulaRange)
    {
        Document? document = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        TabStop? formulaTab = null;
        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        Range? start = null;
        Range? end = null;
        try
        {
            if (!ContainsNativeOmml(formulaRange)
                || IsNumberedEquationTable(formulaRange))
                return;

            document = formulaRange.Document;
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            format = paragraph.Format;

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
                // Use the same standard-page fallback as ConfigureEquationParagraph.
            }
            var expected = CalculateEquationTabStops(
                pageWidth,
                leftMargin,
                rightMargin,
                0f,
                0f);

            var formulaText = formulaRange.Text ?? string.Empty;
            var visibleStart = formulaRange.Start
                + (formulaText.StartsWith("\t", StringComparison.Ordinal) ? 1 : 0);
            visibleStart = Math.Min(visibleStart, formulaRange.End);

            // The width returned by Window.GetPoint is quantized to physical pixels.
            // After save/reopen Word can resolve the same OpenType glyphs a fraction
            // of a point wider or narrower than that first estimate. Close the loop
            // against Word's actual post-layout Range geometry and nudge only the
            // left formula tab; the right number tab and every field/bookmark remain
            // untouched. The correction is linear with the left tab position, so
            // one post-layout adjustment is sufficient. A second adjustment before
            // Word has repaginated can read the same stale box and apply the delta
            // twice. The hard guard avoids reacting to undefined COM geometry in
            // background/protected stories.
            {
                Release(end);
                end = null;
                Release(start);
                start = document.Range(visibleStart, visibleStart);
                end = formulaRange.Duplicate;
                end.Collapse(WdCollapseDirection.wdCollapseEnd);
                var startX = Convert.ToSingle(
                    start.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
                var endX = Convert.ToSingle(
                    end.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
                if (float.IsNaN(startX)
                    || float.IsInfinity(startX)
                    || float.IsNaN(endX)
                    || float.IsInfinity(endX)
                    || startX < -10000f
                    || endX < -10000f
                    || endX <= startX)
                    return;

                var actualCenter = (startX + endX) / 2f;
                var correction = expected.Center - actualCenter;
                if (Math.Abs(correction) <= 0.1f) return;
                if (Math.Abs(correction) > 12f) return;

                Release(formulaTab);
                formulaTab = null;
                Release(tabStops);
                tabStops = format.TabStops;
                for (var index = 1; index <= tabStops.Count; index++)
                {
                    TabStop? candidate = null;
                    try
                    {
                        candidate = tabStops[index];
                        if (candidate.Alignment != WdTabAlignment.wdAlignTabLeft
                            || candidate.Position >= expected.Right - 0.1f)
                            continue;
                        formulaTab = candidate;
                        candidate = null;
                        break;
                    }
                    finally { Release(candidate); }
                }
                if (formulaTab is null) return;

                var correctedPosition = Math.Max(
                    0.1f,
                    Math.Min(expected.Right - 0.1f, formulaTab.Position + correction));
                // Moving the existing custom tab preserves the right-aligned number
                // stop. Clearing and re-adding the formula stop can make Word discard
                // the sibling custom tab in some builds, which leaves the paragraph
                // without a usable numbering layout.
                formulaTab.Position = correctedPosition;
                if (string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                {
                    Console.WriteLine(
                        $"    [layout] corrected OMML center: actual={actualCenter:0.###}, target={expected.Center:0.###}, tab={correctedPosition:0.###}");
                }
            }
        }
        catch
        {
            // Physical centering is a layout refinement. If Word refuses geometry
            // queries in a protected/background story, retain the width-based tab
            // computed by ConfigureEquationParagraph rather than breaking numbering.
        }
        finally
        {
            Release(end);
            Release(start);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(formulaTab);
            Release(tabStops);
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(document);
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
            var formulaText = formulaRange.Text ?? string.Empty;
            if (formulaText.StartsWith("\t", StringComparison.Ordinal)) return;
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

    private static void RemoveTrailingOmmlDisplaySeparators(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? visibleRange = null;
        Range? character = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            visibleRange = FindVisibleEquationNumberRange(document, formulaId);
            var scanEnd = visibleRange is not null
                ? visibleRange.Start
                : Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            // Word can expose one or more U+000B display-math separators after a
            // native OMath. Inside a table these were previously removed by the
            // cell normalizer; in the tab paragraph they would act as a manual
            // line break and place the right-tab number on another visual line.
            // Delete only separators strictly outside OMath.Range, scanning
            // backwards so the live equation and visible-number ranges stay valid.
            for (var position = scanEnd - 1;
                 position >= formulaRange.End;
                 position--)
            {
                Release(character);
                character = document.Range(position, position + 1);
                if (string.Equals(character.Text, "\v", StringComparison.Ordinal))
                    character.Delete();
            }
        }
        finally
        {
            Release(character);
            Release(visibleRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
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
            var formulaText = formulaRange.Text ?? string.Empty;
            if (formulaText.StartsWith("\t", StringComparison.Ordinal)
                && formulaRange.Start < formulaRange.End)
            {
                preceding = document.Range(formulaRange.Start, formulaRange.Start + 1);
                preceding.Delete();
                return;
            }
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
        if (ContainsNativeOmml(formulaRange))
        {
            if (ContainsVisualTeXSequenceInsideOmml(formulaRange))
                return;
            throw new InvalidOperationException(
                "VisualTeX refused to create an external caption for native OMML. Numbered OMML must be materialized atomically as wdOMathDisplay #(SEQ VisualTeXEquation).");
        }

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
        string? plannedPrefix = null,
        bool allowNativeOmmlAcceptanceFixture = false)
    {
        if (ContainsNativeOmml(formulaRange)
            && !allowNativeOmmlAcceptanceFixture)
            throw new InvalidOperationException(
                "VisualTeX refused to create a hidden SEQ caption beside native OMML. The SEQ field belongs inside the Word-native #() number slot.");
        if (allowNativeOmmlAcceptanceFixture
            && !string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            throw new InvalidOperationException(
                "The native-OMML legacy caption bypass is available only to isolated acceptance fixture construction.");

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
            // All newly produced VisualTeX OLE and OMML numbers share the exact
            // product-owned sequence used by Word's native #(SEQ) math host. The
            // localized built-in Equation label remains readable as legacy input,
            // but creating another localized SEQ would split mixed OLE/OMML order.
            object fieldCode = plannedOrdinal.HasValue
                ? $"SEQ {LegacyEquationSequenceName} \\r {plannedOrdinal.Value} \\* ARABIC"
                : $"SEQ {LegacyEquationSequenceName} \\* ARABIC";
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
        AssertCaptionRangeIsOutsideOmml(captionRange);

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
            CompressNativeCaptionParagraphFlow(captionRange);
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
        AssertCaptionRangeIsOutsideOmml(captionRange);

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
            CompressNativeCaptionParagraphFlow(captionRange);
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

    private static void CompressNativeCaptionParagraphFlow(Range captionRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        ParagraphFormat? format = null;
        Microsoft.Office.Interop.Word.Font? markFont = null;
        try
        {
            paragraphs = captionRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.End <= paragraphRange.Start) return;

            // The SEQ target remains black 11pt text inside its 0.1pt clipping
            // Frame so Word's native REF fields inherit normal character styling.
            // The final paragraph mark sits outside that Frame; at ordinary 11pt
            // single spacing it creates a full visible blank line below every
            // numbered VisualTeX OLE. Compress only that body-flow paragraph.
            format = paragraph.Format;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            format.LineSpacing = 1f;
            format.KeepTogether = 0;
            format.KeepWithNext = 0;
            format.PageBreakBefore = 0;
            format.WidowControl = 0;
            try { format.DisableLineHeightGrid = -1; } catch { }

            paragraphMark = paragraphRange.Document.Range(
                paragraphRange.End - 1,
                paragraphRange.End);
            markFont = paragraphMark.Font;
            markFont.Hidden = 0;
            markFont.Size = 1f;
            markFont.Position = 0;
        }
        catch
        {
            // The clipping Frame still protects the SEQ target if a protected or
            // custom paragraph refuses the optional flow compaction.
        }
        finally
        {
            Release(markFont);
            Release(format);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void AssertCaptionRangeIsOutsideOmml(Range captionRange)
    {
        OMaths? maths = null;
        try
        {
            maths = captionRange.OMaths;
            if (maths.Count > 0)
                throw new InvalidOperationException(
                    "VisualTeX refused to style or frame a caption range inside OMML. Native equation numbering is owned by #(SEQ) and must remain free of external caption formatting.");
        }
        finally
        {
            Release(maths);
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
                    if (tableLayout && IsVisualTeXSequenceFieldCode(code.Text)) return true;
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
            var reusesExistingRightTab = !tableLayout
                && HasTabImmediatelyBeforePosition(document, suffixStart);
            var labelStart = reusesExistingRightTab
                ? suffixStart - 1
                : suffixStart;
            // Normal VisualTeX insertion/edit keeps the long-standing stable
            // "()" scaffold. Only MathType/OMML -> VisualTeX conversion uses the
            // isolated REF-first path because migrated MathType paragraphs can
            // make Word consume a pre-seeded parenthesis while the field is born.
            var scaffold = tableLayout
                ? useConversionSafeVisibleNumber
                    ? (Text: string.Empty, FieldOffset: 0)
                    : (Text: "()", FieldOffset: 1)
                : reusesExistingRightTab
                    ? (Text: "()", FieldOffset: 1)
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
                NormalizeReferenceResult(
                    field,
                    formulaFontSizePoints,
                    inheritParagraphFormatting: !tableLayout);
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
                    labelStart,
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
                formulaHeightPoints: 0f,
                formulaFontSizePoints,
                inheritParagraphFormatting: !tableLayout);
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

    private static bool HasTabImmediatelyBeforePosition(
        Document document,
        int position)
    {
        Range? preceding = null;
        try
        {
            if (position <= document.Content.Start) return false;
            preceding = document.Range(position - 1, position);
            return string.Equals(preceding.Text, "\t", StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
        finally { Release(preceding); }
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
        var removedTrueDisplayHost = RemoveNativeDisplayNumberShapeAndAnchor(
            document,
            formulaId);
        if (!removedTrueDisplayHost
            || HasVisibleEquationNumberBookmark(document, formulaId))
            RemoveVisibleEquationNumberLegacy(document, formulaId);
    }

    private static void RemoveVisibleEquationNumberLegacy(Document document, string formulaId)
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
            if (containingMath is not null)
            {
                Range? containingRange = null;
                try
                {
                    containingRange = containingMath.Range;
                    if (WordOmmlConverter.HasVisualTeXNativeEquationNumber(
                            containingRange.WordOpenXML)
                        || HasManagedNativeOmmlHashSequenceHost(
                            document,
                            formulaId))
                    {
                        // The visible result is part of one generated display OMath.
                        // Deleting its result range would leave a damaged field and
                        // can make Word linearize the complete equation. Remove only
                        // the lookup bookmark; replacement/unnumbering rebuilds the
                        // professional OMath atomically.
                        bookmark.Delete();
                        return;
                    }
                }
                finally { Release(containingRange); }
            }
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
        if (TryRemoveNativeHashSequenceCaptionBookmarks(document, formulaId))
            return;
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
        var originalParagraphStart = -1;
        var originalParagraphEnd = -1;
        var originalOwnedWholeParagraph = false;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            range = bookmark.Range;
            Paragraphs? originalParagraphs = null;
            Paragraph? originalParagraph = null;
            Range? originalParagraphRange = null;
            try
            {
                originalParagraphs = range.Paragraphs;
                if (originalParagraphs.Count == 1)
                {
                    originalParagraph = originalParagraphs[1];
                    originalParagraphRange = originalParagraph.Range;
                    originalParagraphStart = originalParagraphRange.Start;
                    originalParagraphEnd = originalParagraphRange.End;
                    originalOwnedWholeParagraph = range.Start == originalParagraphStart
                        && range.End == originalParagraphEnd
                        && (range.Text ?? string.Empty).EndsWith("\r", StringComparison.Ordinal);
                }
            }
            finally
            {
                Release(originalParagraphRange);
                Release(originalParagraph);
                Release(originalParagraphs);
            }
            OMaths? bookmarkMaths = null;
            try
            {
                bookmarkMaths = range.OMaths;
                if (bookmarkMaths.Count > 0)
                {
                    // VTEqCap can be an alias inside a native #(SEQ) result. A
                    // bookmark may transiently survive FormattedText replacement;
                    // deleting its Range would then remove part/all of the new
                    // equation. Mathematical aliases are identity only and are
                    // always removed without deleting content.
                    bookmark.Delete();
                    return;
                }
            }
            finally
            {
                Release(bookmarkMaths);
            }

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
                // Frame.Delete preserves the SEQ text but can shrink VTEqCap from
                // the original full paragraph to only the framed contents. Deleting
                // that mutated bookmark range leaves its paragraph mark behind as a
                // real empty line after VisualTeX -> MathType conversion. Freeze the
                // original whole-paragraph coordinates before touching the Frame and
                // remove that exact paragraph after detaching bookmark ownership.
                bookmark.Delete();
                if (originalOwnedWholeParagraph
                    && originalParagraphStart >= document.Content.Start
                    && originalParagraphEnd <= document.Content.End
                    && originalParagraphEnd > originalParagraphStart)
                {
                    Range? completeCaptionParagraph = null;
                    try
                    {
                        completeCaptionParagraph = document.Range(
                            originalParagraphStart,
                            originalParagraphEnd);
                        completeCaptionParagraph.Delete();
                    }
                    finally { Release(completeCaptionParagraph); }
                }
                else
                {
                    range.Delete();
                }
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
                OMaths? maths = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryFormulaIdFromBookmark(bookmark.Name, prefix, out var formulaId)
                        || activeFormulaIds.Contains(formulaId))
                        continue;
                    if (deleteRange)
                    {
                        range = bookmark.Range;
                        maths = range.OMaths;
                        if (maths.Count > 0)
                        {
                            // Native #(SEQ) aliases live inside the mathematical
                            // field result. Removing their text would corrupt the
                            // OMath/field tree. Drop only the orphan identity; an
                            // explicit unnumber/migration operation rebuilds the
                            // complete OMath atomically when content must disappear.
                            bookmark.Delete();
                        }
                        else
                        {
                            range.Delete();
                        }
                    }
                    else
                    {
                        bookmark.Delete();
                    }
                }
                finally
                {
                    Release(maths);
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
        // Native numbered OMML owns no generated REF label to restyle. Refresh its
        // in-math SEQ field only; touching the VTEq_ bookmark's Font/Position would
        // mutate the mathematical delimiter and can destabilize Word's #() host.
        if (IsNumberedNativeOmmlHashSequenceFormula(document, formulaId))
        {
            if (!TryRefreshNumberedNativeOmmlHashSequence(document, formulaId))
                throw new InvalidOperationException(
                    $"Native #(SEQ) OMML {formulaId} could not refresh its field without rewriting Field.Code.Text.");
            return;
        }

        var tableLayout = true;
        var inheritParagraphFormatting = false;
        Range? layoutProbe = null;
        try
        {
            layoutProbe = FindVisibleEquationNumberRange(document, formulaId);
            if (layoutProbe is not null)
            {
                tableLayout = IsNumberedEquationTable(layoutProbe);
                // Only the VisualTeX OLE/MathType-style tab row lives in the main
                // text story and inherits its paragraph mark. True-display OMML
                // numbers live in a TextFrame story and use explicit Shape text
                // formatting; treating those story-relative coordinates as main-
                // document offsets produces Word 0x800A1200 during OMML editing.
                inheritParagraphFormatting = !tableLayout
                    && layoutProbe.StoryType == WdStoryType.wdMainTextStory;
            }
        }
        finally { Release(layoutProbe); }

        UpdateFieldInBookmark(
            document,
            EquationBookmarkName(formulaId),
            code => IsReferenceToBookmark(code, NativeNumberBookmarkName(formulaId)),
            inheritParagraphFormatting);

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
            // The tab-layout number follows the paragraph baseline exactly, like
            // a native MathType display row. The VisualTeX OLE itself owns the
            // baseline offset; never move the number upward to compensate for the
            // object's height. OMML table numbers keep their established explicit
            // compatibility formatting, also at Position=0.
            AlignEquationNumberVertically(
                range,
                formulaHeightPoints: 0f,
                formulaFontSizePoints: formulaFontSizePoints,
                inheritParagraphFormatting);
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
            TraceRangeStage("prepare-suffix");
            // Never call Range.Fields.Update() here. A numbered document can also
            // contain EMBED/LINK/INCLUDE/DDE or unrelated user fields, and Word may
            // display its external-field confirmation dialog when those are swept
            // into a batch update. Refresh only VisualTeX's own SEQ/REF fields.
            if (UpdateVisualTeXOwnedFieldsInRange(suffixRange) == 0) return false;
            TraceRangeStage("suffix-pass-1");
            UpdateVisualTeXOwnedFieldsInRange(suffixRange);
            TraceRangeStage("suffix-pass-2");

            // References located before the insertion point can target a formula
            // whose number just shifted. Refresh only that prefix once after the
            // suffix SEQ values are stable. Previous formulas' own SEQ fields are
            // unchanged, and Word simply reproduces the same results for them.
            if (captionRange.Start > content.Start)
            {
                prefixRange = document.Range(content.Start, captionRange.Start);
                UpdateVisualTeXOwnedFieldsInRange(prefixRange);
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

    private static bool TryUpdateInsertedEquationSequenceSuffixWithinScope(
        Document document,
        string formulaId,
        out Dictionary<string, string> changedFormulaNumbers,
        out int referencesAlreadyUpdatedFrom)
    {
        changedFormulaNumbers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        referencesAlreadyUpdatedFrom = -1;
        var format = ReadEquationNumberFormat(document);

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
                    out var currentPrefix,
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

            var allOrderedSuffix = suffixEntries
                .OrderBy(entry => entry.Position)
                .ToArray();
            var scopedSuffix = new List<NativeEquationCaptionEntry>();
            foreach (var entry in allOrderedSuffix)
            {
                if (!TryParseNativeEquationNumber(
                        entry.NumberText,
                        format,
                        out var entryPrefix,
                        out _))
                    return false;
                if (format.UsesHeading
                    && !string.Equals(entryPrefix, currentPrefix, StringComparison.Ordinal))
                    break;
                scopedSuffix.Add(entry);
            }
            var orderedSuffix = scopedSuffix.ToArray();
            for (var index = 0; index < orderedSuffix.Length; index++)
            {
                if (!TryParseNativeEquationNumber(
                        orderedSuffix[index].NumberText,
                        format,
                        out var existingPrefix,
                        out var existingOrdinal)
                    || (format.UsesHeading
                        && !string.Equals(existingPrefix, currentPrefix, StringComparison.Ordinal)))
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
                        currentPrefix
                        + expectedOrdinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
            }

            // A SEQ field result must be recalculated by Word; replacing that
            // result text directly can destroy the field/bookmark relationship.
            // Use one native suffix update (with our OLE EMBED fields locked),
            // then repair only VisualTeX REF results directly. This keeps Word's
            // sequence engine authoritative while avoiding the previous second
            // full Fields.Update pass.
            var suffixHasOnlyGeneratedReferences = false;
            var usedSingleBatchUpdate = !format.UsesHeading
                && changedFormulaNumbers.Count > 0
                && suffixRange is not null
                && TryBatchUpdateHealthyVisualTeXSuffixFields(
                    suffixRange,
                    nativeSequenceName,
                    changedFormulaNumbers.Count,
                    out suffixHasOnlyGeneratedReferences);
            if (format.UsesHeading)
            {
                foreach (var entry in orderedSuffix)
                {
                    if (!changedFormulaNumbers.ContainsKey(entry.FormulaId))
                        continue;
                    if (!TryParseNativeEquationNumber(
                            entry.NumberText,
                            format,
                            out _,
                            out var oldOrdinal))
                        return false;
                    var expectedOrdinal = oldOrdinal + 1;
                    if (!TryUpdateDirectTableSequenceNumber(
                            document,
                            entry.FormulaId,
                            nativeSequenceName,
                            expectedOrdinal,
                            currentPrefix,
                            formatOnly: false,
                            knownBookmarks: bookmarks))
                        return false;
                }
            }
            else if (usedSingleBatchUpdate)
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
                            out var refreshedPrefix,
                            out var refreshedOrdinal)
                        || refreshedOrdinal != expectedOrdinal
                        || (format.UsesHeading
                            && !string.Equals(refreshedPrefix, currentPrefix, StringComparison.Ordinal)))
                        return false;
                }
                finally
                {
                    Release(refreshedNumberRange);
                    Release(refreshedNumberBookmark);
                }
            }

            if (orderedSuffix.Length == allOrderedSuffix.Length)
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

    private static bool TryReadTargetedRemovalSnapshot(
        Document document,
        string removedFormulaId,
        out IReadOnlyList<NativeEquationCaptionEntry> captions,
        out HashSet<string> externalReferenceFormulaIds)
    {
        captions = Array.Empty<NativeEquationCaptionEntry>();
        externalReferenceFormulaIds = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase);
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            var current = GetNativeEquationCaptionEntries(
                document,
                GetNativeEquationSequenceName(document));
            if (!current.Any(item => string.Equals(
                    item.FormulaId,
                    removedFormulaId,
                    StringComparison.OrdinalIgnoreCase)))
                return false;

            // Current direct-SEQ number cells are ordinary SEQ fields; a wdFieldRef
            // here is therefore a real body/cross-reference, not the generated
            // visible number. Enumerate only those REF fields once instead of
            // serializing/parsing the whole 100-OMML document.
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = null;
                field = fields[index];
                if (field.Type != WdFieldType.wdFieldRef) continue;
                code = field.Code;
                if (!TryResolveVisualTeXReferenceBookmark(
                        document,
                        code.Text ?? string.Empty,
                        out var bookmarkName)
                    || !TryFormulaIdFromBookmark(
                        bookmarkName,
                        NativeNumberBookmarkPrefix,
                        out var referencedFormulaId))
                    continue;
                externalReferenceFormulaIds.Add(referencedFormulaId);
            }
            if (externalReferenceFormulaIds.Contains(removedFormulaId))
                return false;
            captions = current;
            return true;
        }
        catch
        {
            captions = Array.Empty<NativeEquationCaptionEntry>();
            externalReferenceFormulaIds.Clear();
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static bool TryUpdateRemovedEquationSequenceSuffixWithinScope(
        Document document,
        string removedFormulaId,
        IReadOnlyList<NativeEquationCaptionEntry> preRemovalCaptions,
        out Dictionary<string, string> changedFormulaNumbers)
    {
        changedFormulaNumbers = new Dictionary<string, string>(
            StringComparer.OrdinalIgnoreCase);
        var removedIndex = -1;
        for (var index = 0; index < preRemovalCaptions.Count; index++)
        {
            if (!string.Equals(
                    preRemovalCaptions[index].FormulaId,
                    removedFormulaId,
                    StringComparison.OrdinalIgnoreCase))
                continue;
            removedIndex = index;
            break;
        }
        if (removedIndex < 0) return false;

        var format = ReadEquationNumberFormat(document);
        if (!TryParseNativeEquationNumber(
                preRemovalCaptions[removedIndex].NumberText,
                format,
                out var removedPrefix,
                out var removedOrdinal))
            return false;

        Bookmarks? bookmarks = null;
        Bookmark? verifyBookmark = null;
        Range? verifyRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var nativeSequenceName = GetNativeEquationSequenceName(document);
            for (var index = removedIndex + 1;
                 index < preRemovalCaptions.Count;
                 index++)
            {
                var item = preRemovalCaptions[index];
                if (!TryParseNativeEquationNumber(
                        item.NumberText,
                        format,
                        out var itemPrefix,
                        out var oldOrdinal))
                    return false;

                if (format.UsesHeading
                    && !string.Equals(
                        itemPrefix,
                        removedPrefix,
                        StringComparison.Ordinal))
                    break;

                var expectedOldOrdinal = removedOrdinal + (index - removedIndex);
                if (oldOrdinal != expectedOldOrdinal)
                    return false;
                var newOrdinal = oldOrdinal - 1;
                var expectedText = itemPrefix
                    + newOrdinal.ToString(
                        System.Globalization.CultureInfo.InvariantCulture);

                UpdateNativeEquationCaptionNumber(
                    document,
                    item.FormulaId,
                    nativeSequenceName,
                    newOrdinal,
                    itemPrefix,
                    formatOnly: false,
                    cleanupLegacyFrames: false,
                    knownBookmarks: bookmarks);

                var numberName = NativeNumberBookmarkName(item.FormulaId);
                if (!bookmarks.Exists(numberName)) return false;
                Release(verifyRange); verifyRange = null;
                Release(verifyBookmark); verifyBookmark = null;
                verifyBookmark = bookmarks[numberName];
                verifyRange = verifyBookmark.Range;
                if (!string.Equals(
                        NormalizeNativeEquationNumberText(verifyRange.Text),
                        expectedText,
                        StringComparison.Ordinal))
                    return false;
                changedFormulaNumbers[item.FormulaId] = expectedText;
            }

            WriteNativeEquationTailFormulaId(
                document,
                preRemovalCaptions
                    .Where(item => !string.Equals(
                        item.FormulaId,
                        removedFormulaId,
                        StringComparison.OrdinalIgnoreCase))
                    .LastOrDefault()?.FormulaId);
            return true;
        }
        catch
        {
            changedFormulaNumbers.Clear();
            return false;
        }
        finally
        {
            Release(verifyRange);
            Release(verifyBookmark);
            Release(bookmarks);
        }
    }

    private static bool TryPatchVisibleEquationNumberResult(
        Bookmarks bookmarks,
        string formulaId,
        string expectedNumber)
    {
        Bookmark? visibleBookmark = null;
        Range? visibleRange = null;
        Document? visibleDocument = null;
        OMaths? maths = null;
        Tables? tables = null;
        Table? table = null;
        Fields? directFields = null;
        Field? directField = null;
        Range? directCode = null;
        try
        {
            var visibleName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(visibleName)) return false;
            visibleBookmark = bookmarks[visibleName];
            visibleRange = visibleBookmark.Range;
            visibleDocument = visibleRange.Document;
            maths = visibleRange.OMaths;

            // Current OMML numbering owns its visible label directly in cell
            // (1,3): TAB + ordinary SEQ + paragraph mark. There is no generated
            // REF result to patch. Prove that exact host locally and accept the
            // already-refreshed SEQ result instead of falling through to the
            // document-wide Fields scan (which made five numbers take ~1 s).
            if (maths.Count == 0
                && (bool)visibleRange.get_Information(WdInformation.wdWithInTable))
            {
                tables = visibleRange.Tables;
                if (tables.Count == 1)
                {
                    table = tables[1];
                    if (table.Rows.Count == 1 && table.Columns.Count == 3)
                    {
                        directFields = visibleRange.Fields;
                        if (directFields.Count == 1)
                        {
                            directField = directFields[1];
                            directCode = directField.Code;
                            if (IsVisualTeXSequenceFieldCode(directCode.Text))
                                return string.Equals(
                                    NormalizeNativeEquationNumberText(visibleRange.Text),
                                    expectedNumber,
                                    StringComparison.Ordinal);
                        }
                    }
                }
            }

            if (maths.Count > 0
                || HasManagedNativeOmmlHashSequenceHost(
                    visibleDocument,
                    formulaId))
            {
                // VTEq_ is an alias around the result of the mathematical SEQ.
                // Never replace that result text directly: Word must update the
                // unchanged field itself, or the entire numbered OMath must be
                // rebuilt when its format instruction changes.
                return string.Equals(
                    NormalizeNativeEquationNumberText(visibleRange.Text),
                    expectedNumber,
                    StringComparison.Ordinal);
            }
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(directCode);
            Release(directField);
            Release(directFields);
            Release(table);
            Release(tables);
            Release(maths);
            Release(visibleDocument);
            Release(visibleRange);
            Release(visibleBookmark);
        }

        return TryPatchVisibleEquationNumberResultCore(
            bookmarks,
            formulaId,
            expectedNumber);
    }

    private static bool TryPatchVisibleEquationNumberResultCore(
        Bookmarks bookmarks,
        string formulaId,
        string expectedNumber)
    {
        // VTEqNum_* for current native OMML encloses the live mathematical
        // SEQ result. Never replace Range.Text there: doing so unlinks/destroys
        // the field. Recalculate the unchanged field and let ordinary body REF
        // fields read the same bookmark. Prefix/field-switch changes are handled
        // by an atomic OMath rebuild before this result-refresh path is entered.
        Bookmark? nativeHashBookmark = null;
        Range? nativeHashSequenceRange = null;
        OMaths? nativeHashMaths = null;
        Fields? nativeHashFields = null;
        try
        {
            var nativeHashName = NativeNumberBookmarkName(formulaId);
            if (bookmarks.Exists(nativeHashName))
            {
                nativeHashBookmark = bookmarks[nativeHashName];
                nativeHashSequenceRange = nativeHashBookmark.Range;
                nativeHashMaths = nativeHashSequenceRange.OMaths;
                if (nativeHashMaths.Count > 0)
                {
                    nativeHashFields = nativeHashSequenceRange.Fields;
                    for (var fieldIndex = 1;
                         fieldIndex <= nativeHashFields.Count;
                         fieldIndex++)
                    {
                        Field? nativeHashField = null;
                        Range? nativeHashCode = null;
                        try
                        {
                            nativeHashField = nativeHashFields[fieldIndex];
                            nativeHashCode = nativeHashField.Code;
                            if (!IsVisualTeXSequenceFieldCode(nativeHashCode.Text))
                                continue;
                            nativeHashField.Update();
                            return string.Equals(
                                NormalizeNativeEquationNumberText(
                                    nativeHashSequenceRange.Text),
                                expectedNumber,
                                StringComparison.Ordinal);
                        }
                        finally
                        {
                            Release(nativeHashCode);
                            Release(nativeHashField);
                        }
                    }
                    return false;
                }
            }
        }
        finally
        {
            Release(nativeHashFields);
            Release(nativeHashMaths);
            Release(nativeHashSequenceRange);
            Release(nativeHashBookmark);
        }


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
                    if (range.StoryType == WdStoryType.wdMainTextStory)
                        ApplyParagraphEquationNumberFont(
                            range,
                            fallbackSize: 0f,
                            position: 0);
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

            // Do not batch-update the raw Fields collection. It can contain OLE
            // EMBED or external-link fields unrelated to equation numbering. Count
            // and update only VisualTeX-owned SEQ/REF fields so interactive Word
            // never asks the user whether to refresh fields from other files.
            var updated = UpdateVisualTeXOwnedFieldsInRange(
                suffixRange,
                out referenceFieldCount);
            if (updated == 0) return false;
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
        if (IsNumberedNativeOmmlHashSequenceFormula(document, formulaId))
        {
            if (!TryRefreshNumberedNativeOmmlHashSequence(document, formulaId))
                return false;
            Bookmark? directNumberBookmark = null;
            Range? directNumberRange = null;
            try
            {
                var directNumberName = NativeNumberBookmarkName(formulaId);
                if (!bookmarks.Exists(directNumberName)) return false;
                directNumberBookmark = bookmarks[directNumberName];
                directNumberRange = directNumberBookmark.Range;
                return string.Equals(
                    NormalizeNativeEquationNumberText(directNumberRange.Text),
                    expectedOrdinal.ToString(
                        System.Globalization.CultureInfo.InvariantCulture),
                    StringComparison.Ordinal);
            }
            finally
            {
                Release(directNumberRange);
                Release(directNumberBookmark);
            }
        }

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
                EnsureSequenceFieldCodeCanBeRewritten(field);
                code.Text = $" SEQ {LegacyEquationSequenceName} \\* ARABIC ";
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
        EquationNumberFormat format,
        bool trustedHealthyDirectTables = false)
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
        Bookmarks? sharedBookmarks = null;
        try
        {
            sharedBookmarks = document.Bookmarks;

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
                    cleanupLegacyFrames: false,
                    knownBookmarks: sharedBookmarks,
                    trustedHealthyDirectTable:
                        trustedHealthyDirectTables);
                changedFormulaNumbers[caption.FormulaId] = expectedNumberText;
            }

            WriteNativeEquationTailFormulaId(document, captions.LastOrDefault()?.FormulaId);
            return changedFormulaNumbers;
        }
        finally { Release(sharedBookmarks); }
    }

    private static bool IsFieldInsideOmml(Field field)
    {
        if (field is null) return false;
        Range? code = null;
        OMaths? codeMaths = null;
        try
        {
            // Rewriting safety is determined by the field instruction, not by its
            // rendered result. A perfectly ordinary body REF can target VTEqNum
            // inside an OMath and Word may then expose Field.Result.OMaths.Count>0;
            // that does not make the REF field itself mathematical. Conversely a
            // SEQ/REF whose Code is inside professional OMML must never have
            // Field.Code.Text assigned in place.
            code = field.Code;
            codeMaths = code.OMaths;
            return codeMaths.Count > 0;
        }
        catch
        {
            // An inaccessible field code is never safe for in-place rewriting.
            // Conservatively keep it on structural migration paths.
            return true;
        }
        finally
        {
            Release(codeMaths);
            Release(code);
        }
    }

    private static void EnsureSequenceFieldCodeCanBeRewritten(Field field)
    {
        if (field is null) throw new ArgumentNullException(nameof(field));
        if (IsFieldInsideOmml(field))
            throw new InvalidOperationException(
                "VisualTeX refused to rewrite Field.Code.Text inside an OMML equation. Native #(SEQ) numbering must be rebuilt atomically or updated with F9/Field.Update only.");
    }

    private static void EnsureReferenceFieldCodeCanBeRewritten(Field field)
    {
        if (field is null) throw new ArgumentNullException(nameof(field));
        if (IsFieldInsideOmml(field))
            throw new InvalidOperationException(
                "VisualTeX refused to rewrite a REF field inside OMML. Legacy #(REF) numbering must be migrated by atomically replacing the complete equation.");
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
        Func<string?, bool> predicate,
        bool inheritParagraphFormatting = false)
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
                            knownFontSize: null,
                            inheritParagraphFormatting);
                    }
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            if (inheritParagraphFormatting
                && range.StoryType == WdStoryType.wdMainTextStory)
                ApplyParagraphEquationNumberFont(
                    range,
                    fallbackSize: 0f,
                    position: 0);
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
        float formulaFontSizePoints,
        bool inheritParagraphFormatting = false)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = numberRange.Font;
            var numberFontSize = FormulaFontSize.Normalize(formulaFontSizePoints);
            var position = CalculateEquationNumberFontPosition(
                formulaHeightPoints,
                numberFontSize);
            if (inheritParagraphFormatting)
            {
                ApplyParagraphEquationNumberFont(
                    numberRange,
                    numberFontSize,
                    position);
                return;
            }

            // The native caption target is deliberately white and one point.
            // Word propagates that appearance into REF results unless the
            // visible range is normalized after every field update. Normalize
            // the brackets and field result as one run so locale-specific body
            // fonts cannot put the digits and parentheses on different baselines.
            ApplyEquationNumberFont(
                font,
                numberFontSize,
                position,
                inheritParagraphFormatting: false);
        }
        finally { Release(font); }
    }

    private static void ApplyEquationNumberFont(
        Microsoft.Office.Interop.Word.Font font,
        float size,
        int position,
        bool inheritParagraphFormatting = false)
    {
        if (inheritParagraphFormatting)
        {
            // MathType's display-number style inherits the host paragraph typeface.
            // Reset first to remove direct Cambria Math formatting left by an older
            // build or CHARFORMAT. During conversion the paragraph mark can receive
            // its final semantic size only after the REF is born, so apply a known
            // positive size here without imposing a separate typeface.
            try { font.Reset(); } catch { }
            if (size > 0f)
                font.Size = size;
        }

        font.Hidden = 0;
        font.Color = WdColor.wdColorAutomatic;
        if (!inheritParagraphFormatting)
        {
            font.Size = size;
            try { font.Name = EquationNumberFontName; } catch { }
            try { font.NameAscii = EquationNumberFontName; } catch { }
            try { font.NameFarEast = EquationNumberFontName; } catch { }
            try { font.NameBi = EquationNumberFontName; } catch { }
        }
        font.Position = position;
        try { font.Bold = 0; } catch { }
        try { font.Italic = 0; } catch { }
        try { font.Superscript = 0; } catch { }
        try { font.Subscript = 0; } catch { }
        try { font.Scaling = 100; } catch { }
        try { font.Spacing = 0f; } catch { }
        try { font.Kerning = 0f; } catch { }
    }

    private static void ApplyParagraphEquationNumberFont(
        Range targetRange,
        float fallbackSize,
        int position)
    {
        Document? document = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        Microsoft.Office.Interop.Word.Font? paragraphFont = null;
        Microsoft.Office.Interop.Word.Font? targetFont = null;
        try
        {
            // Paragraph coordinates are story-relative. A text-box REF can have the
            // same numeric Start/End values as the main document, but passing those
            // coordinates through main-story paragraph logic raises Word 0x800A1200.
            // True-display OMML numbers deliberately live in a TextFrame story and
            // therefore use explicit number formatting instead of paragraph-mark
            // inheritance.
            if (targetRange.StoryType != WdStoryType.wdMainTextStory)
            {
                var explicitSize = fallbackSize > 0f
                    ? FormulaFontSize.Normalize(fallbackSize)
                    : ReadUsableFontSize(targetRange) ?? 11f;
                targetFont = targetRange.Font;
                ApplyEquationNumberFont(
                    targetFont,
                    explicitSize,
                    position,
                    inheritParagraphFormatting: false);
                return;
            }

            document = targetRange.Document;
            paragraphs = targetRange.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
                if (paragraphRange.End > paragraphRange.Start)
                {
                    paragraphMark = paragraphRange.Duplicate;
                    paragraphMark.SetRange(
                        paragraphRange.End - 1,
                        paragraphRange.End);
                    paragraphFont = paragraphMark.Font;
                }
            }

            var size = paragraphMark is not null
                ? ReadUsableFontSize(paragraphMark)
                : null;
            var normalizedSize = size
                ?? (fallbackSize > 0f
                    ? FormulaFontSize.Normalize(fallbackSize)
                    : 11f);
            targetFont = targetRange.Font;
            ApplyEquationNumberFont(
                targetFont,
                normalizedSize,
                position,
                inheritParagraphFormatting: true);

            if (paragraphFont is null) return;
            static void CopyName(
                Func<string?> read,
                Action<string> write)
            {
                try
                {
                    var value = read();
                    if (!string.IsNullOrWhiteSpace(value)) write(value!);
                }
                catch { }
            }
            CopyName(
                () => paragraphFont.Name,
                value => targetFont.Name = value);
            CopyName(
                () => paragraphFont.NameAscii,
                value => targetFont.NameAscii = value);
            CopyName(
                () => paragraphFont.NameFarEast,
                value => targetFont.NameFarEast = value);
            CopyName(
                () => paragraphFont.NameBi,
                value => targetFont.NameBi = value);
        }
        finally
        {
            Release(targetFont);
            Release(paragraphFont);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(document);
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
        var frozen = WordEquationReferenceFields.FreezeNavigableReferences(
            document,
            targetBookmarkName);
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
                    if (!matches || IsFieldInsideOmml(field)) continue;
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
                            knownFontSize: null,
                            inheritParagraphFormatting: generatedTabLayout);
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
                catch
                {
                    // A protected/legacy field can reject direct result mutation.
                    // Fall back to Word's normal field update for that one field;
                    // healthy generated references never pay this heavier path.
                    try
                    {
                        generatedTabLayout =
                            IsGeneratedTabLayoutEquationNumberReference(
                                document,
                                field!,
                                formulaId);
                        NormalizeReferenceResult(
                            field!,
                            knownFontSize: null,
                            inheritParagraphFormatting: generatedTabLayout);
                        if (generatedTabLayout)
                            NormalizeGeneratedTabLayoutEquationNumberLabel(
                                document,
                                formulaId);
                        updated++;
                    }
                    catch { }
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

                    if (!TryFormulaIdFromBookmark(
                            visualTeXBookmark,
                            NativeNumberBookmarkPrefix,
                            out var formulaId))
                        continue;
                    if (targetFormulaIds is not null
                        && !targetFormulaIds.Contains(formulaId))
                        continue;

                    if (!IsReferenceToBookmark(code.Text, visualTeXBookmark))
                    {
                        EnsureReferenceFieldCodeCanBeRewritten(field);
                        code.Text = $" REF {visualTeXBookmark} \\h ";
                    }

                    // NormalizeReferenceResult performs the update itself after
                    // applying CHARFORMAT. Restrict this pass to VisualTeX
                    // equation references; touching unrelated REF fields made a
                    // local formula edit scale with the rest of the document.
                    // The generated right-side number in a tab-layout OLE is the
                    // one exception to the historical Cambria compatibility font:
                    // it must continue inheriting the display paragraph style.
                    var generatedTabLayout =
                        IsGeneratedTabLayoutEquationNumberReference(
                            document,
                            field,
                            formulaId);
                    NormalizeReferenceResult(
                        field,
                        knownFontSize: null,
                        inheritParagraphFormatting: generatedTabLayout);
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

        // VisualTeX's double-clickable equation reference is an outer GOTOBUTTON
        // whose live REF is nested inside GOTOBUTTON.Code. Word omits that nested
        // field from document.Fields, so the ordinary pass above cannot refresh it.
        // Update only those nested body REF fields here; the mathematical #(SEQ)
        // targets and their field codes are never touched.
        ISet<string>? targetBookmarkNames = targetFormulaIds is null
            ? null
            : targetFormulaIds
                .Select(NativeNumberBookmarkName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        updated += WordEquationReferenceFields.UpdateNavigableReferences(
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

    private static void NormalizeReferenceResult(
        Field field,
        float? knownFontSize = null,
        bool inheritParagraphFormatting = false)
    {
        if (IsFieldInsideOmml(field))
            throw new InvalidOperationException(
                "VisualTeX refused to normalize or rewrite a REF field inside OMML. Legacy #(REF) numbering must be migrated by atomically replacing the complete equation.");

        Range? code = null;
        Range? result = null;
        Microsoft.Office.Interop.Word.Font? codeFont = null;
        Microsoft.Office.Interop.Word.Font? resultFont = null;
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

            // CHARFORMAT makes Word use the field-code appearance instead of
            // copying the hidden one-point SEQ target appearance into the REF.
            // Re-check the actual field story here instead of trusting the caller:
            // bookmarks can move from the main story into a Shape during numbered
            // OMML migration while an older Range RCW still reports the old host.
            var inheritCodeParagraphFormatting = inheritParagraphFormatting
                && code.StoryType == WdStoryType.wdMainTextStory;
            if (inheritCodeParagraphFormatting)
            {
                ApplyParagraphEquationNumberFont(code, size, position: 0);
            }
            else
            {
                codeFont = code.Font;
                ApplyEquationNumberFont(
                    codeFont,
                    size,
                    position: 0,
                    inheritParagraphFormatting: false);
            }
            field.Update();

            result = field.Result;
            var inheritResultParagraphFormatting = inheritParagraphFormatting
                && result.StoryType == WdStoryType.wdMainTextStory;
            if (inheritResultParagraphFormatting)
            {
                ApplyParagraphEquationNumberFont(result, size, position: 0);
            }
            else
            {
                resultFont = result.Font;
                ApplyEquationNumberFont(
                    resultFont,
                    size,
                    position: 0,
                    inheritParagraphFormatting: false);
            }
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
        if (document is null) throw new ArgumentNullException(nameof(document));

        // Every VisualTeX-owned equation host must participate in one locale-neutral
        // Word sequence. Native OMML #(SEQ), table-free VisualTeX OLE and converted
        // MathType formulas can coexist in the same document; using Word's localized
        // built-in “Equation” caption name for only the OLE side would split F9 into
        // independent sequences on non-English Word installations. Older localized
        // caption fields are migration input: EnsureNativeCaption removes/recreates
        // their VisualTeX-owned VTEqCap/VTEqNum scaffold with this fixed sequence.
        return LegacyEquationSequenceName;
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
        return IsVisualTeXSequenceFieldCode(code)
            || code!.IndexOf(
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
    private static void RewriteSequenceFieldCodeOutsideMath(
        Field field,
        Range code,
        string replacementCode)
    {
        EnsureSequenceFieldCodeCanBeRewritten(field);
        code.Text = replacementCode;
    }

    private static string ExtractNativeHashSequenceResultFromBookmarkXml(
        string bookmarkSegment)
    {
        if (string.IsNullOrEmpty(bookmarkSegment)) return string.Empty;
        var options = RegexOptions.IgnoreCase
            | RegexOptions.CultureInvariant
            | RegexOptions.Singleline;
        var begin = Regex.Match(
            bookmarkSegment,
            @"<w:fldChar\b[^>]*\bw:fldCharType=""begin""[^>]*/?>",
            options);
        var separate = Regex.Match(
            bookmarkSegment,
            @"<w:fldChar\b[^>]*\bw:fldCharType=""separate""[^>]*/?>",
            options);
        var end = Regex.Match(
            bookmarkSegment,
            @"<w:fldChar\b[^>]*\bw:fldCharType=""end""[^>]*/?>",
            options);
        if (!begin.Success || !separate.Success || !end.Success
            || begin.Index >= separate.Index
            || separate.Index >= end.Index)
            return string.Empty;

        static int LastMathRunStart(string xml, int beforeIndex)
        {
            var bare = xml.LastIndexOf(
                "<m:r>",
                beforeIndex,
                StringComparison.OrdinalIgnoreCase);
            var attributed = xml.LastIndexOf(
                "<m:r ",
                beforeIndex,
                StringComparison.OrdinalIgnoreCase);
            return Math.Max(bare, attributed);
        }

        var beginRunStart = LastMathRunStart(bookmarkSegment, begin.Index);
        var separateRunEnd = bookmarkSegment.IndexOf(
            "</m:r>",
            separate.Index,
            StringComparison.OrdinalIgnoreCase);
        var endRunStart = LastMathRunStart(bookmarkSegment, end.Index);
        if (beginRunStart < 0 || separateRunEnd < separate.Index
            || endRunStart <= separateRunEnd)
            return string.Empty;
        separateRunEnd += "</m:r>".Length;

        // VTEqNum wraps the complete displayed number. Static chapter text lives
        // before fldChar(begin), while the live SEQ result lives between
        // fldChar(separate) and fldChar(end). The instruction run between begin and
        // separate must never be interpreted as rendered number text.
        var visibleXml = bookmarkSegment.Substring(0, beginRunStart)
            + bookmarkSegment.Substring(
                separateRunEnd,
                endRunStart - separateRunEnd);
        var text = string.Concat(
            Regex.Matches(
                    visibleXml,
                    @"<(?:w|m):t(?:\s[^>]*)?>(?<text>.*?)</(?:w|m):t>",
                    options)
                .Cast<Match>()
                .Select(match => System.Net.WebUtility.HtmlDecode(
                    match.Groups["text"].Value)));
        return NormalizeNativeEquationNumberText(text);
    }

    private static bool HasLegacyNumberedOmmlShapeArtifact(Document document)
    {
        Range? content = null;
        try
        {
            content = document.Content;
            var xml = content.WordOpenXML ?? string.Empty;
            return Regex.IsMatch(
                    xml,
                    @"\bVTEqShape_[0-9A-F]{32}\b",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                || xml.IndexOf(
                    NativeDisplayNumberShapeAlternativeTextPrefix,
                    StringComparison.Ordinal) >= 0;
        }
        catch
        {
            return false;
        }
        finally { Release(content); }
    }

}
