using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordEquationNumbering;
using static VisualTeX.WordVsto.WordEquationHeadings;
using static VisualTeX.WordVsto.WordEquationNumberReferences;
using static VisualTeX.WordVsto.WordEquationSequenceFields;
using static VisualTeX.WordVsto.WordNumberedDisplayLayout;
using static VisualTeX.WordVsto.WordNumberedTableDisplay;
using static VisualTeX.WordVsto.WordNumberedTrueDisplay;
using static VisualTeX.WordVsto.WordNumberedTypingParagraphs;
using static VisualTeX.WordVsto.WordNumberingFields;
using static VisualTeX.WordVsto.WordNumberingReconcile;
using static VisualTeX.WordVsto.WordNumberingState;

namespace VisualTeX.WordVsto;

/// <summary>
/// The equation number format and the user preferences: the saved format, the document's own format variable,
/// legacy formats and applying a format to every number.
/// </summary>
internal static partial class WordEquationNumberFormat
{
    // Earlier versions stored the chosen style in this document variable. It is
    // still honoured when read and is removed when VisualTeX next numbers the
    // document; nothing writes it any more.
    private const string LegacyEquationNumberFormatVariableName = "VisualTeXEquationNumberFormat";
    private const string EquationNumberTailFormulaVariableName = "VisualTeXEquationNumberTailFormulaId";
    private const string UserPreferenceRegistryPath = @"Software\VisualTeX\Word";
    private const string DefaultNumberedPreferenceName = "DefaultDisplayEquationNumbered";
    private const string DefaultNumberFormatPreferenceName = "DefaultEquationNumberFormat";
    private const string DefaultCreateObjectModePreferenceName = "DefaultCreateFormulaObjectMode";
    private const string DefaultMathTypeNumberPositionPreferenceName = "DefaultMathTypeNumberPosition";

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

    internal sealed class EquationNumberFormatVariableSnapshot
    {
        internal EquationNumberFormatVariableSnapshot(bool exists, string value)
        {
            Exists = exists;
            Value = value;
        }

        internal bool Exists { get; }
        internal string Value { get; }
    }

    // Word documents stay purely native: the chosen numbering style is never
    // written into the document. The style chosen (or worked out) for a
    // document in this Word session is kept here; otherwise it is read from the
    // document's own equation numbers, and a document without numbers uses the
    // user's default.
    private static readonly Dictionary<string, string> SessionEquationNumberFormats =
        new(StringComparer.OrdinalIgnoreCase);

    [ThreadStatic]
    private static bool inferringEquationNumberFormat;

    // The style the document's own equation numbers show (WordVsto:
    // WordEquationNumberFormatAuthority); null when it has none.
    static partial void InferDocumentEquationNumberFormat(
        Document document,
        ref string? formatId);

    private static string SessionEquationNumberFormatKey(Document document)
    {
        try { return document.FullName ?? string.Empty; }
        catch { return string.Empty; }
    }

    private static void RememberEquationNumberFormat(Document document, string formatId)
    {
        var key = SessionEquationNumberFormatKey(document);
        if (key.Length == 0) return;
        lock (SessionEquationNumberFormats)
            SessionEquationNumberFormats[key] = EquationNumberFormat.Resolve(formatId).Id;
    }

    // The style chosen for this document, without looking at its numbers.
    internal static EquationNumberFormatVariableSnapshot CaptureEquationNumberFormatVariable(
        Document document)
    {
        var key = SessionEquationNumberFormatKey(document);
        if (key.Length > 0)
        {
            lock (SessionEquationNumberFormats)
            {
                if (SessionEquationNumberFormats.TryGetValue(key, out var chosen))
                    return new EquationNumberFormatVariableSnapshot(exists: true, chosen);
            }
        }
        return TryReadLegacyEquationNumberFormat(document, out var legacy)
            ? new EquationNumberFormatVariableSnapshot(exists: true, legacy)
            : new EquationNumberFormatVariableSnapshot(exists: false, string.Empty);
    }

    internal static void RestoreEquationNumberFormatVariable(
        Document document,
        EquationNumberFormatVariableSnapshot snapshot)
    {
        if (snapshot.Exists)
        {
            RememberEquationNumberFormat(document, snapshot.Value);
            return;
        }
        var key = SessionEquationNumberFormatKey(document);
        if (key.Length == 0) return;
        lock (SessionEquationNumberFormats)
            SessionEquationNumberFormats.Remove(key);
    }

    private static bool TryReadLegacyEquationNumberFormat(Document document, out string formatId)
    {
        formatId = string.Empty;
        Variables? variables = null;
        Variable? variable = null;
        try
        {
            variables = document.Variables;
            if (variables.Count == 0) return false;
            object index = LegacyEquationNumberFormatVariableName;
            try
            {
                variable = variables.get_Item(ref index);
                formatId = variable.Value ?? string.Empty;
                return true;
            }
            catch (COMException)
            {
                return false;
            }
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(variable);
            Release(variables);
        }
    }

    internal static bool HasLegacyEquationNumberFormat(Document document) =>
        TryReadLegacyEquationNumberFormat(document, out _);

    // An earlier version's variable goes as soon as VisualTeX changes the
    // document; the style it recorded stays chosen for this session.
    internal static void RetireLegacyEquationNumberFormat(Document document)
    {
        if (!TryReadLegacyEquationNumberFormat(document, out var legacy)) return;
        var key = SessionEquationNumberFormatKey(document);
        if (key.Length > 0)
        {
            lock (SessionEquationNumberFormats)
            {
                if (!SessionEquationNumberFormats.ContainsKey(key))
                    SessionEquationNumberFormats[key] = EquationNumberFormat.Resolve(legacy).Id;
            }
        }
        DeleteLegacyEquationNumberFormat(document);
    }

    private static void DeleteLegacyEquationNumberFormat(Document document)
    {
        Variables? variables = null;
        Variable? variable = null;
        try
        {
            variables = document.Variables;
            if (variables.Count == 0) return;
            object index = LegacyEquationNumberFormatVariableName;
            try
            {
                variable = variables.get_Item(ref index);
                variable.Delete();
            }
            catch (COMException ignored)
            {
                // Absent: nothing to remove.
                IgnoredErrors.Ignore(ignored);
            }
        }
        finally
        {
            Release(variable);
            Release(variables);
        }
    }

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
            var xml = WordDocumentXml.Read(document) ?? string.Empty;
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

    internal static int UpdateEquationNumbers(
        Document document,
        IReadOnlyCollection<string>? expectedNumberedFormulaIds = null)
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
        // A complete existing scaffold says nothing about newly inserted formulas.
        // Batch callers already know their required identities; the explicit Ribbon
        // command obtains the managed OMML set once from persisted metadata.
        var requiredIds = expectedNumberedFormulaIds ?? WordOmmlFormulaStore.StoredFormulaIds(document)
            .Where(id =>
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, id);
                return metadata?.Numbered == true && metadata.DisplayMode == "block";
            }).ToArray();
        // Healthy current direct-SEQ documents need only one OpenXML inventory.
        // Running the legacy spacing-repair scan first took a second full snapshot
        // of large documents even when no repair was possible. Keep repair as the
        // compatibility fallback, then retry the strict fast path before entering
        // full structural reconciliation.
        var fastPath = TryRefreshHealthyEquationNumbersInPlace(document, out var updated, requiredIds);
        if (!fastPath)
        {
            RepairNumberedDisplaySpacing(document);
            fastPath = TryRefreshHealthyEquationNumbersInPlace(document, out updated, requiredIds);
        }
        // An XML-proven document with no VisualTeX objects, aliases or references
        // has nothing for this numbering engine. MathType is refreshed separately.
        if (fastPath && updated == 0 && requiredIds.Count == 0) return 0;
        var result = fastPath ? updated : Reconcile(document);
        // The whole-document health proof also fails for unrelated leftovers
        // elsewhere; what this operation needs is a number host for each of
        // its own formulas.
        if (!fastPath && requiredIds.Count > 0)
        {
            if (TryRefreshHealthyEquationNumbersInPlace(document, out var healthyCount, requiredIds))
                result = healthyCount;
            else if (!RequiredFormulasHaveNumberHosts(document, requiredIds))
                throw new InvalidDataException("Numbering did not create a complete host for every required formula.");
        }
        // References retain their original names across format conversions.
        // The numeric-identity inventory alone cannot enumerate inherited ZEqnNum
        // references. Finalize both direct and nested references by their actual
        // target bookmarks, after either numbering implementation completes.
        // The healthy fast path already updates changed target references and
        // repairs every navigable REF discovered by its OpenXML inventory. A
        // second document-wide Fields scan here was pure duplicate work in large
        // sparse documents. Reconciliation does not share that inventory, so it
        // retains the compatibility-wide refresh.
        if (!fastPath)
            WordEquationReferenceFields.UpdateReferences(document);
        if (watch is not null)
        {
            watch.Stop();
            Console.WriteLine(
                $"    [perf] UpdateEquationNumbers.{(fastPath ? "fast" : "reconcile")}: {result} formulas in {watch.ElapsedMilliseconds}ms");
        }
        return result;
    }

    private static bool RequiredFormulasHaveNumberHosts(Document document, IEnumerable<string> formulaIds)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var id in formulaIds)
            {
                if (!bookmarks.Exists(EquationBookmarkName(id))
                    || !bookmarks.Exists(NativeCaptionBookmarkName(id))
                    || !bookmarks.Exists(NativeNumberBookmarkName(id)))
                    return false;
            }
            WordDoubleClickHook.TraceMessage("equation-number-health-local-proof");
            return true;
        }
        catch (COMException) { return false; }
        finally { Release(bookmarks); }
    }

    private static bool TryRefreshHealthyEquationNumbersInPlace(
        Document document,
        out int updated,
        IReadOnlyCollection<string>? requiredFormulaIds = null,
        bool deferExternalReferences = false)
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
        if (requiredFormulaIds is not null)
        {
            var materialized = new HashSet<string>(openXmlCaptions.Select(entry => entry.FormulaId),
                StringComparer.OrdinalIgnoreCase);
            if (requiredFormulaIds.Any(id => !materialized.Contains(id))) return false;
        }

        if (openXmlCaptions.Count == 0) return true;
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
        if (changedFormulaNumbers.Count > 0 && !deferExternalReferences)
        {
            UpdateHealthyNativeCrossReferencesAfterRenumbering(
                document,
                changedFormulaNumbers,
                knownReferenceCounts: referenceCounts);
            TraceStage("reference-refresh");
        }

        // Recreated targets can retain the same number while their body REF still
        // contains Word's missing-bookmark result. REF health is independent of
        // whether a sequence value changed. Reuse the existing XML inventory to
        // avoid opening the Fields collection in documents with no body references.
        var referencedBookmarks = new HashSet<string>(referenceCounts
            .Where(entry => entry.Value > 1)
            .Select(entry => NativeNumberBookmarkName(entry.Key)),
            StringComparer.OrdinalIgnoreCase);
        if (referencedBookmarks.Count > 0 && !deferExternalReferences)
            WordEquationReferenceFields.UpdateNavigableReferences(document, referencedBookmarks);

        updated = captions.Count;
        return true;
    }

    internal static bool TryFinalizeHealthyConversionNumbering(
        Document document,
        out int updated,
        IReadOnlyCollection<string>? requiredFormulaIds = null)
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
                return requiredFormulaIds is null || requiredFormulaIds.Count == 0;

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
                    nativeTargetsAlreadyPlanned: true,
                    deferExternalReferences: true))
                return false;
            TraceStage("field-batch");

            if (!TryRefreshHealthyEquationNumbersInPlace(document, out updated, requiredFormulaIds, deferExternalReferences: true))
                return false;
            TraceStage("healthy-refresh");
            return true;
        }
        catch (Exception error)
        {
            TraceNumberingPerformance($"[perf] conversion-numbering-finalize.failed error={error}");
            throw;
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
        WriteEquationNumberFormat(document, EquationNumberFormat.Resolve(formatId).Id);
        var fastPath = TryApplyEquationNumberFormatInPlace(document, out var updated);
        var result = fastPath ? updated : Reconcile(document);
        WordEquationReferenceFields.UpdateReferences(document);
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

    internal static bool ContainsVisualTeXSequenceInsideOmml(Range formulaRange)
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
        IReadOnlyDictionary<string, int>? knownReferenceCounts = null,
        bool deferExternalReferences = false)
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
            if (!generatedReferencesPatched && deferExternalReferences)
                throw new InvalidDataException("The conversion could not update its own visible equation numbers.");
            if ((!generatedReferencesPatched || !hasOnlyGeneratedReferences) && !deferExternalReferences)
                UpdateNativeCrossReferences(document);
            TraceBatch("extra-references");

            WriteNativeEquationTailFormulaId(document, captions.LastOrDefault()?.FormulaId);
            updated = captions.Count;
            return true;
        }
        catch (Exception error)
        {
            TraceNumberingPerformance($"[perf] equation-number-field-batch.failed mutationStarted={mutationStarted} error={error}");
            if (!mutationStarted)
            {
                updated = 0;
                return false;
            }
            // The enclosing Word transaction owns recovery. A second full
            // structural reconstruction would hide the original partial failure.
            throw;
        }
        finally { Release(bookmarks); }
    }

    internal static void SetEquationNumberFormatPreference(
        Document document,
        string formatId)
    {
        var format = EquationNumberFormat.Resolve(formatId);

        // Acceptance and migration fixtures use this helper to configure the
        // document under test. They must never mutate the real user's default
        // numbering preference in HKCU; otherwise a heading-numbered fixture can
        // leak into the next ordinary document opened by the user.
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
        {
            SetDefaultEquationNumberFormatPreference(format.Id);
        }

        WriteEquationNumberFormat(document, format.Id);
    }

    internal static void RestoreEquationNumberFormatForConversion(
        Document document,
        string formatId)
    {
        using var metric = WordOperationMetrics.Measure("Numbering.RestoreFormatForConversion");
        var resolved = EquationNumberFormat.Resolve(formatId).Id;
        // Only a stored choice can make the write unnecessary. Without one the
        // end state is the remembered `resolved` format whatever the existing
        // numbers would suggest, so do not infer it from them: that scan reads
        // every native and MathType number of a just-numbered document.
        var stored = CaptureEquationNumberFormatVariable(document);
        if (stored.Exists
            && string.Equals(
                EquationNumberFormat.Resolve(stored.Value).Id,
                resolved,
                StringComparison.Ordinal))
        {
            RememberEquationNumberFormat(document, resolved);
            return;
        }
        WriteEquationNumberFormat(document, resolved);
    }

    internal static EquationNumberFormat ReadEquationNumberFormat(Document document)
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
        var stored = CaptureEquationNumberFormatVariable(document);
        var found = stored.Exists ? stored.Value : null;
        if (found is null && !inferringEquationNumberFormat)
        {
            inferringEquationNumberFormat = true;
            try { InferDocumentEquationNumberFormat(document, ref found); }
            catch { found = null; }
            finally { inferringEquationNumberFormat = false; }
        }
        if (found is null) return false;
        formatId = EquationNumberFormat.Resolve(found).Id;
        RememberEquationNumberFormat(document, formatId);
        return true;
    }

    internal static void EnsureDocumentEquationNumberFormatPreference(Document document)
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
            catch (Exception ignored)
            {
                // Preferences are optional. A locked or unavailable registry
                // must never block formula insertion or numbering.
                IgnoredErrors.Ignore(ignored);
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
            catch (Exception ignored)
            {
                // Keep document operations functional even if persistence is
                // unavailable on a restricted machine.
                IgnoredErrors.Ignore(ignored);
            }
        }
    }

    // Called only while VisualTeX numbers the document, so removing an old
    // version's variable here never marks an untouched document as changed.
    private static void WriteEquationNumberFormat(Document document, string formatId)
    {
        RememberEquationNumberFormat(document, formatId);
        DeleteLegacyEquationNumberFormat(document);
    }

    internal static bool TryReadNativeEquationTailFormulaId(
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

    internal static void WriteNativeEquationTailFormulaId(
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
}
