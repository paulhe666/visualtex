using System.Linq;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static class WordOmmlNativeSource
{
    // Keep the actual OMath returned by insertion until the batch commits.
    // A collapsed bookmark has insertion gravity and can move to its neighbour;
    // the retained physical object is independent evidence of fresh ownership.
    internal sealed class LiveInsertionOwners : IDisposable
    {
        private readonly Dictionary<string, OMath> owners = new(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> verifiedOwners = new(StringComparer.OrdinalIgnoreCase);

        internal void Capture(string formulaId, Range inserted)
        {
            OMaths? maths = null;
            OMath? math = null;
            Range? range = null;
            try
            {
                maths = inserted.OMaths;
                if (maths.Count != 1 || owners.ContainsKey(formulaId))
                    throw new InvalidDataException("Fresh OMML insertion has no unique physical owner.");
                math = maths[1];
                range = math.Range;
                if (range.StoryType != inserted.StoryType || range.Start != inserted.Start || range.End != inserted.End)
                    throw new InvalidDataException("Fresh OMML insertion range differs from its physical object.");
                owners.Add(formulaId, math);
                verifiedOwners.Remove(formulaId);
                math = null;
            }
            finally { Release(range); Release(math); Release(maths); }
        }

        internal Range? TryReadOnlyOwner(string formulaId)
        {
            if (owners.Count != 1 || !owners.TryGetValue(formulaId, out var math)) return null;
            return math.Range.Duplicate;
        }

        internal void MarkVerifiedOwners(IEnumerable<string> formulaIds)
        {
            foreach (var formulaId in formulaIds.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (!owners.ContainsKey(formulaId))
                    throw new InvalidDataException($"Fresh OMML owner {formulaId} was not retained for verified reuse.");
                verifiedOwners.Add(formulaId);
            }
        }

        internal Range? TryReadVerifiedOwner(string formulaId)
        {
            if (!verifiedOwners.Contains(formulaId)
                || !owners.TryGetValue(formulaId, out var math))
                return null;

            Range? range = null;
            OMaths? maths = null;
            OMath? resolved = null;
            Range? resolvedRange = null;
            try
            {
                range = math.Range.Duplicate;
                if (range.Start >= range.End) return null;
                maths = range.OMaths;
                if (maths.Count != 1) return null;
                resolved = maths[1];
                resolvedRange = resolved.Range;
                if (resolvedRange.StoryType != range.StoryType
                    || resolvedRange.Start != range.Start
                    || resolvedRange.End != range.End)
                    return null;
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
                Release(resolvedRange);
                Release(resolved);
                Release(maths);
                Release(range);
            }
        }

        internal IReadOnlyDictionary<string, int> CaptureIndices(OMaths maths)
        {
            var actual = new Dictionary<(WdStoryType Story, int Start, int End), int>();
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    actual.Add((range.StoryType, range.Start, range.End), index - 1);
                }
                finally { Release(range); Release(math); }
            }
            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var owner in owners)
            {
                Range? range = null;
                try
                {
                    range = owner.Value.Range;
                    if (!actual.TryGetValue((range.StoryType, range.Start, range.End), out var index))
                        throw new InvalidDataException($"Fresh OMML object {owner.Key} did not survive the batch.");
                    result.Add(owner.Key, index);
                }
                finally { Release(range); }
            }
            return result;
        }

        public void Dispose()
        {
            foreach (var math in owners.Values) Release(math);
            owners.Clear();
            verifiedOwners.Clear();
        }
    }

    // Redraw replaces source paragraphs from the end of the story toward the
    // beginning. Word can rebind both collapsed bookmarks and retained OMath RCWs
    // while those earlier paragraphs are deleted/recreated. A temporary bookmark
    // spanning the complete fresh OMath is materially stronger: it owns content,
    // not an insertion boundary, and therefore tracks that equation as surrounding
    // story coordinates shift. These bookmarks exist only inside the redraw edit
    // transaction and are deleted before success or rollback verification.
    internal sealed class StableRedrawInsertionOwners : IDisposable
    {
        private readonly Dictionary<string, Bookmark> owners = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> materializedSignatures = new(StringComparer.OrdinalIgnoreCase);

        internal IReadOnlyDictionary<string, string> MaterializedSignatures => materializedSignatures;

        internal void Capture(string formulaId, Range inserted, string preparedOmml,
            bool verifyDuringCapture = true)
        {
            if (!Guid.TryParse(formulaId, out var parsed))
                throw new InvalidDataException("Fresh OMML redraw formulaId is invalid.");
            if (owners.ContainsKey(formulaId))
                throw new InvalidDataException("Fresh OMML redraw insertion identity is duplicated.");

            OMaths? maths = null;
            OMath? math = null;
            Range? mathRange = null;
            Document? document = null;
            Bookmarks? bookmarks = null;
            Bookmark? owner = null;
            Range? ownerRange = null;
            try
            {
                maths = inserted.OMaths;
                if (maths.Count != 1)
                    throw new InvalidDataException("Fresh OMML redraw insertion has no unique physical equation.");
                math = maths[1];
                mathRange = math.Range;
                if (mathRange.StoryType != inserted.StoryType
                    || mathRange.Start != inserted.Start
                    || mathRange.End != inserted.End)
                    throw new InvalidDataException("Fresh OMML redraw range differs from its physical equation.");

                var materializedOmml = preparedOmml;
                if (verifyDuringCapture)
                {
                    materializedOmml = WordOmmlConverter.ExtractSingleOMath(mathRange.WordOpenXML);
                    WordOmmlConverter.ComputeVerifiedMaterializedOmmlFingerprint(preparedOmml, materializedOmml);
                }
                // Deferred callers supplied a Word-materialized, prevalidated
                // source and must compare every final equation after Undo ends.
                // Serializing here would split that custom Undo transaction.
                materializedSignatures.Add(
                    formulaId,
                    WordOmmlConverter.ComputeImportedOmmlContentSignature(materializedOmml));

                document = inserted.Document;
                bookmarks = document.Bookmarks;
                var bookmarkName = "VTRD_" + parsed.ToString("N");
                if (bookmarks.Exists(bookmarkName))
                    throw new InvalidDataException("Fresh OMML redraw owner bookmark already exists.");
                owner = bookmarks.Add(bookmarkName, mathRange);
                ownerRange = owner.Range;
                if (ownerRange.Start == ownerRange.End
                    || ownerRange.StoryType != mathRange.StoryType
                    || ownerRange.Start != mathRange.Start
                    || ownerRange.End != mathRange.End)
                    throw new InvalidDataException("Word did not retain the fresh OMML redraw owner range.");
                owners.Add(formulaId, owner);
                owner = null;
            }
            finally
            {
                Release(ownerRange);
                Release(owner);
                Release(bookmarks);
                Release(document);
                Release(mathRange);
                Release(math);
                Release(maths);
            }
        }

        internal IReadOnlyDictionary<string, int> CaptureIndices(OMaths maths)
        {
            var actual = new Dictionary<(WdStoryType Story, int Start, int End), int>();
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    var key = (range.StoryType, range.Start, range.End);
                    if (actual.ContainsKey(key))
                        throw new InvalidDataException("Two physical OMML equations share one redraw owner range.");
                    actual.Add(key, index - 1);
                }
                finally { Release(range); Release(math); }
            }

            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var owner in owners)
            {
                Range? range = null;
                OMaths? ownerMaths = null;
                OMath? ownerMath = null;
                Range? equationRange = null;
                try
                {
                    range = owner.Value.Range;
                    if (range.Start == range.End)
                        throw new InvalidDataException($"Fresh OMML redraw owner {owner.Key} collapsed during the batch.");
                    ownerMaths = range.OMaths;
                    if (ownerMaths.Count != 1)
                        throw new InvalidDataException($"Fresh OMML redraw owner {owner.Key} no longer contains exactly one equation.");
                    ownerMath = ownerMaths[1];
                    equationRange = ownerMath.Range;
                    if (!actual.TryGetValue(
                            (equationRange.StoryType, equationRange.Start, equationRange.End),
                            out var index))
                        throw new InvalidDataException($"Fresh OMML redraw object {owner.Key} did not survive the batch.");
                    result.Add(owner.Key, index);
                }
                finally
                {
                    Release(equationRange);
                    Release(ownerMath);
                    Release(ownerMaths);
                    Release(range);
                }
            }
            return result;
        }

        internal void Clear()
        {
            foreach (var owner in owners.Values)
            {
                try { owner.Delete(); } catch { }
                Release(owner);
            }
            owners.Clear();
            materializedSignatures.Clear();
        }

        public void Dispose() => Clear();
    }

    internal static FormulaMetadata CreateForNative(
        Document document,
        Range equationRange)
    {
        var formulaId = Guid.NewGuid().ToString("D");
        var wordOpenXml = ReadCompleteEquationWordOpenXml(
            document,
            equationRange,
            formulaId);
        var displayMode = ReadDisplayMode(equationRange);
        var numbered = string.Equals(
                displayMode,
                "block",
                StringComparison.Ordinal)
            && WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                wordOpenXml);
        var semanticWordOpenXml = numbered
            ? WordOmmlConverter.StripManagedVisualTeXNativeEquationNumber(
                wordOpenXml)
            : wordOpenXml;
        var mathMl = WordOmmlConverter.TransformOmmlToMathMl(
            semanticWordOpenXml,
            display: string.Equals(displayMode, "block", StringComparison.Ordinal));
        var latex = SanitizeFormulaBoundaryArtifacts(
            MathMlToLatexConverter.Convert(mathMl));
        if (string.IsNullOrWhiteSpace(latex))
            throw new InvalidDataException(
                "The Word-native OMML equation could not be converted back to editable LaTeX.");

        var fontSize = ReadFontSize(equationRange);
        var now = DateTimeOffset.UtcNow.ToString("O");
        var metadata = new FormulaMetadata
        {
            FormulaId = formulaId,
            Title = "Word Formula",
            Latex = latex,
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
            },
            CodeFormat = "raw",
            DisplayMode = displayMode,
            Numbered = numbered,
            FontSizePt = fontSize,
            RenderFontSizePt = fontSize,
            NativeOmmlFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                semanticWordOpenXml),
            CreatedWithVersion = "1.2.7",
            UpdatedWithVersion = "1.2.7",
            CreatedAt = now,
            UpdatedAt = now,
        };
        metadata.Validate();
        return metadata;
    }

    internal static FormulaMetadata RefreshForVisualTeX(
        Document document,
        Bookmark bookmark,
        FormulaMetadata stored)
    {
        Range? equationRange = null;
        try
        {
            equationRange = WordOmmlFormulaStore.GetEquationRangeForCurrentRead(
                document, bookmark, stored);
            var wordOpenXml = ReadCompleteEquationWordOpenXml(
                document,
                equationRange,
                stored.FormulaId);
            return RefreshForVisualTeXFromCapturedWordOpenXml(stored, wordOpenXml);
        }
        finally
        {
            Release(equationRange);
        }
    }

    internal static FormulaMetadata RefreshForVisualTeXFromCapturedWordOpenXml(
        FormulaMetadata stored,
        string wordOpenXml)
    {
        if (stored is null) throw new ArgumentNullException(nameof(stored));
        if (string.IsNullOrWhiteSpace(wordOpenXml))
            throw new ArgumentException("Captured Word OMML is required.", nameof(wordOpenXml));

        var sanitizedStored = Clone(stored);
        SanitizeMetadataBoundaryArtifacts(sanitizedStored);
        var fingerprint = WordOmmlConverter.ComputeOmmlFingerprint(wordOpenXml);
        if (string.Equals(
                stored.NativeOmmlFingerprint,
                fingerprint,
                StringComparison.OrdinalIgnoreCase))
            return sanitizedStored;

        var mathMl = WordOmmlConverter.TransformOmmlToMathMl(
            wordOpenXml,
            display: string.Equals(
                stored.DisplayMode,
                "block",
                StringComparison.Ordinal));
        var latex = SanitizeFormulaBoundaryArtifacts(
            MathMlToLatexConverter.Convert(mathMl));
        if (string.IsNullOrWhiteSpace(latex))
            throw new InvalidDataException(
                "The Word-native OMML equation could not be converted back to editable LaTeX.");

        var refreshed = Clone(stored);
        var lineId = refreshed.Lines.FirstOrDefault()?.Id;
        if (string.IsNullOrWhiteSpace(lineId)) lineId = Guid.NewGuid().ToString();
        refreshed.Latex = latex;
        refreshed.Lines = new List<FormulaLine>
        {
            new() { Id = lineId!, Latex = latex },
        };
        refreshed.CodeFormat = "raw";
        refreshed.NativeOmmlFingerprint = fingerprint;
        refreshed.Validate();

        // Persist only when the owning edit transaction commits. Reading or
        // cancelling an editor must not mutate the document's identity data.
        return refreshed;
    }

    internal static void StampFingerprint(FormulaMetadata metadata, Range equationRange)
    {
        Document? document = null;
        try
        {
            document = equationRange.Document;
            metadata.NativeOmmlFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                ReadCompleteEquationWordOpenXml(
                    document,
                    equationRange,
                    metadata.FormulaId));
        }
        finally { Release(document); }
    }

    internal static void StampFingerprintFromResolvedRange(
        FormulaMetadata metadata,
        Range equationRange)
    {
        // A live OMath Range can still export Word's transient empty scratch body.
        // Use the same verified, read-only export as editor and conversion capture.
        StampFingerprint(metadata, equationRange);
    }

    internal static int RefreshFingerprintsFromDocumentOpenXml(
        Document document,
        IReadOnlyCollection<string> formulaIds,
        IReadOnlyDictionary<string, string>? freshSourceOmml = null,
        LiveInsertionOwners? liveInsertions = null,
        IDictionary<string, FormulaMetadata>? verifiedMetadata = null)
    {
        if (formulaIds is null || formulaIds.Count == 0) return 0;
        if (formulaIds.Count == 1)
        {
            // A single conversion already owns one physical insertion. Verify
            // that exact object and its prepared contents, rather than exporting
            // and indexing every unrelated equation/OLE in the document.
            var id = formulaIds.Single();
            var item = WordOmmlFormulaStore.TryRead(document, id)
                ?? throw new InvalidDataException($"Converted OMML metadata {id} is missing.");
            Range? local = null;
            Bookmark? anchor = null;
            try
            {
                if (freshSourceOmml is not null)
                {
                    if (freshSourceOmml.Count != 1 || !freshSourceOmml.TryGetValue(id, out var prepared)
                        || WordOmmlConverter.ComputeOmmlFingerprint(prepared) != item.NativeOmmlFingerprint)
                        throw new InvalidDataException("Initial OMML finalization does not match its prepared source identity.");
                    local = liveInsertions?.TryReadOnlyOwner(id);
                    // Legacy callers without retained physical ownership keep the
                    // existing document-level recovery path below.
                    if (local is not null)
                    {
                        var actual = ReadCompleteEquationWordOpenXml(document, local, id);
                        item.NativeOmmlFingerprint = WordOmmlConverter.ComputeVerifiedMaterializedOmmlFingerprint(prepared, actual);
                        anchor = WordOmmlFormulaStore.FindByFormulaId(document, id);
                        if (!WordOmmlFormulaStore.IsCanonicalAnchor(anchor, local))
                        {
                            Release(anchor); anchor = null;
                            anchor = WordOmmlFormulaStore.Wrap(document, local, item);
                        }
                        if (!WordOmmlFormulaStore.IsCanonicalAnchor(anchor, local))
                            throw new InvalidDataException("Word did not retain the single converted OMML anchor.");
                        if (!WordOmmlFormulaStore.TrySaveKnownCachedPart(document, item))
                            WordOmmlFormulaStore.Save(document, item);
                        if (verifiedMetadata is not null)
                            verifiedMetadata[id] = item;
                        liveInsertions?.MarkVerifiedOwners(new[] { id });
                        return 1;
                    }
                }
                else
                {
                    if (liveInsertions is not null)
                        throw new InvalidDataException("Live insertion ownership requires the prepared source.");
                    // The normal strict resolver checks the durable anchor, its
                    // numbered row and captured fingerprint. Do not replace a
                    // content mismatch with a new fingerprint here.
                    local = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(document, id, item);
                    var actual = ReadCompleteEquationWordOpenXml(document, local, id);
                    if (WordOmmlConverter.ComputeOmmlFingerprint(actual) != item.NativeOmmlFingerprint)
                        throw new InvalidDataException("The final single converted OMML content changed.");
                    if (verifiedMetadata is not null)
                        verifiedMetadata[id] = item;
                    return 1;
                }
            }
            finally { Release(anchor); Release(local); }
        }
        Range? content = null;
        OMaths? maths = null;
        try
        {
            content = document.Content;
            var package = XDocument.Parse(WordDocumentXml.Read(document));
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var body = package.Descendants(w + "body").Single();
            maths = document.OMaths;
            var xmlEquations = body.Descendants(m + "oMath").ToArray();
            if (maths.Count != xmlEquations.Length)
            {
                WordDoubleClickHook.TraceMessage($"omml-document-export-mismatch document={document.FullName} range={content.Start}:{content.End} com={maths.Count} xml={xmlEquations.Length}");
                if (Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_RECOVERY_XML") == "1")
                {
                    var trace = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH")
                        ?? throw new InvalidOperationException("Diagnostic evidence needs a configured trace path.");
                    File.WriteAllText(trace + ".export-mismatch.xml", package.ToString());
                    Range? probe = null;
                    try
                    {
                        probe = document.Range(content.Start, content.End);
                        File.WriteAllText(trace + ".export-explicit-range.xml", probe.WordOpenXML);
                    }
                    finally { Release(probe); }
                }
                throw new InvalidDataException("COM and Word XML disagree on the complete equation inventory.");
            }
            var metadata = formulaIds.Distinct(StringComparer.OrdinalIgnoreCase).ToDictionary(
                id => id, id => WordOmmlFormulaStore.TryRead(document, id)
                    ?? throw new InvalidDataException($"Converted OMML metadata {id} is missing."),
                StringComparer.OrdinalIgnoreCase);
            if (freshSourceOmml is not null && (freshSourceOmml.Count != metadata.Count
                || metadata.Any(pair => !freshSourceOmml.TryGetValue(pair.Key, out var source)
                    || WordOmmlConverter.ComputeOmmlFingerprint(source) != pair.Value.NativeOmmlFingerprint)))
                throw new InvalidDataException("Initial OMML finalization does not match its prepared source identities.");
            Func<string, string> fingerprint = freshSourceOmml is null
                ? WordOmmlConverter.ComputeOmmlFingerprint
                : WordOmmlConverter.ComputeImportedOmmlContentSignature;
            var expected = metadata.ToDictionary(pair => pair.Key,
                pair => freshSourceOmml is null ? pair.Value.NativeOmmlFingerprint ?? string.Empty
                    : fingerprint(freshSourceOmml[pair.Key]), StringComparer.OrdinalIgnoreCase);
            IReadOnlyDictionary<string, (int Index, bool Canonical)> identities;
            try
            {
                if (liveInsertions is not null && freshSourceOmml is null)
                    throw new InvalidDataException("Live insertion ownership requires the prepared batch sources.");
                var exactAnchorOwners = CaptureExactAnchorOwners(document, maths, expected.Keys.ToArray());
                identities = IndexConversionEquationIdentities(body, expected, fingerprint, exactAnchorOwners,
                    liveInsertions?.CaptureIndices(maths));
            }
            catch
            {
                if (Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_RECOVERY_XML") == "1")
                {
                    var trace = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH")
                        ?? throw new InvalidOperationException("Diagnostic evidence needs a configured trace path.");
                    File.WriteAllText(trace + ".identity-body.xml", body.ToString());
                    if (freshSourceOmml is not null)
                        foreach (var source in freshSourceOmml)
                            File.WriteAllText(trace + ".identity-" + source.Key + ".xml", source.Value);
                }
                throw;
            }
            // The XML pass proves the entire mapping before any anchor changes.
            // Healthy rows need no per-equation WordOpenXML serialization. Only a
            // proven displaced anchor pays for a complete COM/ XML cross-check.
            foreach (var identity in identities)
            {
                var item = metadata[identity.Key];
                if (freshSourceOmml is not null)
                    item.NativeOmmlFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                        xmlEquations[identity.Value.Index].ToString(SaveOptions.DisableFormatting));
                if (identity.Value.Canonical)
                {
                    if (freshSourceOmml is not null) WordOmmlFormulaStore.Save(document, item);
                    continue;
                }
                OMath? math = null;
                Range? range = null;
                Bookmark? rebound = null;
                try
                {
                    math = maths[identity.Value.Index + 1];
                    range = math.Range;
                    var actual = WordOmmlConverter.ComputeOmmlFingerprint(range.WordOpenXML);
                    if (!string.Equals(actual, item.NativeOmmlFingerprint, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidDataException($"COM equation content does not own OMML identity {identity.Key}.");
                    rebound = WordOmmlFormulaStore.Wrap(document, range, item);
                    if (!WordOmmlFormulaStore.IsCanonicalAnchor(rebound, range))
                        throw new InvalidDataException($"Word did not retain the repaired OMML anchor {identity.Key}.");
                    if (freshSourceOmml is not null) WordOmmlFormulaStore.Save(document, item);
                }
                finally { Release(rebound); Release(range); Release(math); }
            }
            if (verifiedMetadata is not null)
            {
                foreach (var pair in metadata)
                    verifiedMetadata[pair.Key] = pair.Value;
            }
            liveInsertions?.MarkVerifiedOwners(identities.Keys);
            return identities.Count;
        }
        finally { Release(maths); Release(content); }
    }

    internal static int FinalizeNewBatchAnchorsWithoutExport(
        Document document,
        IReadOnlyList<FormulaMetadata> metadataItems,
        StableRedrawInsertionOwners liveInsertions)
    {
        OMaths? maths = null;
        try
        {
            maths = document.OMaths;
            var indices = liveInsertions.CaptureIndices(maths);
            if (indices.Count != metadataItems.Count || indices.Values.Distinct().Count() != metadataItems.Count)
                throw new InvalidDataException("The completed OMML batch does not have independent physical owners.");
            foreach (var metadata in metadataItems)
            {
                if (!indices.TryGetValue(metadata.FormulaId, out var index))
                    throw new InvalidDataException("The completed OMML batch lost an insertion identity.");
                OMath? math = null;
                Range? range = null;
                Bookmark? bookmark = null;
                try
                {
                    math = maths[index + 1];
                    range = math.Range;
                    bookmark = WordOmmlFormulaStore.Wrap(document, range, metadata, replaceExisting: true);
                    if (!WordOmmlFormulaStore.IsCanonicalAnchor(bookmark, range))
                        throw new InvalidDataException("The completed OMML anchor did not retain its exact physical owner.");
                }
                finally { Release(bookmark); Release(range); Release(math); }
            }
            return metadataItems.Count;
        }
        finally { Release(maths); }
    }

    internal static int FinalizeNewBatchFingerprintsFromDocumentOpenXml(
        Document document,
        IReadOnlyList<FormulaMetadata> metadataItems,
        IReadOnlyDictionary<string, string> freshSourceOmml,
        StableRedrawInsertionOwners? liveInsertions = null)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (metadataItems is null) throw new ArgumentNullException(nameof(metadataItems));
        if (freshSourceOmml is null) throw new ArgumentNullException(nameof(freshSourceOmml));
        if (metadataItems.Count == 0) return 0;

        var metadata = metadataItems.ToDictionary(
            item => item.FormulaId,
            item => item,
            StringComparer.OrdinalIgnoreCase);
        if (metadata.Count != metadataItems.Count || freshSourceOmml.Count != metadata.Count)
            throw new InvalidDataException("The fresh OMML redraw batch has duplicate or missing identities.");
        foreach (var pair in metadata)
        {
            if (!freshSourceOmml.TryGetValue(pair.Key, out var source)
                || !string.Equals(
                    WordOmmlConverter.ComputeOmmlFingerprint(source),
                    pair.Value.NativeOmmlFingerprint,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException($"Fresh OMML redraw source {pair.Key} does not match its prepared identity.");
        }

        Range? content = null;
        OMaths? maths = null;
        try
        {
            content = document.Content;
            var package = XDocument.Parse(WordDocumentXml.Read(document));
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var body = package.Descendants(w + "body").Single();
            maths = document.OMaths;
            var xmlEquations = body.Descendants(m + "oMath").ToArray();
            if (maths.Count != xmlEquations.Length)
                throw new InvalidDataException("COM and Word XML disagree on the OMML inventory before redraw finalization.");

            Dictionary<string, string> expected;
            if (liveInsertions is not null)
            {
                if (liveInsertions.MaterializedSignatures.Count != metadata.Count
                    || metadata.Keys.Any(id => !liveInsertions.MaterializedSignatures.ContainsKey(id)))
                    throw new InvalidDataException(
                        "The fresh OMML redraw batch did not capture every materialized Word equation signature.");
                expected = liveInsertions.MaterializedSignatures.ToDictionary(
                    pair => pair.Key,
                    pair => pair.Value,
                    StringComparer.OrdinalIgnoreCase);
            }
            else
            {
                expected = freshSourceOmml.ToDictionary(
                    pair => pair.Key,
                    pair => WordOmmlConverter.ComputeImportedOmmlContentSignature(pair.Value),
                    StringComparer.OrdinalIgnoreCase);
            }
            var exactAnchorOwners = CaptureExactAnchorOwners(
                document,
                maths,
                expected.Keys.ToArray());
            // A collapsed Word bookmark has insertion gravity. During a large
            // redraw, later paragraph deletions/replacements can move that bookmark
            // onto a neighboring equation even though the freshly inserted OMath
            // itself remains healthy. Retain the actual OMath RCWs from insertion
            // and use them as independent physical ownership evidence. Existing
            // managed documents still rely on the strict anchor/fingerprint path;
            // only this fresh batch may repair a drifted anchor.
            var identities = IndexConversionEquationIdentities(
                body,
                expected,
                WordOmmlConverter.ComputeImportedOmmlContentSignature,
                exactAnchorOwners,
                liveInsertions?.CaptureIndices(maths),
                allowFreshOwnerFingerprintRecovery: true);
            if (identities.Count != metadata.Count)
                throw new InvalidDataException("The fresh OMML redraw identity index is incomplete.");

            foreach (var identity in identities)
            {
                var item = metadata[identity.Key];
                item.NativeOmmlFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                    xmlEquations[identity.Value.Index].ToString(SaveOptions.DisableFormatting));
                if (identity.Value.Canonical) continue;

                OMath? math = null;
                Range? range = null;
                Bookmark? rebound = null;
                try
                {
                    math = maths[identity.Value.Index + 1];
                    range = math.Range;
                    var actual = WordOmmlConverter.ComputeOmmlFingerprint(range.WordOpenXML);
                    if (!string.Equals(actual, item.NativeOmmlFingerprint, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidDataException($"COM equation content does not own fresh OMML identity {identity.Key}.");
                    rebound = WordOmmlFormulaStore.Wrap(
                        document,
                        range,
                        item,
                        replaceExisting: true);
                    if (!WordOmmlFormulaStore.IsCanonicalAnchor(rebound, range))
                        throw new InvalidDataException($"Word did not retain repaired fresh OMML anchor {identity.Key}.");
                }
                finally
                {
                    Release(rebound);
                    Release(range);
                    Release(math);
                }
            }
            return identities.Count;
        }
        finally
        {
            Release(maths);
            Release(content);
        }
    }

    private static IReadOnlyDictionary<string, int> CaptureExactAnchorOwners(
        Document document,
        OMaths maths,
        IReadOnlyCollection<string> formulaIds)
    {
        // Word may export a collapsed bookmark at an equation start as a child of
        // w:body, outside the equation paragraph. XML proximity cannot establish
        // that ownership; exact live positions can, without serializing each math.
        var starts = new Dictionary<(WdStoryType Story, int Start), int>();
        for (var index = 1; index <= maths.Count; index++)
        {
            OMath? math = null;
            Range? range = null;
            try
            {
                math = maths[index];
                range = math.Range;
                var key = (range.StoryType, range.Start);
                if (starts.ContainsKey(key))
                    throw new InvalidDataException("Two physical OMML equations have the same live start.");
                starts.Add(key, index - 1);
            }
            finally { Release(range); Release(math); }
        }
        var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var formulaId in formulaIds)
            {
                var name = WordOmmlFormulaStore.BookmarkName(formulaId);
                if (!bookmarks.Exists(name)) continue;
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    bookmark = bookmarks[name];
                    range = bookmark.Range;
                    if (range.Start == range.End
                        && starts.TryGetValue((range.StoryType, range.Start), out var index))
                        result.Add(formulaId, index);
                }
                finally { Release(range); Release(bookmark); }
            }
        }
        finally { Release(bookmarks); }
        return result;
    }

    // Exact live anchors and numbered rows establish physical ownership. XML
    // proximity alone never lets a body-level bookmark claim the following math.
    // Content fingerprints remain mandatory, and drift recovery must be unique.
    internal static IReadOnlyDictionary<string, (int Index, bool Canonical)> IndexConversionEquationIdentities(
        XElement body,
        IReadOnlyDictionary<string, string> expectedFingerprints,
        Func<string, string>? fingerprint = null,
        IReadOnlyDictionary<string, int>? exactAnchorOwners = null,
        IReadOnlyDictionary<string, int>? liveInsertedOwners = null,
        bool allowFreshOwnerFingerprintRecovery = false)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var equations = body.Descendants(m + "oMath").ToArray();
        fingerprint ??= WordOmmlConverter.ComputeOmmlFingerprint;
        var fingerprints = equations.Select(e => fingerprint(
            e.ToString(SaveOptions.DisableFormatting))).ToArray();
        var bookmarks = body.Descendants(w + "bookmarkStart")
            .GroupBy(e => (string?)e.Attribute(w + "name") ?? string.Empty, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.ToArray(), StringComparer.OrdinalIgnoreCase);
        var result = new Dictionary<string, (int Index, bool Canonical)>(StringComparer.OrdinalIgnoreCase);
        var claimed = new HashSet<int>();
        if (liveInsertedOwners is not null && liveInsertedOwners.Keys.Any(id => !expectedFingerprints.ContainsKey(id)))
            throw new InvalidDataException("A retained OMath is not owned by this fresh conversion batch.");
        foreach (var entry in expectedFingerprints)
        {
            if (string.IsNullOrWhiteSpace(entry.Value))
                throw new InvalidDataException($"Converted OMML {entry.Key} has no captured content fingerprint.");
            XElement? anchor = null;
            XElement? candidate = null;
            var name = WordOmmlFormulaStore.BookmarkName(entry.Key);
            if (bookmarks.TryGetValue(name, out var anchors))
            {
                if (anchors.Length != 1) throw new InvalidDataException($"OMML identity {name} is duplicated.");
                anchor = anchors[0];
            }
            var exactAnchor = exactAnchorOwners is not null
                && exactAnchorOwners.TryGetValue(entry.Key, out _);
            if (liveInsertedOwners is not null && liveInsertedOwners.TryGetValue(entry.Key, out var insertedOwner))
            {
                if (insertedOwner < 0 || insertedOwner >= equations.Length)
                    throw new InvalidDataException($"Retained OMML object {entry.Key} has no live equation index.");
                if (string.Equals(
                        fingerprints[insertedOwner],
                        entry.Value,
                        StringComparison.OrdinalIgnoreCase))
                {
                    candidate = equations[insertedOwner];
                    exactAnchor = exactAnchor
                        && anchor is not null
                        && exactAnchorOwners![entry.Key] == insertedOwner;
                }
                else if (allowFreshOwnerFingerprintRecovery)
                {
                    // Fresh redraw ownership is recovery evidence, not a reason to
                    // discard an otherwise healthy batch. Word can rebind both a
                    // collapsed VTOMML anchor and an insertion-time locator while
                    // earlier source paragraphs are replaced. Fall through to the
                    // prepared-content fingerprint resolver below. It still requires
                    // a unique physical equation, and claimed-index validation still
                    // prevents two logical identities from sharing one OMath.
                    candidate = null;
                    exactAnchor = false;
                }
                else
                {
                    throw new InvalidDataException(
                        $"Retained OMML object {entry.Key} owns different formula content.");
                }
            }
            else if (exactAnchor)
            {
                var owner = exactAnchorOwners![entry.Key];
                if (anchor is null || owner < 0 || owner >= equations.Length)
                    throw new InvalidDataException($"COM and XML disagree on OMML anchor {entry.Key}.");
                if (!string.Equals(fingerprints[owner], entry.Value, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException($"OMML anchor {entry.Key} owns different formula content.");
                candidate = equations[owner];
            }
            if (bookmarks.TryGetValue(WordEquationNumbering.NativeNumberBookmarkName(entry.Key), out var numbers))
            {
                if (numbers.Length != 1) throw new InvalidDataException($"OMML number identity {entry.Key} is duplicated.");
                XElement? numberOwner;
                var cell = numbers[0].Ancestors(w + "tc").FirstOrDefault();
                var row = cell?.Parent;
                if (row?.Name == w + "tr")
                {
                    var cells = row.Elements(w + "tc").ToArray();
                    if (cells.Length != 3 || cells[2] != cell)
                        throw new InvalidDataException($"OMML number {entry.Key} does not own column three of its row.");
                    var rowEquations = cells[1].Descendants(m + "oMath").ToArray();
                    if (rowEquations.Length != 1)
                        throw new InvalidDataException($"OMML number {entry.Key} has no single center equation.");
                    numberOwner = rowEquations[0];
                }
                else numberOwner = numbers[0].Ancestors(m + "oMath").FirstOrDefault();
                if (numberOwner is null)
                    throw new InvalidDataException($"OMML number {entry.Key} has no physical equation owner.");
                if (candidate is not null && candidate != numberOwner)
                    throw new InvalidDataException($"OMML anchor and number {entry.Key} own different equations.");
                candidate = numberOwner;
                if (!string.Equals(fingerprints[Array.IndexOf(equations, candidate)], entry.Value, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException($"OMML number {entry.Key} owns different formula content.");
            }
            var paragraph = anchor?.Ancestors(w + "p").FirstOrDefault();
            var adjacent = paragraph?.Descendants(m + "oMath")
                .FirstOrDefault(e => XNode.CompareDocumentOrder(anchor!, e) < 0);
            if (candidate is null && adjacent is not null
                && string.Equals(fingerprints[Array.IndexOf(equations, adjacent)], entry.Value, StringComparison.OrdinalIgnoreCase))
                candidate = adjacent;
            if (candidate is null)
            {
                var matches = Enumerable.Range(0, equations.Length)
                    .Where(i => string.Equals(fingerprints[i], entry.Value, StringComparison.OrdinalIgnoreCase)).ToArray();
                if (matches.Length != 1)
                    throw new InvalidDataException($"OMML identity {entry.Key} could not be recovered uniquely ({matches.Length} matching equations).");
                candidate = equations[matches[0]];
            }
            var index = Array.IndexOf(equations, candidate);
            if (!claimed.Add(index))
                throw new InvalidDataException("Two logical OMML identities claim the same physical equation.");
            var canonical = exactAnchor || (anchor is not null && adjacent == candidate);
            if (canonical && !exactAnchor)
            {
                var id = (string?)anchor!.Attribute(w + "id");
                var ends = paragraph!.Descendants(w + "bookmarkEnd").Where(e => (string?)e.Attribute(w + "id") == id).ToArray();
                canonical = ends.Length == 1 && XNode.CompareDocumentOrder(anchor, ends[0]) < 0
                    && XNode.CompareDocumentOrder(ends[0], candidate) < 0
                    && !paragraph.Descendants().Any(e => (e.Name == m + "t" || e.Name == w + "t" || e.Name == w + "instrText")
                        && XNode.CompareDocumentOrder(anchor, e) < 0 && XNode.CompareDocumentOrder(e, ends[0]) < 0);
            }
            result.Add(entry.Key, (index, canonical));
        }
        return result;
    }

    internal static string ReadCompleteEquationWordOpenXml(
        Document document,
        Range equationRange,
        string formulaId)
    {
        Range? content = null;
        Range? probe = null;
        Bookmarks? bookmarks = null;
        Bookmark? boundaryBookmark = null;
        Range? boundaryRange = null;
        try
        {
            content = document.Content;
            var probeEnd = equationRange.End;
            if (Guid.TryParse(formulaId, out var parsed))
            {
                bookmarks = document.Bookmarks;
                var boundaryName = "VTBL_" + parsed.ToString("N");
                if (bookmarks.Exists(boundaryName))
                {
                    boundaryBookmark = bookmarks[boundaryName];
                    boundaryRange = boundaryBookmark.Range;
                    // VTBL owns the ordinary-text typing anchor immediately after
                    // an inline formula. The anchor is a hard serialization bound,
                    // never part of the native equation.
                    probeEnd = Math.Max(probeEnd, boundaryRange.Start);
                }
            }

            // A field boundary can span dozens of Word structure characters.
            // Include the complete field, not merely its first marker; otherwise
            // Word may serialize only the leading fragment of a compound OMath.
            object start = equationRange.Start;
            object end = Math.Min(content.End, Math.Max(probeEnd, equationRange.End));
            var originalStart = equationRange.Start;
            var originalEnd = equationRange.End;
            var contentEnd = content.End;
            for (var attempt = 0; attempt < 2; attempt++)
            {
                probe = document.Range(ref start, ref end);
                var maths = probe.OMaths;
                try
                {
                    if (maths.Count != 1)
                        throw new InvalidDataException("The captured source does not contain exactly one live Word equation.");
                    var xml = probe.WordOpenXML;
                    if (equationRange.Start != originalStart || equationRange.End != originalEnd
                        || probe.Start != (int)start || probe.End != (int)end
                        || content.End != contentEnd || maths.Count != 1)
                        throw new InvalidDataException("The Word equation changed during native XML capture.");
                    XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
                    var count = XDocument.Parse(xml).Descendants(math + "oMath").Count();
                    if (count == 1)
                    {
                        WordOmmlConverter.ExtractSingleOMath(xml);
                        return xml;
                    }
                    if (count != 0)
                        throw new InvalidDataException("Word exported more than one equation for the captured source.");
                    WordDoubleClickHook.TraceMessage($"omml-xml-incomplete-export formulaId={formulaId} attempt={attempt + 1} range={originalStart}:{originalEnd}");
                }
                finally { Release(maths); Release(probe); probe = null; }
            }
            throw new InvalidDataException("Word could not export XML matching the captured live equation.");
        }
        finally
        {
            Release(boundaryRange);
            Release(boundaryBookmark);
            Release(bookmarks);
            Release(probe);
            Release(content);
        }
    }

    private static string ReadDisplayMode(Range equationRange)
    {
        OMaths? maths = null;
        OMath? selected = null;
        try
        {
            maths = equationRange.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? candidate = null;
                Range? candidateRange = null;
                try
                {
                    candidate = maths[index];
                    candidateRange = candidate.Range;
                    if (selected is null
                        || candidateRange.Start == equationRange.Start
                            && candidateRange.End == equationRange.End)
                    {
                        Release(selected);
                        selected = candidate;
                        candidate = null;
                        if (candidateRange.Start == equationRange.Start
                            && candidateRange.End == equationRange.End)
                            break;
                    }
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }
            return selected?.Type == WdOMathType.wdOMathDisplay
                ? "block"
                : "inline";
        }
        catch { return "inline"; }
        finally
        {
            Release(selected);
            Release(maths);
        }
    }

    private static double ReadFontSize(Range equationRange)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = equationRange.Font;
            var size = font.Size;
            return size > 0 && !float.IsNaN(size) && !float.IsInfinity(size)
                ? FormulaFontSize.Normalize(size)
                : FormulaFontSize.DefaultPt;
        }
        catch { return FormulaFontSize.DefaultPt; }
        finally { Release(font); }
    }

    private static void SanitizeMetadataBoundaryArtifacts(FormulaMetadata metadata)
    {
        foreach (var line in metadata.Lines)
            line.Latex = SanitizeFormulaBoundaryArtifacts(line.Latex);
        metadata.Latex = metadata.Lines.Count > 0
            ? string.Join("\n", metadata.Lines.Select(line => line.Latex))
            : SanitizeFormulaBoundaryArtifacts(metadata.Latex);
    }

    private static string SanitizeFormulaBoundaryArtifacts(string? latex)
    {
        if (string.IsNullOrEmpty(latex)) return string.Empty;
        static bool IsBoundaryArtifact(char character) =>
            character is '\u200B' or '\u200C' or '\u2060' or '\uFEFF';

        var value = latex!;
        for (var index = 0; index < value.Length; index++)
        {
            if (!IsBoundaryArtifact(value[index])) continue;
            var runEnd = index + 1;
            while (runEnd < value.Length && IsBoundaryArtifact(value[runEnd]))
                runEnd++;
            var left = value.Substring(0, index).Trim();
            var right = value.Substring(runEnd).Trim();
            if (left.Length > 0 && string.Equals(left, right, StringComparison.Ordinal))
                return left;
            index = runEnd - 1;
        }

        return new string(value.Where(character => !IsBoundaryArtifact(character)).ToArray());
    }

    private static FormulaMetadata Clone(FormulaMetadata metadata)
    {
        var clone = FormulaMetadataCodec.Decode(FormulaMetadataCodec.Encode(metadata));
        return clone
            ?? throw new InvalidDataException("Unable to clone VisualTeX formula metadata.");
    }

    private static void Release(object? value)
    {
        if (value is null || !System.Runtime.InteropServices.Marshal.IsComObject(value)) return;
        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(value); } catch { }
    }
}
