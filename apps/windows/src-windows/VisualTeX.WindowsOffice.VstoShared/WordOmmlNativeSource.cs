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
                math = null;
            }
            finally { Release(range); Release(math); Release(maths); }
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
        }
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
            CreatedWithVersion = "1.2.6",
            UpdatedWithVersion = "1.2.6",
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
            var sanitizedStored = Clone(stored);
            SanitizeMetadataBoundaryArtifacts(sanitizedStored);
            var wordOpenXml = ReadCompleteEquationWordOpenXml(
                document,
                equationRange,
                stored.FormulaId);
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
        finally
        {
            Release(equationRange);
        }
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
        LiveInsertionOwners? liveInsertions = null)
    {
        if (formulaIds is null || formulaIds.Count == 0) return 0;
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
            return identities.Count;
        }
        finally { Release(maths); Release(content); }
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
        IReadOnlyDictionary<string, int>? liveInsertedOwners = null)
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
                if (anchor is null || insertedOwner < 0 || insertedOwner >= equations.Length
                    || !string.Equals(fingerprints[insertedOwner], entry.Value, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException($"Retained OMML object {entry.Key} owns different formula content.");
                candidate = equations[insertedOwner];
                exactAnchor = exactAnchor && exactAnchorOwners![entry.Key] == insertedOwner;
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
