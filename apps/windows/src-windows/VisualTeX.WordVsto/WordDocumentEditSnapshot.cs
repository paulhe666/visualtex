using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

// Read-only evidence for one owned Word undo record. The body is never copied
// back into Word. After native Undo, metadata/preferences and proven owned
// bookmark spans may be restored explicitly, followed by full verification.
internal sealed class WordDocumentEditSnapshot
{
    private readonly string bodySignature;
    private readonly string normalizedOriginalBody;
    private readonly int documentEnd;
    private readonly int evidenceStart = -1;
    private readonly int evidenceEnd;
    private readonly string mathFont;
    private readonly IReadOnlyList<string> metadata;
    private readonly IReadOnlyDictionary<string, string> variables;
    private readonly WordUndoHistorySnapshot undoHistory;
    private readonly string? diagnosticPath;
    private readonly WordBookmarkRecoverySnapshot? bookmarkRecovery;

    internal WordDocumentEditSnapshot(Document document, IEnumerable<string>? ownedBookmarkNames = null,
        Range? isolatedMutationScope = null)
    {
        var watch = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF") == "1"
            ? System.Diagnostics.Stopwatch.StartNew() : null;
        long checkpoint = 0;
        void Trace(string stage)
        {
            if (watch is null) return;
            var elapsed = watch.ElapsedMilliseconds;
            WordDoubleClickHook.TraceMessage($"document-checkpoint-perf stage={stage} deltaMs={elapsed - checkpoint} totalMs={elapsed}");
            checkpoint = elapsed;
        }
        Range? content = null;
        try
        {
            content = document.Content;
            documentEnd = content.End;
            if (isolatedMutationScope is not null)
            {
                if (isolatedMutationScope.StoryType != WdStoryType.wdMainTextStory
                    || isolatedMutationScope.Start < content.Start || isolatedMutationScope.End > content.End
                    || isolatedMutationScope.End <= isolatedMutationScope.Start)
                    throw new InvalidDataException("The isolated conversion recovery scope is not a valid body range.");
                evidenceStart = isolatedMutationScope.Start;
                evidenceEnd = isolatedMutationScope.End;
                Release(content); content = ReadEvidenceRange(document);
            }
            var originalXml = WordDocumentXml.Read(document, content);
            Trace("xml");
            var geometry = WordInlineObjectGeometry.Capture(content);
            bodySignature = WordLocalEditSnapshot.Signature(originalXml, geometry);
            normalizedOriginalBody = WordLocalEditSnapshot.NormalizedBody(originalXml, geometry);
            Trace("signature");
            if (ownedBookmarkNames is not null)
                bookmarkRecovery = new WordBookmarkRecoverySnapshot(document, content,
                    WordLocalEditSnapshot.NormalizedBody(originalXml, geometry), ownedBookmarkNames);
            Trace("bookmarks");
            if (Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_RECOVERY_XML") == "1")
            {
                diagnosticPath = (Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH")
                    ?? throw new InvalidOperationException("Diagnostic evidence needs a configured trace path."))
                    + ".recovery-" + Guid.NewGuid().ToString("N");
                File.WriteAllText(diagnosticPath + "-before.xml", WordLocalEditSnapshot.NormalizedBody(originalXml, geometry));
            }
            mathFont = document.OMathFontName;
            metadata = ReadMetadata(document);
            variables = ReadVariables(document);
            Trace("metadata-variables");
            undoHistory = new WordUndoHistorySnapshot(document);
            Trace("undo-history");
        }
        finally { Release(content); }
    }

    private Range ReadEvidenceRange(Document document) => evidenceStart < 0
        ? document.Content : document.Range(evidenceStart, evidenceEnd);

    private bool BodyMatches(Document document)
    {
        Range? content = null;
        try
        {
            content = document.Content;
            if (content.End != documentEnd) return false;
            Release(content); content = ReadEvidenceRange(document);
            var xml = WordDocumentXml.Read(document, content);
            var geometry = WordInlineObjectGeometry.Capture(content);
            var matches = WordLocalEditSnapshot.Signature(xml, geometry) == bodySignature;
            if (!matches && WordRecoveryXmlEquivalence.HasOnlyAutomaticChineseProofingAdditions(
                    normalizedOriginalBody, WordLocalEditSnapshot.NormalizedBody(xml, geometry)))
            {
                // Native Undo restores the actual text and layout. Word may then
                // fill previously implicit Chinese proofing annotations while
                // repaginating; explicit languages, fonts and content still match.
                WordDoubleClickHook.TraceMessage("word-undo-body-equivalent automaticChineseProofingOnly=True");
                matches = true;
            }
            if (!matches && diagnosticPath is not null)
                File.WriteAllText(diagnosticPath + "-after.xml", WordLocalEditSnapshot.NormalizedBody(xml, geometry));
            return matches;
        }
        finally { Release(content); }
    }

    internal void Restore(Document document)
    {
        if (diagnosticPath is not null)
        {
            // Preserve the actual failed structure before native Undo restores it.
            // This optional read-only evidence must never prevent recovery.
            try { File.WriteAllText(diagnosticPath + "-failed.xml", WordDocumentXml.Read(document)); }
            catch (Exception error)
            {
                WordDoubleClickHook.TraceMessage($"format-conversion-failed-evidence-error error={error}");
            }
        }
        // The caller ended any remaining custom record. Serialization may have
        // split it earlier; verify and undo exactly this operation's native prefix.
        undoHistory.UndoChanges(document);
        var undoBodyMatches = BodyMatches(document);
        WordDoubleClickHook.TraceMessage($"format-conversion-recovery afterUndoBodyMatches={undoBodyMatches}");
        if (bookmarkRecovery is not null && !undoBodyMatches)
        {
            Range? content = null;
            try
            {
                content = document.Content;
                if (content.End != documentEnd)
                    throw new InvalidDataException("Word undo has not restored the original document extent.");
                Release(content); content = ReadEvidenceRange(document);
                bookmarkRecovery.Restore(document, WordLocalEditSnapshot.NormalizedBody(
                    WordDocumentXml.Read(document, content), WordInlineObjectGeometry.Capture(content)));
            }
            finally { Release(content); }
        }
        if (document.OMathFontName != mathFont) document.OMathFontName = mathFont;
        if (!ReadMetadata(document).SequenceEqual(metadata))
            RestoreMetadata(document, metadata);
        RestoreVariables(document, variables);
        WordOmmlFormulaStore.InvalidateDocumentCache(document);
        var bodyRestored = BodyMatches(document);
        var fontRestored = document.OMathFontName == mathFont;
        var metadataRestored = ReadMetadata(document).SequenceEqual(metadata);
        var variablesRestored = ReadVariables(document).OrderBy(p => p.Key).SequenceEqual(variables.OrderBy(p => p.Key));
        WordDoubleClickHook.TraceMessage(
            $"format-conversion-recovery-verified body={bodyRestored} scope={(evidenceStart < 0 ? "document" : evidenceStart + ":" + evidenceEnd)} mathFont={fontRestored} metadata={metadataRestored} variables={variablesRestored}");
        if (!bodyRestored || !fontRestored || !metadataRestored || !variablesRestored)
            throw new InvalidDataException("Word undo did not restore the original content, numbering, bookmarks, formatting and metadata.");
    }

    private static IReadOnlyList<string> ReadMetadata(Document document)
    {
        object? parts = null;
        object? selected = null;
        var values = new List<string>();
        try
        {
            parts = ((dynamic)document).CustomXMLParts;
            selected = ((dynamic)parts).SelectByNamespace(WordOmmlFormulaStore.NamespaceUri);
            for (var i = 1; i <= (int)((dynamic)selected).Count; i++)
            {
                object? part = null;
                try
                {
                    part = ((dynamic)selected)[i];
                    values.Add((string)((dynamic)part).XML);
                }
                finally { Release(part); }
            }
            return values.OrderBy(x => x, StringComparer.Ordinal).ToArray();
        }
        finally { Release(selected); Release(parts); }
    }

    private static void RestoreMetadata(Document document, IReadOnlyList<string> original)
    {
        object? parts = null;
        object? selected = null;
        try
        {
            parts = ((dynamic)document).CustomXMLParts;
            selected = ((dynamic)parts).SelectByNamespace(WordOmmlFormulaStore.NamespaceUri);
            var missing = original.ToList();
            for (var i = (int)((dynamic)selected).Count; i >= 1; i--)
            {
                object? part = null;
                try
                {
                    part = ((dynamic)selected)[i];
                    var xml = (string)((dynamic)part).XML;
                    if (!missing.Remove(xml)) ((dynamic)part).Delete();
                }
                finally { Release(part); }
            }
            foreach (var xml in missing)
            {
                object? part = null;
                try { part = ((dynamic)parts).Add(xml); }
                finally { Release(part); }
            }
        }
        finally { Release(selected); Release(parts); }
    }

    private static IReadOnlyDictionary<string, string> ReadVariables(Document document)
    {
        Variables? collection = null;
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        try
        {
            collection = document.Variables;
            for (var i = 1; i <= collection.Count; i++)
            {
                Variable? variable = null;
                try
                {
                    variable = collection[i];
                    values.Add(variable.Name, variable.Value);
                }
                finally { Release(variable); }
            }
            return values;
        }
        finally { Release(collection); }
    }

    private static void RestoreVariables(Document document, IReadOnlyDictionary<string, string> original)
    {
        Variables? collection = null;
        try
        {
            collection = document.Variables;
            var current = ReadVariables(document);
            foreach (var item in current)
            {
                if (original.TryGetValue(item.Key, out var value) && value == item.Value) continue;
                if (!item.Key.StartsWith("VisualTeX", StringComparison.Ordinal))
                    throw new InvalidDataException("An unrelated document variable changed during conversion.");
                Variable? variable = null;
                try
                {
                    variable = collection[item.Key];
                    if (original.ContainsKey(item.Key)) variable.Value = original[item.Key];
                    else variable.Delete();
                }
                finally { Release(variable); }
            }
            foreach (var item in original.Where(p => !current.ContainsKey(p.Key)))
            {
                Variable? variable = null;
                try { variable = collection.Add(item.Key, item.Value); }
                finally { Release(variable); }
            }
        }
        finally { Release(collection); }
    }

    private static void Release(object? value)
    {
        if (value is not null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            System.Runtime.InteropServices.Marshal.ReleaseComObject(value);
    }
}
