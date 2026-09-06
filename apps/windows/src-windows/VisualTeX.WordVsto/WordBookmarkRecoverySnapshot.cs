using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

// Word's native Undo can restore caption text/Frames while retaining a shortened
// bookmark. Restore only names owned by this operation, after proving that every
// other part of the body (including unrelated bookmarks) has already recovered.
internal sealed class WordBookmarkRecoverySnapshot
{
    private sealed class Span
    {
        internal string Name = string.Empty;
        internal int Start;
        internal int End;
        internal string Text = string.Empty;
    }

    private readonly List<Span> spans = new();
    private readonly string originalBody;

    internal static IEnumerable<string> NamesForFormula(string formulaId)
    {
        var suffix = Guid.Parse(formulaId).ToString("N");
        return new[] { "VTO_", "VTOMML_", "VTBL_", "VTEq_", "VTEqCap_", "VTEqNum_", "VTEqAnc_", "VTAncR_" }
            .Select(prefix => prefix + suffix);
    }

    internal WordBookmarkRecoverySnapshot(Document document, Range scope, string normalizedBody,
        IEnumerable<string> ownedNames)
    {
        originalBody = normalizedBody;
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var name in ownedNames.Distinct(StringComparer.Ordinal))
            {
                if (!bookmarks.Exists(name)) continue;
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    bookmark = bookmarks[name];
                    range = bookmark.Range;
                    if (range.StoryType != scope.StoryType || range.Start < scope.Start || range.End > scope.End)
                        continue;
                    spans.Add(new Span { Name = name, Start = range.Start, End = range.End, Text = range.Text ?? string.Empty });
                }
                finally { Release(range); Release(bookmark); }
            }
        }
        finally { Release(bookmarks); }
    }

    internal void Restore(Document document, string currentBody)
    {
        Bookmarks? bookmarks = null;
        var changed = new List<Span>();
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var span in spans)
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    if (!bookmarks.Exists(span.Name)) { changed.Add(span); continue; }
                    bookmark = bookmarks[span.Name];
                    range = bookmark.Range;
                    if (range.StoryType != WdStoryType.wdMainTextStory)
                        throw new InvalidDataException("An owned bookmark moved to another Word story during recovery.");
                    if (range.Start != span.Start || range.End != span.End || (range.Text ?? string.Empty) != span.Text)
                        changed.Add(span);
                }
                finally { Release(range); Release(bookmark); }
            }
            if (changed.Count == 0) return;
            VerifyOnlyBookmarkSpansChanged(originalBody, currentBody, changed.Select(span => span.Name).ToArray());
            // Validate every destination before the first mutation. This does not
            // delete bookmark contents, Frames, fields or reference aliases.
            foreach (var span in changed)
            {
                Range? range = null;
                try
                {
                    range = document.Range(span.Start, span.End);
                    if ((range.Text ?? string.Empty) != span.Text)
                        throw new InvalidDataException("The original bookmark contents have not recovered.");
                }
                finally { Release(range); }
            }
            foreach (var span in changed)
            {
                Range? range = null;
                Bookmark? restored = null;
                try
                {
                    range = document.Range(span.Start, span.End);
                    restored = bookmarks.Add(span.Name, range);
                    WordDoubleClickHook.TraceMessage($"word-undo-bookmark-restored name={span.Name} range={span.Start}:{span.End}");
                }
                finally { Release(restored); Release(range); }
            }
            // The caller must still compare its original, bookmark-inclusive body
            // signature after these exact same-name COM bindings.
        }
        finally { Release(bookmarks); }
    }

    internal static void VerifyOnlyBookmarkSpansChanged(string before, string after, IReadOnlyCollection<string> changedNames)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var original = XElement.Parse(before);
        var current = XElement.Parse(after);
        foreach (var name in changedNames)
        {
            if (original.Descendants(w + "bookmarkStart").Count(e => (string?)e.Attribute(w + "name") == name) != 1
                || original.Descendants(w + "bookmarkEnd").Count(e => (string?)e.Attribute(w + "id") == name) != 1)
                throw new InvalidDataException("Bookmark recovery requires one complete original owned span.");
            foreach (var body in new[] { original, current })
                body.Descendants().Where(e =>
                    (e.Name == w + "bookmarkStart" || e.Name == w + "bookmarkEnd")
                    && (string?)e.Attribute(w + "id") == name).Remove();
        }
        if (!XNode.DeepEquals(original, current))
            throw new InvalidDataException("Content, formatting or unrelated bookmarks still differ after Word undo; owned bookmark recovery was not applied.");
    }

    private static void Release(object? value)
    {
        if (value is not null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            System.Runtime.InteropServices.Marshal.ReleaseComObject(value);
    }
}
