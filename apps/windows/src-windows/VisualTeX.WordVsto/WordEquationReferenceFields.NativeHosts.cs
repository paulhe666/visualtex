using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static partial class WordEquationReferenceFields
{
    // User REF fields whose target bookmark lies inside a numbered native OMML
    // host (Word's hidden _Ref bound to the #(...) number). Format conversion
    // deletes that host together with the bookmark, so these references must be
    // migrated to the converted formula's number target.
    internal static IReadOnlyDictionary<string, int> CaptureNativeOmmlNumberReferences(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        if (host is null
            || host.Kind != WordFormulaHostKind.Omml
            || !host.Display)
            return new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        return NativeOmmlReferenceIndex.Capture(document).CountInside(host);
    }

    // The document's REF fields and their bookmark targets, read once. A batch
    // conversion asks this for every numbered OMML source; walking
    // document.Fields per source made the preflight quadratic in document size.
    // Valid only while the document is unchanged (conversion preflight).
    internal sealed class NativeOmmlReferenceIndex
    {
        private readonly List<(WdStoryType Story, int Start, int End, string Name)> _references = new();
        private readonly Dictionary<string, (WdStoryType Story, int Start, int End)?> _targets =
            new(StringComparer.OrdinalIgnoreCase);

        private NativeOmmlReferenceIndex() { }

        internal static NativeOmmlReferenceIndex Capture(Document document)
        {
            if (document is null)
                throw new ArgumentNullException(nameof(document));
            var index = new NativeOmmlReferenceIndex();
            Fields? fields = null;
            Bookmarks? bookmarks = null;
            try
            {
                fields = document.Fields;
                bookmarks = document.Bookmarks;
                var count = fields.Count;
                for (var position = 1; position <= count; position++)
                {
                    Field? field = null;
                    Range? code = null;
                    Bookmark? bookmark = null;
                    Range? target = null;
                    try
                    {
                        field = fields[position];
                        if (field.Type != WdFieldType.wdFieldRef) continue;
                        code = field.Code;
                        if (!TryReadReferenceBookmark(code.Text, out var name))
                            continue;
                        index._references.Add((code.StoryType, code.Start, code.End, name));
                        if (index._targets.ContainsKey(name)) continue;
                        (WdStoryType Story, int Start, int End)? resolved = null;
                        if (bookmarks.Exists(name))
                        {
                            bookmark = bookmarks[name];
                            target = bookmark.Range;
                            resolved = (target.StoryType, target.Start, target.End);
                        }
                        index._targets[name] = resolved;
                    }
                    finally
                    {
                        Release(target);
                        Release(bookmark);
                        Release(code);
                        Release(field);
                    }
                }
                return index;
            }
            finally
            {
                Release(bookmarks);
                Release(fields);
            }
        }

        // REF fields outside a display OMML host whose bookmark lies inside it.
        internal IReadOnlyDictionary<string, int> CountInside(WordFormulaHostDescriptor host)
        {
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (host is null
                || host.Kind != WordFormulaHostKind.Omml
                || !host.Display
                || _references.Count == 0)
                return counts;
            var story = host.Range.StoryType;
            var start = host.Range.Start;
            var end = host.Range.End;
            foreach (var reference in _references)
            {
                if (reference.Story == story && reference.Start >= start && reference.End <= end)
                    continue;
                if (_targets[reference.Name] is not { } target
                    || target.Story != story || target.Start < start || target.End > end)
                    continue;
                counts.TryGetValue(reference.Name, out var count);
                counts[reference.Name] = count + 1;
            }
            return counts;
        }
    }

    internal static string CreateNativeOmmlNumberReferenceBookmark(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        if (host is null
            || host.Kind != WordFormulaHostKind.Omml)
            throw new ArgumentException(
                "A native Word reference bookmark requires one OMML host.",
                nameof(host));

        var numbering =
            WordFormulaNumberingResolver.ResolveLocal(
                document,
                host);
        if (!numbering.Numbered
            || numbering.ContainerKind !=
                WordFormulaNumberingContainerKind.CanonicalNativeOmml
            || numbering.NumberRange is null)
            throw new InvalidDataException(
                "The converted OMML host has no canonical native equation-number range.");

        Range? numberRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? created = null;
        Range? createdRange = null;
        try
        {
            numberRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    numbering.NumberRange);
            bookmarks = document.Bookmarks;

            string? name = null;
            for (var attempt = 0;
                 attempt < 64;
                 attempt++)
            {
                var bytes =
                    Guid.NewGuid().ToByteArray();
                var value =
                    BitConverter.ToUInt32(
                        bytes,
                        0)
                    % 1_000_000_000u;
                var candidate =
                    "_Ref"
                    + value.ToString("D9");
                if (bookmarks.Exists(candidate))
                    continue;
                name = candidate;
                break;
            }
            if (name is null)
                throw new InvalidOperationException(
                    "VisualTeX could not allocate a unique native Word _Ref bookmark.");

            // A referenced pure OMML equation may own Word's ordinary hidden
            // _Ref bookmark, but never a VisualTeX FormulaId/private numbering
            // bookmark. Bind only the exact native number result (STYLEREF+SEQ
            // when applicable), so an existing number-only REF keeps displaying
            // exactly the same label after conversion.
            created =
                bookmarks.Add(
                    name,
                    numberRange);
            createdRange =
                created.Range.Duplicate;
            if (createdRange.StoryType !=
                    numberRange.StoryType
                || createdRange.Start !=
                    numberRange.Start
                || createdRange.End !=
                    numberRange.End)
                throw new InvalidDataException(
                    "Word changed the native equation-number range while creating its _Ref target.");

            return name;
        }
        finally
        {
            Release(createdRange);
            Release(created);
            Release(bookmarks);
            Release(numberRange);
        }
    }
}
