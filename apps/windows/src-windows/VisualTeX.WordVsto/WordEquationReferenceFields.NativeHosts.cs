using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static partial class WordEquationReferenceFields
{
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
