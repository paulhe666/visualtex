using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Active local lookup must be observational. Word's layout-backed
/// wdWithInTable query can discard existing math-format revision provenance
/// without adding an Undo item. Use the frozen resolver's structural scope
/// index for native discovery; retain its existing VisualTeX implementation.
/// The frozen R3 binary/source and held R4 are deliberately not modified.
/// </summary>
internal static class WordFormulaHostLookup
{
    internal static WordFormulaHostDescriptor? ResolveLocal(Document document,
        Range selection, WordFormulaHostKind kind)
    {
        if (kind != WordFormulaHostKind.Omml)
            return WordFormulaHostResolver.ResolveLocal(document, selection, kind);
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (selection is null) throw new ArgumentNullException(nameof(selection));
        Range? preferred = null;
        Range? probe = null;
        Range? exact = null;
        OMaths? maths = null;
        OMath? math = null;
        try
        {
            // Preserve the existing exact/one-native-host preference before
            // broadening to the bounded local seam search.
            maths = selection.OMaths;
            if (maths.Count == 1)
            {
                math = maths[1];
                var current = math.Range;
                if (Touches(selection, current.Start, current.End, current.StoryType))
                    preferred = current;
                else Release(current);
            }
            probe = (preferred ?? selection).Duplicate;
            if (preferred is null)
            {
                probe.MoveStart(WdUnits.wdCharacter, -2);
                probe.MoveEnd(WdUnits.wdCharacter, 2);
            }
            // REF.Result can contain cached math. The containing local
            // paragraphs supply the actual field owners for exclusion, while
            // final selection filtering prevents adopting a neighboring host.
            probe.Expand(WdUnits.wdParagraph);
            var matches = WordFormulaHostResolver.CaptureScopeKindIndex(
                document, probe, WordFormulaHostKind.Omml).Where(host =>
                preferred is not null
                    ? host.Range.Start == preferred.Start && host.Range.End == preferred.End
                        && host.Range.StoryType == preferred.StoryType
                    : Touches(selection, host.Range.Start, host.Range.End, host.Range.StoryType))
                .GroupBy(host => (host.Range.StoryType, host.Range.Start, host.Range.End))
                .Select(group => group.First()).ToArray();
            if (matches.Length == 0) return null;
            if (matches.Length != 1)
                throw new InvalidDataException("The selected range resolves to more than one native OMML host.");
            var selected = matches[0];
            exact = WordFormulaHostSemanticReader.CreateRange(document, selected.Range);
            selected.Numbering = !selected.Display
                ? new WordFormulaNumberingDescriptor
                  { Numbered = false, Position = "right", ContainerKind = WordFormulaNumberingContainerKind.None }
                : WordNativeOmmlNumbering.TryResolveUnownedNativeStructure(document, exact, out var number)
                    ? number : WordFormulaNumberingResolver.ResolveOmmlAfterNativeMiss(exact);
            return selected;
        }
        finally { Release(exact); Release(probe); Release(preferred); Release(math); Release(maths); }
    }

    internal static bool IsWhollyInsideTable(Range range)
    {
        Tables? tables = null;
        Table? table = null;
        Range? bounds = null;
        try
        {
            tables = range.Tables;
            for (var index = 1; index <= tables.Count; index++)
            {
                Release(bounds); bounds = null;
                Release(table); table = tables[index];
                bounds = table.Range;
                if (range.InStory(bounds) && bounds.Start <= range.Start && bounds.End >= range.End)
                    return true;
            }
            return false;
        }
        finally { Release(bounds); Release(table); Release(tables); }
    }

    private static bool Touches(Range selection, int start, int end, WdStoryType story) =>
        selection.StoryType == story && (selection.Start == selection.End
            ? start <= selection.Start && selection.Start <= end
            : selection.Start < end && selection.End > start);

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
