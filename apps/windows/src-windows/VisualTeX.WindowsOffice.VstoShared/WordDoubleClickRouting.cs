using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.VstoShared;

internal static class WordDoubleClickRouting
{
    internal static bool ShouldOpenVisualTeX(OfficeSelection? selection)
    {
        if (selection?.Metadata is null
            || string.IsNullOrWhiteSpace(selection.FormulaId))
            return false;

        // Every VisualTeX-owned object, including Word-native OMML, reopens the
        // VisualTeX Session editor. Ordinary Word equations have no VisualTeX
        // metadata and therefore still keep Word's native equation editor.
        return true;
    }
}
