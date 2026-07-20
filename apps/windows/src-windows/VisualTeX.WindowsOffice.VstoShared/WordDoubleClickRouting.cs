using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.VstoShared;

internal static class WordDoubleClickRouting
{
    internal static bool ShouldOpenVisualTeX(OfficeSelection? selection)
    {
        if (selection?.Metadata is null
            || string.IsNullOrWhiteSpace(selection.FormulaId))
            return false;

        // Word-native OMML must keep Word's own equation editor. Embedded OLE
        // and cross-platform picture formulas are edited by the VisualTeX
        // session window instead of invoking the OLE server's default verb.
        return !string.Equals(
            selection.ObjectMode,
            FormulaOleContract.WordOmmlMode,
            StringComparison.Ordinal);
    }
}
