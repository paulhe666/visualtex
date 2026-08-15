using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.VstoShared;

internal static class WordDoubleClickRouting
{
    internal static bool ShouldOpenVisualTeX(OfficeSelection? selection)
    {
        if (selection?.Metadata is null
            || string.IsNullOrWhiteSpace(selection.FormulaId))
            return false;

        // At this point the selection carries VisualTeX metadata. Native Word
        // OMML may receive that metadata during its first coordinate hit-test,
        // so unmanaged OMath objects (including MathType-converted OMML) can
        // open VisualTeX immediately on the first double-click.
        return true;
    }

    internal static bool ScreenPointHitsFormulaRectangle(
        int screenX,
        int screenY,
        int left,
        int top,
        int width,
        int height,
        int padding = 0)
    {
        if (width <= 0 || height <= 0 || padding < 0) return false;

        var paddedLeft = (long)left - padding;
        var paddedTop = (long)top - padding;
        var paddedRight = (long)left + width + padding;
        var paddedBottom = (long)top + height + padding;
        return screenX >= paddedLeft
            && screenX <= paddedRight
            && screenY >= paddedTop
            && screenY <= paddedBottom;
    }
}
