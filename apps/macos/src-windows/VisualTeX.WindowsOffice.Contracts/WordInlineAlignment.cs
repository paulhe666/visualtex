using System;

namespace VisualTeX.WindowsOffice.Contracts;

public static class WordInlineAlignment
{
    public static int CalculateFontPosition(
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline)
    {
        if (!(actualHeightPoints > 0)
            || !(exportedHeight > 0)
            || exportedBaseline is null
            || exportedBaseline < 0
            || exportedBaseline > exportedHeight)
            return 0;

        var descentRatio = (exportedHeight - exportedBaseline.Value) / exportedHeight;
        var downwardShiftPoints = actualHeightPoints * descentRatio;
        if (!(downwardShiftPoints > 0) || float.IsInfinity(downwardShiftPoints))
            return 0;

        return -Math.Max(
            0,
            (int)Math.Round(downwardShiftPoints, MidpointRounding.AwayFromZero));
    }
}
