using System;

namespace VisualTeX.WindowsOffice.Contracts;

public static class WordInlineAlignment
{
    public const float LegacyDescentRatio = 0.25f;
    public const float WholePointSnapTolerancePoints = 0.0101f;

    public static int CalculateFontPositionWithLegacyFallback(
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline,
        float? existingFontPosition,
        double sourceSemanticFontSizePoints,
        double targetSemanticFontSizePoints)
    {
        if (HasValidExportedBaseline(exportedHeight, exportedBaseline))
            return CalculateFontPosition(
                actualHeightPoints,
                exportedHeight,
                exportedBaseline);

        var sourceSize = FormulaFontSize.Normalize(sourceSemanticFontSizePoints);
        var targetSize = FormulaFontSize.Normalize(targetSemanticFontSizePoints);
        if (existingFontPosition.HasValue
            && IsFinite(existingFontPosition.Value)
            && Math.Abs(existingFontPosition.Value) <= 256f
            && Math.Abs(existingFontPosition.Value) >= 0.01f)
        {
            // Word stores Font.Position as whole points. Remove half a point of
            // quantisation before scaling so an old -4 pt position at 14 pt maps
            // to -11 pt at 42 pt, rather than magnifying the original rounding
            // error to -12 pt. The same rule maps -11 pt back to -4 pt.
            var sign = Math.Sign(existingFontPosition.Value);
            var dequantizedMagnitude = Math.Max(
                0,
                Math.Abs(existingFontPosition.Value) - 0.5f);
            return sign * Math.Max(
                0,
                (int)Math.Round(
                    dequantizedMagnitude * targetSize / sourceSize,
                    MidpointRounding.AwayFromZero));
        }

        if (!(actualHeightPoints > 0) || !IsFinite(actualHeightPoints))
            return 0;
        return -Math.Max(
            0,
            (int)Math.Round(
                actualHeightPoints * LegacyDescentRatio,
                MidpointRounding.AwayFromZero));
    }

    public static int CalculateFontPosition(
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline)
    {
        if (!(actualHeightPoints > 0)
            || !IsFinite(actualHeightPoints)
            || !HasValidExportedBaseline(exportedHeight, exportedBaseline))
            return 0;

        var baseline = exportedBaseline.GetValueOrDefault();
        var descentRatio = (exportedHeight - baseline) / exportedHeight;
        var downwardShiftPoints = actualHeightPoints * descentRatio;
        if (!(downwardShiftPoints > 0) || float.IsInfinity(downwardShiftPoints))
            return 0;

        // Word exposes Font.Position only in whole points. Lower the object by
        // the complete whole-point part of its rendered descent. Float32 Word/
        // Office geometry can report an exact 10 pt descent as 9.994... pt, so
        // snap only values already within one hundredth of the next integer;
        // the extra 0.0001 pt absorbs Float32 representation at that boundary.
        // This tolerance is far below Word's one-point Position resolution and
        // avoids formula-type heuristics while preserving genuine sub-point
        // remainders such as 0.6 or 0.7 pt.
        var wholePointDescent = Math.Max(
            0,
            (int)Math.Floor(
                downwardShiftPoints + WholePointSnapTolerancePoints));
        return -wholePointDescent;
    }

    private static bool HasValidExportedBaseline(
        float exportedHeight,
        float? exportedBaseline) =>
        exportedHeight > 0
        && IsFinite(exportedHeight)
        && exportedBaseline.HasValue
        && IsFinite(exportedBaseline.Value)
        && exportedBaseline.Value >= 0
        && exportedBaseline.Value < exportedHeight;

    private static bool IsFinite(float value) =>
        !float.IsNaN(value) && !float.IsInfinity(value);
}
