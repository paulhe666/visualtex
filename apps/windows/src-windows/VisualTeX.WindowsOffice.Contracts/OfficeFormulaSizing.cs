using System;

namespace VisualTeX.WindowsOffice.Contracts;

public static class OfficeFormulaSizing
{
    private const float PointsPerPixel = 0.75f;
    private const float MinimumDimensionPoints = 1f;

    public static (float Width, float Height) NaturalSize(
        float renderWidth,
        float renderHeight)
    {
        // Render dimensions are CSS pixels at 96 dpi; Word stores object sizes
        // in 72 dpi points. Do not impose a 12 pt selection-box floor here:
        // scaling a narrow inline formula such as x up to 12 pt width enlarges
        // its glyph by two to three times even though metadata still says 11 pt.
        return (
            Math.Max(MinimumDimensionPoints, Math.Max(1f, renderWidth) * PointsPerPixel),
            Math.Max(MinimumDimensionPoints, Math.Max(1f, renderHeight) * PointsPerPixel));
    }

    public static (float Width, float Height) EditedSize(
        float currentWidth,
        float currentHeight,
        double? originalRenderWidth,
        double? originalRenderHeight,
        float newRenderWidth,
        float newRenderHeight,
        float maximumWidth = float.PositiveInfinity,
        float maximumHeight = float.PositiveInfinity,
        double? originalFontSizePt = null,
        double? originalRenderFontSizePt = null,
        bool preserveIndependentAxisScale = false)
    {
        var next = NaturalSize(newRenderWidth, newRenderHeight);
        var scale = 1f;
        var horizontalScale = float.NaN;
        var verticalScale = float.NaN;
        if (originalRenderWidth is > 0 && originalRenderHeight is > 0
            && currentWidth > 0 && currentHeight > 0)
        {
            // Use the preview's real physical dimensions when recovering the
            // existing object scale. NaturalSize applies a 12 pt minimum box,
            // which is useful for creating/selecting tiny objects but is not a
            // valid font-scale reference. A 25x11 px inline formula is really
            // 18.75x8.25 pt; comparing its 8.5 pt Word height with a clamped
            // 12 pt height incorrectly shrinks a wider edited formula to 71%.
            var previousWidth = Math.Max(0.01f,
                (float)originalRenderWidth.Value * PointsPerPixel);
            var previousHeight = Math.Max(0.01f,
                (float)originalRenderHeight.Value * PointsPerPixel);
            horizontalScale = currentWidth / previousWidth;
            verticalScale = currentHeight / previousHeight;

            // Formula height is the visual font-size reference. Prefer it over
            // the geometric mean so picture→OLE conversion cannot become
            // shorter and wider when the old picture box was non-uniform or
            // came from a legacy raster export.
            if (IsPositiveFinite(verticalScale))
                scale = verticalScale;
            else if (IsPositiveFinite(horizontalScale))
                scale = horizontalScale;
        }
        else if (currentHeight > 0 && IsPositiveFinite(next.Height))
        {
            // Legacy picture metadata can lack natural render dimensions.
            // Preserve its physical height and recover the new width from the
            // replacement formula's natural aspect ratio.
            scale = currentHeight / next.Height;
        }
        else if (currentWidth > 0 && IsPositiveFinite(next.Width))
        {
            scale = currentWidth / next.Width;
        }
        var semanticScale = 1f;
        if (originalFontSizePt is > 0 && originalRenderFontSizePt is > 0)
        {
            semanticScale = FormulaFontSize.Normalize(originalFontSizePt)
                / FormulaFontSize.Normalize(originalRenderFontSizePt);
            if (IsPositiveFinite(semanticScale)) scale /= semanticScale;
            else semanticScale = 1f;
        }
        if (!IsPositiveFinite(scale)) scale = 1f;
        scale = Math.Max(0.1f, Math.Min(10f, scale));

        float width;
        float height;
        if (preserveIndependentAxisScale
            && IsPositiveFinite(horizontalScale)
            && IsPositiveFinite(verticalScale))
        {
            // A VisualTeX inline OLE can intentionally carry different Word X/Y
            // presentation scales after a format conversion. Treat those scales as
            // host presentation state, not as semantic font-size evidence. Applying
            // only the height scale on every edit slowly changes the width even when
            // the user restores the original LaTeX. Preserve each axis independently
            // so content edits are reversible while the semantic/render font size is
            // unchanged. Other families keep the established uniform-height rule.
            var widthScale = horizontalScale / semanticScale;
            var heightScale = verticalScale / semanticScale;
            if (!IsPositiveFinite(widthScale)) widthScale = scale;
            if (!IsPositiveFinite(heightScale)) heightScale = scale;
            widthScale = Math.Max(0.1f, Math.Min(10f, widthScale));
            heightScale = Math.Max(0.1f, Math.Min(10f, heightScale));
            width = next.Width * widthScale;
            height = next.Height * heightScale;
        }
        else
        {
            width = next.Width * scale;
            height = next.Height * scale;
        }
        var fitScale = Math.Min(
            1f,
            Math.Min(
                IsPositiveFinite(maximumWidth) ? maximumWidth / width : 1f,
                IsPositiveFinite(maximumHeight) ? maximumHeight / height : 1f));
        if (IsPositiveFinite(fitScale) && fitScale < 1f)
        {
            width *= fitScale;
            height *= fitScale;
        }
        return (Math.Max(1f, width), Math.Max(1f, height));
    }

    private static bool IsPositiveFinite(float value) =>
        value > 0 && !float.IsNaN(value) && !float.IsInfinity(value);
}
