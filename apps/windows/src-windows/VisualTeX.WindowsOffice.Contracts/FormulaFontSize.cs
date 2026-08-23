using System;
using System.Collections.Generic;
using System.Globalization;

namespace VisualTeX.WindowsOffice.Contracts;

public static class FormulaFontSize
{
    public const float DefaultPt = 14f;
    public const float MinimumPt = 5f;
    public const float MaximumPt = 200f;
    public const float StepPt = 0.5f;
    public const float PrecisionPt = 0.01f;

    private static readonly float[] Presets =
    {
        5f, 5.5f, 6.5f, 7.5f, 8f, 9f, 10f, 10.5f, 11f, 12f,
        14f, 15f, 16f, 18f, 20f, 22f, 24f, 26f, 28f, 36f,
        42f, 48f, 72f, 96f,
    };

    private static readonly (string Name, float Points)[] ChineseSizes =
    {
        ("初号", 42f),
        ("小初", 36f),
        ("一号", 26f),
        ("小一", 24f),
        ("二号", 22f),
        ("小二", 18f),
        ("三号", 16f),
        ("小三", 15f),
        ("四号", 14f),
        ("小四", 12f),
        ("五号", 10.5f),
        ("小五", 9f),
        ("六号", 7.5f),
        ("小六", 6.5f),
        ("七号", 5.5f),
        ("八号", 5f),
    };

    public static IReadOnlyList<float> StandardPresets => Presets;

    public static float Normalize(double? value, float fallback = DefaultPt)
    {
        var resolved = value.HasValue && IsPositiveFinite(value.Value)
            ? value.Value
            : fallback;
        resolved = Math.Max(MinimumPt, Math.Min(MaximumPt, resolved));
        // Point sizes are persisted as semantic document data, not as entries in
        // the half-point preset menu. Word and the Office dialog both allow typed
        // custom sizes, so quantizing here silently changed values such as 13.25
        // to 13.5 every time a formula was saved/reopened. Keep centipoint
        // precision for stable JSON/COM round-trips; StepPt remains the navigation
        // increment used by preset UI actions.
        return (float)Math.Round(resolved, 2, MidpointRounding.AwayFromZero);
    }

    public static float NextPreset(double? value)
    {
        var current = Normalize(value);
        foreach (var preset in Presets)
        {
            if (preset > current + 0.001f) return preset;
        }
        return Normalize(current + StepPt);
    }

    public static float PreviousPreset(double? value)
    {
        var current = Normalize(value);
        for (var index = Presets.Length - 1; index >= 0; index--)
        {
            if (Presets[index] < current - 0.001f) return Presets[index];
        }
        return Normalize(current - StepPt);
    }

    public static float Parse(string? value)
    {
        if (TryParse(value, out var result)) return result;
        throw new InvalidOperationException(
            "请输入 5–200 之间的公式字号，或输入初号、小初、一号、小一等中文字号。");
    }

    public static bool TryParse(string? value, out float fontSizePt)
    {
        fontSizePt = DefaultPt;
        var source = (value ?? string.Empty).Trim();
        if (source.Length == 0) return false;

        var compact = source
            .Replace(" ", string.Empty)
            .Replace("　", string.Empty)
            .Replace("字号", string.Empty);
        foreach (var size in ChineseSizes)
        {
            if (!string.Equals(compact, size.Name, StringComparison.Ordinal)) continue;
            fontSizePt = size.Points;
            return true;
        }

        var numeric = source
            .Replace("pt", string.Empty)
            .Replace("PT", string.Empty)
            .Replace("Pt", string.Empty)
            .Replace("pT", string.Empty)
            .Replace("磅", string.Empty)
            .Trim();
        if (!double.TryParse(
                numeric,
                NumberStyles.Float,
                CultureInfo.CurrentCulture,
                out var parsed)
            && !double.TryParse(
                numeric,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out parsed))
            return false;
        fontSizePt = Normalize(parsed);
        return true;
    }

    public static string FormatDisplay(double? value)
    {
        var normalized = Normalize(value);
        foreach (var size in ChineseSizes)
        {
            if (Math.Abs(size.Points - normalized) < 0.001f) return size.Name;
        }
        return normalized.ToString("0.##", CultureInfo.InvariantCulture);
    }

    public static string Describe(double? value)
    {
        var normalized = Normalize(value);
        var display = FormatDisplay(normalized);
        return display.EndsWith("号", StringComparison.Ordinal)
            || display.StartsWith("小", StringComparison.Ordinal)
            ? $"{display}（{normalized.ToString("0.##", CultureInfo.InvariantCulture)} 磅）"
            : $"{display} 磅";
    }

    public static float ResolveRenderFontSize(FormulaMetadata? metadata) =>
        Normalize(metadata?.RenderFontSizePt ?? metadata?.FontSizePt ?? DefaultPt);

    public static float ResolveSemanticFontSize(FormulaMetadata? metadata) =>
        Normalize(metadata?.FontSizePt ?? metadata?.RenderFontSizePt ?? DefaultPt);

    public static float InferOleFontSize(
        float currentWidthPoints,
        float currentHeightPoints,
        FormulaMetadata? metadata)
    {
        var fallback = ResolveSemanticFontSize(metadata);

        // Inline OLE formulas persist the exact Word dimensions that belonged to
        // the semantic font size at the last successful save. Prefer that
        // document-native checkpoint over re-deriving a size from browser render
        // pixels: Word/OLE can quantize the displayed extent slightly across a
        // save/reopen cycle even though the user never resized the formula.
        if (metadata?.WordInlineOleWidthPt is > 0
            && metadata.WordInlineOleHeightPt is > 0)
        {
            var storedWidth = (float)metadata.WordInlineOleWidthPt.Value;
            var storedHeight = (float)metadata.WordInlineOleHeightPt.Value;
            const float storedGeometryTolerancePoints = 0.75f;
            var widthMatchesStored = currentWidthPoints <= 0
                || Math.Abs(currentWidthPoints - storedWidth) <= storedGeometryTolerancePoints;
            var heightMatchesStored = currentHeightPoints <= 0
                || Math.Abs(currentHeightPoints - storedHeight) <= storedGeometryTolerancePoints;
            if (widthMatchesStored && heightMatchesStored) return fallback;

            var storedHeightScale = currentHeightPoints > 0
                ? currentHeightPoints / storedHeight
                : float.NaN;
            var storedWidthScale = currentWidthPoints > 0
                ? currentWidthPoints / storedWidth
                : float.NaN;
            var storedScale = IsPositiveFinite(storedHeightScale)
                ? storedHeightScale
                : storedWidthScale;
            if (IsPositiveFinite(storedScale))
                return Normalize(fallback * storedScale, fallback);
        }

        if (metadata?.RenderWidthPx is not > 0 || metadata.RenderHeightPx is not > 0)
            return fallback;

        // Font-size inference must use the preview's raw physical dimensions.
        // OfficeFormulaSizing.NaturalSize intentionally applies a 12 pt floor
        // for newly created object boxes. Small inline formulas are routinely
        // shorter than that floor (for example a 10.926 px preview is only
        // 8.195 pt high). Comparing the real Word height with the clamped 12 pt
        // value incorrectly turns an 11 pt formula into 7.5 pt on double-click.
        const float pointsPerPixel = 0.75f;
        var rawWidth = Math.Max(0.01f, (float)metadata.RenderWidthPx.Value * pointsPerPixel);
        var rawHeight = Math.Max(0.01f, (float)metadata.RenderHeightPx.Value * pointsPerPixel);
        // Word stores InlineShape dimensions at a coarse physical resolution.
        // For very small formulas the read-back height can differ from the raw
        // preview by several tenths of a point, enough to cross a 0.5 pt font
        // preset boundary even though the user never resized the object. Keep
        // the persisted semantic size whenever both dimensions are still within
        // normal Word geometry quantisation; only infer a new size after a real
        // physical resize.
        const float wordGeometryTolerancePoints = 0.75f;
        var widthMatches = currentWidthPoints <= 0
            || Math.Abs(currentWidthPoints - rawWidth) <= wordGeometryTolerancePoints;
        var heightMatches = currentHeightPoints <= 0
            || Math.Abs(currentHeightPoints - rawHeight) <= wordGeometryTolerancePoints;
        if (widthMatches && heightMatches) return fallback;

        var heightScale = currentHeightPoints > 0
            ? currentHeightPoints / rawHeight
            : float.NaN;
        var widthScale = currentWidthPoints > 0
            ? currentWidthPoints / rawWidth
            : float.NaN;
        var scale = IsPositiveFinite(heightScale)
            ? heightScale
            : widthScale;
        if (!IsPositiveFinite(scale)) return fallback;
        return Normalize(ResolveRenderFontSize(metadata) * scale, fallback);
    }

    public static (float Width, float Height) OleSizeAt(
        FormulaMetadata metadata,
        double targetFontSizePt,
        float maximumWidth = float.PositiveInfinity,
        float maximumHeight = float.PositiveInfinity)
    {
        if (metadata.RenderWidthPx is not > 0 || metadata.RenderHeightPx is not > 0)
            throw new InvalidOperationException(
                "VisualTeX formula metadata requires natural render dimensions for font sizing.");

        var natural = OfficeFormulaSizing.NaturalSize(
            (float)metadata.RenderWidthPx.Value,
            (float)metadata.RenderHeightPx.Value);
        var scale = Normalize(targetFontSizePt) / ResolveRenderFontSize(metadata);
        var width = natural.Width * scale;
        var height = natural.Height * scale;
        var fit = Math.Min(
            1f,
            Math.Min(
                IsPositiveFinite(maximumWidth) ? maximumWidth / width : 1f,
                IsPositiveFinite(maximumHeight) ? maximumHeight / height : 1f));
        if (fit > 0 && fit < 1f)
        {
            width *= fit;
            height *= fit;
        }
        return (Math.Max(1f, width), Math.Max(1f, height));
    }

    private static bool IsPositiveFinite(double value) =>
        value > 0 && !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool IsPositiveFinite(float value) =>
        value > 0 && !float.IsNaN(value) && !float.IsInfinity(value);
}
