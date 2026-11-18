using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace VisualTeX.WindowsOffice.Contracts;

public sealed class FormulaMetadata
{
    [JsonPropertyName("schema")]
    public string Schema { get; set; } = "visualtex-formula";

    [JsonPropertyName("schemaVersion")]
    public int SchemaVersion { get; set; } = 1;

    [JsonPropertyName("formulaId")]
    public string FormulaId { get; set; } = string.Empty;

    [JsonPropertyName("title")]
    public string Title { get; set; } = string.Empty;

    [JsonPropertyName("latex")]
    public string Latex { get; set; } = string.Empty;

    [JsonPropertyName("lines")]
    public List<FormulaLine> Lines { get; set; } = new();

    [JsonPropertyName("codeFormat")]
    public string CodeFormat { get; set; } = string.Empty;

    [JsonPropertyName("displayMode")]
    public string DisplayMode { get; set; } = "block";

    [JsonPropertyName("numbered")]
    public bool Numbered { get; set; }

    [JsonPropertyName("equationTag")]
    public string? EquationTag { get; set; }

    [JsonPropertyName("renderWidthPx")]
    public double? RenderWidthPx { get; set; }

    [JsonPropertyName("renderHeightPx")]
    public double? RenderHeightPx { get; set; }

    [JsonPropertyName("baseline")]
    public double? Baseline { get; set; }

    [JsonPropertyName("fontSizePt")]
    public double? FontSizePt { get; set; }

    [JsonPropertyName("renderFontSizePt")]
    public double? RenderFontSizePt { get; set; }

    [JsonPropertyName("formulaLetterFont")]
    public string? FormulaLetterFont { get; set; }

    [JsonPropertyName("formulaChineseFont")]
    public string? FormulaChineseFont { get; set; }

    /// <summary>The fonts for letters, Greek letters and digits the formula was drawn in (absent: none).</summary>
    [JsonPropertyName("formulaFaces")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public FormulaFaces? FormulaFaces { get; set; }

    [JsonPropertyName("wordInlineOleWidthPt")]
    public double? WordInlineOleWidthPt { get; set; }

    [JsonPropertyName("wordInlineOleHeightPt")]
    public double? WordInlineOleHeightPt { get; set; }

    [JsonPropertyName("wordInlineOlePositionPt")]
    public int? WordInlineOlePositionPt { get; set; }

    [JsonPropertyName("wordInlineSourceBottomWhitespacePt")]
    public double? WordInlineSourceBottomWhitespacePt { get; set; }

    [JsonPropertyName("wordInlinePreviewShiftPx")]
    public double? WordInlinePreviewShiftPx { get; set; }

    [JsonPropertyName("wordDisplayPreviewInkHeightRatio")]
    public double? WordDisplayPreviewInkHeightRatio { get; set; }

    [JsonPropertyName("wordDisplayPreviewBottomWhitespaceRatio")]
    public double? WordDisplayPreviewBottomWhitespaceRatio { get; set; }

    [JsonPropertyName("nativeOmmlFingerprint")]
    public string? NativeOmmlFingerprint { get; set; }

    [JsonPropertyName("createdWithVersion")]
    public string CreatedWithVersion { get; set; } = string.Empty;

    [JsonPropertyName("updatedWithVersion")]
    public string UpdatedWithVersion { get; set; } = string.Empty;

    [JsonPropertyName("createdAt")]
    public string CreatedAt { get; set; } = string.Empty;

    [JsonPropertyName("updatedAt")]
    public string UpdatedAt { get; set; } = string.Empty;

    /// <summary>
    /// Makes metadata read from a document usable: the formula's identity and
    /// LaTeX stay; derived layout and presentation values that are out of range
    /// (another build or platform, a rounding residue) are dropped so the add-in
    /// recomputes them. Validate() stays strict for metadata the add-in writes.
    /// </summary>
    public void NormalizeForRead()
    {
        static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
        static bool SupportedSize(double value) =>
            Finite(value) && value >= FormulaFontSize.MinimumPt && value <= FormulaFontSize.MaximumPt;
        var display = string.Equals(DisplayMode, "block", StringComparison.Ordinal);
        var inline = string.Equals(DisplayMode, "inline", StringComparison.OrdinalIgnoreCase);
        Lines ??= new List<FormulaLine>();
        if (Lines.Count == 0 && !string.IsNullOrWhiteSpace(Latex))
            Lines.Add(new FormulaLine { Id = Guid.NewGuid().ToString("D"), Latex = Latex });
        if (Numbered && !display) Numbered = false;
        if (!string.IsNullOrWhiteSpace(EquationTag) && (!display || EquationTag!.Length > 256)) EquationTag = null;
        if (RenderWidthPx is double width && !(width > 0 && Finite(width))) RenderWidthPx = null;
        if (RenderHeightPx is double height && !(height > 0 && Finite(height))) RenderHeightPx = null;
        if (Baseline is double baseline
            && (!Finite(baseline) || baseline < 0 || (RenderHeightPx.HasValue && baseline > RenderHeightPx.Value)))
            Baseline = null;
        if (FontSizePt is double fontSize && !SupportedSize(fontSize)) FontSizePt = null;
        if (RenderFontSizePt is double renderSize && !SupportedSize(renderSize)) RenderFontSizePt = null;
        if (FormulaLetterFont is not null && !IsSupportedFormulaLetterFont(FormulaLetterFont)) FormulaLetterFont = null;
        if (FormulaChineseFont is not null && !IsSupportedFormulaChineseFont(FormulaChineseFont)) FormulaChineseFont = null;
        if (FormulaFaces is not null
            && ((FormulaFaces.Letter ?? string.Empty).Length > 128 || (FormulaFaces.Greek ?? string.Empty).Length > 128
                || (FormulaFaces.Digit ?? string.Empty).Length > 128))
            FormulaFaces = null;
        if (WordInlineOleWidthPt.HasValue != WordInlineOleHeightPt.HasValue || !inline
            || WordInlineOleWidthPt is double oleWidth && !(oleWidth > 0 && Finite(oleWidth))
            || WordInlineOleHeightPt is double oleHeight && !(oleHeight > 0 && Finite(oleHeight)))
        {
            WordInlineOleWidthPt = null;
            WordInlineOleHeightPt = null;
        }
        if (WordInlineOlePositionPt is int position && (!inline || position < -256 || position > 256))
            WordInlineOlePositionPt = null;
        if (WordInlineSourceBottomWhitespacePt is double bottom && (!inline || !Finite(bottom) || bottom < 0 || bottom > 256))
            WordInlineSourceBottomWhitespacePt = null;
        if (WordInlinePreviewShiftPx is double shift && (!inline || !Finite(shift) || shift < -2 || shift > 2))
            WordInlinePreviewShiftPx = null;
        if (WordDisplayPreviewInkHeightRatio.HasValue != WordDisplayPreviewBottomWhitespaceRatio.HasValue
            || WordDisplayPreviewInkHeightRatio is double ink
                && (!string.Equals(DisplayMode, "block", StringComparison.OrdinalIgnoreCase)
                    || !Finite(ink) || ink <= 0 || ink > 1
                    || WordDisplayPreviewBottomWhitespaceRatio is not double whitespace
                    || !Finite(whitespace) || whitespace < 0 || whitespace > 1 || ink + whitespace > 1.001))
        {
            WordDisplayPreviewInkHeightRatio = null;
            WordDisplayPreviewBottomWhitespaceRatio = null;
        }
    }

    public void Validate()
    {
        if (Schema != "visualtex-formula" || SchemaVersion != 1)
            throw new InvalidOperationException("Unsupported VisualTeX formula metadata schema.");
        if (!Guid.TryParse(FormulaId, out _))
            throw new InvalidOperationException("VisualTeX formulaId must be a UUID.");
        if (Lines.Count == 0)
            throw new InvalidOperationException("VisualTeX formula metadata requires at least one line.");
        if (Numbered && !string.Equals(DisplayMode, "block", StringComparison.Ordinal))
            throw new InvalidOperationException("Only display formulas can use equation numbering.");
        if (!string.IsNullOrWhiteSpace(EquationTag)
            && (!string.Equals(DisplayMode, "block", StringComparison.Ordinal)
                || EquationTag!.Length > 256))
            throw new InvalidOperationException("Equation tags are supported only for display formulas and must not exceed 256 characters.");
        if (RenderWidthPx is <= 0 || double.IsNaN(RenderWidthPx ?? 1) || double.IsInfinity(RenderWidthPx ?? 1))
            throw new InvalidOperationException("VisualTeX renderWidthPx must be a positive finite number.");
        if (RenderHeightPx is <= 0 || double.IsNaN(RenderHeightPx ?? 1) || double.IsInfinity(RenderHeightPx ?? 1))
            throw new InvalidOperationException("VisualTeX renderHeightPx must be a positive finite number.");
        if (Baseline.HasValue
            && (double.IsNaN(Baseline.Value)
                || double.IsInfinity(Baseline.Value)
                || Baseline.Value < 0
                || (RenderHeightPx.HasValue && Baseline.Value > RenderHeightPx.Value)))
            throw new InvalidOperationException("VisualTeX baseline must be within the rendered formula height.");
        if (FontSizePt.HasValue
            && (FontSizePt.Value < FormulaFontSize.MinimumPt
                || FontSizePt.Value > FormulaFontSize.MaximumPt
                || double.IsNaN(FontSizePt.Value)
                || double.IsInfinity(FontSizePt.Value)))
            throw new InvalidOperationException("VisualTeX fontSizePt must be a supported finite point size.");
        if (RenderFontSizePt.HasValue
            && (RenderFontSizePt.Value < FormulaFontSize.MinimumPt
                || RenderFontSizePt.Value > FormulaFontSize.MaximumPt
                || double.IsNaN(RenderFontSizePt.Value)
                || double.IsInfinity(RenderFontSizePt.Value)))
            throw new InvalidOperationException("VisualTeX renderFontSizePt must be a supported finite point size.");
        if (FormulaLetterFont is not null && !IsSupportedFormulaLetterFont(FormulaLetterFont))
            throw new InvalidOperationException("VisualTeX formulaLetterFont is not supported.");
        if (FormulaChineseFont is not null && !IsSupportedFormulaChineseFont(FormulaChineseFont))
            throw new InvalidOperationException("VisualTeX formulaChineseFont is not supported.");
        FormulaFaces?.Validate();
        if (WordInlineOleWidthPt.HasValue != WordInlineOleHeightPt.HasValue)
            throw new InvalidOperationException(
                "VisualTeX Word inline OLE width and height must be stored together.");
        if (WordInlineOleWidthPt.HasValue
            && !string.Equals(DisplayMode, "inline", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "VisualTeX Word inline OLE dimensions are only valid for inline formulas.");
        if (WordInlineOleWidthPt.HasValue
            && (WordInlineOleWidthPt.Value <= 0
                || double.IsNaN(WordInlineOleWidthPt.Value)
                || double.IsInfinity(WordInlineOleWidthPt.Value)
                || WordInlineOleHeightPt!.Value <= 0
                || double.IsNaN(WordInlineOleHeightPt.Value)
                || double.IsInfinity(WordInlineOleHeightPt.Value)))
            throw new InvalidOperationException(
                "VisualTeX Word inline OLE dimensions must be positive finite values.");
        if (WordInlineOlePositionPt.HasValue
            && (!string.Equals(DisplayMode, "inline", StringComparison.OrdinalIgnoreCase)
                || WordInlineOlePositionPt.Value < -256
                || WordInlineOlePositionPt.Value > 256))
            throw new InvalidOperationException(
                "VisualTeX Word inline OLE position must be an inline whole-point offset between -256 and 256.");
        if (WordInlineSourceBottomWhitespacePt.HasValue
            && (!string.Equals(DisplayMode, "inline", StringComparison.OrdinalIgnoreCase)
                || WordInlineSourceBottomWhitespacePt.Value < 0
                || WordInlineSourceBottomWhitespacePt.Value > 256
                || double.IsNaN(WordInlineSourceBottomWhitespacePt.Value)
                || double.IsInfinity(WordInlineSourceBottomWhitespacePt.Value)))
            throw new InvalidOperationException(
                "VisualTeX source inline bottom whitespace must be a finite inline point value between 0 and 256.");
        if (WordInlinePreviewShiftPx.HasValue
            && (!string.Equals(DisplayMode, "inline", StringComparison.OrdinalIgnoreCase)
                || WordInlinePreviewShiftPx.Value < -2
                || WordInlinePreviewShiftPx.Value > 2
                || double.IsNaN(WordInlinePreviewShiftPx.Value)
                || double.IsInfinity(WordInlinePreviewShiftPx.Value)))
            throw new InvalidOperationException(
                "VisualTeX inline preview shift must be a finite pixel value between -2 and 2.");
        if (WordDisplayPreviewInkHeightRatio.HasValue
            != WordDisplayPreviewBottomWhitespaceRatio.HasValue)
            throw new InvalidOperationException(
                "VisualTeX display preview ink metrics must be stored together.");
        if (WordDisplayPreviewInkHeightRatio.HasValue
            && (!string.Equals(DisplayMode, "block", StringComparison.OrdinalIgnoreCase)
                || WordDisplayPreviewInkHeightRatio.Value <= 0
                || WordDisplayPreviewInkHeightRatio.Value > 1
                || WordDisplayPreviewBottomWhitespaceRatio!.Value < 0
                || WordDisplayPreviewBottomWhitespaceRatio.Value > 1
                || WordDisplayPreviewInkHeightRatio.Value
                    + WordDisplayPreviewBottomWhitespaceRatio.Value > 1.001
                || double.IsNaN(WordDisplayPreviewInkHeightRatio.Value)
                || double.IsInfinity(WordDisplayPreviewInkHeightRatio.Value)
                || double.IsNaN(WordDisplayPreviewBottomWhitespaceRatio.Value)
                || double.IsInfinity(WordDisplayPreviewBottomWhitespaceRatio.Value)))
            throw new InvalidOperationException(
                "VisualTeX display preview ink metrics must describe finite ratios inside the OLE frame.");
    }

    private static bool IsSupportedFormulaLetterFont(string value)
    {
        return value == "katex"
            || value == "times"
            || value == "cambria"
            || value == "stix"
            || value == "palatino"
            || value == "helvetica";
    }

    private static bool IsSupportedFormulaChineseFont(string value)
    {
        return value == "system"
            || value == "pingfang"
            || value == "songti"
            || value == "kaiti"
            || value == "heiti";
    }
}

public sealed class FormulaLine
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("latex")]
    public string Latex { get; set; } = string.Empty;
}
