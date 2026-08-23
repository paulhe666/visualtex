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

    [JsonPropertyName("wordInlineOleWidthPt")]
    public double? WordInlineOleWidthPt { get; set; }

    [JsonPropertyName("wordInlineOleHeightPt")]
    public double? WordInlineOleHeightPt { get; set; }

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
