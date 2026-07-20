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

    [JsonPropertyName("renderWidthPx")]
    public double? RenderWidthPx { get; set; }

    [JsonPropertyName("renderHeightPx")]
    public double? RenderHeightPx { get; set; }

    [JsonPropertyName("baseline")]
    public double? Baseline { get; set; }

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
    }
}

public sealed class FormulaLine
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("latex")]
    public string Latex { get; set; } = string.Empty;
}
