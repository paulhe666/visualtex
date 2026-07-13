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
    }
}

public sealed class FormulaLine
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("latex")]
    public string Latex { get; set; } = string.Empty;
}
