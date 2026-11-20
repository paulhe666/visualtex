using System.Text.Json.Serialization;

namespace VisualTeX.WindowsOffice.Contracts;

public sealed class SessionInfo
{
    [JsonPropertyName("sessionId")]
    public string SessionId { get; set; } = string.Empty;

    [JsonPropertyName("formulaId")]
    public string FormulaId { get; set; } = string.Empty;

    [JsonPropertyName("host")]
    public string Host { get; set; } = string.Empty;

    [JsonPropertyName("mode")]
    public string Mode { get; set; } = string.Empty;

    [JsonPropertyName("imagePath")]
    public string ImagePath { get; set; } = string.Empty;

    [JsonPropertyName("sourceDocumentId")]
    public string? SourceDocumentId { get; set; }

    [JsonPropertyName("sourceObjectId")]
    public string? SourceObjectId { get; set; }

    [JsonPropertyName("width")]
    public float Width { get; set; }

    [JsonPropertyName("height")]
    public float Height { get; set; }

    [JsonPropertyName("baseline")]
    public float? Baseline { get; set; }

    [JsonPropertyName("metadata")]
    public FormulaMetadata Metadata { get; set; } = new();
}
