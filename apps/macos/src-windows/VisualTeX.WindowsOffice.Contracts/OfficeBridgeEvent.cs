using System.Text.Json.Serialization;

namespace VisualTeX.WindowsOffice.Contracts;

public sealed class OfficeBridgeEvent
{
    [JsonPropertyName("protocolVersion")]
    public int ProtocolVersion { get; set; } = 1;

    [JsonPropertyName("event")]
    public string Event { get; set; } = string.Empty;

    [JsonPropertyName("payload")]
    public object? Payload { get; set; }

    [JsonPropertyName("cursor")]
    public long Cursor { get; set; }
}
