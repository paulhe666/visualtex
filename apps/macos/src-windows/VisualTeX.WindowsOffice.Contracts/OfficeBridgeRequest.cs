using System.Text.Json;
using System.Text.Json.Serialization;

namespace VisualTeX.WindowsOffice.Contracts;

public sealed class OfficeBridgeRequest
{
    [JsonPropertyName("protocolVersion")]
    public int ProtocolVersion { get; set; } = 1;

    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("method")]
    public string Method { get; set; } = string.Empty;

    [JsonPropertyName("params")]
    public JsonElement Params { get; set; }
}
