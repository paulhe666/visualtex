using System.Text.Json.Serialization;

namespace VisualTeX.WindowsOffice.Contracts;

public sealed class OfficeBridgeResponse
{
    [JsonPropertyName("protocolVersion")]
    public int ProtocolVersion { get; set; } = 1;

    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("ok")]
    public bool Ok { get; set; }

    [JsonPropertyName("result")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Result { get; set; }

    [JsonPropertyName("error")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public OfficeBridgeError? Error { get; set; }

    public static OfficeBridgeResponse Success(string id, object? result = null) =>
        new() { Id = id, Ok = true, Result = result };

    public static OfficeBridgeResponse Failure(
        string id,
        string code,
        string message,
        bool retryable = false,
        object? details = null) =>
        new()
        {
            Id = id,
            Ok = false,
            Error = new OfficeBridgeError
            {
                Code = code,
                Message = message,
                Retryable = retryable,
                Details = details,
            },
        };
}

public sealed class OfficeBridgeError
{
    [JsonPropertyName("code")]
    public string Code { get; set; } = string.Empty;

    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;

    [JsonPropertyName("retryable")]
    public bool Retryable { get; set; }

    [JsonPropertyName("details")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Details { get; set; }
}
