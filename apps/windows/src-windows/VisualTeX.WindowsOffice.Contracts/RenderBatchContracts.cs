using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace VisualTeX.WindowsOffice.Contracts;

/// <summary>A formula render request with a caller-owned correlation identifier.</summary>
public sealed class RenderBatchItemRequest
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("request")]
    public CreateVstoSessionRequest Request { get; set; } = new();
}

/// <summary>One inline render result, retaining the caller's exact identifier.</summary>
public sealed class RenderBatchItemResult
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("session")]
    public OfficeSessionDocument Session { get; set; } = null!;
}

/// <summary>
/// An ephemeral batch's fully downloaded results. Items are returned in request
/// order. All exports are inline and remain usable after the server job is released.
/// Session status and error are left intact for the materialization owner.
/// </summary>
public sealed class RenderBatchResult
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("items")]
    public IReadOnlyList<RenderBatchItemResult> Items { get; set; } =
        Array.Empty<RenderBatchItemResult>();

    [JsonPropertyName("error")]
    public string? Error { get; set; }
}

internal sealed class CreateRenderBatchResponse
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;
}
