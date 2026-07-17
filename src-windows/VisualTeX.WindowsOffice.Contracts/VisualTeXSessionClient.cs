using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

namespace VisualTeX.WindowsOffice.Contracts;

public sealed class VisualTeXSessionClient : IDisposable
{
    private static readonly Uri CompanionOrigin = new("https://127.0.0.1:43127");
    private readonly HttpClient _http;
    private string? _installToken;
    private bool _disposed;

    public VisualTeXSessionClient()
    {
        _http = new HttpClient
        {
            BaseAddress = CompanionOrigin,
            Timeout = TimeSpan.FromSeconds(15),
        };
        _http.DefaultRequestHeaders.Accept.ParseAdd("application/json");
    }

    public async Task EnsureHealthyAsync(CancellationToken cancellationToken)
    {
        if (!await TryReadHealthyAsync(cancellationToken).ConfigureAwait(false))
        {
            StartVisualTeXCompanion();
            var deadline = DateTimeOffset.UtcNow + TimeSpan.FromSeconds(20);
            while (DateTimeOffset.UtcNow < deadline)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await Task.Delay(200, cancellationToken).ConfigureAwait(false);
                if (await TryReadHealthyAsync(cancellationToken).ConfigureAwait(false))
                    break;
            }
            if (!await TryReadHealthyAsync(cancellationToken).ConfigureAwait(false))
                throw new TimeoutException(
                    "VisualTeX local companion did not become healthy within 20 seconds.");
        }
        EnsureAuthorizationHeader();
    }

    public async Task<OfficeSessionDocument> CreateSessionAsync(
        CreateVstoSessionRequest request,
        CancellationToken cancellationToken)
    {
        EnsureAuthorizationHeader();
        var json = JsonSerializer.Serialize(request, JsonOptions.Default);
        using var response = await _http.PostAsync(
            "/api/v1/sessions",
            new StringContent(json, Encoding.UTF8, "application/json"),
            cancellationToken).ConfigureAwait(false);
        await EnsureSuccessAsync(response).ConfigureAwait(false);
        return await DeserializeAsync<OfficeSessionDocument>(response).ConfigureAwait(false);
    }

    public async Task OpenEditorAsync(
        string sessionId,
        CancellationToken cancellationToken)
    {
        if (!Guid.TryParse(sessionId, out _))
            throw new InvalidOperationException("VisualTeX Session id must be a UUID.");
        EnsureAuthorizationHeader();
        using var response = await _http.PostAsync(
            $"/api/v1/app/sessions/{Uri.EscapeDataString(sessionId)}/open",
            new StringContent("{}", Encoding.UTF8, "application/json"),
            cancellationToken).ConfigureAwait(false);
        await EnsureSuccessAsync(response).ConfigureAwait(false);
    }

    public async Task CloseEditorAsync(
        string sessionId,
        CancellationToken cancellationToken)
    {
        if (!Guid.TryParse(sessionId, out _))
            throw new InvalidOperationException("VisualTeX Session id must be a UUID.");
        EnsureAuthorizationHeader();
        using var response = await _http.PostAsync(
            $"/api/v1/app/sessions/{Uri.EscapeDataString(sessionId)}/close",
            new StringContent("{}", Encoding.UTF8, "application/json"),
            cancellationToken).ConfigureAwait(false);
        await EnsureSuccessAsync(response).ConfigureAwait(false);
    }

    public async Task<OfficeSessionDocument> WaitForCommitAsync(
        string sessionId,
        TimeSpan timeout,
        CancellationToken cancellationToken)
    {
        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var session = await GetSessionAsync(sessionId, cancellationToken).ConfigureAwait(false);
            switch (session.Status)
            {
                case "committing":
                case "completed":
                case "cancelled":
                case "failed":
                    return session;
            }
            await Task.Delay(150, cancellationToken).ConfigureAwait(false);
        }
        throw new TimeoutException("VisualTeX formula editing session timed out.");
    }

    public async Task<OfficeSessionDocument> GetSessionAsync(
        string sessionId,
        CancellationToken cancellationToken)
    {
        EnsureAuthorizationHeader();
        using var response = await _http.GetAsync(
            $"/api/v1/sessions/{Uri.EscapeDataString(sessionId)}",
            cancellationToken).ConfigureAwait(false);
        await EnsureSuccessAsync(response).ConfigureAwait(false);
        return await DeserializeAsync<OfficeSessionDocument>(response).ConfigureAwait(false);
    }

    public Task<OfficeSessionDocument> CompleteAsync(
        string sessionId,
        CancellationToken cancellationToken) =>
        PatchAsync(sessionId, new { status = "completed", error = (string?)null }, cancellationToken);

    public Task<OfficeSessionDocument> FailAsync(
        string sessionId,
        string error,
        CancellationToken cancellationToken) =>
        PatchAsync(sessionId, new { status = "failed", error }, cancellationToken);

    public async Task<OfficeSessionDocument> PatchAsync(
        string sessionId,
        object update,
        CancellationToken cancellationToken)
    {
        EnsureAuthorizationHeader();
        var json = JsonSerializer.Serialize(update, JsonOptions.Default);
        using var request = new HttpRequestMessage(
            new HttpMethod("PATCH"),
            $"/api/v1/sessions/{Uri.EscapeDataString(sessionId)}")
        {
            Content = new StringContent(json, Encoding.UTF8, "application/json"),
        };
        using var response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false);
        await EnsureSuccessAsync(response).ConfigureAwait(false);
        return await DeserializeAsync<OfficeSessionDocument>(response).ConfigureAwait(false);
    }

    public string MaterializeSvg(OfficeSessionDocument session)
    {
        if (!Guid.TryParse(session.Id, out var sessionId))
            throw new InvalidOperationException("VisualTeX Session id must be a UUID.");
        var export = session.ExportResult
            ?? throw new InvalidOperationException("VisualTeX Session has no SVG export.");
        var svg = export.Svg;
        if (string.IsNullOrWhiteSpace(svg) && !string.IsNullOrWhiteSpace(export.SvgBase64))
        {
            var encoded = export.SvgBase64!;
            var comma = encoded.IndexOf(',');
            if (comma >= 0) encoded = encoded.Substring(comma + 1);
            var bytes = Convert.FromBase64String(encoded);
            if (bytes.Length == 0 || bytes.Length > 16 * 1024 * 1024)
                throw new InvalidDataException("VisualTeX Session SVG export size is invalid.");
            svg = new UTF8Encoding(false, true).GetString(bytes);
        }
        if (string.IsNullOrWhiteSpace(svg))
            throw new InvalidOperationException("VisualTeX Session has no SVG export.");
        var normalized = svg!.Trim();
        if (!normalized.StartsWith("<svg", StringComparison.OrdinalIgnoreCase)
            || normalized.IndexOf("</svg>", StringComparison.OrdinalIgnoreCase) < 0)
            throw new InvalidDataException("VisualTeX Session SVG export is invalid.");
        if (normalized.IndexOf("<foreignObject", StringComparison.OrdinalIgnoreCase) >= 0
            || normalized.IndexOf("<image", StringComparison.OrdinalIgnoreCase) >= 0
            || normalized.IndexOf("<script", StringComparison.OrdinalIgnoreCase) >= 0)
            throw new InvalidDataException("VisualTeX Session SVG export contains forbidden content.");
        var root = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(root);
        var path = Path.Combine(root, $"{sessionId:D}.svg");
        File.WriteAllText(path, normalized, new UTF8Encoding(false, true));
        return path;
    }

    public string MaterializePng(OfficeSessionDocument session)
    {
        if (!Guid.TryParse(session.Id, out var sessionId))
            throw new InvalidOperationException("VisualTeX Session id must be a UUID.");
        var data = session.ExportResult?.PngBase64
            ?? throw new InvalidOperationException("VisualTeX Session has no PNG export.");
        var comma = data.IndexOf(',');
        if (comma >= 0) data = data.Substring(comma + 1);
        var bytes = Convert.FromBase64String(data);
        if (bytes.Length < 8
            || bytes[0] != 137
            || bytes[1] != 80
            || bytes[2] != 78
            || bytes[3] != 71)
            throw new InvalidDataException("VisualTeX Session PNG export is invalid.");
        var root = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(root);
        var path = Path.Combine(root, $"{sessionId:D}.png");
        File.WriteAllBytes(path, bytes);
        return path;
    }

    private async Task<bool> TryReadHealthyAsync(CancellationToken cancellationToken)
    {
        try
        {
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken);
            timeout.CancelAfter(TimeSpan.FromSeconds(2));
            using var response = await _http.GetAsync("/health", timeout.Token)
                .ConfigureAwait(false);
            if (!response.IsSuccessStatusCode) return false;
            var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            using var document = JsonDocument.Parse(json);
            return document.RootElement.TryGetProperty("ok", out var ok)
                && ok.ValueKind == JsonValueKind.True;
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            return false;
        }
        catch (HttpRequestException)
        {
            return false;
        }
    }

    private void EnsureAuthorizationHeader()
    {
        var token = ReadInstallToken();
        if (string.Equals(_installToken, token, StringComparison.Ordinal)) return;
        _http.DefaultRequestHeaders.Remove("X-VisualTeX-Install-Token");
        _http.DefaultRequestHeaders.Add("X-VisualTeX-Install-Token", token);
        _installToken = token;
    }

    public void OpenDesktop()
    {
        var executable = FindVisualTeXExecutable();
        Process.Start(new ProcessStartInfo(executable)
        {
            UseShellExecute = true,
        });
    }

    private static void StartVisualTeXCompanion()
    {
        var executable = FindVisualTeXExecutable();
        Process.Start(new ProcessStartInfo(executable, "--office-background")
        {
            UseShellExecute = true,
            WindowStyle = ProcessWindowStyle.Hidden,
        });
    }

    private static string FindVisualTeXExecutable()
    {
        var candidates = new[]
        {
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Programs",
                "VisualTeX",
                "VisualTeX.exe"),
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "VisualTeX",
                "VisualTeX.exe"),
        };
        foreach (var executable in candidates)
        {
            if (File.Exists(executable)) return executable;
        }
        throw new FileNotFoundException(
            "The installed VisualTeX desktop executable could not be found.");
    }

    private static async Task EnsureSuccessAsync(HttpResponseMessage response)
    {
        if (response.IsSuccessStatusCode) return;
        var detail = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        throw new HttpRequestException(
            $"VisualTeX Session request failed ({(int)response.StatusCode}): {detail}");
    }

    private static async Task<T> DeserializeAsync<T>(HttpResponseMessage response)
    {
        var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        return JsonSerializer.Deserialize<T>(json, JsonOptions.Default)
            ?? throw new InvalidDataException("VisualTeX Session response was empty.");
    }

    private static string ReadInstallToken()
    {
        var candidates = new[]
        {
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "com.visualtex.studio",
                "office",
                "install.json"),
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "com.visualtex.studio",
                "office",
                "install.json"),
        };
        foreach (var path in candidates)
        {
            if (!File.Exists(path)) continue;
            using var document = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
            if (document.RootElement.TryGetProperty("installToken", out var token)
                && token.ValueKind == JsonValueKind.String
                && token.GetString()?.Length == 64)
                return token.GetString()!;
        }
        throw new FileNotFoundException(
            "VisualTeX install token was not found. Start VisualTeX before using the Office add-in.");
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _http.Dispose();
    }
}

public sealed class CreateVstoSessionRequest
{
    [JsonPropertyName("mode")]
    public string Mode { get; set; } = "create";

    [JsonPropertyName("host")]
    public string Host { get; set; } = string.Empty;

    [JsonPropertyName("formulaId")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FormulaId { get; set; }

    [JsonPropertyName("sourceDocumentId")]
    public string? SourceDocumentId { get; set; }

    [JsonPropertyName("sourceObjectId")]
    public string? SourceObjectId { get; set; }

    [JsonPropertyName("title")]
    public string Title { get; set; } = "Office Formula";

    [JsonPropertyName("lines")]
    public List<FormulaLine> Lines { get; set; } = new();

    [JsonPropertyName("activeLineId")]
    public string? ActiveLineId { get; set; }

    [JsonPropertyName("codeFormat")]
    public string CodeFormat { get; set; } = "latex";

    [JsonPropertyName("displayMode")]
    public string DisplayMode { get; set; } = "block";

    [JsonPropertyName("objectMode")]
    public string ObjectMode { get; set; } = "nativeOle";

    [JsonPropertyName("numbered")]
    public bool Numbered { get; set; }

    [JsonPropertyName("originalMetadata")]
    public FormulaMetadata? OriginalMetadata { get; set; }

    [JsonPropertyName("autoCommitOnClose")]
    public bool AutoCommitOnClose { get; set; } = true;
}

public sealed class OfficeSessionDocument
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("mode")]
    public string Mode { get; set; } = string.Empty;

    [JsonPropertyName("host")]
    public string Host { get; set; } = string.Empty;

    [JsonPropertyName("formulaId")]
    public string FormulaId { get; set; } = string.Empty;

    [JsonPropertyName("sourceDocumentId")]
    public string? SourceDocumentId { get; set; }

    [JsonPropertyName("sourceObjectId")]
    public string? SourceObjectId { get; set; }

    [JsonPropertyName("title")]
    public string Title { get; set; } = string.Empty;

    [JsonPropertyName("lines")]
    public List<FormulaLine> Lines { get; set; } = new();

    [JsonPropertyName("codeFormat")]
    public string CodeFormat { get; set; } = string.Empty;

    [JsonPropertyName("displayMode")]
    public string DisplayMode { get; set; } = "block";

    [JsonPropertyName("objectMode")]
    public string ObjectMode { get; set; } = "crossPlatformPicture";

    [JsonPropertyName("numbered")]
    public bool Numbered { get; set; }

    [JsonPropertyName("status")]
    public string Status { get; set; } = string.Empty;

    [JsonPropertyName("dirty")]
    public bool Dirty { get; set; }

    [JsonPropertyName("explicitCancel")]
    public bool ExplicitCancel { get; set; }

    [JsonPropertyName("error")]
    public string? Error { get; set; }

    [JsonPropertyName("originalMetadata")]
    public FormulaMetadata? OriginalMetadata { get; set; }

    [JsonPropertyName("exportResult")]
    public OfficeExportDocument? ExportResult { get; set; }

    public FormulaMetadata ToMetadata()
    {
        var now = DateTimeOffset.UtcNow.ToString("O");
        return new FormulaMetadata
        {
            FormulaId = FormulaId,
            Title = Title,
            Latex = string.Join("\n", Lines.ConvertAll(line => line.Latex)),
            Lines = Lines,
            CodeFormat = CodeFormat,
            DisplayMode = DisplayMode,
            Numbered = Numbered,
            RenderWidthPx = ExportResult?.Width > 0 ? ExportResult.Width : OriginalMetadata?.RenderWidthPx,
            RenderHeightPx = ExportResult?.Height > 0 ? ExportResult.Height : OriginalMetadata?.RenderHeightPx,
            Baseline = ExportResult?.Baseline ?? OriginalMetadata?.Baseline,
            CreatedWithVersion = OriginalMetadata?.CreatedWithVersion ?? "1.0.18",
            UpdatedWithVersion = "1.0.18",
            CreatedAt = OriginalMetadata?.CreatedAt ?? now,
            UpdatedAt = now,
        };
    }
}

public sealed class OfficeExportDocument
{
    [JsonPropertyName("svg")]
    public string? Svg { get; set; }

    [JsonPropertyName("svgBase64")]
    public string? SvgBase64 { get; set; }

    [JsonPropertyName("mathMl")]
    public string? MathMl { get; set; }

    [JsonPropertyName("pngBase64")]
    public string? PngBase64 { get; set; }

    [JsonPropertyName("width")]
    public float Width { get; set; }

    [JsonPropertyName("height")]
    public float Height { get; set; }

    [JsonPropertyName("baseline")]
    public float? Baseline { get; set; }
}
