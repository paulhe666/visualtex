using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Security;
using System.Net.Sockets;
using System.Security.Authentication;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

namespace VisualTeX.WindowsOffice.Contracts;

public sealed class CompanionHealthDiagnostic
{
    [JsonPropertyName("stage")]
    public string Stage { get; set; } = "configuration";

    [JsonPropertyName("failureType")]
    public string FailureType { get; set; } = string.Empty;

    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;

    [JsonPropertyName("exception")]
    public string Exception { get; set; } = string.Empty;

    [JsonPropertyName("innerExceptions")]
    public List<string> InnerExceptions { get; set; } = new();

    [JsonPropertyName("executablePath")]
    public string ExecutablePath { get; set; } = string.Empty;

    [JsonPropertyName("appDataRoot")]
    public string AppDataRoot { get; set; } = string.Empty;

    [JsonPropertyName("certificatePath")]
    public string CertificatePath { get; set; } = string.Empty;

    [JsonPropertyName("expectedCertificateThumbprint")]
    public string ExpectedCertificateThumbprint { get; set; } = string.Empty;

    [JsonPropertyName("serverCertificateThumbprint")]
    public string ServerCertificateThumbprint { get; set; } = string.Empty;

    [JsonPropertyName("companionPort")]
    public int CompanionPort { get; set; }

    [JsonPropertyName("expectedProtocolVersion")]
    public int ExpectedProtocolVersion { get; set; }

    [JsonPropertyName("actualProtocolVersion")]
    public int? ActualProtocolVersion { get; set; }

    [JsonPropertyName("startedProcessId")]
    public int? StartedProcessId { get; set; }

    [JsonPropertyName("startedProcessExitCode")]
    public int? StartedProcessExitCode { get; set; }

    [JsonPropertyName("portListening")]
    public bool PortListening { get; set; }

    [JsonPropertyName("portOwnerProcessId")]
    public int? PortOwnerProcessId { get; set; }

    [JsonPropertyName("portOwnerProcessName")]
    public string PortOwnerProcessName { get; set; } = string.Empty;

    [JsonPropertyName("portOwnerProcessPath")]
    public string PortOwnerProcessPath { get; set; } = string.Empty;

    [JsonPropertyName("tlsPolicyErrors")]
    public string TlsPolicyErrors { get; set; } = string.Empty;

    [JsonPropertyName("healthResponse")]
    public string HealthResponse { get; set; } = string.Empty;

    [JsonPropertyName("installJsonPath")]
    public string InstallJsonPath { get; set; } = string.Empty;

    [JsonPropertyName("startupLogPath")]
    public string StartupLogPath { get; set; } = string.Empty;

    [JsonPropertyName("companionLogPath")]
    public string CompanionLogPath { get; set; } = string.Empty;

    [JsonPropertyName("vstoClientLogPath")]
    public string VstoClientLogPath { get; set; } = string.Empty;

    [JsonPropertyName("attemptedExecutablePaths")]
    public List<string> AttemptedExecutablePaths { get; set; } = new();

    public string ToJson() => JsonSerializer.Serialize(this, JsonOptions.Default);
}

public sealed class VisualTeXCompanionException : Exception
{
    public VisualTeXCompanionException(CompanionHealthDiagnostic diagnostic, Exception? inner = null)
        : base($"VisualTeX companion validation failed at stage '{diagnostic.Stage}': {diagnostic.Message}\n{diagnostic.ToJson()}", inner)
    {
        Diagnostic = diagnostic;
    }

    public CompanionHealthDiagnostic Diagnostic { get; }
}

internal sealed class CompanionConfiguration
{
    public string ExecutablePath { get; set; } = string.Empty;
    public string AppDataRoot { get; set; } = string.Empty;
    public string CertificatePath { get; set; } = string.Empty;
    public string CertificateThumbprint { get; set; } = string.Empty;
    public int CompanionPort { get; set; }
    public int ProtocolVersion { get; set; }
    public string InstallJsonPath => Path.Combine(AppDataRoot, "office", "install.json");
}

public sealed class VisualTeXSessionClient : IDisposable
{
    private readonly CompanionConfiguration _configuration;
    private readonly VisualTeXCompanionException? _configurationException;
    private readonly HttpClient _http;
    private readonly SemaphoreSlim _converterPrewarmGate = new(1, 1);
    private string? _installToken;
    private static readonly object ClientLogLock = new();
    private string _lastServerCertificateThumbprint = string.Empty;
    private string _lastTlsPolicyErrors = string.Empty;
    private bool _hasValidatedServerCertificate;
    private bool _converterPrewarmed;
    private bool _disposed;
    private long _lastHealthyUtcTicks;
    // Word can spend tens of seconds building or scanning a large document
    // after the add-in's startup prewarm. Keep that successful validation long
    // enough to cover the first formula edit; every tracked HTTP failure still
    // invalidates the cache immediately.
    private static readonly TimeSpan HealthCacheDuration = TimeSpan.FromMinutes(2);

    public VisualTeXSessionClient()
    {
        try
        {
            _configuration = LoadConfiguration();
        }
        catch (VisualTeXCompanionException error)
        {
            _configurationException = error;
            LastHealthDiagnostic = error.Diagnostic;
            _configuration = new CompanionConfiguration
            {
                CompanionPort = error.Diagnostic.CompanionPort > 0
                    ? error.Diagnostic.CompanionPort
                    : 43127,
                ProtocolVersion = error.Diagnostic.ExpectedProtocolVersion > 0
                    ? error.Diagnostic.ExpectedProtocolVersion
                    : 1,
                ExecutablePath = error.Diagnostic.ExecutablePath,
                AppDataRoot = error.Diagnostic.AppDataRoot,
                CertificatePath = error.Diagnostic.CertificatePath,
                CertificateThumbprint = error.Diagnostic.ExpectedCertificateThumbprint,
            };
        }
        var handler = new HttpClientHandler
        {
            UseProxy = false,
            Proxy = null,
        };
        try
        {
            handler.SslProtocols = SslProtocols.Tls12;
        }
        catch (PlatformNotSupportedException)
        {
            // .NET Framework 4.8 supports this property. Keep the system TLS
            // default only on runtimes where the property is unavailable.
        }
        handler.ServerCertificateCustomValidationCallback = CaptureServerCertificate;
        _http = new HttpClient(handler)
        {
            BaseAddress = new Uri($"https://127.0.0.1:{_configuration.CompanionPort}"),
            Timeout = TimeSpan.FromSeconds(5),
        };
        _http.DefaultRequestHeaders.Accept.ParseAdd("application/json");
        AppendClientLog("client-created", null, null);
    }

    public CompanionHealthDiagnostic? LastHealthDiagnostic { get; private set; }

    public async Task EnsureHealthyAsync(CancellationToken cancellationToken)
    {
        if (_configurationException != null)
        {
            LastHealthDiagnostic = _configurationException.Diagnostic;
            throw new VisualTeXCompanionException(
                _configurationException.Diagnostic,
                _configurationException);
        }
        if (HasFreshHealth())
        {
            AppendClientLog("health-cache-hit", LastHealthDiagnostic, null);
            return;
        }
        var diagnostic = CreateDiagnostic("configuration");
        LastHealthDiagnostic = diagnostic;
        AppendClientLog("health-start", diagnostic, null);
        try
        {
            ValidateConfiguration(diagnostic);
            if (await TryValidateRunningCompanionAsync(diagnostic, cancellationToken)
                    .ConfigureAwait(false))
            {
                EnsureAuthorizationHeader();
                diagnostic.Stage = "complete";
                diagnostic.FailureType = string.Empty;
                diagnostic.Message = "VisualTeX companion is healthy.";
                MarkHealthy();
                AppendClientLog("health-passed", diagnostic, null);
                return;
            }
            if (diagnostic.PortListening)
            {
                if (!string.IsNullOrWhiteSpace(diagnostic.PortOwnerProcessPath)
                    && !PathsEqual(diagnostic.PortOwnerProcessPath, _configuration.ExecutablePath))
                {
                    diagnostic.FailureType = "PortOccupiedByOtherProcess";
                    diagnostic.Message =
                        $"Port {_configuration.CompanionPort} is occupied by {diagnostic.PortOwnerProcessName} PID={diagnostic.PortOwnerProcessId?.ToString() ?? "unknown"} at '{diagnostic.PortOwnerProcessPath}', not by '{_configuration.ExecutablePath}'.";
                }
                else
                {
                    diagnostic.Message =
                        $"Port {_configuration.CompanionPort} is already listening, but the existing service failed stage '{diagnostic.Stage}'. Owner={diagnostic.PortOwnerProcessName} PID={diagnostic.PortOwnerProcessId?.ToString() ?? "unknown"} path='{diagnostic.PortOwnerProcessPath}'. {diagnostic.Message}";
                }
                throw new VisualTeXCompanionException(diagnostic);
            }

            diagnostic.Stage = "process-start";
            var started = StartVisualTeXCompanion(diagnostic);
            var deadline = DateTimeOffset.UtcNow + TimeSpan.FromSeconds(20);
            while (DateTimeOffset.UtcNow < deadline)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (started.HasExited)
                {
                    diagnostic.StartedProcessExitCode = started.ExitCode;
                    diagnostic.FailureType = "ProcessExited";
                    diagnostic.Message =
                        $"VisualTeX background process exited before the companion became healthy (exit code {started.ExitCode}).";
                    throw new VisualTeXCompanionException(diagnostic);
                }

                if (await TryValidateRunningCompanionAsync(diagnostic, cancellationToken)
                        .ConfigureAwait(false))
                {
                    EnsureAuthorizationHeader();
                    diagnostic.Stage = "complete";
                    diagnostic.FailureType = string.Empty;
                    diagnostic.Message = "VisualTeX companion started and passed all health checks.";
                    MarkHealthy();
                    AppendClientLog("health-passed-after-start", diagnostic, null);
                    return;
                }
                await Task.Delay(250, cancellationToken).ConfigureAwait(false);
            }

            if (string.IsNullOrWhiteSpace(diagnostic.FailureType))
                diagnostic.FailureType = "RuntimeValidationTimeout";
            diagnostic.Message =
                $"VisualTeX companion did not pass stage '{diagnostic.Stage}' within 20 seconds. {diagnostic.Message}".Trim();
            throw new VisualTeXCompanionException(diagnostic);
        }
        catch (VisualTeXCompanionException error)
        {
            InvalidateHealth();
            AppendClientLog("health-failed", error.Diagnostic, error);
            throw;
        }
        catch (Exception error)
        {
            InvalidateHealth();
            PopulateException(diagnostic, error);
            if (string.IsNullOrWhiteSpace(diagnostic.FailureType))
                diagnostic.FailureType = error.GetType().Name;
            if (string.IsNullOrWhiteSpace(diagnostic.Message))
                diagnostic.Message = error.Message;
            AppendClientLog("health-failed-unhandled", diagnostic, error);
            throw new VisualTeXCompanionException(diagnostic, error);
        }
    }

    public async Task<OfficeSessionDocument> CreateSessionAsync(
        CreateVstoSessionRequest request,
        CancellationToken cancellationToken)
    {
        EnsureAuthorizationHeader();
        var json = JsonSerializer.Serialize(request, JsonOptions.Default);
        using var response = await SendTrackedAsync(() => _http.PostAsync(
            "/api/v1/sessions",
            new StringContent(json, Encoding.UTF8, "application/json"),
            cancellationToken)).ConfigureAwait(false);
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
        using var response = await SendTrackedAsync(() => _http.PostAsync(
            $"/api/v1/app/sessions/{Uri.EscapeDataString(sessionId)}/open",
            new StringContent("{}", Encoding.UTF8, "application/json"),
            cancellationToken)).ConfigureAwait(false);
        await EnsureSuccessAsync(response).ConfigureAwait(false);
    }

    public async Task PrewarmConverterAsync(CancellationToken cancellationToken)
    {
        if (_converterPrewarmed) return;
        await _converterPrewarmGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (_converterPrewarmed) return;
            EnsureAuthorizationHeader();
            using var response = await SendTrackedAsync(() => _http.PostAsync(
                "/api/v1/app/converter/prewarm",
                new StringContent("{}", Encoding.UTF8, "application/json"),
                cancellationToken)).ConfigureAwait(false);
            await EnsureSuccessAsync(response).ConfigureAwait(false);
            // WebView creation returns before React and MathJax finish loading.
            // Pay this one-time warm-up cost at add-in startup so the first
            // actual redraw Session cannot race the custom Session event.
            await Task.Delay(500, cancellationToken).ConfigureAwait(false);
            _converterPrewarmed = true;
        }
        finally
        {
            _converterPrewarmGate.Release();
        }
    }

    public async Task OpenConverterAsync(
        string sessionId,
        CancellationToken cancellationToken)
    {
        if (!Guid.TryParse(sessionId, out _))
            throw new InvalidOperationException("VisualTeX Session id must be a UUID.");
        EnsureAuthorizationHeader();
        using var response = await SendTrackedAsync(() => _http.PostAsync(
            $"/api/v1/app/sessions/{Uri.EscapeDataString(sessionId)}/convert",
            new StringContent("{}", Encoding.UTF8, "application/json"),
            cancellationToken)).ConfigureAwait(false);
        await EnsureSuccessAsync(response).ConfigureAwait(false);
    }

    public async Task OpenConverterBatchAsync(
        IReadOnlyList<string> sessionIds,
        CancellationToken cancellationToken)
    {
        if (sessionIds is null)
            throw new ArgumentNullException(nameof(sessionIds));
        if (sessionIds.Count == 0 || sessionIds.Count > 256)
            throw new ArgumentOutOfRangeException(
                nameof(sessionIds),
                "VisualTeX batch conversion requires between 1 and 256 Sessions.");
        var validated = new List<string>(sessionIds.Count);
        foreach (var sessionId in sessionIds)
        {
            if (!Guid.TryParse(sessionId, out _))
                throw new InvalidOperationException(
                    "VisualTeX Session id must be a UUID.");
            validated.Add(sessionId);
        }
        EnsureAuthorizationHeader();
        var json = JsonSerializer.Serialize(
            new { sessionIds = validated },
            JsonOptions.Default);
        using var response = await SendTrackedAsync(() => _http.PostAsync(
            "/api/v1/app/converter/convert-batch",
            new StringContent(json, Encoding.UTF8, "application/json"),
            cancellationToken)).ConfigureAwait(false);
        await EnsureSuccessAsync(response).ConfigureAwait(false);
    }

    public async Task OpenBulkImportAsync(
        string sessionId,
        CancellationToken cancellationToken)
    {
        if (!Guid.TryParse(sessionId, out _))
            throw new InvalidOperationException("VisualTeX Session id must be a UUID.");
        EnsureAuthorizationHeader();
        using var response = await SendTrackedAsync(() => _http.PostAsync(
            $"/api/v1/app/sessions/{Uri.EscapeDataString(sessionId)}/bulk-import",
            new StringContent("{}", Encoding.UTF8, "application/json"),
            cancellationToken)).ConfigureAwait(false);
        await EnsureSuccessAsync(response).ConfigureAwait(false);
    }

    public async Task CloseEditorAsync(
        string sessionId,
        CancellationToken cancellationToken)
    {
        if (!Guid.TryParse(sessionId, out _))
            throw new InvalidOperationException("VisualTeX Session id must be a UUID.");
        EnsureAuthorizationHeader();
        using var response = await SendTrackedAsync(() => _http.PostAsync(
            $"/api/v1/app/sessions/{Uri.EscapeDataString(sessionId)}/close",
            new StringContent("{}", Encoding.UTF8, "application/json"),
            cancellationToken)).ConfigureAwait(false);
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
            // Formula commits are local Companion operations and normally finish
            // within a few hundred milliseconds. A 150 ms poll interval added a
            // visible delay larger than the actual Word update work; one frame
            // keeps the UI responsive without busy-waiting.
            await Task.Delay(16, cancellationToken).ConfigureAwait(false);
        }
        throw new TimeoutException("VisualTeX formula editing session timed out.");
    }

    public async Task<OfficeSessionDocument> GetSessionAsync(
        string sessionId,
        CancellationToken cancellationToken)
    {
        EnsureAuthorizationHeader();
        using var response = await SendTrackedAsync(() => _http.GetAsync(
            $"/api/v1/sessions/{Uri.EscapeDataString(sessionId)}",
            cancellationToken)).ConfigureAwait(false);
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
        using var response = await SendTrackedAsync(() =>
            _http.SendAsync(request, cancellationToken)).ConfigureAwait(false);
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
        var root = ControlledOfficeTempRoot();
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
        var root = ControlledOfficeTempRoot();
        Directory.CreateDirectory(root);
        var path = Path.Combine(root, $"{sessionId:D}.png");
        File.WriteAllBytes(path, bytes);
        return path;
    }

    private CompanionHealthDiagnostic CreateDiagnostic(string stage) => new()
    {
        Stage = stage,
        ExecutablePath = _configuration.ExecutablePath,
        AppDataRoot = _configuration.AppDataRoot,
        CertificatePath = _configuration.CertificatePath,
        ExpectedCertificateThumbprint = NormalizeThumbprint(_configuration.CertificateThumbprint),
        CompanionPort = _configuration.CompanionPort,
        ExpectedProtocolVersion = _configuration.ProtocolVersion,
        InstallJsonPath = _configuration.InstallJsonPath,
        StartupLogPath = Path.Combine(CompanionLogRoot(), "startup.log"),
        CompanionLogPath = Path.Combine(CompanionLogRoot(), "companion.log"),
        VstoClientLogPath = VstoClientLogPath(),
        AttemptedExecutablePaths = new List<string> { _configuration.ExecutablePath },
    };

    private static string ControlledOfficeTempRoot() => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "VisualTeX",
        "office",
        "temp");

    private static string CompanionLogRoot() => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "VisualTeX",
        "office",
        "logs");

    private static string VstoClientLogPath() =>
        Path.Combine(CompanionLogRoot(), "vsto-client.log");

    private static CompanionConfiguration LoadConfiguration()
    {
        const string registryPath = @"HKCU\Software\VisualTeX\OfficeIntegration";
        var diagnostic = new CompanionHealthDiagnostic
        {
            Stage = "configuration",
            FailureType = "ConfigurationMissing",
            StartupLogPath = Path.Combine(CompanionLogRoot(), "startup.log"),
            CompanionLogPath = Path.Combine(CompanionLogRoot(), "companion.log"),
            VstoClientLogPath = VstoClientLogPath(),
        };
        var values = QueryIntegrationRegistry(registryPath, diagnostic);
        var configuration = new CompanionConfiguration
        {
            ExecutablePath = ReadRegistryString(values, "ExecutablePath").Trim().Trim('"'),
            AppDataRoot = ReadRegistryString(values, "AppDataRoot").Trim().Trim('"'),
            CertificatePath = ReadRegistryString(values, "CertificatePath").Trim().Trim('"'),
            CertificateThumbprint = ReadRegistryString(values, "CertificateThumbprint"),
            CompanionPort = ReadRegistryInteger(values, "CompanionPort"),
            ProtocolVersion = ReadRegistryInteger(values, "ProtocolVersion"),
        };
        diagnostic.ExecutablePath = configuration.ExecutablePath;
        diagnostic.AppDataRoot = configuration.AppDataRoot;
        diagnostic.CertificatePath = configuration.CertificatePath;
        diagnostic.ExpectedCertificateThumbprint = NormalizeThumbprint(configuration.CertificateThumbprint);
        diagnostic.CompanionPort = configuration.CompanionPort;
        diagnostic.ExpectedProtocolVersion = configuration.ProtocolVersion;
        diagnostic.InstallJsonPath = configuration.InstallJsonPath;
        if (!string.IsNullOrWhiteSpace(configuration.ExecutablePath))
            diagnostic.AttemptedExecutablePaths.Add(configuration.ExecutablePath);

        var missing = new List<string>();
        if (string.IsNullOrWhiteSpace(configuration.ExecutablePath)) missing.Add("ExecutablePath");
        if (string.IsNullOrWhiteSpace(configuration.AppDataRoot)) missing.Add("AppDataRoot");
        if (string.IsNullOrWhiteSpace(configuration.CertificatePath)) missing.Add("CertificatePath");
        if (string.IsNullOrWhiteSpace(configuration.CertificateThumbprint)) missing.Add("CertificateThumbprint");
        if (configuration.CompanionPort <= 0 || configuration.CompanionPort > 65535) missing.Add("CompanionPort");
        if (configuration.ProtocolVersion <= 0) missing.Add("ProtocolVersion");
        if (missing.Count > 0)
        {
            diagnostic.Message =
                $"The shared Office configuration is incomplete. Missing or invalid values: {string.Join(", ", missing)}.";
            throw new VisualTeXCompanionException(diagnostic);
        }
        return configuration;
    }

    private static Dictionary<string, string> QueryIntegrationRegistry(
        string registryPath,
        CompanionHealthDiagnostic diagnostic)
    {
        try
        {
            using var process = Process.Start(new ProcessStartInfo(
                "reg.exe",
                $"query \"{registryPath}\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            });
            if (process == null)
                throw new InvalidOperationException("Process.Start returned null for reg.exe.");
            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();
            if (process.ExitCode != 0)
            {
                diagnostic.Message =
                    $"Unable to read {registryPath}. reg.exe exit code={process.ExitCode}; stdout={stdout}; stderr={stderr}";
                diagnostic.Exception = stderr;
                throw new VisualTeXCompanionException(diagnostic);
            }

            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var rawLine in stdout.Split(
                         new[] { '\r', '\n' },
                         StringSplitOptions.RemoveEmptyEntries))
            {
                var line = rawLine.Trim();
                var typeName = string.Empty;
                var typeIndex = -1;
                foreach (var candidate in new[] { "REG_SZ", "REG_EXPAND_SZ", "REG_DWORD" })
                {
                    var index = line.IndexOf(candidate, StringComparison.OrdinalIgnoreCase);
                    if (index < 0 || typeIndex >= 0 && index >= typeIndex) continue;
                    typeIndex = index;
                    typeName = candidate;
                }
                if (typeIndex <= 0) continue;
                var name = line.Substring(0, typeIndex).Trim();
                var value = line.Substring(typeIndex + typeName.Length).Trim();
                if (!string.IsNullOrWhiteSpace(name)) values[name] = value;
            }
            return values;
        }
        catch (VisualTeXCompanionException)
        {
            throw;
        }
        catch (Exception error)
        {
            diagnostic.Message = $"Unable to read {registryPath}: {error.Message}";
            PopulateException(diagnostic, error);
            throw new VisualTeXCompanionException(diagnostic, error);
        }
    }

    private static string ReadRegistryString(
        IReadOnlyDictionary<string, string> values,
        string name) => values.TryGetValue(name, out var value) ? value : string.Empty;

    private static int ReadRegistryInteger(
        IReadOnlyDictionary<string, string> values,
        string name)
    {
        if (!values.TryGetValue(name, out var value)) return 0;
        value = value.Trim();
        if (value.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(
                value.Substring(2),
                NumberStyles.HexNumber,
                CultureInfo.InvariantCulture,
                out var hexadecimal))
            return hexadecimal;
        return int.TryParse(
            value,
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out var decimalValue)
            ? decimalValue
            : 0;
    }

    private void ValidateConfiguration(CompanionHealthDiagnostic diagnostic)
    {
        diagnostic.Stage = "configuration";
        if (!File.Exists(_configuration.ExecutablePath))
        {
            diagnostic.FailureType = "ExecutableMissing";
            diagnostic.Message =
                $"The registered VisualTeX.exe does not exist: {_configuration.ExecutablePath}. Run repair with the exact -VisualTeXPath.";
            throw new VisualTeXCompanionException(diagnostic);
        }
        if (!Directory.Exists(_configuration.AppDataRoot))
        {
            diagnostic.FailureType = "AppDataRootMissing";
            diagnostic.Message = $"The registered AppDataRoot does not exist: {_configuration.AppDataRoot}";
            throw new VisualTeXCompanionException(diagnostic);
        }
        if (!File.Exists(_configuration.CertificatePath))
        {
            diagnostic.FailureType = "CertificateFileMissing";
            diagnostic.Message = $"The registered certificate file does not exist: {_configuration.CertificatePath}";
            throw new VisualTeXCompanionException(diagnostic);
        }

        using (var certificate = new X509Certificate2(_configuration.CertificatePath))
        {
            var fileThumbprint = NormalizeThumbprint(certificate.Thumbprint);
            if (!string.Equals(
                    fileThumbprint,
                    NormalizeThumbprint(_configuration.CertificateThumbprint),
                    StringComparison.OrdinalIgnoreCase))
            {
                diagnostic.FailureType = "CertificateRegistryMismatch";
                diagnostic.Message =
                    $"CertificatePath contains thumbprint {fileThumbprint}, but the registry expects {NormalizeThumbprint(_configuration.CertificateThumbprint)}.";
                throw new VisualTeXCompanionException(diagnostic);
            }
        }
        if (!IsCertificateInCurrentUserRoot(_configuration.CertificateThumbprint))
        {
            diagnostic.FailureType = "CertificateNotTrusted";
            diagnostic.Message =
                $"Certificate {_configuration.CertificateThumbprint} is not present in the current-user Root certificate store.";
            throw new VisualTeXCompanionException(diagnostic);
        }

        ValidateInstallJson(diagnostic);
    }

    private void ValidateInstallJson(CompanionHealthDiagnostic diagnostic)
    {
        diagnostic.Stage = "install-json";
        if (!File.Exists(_configuration.InstallJsonPath))
        {
            diagnostic.FailureType = "InstallJsonMissing";
            diagnostic.Message = $"VisualTeX install.json is missing: {_configuration.InstallJsonPath}";
            throw new VisualTeXCompanionException(diagnostic);
        }
        try
        {
            using var document = JsonDocument.Parse(
                File.ReadAllText(_configuration.InstallJsonPath, Encoding.UTF8));
            var root = document.RootElement;
            if (!root.TryGetProperty("installToken", out var token)
                || token.ValueKind != JsonValueKind.String
                || token.GetString()?.Length != 64)
                throw new InvalidDataException("installToken is missing or is not a 64-character token.");
            if (!root.TryGetProperty("port", out var port)
                || !port.TryGetInt32(out var actualPort)
                || actualPort != _configuration.CompanionPort)
                throw new InvalidDataException(
                    $"install.json port does not match registry CompanionPort={_configuration.CompanionPort}.");
            if (!root.TryGetProperty("protocolVersion", out var protocol)
                || !protocol.TryGetInt32(out var actualProtocol)
                || actualProtocol != _configuration.ProtocolVersion)
                throw new InvalidDataException(
                    $"install.json protocolVersion does not match registry ProtocolVersion={_configuration.ProtocolVersion}.");
        }
        catch (VisualTeXCompanionException)
        {
            throw;
        }
        catch (Exception error)
        {
            diagnostic.FailureType = "InstallJsonInvalid";
            diagnostic.Message = $"VisualTeX install.json is invalid: {error.Message}";
            PopulateException(diagnostic, error);
            throw new VisualTeXCompanionException(diagnostic, error);
        }
    }

    private bool CaptureServerCertificate(
        HttpRequestMessage _,
        X509Certificate2? certificate,
        X509Chain? __,
        SslPolicyErrors sslPolicyErrors)
    {
        _lastServerCertificateThumbprint = NormalizeThumbprint(certificate?.Thumbprint);
        _lastTlsPolicyErrors = sslPolicyErrors.ToString();
        var expectedThumbprint = NormalizeThumbprint(_configuration.CertificateThumbprint);
        var accepted = IsPinnedCertificateAccepted(
            certificate,
            expectedThumbprint,
            DateTime.UtcNow,
            out var certificateTimeValid,
            out var pinnedCertificateMatches);

        // This is a loopback-only service using a per-user certificate. The
        // certificate file, current-user Root trust, and registry thumbprint
        // are validated before the request. Office-hosted .NET Framework can
        // report chain/name policy errors that differ from PowerShell on the
        // same machine, so accept only the exact pinned, currently valid cert.
        var diagnostic = CreateDiagnostic("tls-certificate-callback");
        diagnostic.ServerCertificateThumbprint = _lastServerCertificateThumbprint;
        diagnostic.TlsPolicyErrors = _lastTlsPolicyErrors;
        diagnostic.FailureType = accepted ? string.Empty : "PinnedCertificateRejected";
        diagnostic.Message =
            $"TLS callback accepted={accepted}; policyErrors={sslPolicyErrors}; expected={expectedThumbprint}; actual={_lastServerCertificateThumbprint}; timeValid={certificateTimeValid}.";
        _hasValidatedServerCertificate = accepted;
        AppendClientLog("tls-certificate", diagnostic, null);
        return accepted;
    }

    private async Task<bool> TryValidateRunningCompanionAsync(
        CompanionHealthDiagnostic diagnostic,
        CancellationToken cancellationToken)
    {
        diagnostic.Stage = "port-listen";
        diagnostic.PortListening = await IsPortListeningAsync(
                _configuration.CompanionPort,
                cancellationToken)
            .ConfigureAwait(false);
        PopulatePortOwner(diagnostic);
        if (!diagnostic.PortListening)
        {
            diagnostic.FailureType = "PortNotListening";
            diagnostic.Message =
                $"No process is listening on 127.0.0.1:{_configuration.CompanionPort}.";
            return false;
        }

        diagnostic.Stage = "https-handshake";
        try
        {
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(TimeSpan.FromSeconds(3));
            using var response = await _http.GetAsync("/health", timeout.Token)
                .ConfigureAwait(false);
            diagnostic.ServerCertificateThumbprint = _lastServerCertificateThumbprint;
            diagnostic.TlsPolicyErrors = _lastTlsPolicyErrors;
            diagnostic.HealthResponse = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
            {
                diagnostic.FailureType = "HealthHttpStatus";
                diagnostic.Message =
                    $"The companion /health endpoint returned HTTP {(int)response.StatusCode}.";
                return false;
            }
        }
        catch (OperationCanceledException error) when (!cancellationToken.IsCancellationRequested)
        {
            diagnostic.FailureType = "HttpsTimeout";
            diagnostic.Message = "The HTTPS /health request timed out after 3 seconds.";
            diagnostic.ServerCertificateThumbprint = _lastServerCertificateThumbprint;
            diagnostic.TlsPolicyErrors = _lastTlsPolicyErrors;
            PopulateException(diagnostic, error);
            return false;
        }
        catch (HttpRequestException error)
        {
            diagnostic.ServerCertificateThumbprint = _lastServerCertificateThumbprint;
            diagnostic.TlsPolicyErrors = _lastTlsPolicyErrors;
            diagnostic.FailureType = ContainsException<AuthenticationException>(error)
                || !string.IsNullOrWhiteSpace(_lastTlsPolicyErrors)
                    ? "TlsValidationFailed"
                    : diagnostic.PortOwnerProcessId.HasValue
                        ? "PortOccupiedOrHttpsUnavailable"
                        : "HttpsRequestFailed";
            diagnostic.Message =
                $"HTTPS /health failed: {error.Message}. TLS={_lastTlsPolicyErrors}; port owner={diagnostic.PortOwnerProcessName} ({diagnostic.PortOwnerProcessId?.ToString() ?? "unknown"}) path='{diagnostic.PortOwnerProcessPath}'.";
            PopulateException(diagnostic, error);
            return false;
        }

        diagnostic.Stage = "certificate-match";
        if (!_hasValidatedServerCertificate
            || !string.Equals(
                NormalizeThumbprint(diagnostic.ServerCertificateThumbprint),
                NormalizeThumbprint(_configuration.CertificateThumbprint),
                StringComparison.OrdinalIgnoreCase))
        {
            diagnostic.FailureType = "ServerCertificateMismatch";
            diagnostic.Message =
                $"The HTTPS server certificate thumbprint {diagnostic.ServerCertificateThumbprint} does not match the registered certificate {_configuration.CertificateThumbprint}.";
            return false;
        }

        diagnostic.Stage = "health-json";
        try
        {
            using var document = JsonDocument.Parse(diagnostic.HealthResponse);
            var root = document.RootElement;
            if (!root.TryGetProperty("ok", out var ok) || ok.ValueKind != JsonValueKind.True)
            {
                diagnostic.FailureType = "HealthNotOk";
                diagnostic.Message = "The /health response did not contain ok=true.";
                return false;
            }
            if (!root.TryGetProperty("protocolVersion", out var protocol)
                || !protocol.TryGetInt32(out var actualProtocol))
            {
                diagnostic.FailureType = "HealthProtocolMissing";
                diagnostic.Message = "The /health response did not contain a numeric protocolVersion.";
                return false;
            }
            diagnostic.ActualProtocolVersion = actualProtocol;
            diagnostic.Stage = "protocol-version";
            if (actualProtocol != _configuration.ProtocolVersion)
            {
                diagnostic.FailureType = "ProtocolVersionMismatch";
                diagnostic.Message =
                    $"Companion protocolVersion={actualProtocol}; expected {_configuration.ProtocolVersion}.";
                return false;
            }
            diagnostic.FailureType = string.Empty;
            diagnostic.Message = "Port, HTTPS, certificate, health JSON and protocol version are valid.";
            return true;
        }
        catch (JsonException error)
        {
            diagnostic.FailureType = "HealthJsonInvalid";
            diagnostic.Message = $"The /health response is not valid JSON: {error.Message}";
            PopulateException(diagnostic, error);
            return false;
        }
    }

    private static async Task<bool> IsPortListeningAsync(
        int port,
        CancellationToken cancellationToken)
    {
        using var client = new TcpClient();
        try
        {
            var connect = client.ConnectAsync(IPAddress.Loopback, port);
            var completed = await Task.WhenAny(
                    connect,
                    Task.Delay(TimeSpan.FromMilliseconds(500), cancellationToken))
                .ConfigureAwait(false);
            if (completed != connect) return false;
            await connect.ConfigureAwait(false);
            return client.Connected;
        }
        catch (SocketException)
        {
            return false;
        }
    }

    private static void PopulatePortOwner(CompanionHealthDiagnostic diagnostic)
    {
        try
        {
            using var process = Process.Start(new ProcessStartInfo("netstat.exe", "-ano -p tcp")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            });
            if (process == null) return;
            var output = process.StandardOutput.ReadToEnd();
            process.WaitForExit(2000);
            foreach (var line in output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (line.IndexOf("LISTEN", StringComparison.OrdinalIgnoreCase) < 0) continue;
                var columns = line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                if (columns.Length < 4 || !columns[1].EndsWith(
                        $":{diagnostic.CompanionPort}",
                        StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!int.TryParse(columns[columns.Length - 1], out var processId)) continue;
                diagnostic.PortOwnerProcessId = processId;
                try
                {
                    using var owner = Process.GetProcessById(processId);
                    diagnostic.PortOwnerProcessName = owner.ProcessName;
                    try
                    {
                        diagnostic.PortOwnerProcessPath = owner.MainModule?.FileName ?? string.Empty;
                    }
                    catch
                    {
                        diagnostic.PortOwnerProcessPath = string.Empty;
                    }
                }
                catch
                {
                    diagnostic.PortOwnerProcessName = "<unavailable>";
                    diagnostic.PortOwnerProcessPath = string.Empty;
                }
                return;
            }
        }
        catch (Exception error)
        {
            diagnostic.InnerExceptions.Add($"Port owner lookup failed: {error}");
        }
    }

    private static bool PathsEqual(string left, string right)
    {
        try
        {
            return string.Equals(
                Path.GetFullPath(left),
                Path.GetFullPath(right),
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
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
        if (!File.Exists(_configuration.ExecutablePath))
            throw new FileNotFoundException(
                $"The registered VisualTeX.exe does not exist: {_configuration.ExecutablePath}",
                _configuration.ExecutablePath);
        Process.Start(new ProcessStartInfo(_configuration.ExecutablePath)
        {
            UseShellExecute = true,
        });
    }

    private Process StartVisualTeXCompanion(CompanionHealthDiagnostic diagnostic)
    {
        diagnostic.Stage = "process-start";
        var process = Process.Start(new ProcessStartInfo(
            _configuration.ExecutablePath,
            "--office-background")
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            WorkingDirectory = Path.GetDirectoryName(_configuration.ExecutablePath) ?? string.Empty,
        });
        if (process == null)
        {
            diagnostic.FailureType = "ProcessStartReturnedNull";
            diagnostic.Message =
                $"Process.Start returned null for {_configuration.ExecutablePath}.";
            throw new VisualTeXCompanionException(diagnostic);
        }
        diagnostic.StartedProcessId = process.Id;
        diagnostic.Message =
            $"Started VisualTeX background process PID={process.Id} from {_configuration.ExecutablePath}.";
        return process;
    }

    private static bool IsCertificateInCurrentUserRoot(string thumbprint)
    {
        var expected = NormalizeThumbprint(thumbprint);
        using var store = new X509Store(StoreName.Root, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadOnly);
        foreach (var certificate in store.Certificates)
        {
            if (string.Equals(
                    NormalizeThumbprint(certificate.Thumbprint),
                    expected,
                    StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    internal static bool IsPinnedCertificateAccepted(
        X509Certificate2? certificate,
        string? expectedThumbprint,
        DateTime utcNow,
        out bool certificateTimeValid,
        out bool pinnedCertificateMatches)
    {
        var normalizedExpected = NormalizeThumbprint(expectedThumbprint);
        var normalizedActual = NormalizeThumbprint(certificate?.Thumbprint);
        certificateTimeValid = certificate is not null
            && utcNow >= certificate.NotBefore.ToUniversalTime()
            && utcNow <= certificate.NotAfter.ToUniversalTime();
        pinnedCertificateMatches = certificate is not null
            && !string.IsNullOrWhiteSpace(normalizedExpected)
            && string.Equals(
                normalizedActual,
                normalizedExpected,
                StringComparison.OrdinalIgnoreCase);
        return pinnedCertificateMatches && certificateTimeValid;
    }

    private static string NormalizeThumbprint(string? value) =>
        (value ?? string.Empty).Replace(" ", string.Empty).ToUpperInvariant();

    private void AppendClientLog(
        string eventName,
        CompanionHealthDiagnostic? diagnostic,
        Exception? error)
    {
        try
        {
            var logPath = VstoClientLogPath();
            Directory.CreateDirectory(Path.GetDirectoryName(logPath) ?? CompanionLogRoot());
            using var process = Process.GetCurrentProcess();
            string processPath;
            try
            {
                processPath = process.MainModule?.FileName ?? string.Empty;
            }
            catch
            {
                processPath = string.Empty;
            }
            var assembly = typeof(VisualTeXSessionClient).Assembly;
            var builder = new StringBuilder();
            builder.Append(DateTimeOffset.Now.ToString("o", CultureInfo.InvariantCulture));
            builder.Append(" event=").Append(eventName);
            builder.Append(" process=").Append(process.ProcessName);
            builder.Append(" pid=").Append(process.Id.ToString(CultureInfo.InvariantCulture));
            builder.Append(" processPath=\"").Append(processPath).Append('"');
            builder.Append(" clr=\"").Append(Environment.Version).Append('"');
            builder.Append(" os=\"").Append(Environment.OSVersion).Append('"');
            builder.Append(" assembly=\"").Append(assembly.Location).Append('"');
            builder.Append(" assemblyVersion=\"").Append(assembly.GetName().Version).Append('"');
            builder.Append(" executable=\"").Append(_configuration.ExecutablePath).Append('"');
            builder.Append(" port=").Append(_configuration.CompanionPort.ToString(CultureInfo.InvariantCulture));
            if (diagnostic is not null)
                builder.Append(" diagnostic=").Append(diagnostic.ToJson());
            if (error is not null)
                builder.Append(" exception=").Append(JsonSerializer.Serialize(error.ToString()));
            builder.AppendLine();
            lock (ClientLogLock)
            {
                File.AppendAllText(logPath, builder.ToString(), new UTF8Encoding(false));
            }
        }
        catch
        {
            // Diagnostics must never block Office operations.
        }
    }

    private static bool ContainsException<T>(Exception error) where T : Exception
    {
        for (Exception? current = error; current != null; current = current.InnerException)
        {
            if (current is T) return true;
        }
        return false;
    }

    private static void PopulateException(
        CompanionHealthDiagnostic diagnostic,
        Exception error)
    {
        diagnostic.Exception = error.ToString();
        diagnostic.InnerExceptions.Clear();
        for (var current = error.InnerException; current != null; current = current.InnerException)
            diagnostic.InnerExceptions.Add(current.ToString());
    }

    private bool HasFreshHealth()
    {
        var lastHealthyTicks = Interlocked.Read(ref _lastHealthyUtcTicks);
        if (lastHealthyTicks <= 0) return false;
        var elapsedTicks = DateTime.UtcNow.Ticks - lastHealthyTicks;
        return elapsedTicks >= 0 && elapsedTicks <= HealthCacheDuration.Ticks;
    }

    private void MarkHealthy() =>
        Interlocked.Exchange(ref _lastHealthyUtcTicks, DateTime.UtcNow.Ticks);

    private void InvalidateHealth() =>
        Interlocked.Exchange(ref _lastHealthyUtcTicks, 0);

    private async Task<HttpResponseMessage> SendTrackedAsync(
        Func<Task<HttpResponseMessage>> operation)
    {
        try
        {
            var response = await operation().ConfigureAwait(false);
            if (response.IsSuccessStatusCode) MarkHealthy();
            else InvalidateHealth();
            return response;
        }
        catch
        {
            InvalidateHealth();
            throw;
        }
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

    private string ReadInstallToken()
    {
        try
        {
            using var document = JsonDocument.Parse(
                File.ReadAllText(_configuration.InstallJsonPath, Encoding.UTF8));
            if (document.RootElement.TryGetProperty("installToken", out var token)
                && token.ValueKind == JsonValueKind.String
                && token.GetString()?.Length == 64)
                return token.GetString()!;
            throw new InvalidDataException("installToken is missing or invalid.");
        }
        catch (Exception error)
        {
            var diagnostic = CreateDiagnostic("install-json");
            diagnostic.FailureType = "InstallJsonInvalid";
            diagnostic.Message =
                $"Unable to read VisualTeX install token from {_configuration.InstallJsonPath}: {error.Message}";
            PopulateException(diagnostic, error);
            LastHealthDiagnostic = diagnostic;
            throw new VisualTeXCompanionException(diagnostic, error);
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _converterPrewarmGate.Dispose();
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

    [JsonPropertyName("mathTypeNumberPosition")]
    public string MathTypeNumberPosition { get; set; } = "right";

    [JsonPropertyName("fontSizePt")]
    public double FontSizePt { get; set; } = FormulaFontSize.DefaultPt;

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

    [JsonPropertyName("mathTypeNumberPosition")]
    public string MathTypeNumberPosition { get; set; } = "right";

    [JsonPropertyName("fontSizePt")]
    public double FontSizePt { get; set; } = FormulaFontSize.DefaultPt;

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
        var sanitizedLines = Lines.ConvertAll(line => new FormulaLine
        {
            Id = line.Id,
            Latex = SanitizeFormulaBoundaryArtifacts(line.Latex),
        });
        return new FormulaMetadata
        {
            FormulaId = FormulaId,
            Title = Title,
            Latex = string.Join("\n", sanitizedLines.ConvertAll(line => line.Latex)),
            Lines = sanitizedLines,
            CodeFormat = CodeFormat,
            DisplayMode = DisplayMode,
            Numbered = Numbered,
            EquationTag = string.Equals(DisplayMode, "block", StringComparison.Ordinal)
                ? OriginalMetadata?.EquationTag
                : null,
            RenderWidthPx = ExportResult?.Width > 0 ? ExportResult.Width : OriginalMetadata?.RenderWidthPx,
            RenderHeightPx = ExportResult?.Height > 0 ? ExportResult.Height : OriginalMetadata?.RenderHeightPx,
            Baseline = ExportResult?.Baseline ?? OriginalMetadata?.Baseline,
            FontSizePt = FormulaFontSize.Normalize(FontSizePt),
            RenderFontSizePt = ExportResult is not null
                ? FormulaFontSize.Normalize(FontSizePt)
                : OriginalMetadata?.RenderFontSizePt ?? FormulaFontSize.Normalize(FontSizePt),
            FormulaLetterFont = ExportResult?.FormulaLetterFont ?? OriginalMetadata?.FormulaLetterFont,
            FormulaChineseFont = ExportResult?.FormulaChineseFont ?? OriginalMetadata?.FormulaChineseFont,
            WordInlineOleWidthPt = OriginalMetadata?.WordInlineOleWidthPt,
            WordInlineOleHeightPt = OriginalMetadata?.WordInlineOleHeightPt,
            CreatedWithVersion = OriginalMetadata?.CreatedWithVersion ?? "1.0.18",
            UpdatedWithVersion = "1.0.18",
            CreatedAt = OriginalMetadata?.CreatedAt ?? now,
            UpdatedAt = now,
        };
    }

    private static string SanitizeFormulaBoundaryArtifacts(string? latex)
    {
        if (string.IsNullOrEmpty(latex)) return string.Empty;
        static bool IsBoundaryArtifact(char character) =>
            character is '\u200B' or '\u200C' or '\u2060' or '\uFEFF';

        var value = latex!;
        for (var index = 0; index < value.Length; index++)
        {
            if (!IsBoundaryArtifact(value[index])) continue;
            var runEnd = index + 1;
            while (runEnd < value.Length && IsBoundaryArtifact(value[runEnd]))
                runEnd++;
            var left = value.Substring(0, index).Trim();
            var right = value.Substring(runEnd).Trim();
            if (left.Length > 0 && string.Equals(left, right, StringComparison.Ordinal))
                return left;
            index = runEnd - 1;
        }

        return new string(value.Where(character => !IsBoundaryArtifact(character)).ToArray());
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

    [JsonPropertyName("formulaLetterFont")]
    public string? FormulaLetterFont { get; set; }

    [JsonPropertyName("formulaChineseFont")]
    public string? FormulaChineseFont { get; set; }
}
