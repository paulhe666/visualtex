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
    // Acceptance runs can read the existing identity while keeping all runtime
    // data under a distinct artifact root. Production configuration never sets
    // these fields and retains its original install.json location/port checks.
    public string IdentityRoot { get; set; } = string.Empty;
    public int? IdentityPort { get; set; }
    public string InstallJsonPath => Path.Combine(
        string.IsNullOrWhiteSpace(IdentityRoot) ? AppDataRoot : IdentityRoot,
        "office", "install.json");
    public int InstallJsonPort => IdentityPort ?? CompanionPort;
}

public sealed class VisualTeXSessionClient : IDisposable
{
    internal const int ConverterTransportBatchSize = 256;

    internal static IReadOnlyList<IReadOnlyList<string>> ChunkConverterSessionIds(
        IReadOnlyList<string> sessionIds)
    {
        if (sessionIds is null) throw new ArgumentNullException(nameof(sessionIds));
        var chunks = new List<IReadOnlyList<string>>();
        for (var offset = 0; offset < sessionIds.Count; offset += ConverterTransportBatchSize)
        {
            var count = Math.Min(ConverterTransportBatchSize, sessionIds.Count - offset);
            var batch = new string[count];
            for (var index = 0; index < count; index++)
                batch[index] = sessionIds[offset + index];
            chunks.Add(batch);
        }
        return chunks;
    }

    private readonly CompanionConfiguration _configuration;
    private readonly VisualTeXCompanionException? _configurationException;
    private readonly HttpClient _http;
    private readonly HttpClient _renderBatchHttp;
    private readonly SemaphoreSlim _editorPrewarmGate = new(1, 1);
    private readonly SemaphoreSlim _converterPrewarmGate = new(1, 1);
    private string? _installToken;
    private static readonly object ClientLogLock = new();
    private string _lastServerCertificateThumbprint = string.Empty;
    private string _lastTlsPolicyErrors = string.Empty;
    private bool _hasValidatedServerCertificate;
    private bool _editorPrewarmed;
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
            // .NET Framework 4.7.2 supports this property. Keep the system TLS
            // default only on runtimes where the property is unavailable.
        }
        handler.ServerCertificateCustomValidationCallback = CaptureServerCertificate;
        _http = new HttpClient(handler)
        {
            BaseAddress = new Uri($"https://127.0.0.1:{_configuration.CompanionPort}"),
            Timeout = TimeSpan.FromSeconds(5),
        };
        _http.DefaultRequestHeaders.Accept.ParseAdd("application/json");
        // Only the batch result wait may outlive the existing five-second API
        // timeout. Share the exact TLS/proxy handler without taking ownership.
        // RenderBatchAsync supplies a caller-owned timeout and cancellation token.
        _renderBatchHttp = new HttpClient(handler, disposeHandler: false)
        {
            BaseAddress = _http.BaseAddress,
            Timeout = Timeout.InfiniteTimeSpan,
        };
        _renderBatchHttp.DefaultRequestHeaders.Accept.ParseAdd("application/json");
        AppendClientLog("client-created", null, null);
    }

    public CompanionHealthDiagnostic? LastHealthDiagnostic { get; private set; }

    public async Task EnsureHealthyAsync(CancellationToken cancellationToken)
    {
        // A cancelled caller (e.g. a Stop button) is not a missing companion:
        // a check cut short by cancellation must never start a VisualTeX
        // process or be reported as a companion failure.
        cancellationToken.ThrowIfCancellationRequested();
        if (_configurationException != null)
        {
            LastHealthDiagnostic = _configurationException.Diagnostic;
            throw new VisualTeXCompanionException(
                _configurationException.Diagnostic,
                _configurationException);
        }
        if (HasFreshHealth())
        {
            // The user may have closed VisualTeX since the last successful
            // request. A loopback connect is cheap; without it the cached
            // result skips the auto-start and the next request fails.
            if (await IsPortListeningAsync(_configuration.CompanionPort, cancellationToken)
                    .ConfigureAwait(false))
            {
                AppendClientLog("health-cache-hit", LastHealthDiagnostic, null);
                return;
            }
            InvalidateHealth();
            AppendClientLog("health-cache-stale", LastHealthDiagnostic, null);
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
            cancellationToken.ThrowIfCancellationRequested();
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
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            AppendClientLog("health-cancelled", diagnostic, null);
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

    public async Task PrewarmEditorAsync(CancellationToken cancellationToken)
    {
        if (_editorPrewarmed) return;
        await _editorPrewarmGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (_editorPrewarmed) return;
            EnsureAuthorizationHeader();
            using var response = await SendTrackedAsync(() => _http.PostAsync(
                "/api/v1/app/editor/prewarm",
                new StringContent("{}", Encoding.UTF8, "application/json"),
                cancellationToken)).ConfigureAwait(false);
            await EnsureSuccessAsync(response).ConfigureAwait(false);
            _editorPrewarmed = true;
        }
        finally
        {
            _editorPrewarmGate.Release();
        }
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
        if (sessionIds.Count == 0)
            throw new ArgumentOutOfRangeException(
                nameof(sessionIds),
                "VisualTeX batch conversion requires at least one Session.");
        var validated = new List<string>(sessionIds.Count);
        foreach (var sessionId in sessionIds)
        {
            if (!Guid.TryParse(sessionId, out _))
                throw new InvalidOperationException(
                    "VisualTeX Session id must be a UUID.");
            validated.Add(sessionId);
        }
        EnsureAuthorizationHeader();

        // Keep the transport batch small enough for older VisualTeX Companion
        // builds while allowing an arbitrary total number of Sessions. Companion
        // queues append requests in order, so 257 Sessions become 256 + 1 rather
        // than failing at the historical total-count ceiling.
        foreach (var batch in ChunkConverterSessionIds(validated))
        {
            var json = JsonSerializer.Serialize(
                new { sessionIds = batch },
                JsonOptions.Default);
            using var response = await SendTrackedAsync(() => _http.PostAsync(
                "/api/v1/app/converter/convert-batch",
                new StringContent(json, Encoding.UTF8, "application/json"),
                cancellationToken)).ConfigureAwait(false);
            await EnsureSuccessAsync(response).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Submits one ephemeral render operation and awaits its complete inline result.
    /// No durable sessions are created and no per-item polling is performed.
    /// The server job is released before this method returns; the returned managed
    /// export data remains available for subsequent local materialization.
    /// </summary>
    public Task<RenderBatchResult> RenderBatchAsync(
        IReadOnlyList<RenderBatchItemRequest> items,
        TimeSpan timeout,
        CancellationToken cancellationToken) =>
        SendRenderBatchAsync(
            items,
            timeout,
            cancellationToken,
            (request, isResultWait, token) =>
            {
                EnsureAuthorizationHeader();
                return SendTrackedAsync(() =>
                    (isResultWait ? _renderBatchHttp : _http).SendAsync(request, token));
            });

    // The injected sender keeps wire/lifetime tests independent of registry, TLS,
    // companion startup, Word, and the native renderer.
    internal static async Task<RenderBatchResult> SendRenderBatchAsync(
        IReadOnlyList<RenderBatchItemRequest> items,
        TimeSpan timeout,
        CancellationToken cancellationToken,
        Func<HttpRequestMessage, bool, CancellationToken, Task<HttpResponseMessage>> send)
    {
        if (items is null) throw new ArgumentNullException(nameof(items));
        if (send is null) throw new ArgumentNullException(nameof(send));
        if (items.Count == 0)
            throw new ArgumentException("A render batch requires at least one item.", nameof(items));
        if (timeout <= TimeSpan.Zero)
            throw new ArgumentOutOfRangeException(nameof(timeout), "A render batch timeout must be positive.");

        // Snapshot the identifiers before the first await so later caller edits
        // cannot change the correlation used to validate the response.
        var submitted = new RenderBatchItemRequest[items.Count];
        var requestedIds = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < items.Count; index++)
        {
            var item = items[index];
            if (item is null || string.IsNullOrWhiteSpace(item.Id) || item.Request is null)
                throw new ArgumentException("Every render item requires an id and request.", nameof(items));
            if (!requestedIds.Add(item.Id))
                throw new ArgumentException($"Duplicate render item id '{item.Id}'.", nameof(items));
            submitted[index] = new RenderBatchItemRequest { Id = item.Id, Request = item.Request };
        }
        cancellationToken.ThrowIfCancellationRequested();
        // The caller owns the id before submission, so even a lost POST response
        // can be followed by an idempotent DELETE of the accepted server job.
        var batchId = Guid.NewGuid().ToString("D");
        var json = JsonSerializer.Serialize(new { id = batchId, items = submitted }, JsonOptions.Default);
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        operation.CancelAfter(timeout);
        Exception? primaryFailure = null;
        RenderBatchResult? result = null;
        try
        {
            using (var request = new HttpRequestMessage(
                       HttpMethod.Post, "/api/v1/app/converter/render-batches")
                   {
                       Content = new StringContent(json, Encoding.UTF8, "application/json"),
                   })
            using (var response = await send(request, false, operation.Token).ConfigureAwait(false))
            {
                await EnsureSuccessAsync(response).ConfigureAwait(false);
                var created = await DeserializeAsync<CreateRenderBatchResponse>(response).ConfigureAwait(false);
                if (!string.Equals(created.Id, batchId, StringComparison.Ordinal))
                    throw new InvalidDataException("VisualTeX render batch submission did not echo the requested job id.");
            }

            using (var request = new HttpRequestMessage(
                       HttpMethod.Get,
                       $"/api/v1/app/converter/render-batches/{Uri.EscapeDataString(batchId)}/result"))
            using (var response = await send(request, true, operation.Token).ConfigureAwait(false))
            {
                await EnsureSuccessAsync(response).ConfigureAwait(false);
                result = await DeserializeAsync<RenderBatchResult>(response).ConfigureAwait(false);
            }
            operation.Token.ThrowIfCancellationRequested();
            if (!string.Equals(result.Id, batchId, StringComparison.Ordinal))
                throw new InvalidDataException("VisualTeX render batch response id did not match the submitted job.");
            if (!string.IsNullOrWhiteSpace(result.Error))
            {
                var error = new InvalidOperationException($"VisualTeX render batch failed: {result.Error}");
                error.Data["RenderBatchResult"] = result;
                throw error;
            }

            // Validate the complete one-to-one mapping before handing any formula
            // to Word. A server may return items in a different completion order.
            var returned = new Dictionary<string, RenderBatchItemResult>(StringComparer.Ordinal);
            foreach (var item in result.Items
                         ?? throw new InvalidDataException("VisualTeX render batch response had no items."))
            {
                if (item is null || string.IsNullOrWhiteSpace(item.Id) || item.Session is null)
                    throw new InvalidDataException("VisualTeX render batch returned an incomplete item.");
                if (!requestedIds.Contains(item.Id))
                    throw new InvalidDataException($"VisualTeX render batch returned unexpected item '{item.Id}'.");
                if (returned.ContainsKey(item.Id))
                    throw new InvalidDataException($"VisualTeX render batch returned duplicate item '{item.Id}'.");
                returned.Add(item.Id, item);
            }
            var ordered = new RenderBatchItemResult[submitted.Length];
            for (var index = 0; index < submitted.Length; index++)
            {
                if (!returned.TryGetValue(submitted[index].Id, out var item))
                    throw new InvalidDataException($"VisualTeX render batch omitted item '{submitted[index].Id}'.");
                ordered[index] = item;
            }
            result.Items = ordered;
            return result;
        }
        catch (OperationCanceledException error)
            when (!cancellationToken.IsCancellationRequested && operation.IsCancellationRequested)
        {
            primaryFailure = new TimeoutException("VisualTeX render batch timed out.", error);
            throw primaryFailure;
        }
        catch (Exception error)
        {
            primaryFailure = error;
            throw;
        }
        finally
        {
            // One idempotent DELETE uses the ordinary transport timeout and is
            // independent of caller cancellation. Inline exports are already in
            // managed memory and do not depend on the released server job.
            try
            {
                using var request = new HttpRequestMessage(
                    HttpMethod.Delete,
                    $"/api/v1/app/converter/render-batches/{Uri.EscapeDataString(batchId)}");
                using var response = await send(request, false, CancellationToken.None).ConfigureAwait(false);
                if (response.StatusCode != HttpStatusCode.NotFound)
                    await EnsureSuccessAsync(response).ConfigureAwait(false);
            }
            catch (Exception cleanupFailure)
            {
                if (primaryFailure != null)
                    throw new AggregateException(
                        $"VisualTeX render batch '{batchId}' failed and could not be released.",
                        primaryFailure, cleanupFailure);
                throw;
            }
        }
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

    public async Task DeleteSessionsBatchAsync(
        IReadOnlyCollection<string> sessionIds,
        CancellationToken cancellationToken)
    {
        if (sessionIds.Count == 0) return;
        EnsureAuthorizationHeader();
        var json = JsonSerializer.Serialize(
            new { sessionIds = sessionIds.ToArray() },
            JsonOptions.Default);
        using var response = await SendTrackedAsync(() => _http.PostAsync(
            "/api/v1/sessions/batch-delete",
            new StringContent(json, Encoding.UTF8, "application/json"),
            cancellationToken)).ConfigureAwait(false);
        await EnsureSuccessAsync(response).ConfigureAwait(false);
    }

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

    private static string CompanionLogRoot() => ResolveCompanionLogRoot(
        Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_OFFICE_COMPANION_PORT"),
        Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_OFFICE_APP_DATA_ROOT"));

    internal static string ResolveCompanionLogRoot(
        string? acceptancePortText,
        string? acceptanceAppDataRoot)
    {
        // Use the same explicit acceptance-mode selector as LoadConfiguration.
        // A configured test runtime must never write its diagnostics to production.
        if (int.TryParse(acceptancePortText, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var acceptancePort)
            && acceptancePort is > 0 and <= 65535)
        {
            var root = (acceptanceAppDataRoot ?? string.Empty).Trim().Trim('"');
            if (string.IsNullOrWhiteSpace(root))
                throw new InvalidDataException("Acceptance diagnostic logging requires APP_DATA_ROOT.");
            return Path.Combine(root, "office", "logs");
        }
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX", "office", "logs");
    }

    private static string VstoClientLogPath() =>
        Path.Combine(CompanionLogRoot(), "vsto-client.log");

    private static CompanionConfiguration LoadConfiguration()
    {
        var acceptancePortText = Environment.GetEnvironmentVariable(
            "VISUALTEX_ACCEPTANCE_OFFICE_COMPANION_PORT");
        if (int.TryParse(
                acceptancePortText,
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out var acceptancePort)
            && acceptancePort is > 0 and <= 65535)
        {
            var acceptanceProtocolText = Environment.GetEnvironmentVariable(
                "VISUALTEX_ACCEPTANCE_OFFICE_PROTOCOL_VERSION");
            _ = int.TryParse(
                acceptanceProtocolText,
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out var acceptanceProtocol);
            var acceptanceIdentityPortText = Environment.GetEnvironmentVariable(
                "VISUALTEX_ACCEPTANCE_OFFICE_IDENTITY_PORT");
            var acceptanceConfiguration = new CompanionConfiguration
            {
                ExecutablePath = (
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_ACCEPTANCE_OFFICE_EXECUTABLE_PATH")
                    ?? string.Empty).Trim().Trim('"'),
                AppDataRoot = (
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_ACCEPTANCE_OFFICE_APP_DATA_ROOT")
                    ?? string.Empty).Trim().Trim('"'),
                CertificatePath = (
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_ACCEPTANCE_OFFICE_CERTIFICATE_PATH")
                    ?? string.Empty).Trim().Trim('"'),
                CertificateThumbprint = (
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_ACCEPTANCE_OFFICE_CERTIFICATE_THUMBPRINT")
                    ?? string.Empty).Trim(),
                CompanionPort = acceptancePort,
                ProtocolVersion = acceptanceProtocol > 0
                    ? acceptanceProtocol
                    : 1,
                IdentityRoot = (
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_ACCEPTANCE_OFFICE_IDENTITY_ROOT")
                    ?? string.Empty).Trim().Trim('"'),
                IdentityPort = string.IsNullOrWhiteSpace(acceptanceIdentityPortText)
                    ? (int?)null
                    : int.TryParse(acceptanceIdentityPortText, NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out var identityPort)
                        ? identityPort
                        : 0,
            };
            ValidateAcceptanceIdentityConfiguration(acceptanceConfiguration);
            return acceptanceConfiguration;
        }

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

    internal static void ValidateAcceptanceIdentityConfiguration(CompanionConfiguration configuration)
    {
        if (string.IsNullOrWhiteSpace(configuration.IdentityRoot) && !configuration.IdentityPort.HasValue)
            return;
        try
        {
            if (string.IsNullOrWhiteSpace(configuration.IdentityRoot)
                || !configuration.IdentityPort.HasValue
                || configuration.IdentityPort.Value <= 0
                || configuration.IdentityPort.Value > 65535
                || configuration.IdentityPort.Value == configuration.CompanionPort)
                throw new InvalidDataException(
                    "Acceptance identity root and a valid source identity port must be supplied together; the runtime port must be distinct.");
            static bool IsAbsolutePath(string path) => Path.IsPathRooted(path)
                && Uri.TryCreate(path, UriKind.Absolute, out var uri) && uri.IsFile;
            if (!IsAbsolutePath(configuration.IdentityRoot)
                || !IsAbsolutePath(configuration.AppDataRoot))
                throw new InvalidDataException("Acceptance identity and runtime data roots must be absolute paths.");
            var identityRoot = Path.GetFullPath(configuration.IdentityRoot)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            var dataRoot = Path.GetFullPath(configuration.AppDataRoot)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            if (identityRoot.StartsWith(dataRoot, StringComparison.OrdinalIgnoreCase)
                || dataRoot.StartsWith(identityRoot, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Acceptance identity and runtime data roots must not overlap.");
        }
        catch (Exception error)
        {
            throw new VisualTeXCompanionException(new CompanionHealthDiagnostic
            {
                Stage = "configuration",
                FailureType = "AcceptanceIdentityInvalid",
                Message = error.Message,
                ExecutablePath = configuration.ExecutablePath,
                AppDataRoot = configuration.AppDataRoot,
                CertificatePath = configuration.CertificatePath,
                ExpectedCertificateThumbprint = configuration.CertificateThumbprint,
                CompanionPort = configuration.CompanionPort,
                ExpectedProtocolVersion = configuration.ProtocolVersion,
            }, error);
        }
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
                || actualPort != _configuration.InstallJsonPort)
                throw new InvalidDataException(
                    $"install.json port does not match the configured source identity port={_configuration.InstallJsonPort}.");
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
        _renderBatchHttp.DefaultRequestHeaders.Remove("X-VisualTeX-Install-Token");
        _renderBatchHttp.DefaultRequestHeaders.Add("X-VisualTeX-Install-Token", token);
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
        _editorPrewarmGate.Dispose();
        _converterPrewarmGate.Dispose();
        _renderBatchHttp.Dispose();
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
            WordInlineOlePositionPt = OriginalMetadata?.WordInlineOlePositionPt,
            WordInlineSourceBottomWhitespacePt = OriginalMetadata?.WordInlineSourceBottomWhitespacePt,
            WordInlinePreviewShiftPx = OriginalMetadata?.WordInlinePreviewShiftPx,
            CreatedWithVersion = OriginalMetadata?.CreatedWithVersion ?? "1.0.18",
            UpdatedWithVersion = "1.0.18",
            CreatedAt = OriginalMetadata?.CreatedAt ?? now,
            UpdatedAt = now,
        };
    }

    // Render input and persisted metadata must use the same source. Expose the
    // established Office-boundary cleanup so converter sessions cannot render
    // stale zero-width Word anchors that ToMetadata would later remove.
    public static string SanitizeFormulaBoundaryArtifacts(string? latex)
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
