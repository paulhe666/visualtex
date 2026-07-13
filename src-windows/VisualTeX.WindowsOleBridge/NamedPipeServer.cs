using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOleBridge;

internal sealed class NamedPipeServer
{
    private const int MaxLineLength = 1024 * 1024;
    private static readonly TimeSpan HandshakeTimeout = TimeSpan.FromSeconds(5);
    private static readonly TimeSpan RequestTimeout = TimeSpan.FromSeconds(30);
    private readonly string _pipeName;
    private readonly string _token;
    private readonly IWindowsOfficeBackend _backend;
    private readonly FileLogger _logger;
    private readonly CancellationTokenSource _shutdown = new();

    public NamedPipeServer(
        string pipeName,
        string token,
        IWindowsOfficeBackend backend,
        FileLogger logger)
    {
        _pipeName = NormalizePipeName(pipeName);
        _token = token;
        _backend = backend;
        _logger = logger;
    }

    public async Task RunAsync(CancellationToken cancellationToken)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken,
            _shutdown.Token);
        _logger.Info($"Windows Office bridge listening on {_pipeName}.");
        while (!linked.IsCancellationRequested)
        {
            var pipe = CreatePipe();
            try
            {
                await pipe.WaitForConnectionAsync(linked.Token);
                _ = HandleClientAsync(pipe, linked.Token);
            }
            catch (OperationCanceledException)
            {
                pipe.Dispose();
                break;
            }
            catch (Exception error)
            {
                pipe.Dispose();
                _logger.Error("Named pipe accept failed.", error);
                await Task.Delay(250, linked.Token);
            }
        }
    }

    private async Task HandleClientAsync(
        NamedPipeServerStream pipe,
        CancellationToken cancellationToken)
    {
        await using var ownedPipe = pipe;
        using var reader = new StreamReader(
            pipe,
            new UTF8Encoding(false),
            detectEncodingFromByteOrderMarks: false,
            bufferSize: 16 * 1024,
            leaveOpen: true);
        using var writer = new StreamWriter(
            pipe,
            new UTF8Encoding(false),
            bufferSize: 16 * 1024,
            leaveOpen: true)
        {
            AutoFlush = true,
            NewLine = "\n",
        };
        try
        {
            using var handshakeTimeout = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken);
            handshakeTimeout.CancelAfter(HandshakeTimeout);
            var handshakeLine = await ReadLineAsync(reader, handshakeTimeout.Token);
            var handshake = JsonSerializer.Deserialize<OfficeBridgeRequest>(
                handshakeLine,
                JsonOptions.Default)
                ?? throw new InvalidDataException("Missing pipe handshake request.");
            if (!string.Equals(handshake.Method, "handshake", StringComparison.Ordinal)
                || !handshake.Params.TryGetProperty("token", out var tokenProperty)
                || tokenProperty.ValueKind != JsonValueKind.String
                || !ConstantTimeEquals(tokenProperty.GetString() ?? string.Empty, _token))
            {
                await WriteAsync(
                    writer,
                    OfficeBridgeResponse.Failure(
                        handshake.Id,
                        "unauthorized",
                        "Invalid Windows Office bridge token."));
                return;
            }
            await WriteAsync(writer, OfficeBridgeResponse.Success(handshake.Id, new { authenticated = true }));

            while (pipe.IsConnected && !cancellationToken.IsCancellationRequested)
            {
                string line;
                try { line = await ReadLineAsync(reader, cancellationToken); }
                catch (EndOfStreamException) { break; }
                var request = JsonSerializer.Deserialize<OfficeBridgeRequest>(line, JsonOptions.Default);
                if (request is null)
                {
                    await WriteAsync(writer, OfficeBridgeResponse.Failure(
                        string.Empty,
                        "invalid_json",
                        "Invalid Office bridge request."));
                    continue;
                }
                using var requestTimeout = CancellationTokenSource.CreateLinkedTokenSource(
                    cancellationToken);
                if (!string.Equals(request.Method, "shutdown", StringComparison.Ordinal))
                    requestTimeout.CancelAfter(RequestTimeout);

                var response = await _backend.HandleAsync(request, requestTimeout.Token);
                var timedOut = requestTimeout.IsCancellationRequested
                    && !cancellationToken.IsCancellationRequested
                    && !string.Equals(request.Method, "shutdown", StringComparison.Ordinal);
                if (timedOut)
                {
                    response = OfficeBridgeResponse.Failure(
                        request.Id,
                        "office_operation_timeout",
                        $"Office bridge method {request.Method} exceeded {RequestTimeout.TotalSeconds:0} seconds.",
                        retryable: true);
                }
                await WriteAsync(writer, response);
                if (timedOut)
                {
                    await writer.FlushAsync();
                    _logger.Error(
                        $"Office bridge method timed out and the sidecar will restart: {request.Method}");
                    Environment.Exit(124);
                }
                if (request.Method == "shutdown" && response.Ok)
                {
                    _shutdown.Cancel();
                    break;
                }
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception error)
        {
            _logger.Error("Named pipe client failed.", error);
        }
    }

    private NamedPipeServerStream CreatePipe()
    {
        using var identity = WindowsIdentity.GetCurrent();
        var sid = identity.User
            ?? throw new InvalidOperationException("Current Windows user has no SID.");
        var security = new PipeSecurity();
        security.SetOwner(sid);
        security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);
        security.AddAccessRule(new PipeAccessRule(
            sid,
            PipeAccessRights.ReadWrite | PipeAccessRights.CreateNewInstance,
            AccessControlType.Allow));
        return NamedPipeServerStreamAcl.Create(
            _pipeName,
            PipeDirection.InOut,
            8,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous | PipeOptions.WriteThrough,
            64 * 1024,
            64 * 1024,
            security);
    }

    private static async Task<string> ReadLineAsync(
        StreamReader reader,
        CancellationToken cancellationToken)
    {
        var line = await reader.ReadLineAsync(cancellationToken);
        if (line is null) throw new EndOfStreamException();
        if (line.Length > MaxLineLength)
            throw new InvalidDataException("Office bridge request exceeds 1 MiB.");
        return line;
    }

    private static Task WriteAsync(StreamWriter writer, OfficeBridgeResponse response) =>
        writer.WriteLineAsync(JsonSerializer.Serialize(response, JsonOptions.Default));

    internal static string NormalizePipeName(string value)
    {
        const string prefix = @"\\.\pipe\";
        var name = value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
            ? value[prefix.Length..]
            : value;
        if (string.IsNullOrWhiteSpace(name) || name.Contains('\\') || name.Contains('/'))
            throw new ArgumentException("Invalid Windows Office pipe name.", nameof(value));
        return name;
    }

    internal static bool ConstantTimeEquals(string left, string right)
    {
        var leftBytes = Encoding.UTF8.GetBytes(left);
        var rightBytes = Encoding.UTF8.GetBytes(right);
        if (leftBytes.Length != rightBytes.Length) return false;
        var difference = 0;
        for (var index = 0; index < leftBytes.Length; index++)
            difference |= leftBytes[index] ^ rightBytes[index];
        return difference == 0;
    }
}
