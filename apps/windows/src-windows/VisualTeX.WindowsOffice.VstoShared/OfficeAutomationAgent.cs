using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Forms;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.VstoShared;

/// <summary>
/// What an add-in can do for an agent (docs/office-mcp/DESIGN.md §4.4). Execute
/// runs on the Office UI thread and returns the command's data (serialized as JSON)
/// or throws <see cref="OfficeAutomationException"/>.
/// </summary>
internal interface IOfficeAutomationHost
{
    string App { get; }
    string ProductVersion { get; }
    IReadOnlyList<string> Capabilities { get; }
    // reads: run on the Office UI thread
    object? Execute(string operation, JsonElement arguments);
    // writes: the add-in's own commands, started from a background thread;
    // null when the operation is not a write
    Task<object?>? ExecuteWriteAsync(string operation, JsonElement arguments);
}

internal sealed class OfficeAutomationException : Exception
{
    internal OfficeAutomationException(string code, string message, object? details = null)
        : base(message)
    {
        Code = code;
        Details = details;
    }

    internal string Code { get; }
    internal object? Details { get; }
}

/// <summary>
/// Registers the add-in with the VisualTeX companion and runs the commands an agent
/// sends (DESIGN.md §4.3): a long poll on a background task, each command on the
/// Office UI thread, its result posted back. The add-in opens no port of its own.
/// </summary>
internal sealed class OfficeAutomationAgent : IDisposable
{
    private static readonly TimeSpan PollLimit = TimeSpan.FromSeconds(40);
    private static readonly TimeSpan RequestLimit = TimeSpan.FromSeconds(10);
    // A command that cannot get the UI thread in this time fails (a modal dialog is open).
    private static readonly TimeSpan UiStartLimit = TimeSpan.FromSeconds(20);
    private static readonly TimeSpan MaxBackoff = TimeSpan.FromSeconds(30);
    private static readonly JsonSerializerOptions ResultJson = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    private readonly VisualTeXSessionClient _client;
    private readonly IOfficeAutomationHost _host;
    private readonly Func<Func<object?>, Task<object?>> _invokeOnUi;
    private readonly Action<string>? _trace;
    private readonly CancellationTokenSource _lifetime = new();
    private string? _hostId;
    private string? _secret;
    private int _started;

    internal OfficeAutomationAgent(
        VisualTeXSessionClient client,
        IOfficeAutomationHost host,
        Func<Func<object?>, Task<object?>> invokeOnUi,
        Action<string>? trace = null)
    {
        _client = client;
        _host = host;
        _invokeOnUi = invokeOnUi;
        _trace = trace;
    }

    internal void Start()
    {
        if (Interlocked.Exchange(ref _started, 1) != 0) return;
        _ = Task.Run(() => RunAsync(_lifetime.Token));
    }

    private void Trace(string message)
    {
        try { _trace?.Invoke("automation " + message); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
    }

    private async Task RunAsync(CancellationToken cancellation)
    {
        var backoff = TimeSpan.FromSeconds(2);
        while (!cancellation.IsCancellationRequested)
        {
            try
            {
                if (_hostId is null)
                {
                    // Never start VisualTeX for this: the add-in's own startup and the
                    // agent's launcher do that. Only connect when it is listening.
                    if (!await _client.IsCompanionListeningAsync(cancellation).ConfigureAwait(false))
                    {
                        await Task.Delay(backoff, cancellation).ConfigureAwait(false);
                        backoff = Next(backoff);
                        continue;
                    }
                    await RegisterAsync(cancellation).ConfigureAwait(false);
                }

                var poll = await _client.SendAutomationAsync(
                    "GET",
                    $"/api/v1/automation/hosts/{Uri.EscapeDataString(_hostId!)}/commands?wait=25",
                    null,
                    _secret,
                    PollLimit,
                    cancellation).ConfigureAwait(false);
                if (poll.StatusCode == 404)
                {
                    // the companion restarted or dropped this host: register again
                    _hostId = null;
                    continue;
                }
                if (poll.StatusCode != 200)
                    throw new InvalidOperationException($"poll answered HTTP {poll.StatusCode}");
                backoff = TimeSpan.FromSeconds(2);
                using var document = JsonDocument.Parse(poll.Body);
                if (!document.RootElement.TryGetProperty("commands", out var commands)
                    || commands.ValueKind != JsonValueKind.Array)
                    continue;
                foreach (var command in commands.EnumerateArray())
                {
                    var copy = command.Clone();
                    _ = Task.Run(() => RunCommandAsync(copy, cancellation));
                }
            }
            catch (OperationCanceledException) when (cancellation.IsCancellationRequested)
            {
                break;
            }
            catch (Exception error)
            {
                Trace($"connection failed: {error.GetType().Name}: {error.Message}");
                _hostId = null;
                try { await Task.Delay(backoff, cancellation).ConfigureAwait(false); }
                catch (OperationCanceledException) { break; }
                backoff = Next(backoff);
            }
        }
    }

    private static TimeSpan Next(TimeSpan backoff) =>
        TimeSpan.FromTicks(Math.Min(MaxBackoff.Ticks, backoff.Ticks * 2));

    private async Task RegisterAsync(CancellationToken cancellation)
    {
        var registration = JsonSerializer.Serialize(new
        {
            app = _host.App,
            vendor = OfficeHostEnvironment.IsWps ? "wps" : "microsoft",
            productVersion = _host.ProductVersion,
            addinVersion = typeof(OfficeAutomationAgent).Assembly.GetName().Version?.ToString() ?? string.Empty,
            pid = Process.GetCurrentProcess().Id,
            capabilities = _host.Capabilities,
        });
        var response = await _client.SendAutomationAsync(
            "POST",
            "/api/v1/automation/hosts",
            registration,
            null,
            RequestLimit,
            cancellation).ConfigureAwait(false);
        if (response.StatusCode != 201)
            throw new InvalidOperationException($"registration answered HTTP {response.StatusCode}");
        using var document = JsonDocument.Parse(response.Body);
        _hostId = document.RootElement.GetProperty("hostId").GetString();
        _secret = document.RootElement.GetProperty("secret").GetString();
        Trace($"registered {_hostId}");
    }

    private async Task RunCommandAsync(JsonElement command, CancellationToken cancellation)
    {
        var hostId = _hostId;
        var secret = _secret;
        string id;
        string operation;
        try
        {
            id = command.GetProperty("id").GetString() ?? string.Empty;
            operation = command.GetProperty("operation").GetString() ?? string.Empty;
        }
        catch
        {
            return;
        }
        var arguments = command.TryGetProperty("arguments", out var value) ? value : default;
        var deadlineMs = command.TryGetProperty("deadlineMs", out var deadlineValue)
            && deadlineValue.TryGetInt64(out var parsed) ? parsed : 30_000;
        var deadline = DateTime.UtcNow + TimeSpan.FromMilliseconds(Math.Max(0, deadlineMs));

        string result;
        var started = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var abandoned = 0;
        try
        {
            var write = _host.ExecuteWriteAsync(operation, arguments);
            if (write is not null)
            {
                var written = await write.ConfigureAwait(false);
                result = JsonSerializer.Serialize(new { ok = true, data = written }, ResultJson);
                goto Post;
            }
            var work = _invokeOnUi(() =>
            {
                // never start late: the agent has already been told it failed
                if (Volatile.Read(ref abandoned) != 0 || DateTime.UtcNow > deadline)
                    throw new OfficeAutomationException("host_unresponsive", "The command was not started before its deadline.");
                started.TrySetResult(true);
                return _host.Execute(operation, arguments);
            });
            var first = await Task.WhenAny(work, started.Task, Task.Delay(UiStartLimit, cancellation)).ConfigureAwait(false);
            if (first != work && first != started.Task)
            {
                Interlocked.Exchange(ref abandoned, 1);
                if (!started.Task.IsCompleted)
                    throw new OfficeAutomationException(
                        "host_unresponsive",
                        OfficeHostEnvironment.IsWps
                            ? "WPS did not respond: a dialog may be open, or WPS is busy."
                            : $"{(_host.App == "word" ? "Word" : "PowerPoint")} did not respond: a dialog may be open, or it is busy.");
            }
            var data = await work.ConfigureAwait(false);
            result = JsonSerializer.Serialize(new { ok = true, data }, ResultJson);
        }
        catch (OfficeAutomationException error)
        {
            result = JsonSerializer.Serialize(
                new { ok = false, error = new { code = error.Code, message = error.Message, details = error.Details } },
                ResultJson);
        }
        catch (Exception error)
        {
            Trace($"{operation} failed: {error}");
            result = JsonSerializer.Serialize(
                new { ok = false, error = new { code = "internal", message = OfficeErrorText.Describe(error) } },
                ResultJson);
        }

    Post:
        if (hostId is null || cancellation.IsCancellationRequested) return;
        try
        {
            await _client.SendAutomationAsync(
                "POST",
                $"/api/v1/automation/hosts/{Uri.EscapeDataString(hostId)}/commands/{Uri.EscapeDataString(id)}/result",
                result,
                secret,
                RequestLimit,
                cancellation).ConfigureAwait(false);
        }
        catch (Exception error)
        {
            Trace($"result of {operation} not delivered: {error.Message}");
        }
    }

    public void Dispose()
    {
        var hostId = _hostId;
        var secret = _secret;
        _lifetime.Cancel();
        if (hostId is null) return;
        // best effort: the companion also drops a host that stops polling
        try
        {
            _ = _client.SendAutomationAsync(
                "DELETE",
                $"/api/v1/automation/hosts/{Uri.EscapeDataString(hostId)}",
                null,
                secret,
                TimeSpan.FromSeconds(2),
                CancellationToken.None);
        }
        catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
    }
}

/// <summary>Small helpers shared by the Word and PowerPoint automation hosts.</summary>
internal static class OfficeAutomationText
{
    /// The add-in's key for an open document: a short hash of its full name
    /// (its name while unsaved). Nothing is written into the document.
    internal static string DocumentKey(string identity)
    {
        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(identity ?? string.Empty));
        return Hex(hash, 6);
    }

    internal static string Hash8(string? value)
    {
        using var sha = SHA256.Create();
        return Hex(sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty)), 4);
    }

    private static string Hex(byte[] bytes, int count)
    {
        var builder = new StringBuilder(count * 2);
        for (var index = 0; index < count && index < bytes.Length; index++)
            builder.Append(bytes[index].ToString("x2"));
        return builder.ToString();
    }

    /// Document text for an agent: control characters Word uses for objects and
    /// fields removed, paragraph marks as newlines.
    internal static string Clean(string? text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        var builder = new StringBuilder(text!.Length);
        foreach (var character in text)
        {
            switch (character)
            {
                case '\r':
                case '\v':
                    builder.Append('\n');
                    break;
                case '\t':
                    builder.Append(' ');
                    break;
                case '\a':
                case '\u0001':
                case '\u0013':
                case '\u0014':
                case '\u0015':
                case '\uFFFC':
                    break;
                default:
                    if (!char.IsControl(character) || character == '\n') builder.Append(character);
                    break;
            }
        }
        return builder.ToString();
    }

    internal static string? String(JsonElement arguments, string name) =>
        arguments.ValueKind == JsonValueKind.Object
        && arguments.TryGetProperty(name, out var value)
        && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    internal static string RequiredString(JsonElement arguments, string name) =>
        String(arguments, name) is { Length: > 0 } value
            ? value
            : throw new OfficeAutomationException("invalid_arguments", $"{name} is required.");

    internal static int Integer(JsonElement arguments, string name, int fallback, int minimum, int maximum)
    {
        if (arguments.ValueKind != JsonValueKind.Object
            || !arguments.TryGetProperty(name, out var value)
            || !value.TryGetInt32(out var number))
            return fallback;
        return Math.Max(minimum, Math.Min(maximum, number));
    }

    internal static HashSet<string> Kinds(JsonElement arguments)
    {
        var kinds = new HashSet<string>(StringComparer.Ordinal);
        if (arguments.ValueKind == JsonValueKind.Object
            && arguments.TryGetProperty("kinds", out var value)
            && value.ValueKind == JsonValueKind.Array)
        {
            foreach (var kind in value.EnumerateArray())
                if (kind.ValueKind == JsonValueKind.String) kinds.Add(kind.GetString()!);
        }
        return kinds;
    }

    internal static bool? Bool(JsonElement arguments, string name) =>
        arguments.ValueKind == JsonValueKind.Object
        && arguments.TryGetProperty(name, out var value)
        && value.ValueKind is JsonValueKind.True or JsonValueKind.False
            ? value.GetBoolean()
            : null;

    internal static double? Number(JsonElement arguments, string name) =>
        arguments.ValueKind == JsonValueKind.Object
        && arguments.TryGetProperty(name, out var value)
        && value.ValueKind == JsonValueKind.Number
            ? value.GetDouble()
            : null;

    internal static JsonElement Element(JsonElement arguments, string name) =>
        arguments.ValueKind == JsonValueKind.Object && arguments.TryGetProperty(name, out var value)
            ? value
            : default;

    /// The formula's rows: "lines" (one row each) or "latex".
    internal static List<string>? Lines(JsonElement arguments)
    {
        var lines = new List<string>();
        var array = Element(arguments, "lines");
        if (array.ValueKind == JsonValueKind.Array)
        {
            foreach (var line in array.EnumerateArray())
                if (line.ValueKind == JsonValueKind.String) lines.Add(line.GetString()!);
        }
        else if (String(arguments, "latex") is { } latex)
        {
            lines.Add(latex);
        }
        lines = lines.Where(line => !string.IsNullOrWhiteSpace(line)).ToList();
        return lines.Count == 0 ? null : lines;
    }

    /// Agent format names → the add-ins' object modes.
    internal static string? ObjectMode(string? format) => format switch
    {
        null or "" => null,
        "visualtex_ole" => FormulaOleContract.NativeOleMode,
        "office_equation" => FormulaOleContract.WordOmmlMode,
        "mathtype" => FormulaOleContract.MathTypeOleMode,
        "picture" => FormulaOleContract.CrossPlatformPictureMode,
        _ => throw new OfficeAutomationException("invalid_arguments", $"Unknown format {format}: use visualtex_ole, office_equation, mathtype or picture."),
    };

    internal static OfficeAutomationException DocumentNotFound() =>
        new("document_not_found", "The document is no longer open, or it was saved under another name; call list_documents again.");

    internal static OfficeAutomationException UnknownOperation(string operation) =>
        new("unsupported_in_host", $"This add-in does not support {operation}.");

    /// Whether this process owns the foreground window (the user is working in it).
    internal static bool ProcessIsForeground()
    {
        try
        {
            var window = GetForegroundWindow();
            if (window == IntPtr.Zero) return false;
            GetWindowThreadProcessId(window, out var processId);
            return processId == (uint)Process.GetCurrentProcess().Id;
        }
        catch
        {
            return false;
        }
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}

/// <summary>
/// One agent write running the add-in's own command. While it runs, the command's
/// questions are answered "yes" (the agent asked for the operation), its messages
/// and errors are collected for the result instead of shown, and its status-bar
/// text becomes the result message.
/// </summary>
internal sealed class OfficeAutomationRun : IDisposable
{
    private static OfficeAutomationRun? _current;
    private readonly object _gate = new();
    private readonly List<string> _messages = new();
    private readonly List<string> _warnings = new();
    private readonly List<string> _errors = new();

    internal static OfficeAutomationRun? Current => Volatile.Read(ref _current);

    internal static OfficeAutomationRun Begin()
    {
        var run = new OfficeAutomationRun();
        if (Interlocked.CompareExchange(ref _current, run, null) is not null)
            throw new OfficeAutomationException("host_busy", "Another agent command is running in this Office application.");
        return run;
    }

    internal object? Result { get; set; }
    internal bool NumberDisplayFormulas { get; set; }
    internal string? BulkSource { get; set; }
    internal string BulkSourceFormat { get; set; } = "auto";
    internal string? BulkObjectMode { get; set; }

    internal void Status(string message)
    {
        if (string.IsNullOrWhiteSpace(message)) return;
        lock (_gate) _messages.Add(message.Trim());
    }

    internal void Warn(string message)
    {
        if (string.IsNullOrWhiteSpace(message)) return;
        lock (_gate) _warnings.Add(message.Trim());
    }

    internal void Fail(string message)
    {
        if (string.IsNullOrWhiteSpace(message)) return;
        lock (_gate) _errors.Add(message.Trim());
    }

    /// Stops a multi-step command after a failed step.
    internal void ThrowIfFailed()
    {
        lock (_gate)
        {
            if (_errors.Count > 0)
                throw new OfficeAutomationException("operation_failed", string.Join("\n", _errors.Distinct()));
        }
    }

    internal string? LastMessage
    {
        get { lock (_gate) return _messages.Count == 0 ? null : _messages[_messages.Count - 1]; }
    }

    /// The command's outcome: its errors as one failure, else its last message and warnings.
    internal Dictionary<string, object?> Finish(object? data = null)
    {
        lock (_gate)
        {
            if (_errors.Count > 0)
                throw new OfficeAutomationException(
                    "operation_failed",
                    string.Join("\n", _errors.Distinct()),
                    _warnings.Count == 0 ? null : new { warnings = _warnings.Distinct().ToList() });
            var result = new Dictionary<string, object?>
            {
                ["message"] = _messages.Count == 0 ? null : _messages[_messages.Count - 1],
            };
            if (_warnings.Count > 0) result["warnings"] = _warnings.Distinct().ToList();
            if (data is not null) result["result"] = data;
            return result;
        }
    }

    public void Dispose() => Interlocked.CompareExchange(ref _current, null, this);
}

/// <summary>
/// The add-ins' message boxes. During an agent command a question is answered with
/// its first button (the add-ins ask "Continue?" with Yes first), and the text is
/// collected: errors fail the command, warnings and notes come back with it.
/// </summary>
internal static class AutomationMessageBox
{
    private static DialogResult Answer(string text, MessageBoxButtons buttons, MessageBoxIcon icon)
    {
        var run = OfficeAutomationRun.Current!;
        switch (icon)
        {
            case MessageBoxIcon.Error:
                run.Fail(text);
                break;
            case MessageBoxIcon.Warning:
                run.Warn(text);
                break;
            default:
                if (buttons == MessageBoxButtons.OK) run.Status(text);
                break;
        }
        return buttons switch
        {
            MessageBoxButtons.YesNo or MessageBoxButtons.YesNoCancel => DialogResult.Yes,
            MessageBoxButtons.RetryCancel => DialogResult.Cancel,
            MessageBoxButtons.AbortRetryIgnore => DialogResult.Abort,
            _ => DialogResult.OK,
        };
    }

    internal static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon) =>
        OfficeAutomationRun.Current is null ? MessageBox.Show(text, caption, buttons, icon) : Answer(text, buttons, icon);

    internal static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
        MessageBoxDefaultButton defaultButton) =>
        OfficeAutomationRun.Current is null ? MessageBox.Show(text, caption, buttons, icon, defaultButton) : Answer(text, buttons, icon);

    internal static DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons,
        MessageBoxIcon icon) =>
        OfficeAutomationRun.Current is null ? MessageBox.Show(owner, text, caption, buttons, icon) : Answer(text, buttons, icon);

    internal static DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons,
        MessageBoxIcon icon, MessageBoxDefaultButton defaultButton) =>
        OfficeAutomationRun.Current is null
            ? MessageBox.Show(owner, text, caption, buttons, icon, defaultButton)
            : Answer(text, buttons, icon);
}
