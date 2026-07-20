using System.Collections.Concurrent;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOleBridge;

internal sealed class WindowsOfficeBackend : IWindowsOfficeBackend, IDisposable
{
    private readonly OfficeStaDispatcher _dispatcher;
    private readonly WordOleService _word = new();
    private readonly PowerPointOleService _powerPoint = new();
    private readonly TempPathGuard _tempPathGuard;
    private readonly FileLogger _logger;
    private readonly ConcurrentQueue<OfficeBridgeEvent> _events = new();
    private readonly OfficeDoubleClickHook _doubleClickHook;
    private readonly FormulaDoubleClickDeduplicator _doubleClickDeduplicator = new();
    private long _eventCursor;
    private bool _disposed;

    public WindowsOfficeBackend(string tempRoot, FileLogger logger)
    {
        _logger = logger;
        _tempPathGuard = new TempPathGuard(tempRoot);
        _dispatcher = new OfficeStaDispatcher(logger);
        _doubleClickHook = new OfficeDoubleClickHook(OnOfficeDoubleClick, logger);
        _doubleClickHook.Start();
    }

    public async Task<OfficeBridgeResponse> HandleAsync(
        OfficeBridgeRequest request,
        CancellationToken cancellationToken)
    {
        if (request.ProtocolVersion != 1)
            return OfficeBridgeResponse.Failure(
                request.Id,
                "unsupported_protocol",
                $"Unsupported protocol version {request.ProtocolVersion}.");
        if (string.IsNullOrWhiteSpace(request.Id) || string.IsNullOrWhiteSpace(request.Method))
            return OfficeBridgeResponse.Failure(
                request.Id,
                "invalid_request",
                "Request id and method are required.");

        try
        {
            object? result = request.Method switch
            {
                "health" => new
                {
                    ok = true,
                    processId = Environment.ProcessId,
                    staThreadId = _dispatcher.ManagedThreadId,
                },
                "office.detect" => await OnStaAsync(
                    () => (object)new
                    {
                        wordRunning = RunningOfficeLocator.IsRunning("Word.Application"),
                        powerPointRunning = RunningOfficeLocator.IsRunning("PowerPoint.Application"),
                    },
                    cancellationToken),
                "word.getSelection" => await OnStaAsync(() => _word.GetSelection(), cancellationToken),
                "word.insertInlineFormula" => await WithSessionAsync(
                    request,
                    session => _word.InsertInlineFormula(session),
                    cancellationToken),
                "word.insertDisplayFormula" => await WithSessionAsync(
                    request,
                    session => _word.InsertDisplayFormula(session),
                    cancellationToken),
                "word.replaceFormula" => await WithSessionAsync(
                    request,
                    session => _word.ReplaceFormula(session),
                    cancellationToken),
                "word.updateEquationNumbers" => await OnStaAsync(
                    () => (object)new { updated = _word.UpdateEquationNumbers() },
                    cancellationToken),
                "powerpoint.getSelection" => await OnStaAsync(
                    () => _powerPoint.GetSelection(),
                    cancellationToken),
                "powerpoint.insertFormula" => await WithSessionAsync(
                    request,
                    session => _powerPoint.InsertFormula(session),
                    cancellationToken),
                "powerpoint.replaceFormula" => await WithSessionAsync(
                    request,
                    session => _powerPoint.ReplaceFormula(session),
                    cancellationToken),
                "powerpoint.markFormula" => await OnStaAsync(
                    () => _powerPoint.MarkFormula(RequiredString(request.Params, "formulaId")),
                    cancellationToken),
                "powerpoint.deleteFormula" => await OnStaAsync(
                    () =>
                    {
                        _powerPoint.DeleteFormula(RequiredString(request.Params, "formulaId"));
                        return new { deleted = true };
                    },
                    cancellationToken),
                "office.openWord" => OpenWord(),
                "office.openPowerPoint" => OpenPowerPoint(),
                "events.after" => new
                {
                    events = GetEventsAfter(OptionalLong(request.Params, "cursor")),
                },
                "shutdown" => new { shuttingDown = true },
                _ => throw new NotSupportedException($"Unsupported Office bridge method: {request.Method}"),
            };
            return OfficeBridgeResponse.Success(request.Id, result);
        }
        catch (OperationCanceledException)
        {
            return OfficeBridgeResponse.Failure(
                request.Id,
                "cancelled",
                "The Office operation was cancelled.",
                retryable: true);
        }
        catch (Exception error)
        {
            _logger.Error($"Office bridge method failed: {request.Method}", error);
            return OfficeBridgeResponse.Failure(
                request.Id,
                ErrorCode(error),
                error.Message,
                IsRetryable(error));
        }
    }

    public IReadOnlyList<OfficeBridgeEvent> GetEventsAfter(long cursor) =>
        _events.Where(item => item.Cursor > cursor).OrderBy(item => item.Cursor).Take(100).ToArray();

    private Task<T> OnStaAsync<T>(Func<T> operation, CancellationToken cancellationToken) =>
        _dispatcher.InvokeAsync(operation, cancellationToken);

    private async Task<OfficeObjectResult> WithSessionAsync(
        OfficeBridgeRequest request,
        Func<SessionInfo, OfficeObjectResult> operation,
        CancellationToken cancellationToken)
    {
        var session = request.Params.Deserialize<SessionInfo>(JsonOptions.Default)
            ?? throw new InvalidOperationException("Office formula session parameters are missing.");
        session.ImagePath = string.Equals(
                session.Host,
                "powerpoint",
                StringComparison.OrdinalIgnoreCase)
            ? _tempPathGuard.ValidateSvg(session.ImagePath)
            : _tempPathGuard.ValidatePng(session.ImagePath);
        if (!Guid.TryParse(session.SessionId, out _))
            throw new InvalidOperationException("Session id must be a UUID.");
        if (!Guid.TryParse(session.FormulaId, out _))
            throw new InvalidOperationException("Formula id must be a UUID.");
        if (!string.Equals(session.FormulaId, session.Metadata.FormulaId, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Session and metadata formula ids do not match.");
        return await OnStaAsync(() => operation(session), cancellationToken);
    }

    private object OpenWord()
    {
        RunningOfficeLocator.OpenWord();
        return new { opened = true };
    }

    private object OpenPowerPoint()
    {
        RunningOfficeLocator.OpenPowerPoint();
        return new { opened = true };
    }

    private void OnOfficeDoubleClick(string host)
    {
        if (_disposed) return;
        _ = CaptureDoubleClickTargetAsync(host);
    }

    private async Task CaptureDoubleClickTargetAsync(string host)
    {
        try
        {
            // The low-level mouse hook fires on the second button-down, before
            // Word has finished moving its COM selection onto the InlineShape.
            // Waiting for the host to process the click prevents the subsequent
            // editor command from seeing the caret beside the formula instead.
            await Task.Delay(
                string.Equals(host, "word", StringComparison.Ordinal) ? 120 : 40
            ).ConfigureAwait(false);
            if (_disposed) return;
            await _dispatcher.InvokeAsync(
                () =>
                {
                    OfficeSelection selection = string.Equals(
                        host,
                        "word",
                        StringComparison.Ordinal)
                            ? _word.GetSelection()
                            : _powerPoint.GetSelection();
                    if (selection.Metadata is not null)
                        AddDoubleClickEvent(selection);
                    return true;
                },
                CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception error)
        {
            _logger.Error($"Unable to capture the {host} formula double-click target.", error);
        }
    }

    private void AddDoubleClickEvent(OfficeSelection selection)
    {
        if (!_doubleClickDeduplicator.ShouldAccept(
                selection.Host,
                selection.DocumentId,
                selection.ObjectId,
                selection.FormulaId,
                DateTimeOffset.UtcNow))
            return;
        var cursor = Interlocked.Increment(ref _eventCursor);
        _events.Enqueue(new OfficeBridgeEvent
        {
            Cursor = cursor,
            Event = "office.formulaDoubleClick",
            Payload = new
            {
                host = selection.Host,
                formulaId = selection.FormulaId,
                documentId = selection.DocumentId,
                objectId = selection.ObjectId,
                metadata = selection.Metadata,
            },
        });
        while (_events.Count > 200 && _events.TryDequeue(out _)) { }
    }

    private static string RequiredString(JsonElement element, string name)
    {
        if (!element.TryGetProperty(name, out var property) || property.ValueKind != JsonValueKind.String)
            throw new InvalidOperationException($"Required parameter is missing: {name}");
        return property.GetString()!;
    }

    private static long OptionalLong(JsonElement element, string name) =>
        element.TryGetProperty(name, out var property) && property.TryGetInt64(out var value)
            ? value
            : 0;

    private static string ErrorCode(Exception error) => error switch
    {
        UnauthorizedAccessException => "document_read_only",
        FileNotFoundException => "formula_image_missing",
        InvalidDataException => "invalid_formula_image",
        NotSupportedException => "unsupported_method",
        InvalidOperationException when error.Message.Contains("not running", StringComparison.OrdinalIgnoreCase) => "office_not_running",
        InvalidOperationException when error.Message.Contains("no longer exists", StringComparison.OrdinalIgnoreCase) => "formula_not_found",
        InvalidOperationException when error.Message.Contains("slide show", StringComparison.OrdinalIgnoreCase) => "powerpoint_slide_show",
        _ => "office_operation_failed",
    };

    private static bool IsRetryable(Exception error) => error is IOException
        || (error is InvalidOperationException
            && error.Message.Contains("not running", StringComparison.OrdinalIgnoreCase));

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _doubleClickHook.Dispose();
        _dispatcher.Dispose();
    }
}
