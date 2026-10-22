using System;
using System.Collections;
using System.Net.Http;
using System.Text.Json;

namespace VisualTeX.WindowsOffice.Contracts;

/// <summary>
/// What went wrong talking to VisualTeX, recorded on the exception's Data so the
/// add-ins can word each cause specifically in the user's language
/// (VstoShared OfficeErrorText). The exception types stay the familiar ones:
/// TimeoutException (a request or batch took too long; never an
/// OperationCanceledException, which means the user stopped), HttpRequestException
/// (VisualTeX unreachable or an HTTP error) and InvalidOperationException (the
/// render batch failed as a whole).
/// </summary>
public static class VisualTeXServiceErrors
{
    /// <summary>One of the Failure* values below.</summary>
    public const string FailureKey = "VisualTeX.Failure";
    /// <summary>The request, e.g. "POST /api/v1/app/converter/render-batches".</summary>
    public const string OperationKey = "VisualTeX.Operation";
    /// <summary>HTTP status code (int).</summary>
    public const string StatusCodeKey = "VisualTeX.StatusCode";
    /// <summary>The server's own error text ({"error": ...}).</summary>
    public const string ServerErrorKey = "VisualTeX.ServerError";
    /// <summary>Time limit in seconds (double).</summary>
    public const string LimitSecondsKey = "VisualTeX.LimitSeconds";
    /// <summary>Number of formulas in the render batch (int).</summary>
    public const string FormulaCountKey = "VisualTeX.FormulaCount";
    /// <summary>The render batch's error code ("cancelled", "converter-not-started", "converter-failed").</summary>
    public const string BatchErrorCodeKey = "VisualTeX.BatchErrorCode";

    public const string FailureRequestTimedOut = "request-timed-out";
    public const string FailureUnreachable = "unreachable";
    public const string FailureHttpStatus = "http-status";
    public const string FailureBatchTimedOut = "batch-timed-out";
    public const string FailureBatchFailed = "batch-failed";

    internal static T With<T>(T error, string failure, string? operation) where T : Exception
    {
        error.Data[FailureKey] = failure;
        if (!string.IsNullOrEmpty(operation)) error.Data[OperationKey] = operation;
        return error;
    }

    /// <summary>The value recorded under <paramref name="key"/>, if any.</summary>
    public static object? Get(Exception error, string key) =>
        error?.Data is IDictionary data && data.Contains(key) ? data[key] : null;

    /// <summary>The innermost exception's message (the actual socket/TLS cause of an HttpRequestException).</summary>
    public static string InnermostMessage(Exception error)
    {
        var current = error;
        while (current.InnerException != null) current = current.InnerException;
        return current.Message;
    }

    /// <summary>The "error" of a JSON error body, or the body itself (trimmed), or null.</summary>
    public static string? ServerErrorText(string? body)
    {
        if (string.IsNullOrWhiteSpace(body)) return null;
        var text = body!.Trim();
        if (text.StartsWith("{", StringComparison.Ordinal))
        {
            try
            {
                using var document = JsonDocument.Parse(text);
                if (document.RootElement.ValueKind == JsonValueKind.Object
                    && document.RootElement.TryGetProperty("error", out var value)
                    && value.ValueKind == JsonValueKind.String
                    && !string.IsNullOrWhiteSpace(value.GetString()))
                    return value.GetString()!.Trim();
            }
            catch (JsonException)
            {
                // not JSON after all: the body itself is the detail
            }
        }
        return text.Length > 1000 ? text.Substring(0, 1000) + "…" : text;
    }

    internal static TimeoutException RequestTimedOut(string operation, TimeSpan limit, Exception inner)
    {
        var error = With(
            new TimeoutException(
                $"VisualTeX did not answer '{operation}' within {limit.TotalSeconds:0.#} s.", inner),
            FailureRequestTimedOut, operation);
        error.Data[LimitSecondsKey] = limit.TotalSeconds;
        return error;
    }

    internal static HttpRequestException Unreachable(string operation, Exception inner) =>
        With(
            new HttpRequestException(
                $"VisualTeX could not be reached for '{operation}': {InnermostMessage(inner)}", inner),
            FailureUnreachable, operation);
}
