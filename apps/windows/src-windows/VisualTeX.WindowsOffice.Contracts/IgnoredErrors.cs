using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// In the root namespace so that every VisualTeX namespace reaches it without a using directive.
namespace VisualTeX;

/// <summary>
/// Errors VisualTeX ignores on purpose: an Office call that may be refused, a cleanup that may fail, a probe whose
/// failure is an answer. Every place that ignores an error hands it to <see cref="Ignore"/>, so it can be seen: with
/// VISUALTEX_IGNORED_ERRORS_TRACE_PATH set, each ignored error is appended to that file with the place that ignored
/// it; without it nothing happens.
/// </summary>
public static class IgnoredErrors
{
    private static readonly string? TracePath = ReadTracePath();
    private static readonly object TraceSync = new();

    public static void Ignore(
        Exception error,
        [CallerMemberName] string member = "",
        [CallerFilePath] string file = "",
        [CallerLineNumber] int line = 0)
    {
        if (TracePath is null) return;
        try
        {
            var code = error is ExternalException external ? $" 0x{external.ErrorCode:X8}" : string.Empty;
            using var process = Process.GetCurrentProcess();
            var entry = $"{DateTimeOffset.Now:O} pid={process.Id} tid={Environment.CurrentManagedThreadId} "
                + $"{Path.GetFileName(file)}:{line} {member}: {error.GetType().Name}{code}: {error.Message}"
                + Environment.NewLine;
            lock (TraceSync) File.AppendAllText(TracePath, entry);
        }
        catch
        {
            // Tracing must never change what the code that ignored the error does.
        }
    }

    private static string? ReadTracePath()
    {
        try
        {
            var path = Environment.GetEnvironmentVariable("VISUALTEX_IGNORED_ERRORS_TRACE_PATH");
            return string.IsNullOrWhiteSpace(path) ? null : path;
        }
        catch
        {
            return null;
        }
    }
}
