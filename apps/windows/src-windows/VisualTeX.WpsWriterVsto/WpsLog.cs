using System.IO;
using System.Text;

namespace VisualTeX.WpsWriterVsto;

internal static class WpsLog
{
    private static readonly object Gate = new();

    internal static string LogPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "VisualTeX",
        "office",
        "logs",
        "wps-writer-vsto.log");

    internal static void Write(string message)
    {
        try
        {
            lock (Gate)
            {
                var directory = Path.GetDirectoryName(LogPath)!;
                Directory.CreateDirectory(directory);
                File.AppendAllText(
                    LogPath,
                    $"{DateTimeOffset.Now:O} {message}{Environment.NewLine}",
                    new UTF8Encoding(false));
            }
        }
        catch
        {
            // Add-in logging must never break the host.
        }
    }

    internal static void WriteException(string stage, Exception error) =>
        Write($"{stage}: {error.GetType().Name}: {error.Message}{Environment.NewLine}{error}");
}
