using System.Text;

namespace VisualTeX.WindowsOleBridge;

internal sealed class FileLogger
{
    private readonly string _path;
    private readonly object _gate = new();

    public FileLogger(string root)
    {
        Directory.CreateDirectory(root);
        _path = Path.Combine(root, "windows-office-bridge.log");
    }

    public void Info(string message) => Write("INFO", message);
    public void Error(string message, Exception? exception = null) =>
        Write("ERROR", exception is null ? message : $"{message}\n{exception}");

    private void Write(string level, string message)
    {
        var line = $"{DateTimeOffset.Now:O} [{level}] {message}{Environment.NewLine}";
        lock (_gate)
        {
            File.AppendAllText(_path, line, Encoding.UTF8);
        }
    }
}
