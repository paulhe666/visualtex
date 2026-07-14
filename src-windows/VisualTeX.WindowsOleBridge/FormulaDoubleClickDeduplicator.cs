namespace VisualTeX.WindowsOleBridge;

internal sealed class FormulaDoubleClickDeduplicator
{
    private readonly TimeSpan _window;
    private readonly object _gate = new();
    private string _lastKey = string.Empty;
    private DateTimeOffset _lastAcceptedAt = DateTimeOffset.MinValue;

    public FormulaDoubleClickDeduplicator(TimeSpan? window = null)
    {
        _window = window ?? TimeSpan.FromSeconds(1);
    }

    public bool ShouldAccept(
        string host,
        string? documentId,
        string? objectId,
        string? formulaId,
        DateTimeOffset now)
    {
        if (string.IsNullOrWhiteSpace(formulaId)) return false;
        var key = string.Join(
            "\u001f",
            host,
            documentId ?? string.Empty,
            objectId ?? string.Empty,
            formulaId);
        lock (_gate)
        {
            if (string.Equals(key, _lastKey, StringComparison.Ordinal)
                && now - _lastAcceptedAt < _window)
                return false;
            _lastKey = key;
            _lastAcceptedAt = now;
            return true;
        }
    }
}
