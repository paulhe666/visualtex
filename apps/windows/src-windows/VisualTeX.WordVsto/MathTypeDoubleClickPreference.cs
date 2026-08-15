using System;
using System.IO;
using System.Text.Json;

namespace VisualTeX.WordVsto;

internal static class MathTypeDoubleClickPreference
{
    private const bool DefaultEnabled = true;
    private const int RefreshIntervalMilliseconds = 250;
    private static readonly object Gate = new();
    private static bool _cachedEnabled = DefaultEnabled;
    private static DateTime _nextRefreshUtc = DateTime.MinValue;
    private static DateTime _lastWriteUtc = DateTime.MinValue;

    public static bool IsEnabled()
    {
        var now = DateTime.UtcNow;
        lock (Gate)
        {
            if (now < _nextRefreshUtc)
                return _cachedEnabled;
            _nextRefreshUtc = now.AddMilliseconds(RefreshIntervalMilliseconds);

            try
            {
                var path = ResolvePreferencesPath();
                if (!File.Exists(path))
                {
                    _cachedEnabled = DefaultEnabled;
                    _lastWriteUtc = DateTime.MinValue;
                    return _cachedEnabled;
                }

                var lastWriteUtc = File.GetLastWriteTimeUtc(path);
                if (lastWriteUtc == _lastWriteUtc)
                    return _cachedEnabled;

                using var document = JsonDocument.Parse(File.ReadAllText(path));
                if (document.RootElement.ValueKind == JsonValueKind.Object
                    && document.RootElement.TryGetProperty(
                        "mathtypeDoubleClickEditEnabled",
                        out var enabledElement)
                    && (enabledElement.ValueKind == JsonValueKind.True
                        || enabledElement.ValueKind == JsonValueKind.False))
                {
                    _cachedEnabled = enabledElement.GetBoolean();
                }
                else
                {
                    // Existing installations do not have this key yet. Keep the
                    // product behavior backward compatible: VisualTeX intercepts
                    // MathType OLE double-clicks unless the user explicitly opts out.
                    _cachedEnabled = DefaultEnabled;
                }
                _lastWriteUtc = lastWriteUtc;
            }
            catch
            {
                // A transient read while the desktop app atomically replaces the
                // preference file must never flip behavior. Retain the last known
                // value and retry on the next refresh interval.
            }
            return _cachedEnabled;
        }
    }

    private static string ResolvePreferencesPath()
    {
        var acceptanceOverride = Environment.GetEnvironmentVariable(
            "VISUALTEX_OFFICE_PREFERENCES_PATH");
        if (!string.IsNullOrWhiteSpace(acceptanceOverride))
            return acceptanceOverride;

        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "com.visualtex.studio",
            "office",
            "office-preferences.json");
    }
}
