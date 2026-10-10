using Microsoft.Win32;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

/// <summary>
/// Remembers which configured formula chords Word kept for its own bindings, so
/// the VisualTeX settings page can show conflicts that only the running Word
/// knows (custom keys, macros, add-in templates). The value is ASCII only,
/// e.g. "Ctrl+D=1:FormatFont;Ctrl+Alt+C=6:", because the app reads it through
/// reg.exe; category numbers are WdKeyCategory values.
/// </summary>
internal static class WordShortcutConflictRecord
{
    private const string RegistryPath = @"Software\VisualTeX\OfficeIntegration";
    internal const string ValueName = "WordFormulaShortcutConflicts";
    private static readonly object Gate = new();

    internal static void Write(WordShortcutConflict conflict) =>
        Update(conflict.Shortcut.Format(),
            $"{(int)conflict.Category}:{AsciiName(conflict.Command)}");

    internal static void Clear(OfficeFormulaShortcut shortcut) =>
        Update(shortcut.Format(), null);

    internal static string Merge(string? current, string chord, string? entry)
    {
        var entries = new List<string>();
        foreach (var part in (current ?? string.Empty).Split(';'))
        {
            var separator = part.IndexOf('=');
            if (separator <= 0) continue;
            if (string.Equals(part.Substring(0, separator), chord, StringComparison.Ordinal)) continue;
            entries.Add(part);
        }
        if (entry is not null) entries.Add(chord + "=" + entry);
        return string.Join(";", entries);
    }

    internal static string AsciiName(string? name)
    {
        if (string.IsNullOrEmpty(name)) return string.Empty;
        var kept = new string(name!.Where(c => c < 0x80
            && (char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '-')).ToArray());
        // A partly non-ASCII name (a Chinese style or macro) would be misleading.
        return kept.Length == name.Length && kept.Length <= 64 ? kept : string.Empty;
    }

    private static void Update(string chord, string? entry)
    {
        try
        {
            lock (Gate)
            {
                using var key = Registry.CurrentUser.CreateSubKey(RegistryPath, writable: true);
                if (key is null) return;
                var current = key.GetValue(ValueName, null) as string;
                var next = Merge(current, chord, entry);
                if (string.Equals(current ?? string.Empty, next, StringComparison.Ordinal)) return;
                if (next.Length == 0) key.DeleteValue(ValueName, throwOnMissingValue: false);
                else key.SetValue(ValueName, next, RegistryValueKind.String);
            }
        }
        catch
        {
            // Only an explanation for the settings page; never affects routing.
        }
    }
}
