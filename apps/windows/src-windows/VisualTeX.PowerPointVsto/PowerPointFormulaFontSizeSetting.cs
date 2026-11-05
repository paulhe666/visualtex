using System.Globalization;
using Microsoft.Win32;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.PowerPointVsto;

/// <summary>
/// The per-user inline and display formula sizes (Ribbon 行内 / 行间, HKCU), as in Word. A size, once chosen, is given
/// to every formula of that kind in the presentation, and LaTeX redraw and format conversion write their new formulas
/// of that kind at it. None: the formulas keep the size found for them.
/// </summary>
internal static class PowerPointFormulaFontSizeSetting
{
    private const string RegistryPath = @"Software\VisualTeX\PowerPointVsto";
    private const string InlineValueName = "InlineFormulaFontSize";
    private const string DisplayValueName = "DisplayFormulaFontSize";

    private static readonly object Gate = new();
    private static bool _loaded;
    private static float? _inline;
    private static float? _display;

    internal static float? Get(bool display)
    {
        lock (Gate)
        {
            if (!_loaded)
            {
                _inline = Read(InlineValueName);
                _display = Read(DisplayValueName);
                _loaded = true;
            }
            return display ? _display : _inline;
        }
    }

    /// <summary>The size a new formula of this kind gets: the setting, or the size found for it.</summary>
    internal static double Resolve(bool display, double found) => Get(display) ?? found;

    internal static double Resolve(string? displayMode, double found) =>
        Resolve(!string.Equals(displayMode, "inline", StringComparison.OrdinalIgnoreCase), found);

    internal static void Save(bool display, float? size)
    {
        lock (Gate)
        {
            _ = Get(display);
            if (display) _display = size; else _inline = size;
        }
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(RegistryPath, writable: true);
            var name = display ? DisplayValueName : InlineValueName;
            if (size is { } value)
                key?.SetValue(name, value.ToString(CultureInfo.InvariantCulture), RegistryValueKind.String);
            else
                key?.DeleteValue(name, throwOnMissingValue: false);
        }
        catch
        {
            // Keep the in-process choice even if the profile rejects persistence.
        }
    }

    private static float? Read(string name)
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RegistryPath, writable: false);
            if (key?.GetValue(name) is string stored
                && float.TryParse(stored, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
                && value >= FormulaFontSize.MinimumPt
                && value <= FormulaFontSize.MaximumPt)
                return FormulaFontSize.Normalize(value);
        }
        catch { }
        return null;
    }
}
