using System.Diagnostics;

namespace VisualTeX.WindowsOffice.VstoShared;

/// <summary>
/// The Office-compatible host the add-in runs in. WPS Office (Kingsoft) loads the same
/// COM add-ins and implements the Word/PowerPoint object models, but a few members fail
/// there. Inside a custom UndoRecord any failed call makes WPS roll the whole record back
/// at EndCustomRecord, so WPS-only paths must avoid those members instead of relying on a
/// swallowed exception.
/// </summary>
internal static class OfficeHostEnvironment
{
    private static readonly Lazy<bool> Wps = new(DetectWps);

    /// <summary>True inside WPS Writer (wps.exe) or WPS Presentation (wpp.exe).</summary>
    internal static bool IsWps => Wps.Value;

    private static bool DetectWps()
    {
        try
        {
            using var process = Process.GetCurrentProcess();
            var name = process.ProcessName;
            return string.Equals(name, "wps", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "wpp", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "et", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }
}
