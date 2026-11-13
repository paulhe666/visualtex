using System.Diagnostics;
using System.Globalization;
using Microsoft.Win32;

namespace VisualTeX.WindowsOffice.VstoShared;

/// <summary>
/// Which hosts the user installed the add-ins for (the installer's components page): Microsoft Office and WPS
/// Office load the same COM add-ins from the same registration, so the add-in itself stays inert in a host the
/// user left out. HKCU\Software\VisualTeX\OfficeIntegration, DWORD MicrosoftOffice / Wps: 0 = off; missing = on.
/// </summary>
internal static class OfficeHostSelection
{
    private const string KeyPath = @"Software\VisualTeX\OfficeIntegration";

    /// <summary>Whether the add-in should run in the current host.</summary>
    internal static bool IsCurrentHostEnabled() => IsEnabled(OfficeHostEnvironment.IsWps ? "Wps" : "MicrosoftOffice");

    /// <summary>
    /// Notes that a host loaded the add-in, so VisualTeX's settings can tell whether WPS (or
    /// Office) loaded VisualTeX at all: string value &lt;host&gt;Loaded = "local time|host
    /// version|bitness|on or off" (off: left out at installation, no ribbon).
    /// </summary>
    internal static void RecordLoad(string host)
    {
        try
        {
            var version = string.Empty;
            try
            {
                using var process = Process.GetCurrentProcess();
                version = process.MainModule?.FileVersionInfo.ProductVersion ?? string.Empty;
            }
            catch { }
            var value = string.Join("|",
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
                version.Replace("|", " "),
                Environment.Is64BitProcess ? "x64" : "x86",
                IsCurrentHostEnabled() ? "on" : "off");
            using var key = Registry.CurrentUser.CreateSubKey(KeyPath);
            key?.SetValue(host + "Loaded", value, RegistryValueKind.String);
        }
        catch { }
    }

    private static bool IsEnabled(string name)
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(KeyPath);
            return key?.GetValue(name) is not int value || value != 0;
        }
        catch
        {
            return true;
        }
    }
}
