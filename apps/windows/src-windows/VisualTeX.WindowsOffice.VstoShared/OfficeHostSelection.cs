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
