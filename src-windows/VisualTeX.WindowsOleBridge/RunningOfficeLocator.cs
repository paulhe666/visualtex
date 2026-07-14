using System.Runtime.InteropServices;

namespace VisualTeX.WindowsOleBridge;

internal static class RunningOfficeLocator
{
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string progId, out Guid clsid);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(
        ref Guid rclsid,
        IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object? instance);

    public static object GetWordApplication() => GetRequiredActiveObject("Word.Application");
    public static object GetPowerPointApplication() =>
        GetRequiredActiveObject("PowerPoint.Application");

    public static bool IsRunning(string progId)
    {
        object? instance = null;
        try
        {
            instance = GetRequiredActiveObject(progId);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            ComRelease.Final(instance);
        }
    }

    public static void OpenWord() =>
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("winword.exe")
        {
            UseShellExecute = true,
        });

    public static void OpenPowerPoint() =>
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("powerpnt.exe")
        {
            UseShellExecute = true,
        });

    private static object GetRequiredActiveObject(string progId)
    {
        var hr = CLSIDFromProgID(progId, out var clsid);
        if (hr < 0) Marshal.ThrowExceptionForHR(hr);
        hr = GetActiveObject(ref clsid, IntPtr.Zero, out var instance);
        if (hr < 0 || instance is null)
            throw new InvalidOperationException($"{progId} is not running.");
        return instance;
    }
}

internal static class ComRelease
{
    public static void Final(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.FinalReleaseComObject(value); } catch { }
    }
}
