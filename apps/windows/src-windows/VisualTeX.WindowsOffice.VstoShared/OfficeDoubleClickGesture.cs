using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace VisualTeX.WindowsOffice.VstoShared;

// A double-click may open a formula only when both presses land on the same
// document page window, which is also what Windows requires before it reports
// a double-click. Presses on anything else Office shows (a dialog, gallery,
// menu, task pane or the Ribbon, even one lying over a formula) never pair, so
// they can neither open the formula underneath nor lose their second press.
internal sealed class OfficeDoubleClickPairing
{
    private readonly object _gate = new();
    private bool _hasPrevious;
    private IntPtr _previousWindow;
    private long _previousMilliseconds;
    private int _previousX;
    private int _previousY;

    // True when this press completes a double-click on the page. The pair is
    // consumed, so a third quick press starts a new gesture.
    internal bool Register(
        IntPtr window,
        bool onDocumentPage,
        int screenX,
        int screenY,
        long timestampMilliseconds,
        long doubleClickMilliseconds,
        int maxDeltaX,
        int maxDeltaY)
    {
        lock (_gate)
        {
            if (!onDocumentPage || window == IntPtr.Zero)
            {
                _hasPrevious = false;
                return false;
            }
            var elapsed = timestampMilliseconds - _previousMilliseconds;
            // Widen to Int64 before subtracting: coordinates may be anywhere on
            // a multi-monitor desktop.
            var paired = _hasPrevious
                && window == _previousWindow
                && elapsed >= 0
                && elapsed <= doubleClickMilliseconds
                && Math.Abs((long)screenX - _previousX) <= Math.Max(1, maxDeltaX)
                && Math.Abs((long)screenY - _previousY) <= Math.Max(1, maxDeltaY);
            if (paired)
            {
                _hasPrevious = false;
                return true;
            }
            _hasPrevious = true;
            _previousWindow = window;
            _previousMilliseconds = timestampMilliseconds;
            _previousX = screenX;
            _previousY = screenY;
            return false;
        }
    }

    internal void Reset()
    {
        lock (_gate) _hasPrevious = false;
    }

    // Monotonic milliseconds. Multiplying raw ticks by 1000 in Int64 could
    // overflow after weeks of uptime on a high-frequency counter.
    internal static long NowMilliseconds() =>
        (long)(Stopwatch.GetTimestamp() * (1000d / Stopwatch.Frequency));
}

internal static class OfficeDocumentSurface
{
    // Word draws the page in "_WwG". PowerPoint draws thumbnails, slide and notes
    // inside its document window, "mdiClass" (older versions used separate
    // "paneClassDC" panes inside it). Everything Office shows over them is a
    // window of another class.
    internal static readonly string[] WordPageClasses = { "_WwG" };
    internal static readonly string[] PowerPointPageClasses = { "mdiClass", "paneClassDC" };

    private const uint GaRoot = 2;

    // The window a press at this point reaches. No messages are sent to other
    // threads, so the low-level mouse hook may call it.
    internal static IntPtr WindowAt(int screenX, int screenY) =>
        WindowFromPoint(new NativePoint { X = screenX, Y = screenY });

    internal static bool IsPage(
        IntPtr window,
        string[] pageClasses,
        int ownerProcessId,
        int screenX,
        int screenY)
    {
        if (window == IntPtr.Zero || ownerProcessId <= 0) return false;
        GetWindowThreadProcessId(window, out var processId);
        if (processId != unchecked((uint)ownerProcessId)) return false;
        var name = new StringBuilder(32);
        if (GetClassName(window, name, name.Capacity) == 0
            || Array.IndexOf(pageClasses, name.ToString()) < 0)
            return false;
        if (!IsWindowVisible(window) || !IsWindowEnabled(window)) return false;
        // A modal dialog disables the document's frame; a press on the page
        // behind it only beeps and must not open anything either.
        var root = GetAncestor(window, GaRoot);
        if (root == IntPtr.Zero || !IsWindowEnabled(root)) return false;
        var point = new NativePoint { X = screenX, Y = screenY };
        if (!ScreenToClient(window, ref point) || !GetClientRect(window, out var client))
            return false;
        return point.X >= client.Left && point.X < client.Right
            && point.Y >= client.Top && point.Y < client.Bottom;
    }

    // Low-level mouse hooks report per-monitor (physical) coordinates; let the
    // hook thread hit-test windows in the same coordinate space.
    internal static void UsePhysicalCoordinatesOnThisThread()
    {
        try { _ = SetThreadDpiAwarenessContext(new IntPtr(-4)); } // PER_MONITOR_AWARE_V2
        catch (EntryPointNotFoundException) { }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativePoint { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect { public int Left; public int Top; public int Right; public int Bottom; }

    [DllImport("user32.dll")]
    private static extern IntPtr WindowFromPoint(NativePoint point);
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")]
    private static extern bool IsWindowEnabled(IntPtr window);
    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr window, uint flags);
    [DllImport("user32.dll")]
    private static extern bool ScreenToClient(IntPtr window, ref NativePoint point);
    [DllImport("user32.dll")]
    private static extern bool GetClientRect(IntPtr window, out NativeRect rect);
    [DllImport("user32.dll")]
    private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
}
