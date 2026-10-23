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

// Where a press reaches, as judged by a mouse hook. Page is the document page
// window that receives the press, or zero when the press lands on something
// else; Detail says why, for the hook trace.
internal readonly struct OfficePagePress
{
    internal OfficePagePress(IntPtr page, string detail)
    {
        Page = page;
        Detail = detail;
    }

    internal IntPtr Page { get; }
    internal string Detail { get; }
    internal bool OnPage => Page != IntPtr.Zero;

    internal static OfficePagePress Off(string detail) => new(IntPtr.Zero, "off " + detail);

    // The hook cannot tell windows apart in physical coordinates, so it pairs
    // presses by time and distance only, as before the page check existed.
    internal static readonly OfficePagePress Unclassified =
        new(new IntPtr(-1), "unclassified thread-not-per-monitor-aware");
}

internal static class OfficeDocumentSurface
{
    // Word draws the page in "_WwG". PowerPoint draws thumbnails, slide and notes
    // inside its document window, "mdiClass" (older versions used separate
    // "paneClassDC" panes inside it). Everything Office shows over them is a
    // window of another class.
    internal static readonly string[] WordPageClasses = { "_WwG" };
    internal static readonly string[] PowerPointPageClasses = { "mdiClass", "paneClassDC" };

    // Tooltips answer WM_NCHITTEST with HTTRANSPARENT, so a press on one goes to
    // the window below it. WindowFromPoint only asks windows of the calling
    // thread, so from the hook thread it still returns the tooltip.
    private static readonly string[] ClickThroughClasses = { "tooltips_class32", "tooltips_class" };

    private const uint GaRoot = 2;
    private const uint GwHwndNext = 2;
    private const int GwlExStyle = -20;
    private const long WsExLayered = 0x00080000;
    private const long WsExTransparent = 0x00000020;
    private const uint CwpSkipInvisible = 0x0001;
    private const uint CwpSkipDisabled = 0x0002;
    private const uint CwpSkipTransparent = 0x0004;

    // The window a press at this point reaches. No messages are sent to other
    // threads, so the low-level mouse hook may call it.
    internal static IntPtr WindowAt(int screenX, int screenY) =>
        WindowFromPoint(new NativePoint { X = screenX, Y = screenY });

    // hit is the window WindowFromPoint returned for the point. A press is on
    // the page when it reaches a page window of the owner process, directly or
    // through a click-through window lying over it. Presses on Office's own
    // dialogs, galleries, menus, task panes and Ribbon, and on other programs'
    // windows (which take the press themselves), are off the page.
    internal static OfficePagePress ResolvePress(
        IntPtr hit,
        string[] pageClasses,
        int ownerProcessId,
        int screenX,
        int screenY)
    {
        if (hit == IntPtr.Zero) return OfficePagePress.Off("no-window");
        if (ownerProcessId <= 0) return OfficePagePress.Off("no-owner");
        var target = hit;
        var top = GetAncestor(hit, GaRoot);
        var through = "";
        // Bounded: only windows stacked at this point are visited.
        for (var depth = 0; depth < 32; depth++)
        {
            var className = ClassOf(target);
            GetWindowThreadProcessId(target, out var processId);
            if (processId != unchecked((uint)ownerProcessId))
                return OfficePagePress.Off($"other-process class={className}{through}");
            if (Array.IndexOf(pageClasses, className) >= 0)
            {
                return IsPage(target, pageClasses, ownerProcessId)
                    ? new OfficePagePress(target, $"page class={className}{through}")
                    : OfficePagePress.Off($"page-unavailable class={className}{through}");
            }
            if (!IsClickThrough(target, className))
                return OfficePagePress.Off($"class={className}{through}");
            through += $" through={className}";
            top = NextTopLevelAt(top, screenX, screenY);
            if (top == IntPtr.Zero) return OfficePagePress.Off($"nothing-below{through}");
            target = DeepestChildAt(top, screenX, screenY);
        }
        return OfficePagePress.Off($"too-deep{through}");
    }

    internal static bool IsPage(
        IntPtr window,
        string[] pageClasses,
        int ownerProcessId)
    {
        if (window == IntPtr.Zero || ownerProcessId <= 0) return false;
        GetWindowThreadProcessId(window, out var processId);
        if (processId != unchecked((uint)ownerProcessId)) return false;
        if (Array.IndexOf(pageClasses, ClassOf(window)) < 0) return false;
        if (!IsWindowVisible(window) || !IsWindowEnabled(window)) return false;
        // A modal dialog disables the document's frame; a press on the page
        // behind it only beeps and must not open anything either.
        var root = GetAncestor(window, GaRoot);
        return root != IntPtr.Zero && IsWindowEnabled(root);
    }

    // Low-level mouse hooks report per-monitor (physical) coordinates; let the
    // hook thread hit-test windows in the same coordinate space. False when
    // the thread could not be made per-monitor aware: hit-testing would then
    // use scaled coordinates and find the wrong window away from the top-left
    // corner of the screen.
    internal static bool UsePhysicalCoordinatesOnThisThread()
    {
        try
        {
            _ = SetThreadDpiAwarenessContext(new IntPtr(-4)); // PER_MONITOR_AWARE_V2
            // PROCESS_PER_MONITOR_DPI_AWARE (v1 or v2)
            return GetAwarenessFromDpiAwarenessContext(GetThreadDpiAwarenessContext()) == 2;
        }
        catch (EntryPointNotFoundException)
        {
            // Before Windows 10 1607 every window shares one DPI, so hook and
            // window coordinates already agree.
            return true;
        }
    }

    private static bool IsClickThrough(IntPtr window, string className)
    {
        if (Array.IndexOf(ClickThroughClasses, className) >= 0) return true;
        var exStyle = GetWindowLongPtr(window, GwlExStyle).ToInt64();
        return (exStyle & (WsExLayered | WsExTransparent)) == (WsExLayered | WsExTransparent);
    }

    // The next visible top-level window below top that contains the point.
    private static IntPtr NextTopLevelAt(IntPtr top, int screenX, int screenY)
    {
        var window = top;
        for (var count = 0; count < 512; count++)
        {
            window = GetWindow(window, GwHwndNext);
            if (window == IntPtr.Zero) return IntPtr.Zero;
            if (!IsWindowVisible(window) || !GetWindowRect(window, out var rect)) continue;
            if (screenX >= rect.Left && screenX < rect.Right
                && screenY >= rect.Top && screenY < rect.Bottom)
                return window;
        }
        return IntPtr.Zero;
    }

    private static IntPtr DeepestChildAt(IntPtr top, int screenX, int screenY)
    {
        var current = top;
        for (var depth = 0; depth < 32; depth++)
        {
            var point = new NativePoint { X = screenX, Y = screenY };
            if (!ScreenToClient(current, ref point)) break;
            var child = ChildWindowFromPointEx(
                current,
                point,
                CwpSkipInvisible | CwpSkipDisabled | CwpSkipTransparent);
            if (child == IntPtr.Zero || child == current) break;
            current = child;
        }
        return current;
    }

    private static string ClassOf(IntPtr window)
    {
        var name = new StringBuilder(64);
        return GetClassName(window, name, name.Capacity) == 0 ? "" : name.ToString();
    }

    private static IntPtr GetWindowLongPtr(IntPtr window, int index) =>
        IntPtr.Size == 8
            ? GetWindowLongPtr64(window, index)
            : new IntPtr(GetWindowLong32(window, index));

    [StructLayout(LayoutKind.Sequential)]
    private struct NativePoint { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect { public int Left; public int Top; public int Right; public int Bottom; }

    [DllImport("user32.dll")]
    private static extern IntPtr WindowFromPoint(NativePoint point);
    [DllImport("user32.dll")]
    private static extern IntPtr ChildWindowFromPointEx(IntPtr parent, NativePoint point, uint flags);
    [DllImport("user32.dll")]
    private static extern IntPtr GetWindow(IntPtr window, uint command);
    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr window, out NativeRect rect);
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
    private static extern IntPtr GetWindowLongPtr64(IntPtr window, int index);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
    private static extern int GetWindowLong32(IntPtr window, int index);
    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")]
    private static extern bool IsWindowEnabled(IntPtr window);
    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr window, uint flags);
    [DllImport("user32.dll")]
    private static extern bool ScreenToClient(IntPtr window, ref NativePoint point);
    [DllImport("user32.dll")]
    private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll")]
    private static extern IntPtr GetThreadDpiAwarenessContext();
    [DllImport("user32.dll")]
    private static extern int GetAwarenessFromDpiAwarenessContext(IntPtr context);
}
