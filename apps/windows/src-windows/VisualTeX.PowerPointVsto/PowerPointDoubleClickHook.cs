using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace VisualTeX.PowerPointVsto;

/// <summary>
/// Watches left-button double-clicks on PowerPoint's own UI thread (WH_MOUSE).
/// PowerPoint retrieves mouse messages in order, so when the second click is
/// retrieved the first one has been fully processed, and the callback may ask
/// PowerPoint what lies under the pointer. A desktop-wide low-level hook sees the
/// second click earlier, on another thread, before PowerPoint has caught up, and
/// cannot call Office at all.
/// </summary>
internal sealed class PowerPointDoubleClickHook : IDisposable
{
    private const int WhMouse = 7;
    private const int HcAction = 0;
    private const int WmLButtonDown = 0x0201;
    private const int WmLButtonDblClk = 0x0203;
    private const int WmNcLButtonDown = 0x00A1;
    private const int WmNcLButtonDblClk = 0x00A3;
    private const int SmCxDoubleClick = 36;
    private const int SmCyDoubleClick = 37;
    private const long TraceLimitBytes = 512 * 1024;
    private static readonly object TraceGate = new();
    private static readonly int OwnerProcessId = Process.GetCurrentProcess().Id;
    private readonly Func<int, int, bool> _claimSecondClick;
    private readonly Action<bool, int, int> _callbackAction;
    private HookProc? _hookCallback;
    private IntPtr _hook;
    private uint _lastClickTime;
    private bool _hasLastClick;
    private int _lastClickX;
    private int _lastClickY;

    // Both callbacks run on the Office UI thread inside the hook. When
    // claimSecondClick returns true the second click is swallowed, so PowerPoint
    // never runs its own double-click (which would activate the OLE server).
    public PowerPointDoubleClickHook(
        Func<int, int, bool> claimSecondClick,
        Action<bool, int, int> callbackAction)
    {
        _claimSecondClick = claimSecondClick;
        _callbackAction = callbackAction;
    }

    // Call on the Office UI thread.
    public void Start()
    {
        if (_hook != IntPtr.Zero) return;
        _hookCallback = HookCallback;
        _hook = SetWindowsHookEx(WhMouse, _hookCallback, IntPtr.Zero, GetCurrentThreadId());
        if (_hook == IntPtr.Zero)
            throw new Win32Exception(
                Marshal.GetLastWin32Error(),
                "PowerPoint double-click hook could not be installed.");
        TraceMessage($"hook-started handle=0x{_hook.ToInt64():X} thread={GetCurrentThreadId()}");
    }

    // %LOCALAPPDATA%\VisualTeX\office\logs\powerpoint-double-click.log. Written
    // on the thread pool so the UI thread never waits for the disk.
    internal static void TraceMessage(string message)
    {
        var line = $"{DateTimeOffset.Now:O} pid={OwnerProcessId} {message}{Environment.NewLine}";
        ThreadPool.QueueUserWorkItem(_ =>
        {
            try
            {
                var directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "VisualTeX", "office", "logs");
                var path = Path.Combine(directory, "powerpoint-double-click.log");
                lock (TraceGate)
                {
                    Directory.CreateDirectory(directory);
                    var file = new FileInfo(path);
                    if (file.Exists && file.Length > TraceLimitBytes)
                    {
                        var previous = path + ".1";
                        if (File.Exists(previous)) File.Delete(previous);
                        File.Move(path, previous);
                    }
                    File.AppendAllText(path, line);
                }
            }
            catch { }
        });
    }

    private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
    {
        // HC_NOREMOVE is only a peek; act once, when the message is retrieved.
        if (code != HcAction) return CallNextHookEx(_hook, code, wParam, lParam);
        var message = wParam.ToInt32();
        if (message is WmNcLButtonDown or WmNcLButtonDblClk)
        {
            _hasLastClick = false;
            return CallNextHookEx(_hook, code, wParam, lParam);
        }
        if (message is not (WmLButtonDown or WmLButtonDblClk))
            return CallNextHookEx(_hook, code, wParam, lParam);

        var input = Marshal.PtrToStructure<MouseHookInput>(lParam);
        var now = unchecked((uint)Environment.TickCount);
        // Windows reports the second click as WM_LBUTTONDBLCLK only for windows
        // whose class asks for it; otherwise compare with the previous click.
        var isDouble = message == WmLButtonDblClk
            || _hasLastClick
                && unchecked(now - _lastClickTime) <= GetDoubleClickTime()
                && Math.Abs((long)input.Pt.X - _lastClickX) <= Math.Max(1, GetSystemMetrics(SmCxDoubleClick))
                && Math.Abs((long)input.Pt.Y - _lastClickY) <= Math.Max(1, GetSystemMetrics(SmCyDoubleClick));
        if (!isDouble)
        {
            _hasLastClick = true;
            _lastClickTime = now;
            _lastClickX = input.Pt.X;
            _lastClickY = input.Pt.Y;
            return CallNextHookEx(_hook, code, wParam, lParam);
        }

        _hasLastClick = false;
        var started = Stopwatch.GetTimestamp();
        bool claimed;
        try { claimed = _claimSecondClick(input.Pt.X, input.Pt.Y); }
        catch (Exception error)
        {
            TraceMessage($"claim-error {error.GetType().Name}: {error.Message}");
            claimed = false;
        }
        var elapsed = (Stopwatch.GetTimestamp() - started) * 1000d / Stopwatch.Frequency;
        TraceMessage($"double-click x={input.Pt.X} y={input.Pt.Y} message=0x{message:X} claimed={claimed} decideMs={elapsed:0.###}");
        try { _callbackAction(claimed, input.Pt.X, input.Pt.Y); }
        catch (Exception error) { TraceMessage($"callback-error {error.GetType().Name}: {error.Message}"); }
        return claimed ? new IntPtr(1) : CallNextHookEx(_hook, code, wParam, lParam);
    }

    public void Dispose()
    {
        if (_hook != IntPtr.Zero)
        {
            try { UnhookWindowsHookEx(_hook); } catch { }
            _hook = IntPtr.Zero;
        }
        _hookCallback = null;
    }

    private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct Point { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct MouseHookInput
    {
        public Point Pt;
        public IntPtr Window;
        public uint HitTestCode;
        public UIntPtr ExtraInfo;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, HookProc callback, IntPtr module, uint threadId);
    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hook);
    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);
    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")]
    private static extern uint GetDoubleClickTime();
    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int index);
}
