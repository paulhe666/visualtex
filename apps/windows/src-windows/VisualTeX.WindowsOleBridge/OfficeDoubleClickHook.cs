using System.Diagnostics;
using System.Runtime.InteropServices;

namespace VisualTeX.WindowsOleBridge;

internal sealed class OfficeDoubleClickHook : IDisposable
{
    private const int WhMouseLl = 14;
    private const int WmLButtonDown = 0x0201;
    private const int WmQuit = 0x0012;
    private const int SmCxDoubleClk = 36;
    private const int SmCyDoubleClk = 37;
    private readonly Action<string> _onDoubleClick;
    private readonly FileLogger _logger;
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _ready = new(false);
    private HookProc? _callback;
    private IntPtr _hook;
    private uint _threadId;
    private long _lastRaisedTicks;
    private long _lastButtonDownTicks;
    private Point _lastButtonDownPoint;
    private bool _started;

    public OfficeDoubleClickHook(Action<string> onDoubleClick, FileLogger logger)
    {
        _onDoubleClick = onDoubleClick;
        _logger = logger;
        _thread = new Thread(Run)
        {
            Name = "VisualTeX Office Double Click Hook",
            IsBackground = true,
        };
    }

    public void Start()
    {
        if (_started) return;
        _started = true;
        _thread.Start();
        if (!_ready.Wait(TimeSpan.FromSeconds(5)))
            throw new TimeoutException("Office double-click hook did not start in time.");
    }

    private void Run()
    {
        _threadId = GetCurrentThreadId();
        _callback = HookCallback;
        _hook = SetWindowsHookEx(WhMouseLl, _callback, IntPtr.Zero, 0);
        _ready.Set();
        if (_hook == IntPtr.Zero)
        {
            _logger.Error("Unable to install Office double-click hook.");
            return;
        }
        while (GetMessage(out var message, IntPtr.Zero, 0, 0) > 0)
        {
            TranslateMessage(ref message);
            DispatchMessage(ref message);
        }
        UnhookWindowsHookEx(_hook);
        _hook = IntPtr.Zero;
    }

    private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
    {
        if (code >= 0 && wParam.ToInt32() == WmLButtonDown)
        {
            var now = Stopwatch.GetTimestamp();
            var point = Marshal.PtrToStructure<LowLevelMouseInput>(lParam).Pt;
            var previous = _lastButtonDownTicks;
            var elapsedMilliseconds = previous == 0
                ? double.MaxValue
                : (now - previous) * 1000d / Stopwatch.Frequency;
            var withinDoubleClickRectangle =
                Math.Abs(point.X - _lastButtonDownPoint.X) <= Math.Max(1, GetSystemMetrics(SmCxDoubleClk) / 2) &&
                Math.Abs(point.Y - _lastButtonDownPoint.Y) <= Math.Max(1, GetSystemMetrics(SmCyDoubleClk) / 2);
            var foregroundHost = GetOfficeForegroundHost();
            if (foregroundHost is not null &&
                elapsedMilliseconds <= GetDoubleClickTime() &&
                withinDoubleClickRectangle)
            {
                _lastButtonDownTicks = 0;
                var sinceLastRaise = (now - Interlocked.Read(ref _lastRaisedTicks)) /
                    (double)Stopwatch.Frequency;
                if (sinceLastRaise >= 0.75)
                {
                    Interlocked.Exchange(ref _lastRaisedTicks, now);
                    try { _onDoubleClick(foregroundHost); } catch (Exception error) { _logger.Error("Double-click handler failed.", error); }
                }
            }
            else
            {
                _lastButtonDownTicks = now;
                _lastButtonDownPoint = point;
            }
        }
        return CallNextHookEx(_hook, code, wParam, lParam);
    }

    private static string? GetOfficeForegroundHost()
    {
        var window = GetForegroundWindow();
        if (window == IntPtr.Zero) return null;
        GetWindowThreadProcessId(window, out var processId);
        if (processId == 0) return null;
        try
        {
            var name = Process.GetProcessById((int)processId).ProcessName;
            if (string.Equals(name, "WINWORD", StringComparison.OrdinalIgnoreCase))
                return "word";
            if (string.Equals(name, "POWERPNT", StringComparison.OrdinalIgnoreCase))
                return "powerpoint";
            return null;
        }
        catch { return null; }
    }

    public void Dispose()
    {
        if (!_started) return;
        if (_threadId != 0) PostThreadMessage(_threadId, WmQuit, IntPtr.Zero, IntPtr.Zero);
        _thread.Join(TimeSpan.FromSeconds(2));
        _ready.Dispose();
    }

    private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct Point { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct LowLevelMouseInput
    {
        public Point Pt;
        public uint MouseData;
        public uint Flags;
        public uint Time;
        public UIntPtr ExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Message
    {
        public IntPtr HWnd;
        public uint Value;
        public UIntPtr WParam;
        public IntPtr LParam;
        public uint Time;
        public Point Pt;
    }

    [DllImport("user32.dll")]
    private static extern IntPtr SetWindowsHookEx(int idHook, HookProc callback, IntPtr module, uint threadId);
    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hook);
    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    private static extern int GetMessage(out Message message, IntPtr window, uint min, uint max);
    [DllImport("user32.dll")]
    private static extern bool TranslateMessage(ref Message message);
    [DllImport("user32.dll")]
    private static extern IntPtr DispatchMessage(ref Message message);
    [DllImport("user32.dll")]
    private static extern bool PostThreadMessage(uint threadId, int message, IntPtr wParam, IntPtr lParam);
    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("user32.dll")]
    private static extern uint GetDoubleClickTime();
    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int index);
}
