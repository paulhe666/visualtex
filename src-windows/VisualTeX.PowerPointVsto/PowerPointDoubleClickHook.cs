using System.Diagnostics;
using System.Runtime.InteropServices;

namespace VisualTeX.PowerPointVsto;

internal sealed class PowerPointDoubleClickHook : IDisposable
{
    private const int WhMouseLl = 14;
    private const int WmLButtonDblClk = 0x0203;
    private const int WmQuit = 0x0012;
    private readonly Action _callbackAction;
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _ready = new(false);
    private HookProc? _hookCallback;
    private IntPtr _hook;
    private uint _threadId;
    private long _lastRaised;

    public PowerPointDoubleClickHook(Action callbackAction)
    {
        _callbackAction = callbackAction;
        _thread = new Thread(Run)
        {
            IsBackground = true,
            Name = "VisualTeX PowerPoint Double Click",
        };
    }

    public void Start()
    {
        _thread.Start();
        if (!_ready.Wait(TimeSpan.FromSeconds(5)))
            throw new TimeoutException("PowerPoint double-click hook did not start.");
    }

    private void Run()
    {
        _threadId = GetCurrentThreadId();
        _hookCallback = HookCallback;
        _hook = SetWindowsHookEx(WhMouseLl, _hookCallback, IntPtr.Zero, 0);
        _ready.Set();
        if (_hook == IntPtr.Zero) return;
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
        if (code >= 0 && wParam.ToInt32() == WmLButtonDblClk && IsPowerPointForeground())
        {
            var now = Stopwatch.GetTimestamp();
            var elapsed = (now - Interlocked.Read(ref _lastRaised)) /
                (double)Stopwatch.Frequency;
            if (elapsed >= 0.75)
            {
                Interlocked.Exchange(ref _lastRaised, now);
                try { _callbackAction(); } catch { }
            }
        }
        return CallNextHookEx(_hook, code, wParam, lParam);
    }

    private static bool IsPowerPointForeground()
    {
        var window = GetForegroundWindow();
        if (window == IntPtr.Zero) return false;
        GetWindowThreadProcessId(window, out var processId);
        if (processId == 0) return false;
        try
        {
            return string.Equals(
                Process.GetProcessById((int)processId).ProcessName,
                "POWERPNT",
                StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    public void Dispose()
    {
        if (_threadId != 0) PostThreadMessage(_threadId, WmQuit, IntPtr.Zero, IntPtr.Zero);
        _thread.Join(TimeSpan.FromSeconds(2));
        _ready.Dispose();
    }

    private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct Point { public int X; public int Y; }

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
}
