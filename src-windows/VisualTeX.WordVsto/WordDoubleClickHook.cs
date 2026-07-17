using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace VisualTeX.WordVsto;

internal sealed class WordDoubleClickHook : IDisposable
{
    private const int WhMouseLl = 14;
    private const int WmLButtonDown = 0x0201;
    private const int WmQuit = 0x0012;
    private const int SmCxDoubleClick = 36;
    private const int SmCyDoubleClick = 37;
    private static readonly object TraceGate = new();
    private readonly Func<int, int, bool> _shouldHandle;
    private readonly Action _callbackAction;
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _ready = new(false);
    private HookProc? _hookCallback;
    private IntPtr _hook;
    private uint _threadId;
    private long _lastClickTimestamp;
    private int _lastClickX = int.MinValue;
    private int _lastClickY = int.MinValue;

    public WordDoubleClickHook(
        Func<int, int, bool> shouldHandle,
        Action callbackAction)
    {
        _shouldHandle = shouldHandle;
        _callbackAction = callbackAction;
        _thread = new Thread(Run)
        {
            IsBackground = true,
            Name = "VisualTeX Word OLE Double Click",
        };
    }

    public void Start()
    {
        _thread.Start();
        if (!_ready.Wait(TimeSpan.FromSeconds(5)))
            throw new TimeoutException("Word OLE double-click hook did not start.");
        if (_hook == IntPtr.Zero)
            throw new Win32Exception(
                Marshal.GetLastWin32Error(),
                "Word OLE double-click hook could not be installed.");
        TraceMessage($"hook-started handle=0x{_hook.ToInt64():X}");
    }

    internal static void TraceMessage(string message)
    {
        var path = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        if (string.IsNullOrWhiteSpace(path)) return;
        try
        {
            lock (TraceGate)
            {
                File.AppendAllText(
                    path,
                    $"{DateTimeOffset.Now:O} pid={Process.GetCurrentProcess().Id} tid={Environment.CurrentManagedThreadId} {message}{Environment.NewLine}");
            }
        }
        catch { }
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
        if (code < 0 || wParam.ToInt32() != WmLButtonDown)
            return CallNextHookEx(_hook, code, wParam, lParam);

        var input = Marshal.PtrToStructure<LowLevelMouseInput>(lParam);
        var wordForeground = IsWordForeground();
        TraceMessage($"left-down x={input.Pt.X} y={input.Pt.Y} wordForeground={wordForeground}");
        if (!wordForeground)
            return CallNextHookEx(_hook, code, wParam, lParam);
        var now = Stopwatch.GetTimestamp();
        var previous = Interlocked.Read(ref _lastClickTimestamp);
        var elapsedMilliseconds = previous == 0
            ? double.PositiveInfinity
            : (now - previous) * 1000d / Stopwatch.Frequency;
        var withinTime = elapsedMilliseconds <= GetDoubleClickTime();
        var withinX = Math.Abs(input.Pt.X - _lastClickX)
            <= Math.Max(1, GetSystemMetrics(SmCxDoubleClick));
        var withinY = Math.Abs(input.Pt.Y - _lastClickY)
            <= Math.Max(1, GetSystemMetrics(SmCyDoubleClick));

        _lastClickX = input.Pt.X;
        _lastClickY = input.Pt.Y;
        Interlocked.Exchange(ref _lastClickTimestamp, now);
        if (!withinTime || !withinX || !withinY)
            return CallNextHookEx(_hook, code, wParam, lParam);

        Interlocked.Exchange(ref _lastClickTimestamp, 0);
        bool handle;
        try { handle = _shouldHandle(input.Pt.X, input.Pt.Y); }
        catch (Exception error)
        {
            TraceMessage($"hit-test-error {error.GetType().Name}: {error.Message}");
            handle = false;
        }
        TraceMessage(
            $"double-click elapsedMs={elapsedMilliseconds:0.###} withinX={withinX} withinY={withinY} handle={handle}");
        if (!handle)
            return CallNextHookEx(_hook, code, wParam, lParam);

        ThreadPool.QueueUserWorkItem(_ =>
        {
            // The first click already selected the inline OLE object. Suppress
            // the second button-down so Word cannot invoke the OLE default verb,
            // then open the VisualTeX Session on the Office UI thread.
            Thread.Sleep(40);
            try
            {
                TraceMessage("callback-begin");
                _callbackAction();
                TraceMessage("callback-end");
            }
            catch (Exception error)
            {
                TraceMessage($"callback-error {error.GetType().Name}: {error.Message}");
            }
        });
        TraceMessage("second-button-down-suppressed");
        return new IntPtr(1);
    }

    private static bool IsWordForeground()
    {
        var window = GetForegroundWindow();
        if (window == IntPtr.Zero) return false;
        GetWindowThreadProcessId(window, out var processId);
        if (processId == 0) return false;
        try
        {
            return string.Equals(
                Process.GetProcessById((int)processId).ProcessName,
                "WINWORD",
                StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    public void Dispose()
    {
        if (_threadId != 0)
            PostThreadMessage(_threadId, WmQuit, IntPtr.Zero, IntPtr.Zero);
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

    [DllImport("user32.dll", SetLastError = true)]
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
    private static extern uint GetDoubleClickTime();
    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int index);
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
