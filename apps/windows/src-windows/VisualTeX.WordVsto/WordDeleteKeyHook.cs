using System.ComponentModel;
using System.Runtime.InteropServices;

namespace VisualTeX.WordVsto;

/// <summary>
/// Observes Delete/Backspace on this Word UI thread. The hook never suppresses
/// or rewrites Word input; it only schedules a post-key notification after Word
/// has processed the native deletion.
/// </summary>
internal sealed class WordDeleteKeyHook : IDisposable
{
    private const int WhKeyboard = 2;
    private const uint VkBack = 0x08;
    private const uint VkDelete = 0x2E;

    private readonly Action _callback;
    private HookProc? _hookCallback;
    private IntPtr _hook;
    private uint _ownerThreadId;
    private int _ownerProcessId;
    private long _lastDeleteObservationTimestamp;

    internal WordDeleteKeyHook(Action callback)
    {
        _callback = callback ?? throw new ArgumentNullException(nameof(callback));
    }

    internal void BindOwnerWindow(int windowHandle)
    {
        if (windowHandle == 0) return;
        var threadId = GetWindowThreadProcessId(new IntPtr(windowHandle), out var processId);
        if (threadId == 0 || processId == 0) return;
        if (_hook != IntPtr.Zero && _ownerThreadId != 0 && _ownerThreadId != threadId)
        {
            WordDoubleClickHook.TraceMessage(
                $"delete-key-hook-window-thread-changed old={_ownerThreadId} new={threadId}");
            return;
        }
        _ownerThreadId = threadId;
        _ownerProcessId = unchecked((int)processId);
    }

    internal void Start()
    {
        if (_hook != IntPtr.Zero) return;
        if (_ownerThreadId == 0)
            _ownerThreadId = GetCurrentThreadId();
        _hookCallback = HookCallback;
        _hook = SetWindowsHookEx(WhKeyboard, _hookCallback, IntPtr.Zero, _ownerThreadId);
        if (_hook == IntPtr.Zero)
            throw new Win32Exception(
                Marshal.GetLastWin32Error(),
                "Word delete-key observer could not be installed.");
        WordDoubleClickHook.TraceMessage(
            $"delete-key-hook-started handle=0x{_hook.ToInt64():X} thread={_ownerThreadId} process={_ownerProcessId}");
    }

    private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
    {
        if (code < 0)
            return CallNextHookEx(_hook, code, wParam, lParam);

        var virtualKey = unchecked((uint)wParam.ToInt64());
        // WH_KEYBOARD lParam bit 31 is 0 for key-down and 1 for key-up.
        var keyUp = (lParam.ToInt64() & 0x80000000L) != 0;
        if (keyUp || (virtualKey != VkDelete && virtualKey != VkBack))
            return CallNextHookEx(_hook, code, wParam, lParam);

        var now = System.Diagnostics.Stopwatch.GetTimestamp();
        var previous = Interlocked.Read(ref _lastDeleteObservationTimestamp);
        var duplicateWindow = System.Diagnostics.Stopwatch.Frequency / 10; // 100 ms
        if (previous == 0 || now - previous > duplicateWindow)
        {
            Interlocked.Exchange(ref _lastDeleteObservationTimestamp, now);
            WordDoubleClickHook.TraceMessage(
                $"delete-key-observed key={(virtualKey == VkDelete ? "Delete" : "Backspace")}");
            // This callback is intentionally limited to copying an already-verified
            // managed guard and scheduling a later Office-STA check. It performs no
            // Word COM call from inside the keyboard hook and never suppresses input.
            try { _callback(); }
            catch (Exception error)
            {
                WordDoubleClickHook.TraceMessage(
                    $"delete-key-callback-error {error.GetType().Name}:{error.Message}");
            }
        }
        return CallNextHookEx(_hook, code, wParam, lParam);
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

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(
        int idHook,
        HookProc callback,
        IntPtr module,
        uint threadId);

    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hook);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(
        IntPtr hook,
        int code,
        IntPtr wParam,
        IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();
}
