using System.ComponentModel;
using System.Runtime.InteropServices;

namespace VisualTeX.WordVsto;

/// <summary>
/// Observes keyboard Paste before Word handles it. Unlike the clipboard observer,
/// this hook runs on Word's own UI thread and fires before the native Ctrl+V /
/// Shift+Insert command mutates the document. VisualTeX uses that narrow window
/// only to open a Word custom undo record; it never suppresses or replaces the
/// user's native paste command.
/// </summary>
internal sealed class WordPasteKeyHook : IDisposable
{
    private const int WhKeyboard = 2;
    private const uint VkV = 0x56;
    private const uint VkInsert = 0x2D;
    private const int VkControl = 0x11;
    private const int VkShift = 0x10;

    private readonly Action _callback;
    private HookProc? _hookCallback;
    private IntPtr _hook;
    private uint _ownerThreadId;
    private int _ownerProcessId;

    internal WordPasteKeyHook(Action callback)
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
                $"paste-key-hook-window-thread-changed old={_ownerThreadId} new={threadId}");
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
                "Word paste-key observer could not be installed.");
        WordDoubleClickHook.TraceMessage(
            $"paste-key-hook-started handle=0x{_hook.ToInt64():X} thread={_ownerThreadId} process={_ownerProcessId}");
    }

    private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
    {
        if (code < 0)
            return CallNextHookEx(_hook, code, wParam, lParam);

        var key = unchecked((uint)wParam.ToInt64());
        var state = lParam.ToInt64();
        var keyUp = (state & 0x80000000L) != 0;
        var repeated = (state & 0x40000000L) != 0;
        if (keyUp || repeated)
            return CallNextHookEx(_hook, code, wParam, lParam);

        var ctrlPaste = key == VkV && IsKeyDown(VkControl);
        var shiftInsertPaste = key == VkInsert && IsKeyDown(VkShift);
        if (!ctrlPaste && !shiftInsertPaste)
            return CallNextHookEx(_hook, code, wParam, lParam);

        WordDoubleClickHook.TraceMessage(
            $"paste-key-observed key={(ctrlPaste ? "Ctrl+V" : "Shift+Insert")}");
        try { _callback(); }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage(
                $"paste-key-callback-error {error.GetType().Name}:{error.Message}");
        }
        return CallNextHookEx(_hook, code, wParam, lParam);
    }

    private static bool IsKeyDown(int virtualKey) =>
        (GetKeyState(virtualKey) & 0x8000) != 0;

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

    [DllImport("user32.dll")]
    private static extern short GetKeyState(int virtualKey);
}
