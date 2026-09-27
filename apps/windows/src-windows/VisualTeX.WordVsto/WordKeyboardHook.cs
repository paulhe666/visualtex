using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace VisualTeX.WordVsto;

internal enum WordFormulaShortcutKind
{
    Inline,
    Display,
}

internal enum WordFormulaShortcutFormat
{
    Ole,
    Omml,
}

internal readonly struct WordFormulaShortcutCommand
{
    internal WordFormulaShortcutCommand(
        WordFormulaShortcutKind kind,
        WordFormulaShortcutFormat format)
    {
        Kind = kind;
        Format = format;
    }

    internal WordFormulaShortcutKind Kind { get; }
    internal WordFormulaShortcutFormat Format { get; }
}

internal readonly struct WordFormulaShortcut : IEquatable<WordFormulaShortcut>
{
    internal WordFormulaShortcut(uint virtualKey, bool ctrl, bool alt, bool shift)
    {
        VirtualKey = virtualKey;
        Ctrl = ctrl;
        Alt = alt;
        Shift = shift;
    }

    internal uint VirtualKey { get; }
    internal bool Ctrl { get; }
    internal bool Alt { get; }
    internal bool Shift { get; }

    internal bool Matches(uint virtualKey, bool ctrl, bool alt, bool shift) =>
        VirtualKey == virtualKey && Ctrl == ctrl && Alt == alt && Shift == shift;

    public bool Equals(WordFormulaShortcut other) =>
        VirtualKey == other.VirtualKey
        && Ctrl == other.Ctrl
        && Alt == other.Alt
        && Shift == other.Shift;

    public override bool Equals(object? value) =>
        value is WordFormulaShortcut other && Equals(other);

    public override int GetHashCode() =>
        (((int)VirtualKey * 397 ^ Ctrl.GetHashCode()) * 397 ^ Alt.GetHashCode())
        * 397 ^ Shift.GetHashCode();

    public static bool operator ==(WordFormulaShortcut left, WordFormulaShortcut right) =>
        left.Equals(right);

    public static bool operator !=(WordFormulaShortcut left, WordFormulaShortcut right) =>
        !left.Equals(right);

    internal static bool TryParse(string? value, out WordFormulaShortcut shortcut)
    {
        shortcut = default;
        if (string.IsNullOrWhiteSpace(value)) return false;
        var parts = value!.Split('+');
        if (parts.Length < 2) return false;
        var ctrl = false;
        var alt = false;
        var shift = false;
        for (var index = 0; index < parts.Length - 1; index += 1)
        {
            switch (parts[index].Trim().ToUpperInvariant())
            {
                case "CTRL" when !ctrl: ctrl = true; break;
                case "ALT" when !alt: alt = true; break;
                case "SHIFT" when !shift: shift = true; break;
                default: return false;
            }
        }
        if (!ctrl && !alt) return false;
        if (!TryParseVirtualKey(parts[parts.Length - 1].Trim(), out var virtualKey))
            return false;
        if (ctrl && alt && !shift && virtualKey == 0x4F) return false;
        shortcut = new WordFormulaShortcut(virtualKey, ctrl, alt, shift);
        return true;
    }

    private static bool TryParseVirtualKey(string value, out uint virtualKey)
    {
        virtualKey = 0;
        var key = value.ToUpperInvariant();
        if (key.Length == 1)
        {
            var character = key[0];
            if ((character >= 'A' && character <= 'Z')
                || (character >= '0' && character <= '9'))
            {
                virtualKey = character;
                return true;
            }
            return false;
        }
        if (key[0] != 'F'
            || !int.TryParse(key.Substring(1), out var functionNumber)
            || functionNumber < 1
            || functionNumber > 24)
            return false;
        virtualKey = unchecked((uint)(0x70 + functionNumber - 1));
        return true;
    }
}

internal sealed class WordFormulaShortcutConfiguration
{
    private const string RegistryPath = @"Software\VisualTeX\OfficeIntegration";

    private WordFormulaShortcutConfiguration(
        WordFormulaShortcut? inline,
        WordFormulaShortcutFormat inlineFormat,
        WordFormulaShortcut? display,
        WordFormulaShortcutFormat displayFormat)
    {
        Inline = inline;
        InlineFormat = inlineFormat;
        Display = display;
        DisplayFormat = displayFormat;
    }

    internal WordFormulaShortcut? Inline { get; }
    internal WordFormulaShortcutFormat InlineFormat { get; }
    internal WordFormulaShortcut? Display { get; }
    internal WordFormulaShortcutFormat DisplayFormat { get; }

    internal static WordFormulaShortcutConfiguration Read()
    {
        var views = Environment.Is64BitOperatingSystem
            ? new[] { RegistryView.Registry64, RegistryView.Registry32 }
            : new[] { RegistryView.Default };
        foreach (var view in views)
        {
            try
            {
                using var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
                using var key = root.OpenSubKey(RegistryPath, writable: false);
                if (key is null) continue;
                var inlineValue = ReadString(key, "WordInlineFormulaShortcut");
                var displayValue = ReadString(key, "WordDisplayFormulaShortcut");
                if (inlineValue is null && displayValue is null) continue;
                WordFormulaShortcut? inline =
                    WordFormulaShortcut.TryParse(inlineValue, out var inlineShortcut)
                        ? inlineShortcut
                        : null;
                WordFormulaShortcut? display =
                    WordFormulaShortcut.TryParse(displayValue, out var displayShortcut)
                        ? displayShortcut
                        : null;
                if (inline.HasValue && display.HasValue && inline.Value == display.Value)
                    display = null;
                return new WordFormulaShortcutConfiguration(
                    inline,
                    ReadFormat(ReadString(key, "WordInlineFormulaFormat")),
                    display,
                    ReadFormat(ReadString(key, "WordDisplayFormulaFormat")));
            }
            catch
            {
                // Missing or inaccessible settings mean the shortcuts are disabled.
            }
        }
        return new WordFormulaShortcutConfiguration(
            null,
            WordFormulaShortcutFormat.Ole,
            null,
            WordFormulaShortcutFormat.Ole);
    }

    internal WordFormulaShortcutCommand? Match(
        uint virtualKey,
        bool ctrl,
        bool alt,
        bool shift)
    {
        if (Inline?.Matches(virtualKey, ctrl, alt, shift) == true)
            return new WordFormulaShortcutCommand(
                WordFormulaShortcutKind.Inline,
                InlineFormat);
        if (Display?.Matches(virtualKey, ctrl, alt, shift) == true)
            return new WordFormulaShortcutCommand(
                WordFormulaShortcutKind.Display,
                DisplayFormat);
        return null;
    }

    private static string? ReadString(RegistryKey key, string name) =>
        key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames) as string;

    private static WordFormulaShortcutFormat ReadFormat(string? value) =>
        string.Equals(value, "omml", StringComparison.OrdinalIgnoreCase)
            ? WordFormulaShortcutFormat.Omml
            : WordFormulaShortcutFormat.Ole;
}

/// <summary>
/// Routes Word UI-thread keyboard gestures through one native hook. Configured
/// formula shortcuts are consumed; native paste is only observed before Word
/// handles it so VisualTeX can open the matching custom undo record.
/// </summary>
internal sealed class WordKeyboardHook : IDisposable
{
    private const int WhKeyboard = 2;
    private const uint VkV = 0x56;
    private const uint VkInsert = 0x2D;
    private const int VkControl = 0x11;
    private const int VkShift = 0x10;
    private const int VkMenu = 0x12;
    private const int VkLeftWindows = 0x5B;
    private const int VkRightWindows = 0x5C;

    private readonly Action _pasteCallback;
    private readonly Action<WordFormulaShortcutCommand> _formulaCallback;
    private HookProc? _hookCallback;
    private IntPtr _hook;
    private uint _ownerThreadId;
    private int _ownerProcessId;

    internal WordKeyboardHook(
        Action pasteCallback,
        Action<WordFormulaShortcutCommand> formulaCallback)
    {
        _pasteCallback = pasteCallback ?? throw new ArgumentNullException(nameof(pasteCallback));
        _formulaCallback = formulaCallback ?? throw new ArgumentNullException(nameof(formulaCallback));
    }

    internal void BindOwnerWindow(int windowHandle)
    {
        if (windowHandle == 0) return;
        var threadId = GetWindowThreadProcessId(new IntPtr(windowHandle), out var processId);
        if (threadId == 0 || processId == 0) return;
        if (_hook != IntPtr.Zero && _ownerThreadId != 0 && _ownerThreadId != threadId)
        {
            WordDoubleClickHook.TraceMessage(
                $"keyboard-hook-window-thread-changed old={_ownerThreadId} new={threadId}");
            return;
        }
        _ownerThreadId = threadId;
        _ownerProcessId = unchecked((int)processId);
    }

    internal void Start()
    {
        if (_hook != IntPtr.Zero) return;
        if (_ownerThreadId == 0) _ownerThreadId = GetCurrentThreadId();
        _hookCallback = HookCallback;
        _hook = SetWindowsHookEx(WhKeyboard, _hookCallback, IntPtr.Zero, _ownerThreadId);
        if (_hook == IntPtr.Zero)
            throw new Win32Exception(Marshal.GetLastWin32Error(),
                "Word keyboard observer could not be installed.");
        WordDoubleClickHook.TraceMessage(
            $"keyboard-hook-started handle=0x{_hook.ToInt64():X} thread={_ownerThreadId} process={_ownerProcessId}");
    }

    private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
    {
        if (code < 0) return CallNextHookEx(_hook, code, wParam, lParam);
        var state = lParam.ToInt64();
        if ((state & 0xC0000000L) != 0)
            return CallNextHookEx(_hook, code, wParam, lParam);

        var key = unchecked((uint)wParam.ToInt64());
        var ctrl = IsKeyDown(VkControl);
        var alt = IsKeyDown(VkMenu);
        var shift = IsKeyDown(VkShift);
        if ((ctrl || alt)
            && !IsKeyDown(VkLeftWindows)
            && !IsKeyDown(VkRightWindows))
        {
            var command = WordFormulaShortcutConfiguration.Read()
                .Match(key, ctrl, alt, shift);
            if (command.HasValue)
            {
                try { _formulaCallback(command.Value); }
                catch (Exception error)
                {
                    WordDoubleClickHook.TraceMessage(
                        $"formula-shortcut-callback-error {error.GetType().Name}:{error.Message}");
                }
                return new IntPtr(1);
            }
        }

        var ctrlPaste = key == VkV && ctrl;
        var shiftInsertPaste = key == VkInsert && shift;
        if (!ctrlPaste && !shiftInsertPaste)
            return CallNextHookEx(_hook, code, wParam, lParam);
        try { _pasteCallback(); }
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
