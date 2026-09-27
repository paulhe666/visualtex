using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace VisualTeX.WindowsOffice.VstoShared;

internal enum OfficeFormulaShortcutHost
{
    Word,
    PowerPoint,
}

internal enum OfficeFormulaShortcutKind
{
    Inline,
    Display,
    PowerPointSvg,
    PowerPointOle,
    PowerPointInlineOmml,
    PowerPointDisplayOmml,
}

internal enum OfficeFormulaShortcutFormat
{
    Ole,
    Omml,
}

internal readonly struct OfficeFormulaShortcutCommand
{
    internal OfficeFormulaShortcutCommand(
        OfficeFormulaShortcutKind kind,
        OfficeFormulaShortcutFormat format = OfficeFormulaShortcutFormat.Ole)
    {
        Kind = kind;
        Format = format;
    }

    internal OfficeFormulaShortcutKind Kind { get; }
    internal OfficeFormulaShortcutFormat Format { get; }
}

internal readonly struct OfficeFormulaShortcut : IEquatable<OfficeFormulaShortcut>
{
    internal OfficeFormulaShortcut(uint virtualKey, bool ctrl, bool alt, bool shift)
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

    public bool Equals(OfficeFormulaShortcut other) =>
        VirtualKey == other.VirtualKey
        && Ctrl == other.Ctrl
        && Alt == other.Alt
        && Shift == other.Shift;

    public override bool Equals(object? value) =>
        value is OfficeFormulaShortcut other && Equals(other);

    public override int GetHashCode() =>
        (((int)VirtualKey * 397 ^ Ctrl.GetHashCode()) * 397 ^ Alt.GetHashCode())
        * 397 ^ Shift.GetHashCode();

    public static bool operator ==(OfficeFormulaShortcut left, OfficeFormulaShortcut right) =>
        left.Equals(right);

    public static bool operator !=(OfficeFormulaShortcut left, OfficeFormulaShortcut right) =>
        !left.Equals(right);

    internal static bool TryParse(string? value, out OfficeFormulaShortcut shortcut)
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
        shortcut = new OfficeFormulaShortcut(virtualKey, ctrl, alt, shift);
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

internal sealed class OfficeFormulaShortcutConfiguration
{
    private const string RegistryPath = @"Software\VisualTeX\OfficeIntegration";

    private readonly OfficeFormulaShortcutBinding[] _bindings;

    private OfficeFormulaShortcutConfiguration(
        OfficeFormulaShortcutBinding[] bindings) => _bindings = bindings;

    internal static OfficeFormulaShortcutConfiguration Read(
        OfficeFormulaShortcutHost host)
    {
        var valuePrefix = host == OfficeFormulaShortcutHost.Word
            ? "Word"
            : "PowerPoint";
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
                var bindings = host == OfficeFormulaShortcutHost.Word
                    ? ReadWordBindings(key, valuePrefix)
                    : ReadPowerPointBindings(key, valuePrefix);
                if (bindings is null) continue;
                return new OfficeFormulaShortcutConfiguration(bindings);
            }
            catch
            {
                // Missing or inaccessible settings mean the shortcuts are disabled.
            }
        }
        return new OfficeFormulaShortcutConfiguration(
            Array.Empty<OfficeFormulaShortcutBinding>());
    }

    internal OfficeFormulaShortcutCommand? Match(
        uint virtualKey,
        bool ctrl,
        bool alt,
        bool shift)
    {
        foreach (var binding in _bindings)
        {
            if (binding.Shortcut.Matches(virtualKey, ctrl, alt, shift))
                return binding.Command;
        }
        return null;
    }

    private static OfficeFormulaShortcutBinding[]? ReadWordBindings(
        RegistryKey key,
        string valuePrefix)
    {
        var inlineValue = ReadString(key, valuePrefix + "InlineFormulaShortcut");
        var displayValue = ReadString(key, valuePrefix + "DisplayFormulaShortcut");
        if (inlineValue is null && displayValue is null) return null;
        return CreateDistinctBindings(
            (inlineValue, new OfficeFormulaShortcutCommand(
                OfficeFormulaShortcutKind.Inline,
                ReadFormat(ReadString(key, valuePrefix + "InlineFormulaFormat")))),
            (displayValue, new OfficeFormulaShortcutCommand(
                OfficeFormulaShortcutKind.Display,
                ReadFormat(ReadString(key, valuePrefix + "DisplayFormulaFormat")))));
    }

    private static OfficeFormulaShortcutBinding[]? ReadPowerPointBindings(
        RegistryKey key,
        string valuePrefix)
    {
        var svg = ReadString(key, valuePrefix + "SvgFormulaShortcut");
        var ole = ReadString(key, valuePrefix + "OleFormulaShortcut");
        var inlineOmml = ReadString(key, valuePrefix + "InlineOmmlFormulaShortcut");
        var displayOmml = ReadString(key, valuePrefix + "DisplayOmmlFormulaShortcut");
        if (svg is null && ole is null && inlineOmml is null && displayOmml is null)
            return null;
        return CreateDistinctBindings(
            (svg, new OfficeFormulaShortcutCommand(OfficeFormulaShortcutKind.PowerPointSvg)),
            (ole, new OfficeFormulaShortcutCommand(OfficeFormulaShortcutKind.PowerPointOle)),
            (inlineOmml, new OfficeFormulaShortcutCommand(
                OfficeFormulaShortcutKind.PowerPointInlineOmml)),
            (displayOmml, new OfficeFormulaShortcutCommand(
                OfficeFormulaShortcutKind.PowerPointDisplayOmml)));
    }

    private static OfficeFormulaShortcutBinding[] CreateDistinctBindings(
        params (string? Value, OfficeFormulaShortcutCommand Command)[] candidates)
    {
        var bindings = new List<OfficeFormulaShortcutBinding>();
        foreach (var candidate in candidates)
        {
            if (!OfficeFormulaShortcut.TryParse(candidate.Value, out var shortcut)
                || bindings.Exists(binding => binding.Shortcut == shortcut))
                continue;
            bindings.Add(new OfficeFormulaShortcutBinding(shortcut, candidate.Command));
        }
        return bindings.ToArray();
    }

    private static string? ReadString(RegistryKey key, string name) =>
        key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames) as string;

    private static OfficeFormulaShortcutFormat ReadFormat(string? value) =>
        string.Equals(value, "omml", StringComparison.OrdinalIgnoreCase)
            ? OfficeFormulaShortcutFormat.Omml
            : OfficeFormulaShortcutFormat.Ole;

    private readonly struct OfficeFormulaShortcutBinding
    {
        internal OfficeFormulaShortcutBinding(
            OfficeFormulaShortcut shortcut,
            OfficeFormulaShortcutCommand command)
        {
            Shortcut = shortcut;
            Command = command;
        }

        internal OfficeFormulaShortcut Shortcut { get; }
        internal OfficeFormulaShortcutCommand Command { get; }
    }
}

/// <summary>
/// Routes configured formula gestures on one Office UI thread. Word can also
/// observe paste through the same hook; PowerPoint leaves that optional callback
/// unset. Each host reads only its own registry values, so identical gestures can
/// safely be assigned across Word and PowerPoint while duplicates inside one host
/// remain impossible.
/// </summary>
internal sealed class OfficeFormulaShortcutKeyboardHook : IDisposable
{
    private const int WhKeyboard = 2;
    private const uint VkV = 0x56;
    private const uint VkInsert = 0x2D;
    private const int VkControl = 0x11;
    private const int VkShift = 0x10;
    private const int VkMenu = 0x12;
    private const int VkLeftWindows = 0x5B;
    private const int VkRightWindows = 0x5C;

    private readonly OfficeFormulaShortcutHost _host;
    private readonly Action<OfficeFormulaShortcutCommand> _formulaCallback;
    private readonly Action? _pasteCallback;
    private readonly Action<string>? _trace;
    private HookProc? _hookCallback;
    private IntPtr _hook;
    private uint _ownerThreadId;
    private int _ownerProcessId;

    internal OfficeFormulaShortcutKeyboardHook(
        OfficeFormulaShortcutHost host,
        Action<OfficeFormulaShortcutCommand> formulaCallback,
        Action? pasteCallback = null,
        Action<string>? trace = null)
    {
        _host = host;
        _formulaCallback = formulaCallback
            ?? throw new ArgumentNullException(nameof(formulaCallback));
        _pasteCallback = pasteCallback;
        _trace = trace;
    }

    internal void BindOwnerWindow(int windowHandle)
    {
        if (windowHandle == 0) return;
        var threadId = GetWindowThreadProcessId(new IntPtr(windowHandle), out var processId);
        if (threadId == 0 || processId == 0) return;
        if (_hook != IntPtr.Zero && _ownerThreadId != 0 && _ownerThreadId != threadId)
        {
            Trace($"keyboard-hook-window-thread-changed old={_ownerThreadId} new={threadId}");
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
            throw new Win32Exception(
                Marshal.GetLastWin32Error(),
                $"{_host} keyboard observer could not be installed.");
        Trace(
            $"keyboard-hook-started host={_host} handle=0x{_hook.ToInt64():X} "
            + $"thread={_ownerThreadId} process={_ownerProcessId}");
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
            var command = OfficeFormulaShortcutConfiguration.Read(_host)
                .Match(key, ctrl, alt, shift);
            if (command.HasValue)
            {
                try { _formulaCallback(command.Value); }
                catch (Exception error)
                {
                    Trace(
                        $"formula-shortcut-callback-error {error.GetType().Name}:{error.Message}");
                }
                return new IntPtr(1);
            }
        }

        if (_pasteCallback is null)
            return CallNextHookEx(_hook, code, wParam, lParam);
        var ctrlPaste = key == VkV && ctrl;
        var shiftInsertPaste = key == VkInsert && shift;
        if (!ctrlPaste && !shiftInsertPaste)
            return CallNextHookEx(_hook, code, wParam, lParam);
        try { _pasteCallback(); }
        catch (Exception error)
        {
            Trace($"paste-key-callback-error {error.GetType().Name}:{error.Message}");
        }
        return CallNextHookEx(_hook, code, wParam, lParam);
    }

    private void Trace(string message)
    {
        try { _trace?.Invoke(message); } catch { }
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
