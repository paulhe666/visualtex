using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using Application = Microsoft.Office.Interop.Word.Application;

namespace VisualTeX.WordVsto;

/// <summary>
/// Word owns existing commands, custom bindings (including disabled bindings),
/// key-sequence prefixes, menus and non-document input. Never rewrites bindings.
/// FindKey is context-dependent: temporarily inspect document/template scopes,
/// then restore CustomizationContext. Unreadable does not mean unassigned.
/// </summary>
internal static class WordFormulaShortcutPolicy
{
    internal static bool CanExecute(Application? application,
        OfficeFormulaShortcut shortcut, OfficeFormulaShortcut? previous)
    {
        if (application is null) return false;
        try
        {
            return CanExecuteWithWordBindings(application, shortcut, previous,
                HasDocumentInputFocus(application));
        }
        catch { return false; }
    }

    internal static bool CanExecuteWithWordBindings(Application application,
        OfficeFormulaShortcut shortcut, OfficeFormulaShortcut? previous,
        bool documentInputFocused)
    {
        if (!CanExecuteInContext(application, shortcut, previous, documentInputFocused))
            return false;
        object? originalContext = null;
        var capturedContext = false;
        var allowed = false;
        try
        {
            originalContext = application.CustomizationContext;
            if (originalContext is null) return false;
            capturedContext = true;
            allowed = CheckDocumentAndTemplateContexts(application, shortcut, previous);
        }
        catch { allowed = false; }
        finally
        {
            if (capturedContext)
            {
                try { application.CustomizationContext = originalContext; }
                catch { allowed = false; }
            }
            Release(originalContext);
        }
        return allowed;
    }

    private static bool CheckDocumentAndTemplateContexts(Application application,
        OfficeFormulaShortcut shortcut, OfficeFormulaShortcut? previous)
    {
        Document? document = null;
        Template? attached = null;
        Template? normal = null;
        Templates? templates = null;
        try
        {
            document = application.ActiveDocument;
            if (document is null) return false;
            application.CustomizationContext = document;
            if (!CanExecuteInContext(application, shortcut, previous, true)) return false;
            attached = document.get_AttachedTemplate() as Template;
            if (attached is null) return false;
            application.CustomizationContext = attached;
            if (!CanExecuteInContext(application, shortcut, previous, true)) return false;
            normal = application.NormalTemplate;
            if (normal is null) return false;
            application.CustomizationContext = normal;
            if (!CanExecuteInContext(application, shortcut, previous, true)) return false;
            templates = application.Templates;
            for (var index = 1; index <= templates.Count; index++)
            {
                Template? template = null;
                try
                {
                    object templateIndex = index;
                    template = templates[ref templateIndex];
                    if (template.Type != WdTemplateType.wdGlobalTemplate) continue;
                    application.CustomizationContext = template;
                    if (!CanExecuteInContext(application, shortcut, previous, true)) return false;
                }
                finally { Release(template); }
            }
            return true;
        }
        finally
        {
            Release(templates);
            Release(normal);
            Release(attached);
            Release(document);
        }
    }

    internal static bool CanExecuteInContext(Application application,
        OfficeFormulaShortcut shortcut, OfficeFormulaShortcut? previous,
        bool documentInputFocused)
    {
        // Alt-only gestures can be Word Ribbon/menu accelerators without a
        // FindKey entry. Leave them to Word rather than guessing from a key list.
        if (!documentInputFocused || !shortcut.Ctrl) return false;
        try
        {
            if (!TryGetCategory(application, shortcut, out var category)
                || category != WdKeyCategory.wdKeyCategoryNil)
                return false;
            // Word supports two-stroke bindings. The first stroke may not be a
            // configured VisualTeX chord, so the hook remembers passed-through
            // nonmodifier input on the same focused document window.
            if (previous.HasValue
                && (!TryGetCategory(application, previous.Value, out var previousCategory)
                    || previousCategory == WdKeyCategory.wdKeyCategoryPrefix))
                return false;
            return true;
        }
        catch { return false; }
    }

    private static bool TryGetCategory(Application application,
        OfficeFormulaShortcut shortcut, out WdKeyCategory category)
    {
        category = WdKeyCategory.wdKeyCategoryNil;
        var key = shortcut.VirtualKey;
        // These are the WdKey values that match virtual-key values. Word's
        // documented function-key constants stop at F16.
        if (!((key >= 0x30 && key <= 0x39)
            || (key >= 0x41 && key <= 0x5A)
            || (key >= 0x70 && key <= 0x7F)
            || key == 0x1B || key == 0x09 || key == 0x0D
            || key == 0x20 || key == 0x08 || key == 0x2E
            || (key >= 0x21 && key <= 0x28)))
            return false;
        object control = shortcut.Ctrl ? (object)WdKey.wdKeyControl : Type.Missing;
        object alt = shortcut.Alt ? (object)WdKey.wdKeyAlt : Type.Missing;
        object shift = shortcut.Shift ? (object)WdKey.wdKeyShift : Type.Missing;
        var code = application.BuildKeyCode((WdKey)key, ref control, ref alt, ref shift);
        KeyBinding? binding = null;
        try
        {
            binding = application.FindKey[code];
            if (binding is null) return false;
            category = binding.KeyCategory;
            return true;
        }
        finally { Release(binding); }
    }

    private static bool HasDocumentInputFocus(Application application)
    {
        Window? window = null;
        try
        {
            var focus = GetFocus();
            if (focus == IntPtr.Zero) return false;
            var name = new StringBuilder(32);
            if (GetClassName(focus, name, name.Capacity) == 0
                || !string.Equals(name.ToString(), "_WwG", StringComparison.Ordinal))
                return false;
            window = application.ActiveWindow;
            if (window is null) return false;
            var owner = new IntPtr(window.Hwnd);
            var root = GetAncestor(owner, 2); // GA_ROOT
            if (root == IntPtr.Zero || root != GetForegroundWindow()
                || GetAncestor(focus, 2) != root
                || (focus != owner && !IsChild(owner, focus)))
                return false;
            var thread = GetWindowThreadProcessId(focus, out var process);
            return thread == GetCurrentThreadId() && process == GetCurrentProcessId();
        }
        finally { Release(window); }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }

    [DllImport("user32.dll")] private static extern IntPtr GetFocus();
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
    [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
    [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();
}
