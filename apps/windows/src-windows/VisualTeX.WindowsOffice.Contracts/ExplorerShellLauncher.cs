using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace VisualTeX.WindowsOffice.Contracts;

/// <summary>
/// Starts a program through the user's explorer.exe (IShellDispatch2.ShellExecute on the
/// desktop folder view), so the new process is explorer's child rather than the host's.
/// WPS Office denies CreateProcess and in-process ShellExecute to add-ins running inside
/// wps.exe/wpp.exe; this is how an add-in there can still start the VisualTeX companion.
/// </summary>
public static class ExplorerShellLauncher
{
    private static readonly Guid ShellWindowsClsid = new("9BA05972-F6A8-11CF-A442-00A0C90A8F39");
    private static readonly Guid TopLevelBrowserService = new("4C96BE40-915C-11CF-99D3-00AA004AE837");
    private static readonly Guid ShellBrowserIid = new("000214E2-0000-0000-C000-000000000046");
    private static readonly Guid DispatchIid = new("00020400-0000-0000-C000-000000000046");
    private const int CsidlDesktop = 0x0008;
    private const int SwcDesktop = 8;
    private const int SwfoNeedDispatch = 1;
    private const uint SvgioBackground = 0;

    public static void Start(string file, string arguments, string workingDirectory)
    {
        object? shellWindows = null;
        object? desktop = null;
        object? browserObject = null;
        IShellView? view = null;
        object? folderView = null;
        object? shell = null;
        try
        {
            var type = Type.GetTypeFromCLSID(ShellWindowsClsid, throwOnError: true)!;
            shellWindows = Activator.CreateInstance(type)!;
            // IShellWindows.FindWindowSW through IDispatch (the early-bound vtable call
            // finds no window): location and root by reference, hwnd out.
            var arguments2 = new object[] { CsidlDesktop, new object(), SwcDesktop, 0, SwfoNeedDispatch };
            var byReference = new ParameterModifier(5);
            byReference[0] = true;
            byReference[1] = true;
            byReference[3] = true;
            desktop = shellWindows.GetType().InvokeMember(
                    "FindWindowSW",
                    BindingFlags.InvokeMethod,
                    null,
                    shellWindows,
                    arguments2,
                    new[] { byReference },
                    null,
                    null)
                ?? throw new InvalidOperationException("Explorer exposes no desktop window.");
            var topLevelBrowser = TopLevelBrowserService;
            var shellBrowserIid = ShellBrowserIid;
            ((IServiceProvider)desktop).QueryService(ref topLevelBrowser, ref shellBrowserIid, out browserObject);
            ((IShellBrowser)browserObject).QueryActiveShellView(out view);
            var dispatchIid = DispatchIid;
            view.GetItemObject(SvgioBackground, ref dispatchIid, out folderView);
            shell = folderView.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, folderView, null)
                ?? throw new InvalidOperationException("Explorer's desktop view exposes no shell application.");
            shell.GetType().InvokeMember(
                "ShellExecute",
                BindingFlags.InvokeMethod,
                null,
                shell,
                new object[] { file, arguments, workingDirectory, "open", 0 });
        }
        finally
        {
            Release(shell);
            Release(folderView);
            Release(view);
            Release(browserObject);
            Release(desktop);
            Release(shellWindows);
        }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }

    [ComImport, Guid("6D5140C1-7436-11CE-8034-00AA006009FA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IServiceProvider
    {
        void QueryService(ref Guid service, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object result);
    }

    [ComImport, Guid("000214E2-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellBrowser
    {
        void GetWindow(out IntPtr hwnd);
        void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool enterMode);
        void InsertMenusSB(IntPtr shared, IntPtr widths);
        void SetMenuSB(IntPtr shared, IntPtr oleMenu, IntPtr activeObject);
        void RemoveMenusSB(IntPtr shared);
        void SetStatusTextSB(IntPtr text);
        void EnableModelessSB([MarshalAs(UnmanagedType.Bool)] bool enable);
        void TranslateAcceleratorSB(IntPtr message, ushort id);
        void BrowseObject(IntPtr pidl, uint flags);
        void GetViewStateStream(uint mode, out IntPtr stream);
        void GetControlWindow(uint id, out IntPtr hwnd);
        void SendControlMsg(uint id, uint message, IntPtr wParam, IntPtr lParam, out IntPtr result);
        void QueryActiveShellView(out IShellView view);
        void OnViewWindowActive(IShellView view);
        void SetToolbarItems(IntPtr buttons, uint count, uint flags);
    }

    [ComImport, Guid("000214E3-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellView
    {
        void GetWindow(out IntPtr hwnd);
        void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool enterMode);
        void TranslateAccelerator(IntPtr message);
        void EnableModeless([MarshalAs(UnmanagedType.Bool)] bool enable);
        void UIActivate(uint state);
        void Refresh();
        void CreateViewWindow(IShellView previous, IntPtr settings, IntPtr browser, IntPtr rect, out IntPtr hwnd);
        void DestroyViewWindow();
        void GetCurrentInfo(IntPtr settings);
        void AddPropertySheetPages(uint reserved, IntPtr callback, IntPtr parameter);
        void SaveViewState();
        void SelectItem(IntPtr pidl, uint flags);
        void GetItemObject(uint item, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object result);
    }
}
