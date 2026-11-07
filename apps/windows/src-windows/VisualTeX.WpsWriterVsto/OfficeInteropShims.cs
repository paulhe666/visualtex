using System.Runtime.InteropServices;

namespace VisualTeX.WpsWriterVsto;

// Office type libraries are not installable on build agents. These are the
// stable, documented COM definitions the add-in implements; keeping them in
// source lets the project restore and compile on a clean Windows runner.

[ComImport]
[Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IDTExtensibility2
{
    [DispId(1)]
    void OnConnection(
        [MarshalAs(UnmanagedType.IDispatch)] object Application,
        ext_ConnectMode ConnectMode,
        [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)]
        ref Array custom);

    [DispId(2)]
    void OnDisconnection(
        ext_DisconnectMode RemoveMode,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)]
        ref Array custom);

    [DispId(3)]
    void OnAddInsUpdate(
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)]
        ref Array custom);

    [DispId(4)]
    void OnStartupComplete(
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)]
        ref Array custom);

    [DispId(5)]
    void OnBeginShutdown(
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)]
        ref Array custom);
}

public enum ext_ConnectMode
{
    ext_cm_AfterStartup = 0,
    ext_cm_Startup = 1,
    ext_cm_External = 2,
    ext_cm_CommandLine = 3,
    ext_cm_Solution = 4,
    ext_cm_UISetup = 5,
}

public enum ext_DisconnectMode
{
    ext_dm_HostShutdown = 0,
    ext_dm_UserClosed = 1,
    ext_dm_UISetupComplete = 2,
    ext_dm_SolutionClosed = 3,
}

[ComImport]
[Guid("000C0396-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IRibbonExtensibility
{
    [DispId(1)]
    string GetCustomUI(string RibbonID);
}
