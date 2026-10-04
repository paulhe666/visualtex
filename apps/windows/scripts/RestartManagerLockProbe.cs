using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

internal static class RestartManagerLockProbe
{
    private const int CCH_RM_SESSION_KEY = 32;
    private const int CCH_RM_MAX_APP_NAME = 255;
    private const int CCH_RM_MAX_SVC_NAME = 63;
    private const int ERROR_MORE_DATA = 234;

    [StructLayout(LayoutKind.Sequential)]
    private struct RM_UNIQUE_PROCESS
    {
        public int dwProcessId;
        public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime;
    }

    private enum RM_APP_TYPE
    {
        RmUnknownApp = 0,
        RmMainWindow = 1,
        RmOtherWindow = 2,
        RmService = 3,
        RmExplorer = 4,
        RmConsole = 5,
        RmCritical = 1000
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct RM_PROCESS_INFO
    {
        public RM_UNIQUE_PROCESS Process;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_APP_NAME + 1)]
        public string strAppName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_SVC_NAME + 1)]
        public string strServiceShortName;
        public RM_APP_TYPE ApplicationType;
        public uint AppStatus;
        public uint TSSessionId;
        [MarshalAs(UnmanagedType.Bool)]
        public bool bRestartable;
    }

    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    private static extern int RmStartSession(out uint pSessionHandle, int dwSessionFlags, string strSessionKey);

    [DllImport("rstrtmgr.dll")]
    private static extern int RmEndSession(uint pSessionHandle);

    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    private static extern int RmRegisterResources(
        uint pSessionHandle,
        uint nFiles,
        string[] rgsFileNames,
        uint nApplications,
        RM_UNIQUE_PROCESS[]? rgApplications,
        uint nServices,
        string[]? rgsServiceNames);

    [DllImport("rstrtmgr.dll")]
    private static extern int RmGetList(
        uint dwSessionHandle,
        out uint pnProcInfoNeeded,
        ref uint pnProcInfo,
        [In, Out] RM_PROCESS_INFO[]? rgAffectedApps,
        ref uint lpdwRebootReasons);

    public static int Main(string[] args)
    {
        if (args.Length != 1) return 2;
        var path = System.IO.Path.GetFullPath(args[0]);
        var key = Guid.NewGuid().ToString("N").Substring(0, CCH_RM_SESSION_KEY);
        var result = RmStartSession(out var handle, 0, key);
        if (result != 0) { Console.Error.WriteLine("RmStartSession=" + result); return 3; }
        try
        {
            result = RmRegisterResources(handle, 1, new[] { path }, 0, null, 0, null);
            if (result != 0) { Console.Error.WriteLine("RmRegisterResources=" + result); return 4; }
            uint needed = 0, count = 0, reboot = 0;
            result = RmGetList(handle, out needed, ref count, null, ref reboot);
            if (result == 0 && needed == 0)
            {
                Console.WriteLine("NO_LOCKING_PROCESSES");
                return 0;
            }
            if (result != ERROR_MORE_DATA)
            {
                Console.Error.WriteLine("RmGetList(initial)=" + result + " needed=" + needed);
                return 5;
            }
            var infos = new RM_PROCESS_INFO[needed];
            count = needed;
            result = RmGetList(handle, out needed, ref count, infos, ref reboot);
            if (result != 0) { Console.Error.WriteLine("RmGetList=" + result); return 6; }
            for (var i = 0; i < count; i++)
            {
                var pid = infos[i].Process.dwProcessId;
                string procName = "";
                string procPath = "";
                try
                {
                    using var p = Process.GetProcessById(pid);
                    procName = p.ProcessName;
                    try { procPath = p.MainModule?.FileName ?? ""; } catch { }
                }
                catch { }
                Console.WriteLine($"PID={pid} APP={infos[i].strAppName} PROC={procName} PATH={procPath} TYPE={infos[i].ApplicationType} RESTARTABLE={infos[i].bRestartable}");
            }
            return 0;
        }
        finally { RmEndSession(handle); }
    }
}
