param(
    [Parameter(Mandatory=$true)][int]$ProcessId,
    [Parameter(Mandatory=$true)][string]$DumpPath,
    [int]$TimeoutSeconds = 900
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

Add-Type -TypeDefinition @'
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

public static class P13NativeCrashCapture
{
    public const uint EXCEPTION_DEBUG_EVENT = 1;
    public const uint EXIT_PROCESS_DEBUG_EVENT = 5;
    public const uint DBG_CONTINUE = 0x00010002;
    public const uint DBG_EXCEPTION_NOT_HANDLED = 0x80010001;
    public const uint EXCEPTION_ILLEGAL_INSTRUCTION = 0xC000001D;

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool DebugActiveProcess(uint dwProcessId);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool DebugActiveProcessStop(uint dwProcessId);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool DebugSetProcessKillOnExit(bool KillOnExit);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool WaitForDebugEvent(IntPtr lpDebugEvent, uint dwMilliseconds);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool ContinueDebugEvent(uint dwProcessId, uint dwThreadId, uint dwContinueStatus);

    [DllImport("dbghelp.dll", SetLastError = true)]
    public static extern bool MiniDumpWriteDump(
        IntPtr hProcess,
        uint ProcessId,
        IntPtr hFile,
        uint DumpType,
        IntPtr ExceptionParam,
        IntPtr UserStreamParam,
        IntPtr CallbackParam);

    public static string Capture(int pid, string dumpPath, int timeoutSeconds)
    {
        var process = Process.GetProcessById(pid);
        if (!DebugActiveProcess((uint)pid))
            throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "DebugActiveProcess failed");
        DebugSetProcessKillOnExit(false);

        IntPtr ev = Marshal.AllocHGlobal(512);
        var deadline = DateTime.UtcNow.AddSeconds(timeoutSeconds);
        bool captured = false;
        try
        {
            while (DateTime.UtcNow < deadline)
            {
                if (!WaitForDebugEvent(ev, 500))
                {
                    int error = Marshal.GetLastWin32Error();
                    if (error == 121) continue; // ERROR_SEM_TIMEOUT
                    try { if (process.HasExited) break; } catch { break; }
                    continue;
                }

                uint eventCode = unchecked((uint)Marshal.ReadInt32(ev, 0));
                uint eventPid = unchecked((uint)Marshal.ReadInt32(ev, 4));
                uint eventTid = unchecked((uint)Marshal.ReadInt32(ev, 8));
                uint status = DBG_CONTINUE;

                if (eventCode == EXCEPTION_DEBUG_EVENT)
                {
                    // DEBUG_EVENT's union is 8-byte aligned on x64, so the
                    // EXCEPTION_RECORD begins at offset 16. ExceptionCode is first.
                    uint exceptionCode = unchecked((uint)Marshal.ReadInt32(ev, 16));
                    Console.WriteLine("DEBUG_EXCEPTION|pid={0}|tid={1}|code=0x{2:X8}", eventPid, eventTid, exceptionCode);
                    if (exceptionCode == EXCEPTION_ILLEGAL_INSTRUCTION && !captured)
                    {
                        var dumpDirectory = Path.GetDirectoryName(Path.GetFullPath(dumpPath));
                        if (String.IsNullOrEmpty(dumpDirectory))
                            throw new InvalidOperationException("Dump path has no parent directory.");
                        Directory.CreateDirectory(dumpDirectory);
                        using (var stream = new FileStream(dumpPath, FileMode.Create, FileAccess.ReadWrite, FileShare.Read))
                        {
                            // Full memory + handle data + thread info.
                            const uint dumpType = 0x00000002 | 0x00000004 | 0x00001000;
                            if (!MiniDumpWriteDump(process.Handle, (uint)pid, stream.SafeFileHandle.DangerousGetHandle(), dumpType,
                                IntPtr.Zero, IntPtr.Zero, IntPtr.Zero))
                                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "MiniDumpWriteDump failed");
                        }
                        captured = true;
                        Console.WriteLine("DUMP_CAPTURED|{0}", Path.GetFullPath(dumpPath));
                    }
                    // Let Windows exception dispatch continue normally after we
                    // have captured the first-chance state.
                    status = DBG_EXCEPTION_NOT_HANDLED;
                }

                if (!ContinueDebugEvent(eventPid, eventTid, status))
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "ContinueDebugEvent failed");

                if (eventCode == EXIT_PROCESS_DEBUG_EVENT)
                    break;
            }
        }
        finally
        {
            Marshal.FreeHGlobal(ev);
            try { DebugActiveProcessStop((uint)pid); } catch { }
        }

        return captured ? "CAPTURED" : "NO_TARGET_EXCEPTION";
    }
}
'@

$resolved = [IO.Path]::GetFullPath($DumpPath)
Write-Output "DEBUG_ATTACH|pid=$ProcessId|dump=$resolved|timeoutSeconds=$TimeoutSeconds"
$result = [P13NativeCrashCapture]::Capture($ProcessId, $resolved, $TimeoutSeconds)
Write-Output "DEBUG_RESULT|$result"
