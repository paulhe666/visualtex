using System;
using System.ComponentModel;
using System.IO;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32.SafeHandles;

namespace VisualTeX.WindowsOffice.Contracts;

public static class NativeRenderClient
{
    // Read-only SCM check: Office neither starts nor stops a native engine.
    public static bool IsAvailable => GetServiceProcessId() != 0;

    public static async Task<NativeRenderResponse> RenderAsync(NativeRenderRequest request, CancellationToken token)
    {
        NativeRenderProtocol.Validate(request);
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(token);
        deadline.CancelAfter(NativeRenderProtocol.RequestTimeoutMs);
        using var pipe = await ConnectAsync(deadline.Token).ConfigureAwait(false);
        // Authenticate before sending equation data. A local process cannot
        // obtain or edit this SCM service registration without administrator rights.
        if (!GetNamedPipeServerProcessId(pipe.SafePipeHandle, out uint serverId)
            || serverId == 0 || serverId != GetServiceProcessId())
            throw new IOException("The native-render endpoint is not the registered VisualTeX service.");
        await NativeRenderProtocol.WriteAsync(pipe, request,
            NativeRenderProtocol.MaxRequestBytes, deadline.Token).ConfigureAwait(false);
        var result = await NativeRenderProtocol.ReadAsync<NativeRenderResponse>(pipe,
            NativeRenderProtocol.MaxResponseBytes, deadline.Token).ConfigureAwait(false);
        NativeRenderProtocol.Validate(result, request);
        // A pipe Disconnect discards unread server buffers. Acknowledge the full
        // validated frame before the server reuses its one persistent instance.
        await pipe.WriteAsync(new byte[] { 1 }, 0, 1, deadline.Token).ConfigureAwait(false);
        return result;
    }

    private static async Task<NamedPipeClientStream> ConnectAsync(CancellationToken token)
    {
        while (true)
        {
            token.ThrowIfCancellationRequested();
            // Explicit rights exclude FILE_CREATE_PIPE_INSTANCE, which Windows
            // includes in GENERIC_WRITE. SQOS Identification prevents impersonation.
            var handle = CreateFile(@"\\.\pipe\" + NativeRenderProtocol.PipeName,
                0x12019b, 0, IntPtr.Zero, 3, 0x40110000, IntPtr.Zero);
            if (!handle.IsInvalid)
            {
                try { return new NamedPipeClientStream(PipeDirection.InOut, true, true, handle); }
                catch { handle.Dispose(); throw; }
            }
            int error = Marshal.GetLastWin32Error();
            handle.Dispose();
            if (error != 231 && error != 2) throw new Win32Exception(error);
            if (GetServiceProcessId() == 0) throw new IOException("The VisualTeX render service is unavailable.");
            await Task.Delay(100, token).ConfigureAwait(false);
        }
    }

    private static uint GetServiceProcessId()
    {
        IntPtr manager = OpenSCManager(null, null, 1);
        if (manager == IntPtr.Zero) return 0;
        try
        {
            IntPtr service = OpenService(manager, NativeRenderProtocol.ServiceName, 4);
            if (service == IntPtr.Zero) return 0;
            try
            {
                return QueryServiceStatusEx(service, 0, out ServiceStatus status,
                    Marshal.SizeOf(typeof(ServiceStatus)), out _) && status.State == 4
                    ? status.ProcessId : 0;
            }
            finally { CloseServiceHandle(service); }
        }
        finally { CloseServiceHandle(manager); }
    }

    [StructLayout(LayoutKind.Sequential)] private struct ServiceStatus
    { public uint Type, State, Controls, Win32Exit, ServiceExit, Checkpoint, WaitHint, ProcessId, Flags; }
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr OpenSCManager(string? machine, string? database, uint access);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr OpenService(IntPtr manager, string name, uint access);
    [DllImport("advapi32.dll", SetLastError = true)] private static extern bool QueryServiceStatusEx(IntPtr service, int kind, out ServiceStatus status, int size, out int needed);
    [DllImport("advapi32.dll")] private static extern bool CloseServiceHandle(IntPtr service);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool GetNamedPipeServerProcessId(SafePipeHandle pipe, out uint pid);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern SafePipeHandle CreateFile(string path, uint access, uint share, IntPtr security, uint creation, uint flags, IntPtr template);
}
