using System.ComponentModel;
using System.Diagnostics;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.ServiceProcess;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.NativeRenderService;

internal sealed class RenderService : ServiceBase
{
    private readonly CancellationTokenSource stop = new();
    private NamedPipeServerStream? pipe;
    private Task? listener;
    private readonly string data = Path.Combine(AppContext.BaseDirectory, "data");

    private RenderService() { ServiceName = NativeRenderProtocol.ServiceName; AutoLog = false; }

    [MTAThread]
    private static void Main() => ServiceBase.Run(new RenderService());

    protected override void OnStart(string[] args)
    {
        VerifyPrivateIdentity();
        if (!Directory.Exists(data)) throw new InvalidOperationException("Private render data directory is not installed.");
        // Reserve the name before reporting started. A pre-existing server is a
        // hard failure, not permission to connect to somebody else's endpoint.
        pipe = CreateLocalPipe();
        listener = ListenAsync(pipe, stop.Token);
    }

    protected override void OnStop()
    {
        stop.Cancel();
        pipe?.Dispose();
        try { listener?.Wait(5000); } catch (AggregateException) { }
    }

    private async Task ListenAsync(NamedPipeServerStream endpoint, CancellationToken token)
    {
        using var backend = new NativeWorkerBackend(data);
        while (!token.IsCancellationRequested)
        {
            try
            {
                await endpoint.WaitForConnectionAsync(token).ConfigureAwait(false);
                NativeRenderRequest request;
                using (var read = CancellationTokenSource.CreateLinkedTokenSource(token))
                {
                    read.CancelAfter(5000);
                    request = await NativeRenderProtocol.ReadAsync<NativeRenderRequest>(
                        endpoint, NativeRenderProtocol.MaxRequestBytes, read.Token).ConfigureAwait(false);
                }
                NativeRenderProtocol.Validate(request);
                NativeRenderResponse response;
                try { response = await Task.Run(() => backend.Render(request, token), token).ConfigureAwait(false); }
                catch (OperationCanceledException) { throw; }
                catch (Exception error)
                {
                    LogFailure(error);
                    response = new NativeRenderResponse { Error = "Private native rendering failed." };
                }
                using var write = CancellationTokenSource.CreateLinkedTokenSource(token);
                write.CancelAfter(5000);
                await NativeRenderProtocol.WriteAsync(endpoint, response,
                    NativeRenderProtocol.MaxResponseBytes, write.Token).ConfigureAwait(false);
                var acknowledgement = new byte[1];
                if (await endpoint.ReadAsync(acknowledgement, 0, 1, write.Token).ConfigureAwait(false) != 1
                    || acknowledgement[0] != 1)
                    throw new IOException("Render response was not acknowledged.");
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested) { break; }
            catch (ObjectDisposedException) when (token.IsCancellationRequested) { break; }
            catch (Exception error) { LogFailure(error); }
            finally
            {
                // Reuse the first instance so another process cannot acquire the
                // name between requests. Rendering is serialized across Office apps.
                try { if (endpoint.IsConnected) endpoint.Disconnect(); } catch (IOException) { }
                catch (ObjectDisposedException) { }
            }
        }
    }

    private void LogFailure(Exception error)
    {
        // Bounded diagnostics, never formula source, document names or profiles.
        try
        {
            string path = Path.Combine(data, "service.log");
            if (File.Exists(path) && new FileInfo(path).Length > 256 * 1024) File.Delete(path);
            File.AppendAllText(path, DateTime.UtcNow.ToString("O") + " " + error.GetType().Name + Environment.NewLine);
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private static void VerifyPrivateIdentity()
    {
        using var identity = WindowsIdentity.GetCurrent();
        string expected = ((SecurityIdentifier)new NTAccount("NT SERVICE", NativeRenderProtocol.ServiceName)
            .Translate(typeof(SecurityIdentifier))).Value;
        if (identity.User?.Value != expected || Process.GetCurrentProcess().SessionId != 0)
            throw new InvalidOperationException("Rendering requires its dedicated noninteractive service account.");
        using var user = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64);
        NtQueryKey(user.Handle.DangerousGetHandle(), 3, IntPtr.Zero, 0, out int size);
        if (size <= 4 || size > 4096) throw new InvalidOperationException("Private service profile is not loaded.");
        IntPtr buffer = Marshal.AllocHGlobal(size);
        try
        {
            if (NtQueryKey(user.Handle.DangerousGetHandle(), 3, buffer, size, out _) != 0)
                throw new InvalidOperationException("Cannot validate service profile.");
            int length = Marshal.ReadInt32(buffer);
            if (length <= 0 || length > size - 4) throw new InvalidOperationException("Invalid registry name.");
            string name = Marshal.PtrToStringUni(IntPtr.Add(buffer, 4), length / 2);
            if (!string.Equals(name, @"\REGISTRY\USER\" + expected, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Refusing a shared/default service profile.");
        }
        finally { Marshal.FreeHGlobal(buffer); }
    }

    private static NamedPipeServerStream CreateLocalPipe()
    {
        using var identity = WindowsIdentity.GetCurrent();
        // AU can read/write only (not FILE_CREATE_PIPE_INSTANCE). Deny network
        // tokens as well as rejecting remote pipe clients at the kernel boundary.
        string sddl = "D:P(D;;GA;;;NU)(A;;GA;;;SY)(A;;GA;;;" + identity.User!.Value + ")(A;;0x12019b;;;AU)";
        if (!ConvertStringSecurityDescriptorToSecurityDescriptor(sddl, 1, out IntPtr descriptor, out _))
            throw new Win32Exception();
        try
        {
            var security = new SecurityAttributes { Length = Marshal.SizeOf<SecurityAttributes>(), Descriptor = descriptor };
            var handle = CreateNamedPipe(@"\\.\pipe\" + NativeRenderProtocol.PipeName,
                3 | 0x40000000 | 0x00080000, 0x8, 1, 65536, 65536, 0, ref security);
            if (handle.IsInvalid) { handle.Dispose(); throw new Win32Exception(); }
            try { return new NamedPipeServerStream(PipeDirection.InOut, true, false, handle); }
            catch { handle.Dispose(); throw; }
        }
        finally { LocalFree(descriptor); }
    }

    [StructLayout(LayoutKind.Sequential)] private struct SecurityAttributes
    { public int Length; public IntPtr Descriptor; [MarshalAs(UnmanagedType.Bool)] public bool Inherit; }
    [DllImport("ntdll.dll")] private static extern int NtQueryKey(IntPtr key, int kind, IntPtr data, int length, out int needed);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool ConvertStringSecurityDescriptorToSecurityDescriptor(string sddl, uint revision, out IntPtr descriptor, out uint size);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern SafePipeHandle CreateNamedPipe(string name, uint openMode, uint pipeMode, uint instances, uint output, uint input, uint timeout, ref SecurityAttributes security);
    [DllImport("kernel32.dll")] private static extern IntPtr LocalFree(IntPtr memory);
}
