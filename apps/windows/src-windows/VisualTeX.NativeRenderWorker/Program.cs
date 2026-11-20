using System.Diagnostics;
using System.IO.Pipes;
using System.Text;
using System.Security.Principal;

namespace VisualTeX.NativeRenderWorker;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        using var identity = WindowsIdentity.GetCurrent();
        // No interactive fallback. Loading MathPage in an Office/user session
        // would reintroduce the single-instance and shared-profile conflict.
        if (args.Length != 1 || Process.GetCurrentProcess().SessionId != 0
            || identity.User?.Value.StartsWith("S-1-5-80-", StringComparison.Ordinal) != true)
            return 10;
        var runtime = Path.Combine(AppContext.BaseDirectory, "runtime");
        var mathPage = Path.Combine(runtime, "MathPage", "64", "MathPage.wll");
        if (!File.Exists(mathPage)) return 11;
        // This pipe is created with an ACL for the dedicated service SID only.
        // The public Office endpoint remains data-only; request paths here are
        // service-generated files within its private job directory.
        using var pipe = new NamedPipeClientStream(".", args[0], PipeDirection.InOut);
        pipe.Connect(5000);
        using var reader = new BinaryReader(pipe, Encoding.UTF8, leaveOpen: true);
        using var writer = new BinaryWriter(pipe, Encoding.UTF8, leaveOpen: true);
        using var session = new MathTypeNativePreviewCommand.NativeSession(runtime);
        while (true)
        {
            string manifest;
            try { manifest = reader.ReadString(); }
            catch (EndOfStreamException) { return 0; }
            int status = session.Render(manifest);
            writer.Write(status);
            writer.Flush();
            if (status != 0) return status;
        }
    }
}
