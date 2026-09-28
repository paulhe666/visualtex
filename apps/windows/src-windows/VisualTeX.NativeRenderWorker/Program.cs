using System.Diagnostics;
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
        return MathTypeNativePreviewCommand.Run(args[0], runtime);
    }
}
