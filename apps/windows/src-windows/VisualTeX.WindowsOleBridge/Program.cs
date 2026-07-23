using System.Diagnostics;
using System.Security.Principal;

namespace VisualTeX.WindowsOleBridge;

internal static class Program
{
    [STAThread]
    private static async Task<int> Main(string[] args)
    {
        var options = ParseArguments(args);
        var logRoot = Required(options, "log-root");
        var logger = new FileLogger(logRoot);
        try
        {
            var sid = WindowsIdentity.GetCurrent().User?.Value
                ?? throw new InvalidOperationException("Unable to resolve current Windows user SID.");
            var token = Required(options, "token");
            if (token.Length < 64)
                throw new UnauthorizedAccessException("The Office bridge token is too short.");

            if (!int.TryParse(Required(options, "parent-pid"), out var parentPid) ||
                parentPid <= 0 || parentPid == Environment.ProcessId)
                throw new ArgumentException("The Office bridge parent process ID is invalid.");

            var acceptanceMode = options.TryGetValue("acceptance", out var acceptanceValue) &&
                string.Equals(acceptanceValue, "true", StringComparison.OrdinalIgnoreCase);
            var expectedPipe = acceptanceMode
                ? $@"\\.\pipe\VisualTeX.OfficeBridge.{sid}.Acceptance.{parentPid}"
                : $@"\\.\pipe\VisualTeX.OfficeBridge.{sid}";
            var pipeName = Required(options, "pipe-name");
            if (!string.Equals(pipeName, expectedPipe, StringComparison.OrdinalIgnoreCase))
                throw new UnauthorizedAccessException("The Office bridge pipe name does not match the secured current-user endpoint.");

            var expectedTempRoot = Path.GetFullPath(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "VisualTeX",
                "office",
                "temp"));
            var tempRoot = Path.GetFullPath(Required(options, "temp-root"));
            if (!string.Equals(tempRoot, expectedTempRoot, StringComparison.OrdinalIgnoreCase))
                throw new UnauthorizedAccessException("The Office temp root is not the VisualTeX current-user directory.");

            var mutexName = acceptanceMode
                ? $@"Local\VisualTeX.OfficeBridge.{sid}.Acceptance.{parentPid}"
                : $@"Local\VisualTeX.OfficeBridge.{sid}";
            using var singleInstance = new Mutex(
                initiallyOwned: true,
                name: mutexName,
                createdNew: out var createdNew);
            if (!createdNew)
            {
                logger.Info("Another Windows Office bridge instance is already running.");
                return 0;
            }

            using var backend = new WindowsOfficeBackend(tempRoot, logger);
            var server = new NamedPipeServer(pipeName, token, backend, logger);
            using var cancellation = new CancellationTokenSource();
            using var parent = Process.GetProcessById(parentPid);
            void CancelBridge()
            {
                try
                {
                    if (!cancellation.IsCancellationRequested)
                        cancellation.Cancel();
                }
                catch (ObjectDisposedException)
                {
                    // A parent/process-exit notification can race normal disposal.
                }
            }
            EventHandler parentExited = (_, _) => CancelBridge();
            EventHandler processExiting = (_, _) => CancelBridge();
            ConsoleCancelEventHandler consoleCancelled = (_, eventArgs) =>
            {
                eventArgs.Cancel = true;
                CancelBridge();
            };
            parent.EnableRaisingEvents = true;
            parent.Exited += parentExited;
            if (parent.HasExited)
                CancelBridge();
            AppDomain.CurrentDomain.ProcessExit += processExiting;
            Console.CancelKeyPress += consoleCancelled;
            try
            {
                await server.RunAsync(cancellation.Token);
            }
            finally
            {
                parent.Exited -= parentExited;
                AppDomain.CurrentDomain.ProcessExit -= processExiting;
                Console.CancelKeyPress -= consoleCancelled;
            }
            return 0;
        }
        catch (Exception error)
        {
            logger.Error("Windows Office bridge terminated.", error);
            return 1;
        }
    }

    private static Dictionary<string, string> ParseArguments(string[] args)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < args.Length; index++)
        {
            var key = args[index];
            if (!key.StartsWith("--", StringComparison.Ordinal) || index + 1 >= args.Length)
                throw new ArgumentException($"Invalid argument: {key}");
            result[key[2..]] = args[++index];
        }
        return result;
    }

    private static string Required(IReadOnlyDictionary<string, string> values, string key) =>
        values.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : throw new ArgumentException($"Missing required argument --{key}.");
}
