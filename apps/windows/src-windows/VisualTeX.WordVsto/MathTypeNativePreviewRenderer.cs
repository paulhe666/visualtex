using System.Diagnostics;
using System.Text.Json;
using Microsoft.Win32;

namespace VisualTeX.WordVsto;

internal static class MathTypeNativePreviewRenderer
{
    private const int RenderTimeoutMilliseconds = 15_000;
    private const int MaxBatchItemsPerSidecar = 128;
    private const string BridgeFileName = "visualtex-windows-office-bridge.exe";
    private const string BridgeOverrideEnvironment = "VISUALTEX_MATHTYPE_PREVIEW_BRIDGE_PATH";

    private sealed class Request
    {
        public string ResultPath { get; set; } = string.Empty;
        public List<RequestItem> Items { get; set; } = new();
    }

    private sealed class RequestItem
    {
        public string Id { get; set; } = string.Empty;
        public string MtefPath { get; set; } = string.Empty;
        public string WmfPath { get; set; } = string.Empty;
    }

    private sealed class Response
    {
        public string Error { get; set; } = string.Empty;
        public List<ResponseItem> Items { get; set; } = new();
    }

    private sealed class ResponseItem
    {
        public string Id { get; set; } = string.Empty;
        public bool Success { get; set; }
        public string Error { get; set; } = string.Empty;
        public string WmfPath { get; set; } = string.Empty;
        public float WidthPt { get; set; }
        public float HeightPt { get; set; }
        public int WordPosition { get; set; }
    }

    internal sealed class Result : IDisposable
    {
        public string WmfPath { get; set; } = string.Empty;
        public float WidthPt { get; set; }
        public float HeightPt { get; set; }
        public int WordPosition { get; set; }

        public void Dispose()
        {
            if (string.IsNullOrWhiteSpace(WmfPath)) return;
            try { File.Delete(WmfPath); } catch { }
        }
    }

    internal static bool TryRenderBatch(
        IReadOnlyDictionary<string, byte[]> mtefs,
        string outputDirectory,
        out Dictionary<string, Result> results)
    {
        results = new Dictionary<string, Result>(StringComparer.Ordinal);
        if (mtefs is null || mtefs.Count == 0) return false;
        if (mtefs.Any(item => string.IsNullOrWhiteSpace(item.Key)
                || item.Value is null
                || item.Value.Length == 0))
            return false;

        var bridgePath = ResolveBridgePath();
        if (bridgePath is null) return false;

        Directory.CreateDirectory(outputDirectory);
        var mathTypeBaseline = SnapshotMathTypeProcessIds();
        try
        {
            var ordered = mtefs.ToArray();
            for (var offset = 0; offset < ordered.Length; offset += MaxBatchItemsPerSidecar)
            {
                var chunk = new Dictionary<string, byte[]>(StringComparer.Ordinal);
                var end = Math.Min(ordered.Length, offset + MaxBatchItemsPerSidecar);
                for (var index = offset; index < end; index++)
                    chunk[ordered[index].Key] = ordered[index].Value;
                RenderBatchChunkResilient(
                    bridgePath,
                    chunk,
                    outputDirectory,
                    results,
                    mathTypeBaseline);
            }
            return results.Count == mtefs.Count;
        }
        finally
        {
            // Keep MathType's windowless -mtrpc helper alive between isolated
            // sidecars in the same document batch. Starting and killing it once
            // per formula turns 100 conversions into minutes of avoidable latency.
            // Cleanup remains unconditional at the end of the batch and never
            // touches a MathType process that existed before this operation.
            CleanupNewWindowlessMathTypeProcesses(mathTypeBaseline);
        }
    }

    private static void RenderBatchChunkResilient(
        string bridgePath,
        IReadOnlyDictionary<string, byte[]> mtefs,
        string outputDirectory,
        Dictionary<string, Result> aggregate,
        HashSet<int> mathTypeBaseline)
    {
        if (mtefs.Count == 0) return;

        _ = TryRenderBatchCore(
            bridgePath,
            mtefs,
            outputDirectory,
            out var rendered);
        foreach (var pair in rendered)
            aggregate[pair.Key] = pair.Value;

        var missing = mtefs
            .Where(pair => !rendered.ContainsKey(pair.Key))
            .ToArray();
        if (missing.Length == 0 || mtefs.Count == 1) return;

        // MathPage is legacy native code. A repeated transform can occasionally
        // terminate its sidecar, hang near the watchdog, or leave later items in
        // a poisoned API state even when the process exits normally. Restart the
        // windowless helper before retrying only the missing portion in smaller
        // isolated processes; ultimately a single formula is isolated without
        // forcing every successful formula down the expensive one-process-per-
        // formula path.
        CleanupNewWindowlessMathTypeProcesses(mathTypeBaseline);
        var midpoint = Math.Max(1, missing.Length / 2);
        var left = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        var right = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        for (var index = 0; index < missing.Length; index++)
        {
            var destination = index < midpoint ? left : right;
            destination[missing[index].Key] = missing[index].Value;
        }
        RenderBatchChunkResilient(
            bridgePath,
            left,
            outputDirectory,
            aggregate,
            mathTypeBaseline);
        // A left retry can itself terminate in MathPage native code and leave
        // its windowless helper poisoned. Isolate the right sibling as well;
        // otherwise one genuinely bad MTEF can make the next valid formula look
        // like a second failure.
        CleanupNewWindowlessMathTypeProcesses(mathTypeBaseline);
        RenderBatchChunkResilient(
            bridgePath,
            right,
            outputDirectory,
            aggregate,
            mathTypeBaseline);
    }

    private static bool TryRenderBatchCore(
        string bridgePath,
        IReadOnlyDictionary<string, byte[]> mtefs,
        string outputDirectory,
        out Dictionary<string, Result> results)
    {
        results = new Dictionary<string, Result>(StringComparer.Ordinal);
        var requestId = Guid.NewGuid().ToString("N");
        var requestRoot = Path.Combine(
            outputDirectory,
            $"mathtype-native-sidecar-batch-{requestId}");
        var manifestPath = requestRoot + ".request.json";
        var responsePath = requestRoot + ".response.json";
        var request = new Request { ResultPath = responsePath };
        var temporaryMtefPaths = new List<string>();
        var temporaryWmfPaths = new List<string>();
        Process? process = null;
        try
        {
            var index = 0;
            foreach (var pair in mtefs)
            {
                index++;
                var itemRoot = requestRoot + $"-{index:D4}";
                var mtefPath = itemRoot + ".mtef";
                var wmfPath = itemRoot + ".wmf";
                File.WriteAllBytes(mtefPath, pair.Value);
                temporaryMtefPaths.Add(mtefPath);
                temporaryWmfPaths.Add(wmfPath);
                request.Items.Add(new RequestItem
                {
                    Id = pair.Key,
                    MtefPath = mtefPath,
                    WmfPath = wmfPath,
                });
            }
            File.WriteAllText(manifestPath, JsonSerializer.Serialize(request));

            var startInfo = new ProcessStartInfo
            {
                FileName = bridgePath,
                Arguments = "--mathtype-preview-manifest " + QuoteProcessArgument(manifestPath),
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = false,
                RedirectStandardError = false,
                WorkingDirectory = Path.GetDirectoryName(bridgePath) ?? outputDirectory,
            };
            process = Process.Start(startInfo);
            if (process is null) return false;
            // Real MathPage throughput on the 100-formula corpus is about
            // 0.45 seconds/item. The old 45-second ceiling killed a healthy batch
            // after 98/100 checkpoints. Scale generously with batch size while
            // retaining a finite watchdog for a genuinely hung native transform.
            var timeout = Math.Max(
                RenderTimeoutMilliseconds,
                Math.Min(120_000, mtefs.Count * 1_500));
            var completedGracefully = process.WaitForExit(timeout);
            if (!completedGracefully)
            {
                try { process.Kill(); } catch { }
                try { process.WaitForExit(2_000); } catch { }
            }
            else
            {
                try { completedGracefully = process.ExitCode == 0; }
                catch { completedGracefully = false; }
            }
            // The sidecar checkpoints its JSON response after every item. Even if
            // legacy MathPage crashes or the watchdog kills a hung batch, recover
            // all WMFs completed before that point and retry only the missing tail.
            if (!File.Exists(responsePath)) return false;

            Response? response;
            try
            {
                response = JsonSerializer.Deserialize<Response>(File.ReadAllText(responsePath));
            }
            catch
            {
                return false;
            }
            if (response?.Items is null) return false;

            foreach (var item in response.Items)
            {
                if (!item.Success
                    || string.IsNullOrWhiteSpace(item.Id)
                    || !mtefs.ContainsKey(item.Id)
                    || !(item.WidthPt > 0)
                    || !(item.HeightPt > 0)
                    || string.IsNullOrWhiteSpace(item.WmfPath)
                    || !File.Exists(item.WmfPath)
                    || new FileInfo(item.WmfPath).Length <= 22)
                {
                    WordDoubleClickHook.TraceMessage(
                        $"mathtype-native-preview-item-failed id={item.Id} error={item.Error} "
                        + $"success={item.Success} width={item.WidthPt} height={item.HeightPt}");
                    continue;
                }
                results[item.Id] = new Result
                {
                    WmfPath = item.WmfPath,
                    WidthPt = item.WidthPt,
                    HeightPt = item.HeightPt,
                    WordPosition = item.WordPosition,
                };
            }

            foreach (var result in results.Values)
                temporaryWmfPaths.Remove(result.WmfPath);
            return completedGracefully;
        }
        catch
        {
            return false;
        }
        finally
        {
            process?.Dispose();
            foreach (var path in temporaryMtefPaths) TryDelete(path);
            foreach (var path in temporaryWmfPaths) TryDelete(path);
            TryDelete(manifestPath);
            TryDelete(responsePath);
        }
    }

    internal static bool TryRender(byte[] mtef, string outputDirectory, out Result result)
    {
        result = new Result();
        if (mtef is null || mtef.Length == 0) return false;

        var bridgePath = ResolveBridgePath();
        if (bridgePath is null) return false;

        Directory.CreateDirectory(outputDirectory);
        var requestId = Guid.NewGuid().ToString("N");
        var requestRoot = Path.Combine(
            outputDirectory,
            $"mathtype-native-sidecar-{requestId}");
        var mtefPath = requestRoot + ".mtef";
        var wmfPath = requestRoot + ".wmf";
        var manifestPath = requestRoot + ".request.json";
        var responsePath = requestRoot + ".response.json";
        Process? process = null;
        var keepWmf = false;
        var mathTypeBaseline = SnapshotMathTypeProcessIds();
        try
        {
            File.WriteAllBytes(mtefPath, mtef);
            var request = new Request
            {
                ResultPath = responsePath,
                Items = new List<RequestItem>
                {
                    new()
                    {
                        Id = requestId,
                        MtefPath = mtefPath,
                        WmfPath = wmfPath,
                    },
                },
            };
            File.WriteAllText(manifestPath, JsonSerializer.Serialize(request));

            var startInfo = new ProcessStartInfo
            {
                FileName = bridgePath,
                Arguments = "--mathtype-preview-manifest " + QuoteProcessArgument(manifestPath),
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = false,
                RedirectStandardError = false,
                WorkingDirectory = Path.GetDirectoryName(bridgePath) ?? outputDirectory,
            };
            process = Process.Start(startInfo);
            if (process is null) return false;
            if (!process.WaitForExit(RenderTimeoutMilliseconds))
            {
                try { process.Kill(); } catch { }
                return false;
            }

            // A legacy MathPage AccessViolation terminates only the sidecar. In
            // that case there may be no response file at all; treat it as a normal
            // preview miss so Word can safely fall back to the frontend geometry.
            if (process.ExitCode != 0 || !File.Exists(responsePath)) return false;
            Response? response;
            try
            {
                response = JsonSerializer.Deserialize<Response>(File.ReadAllText(responsePath));
            }
            catch
            {
                return false;
            }
            var item = response?.Items?.FirstOrDefault(value =>
                string.Equals(value.Id, requestId, StringComparison.Ordinal));
            if (item is null
                || !item.Success
                || !(item.WidthPt > 0)
                || !(item.HeightPt > 0)
                || string.IsNullOrWhiteSpace(item.WmfPath)
                || !File.Exists(item.WmfPath)
                || new FileInfo(item.WmfPath).Length <= 22)
                return false;

            result = new Result
            {
                WmfPath = item.WmfPath,
                WidthPt = item.WidthPt,
                HeightPt = item.HeightPt,
                WordPosition = item.WordPosition,
            };
            keepWmf = true;
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            process?.Dispose();
            CleanupNewWindowlessMathTypeProcesses(mathTypeBaseline);
            TryDelete(mtefPath);
            TryDelete(manifestPath);
            TryDelete(responsePath);
            if (!keepWmf) TryDelete(wmfPath);
        }
    }

    private static string? ResolveBridgePath()
    {
        var overridden = Environment.GetEnvironmentVariable(BridgeOverrideEnvironment);
        if (!string.IsNullOrWhiteSpace(overridden))
        {
            try
            {
                var full = Path.GetFullPath(overridden.Trim().Trim('"'));
                if (File.Exists(full)) return full;
            }
            catch { }
        }

        // Installed builds register the exact Tauri companion executable. The
        // office bridge sidecar is bundled next to it by Tauri externalBin.
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\VisualTeX\OfficeIntegration");
            var executable = key?.GetValue("ExecutablePath") as string;
            if (!string.IsNullOrWhiteSpace(executable))
            {
                var directory = Path.GetDirectoryName(executable.Trim().Trim('"'));
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    var candidate = Path.Combine(directory, BridgeFileName);
                    if (File.Exists(candidate)) return candidate;
                }
            }
        }
        catch { }

        var localAppCandidate = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            BridgeFileName);
        if (File.Exists(localAppCandidate)) return localAppCandidate;

        // Current-source acceptance runs the VSTO DLL from
        // src-windows/VisualTeX.VstoFlowAcceptance/bin/... . Walk up to the
        // src-windows root and use the freshly-built sidecar rather than whatever
        // version happens to be installed on the machine.
        try
        {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            for (var depth = 0; directory is not null && depth < 12; depth++, directory = directory.Parent)
            {
                if (!string.Equals(directory.Name, "src-windows", StringComparison.OrdinalIgnoreCase))
                    continue;
                var releaseCandidate = Path.Combine(
                    directory.FullName,
                    "VisualTeX.WindowsOleBridge",
                    "bin",
                    "x64",
                    "Release",
                    "net8.0-windows",
                    "win-x64",
                    BridgeFileName);
                if (File.Exists(releaseCandidate)) return releaseCandidate;
                var artifactCandidate = Path.Combine(
                    directory.FullName,
                    "artifacts",
                    "windows-ole-bridge",
                    BridgeFileName);
                if (File.Exists(artifactCandidate)) return artifactCandidate;
                break;
            }
        }
        catch { }

        return null;
    }

    private static HashSet<int> SnapshotMathTypeProcessIds()
    {
        var result = new HashSet<int>();
        try
        {
            foreach (var process in Process.GetProcessesByName("MathType"))
            {
                try { result.Add(process.Id); }
                catch { }
                finally { process.Dispose(); }
            }
        }
        catch { }
        return result;
    }

    private static void CleanupNewWindowlessMathTypeProcesses(HashSet<int> baseline)
    {
        // MTInitAPI can leave a delayed `MathType.exe -mtrpc` server behind even
        // after the rendering sidecar has exited. Never touch a MathType process
        // that existed before this render, and never close a visible MathType UI.
        // This cleanup runs in the surviving VSTO process, so it also executes when
        // legacy MathPage native code terminates the sidecar with AccessViolation.
        try
        {
            // MTTermAPI can schedule the windowless -mtrpc server slightly after
            // the rendering sidecar has already exited. An immediate snapshot can
            // therefore miss the process and leave it behind. Give only that
            // delayed server a short arrival window, then keep the existing rule:
            // never touch a pre-existing PID and never terminate visible MathType UI.
            Thread.Sleep(175);
            for (var attempt = 0; attempt < 10; attempt++)
            {
                var foundNewWindowless = false;
                foreach (var process in Process.GetProcessesByName("MathType"))
                {
                    try
                    {
                        if (baseline.Contains(process.Id)) continue;
                        process.Refresh();
                        if (process.HasExited || process.MainWindowHandle != IntPtr.Zero) continue;
                        foundNewWindowless = true;
                        try { process.Kill(); } catch { }
                    }
                    catch { }
                    finally { process.Dispose(); }
                }
                if (!foundNewWindowless) return;
                Thread.Sleep(100);
            }
        }
        catch { }
    }

    private static string QuoteProcessArgument(string value) =>
        "\"" + value.Replace("\"", "\\\"") + "\"";

    private static void TryDelete(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return;
        try { File.Delete(path); } catch { }
    }
}
