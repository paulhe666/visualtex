using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.NativeRenderService;

internal sealed class NativeWorkerBackend
{
    private readonly string executable = Path.Combine(AppContext.BaseDirectory, "VisualTeX.NativeRenderWorker.exe");
    private readonly string jobs;
    internal NativeWorkerBackend(string dataDirectory)
    {
        jobs = Path.Combine(dataDirectory, "jobs");
        Directory.CreateDirectory(jobs);
    }

    internal NativeRenderResponse Render(NativeRenderRequest input, CancellationToken token)
    {
        NativeRenderProtocol.Validate(input);
        // Every filename is generated here. Client ids are only opaque labels.
        string root = Path.Combine(jobs, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var request = new WorkerRequest { ResultPath = Path.Combine(root, "response.json") };
            for (int index = 0; index < input.Items.Count; index++)
            {
                var item = new WorkerItem
                {
                    Id = index.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    MtefPath = Path.Combine(root, index + ".mtef"),
                    WmfPath = Path.Combine(root, index + ".wmf"),
                };
                File.WriteAllBytes(item.MtefPath, input.Items[index].Mtef);
                request.Items.Add(item);
            }
            string manifest = Path.Combine(root, "request.json");
            File.WriteAllText(manifest, JsonSerializer.Serialize(request));
            string error = string.Empty;
            try
            {
                using var job = new OwnedProcessJob();
                int exitCode = job.Run(executable, manifest, root,
                    Math.Max(20_000, Math.Min(120_000, input.Items.Count * 1500)), token);
                if (exitCode != 0) error = "Native worker failed (" + exitCode + ").";
            }
            catch (TimeoutException) { error = "Native worker timed out."; }

            // Read checkpoints only after the entire owned process tree exits.
            // Completed items survive a legacy-engine crash; no global cleanup.
            var result = new NativeRenderResponse { Error = error };
            if (!File.Exists(request.ResultPath)) return result;
            if (new FileInfo(request.ResultPath).Length > 1024 * 1024)
                throw new InvalidDataException("Worker response exceeds its limit.");
            var response = JsonSerializer.Deserialize<WorkerResponse>(File.ReadAllText(request.ResultPath));
            if (response?.Items == null) return result;
            if (!string.IsNullOrEmpty(response.Error)) result.Error = "Native engine could not render the request.";
            var seen = new HashSet<int>();
            long total = 0;
            foreach (var item in response.Items)
            {
                if (!item.Success || !int.TryParse(item.Id, out int index) || index < 0
                    || index >= input.Items.Count || !seen.Add(index)) continue;
                // Ignore the returned path even though the worker is private.
                string wmf = request.Items[index].WmfPath;
                if (!File.Exists(wmf)) continue;
                long length = new FileInfo(wmf).Length;
                if (length <= 22 || length > NativeRenderProtocol.MaxWmfBytes
                    || total + length > NativeRenderProtocol.MaxResponseBytes / 2) continue;
                total += length;
                result.Items.Add(new NativeRenderOutput
                {
                    Id = input.Items[index].Id, Wmf = File.ReadAllBytes(wmf),
                    WidthPt = item.WidthPt, HeightPt = item.HeightPt, WordPosition = item.WordPosition,
                });
            }
            NativeRenderProtocol.Validate(result, input);
            return result;
        }
        finally
        {
            // Private generated descendant, with no client-controlled path and no
            // write permission for callers. Never touch document/temp user trees.
            try { Directory.Delete(root, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private sealed class WorkerRequest
    { public string ResultPath { get; set; } = string.Empty; public List<WorkerItem> Items { get; set; } = new(); }
    private sealed class WorkerItem
    { public string Id { get; set; } = string.Empty; public string MtefPath { get; set; } = string.Empty; public string WmfPath { get; set; } = string.Empty; }
    private sealed class WorkerResponse
    { public string Error { get; set; } = string.Empty; public List<WorkerOutput> Items { get; set; } = new(); }
    private sealed class WorkerOutput
    { public string Id { get; set; } = string.Empty; public bool Success { get; set; } public float WidthPt { get; set; } public float HeightPt { get; set; } public int WordPosition { get; set; } }
}
