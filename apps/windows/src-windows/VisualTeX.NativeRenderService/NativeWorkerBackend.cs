using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.NativeRenderService;

internal sealed class NativeWorkerBackend : IDisposable
{
    private readonly string executable = Path.Combine(AppContext.BaseDirectory, "VisualTeX.NativeRenderWorker.exe");
    private readonly string jobs;
    private OwnedProcessJob? runtimeJob;
    private string? runtimeDirectory;
    private NamedPipeServerStream? workerPipe;
    internal NativeWorkerBackend(string dataDirectory)
    {
        jobs = Path.Combine(dataDirectory, "jobs");
        Directory.CreateDirectory(jobs);
    }

    internal NativeRenderResponse Render(NativeRenderRequest input, CancellationToken token)
    {
        NativeRenderProtocol.Validate(input);
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(token);
        // At least the worker's 30 s MathType launch time-out plus the render.
        deadline.CancelAfter(Math.Max(60_000, Math.Min(120_000, input.Items.Count * 1500)));
        EnsureRuntimeLifetime(deadline.Token);
        // The worker owns the MathPage API lease across requests. Per-request
        // files are disposable; engine/TEMP/Job live until failure or shutdown.
        string root = Path.Combine(runtimeDirectory!, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        bool complete = false;
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
            var endpoint = workerPipe!;
            // Only this service account can open the private worker pipe. The
            // manifest path is generated above, never supplied by an Office client.
            using (deadline.Token.Register(() => endpoint.Dispose()))
            {
                try
                {
                    using var writer = new BinaryWriter(endpoint, Encoding.UTF8, leaveOpen: true);
                    using var reader = new BinaryReader(endpoint, Encoding.UTF8, leaveOpen: true);
                    writer.Write(manifest);
                    writer.Flush();
                    int status = reader.ReadInt32();
                    if (status == NativeRenderProtocol.WorkerFontsUnavailableExitCode)
                        error = NativeRenderProtocol.FontsUnavailableError;
                    else if (status != 0) error = "Native worker failed (" + status + ").";
                }
                catch (IOException) { error = "Native worker disconnected."; }
                catch (ObjectDisposedException) { error = "Native worker request was cancelled."; }
            }
            if (error.Length != 0)
            {
                // A failed/timed-out worker must be quiescent before reading its
                // last checkpoint. The outer finally resets just this owned job.
                runtimeJob!.Dispose();
                token.ThrowIfCancellationRequested();
            }

            var result = new NativeRenderResponse { Error = error };
            if (!File.Exists(request.ResultPath))
            {
                AddFailures(result, input, new Dictionary<int, string>(), error);
                return result;
            }
            if (new FileInfo(request.ResultPath).Length > 1024 * 1024)
                throw new InvalidDataException("Worker response exceeds its limit.");
            var response = JsonSerializer.Deserialize<WorkerResponse>(File.ReadAllText(request.ResultPath));
            if (response?.Items == null)
            {
                AddFailures(result, input, new Dictionary<int, string>(), error);
                return result;
            }
            if (!string.IsNullOrEmpty(response.Error) && result.Error != NativeRenderProtocol.FontsUnavailableError)
                result.Error = "Native engine could not render the request.";
            var seen = new HashSet<int>();
            var reasons = new Dictionary<int, string>();
            long total = 0;
            foreach (var item in response.Items)
            {
                if (!int.TryParse(item.Id, out int index) || index < 0
                    || index >= input.Items.Count || !seen.Add(index)) continue;
                if (!item.Success)
                {
                    reasons[index] = item.Error;
                    continue;
                }
                // Ignore the returned path even though the worker is private.
                string wmf = request.Items[index].WmfPath;
                if (!File.Exists(wmf)) { reasons[index] = "Rendered picture is missing."; continue; }
                long length = new FileInfo(wmf).Length;
                if (length <= 22 || length > NativeRenderProtocol.MaxWmfBytes
                    || total + length > NativeRenderProtocol.MaxResponseBytes / 2)
                {
                    reasons[index] = $"Rendered picture size {length} bytes is outside the limit.";
                    continue;
                }
                total += length;
                result.Items.Add(new NativeRenderOutput
                {
                    Id = input.Items[index].Id, Wmf = File.ReadAllBytes(wmf),
                    WidthPt = item.WidthPt, HeightPt = item.HeightPt, WordPosition = item.WordPosition,
                });
            }
            AddFailures(result, input, reasons, response.Error);
            NativeRenderProtocol.Validate(result, input);
            complete = string.IsNullOrEmpty(result.Error) && result.Items.Count == input.Items.Count;
            return result;
        }
        finally
        {
            // Private generated descendant, with no client-controlled path and no
            // write permission for callers. Never touch document/temp user trees.
            DeleteOwnedDirectory(root);
            if (!complete) Dispose();
        }
    }

    // Every requested item without output gets the engine's reason. An item the
    // worker never reached (crash, timeout) carries the request-level error.
    private static void AddFailures(
        NativeRenderResponse result, NativeRenderRequest input,
        IReadOnlyDictionary<int, string> reasons, string requestError)
    {
        var rendered = new HashSet<string>(result.Items.Select(item => item.Id), StringComparer.Ordinal);
        for (int index = 0; index < input.Items.Count; index++)
        {
            if (rendered.Contains(input.Items[index].Id)) continue;
            string reason = reasons.TryGetValue(index, out var itemReason) && !string.IsNullOrWhiteSpace(itemReason)
                ? itemReason
                : string.IsNullOrWhiteSpace(requestError) ? "The native worker returned no result for this item." : requestError;
            result.Failures.Add(new NativeRenderFailure
            {
                Id = input.Items[index].Id,
                Reason = NativeRenderProtocol.BoundFailureReason(reason),
            });
        }
    }

    private void EnsureRuntimeLifetime(CancellationToken token)
    {
        if (runtimeJob is not null) return;
        string directory = Path.Combine(jobs, "runtime-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            runtimeDirectory = directory;
            string pipeName = "VisualTeX.NativeWorker." + Guid.NewGuid().ToString("N");
            using var identity = WindowsIdentity.GetCurrent();
            var security = new PipeSecurity();
            security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);
            security.AddAccessRule(new PipeAccessRule(identity.User!, PipeAccessRights.FullControl, AccessControlType.Allow));
            security.AddAccessRule(new PipeAccessRule(new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null),
                PipeAccessRights.FullControl, AccessControlType.Allow));
            workerPipe = new NamedPipeServerStream(pipeName, PipeDirection.InOut, 1,
                PipeTransmissionMode.Byte, PipeOptions.Asynchronous, 4096, 4096, security);
            runtimeJob = new OwnedProcessJob();
            runtimeJob.Start(executable, pipeName, directory);
            workerPipe.WaitForConnectionAsync(token).GetAwaiter().GetResult();
        }
        catch
        {
            Dispose();
            throw;
        }
    }

    public void Dispose()
    {
        // End the exact private job before deleting TEMP files still used by its
        // native children. The next request lazily creates a fresh lifetime.
        try { runtimeJob?.Dispose(); }
        finally
        {
            workerPipe?.Dispose();
            workerPipe = null;
            runtimeJob = null;
            if (runtimeDirectory is not null) DeleteOwnedDirectory(runtimeDirectory);
            runtimeDirectory = null;
        }
    }

    private static void DeleteOwnedDirectory(string directory)
    {
        try { Directory.Delete(directory, recursive: true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private sealed class WorkerRequest
    { public string ResultPath { get; set; } = string.Empty; public List<WorkerItem> Items { get; set; } = new(); }
    private sealed class WorkerItem
    { public string Id { get; set; } = string.Empty; public string MtefPath { get; set; } = string.Empty; public string WmfPath { get; set; } = string.Empty; }
    private sealed class WorkerResponse
    { public string Error { get; set; } = string.Empty; public List<WorkerOutput> Items { get; set; } = new(); }
    private sealed class WorkerOutput
    { public string Id { get; set; } = string.Empty; public bool Success { get; set; } public string Error { get; set; } = string.Empty; public float WidthPt { get; set; } public float HeightPt { get; set; } public int WordPosition { get; set; } }
}
