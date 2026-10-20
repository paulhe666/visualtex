using System.Diagnostics;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WordVsto;

// Office owns document insertion and equation encoding, not the native engine.
// There is deliberately no installed-MathType discovery, prewarm, process launch
// or process cleanup in this adapter. Those lifetimes belong to the private service.
internal static class MathTypeNativePreviewRenderer
{
    internal const string UnavailableMessage =
        "VisualTeX 原生公式排版服务不可用，请修复安装后重试。";
    internal const string IncompleteMessage =
        "原生公式排版未完整完成，文档尚未修改；请重试或检查 VisualTeX 排版服务。";
    internal const string FontsUnavailableMessage =
        "VisualTeX 排版服务找不到 MathType 字体，为避免公式显示错误，文档尚未修改。请在 VisualTeX 设置中点击“修复 Office 集成”后重试；如仍失败，请重启电脑。";

    // The renderer refuses (rather than saving placeholder glyphs) when it
    // cannot find MathType's fonts; callers report that case specifically.
    [ThreadStatic] private static bool lastFailureWasFonts;
    // Engine reason per item id that still had no output after the retry.
    [ThreadStatic] private static Dictionary<string, string>? lastFailures;

    internal static bool LastFailureWasFonts => lastFailureWasFonts;

    internal static IReadOnlyDictionary<string, string> LastFailures =>
        lastFailures ?? new Dictionary<string, string>(StringComparer.Ordinal);

    internal static string FailureMessage =>
        lastFailureWasFonts ? FontsUnavailableMessage : DescribeFailures(_ => null);

    // Names the failed formulas (label returns e.g. its LaTeX, or null) and the
    // engine's reason, so a failed batch says which equation to look at.
    internal static string DescribeFailures(Func<string, string?> label)
    {
        if (lastFailureWasFonts) return FontsUnavailableMessage;
        var failures = LastFailures;
        if (failures.Count == 0) return IncompleteMessage;
        var lines = failures.Take(3).Select(failure =>
        {
            var name = label(failure.Key);
            if (!string.IsNullOrWhiteSpace(name))
            {
                name = name!.Trim();
                if (name.Length > 80) name = name.Substring(0, 80) + "…";
            }
            return string.IsNullOrWhiteSpace(name)
                ? $"原因：{failure.Value}"
                : $"{name}\r\n  原因：{failure.Value}";
        });
        var more = failures.Count > 3 ? $"\r\n（另有 {failures.Count - 3} 个公式失败）" : string.Empty;
        return $"原生公式排版未完整完成，文档尚未修改。{failures.Count} 个公式无法排版：\r\n"
            + string.Join("\r\n", lines) + more
            + "\r\n请重试；如仍失败，请检查这些公式或 VisualTeX 排版服务。";
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
            try { File.Delete(WmfPath); } catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    internal static bool IsAvailable => NativeRenderClient.IsAvailable;

    // One pass over the items (protocol-sized requests). PowerPoint layers its
    // own grouping on top; Word uses TryRenderBatchWithRetry.
    internal static bool TryRenderBatch(
        IReadOnlyDictionary<string, byte[]> mtefs,
        string outputDirectory,
        out Dictionary<string, Result> results)
    {
        var rendered = new Dictionary<string, Result>(StringComparer.Ordinal);
        results = rendered;
        lastFailureWasFonts = false;
        lastFailures = new Dictionary<string, string>(StringComparer.Ordinal);
        if (!CanRender(mtefs)) return false;
        try
        {
            Directory.CreateDirectory(outputDirectory);
            lastFailures = RenderAll(mtefs, outputDirectory, rendered);
            RedrawOversizedPreviews(mtefs, outputDirectory, rendered, lastFailures);
            if (lastFailures.Count != 0)
                Trace.WriteLine("mathtype-private-render-incomplete items=" + lastFailures.Count + " reasons="
                    + string.Join(" | ", lastFailures.Values.Distinct()));
            return lastFailures.Count == 0 && rendered.Count == mtefs.Count;
        }
        catch (Exception error)
        {
            Trace.WriteLine("mathtype-private-render-failed " + error.GetType().Name + ": " + error.Message);
            // The caller owns any completed results and decides its document
            // transaction. Never silently replace native layout with another engine.
            foreach (var key in mtefs.Keys.Where(key => !rendered.ContainsKey(key)))
                lastFailures[key] = error.Message;
            return false;
        }
    }

    // The renderer process can crash part-way through a request, and repeatedly
    // for a while after the service is (re)installed; the service starts a fresh
    // worker after every incomplete request. Re-render what is missing in
    // smaller groups (NativePreviewRetry), so a crash costs time, not the batch.
    // LastFailures keeps each still-missing item's reason from its last attempt.
    internal static bool TryRenderBatchWithRetry(
        IReadOnlyDictionary<string, byte[]> mtefs,
        string outputDirectory,
        out Dictionary<string, Result> results)
    {
        var rendered = new Dictionary<string, Result>(StringComparer.Ordinal);
        results = rendered;
        lastFailureWasFonts = false;
        lastFailures = new Dictionary<string, string>(StringComparer.Ordinal);
        if (!CanRender(mtefs)) return false;
        var reasons = new Dictionary<string, string>(StringComparer.Ordinal);
        var fonts = false;
        var pass = 0;
        NativePreviewRetry.Render<Result>(
            mtefs.Keys.ToList(),
            group =>
            {
                TryRenderBatch(
                    group.ToDictionary(id => id, id => mtefs[id], StringComparer.Ordinal),
                    outputDirectory,
                    out var groupResults);
                fonts |= lastFailureWasFonts;
                foreach (var failure in LastFailures) reasons[failure.Key] = failure.Value;
                foreach (var key in groupResults.Keys) reasons.Remove(key);
                foreach (var item in groupResults)
                {
                    if (rendered.TryGetValue(item.Key, out var stale)) stale.Dispose();
                    rendered[item.Key] = item.Value;
                }
                return groupResults;
            },
            new NativePreviewRetryOptions
            {
                GroupSizes = new[] { NativeRenderProtocol.MaxItems, 4, 1 },
                StopNow = () => fonts,
                Retrying = (attempt, missing, size) =>
                {
                    pass = attempt;
                    Trace.WriteLine($"mathtype-private-render-retry pass={attempt} missing={missing} group={size}");
                },
            });
        lastFailureWasFonts = fonts;
        lastFailures = mtefs.Keys.Where(key => !rendered.ContainsKey(key))
            .ToDictionary(key => key, key => reasons.TryGetValue(key, out var reason) ? reason : "排版服务没有返回结果。",
                StringComparer.Ordinal);
        if (pass > 0)
            Trace.WriteLine($"mathtype-private-render-retried passes={pass} missing={lastFailures.Count}");
        return lastFailures.Count == 0;
    }

    private static bool CanRender(IReadOnlyDictionary<string, byte[]> mtefs) =>
        mtefs is not null && mtefs.Count != 0 && IsAvailable
        && !mtefs.Any(pair => string.IsNullOrEmpty(pair.Key)
            || pair.Value is null || pair.Value.Length == 0
            || pair.Value.Length > NativeRenderProtocol.MaxMtefBytes);

    // MathPage writes 16-bit WMF coordinates at 32 units per point, so a preview
    // wider or taller than 1023 pt wraps around (MathType has the same limit),
    // while the dimensions it reports stay exact. Such an equation is drawn
    // again at a reduced size and keeps the exact dimensions of the first
    // render: the preview is vector and Office scales it to that size.
    private const float MaxPreviewPoints = 1000f;

    private static void RedrawOversizedPreviews(
        IReadOnlyDictionary<string, byte[]> mtefs, string outputDirectory,
        Dictionary<string, Result> rendered, Dictionary<string, string> failures)
    {
        var oversized = rendered
            .Where(pair => Math.Max(pair.Value.WidthPt, pair.Value.HeightPt) > MaxPreviewPoints)
            .ToList();
        if (oversized.Count == 0) return;
        var scaled = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var pair in oversized)
        {
            rendered.Remove(pair.Key);
            pair.Value.Dispose();
            var factor = MaxPreviewPoints * 0.9 / Math.Max(pair.Value.WidthPt, pair.Value.HeightPt);
            try { scaled[pair.Key] = MathTypeMtefCodec.ScaleMtefPreferenceSizes(mtefs[pair.Key], factor); }
            catch (InvalidDataException error) { failures[pair.Key] = error.Message; }
        }
        var redrawn = new Dictionary<string, Result>(StringComparer.Ordinal);
        foreach (var failure in RenderAll(scaled, outputDirectory, redrawn))
            failures[failure.Key] = failure.Value;
        foreach (var pair in oversized)
        {
            if (!redrawn.TryGetValue(pair.Key, out var picture)) continue;
            Trace.WriteLine($"mathtype-private-render-scaled widthPt={pair.Value.WidthPt:0.#} heightPt={pair.Value.HeightPt:0.#}");
            rendered[pair.Key] = new Result
            {
                WmfPath = picture.WmfPath,
                WidthPt = pair.Value.WidthPt,
                HeightPt = pair.Value.HeightPt,
                WordPosition = pair.Value.WordPosition,
            };
        }
    }

    // Renders in protocol-sized chunks; returns the reason per item without output.
    // The protocol limits item ids to 256 characters, and callers key formulas
    // by their full LaTeX, so each request carries short ids mapped back here.
    private static Dictionary<string, string> RenderAll(
        IReadOnlyDictionary<string, byte[]> mtefs, string outputDirectory,
        Dictionary<string, Result> results)
    {
        var failures = new Dictionary<string, string>(StringComparer.Ordinal);
        var request = new NativeRenderRequest();
        var keys = new Dictionary<string, string>(StringComparer.Ordinal);
        var byteCount = 0;
        foreach (var pair in mtefs)
        {
            if (request.Items.Count == NativeRenderProtocol.MaxItems
                || byteCount + pair.Value.Length > NativeRenderProtocol.MaxRequestBytes / 2)
            {
                RenderChunk(request, keys, outputDirectory, results, failures);
                request = new NativeRenderRequest();
                keys = new Dictionary<string, string>(StringComparer.Ordinal);
                byteCount = 0;
            }
            var id = request.Items.Count.ToString(System.Globalization.CultureInfo.InvariantCulture);
            keys[id] = pair.Key;
            request.Items.Add(new NativeRenderInput { Id = id, Mtef = pair.Value });
            byteCount += pair.Value.Length;
        }
        if (request.Items.Count != 0) RenderChunk(request, keys, outputDirectory, results, failures);
        return failures;
    }

    private static void RenderChunk(
        NativeRenderRequest request, IReadOnlyDictionary<string, string> keys, string directory,
        Dictionary<string, Result> results, Dictionary<string, string> failures)
    {
        NativeRenderResponse response;
        try
        {
            response = NativeRenderClient.RenderAsync(request, CancellationToken.None)
                .GetAwaiter().GetResult();
        }
        catch (Exception error) when (error is IOException or InvalidDataException
            or TimeoutException or OperationCanceledException or System.ComponentModel.Win32Exception)
        {
            // A lost connection or rejected frame fails this chunk only; the
            // caller's retry sends it again on a fresh connection.
            foreach (var item in request.Items) failures[keys[item.Id]] = error.Message;
            return;
        }
        if (response.Error == NativeRenderProtocol.FontsUnavailableError)
        {
            lastFailureWasFonts = true;
            throw new InvalidDataException(FontsUnavailableMessage);
        }
        foreach (var item in response.Items)
        {
            string path = Path.Combine(directory, "visualtex-native-" + Guid.NewGuid().ToString("N") + ".wmf");
            try { File.WriteAllBytes(path, item.Wmf); }
            catch
            {
                try { File.Delete(path); } catch (IOException) { }
                catch (UnauthorizedAccessException) { }
                throw;
            }
            var key = keys[item.Id];
            if (results.TryGetValue(key, out var stale)) stale.Dispose();
            results[key] = new Result
            {
                WmfPath = path, WidthPt = item.WidthPt,
                HeightPt = item.HeightPt, WordPosition = item.WordPosition,
            };
        }
        var reasons = response.Failures.ToDictionary(failure => failure.Id, failure => failure.Reason, StringComparer.Ordinal);
        var returned = new HashSet<string>(response.Items.Select(output => output.Id), StringComparer.Ordinal);
        foreach (var item in request.Items)
        {
            if (returned.Contains(item.Id)) continue;
            failures[keys[item.Id]] = reasons.TryGetValue(item.Id, out var reason)
                ? reason
                : string.IsNullOrWhiteSpace(response.Error) ? "排版服务没有返回结果。" : response.Error;
        }
    }

    internal static bool TryRender(byte[] mtef, string outputDirectory, out Result result)
    {
        bool complete = TryRenderBatchWithRetry(new Dictionary<string, byte[]> { ["formula"] = mtef },
            outputDirectory, out var results);
        if (!complete)
        {
            foreach (var preview in results.Values) preview.Dispose();
            result = new Result();
            return false;
        }
        result = results["formula"];
        return complete;
    }
}
