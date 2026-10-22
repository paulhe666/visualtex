using System.Diagnostics;
using System.Globalization;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

// Office owns document insertion and equation encoding, not the native engine.
// There is deliberately no installed-MathType discovery, prewarm, process launch
// or process cleanup in this adapter. Those lifetimes belong to the private service.
internal static class MathTypeNativePreviewRenderer
{
    // Evaluated at use time so the current VisualTeX language applies.
    private static string T(string chinese, string english) => OfficePluginLanguage.Text(chinese, english);

    internal static string UnavailableMessage => T(
        "VisualTeX 原生公式排版服务（VisualTeX.NativeRender）没有运行或未安装，请在 VisualTeX 设置中修复 Office 集成后重试。",
        "VisualTeX's native equation renderer (the VisualTeX.NativeRender service) is not running or not installed. Repair the Office integration in VisualTeX settings and try again.");
    internal static string IncompleteMessage => T(
        "原生公式排版未完整完成，文档尚未修改；请重试或检查 VisualTeX 排版服务。",
        "Native equation rendering did not complete, so the document was not changed. Try again or check VisualTeX's renderer service.");
    internal static string FontsUnavailableMessage => T(
        "VisualTeX 排版服务找不到 MathType 字体，为避免公式显示错误，文档尚未修改。请在 VisualTeX 设置中点击“修复 Office 集成”后重试；如仍失败，请重启电脑。",
        "VisualTeX's renderer cannot find MathType's fonts, so the document was not changed (its symbols would display wrongly). In VisualTeX settings, click “Repair integration” and try again; if it still fails, restart Windows.");
    // The service answered but listed neither an output nor a reason for the item.
    internal static string NoResultMessage => T(
        "排版服务没有返回该公式的结果，也没有说明原因。",
        "The renderer returned no result for this equation and gave no reason.");
    private static string EmptyMtefMessage => T(
        "该公式的 MathType 数据（MTEF）为空，无法排版。",
        "This equation's MathType data (MTEF) is empty, so it cannot be rendered.");
    private static string MissingIdMessage => T(
        "内部错误：该公式没有标识，无法发送给排版服务。",
        "Internal error: the equation has no identifier, so it cannot be sent to the renderer.");
    private static string TooLargeMessage(int bytes) => T(
        $"该公式的 MathType 数据（MTEF）有 {bytes.ToString("N0", CultureInfo.InvariantCulture)} 字节，超过排版服务允许的上限 {NativeRenderProtocol.MaxMtefBytes.ToString("N0", CultureInfo.InvariantCulture)} 字节。",
        $"This equation's MathType data (MTEF) is {bytes.ToString("N0", CultureInfo.InvariantCulture)} bytes, more than the renderer's limit of {NativeRenderProtocol.MaxMtefBytes.ToString("N0", CultureInfo.InvariantCulture)} bytes.");
    private static string TimeoutMessage => T(
        $"排版服务没有在 {NativeRenderProtocol.RequestTimeoutMs / 1000} 秒内完成这一组公式（排版进程可能已停止响应）。",
        $"The renderer did not finish this group of equations within {NativeRenderProtocol.RequestTimeoutMs / 1000} s (its process may have stopped responding).");
    private const int DescribedFailureLimit = 15;

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
        var reasonLabel = T("原因：", "Reason: ");
        var lines = failures.Take(DescribedFailureLimit).Select(failure =>
        {
            var name = label(failure.Key);
            if (!string.IsNullOrWhiteSpace(name))
            {
                name = name!.Replace("\r", " ").Replace("\n", " ").Trim();
                if (name.Length > 80) name = name.Substring(0, 80) + "…";
            }
            var reason = OfficePluginLanguage.SafeErrorMessage(failure.Value);
            return string.IsNullOrWhiteSpace(name)
                ? reasonLabel + reason
                : $"{name}\r\n  {reasonLabel}{reason}";
        });
        var hidden = failures.Count - DescribedFailureLimit;
        var more = hidden > 0
            ? "\r\n" + T($"（另有 {hidden} 个公式失败）", $"(and {hidden} more equation(s) failed)")
            : string.Empty;
        return T(
                $"原生公式排版未完整完成，文档尚未修改。{failures.Count} 个公式无法排版：",
                $"Native equation rendering did not complete, so the document was not changed. {failures.Count} equation(s) could not be rendered:")
            + "\r\n" + string.Join("\r\n", lines) + more
            + "\r\n" + T(
                "请重试；如仍失败，请检查这些公式或 VisualTeX 排版服务。",
                "Try again; if it still fails, check these equations or VisualTeX's renderer service.");
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
        var failures = new Dictionary<string, string>(StringComparer.Ordinal);
        lastFailures = failures;
        // An equation the renderer would reject gets its own reason; the others
        // are still rendered.
        var renderable = Renderable(mtefs, failures);
        if (renderable.Count == 0) return false;
        try
        {
            Directory.CreateDirectory(outputDirectory);
            foreach (var failure in RenderAll(renderable, outputDirectory, rendered))
                failures[failure.Key] = failure.Value;
            RedrawOversizedPreviews(renderable, outputDirectory, rendered, failures);
            if (failures.Count != 0)
                Trace.WriteLine("mathtype-private-render-incomplete items=" + failures.Count + " reasons="
                    + string.Join(" | ", failures.Values.Distinct()));
            return failures.Count == 0 && rendered.Count == mtefs.Count;
        }
        catch (Exception error)
        {
            Trace.WriteLine("mathtype-private-render-failed " + error.GetType().Name + ": " + error.Message);
            // The caller owns any completed results and decides its document
            // transaction. Never silently replace native layout with another engine.
            var reason = OfficeErrorText.Describe(error);
            foreach (var key in renderable.Keys.Where(key => !rendered.ContainsKey(key)))
                failures[key] = reason;
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
        var rejected = new Dictionary<string, string>(StringComparer.Ordinal);
        lastFailures = rejected;
        var renderable = Renderable(mtefs, rejected);
        if (renderable.Count == 0) return false;
        var reasons = new Dictionary<string, string>(StringComparer.Ordinal);
        var fonts = false;
        var pass = 0;
        NativePreviewRetry.Render<Result>(
            renderable.Keys.ToList(),
            group =>
            {
                TryRenderBatch(
                    group.ToDictionary(id => id, id => renderable[id], StringComparer.Ordinal),
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
        var missing = renderable.Keys.Where(key => !rendered.ContainsKey(key))
            .ToDictionary(key => key, key => reasons.TryGetValue(key, out var reason) ? reason : NoResultMessage,
                StringComparer.Ordinal);
        foreach (var rejection in rejected) missing[rejection.Key] = rejection.Value;
        lastFailures = missing;
        if (pass > 0)
            Trace.WriteLine($"mathtype-private-render-retried passes={pass} missing={lastFailures.Count}");
        return lastFailures.Count == 0;
    }

    // The equations the renderer accepts; every other one is recorded in
    // failures with its specific reason (empty, too large, service missing).
    private static Dictionary<string, byte[]> Renderable(
        IReadOnlyDictionary<string, byte[]>? mtefs, Dictionary<string, string> failures)
    {
        var renderable = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        if (mtefs is null || mtefs.Count == 0) return renderable;
        var available = IsAvailable;
        foreach (var pair in mtefs)
        {
            if (string.IsNullOrEmpty(pair.Key)) failures[pair.Key ?? string.Empty] = MissingIdMessage;
            else if (pair.Value is null || pair.Value.Length == 0) failures[pair.Key] = EmptyMtefMessage;
            else if (pair.Value.Length > NativeRenderProtocol.MaxMtefBytes) failures[pair.Key] = TooLargeMessage(pair.Value.Length);
            else if (!available) failures[pair.Key] = UnavailableMessage;
            else renderable[pair.Key] = pair.Value;
        }
        return renderable;
    }

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
            catch (InvalidDataException error)
            {
                failures[pair.Key] = T(
                    $"该公式超过 {MaxPreviewPoints:0} 磅，缩小后重新排版失败：{OfficeErrorText.Describe(error)}",
                    $"The equation is larger than {MaxPreviewPoints:0} pt and could not be scaled down for rendering: {OfficeErrorText.Describe(error)}");
            }
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
            var reason = error is TimeoutException or OperationCanceledException
                ? TimeoutMessage
                : T("与排版服务通信失败：", "Communication with the renderer failed: ") + OfficeErrorText.Describe(error);
            foreach (var item in request.Items) failures[keys[item.Id]] = reason;
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
                : string.IsNullOrWhiteSpace(response.Error) ? NoResultMessage : response.Error;
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
