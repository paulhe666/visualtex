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

    internal static bool TryRenderBatch(
        IReadOnlyDictionary<string, byte[]> mtefs,
        string outputDirectory,
        out Dictionary<string, Result> results)
    {
        results = new Dictionary<string, Result>(StringComparer.Ordinal);
        if (mtefs is null || mtefs.Count == 0 || !IsAvailable) return false;
        if (mtefs.Any(pair => string.IsNullOrWhiteSpace(pair.Key) || pair.Key.Length > 256
            || pair.Value is null || pair.Value.Length == 0
            || pair.Value.Length > NativeRenderProtocol.MaxMtefBytes)) return false;

        try
        {
            Directory.CreateDirectory(outputDirectory);
            var request = new NativeRenderRequest();
            var byteCount = 0;
            foreach (var pair in mtefs)
            {
                if (request.Items.Count == NativeRenderProtocol.MaxItems
                    || byteCount + pair.Value.Length > NativeRenderProtocol.MaxRequestBytes / 2)
                {
                    RenderChunk(request, outputDirectory, results);
                    request = new NativeRenderRequest();
                    byteCount = 0;
                }
                request.Items.Add(new NativeRenderInput { Id = pair.Key, Mtef = pair.Value });
                byteCount += pair.Value.Length;
            }
            if (request.Items.Count != 0) RenderChunk(request, outputDirectory, results);
            return results.Count == mtefs.Count;
        }
        catch (Exception error)
        {
            Trace.WriteLine("mathtype-private-render-failed " + error.GetType().Name + ": " + error.Message);
            // The caller owns any completed results and decides its document
            // transaction. Never silently replace native layout with another engine.
            return false;
        }
    }

    private static void RenderChunk(
        NativeRenderRequest request, string directory, Dictionary<string, Result> results)
    {
        var response = NativeRenderClient.RenderAsync(request, CancellationToken.None)
            .GetAwaiter().GetResult();
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
            results.Add(item.Id, new Result
            {
                WmfPath = path, WidthPt = item.WidthPt,
                HeightPt = item.HeightPt, WordPosition = item.WordPosition,
            });
        }
        if (response.Items.Count != request.Items.Count)
            throw new InvalidDataException(IncompleteMessage);
    }

    internal static bool TryRender(byte[] mtef, string outputDirectory, out Result result)
    {
        bool complete = TryRenderBatch(new Dictionary<string, byte[]> { ["formula"] = mtef },
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
