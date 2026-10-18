using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed class PreparedWordBulkFormula
{
    internal WordBulkRun Run { get; set; } = new();
    internal OfficeSessionDocument Session { get; set; } = new();
    internal string? MathMl { get; set; }
    internal string? PngPath { get; set; }
    internal string? SvgPath { get; set; }
    internal float? RawRenderWidthPx { get; set; }
    internal float? RawRenderHeightPx { get; set; }
    internal float? RawBaselinePx { get; set; }
    internal string? EmfPath { get; set; }
    internal bool MathTypeNativePreviewAttempted { get; set; }
    internal MathTypeNativePreviewRenderer.Result? MathTypeNativePreview { get; set; }
    // Batch MathType preparation may already have produced and semantically
    // validated the exact standalone Equation Native/CFB payload while building
    // the native preview. Reuse that immutable payload during Word insertion
    // instead of running the same codec + storage round-trip a second time.
    internal MathTypeMtefCodec.RewriteResult? MathTypeStandaloneGenerated { get; set; }
    internal byte[]? MathTypeStandaloneCompoundFile { get; set; }
    internal string? MathTypeStandaloneSemanticSignature { get; set; }
}

internal sealed class RenderedWordBulkFormulaTemplate
{
    internal OfficeSessionDocument Session { get; set; } = new();
    internal string? MathMl { get; set; }
    internal string? PngPath { get; set; }
    internal string? SvgPath { get; set; }
    internal float? RawRenderWidthPx { get; set; }
    internal float? RawRenderHeightPx { get; set; }
    internal float? RawBaselinePx { get; set; }
    internal string? EmfPath { get; set; }
}

internal sealed class WordBulkInsertResult
{
    internal int BlockCount { get; set; }
    internal int FormulaCount { get; set; }
    internal List<string> FormulaIds { get; set; } = new();
}

internal sealed class WordLatexRedrawTarget
{
    internal string Id { get; set; } = Guid.NewGuid().ToString("D");
    internal int RelativeStart { get; set; }
    internal int SourceLength { get; set; }
    internal int AbsoluteStart { get; set; } = -1;
    internal int AbsoluteEnd { get; set; } = -1;
    internal string Latex { get; set; } = string.Empty;
    internal string DisplayMode { get; set; } = "inline";
    internal bool PreserveDisplayParagraphBoundary { get; set; }
    internal bool DedicatedDisplayParagraph { get; set; }
    internal double FontSizePt { get; set; }
    // Set only when the LaTeX source was produced from an inline native object and
    // carries a VisualTeX provenance bookmark for its established Word baseline.
    internal int? SourceInlineWordPosition { get; set; }
    internal float? SourceInlineBottomWhitespacePoints { get; set; }
}

internal sealed class WordLatexRedrawPlan
{
    internal string DocumentId { get; set; } = string.Empty;
    internal int ScopeStart { get; set; }
    internal int ScopeEnd { get; set; }
    internal string SourceText { get; set; } = string.Empty;
    internal bool NumberDisplayFormulas { get; set; }
    internal List<WordLatexRedrawTarget> Targets { get; set; } = new();
    internal Dictionary<string, long> CaptureTimings { get; } =
        new(StringComparer.Ordinal);
    internal List<DeletedNumberedOmmlResidue> DeletedNumberedOmmlResidues { get; set; } = new();
}

/// <summary>
/// Redraw formulas that are still rendering while Word is written. The render
/// producer completes each target as soon as its preview exists; the Word writer
/// takes targets in document order and waits only for one that is not ready.
/// The wait does not pump messages, so Word stays exactly as blocked as during
/// a non-streamed apply and no user edit can interleave with the batch.
/// </summary>
internal sealed class StreamingLatexRedrawFormulas : IDisposable
{
    private readonly Dictionary<string, System.Threading.Tasks.TaskCompletionSource<(PreparedWordBulkFormula? Formula, string? Error)>> _targets;
    private readonly System.Threading.AutoResetEvent _progress = new(false);

    internal StreamingLatexRedrawFormulas(string objectMode, IEnumerable<string> targetIds)
    {
        ObjectMode = objectMode;
        _targets = targetIds.ToDictionary(
            id => id,
            _ => new System.Threading.Tasks.TaskCompletionSource<(PreparedWordBulkFormula? Formula, string? Error)>(
                System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously),
            StringComparer.Ordinal);
    }

    internal string ObjectMode { get; }

    // Writer time spent waiting for renders (UI thread only).
    internal long WaitMilliseconds { get; private set; }

    internal bool Contains(string targetId) => _targets.ContainsKey(targetId);

    internal void Complete(string targetId, PreparedWordBulkFormula formula)
    {
        _targets[targetId].TrySetResult((formula, null));
        _progress.Set();
    }

    internal void Fail(string targetId, string error)
    {
        _targets[targetId].TrySetResult((null, error));
        _progress.Set();
    }

    internal void FailRemaining(string error)
    {
        foreach (var item in _targets.Values)
            item.TrySetResult((null, error));
        _progress.Set();
    }

    internal bool TryTake(
        string targetId,
        TimeSpan timeout,
        out PreparedWordBulkFormula? formula,
        out string? error)
    {
        var task = _targets[targetId].Task;
        if (!task.IsCompleted)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            var handle = _progress.SafeWaitHandle.DangerousGetHandle();
            while (!task.IsCompleted && watch.Elapsed < timeout)
                _ = WaitForSingleObject(handle, 200);
            WaitMilliseconds += watch.ElapsedMilliseconds;
            if (!task.IsCompleted)
            {
                formula = null;
                error = "公式渲染超时。";
                return false;
            }
        }
        (formula, error) = task.Result;
        return formula is not null;
    }

    public void Dispose() => _progress.Dispose();

    // kernel32's wait is used directly: the CLR's managed waits pump COM and
    // window messages on an STA thread, which the Word UI thread is.
    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    private static extern uint WaitForSingleObject(IntPtr handle, uint milliseconds);
}

internal sealed class DeletedNumberedOmmlResidue
{
    internal string FormulaId { get; set; } = string.Empty;
    internal int OwnerStart { get; set; }
    internal int OwnerEnd { get; set; }
}

internal sealed class WordLatexRedrawResult
{
    internal int FormulaCount { get; set; }
    internal long TotalSourceMutationMilliseconds { get; set; }
    internal long TotalInsertMilliseconds { get; set; }
    internal long MaxInsertMilliseconds { get; set; }
    // Timings are kept in actual document-order mutation order. The acceptance
    // logger uses them to detect index-dependent growth without adding file I/O
    // to the hot path.
    internal List<long> SourceMutationMilliseconds { get; } = new();
    internal List<long> InsertMilliseconds { get; } = new();
    internal List<string> InsertMetricDetails { get; } = new();
    internal string? FinalizationMetricDetail { get; set; }
    internal Dictionary<string, long> BatchTimings { get; } =
        new(StringComparer.Ordinal);
    internal List<string> FormulaIds { get; set; } = new();
    internal List<(string TargetId, string Latex, string Error)> Failures { get; } =
        new();
    internal List<string> Warnings { get; } = new();
}

internal sealed class WordFormulaToLatexResult
{
    internal int FormulaCount { get; set; }
    internal List<string> FormulaIds { get; set; } = new();
}
