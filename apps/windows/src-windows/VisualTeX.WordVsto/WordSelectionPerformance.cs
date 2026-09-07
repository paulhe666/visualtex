using System.Diagnostics;
using System.Threading;

namespace VisualTeX.WordVsto;

// Optional, read-free event timing. Disabled in normal installations. Never
// obtains a Word RCW or pumps messages while observing a selection callback.
internal sealed class WordSelectionPerformance : IDisposable
{
    private static readonly bool Enabled =
        Environment.GetEnvironmentVariable("VISUALTEX_WORD_SELECTION_PERF") == "1";
    private static long nextId;
    private readonly long id;
    private readonly string operation;
    private readonly Stopwatch watch = Stopwatch.StartNew();
    private long checkpoint;

    private WordSelectionPerformance(string operation)
    {
        this.operation = operation;
        id = Interlocked.Increment(ref nextId);
    }

    internal static WordSelectionPerformance? Start(string operation) =>
        Enabled ? new WordSelectionPerformance(operation) : null;

    internal void Mark(string stage)
    {
        var elapsed = watch.ElapsedTicks;
        WordDoubleClickHook.TraceMessage(
            $"selection-perf id={id} operation={operation} stage={stage} "
            + $"deltaMs={(elapsed - checkpoint) * 1000.0 / Stopwatch.Frequency:F3} "
            + $"totalMs={elapsed * 1000.0 / Stopwatch.Frequency:F3}");
        checkpoint = elapsed;
    }

    public void Dispose() => Mark("complete");
}
