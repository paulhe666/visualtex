namespace VisualTeX.WordVsto;

// Shared by the Word and PowerPoint add-ins: the retry policy for the private
// MathType renderer (VisualTeX.NativeRender service).
// MathType's renderer can crash part-way through a long run, and right after its
// service starts (its own state, not the equation), losing the whole request.
// Render in moderate groups, re-render whatever a pass did not return in smaller
// groups and then one by one, so a crash costs time, not the batch. Extra
// one-by-one passes run only while the renderer is recovering (the previous pass
// rendered something), and everything stops once nothing has rendered for a
// while, so a broken renderer fails fast. Whatever is still missing at the end
// is reported per equation by the caller.
internal sealed class NativePreviewRetryOptions
{
    internal int[] GroupSizes { get; set; } = { 24, 4, 1 };
    internal int RecoveryPasses { get; set; } = 3;
    // A freshly started render service was observed crashing for about a minute.
    internal TimeSpan GiveUpWithoutProgress { get; set; } = TimeSpan.FromSeconds(90);
    internal TimeSpan RetryPause { get; set; } = TimeSpan.FromSeconds(1.5);
    internal Func<DateTime> Now { get; set; } = () => DateTime.UtcNow;
    // Waits between passes; throws when the batch is cancelled.
    internal Action<TimeSpan> Pause { get; set; } = System.Threading.Thread.Sleep;
    // True when a failure must end the batch at once (e.g. missing fonts).
    internal Func<bool> StopNow { get; set; } = () => false;
    internal Action<int, int>? Progress { get; set; }
    internal Action<int, int, int>? Retrying { get; set; }
}

internal static class NativePreviewRetry
{
    internal static Dictionary<string, TResult> Render<TResult>(
        IReadOnlyList<string> ids,
        Func<IReadOnlyList<string>, IReadOnlyDictionary<string, TResult>> renderGroup,
        NativePreviewRetryOptions options)
    {
        var rendered = new Dictionary<string, TResult>(StringComparer.Ordinal);
        var passes = options.GroupSizes.Concat(Enumerable.Repeat(1, options.RecoveryPasses)).ToArray();
        var lastProgress = options.Now();
        bool Stalled() => options.Now() - lastProgress > options.GiveUpWithoutProgress;
        var renderedByPreviousPass = 0;
        for (var pass = 0; pass < passes.Length && !Stalled(); pass++)
        {
            var size = Math.Max(1, passes[pass]);
            var missing = ids.Where(id => !rendered.ContainsKey(id)).ToList();
            if (missing.Count == 0) break;
            if (pass >= options.GroupSizes.Length && renderedByPreviousPass == 0) break;
            if (pass > 0)
            {
                options.Retrying?.Invoke(pass, missing.Count, size);
                options.Pause(options.RetryPause);
            }
            var before = rendered.Count;
            for (var offset = 0; offset < missing.Count && !Stalled(); offset += size)
            {
                options.Progress?.Invoke(rendered.Count, ids.Count);
                var group = missing.Skip(offset).Take(size).ToList();
                var results = renderGroup(group);
                var added = 0;
                foreach (var pair in results)
                {
                    if (!group.Contains(pair.Key, StringComparer.Ordinal) || rendered.ContainsKey(pair.Key)) continue;
                    rendered[pair.Key] = pair.Value;
                    added++;
                }
                if (added > 0) lastProgress = options.Now();
                if (options.StopNow()) return rendered;
            }
            renderedByPreviousPass = rendered.Count - before;
        }
        options.Progress?.Invoke(rendered.Count, ids.Count);
        return rendered;
    }
}
