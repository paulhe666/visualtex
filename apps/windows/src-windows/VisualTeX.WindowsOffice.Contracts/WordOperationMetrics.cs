using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace VisualTeX.WindowsOffice.Contracts;

// Explicit, thread-local measurement session. No Word objects or document data
// are retained; timings are inclusive and must not be summed across nesting.
public sealed class WordOperationMetrics : IDisposable
{
    [ThreadStatic] private static WordOperationMetrics? _current;
    private readonly WordOperationMetrics? _previous;
    private readonly Dictionary<string, Entry> _entries = new Dictionary<string, Entry>();
    public sealed class Entry
    {
        public long Calls { get; internal set; }
        public double Milliseconds { get; internal set; }
        public List<double> SamplesMilliseconds { get; } = new List<double>();
    }
    private sealed class Timing : IDisposable
    {
        private readonly Entry _entry;
        private readonly long _start = Stopwatch.GetTimestamp();
        internal Timing(Entry entry) { _entry = entry; entry.Calls++; }
        public void Dispose()
        {
            var elapsed = (Stopwatch.GetTimestamp() - _start) * 1000.0 / Stopwatch.Frequency;
            _entry.Milliseconds += elapsed;
            _entry.SamplesMilliseconds.Add(elapsed);
        }
    }
    public WordOperationMetrics() { _previous = _current; _current = this; }
    public IReadOnlyDictionary<string, Entry> Entries => _entries;
    public static IDisposable? Measure(string name)
    {
        if (_current is null) return null;
        if (!_current._entries.TryGetValue(name, out var entry))
            _current._entries.Add(name, entry = new Entry());
        return new Timing(entry);
    }
    public void Dispose() { _current = _previous; }
}
