using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

// Word may split one custom record into many native entries while exporting XML.
// Record the native depth and exact boundary labels, not thousands of labels on
// every successful edit. Recovery never undoes below the ORIGINAL DEPTH. If Word
// truncated K old entries, currentDepth-originalDepth underestimates the new
// prefix by K; it cannot authorize undoing a preceding user entry. Boundary labels
// detect branch changes, and the caller's full content/metadata/format snapshot
// still has to verify recovery. A mismatch is an error, never an XML replacement.
internal sealed class WordUndoHistorySnapshot
{
    // Depth is the hard rollback boundary: recovery never authorizes Undo below
    // original.Depth. Labels are secondary branch-change evidence, so sampling a
    // compact head/tail is sufficient and avoids dozens of very expensive
    // CommandBarComboBox.List COM calls on documents with deep undo histories.
    // If Word truncates old entries, depth arithmetic can only under-estimate the
    // owned prefix; the full body/metadata verification then rejects incomplete
    // recovery without crossing a user action.
    private const int BoundarySize = 8;
    private const int TailSize = 4;
    private readonly string documentName;
    private readonly Boundary original;

    private sealed class Boundary
    {
        internal int Depth;
        internal int Offset;
        internal string[] Head = Array.Empty<string>();
        internal string[] Tail = Array.Empty<string>();
    }

    internal WordUndoHistorySnapshot(Document document)
    {
        documentName = document.FullName;
        original = ReadBoundary(document, originalDepth: null);
    }

    internal int UndoChanges(Document document)
    {
        if (!string.Equals(document.FullName, documentName, StringComparison.Ordinal))
            throw new InvalidOperationException("The document changed before undo recovery.");
        var current = ReadBoundary(document, original.Depth);
        var count = CountOwnedPrefixAtBoundary(original.Depth, original.Head,
            current.Depth, current.Head);
        if (!original.Tail.SequenceEqual(current.Tail, StringComparer.Ordinal))
            throw new InvalidDataException("The original undo-history tail changed; recovery was not attempted.");
        WordDoubleClickHook.TraceMessage(
            $"word-undo-history-recovery original={original.Depth} current={current.Depth} owned={count}");
        if (count == 0) return 0;
        object times = count;
        if (!document.Undo(ref times))
            throw new InvalidOperationException("Word could not undo all changes made by the failed operation.");
        var restored = ReadBoundary(document, originalDepth: null);
        if (restored.Depth != original.Depth
            || !restored.Head.SequenceEqual(original.Head, StringComparer.Ordinal)
            || !restored.Tail.SequenceEqual(original.Tail, StringComparer.Ordinal))
            throw new InvalidDataException("Word did not restore the original undo-history boundary.");
        return count;
    }

    // Retained for full-history diagnostics and existing pure-data regression tests.
    internal static int CountOwnedPrefix(IReadOnlyList<string> original, IReadOnlyList<string> current)
    {
        var count = current.Count - original.Count;
        if (count < 0 || !current.Skip(Math.Max(0, count)).SequenceEqual(original, StringComparer.Ordinal))
            throw new InvalidDataException(
                "Word's previous undo history changed or was truncated; recovery cannot safely cross that boundary.");
        return count;
    }

    internal static int CountOwnedPrefixAtBoundary(int originalDepth, IReadOnlyList<string> originalHead,
        int currentDepth, IReadOnlyList<string> currentAtBoundary)
    {
        if (originalDepth < 0 || currentDepth < originalDepth
            || originalHead.Count != Math.Min(BoundarySize, originalDepth)
            || !originalHead.SequenceEqual(currentAtBoundary, StringComparer.Ordinal))
            throw new InvalidDataException("Word's previous undo boundary changed; recovery cannot safely cross it.");
        return currentDepth - originalDepth;
    }

    private static Boundary ReadBoundary(Document document, int? originalDepth)
    {
        var watch = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF") == "1"
            ? System.Diagnostics.Stopwatch.StartNew() : null;
        Microsoft.Office.Interop.Word.Application? application = null;
        Document? active = null;
        CommandBars? bars = null;
        CommandBar? standard = null;
        CommandBarControl? control = null;
        try
        {
            application = document.Application;
            active = application.ActiveDocument;
            if (!string.Equals(active.FullName, document.FullName, StringComparison.Ordinal))
                throw new InvalidOperationException("The undo history belongs to a different active document.");
            bars = application.CommandBars;
            standard = bars["Standard"];
            control = standard.FindControl(Id: 128)
                ?? throw new InvalidOperationException("Word's native undo-history control is unavailable.");
            var history = control as CommandBarComboBox
                ?? throw new InvalidOperationException("Word's native undo control has no readable history.");
            var enabled = bars.GetEnabledMso("Undo");
            if (history.Enabled != enabled)
                throw new InvalidOperationException("Word's undo command and history disagree about availability.");
            if (!enabled)
            {
                if (originalDepth.GetValueOrDefault() != 0)
                    throw new InvalidDataException("Word lost the original undo history.");
                return new Boundary();
            }
            var count = history.ListCount;
            if (count <= 0)
                throw new InvalidDataException("Word enables Undo but exposes no readable history.");
            var retainedDepth = originalDepth ?? count;
            if (count < retainedDepth)
                throw new InvalidDataException("Word truncated the previous undo-history boundary.");
            var offset = count - retainedDepth;
            var headCount = Math.Min(BoundarySize, retainedDepth);
            var tailCount = Math.Min(TailSize, Math.Max(0, retainedDepth - headCount));
            var result = new Boundary { Depth = count, Offset = offset,
                Head = new string[headCount], Tail = new string[tailCount] };
            for (var i = 0; i < headCount; i++) result.Head[i] = history.get_List(offset + i + 1);
            for (var i = 0; i < tailCount; i++) result.Tail[i] = history.get_List(count - tailCount + i + 1);
            if (history.ListCount != count)
                throw new InvalidDataException("Word's undo history changed while its boundary was captured.");
            if (watch is not null)
                WordDoubleClickHook.TraceMessage($"undo-history-perf depth={count} read={headCount + tailCount} offset={offset} totalMs={watch.ElapsedMilliseconds}");
            return result;
        }
        finally
        {
            Release(control); Release(standard); Release(bars); Release(active); Release(application);
        }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
