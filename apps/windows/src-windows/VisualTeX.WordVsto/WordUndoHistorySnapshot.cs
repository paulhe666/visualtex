using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

// Word can end a CustomRecord while serializing WordOpenXML. An operation may
// therefore own several native undo entries even though its outer record object
// is still alive. Capture the complete existing stack before any mutation and
// undo only a proven new prefix; never guess a count or cross the user's history.
internal sealed class WordUndoHistorySnapshot
{
    private readonly string documentName;
    private readonly IReadOnlyList<string> original;

    internal WordUndoHistorySnapshot(Document document)
    {
        documentName = document.FullName;
        original = ReadRecords(document);
    }

    internal int UndoChanges(Document document)
    {
        if (!string.Equals(document.FullName, documentName, StringComparison.Ordinal))
            throw new InvalidOperationException("The document changed before undo recovery.");
        var current = ReadRecords(document);
        var count = CountOwnedPrefix(original, current);
        WordDoubleClickHook.TraceMessage(
            $"word-undo-history-recovery original={original.Count} current={current.Count} owned={count}");
        if (count == 0) return 0;
        object times = count;
        if (!document.Undo(ref times))
            throw new InvalidOperationException("Word could not undo all changes made by the failed operation.");
        if (!ReadRecords(document).SequenceEqual(original, StringComparer.Ordinal))
            throw new InvalidDataException("Word did not restore the original undo-history boundary.");
        return count;
    }

    internal static int CountOwnedPrefix(IReadOnlyList<string> original, IReadOnlyList<string> current)
    {
        var count = current.Count - original.Count;
        if (count < 0 || !current.Skip(Math.Max(0, count)).SequenceEqual(original, StringComparer.Ordinal))
            throw new InvalidDataException(
                "Word's previous undo history changed or was truncated; recovery cannot safely cross that boundary.");
        return count;
    }

    private static IReadOnlyList<string> ReadRecords(Document document)
    {
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
            // FindControl on the collection can return the plain menu button,
            // whose caption exposes only the newest action. The standard bar's
            // built-in split dropdown exposes the complete ordered native list.
            control = standard.FindControl(Id: 128)
                ?? throw new InvalidOperationException("Word's native undo-history control is unavailable.");
            var history = control as CommandBarComboBox
                ?? throw new InvalidOperationException("Word's native undo control has no readable history.");
            var commandEnabled = bars.GetEnabledMso("Undo");
            if (history.Enabled != commandEnabled)
                throw new InvalidOperationException("Word's undo command and history disagree about availability.");
            // On a new document Word disables Undo and ListCount throws E_FAIL.
            // Read the command state first; never reinterpret a failed list read
            // as an empty history when Word says an action can be undone.
            if (!commandEnabled) return Array.Empty<string>();
            var count = history.ListCount;
            if (count <= 0)
                throw new InvalidDataException("Word enables Undo but exposes no readable history.");
            var result = new string[count];
            for (var index = 1; index <= count; index++) result[index - 1] = history.get_List(index);
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
