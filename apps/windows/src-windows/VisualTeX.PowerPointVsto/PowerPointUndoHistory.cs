namespace VisualTeX.PowerPointVsto;

// Which undo entries a failed text edit added, from PowerPoint's undo list
// (newest first) read before and after it. PowerPoint keeps a limited number of
// undo steps (20 by default): once the list is full, every new entry drops the
// oldest one, so the list no longer grows by what the edit added (Word's undo
// list only grows, which is what WordUndoHistorySnapshot relies on). The edit's
// entries are those above the original newest entries, which must still all be
// there, in order, right below them: at least the newest BoundarySize of them,
// or every one when the list was shorter. Repeated labels can only make the
// count smaller than what the edit added, never larger, so recovery then fails
// its final check instead of undoing one of the user's own steps.
internal static class PowerPointUndoHistory
{
    internal const int BoundarySize = 8;

    internal static int CountOwnedEntries(IReadOnlyList<string> original, IReadOnlyList<string> current)
    {
        var boundary = Math.Min(BoundarySize, original.Count);
        for (var owned = Math.Max(0, current.Count - original.Count); owned + boundary <= current.Count; owned++)
        {
            if (current.Skip(owned).SequenceEqual(original.Take(current.Count - owned), StringComparer.Ordinal))
                return owned;
        }
        throw new InvalidDataException(
            "PowerPoint's previous undo boundary changed or was truncated; recovery cannot safely cross it.");
    }
}
