namespace VisualTeX.WordVsto;

/// <summary>
/// Maps a story position captured before a batch of in-place replacements to the
/// position of the same text afterwards. Each replacement is recorded with its
/// original end and the number of characters it added (negative when it removed).
/// Restoring a captured selection by its raw numbers after the source text of
/// several formulas changed length moves its end onto unrelated paragraphs.
/// </summary>
internal sealed class WordCoordinateShifts
{
    private readonly List<(int OriginalEnd, int Delta)> _shifts = new();

    internal void Add(int originalEnd, int delta)
    {
        if (delta != 0) _shifts.Add((originalEnd, delta));
    }

    internal int Map(int originalPosition)
    {
        var mapped = originalPosition;
        foreach (var (originalEnd, delta) in _shifts)
        {
            if (originalEnd <= originalPosition) mapped += delta;
        }
        return mapped;
    }
}
