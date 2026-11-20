using VisualTeX.PowerPointVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// Which undo entries a failed PowerPoint text edit may undo, also once
// PowerPoint's undo list is full and drops its oldest entries.
public sealed class PowerPointUndoHistoryTests
{
    private static string[] Entries(string prefix, int count) =>
        Enumerable.Range(1, count).Select(index => $"{prefix} {index}").ToArray();

    [Fact]
    public void AGrowingListOwnsWhatWasAdded()
    {
        var before = new[] { "User typing", "User table" };
        var after = new[] { "Paste", "Typing", "Paste", "User typing", "User table" };
        Assert.Equal(3, PowerPointUndoHistory.CountOwnedEntries(before, after));
        Assert.Equal(0, PowerPointUndoHistory.CountOwnedEntries(before, before));
        Assert.Equal(2, PowerPointUndoHistory.CountOwnedEntries(Array.Empty<string>(), new[] { "Paste", "Typing" }));
    }

    [Fact]
    public void AFullListThatDroppedItsOldestEntriesStillOwnsWhatWasAdded()
    {
        // 20 steps (PowerPoint's default limit); the edit adds 3, the 3 oldest go.
        var before = Entries("User", 20);
        var after = new[] { "Paste", "Typing", "Paste" }.Concat(before.Take(17)).ToArray();
        Assert.Equal(20, after.Length);
        Assert.Equal(3, PowerPointUndoHistory.CountOwnedEntries(before, after));
        // A list one short of full: 2 added, 1 dropped.
        var almost = Entries("User", 19);
        Assert.Equal(2, PowerPointUndoHistory.CountOwnedEntries(almost, new[] { "Paste", "Paste" }.Concat(almost.Take(18)).ToArray()));
    }

    [Fact]
    public void RepeatedLabelsNeverAuthorizeMoreThanWasAdded()
    {
        var before = Enumerable.Repeat("Paste", 20).ToArray();
        var after = Enumerable.Repeat("Paste", 20).ToArray();
        // 2 entries were added and 2 dropped, but the labels cannot tell: the
        // count may come out smaller (recovery then fails its final check), never larger.
        Assert.InRange(PowerPointUndoHistory.CountOwnedEntries(before, after), 0, 2);
    }

    [Fact]
    public void AChangedOrLostBoundaryIsNeverCrossed()
    {
        var before = Entries("User", 20);
        // A user step below the new entries changed.
        var changed = new[] { "Paste" }.Concat(before.Take(19)).ToArray();
        changed[5] = "Something else";
        Assert.Throws<InvalidDataException>(() => PowerPointUndoHistory.CountOwnedEntries(before, changed));
        // So many entries were added that fewer than 8 of the user's steps are left to prove the boundary.
        var flooded = Entries("Paste", 15).Concat(before.Take(5)).ToArray();
        Assert.Throws<InvalidDataException>(() => PowerPointUndoHistory.CountOwnedEntries(before, flooded));
        // The list lost the user's newest step.
        Assert.Throws<InvalidDataException>(() => PowerPointUndoHistory.CountOwnedEntries(
            new[] { "User typing", "User table" }, new[] { "Paste", "User table" }));
    }
}
