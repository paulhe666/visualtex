using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordUndoHistorySnapshotTests
{
    [Fact]
    public void SplitRecordsIncludeAllNewActionsAndPreserveOlderUserActions()
    {
        var before = new[] { "User typing", "User table" };
        var after = new[] { "Insert OMML", "Range.Delete", "Font.Size", "Outer conversion", "User typing", "User table" };
        Assert.Equal(4, WordUndoHistorySnapshot.CountOwnedPrefix(before, after));
    }

    [Fact]
    public void BoundedCheckpointDoesNotScaleWithNativeHistoryDepth()
    {
        var head = Enumerable.Range(0, 8).Select(i => "Prior action " + i).ToArray();
        Assert.Equal(47, WordUndoHistorySnapshot.CountOwnedPrefixAtBoundary(100000, head, 100047, head));
        Assert.Equal(0, WordUndoHistorySnapshot.CountOwnedPrefixAtBoundary(100000, head, 100000, head));
    }

    [Fact]
    public void BoundedCheckpointRejectsTruncationAndChangedBoundary()
    {
        var head = new[] { "User typing", "User table" };
        Assert.Throws<InvalidDataException>(() => WordUndoHistorySnapshot.CountOwnedPrefixAtBoundary(2, head, 1, head));
        Assert.Throws<InvalidDataException>(() => WordUndoHistorySnapshot.CountOwnedPrefixAtBoundary(2, head, 3, head.Reverse().ToArray()));
        Assert.Throws<InvalidDataException>(() => WordUndoHistorySnapshot.CountOwnedPrefixAtBoundary(1000, head, 1001, head));
        Assert.Equal(0, WordUndoHistorySnapshot.CountOwnedPrefixAtBoundary(0, Array.Empty<string>(), 0, Array.Empty<string>()));
        Assert.Equal(5, WordUndoHistorySnapshot.CountOwnedPrefixAtBoundary(0, Array.Empty<string>(), 5, Array.Empty<string>()));
    }

    [Fact]
    public void DepthLimitedUndoCannotCrossTheOriginalUserHistoryEvenIfOldEntriesWereLost()
    {
        var head = Enumerable.Repeat("Repeated native action", 8).ToArray();
        const int original = 1000;
        for (var lost = 0; lost <= original; lost += 7)
        for (var added = 0; added <= 1200; added += 13)
        {
            var current = original - lost + added;
            if (current < original) continue; // The production guard rejects this state.
            var authorized = WordUndoHistorySnapshot.CountOwnedPrefixAtBoundary(original, head, current, head);
            Assert.InRange(authorized, 0, added);
        }
    }

    [Fact]
    public void NoMutationDoesNotUndoAUserAction()
    {
        var before = new[] { "User typing" };
        Assert.Equal(0, WordUndoHistorySnapshot.CountOwnedPrefix(before, before));
    }

    [Fact]
    public void TruncatedOrReorderedHistoryCannotBeUndone()
    {
        foreach (var after in new[]
                 {
                     new[] { "New operation", "User typing" },
                     new[] { "New operation", "User table", "User typing" },
                     new[] { "User table" },
                 })
            Assert.Throws<InvalidDataException>(() => WordUndoHistorySnapshot.CountOwnedPrefix(
                new[] { "User typing", "User table" }, after));
    }
}
