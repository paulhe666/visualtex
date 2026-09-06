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
