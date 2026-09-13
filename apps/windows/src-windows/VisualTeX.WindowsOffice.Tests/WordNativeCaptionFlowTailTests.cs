using System.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordNativeCaptionFlowTailTests
{
    private const string FormulaId = "11111111-2222-3333-4444-555555555555";

    [Fact]
    public void BookmarkNameFitsWordLimit()
    {
        var name = WordEquationNumbering.NativeCaptionFlowTailBookmarkName(FormulaId);
        Assert.Equal("VTEqFT_11111111222233334444555555555555", name);
        Assert.True(name.Length <= 40);
    }

    [Fact]
    public void ExactOneCharacterEmptyBodyParagraphIsOwned()
        => Assert.True(WordEquationNumbering.IsExactOwnedNativeCaptionFlowTail(
            40, 40, 41, "\r", 40, 41, false, 0, 0, 0, 0, 0));

    [Theory]
    [InlineData(39, 40, 41, "\r", 40, 41, false, 0, 0, 0, 0, 0)]
    [InlineData(40, 40, 42, "\r", 40, 42, false, 0, 0, 0, 0, 0)]
    [InlineData(40, 40, 41, "x\r", 40, 41, false, 0, 0, 0, 0, 0)]
    [InlineData(40, 40, 41, "\r", 39, 41, false, 0, 0, 0, 0, 0)]
    [InlineData(40, 40, 41, "\r", 40, 41, true, 0, 0, 0, 0, 0)]
    [InlineData(40, 40, 41, "\r", 40, 41, false, 1, 0, 0, 0, 0)]
    [InlineData(40, 40, 41, "\r", 40, 41, false, 0, 1, 0, 0, 0)]
    [InlineData(40, 40, 41, "\r", 40, 41, false, 0, 0, 1, 0, 0)]
    [InlineData(40, 40, 41, "\r", 40, 41, false, 0, 0, 0, 1, 0)]
    [InlineData(40, 40, 41, "\r", 40, 41, false, 0, 0, 0, 0, 1)]
    public void ModifiedOrStructuralParagraphIsNotOwned(
        int expectedStart,
        int start,
        int end,
        string text,
        int paragraphStart,
        int paragraphEnd,
        bool withinTable,
        int tables,
        int shapes,
        int maths,
        int fields,
        int frames)
        => Assert.False(WordEquationNumbering.IsExactOwnedNativeCaptionFlowTail(
            expectedStart, start, end, text, paragraphStart, paragraphEnd,
            withinTable, tables, shapes, maths, fields, frames));

    [Fact]
    public void RollbackSnapshotOwnsCaptionFlowTailIdentity()
    {
        var names = WordBookmarkRecoverySnapshot.NamesForFormula(FormulaId).ToArray();
        Assert.Contains("VTEqFT_11111111222233334444555555555555", names);
    }
}
