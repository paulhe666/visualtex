using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// A MathType object rewritten at another size (the Ribbon's formula sizes, an edit in VisualTeX) keeps its baseline
// on the line: the run position follows the rewritten equation's own position, and only what the document added to
// the source's own position is kept. Before, the old position was kept as it was, so a display formula made larger
// rose above the line of its tabs and number (and sank below it when made smaller).
public sealed class MathTypeRunPositionTests
{
    [Fact]
    public void AnObjectAtItsOwnPositionTakesTheRewrittenOwnPosition()
    {
        // 10.5 pt source drawn 6 pt below the line by MathPage, rewritten at 16 pt: 9 pt below
        Assert.Equal(-9, MathTypeRunPosition.Rewritten(-6, -6, -9, 1f));
        // made smaller
        Assert.Equal(-4, MathTypeRunPosition.Rewritten(-6, -6, -4, 1f));
    }

    [Fact]
    public void ANudgeTheDocumentAddedIsKept()
    {
        // the user had lowered the object 2 pt more than MathPage's own position
        Assert.Equal(-11, MathTypeRunPosition.Rewritten(-8, -6, -9, 1f));
        // or raised it 1 pt
        Assert.Equal(-8, MathTypeRunPosition.Rewritten(-5, -6, -9, 1f));
    }

    [Fact]
    public void ADocumentDrawingScaleAppliesToBothOwnPositions()
    {
        // the document draws MathType 1.5 times its native size: own positions -6 -> -9 and -10 -> -15
        Assert.Equal(-15, MathTypeRunPosition.Rewritten(-9, -6, -10, 1.5f));
        Assert.Equal(-16, MathTypeRunPosition.Rewritten(-10, -6, -10, 1.5f));
    }

    [Theory]
    [InlineData(float.NaN)]
    [InlineData(0f)]
    [InlineData(-1f)]
    [InlineData(float.PositiveInfinity)]
    public void AnUnusableScaleCountsAsOne(float scale)
    {
        Assert.Equal(-9, MathTypeRunPosition.Rewritten(-6, -6, -9, scale));
    }

    [Fact]
    public void ThePositionStaysWithinWordsLimits()
    {
        Assert.Equal(-MathTypeRunPosition.WordLimitPt, MathTypeRunPosition.Rewritten(-1580, 0, -100, 1f));
        Assert.Equal(MathTypeRunPosition.WordLimitPt, MathTypeRunPosition.Rewritten(1580, 0, 100, 1f));
    }
}
