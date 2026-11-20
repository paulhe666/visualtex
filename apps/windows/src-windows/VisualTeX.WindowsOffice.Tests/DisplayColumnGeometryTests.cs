using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

// Display formula tab stops come from the column the line is in (Word measures tab positions from the column's left
// edge): Letter paper with 1" margins, two equal 216 pt columns, or unequal 288 + 144 pt columns 36 pt apart.
public sealed class DisplayColumnGeometryTests
{
    private static readonly float[] Equal = { 216f, 216f };
    private static readonly float[] Unequal = { 288f, 144f };
    private static readonly float[] Gap = { 36f, 0f };

    [Fact]
    public void OneOrEqualColumnsGiveTheirWidthWherever()
    {
        Assert.Equal(468f, DisplayColumnGeometry.ColumnAt(new[] { 468f }, new[] { 0f }, null));
        Assert.Equal(216f, DisplayColumnGeometry.ColumnAt(Equal, Gap, 10f));
        Assert.Equal(216f, DisplayColumnGeometry.ColumnAt(Equal, Gap, 300f));
        Assert.Equal(216f, DisplayColumnGeometry.ColumnAt(Equal, Gap, null));
        Assert.True(DisplayColumnGeometry.AreEqual(new[] { 216f, 216.02f }));
        Assert.False(DisplayColumnGeometry.AreEqual(Unequal));
    }

    [Fact]
    public void UnequalColumnsArePickedByPosition()
    {
        Assert.Equal(288f, DisplayColumnGeometry.ColumnAt(Unequal, Gap, 0f));
        Assert.Equal(288f, DisplayColumnGeometry.ColumnAt(Unequal, Gap, 287.5f));
        // the gap after a column still belongs to it
        Assert.Equal(288f, DisplayColumnGeometry.ColumnAt(Unequal, Gap, 320f));
        Assert.Equal(144f, DisplayColumnGeometry.ColumnAt(Unequal, Gap, 324f));
        Assert.Equal(144f, DisplayColumnGeometry.ColumnAt(Unequal, Gap, 460f));
        // past the last column: still the last
        Assert.Equal(144f, DisplayColumnGeometry.ColumnAt(Unequal, Gap, 900f));
    }

    [Fact]
    public void WithoutAPositionTheNarrowestColumnIsSafe()
    {
        Assert.Equal(144f, DisplayColumnGeometry.ColumnAt(Unequal, Gap, null));
        Assert.Equal(144f, DisplayColumnGeometry.ColumnAt(Unequal, Gap, float.NaN));
        Assert.Equal(144f, DisplayColumnGeometry.ColumnAt(Unequal, Gap, -20f));
        Assert.Equal(144f, DisplayColumnGeometry.ColumnAt(new[] { 144f, 288f }, Gap, null));
    }

    [Fact]
    public void StopsAreTheColumnsCentreAndRightEdge()
    {
        Assert.Equal((108f, 216f), DisplayColumnGeometry.TabStops(216f));
        Assert.Equal((144f, 288f), DisplayColumnGeometry.TabStops(288f));
        Assert.Equal((72f, 144f), DisplayColumnGeometry.TabStops(144f));
        // a page-wide stop pair (the old rule) would put the formula at 234 pt: past a 216 pt column's edge
        Assert.True(DisplayColumnGeometry.TabStops(468f).Center > 216f);
    }

    [Fact]
    public void IndentsNarrowTheLineTheStopsAreCentredIn()
    {
        // first line starts at 18 - 9 = 9 pt, right edge 216 - 9 = 207 pt
        Assert.Equal((108f, 207f), DisplayColumnGeometry.TabStops(216f, leftIndent: 18f, firstLineIndent: -9f, rightIndent: 9f));
        // a hanging indent past the left edge counts as zero; a right indent wider than the column keeps a 1 pt line
        Assert.Equal((108f, 216f), DisplayColumnGeometry.TabStops(216f, leftIndent: 0f, firstLineIndent: -20f));
        var squeezed = DisplayColumnGeometry.TabStops(216f, leftIndent: 300f, rightIndent: 500f);
        Assert.Equal(216f, squeezed.Right);
        Assert.True(squeezed.Center <= squeezed.Right);
    }

    [Fact]
    public void AnUnreadableWidthFallsBackToLetterTextWidth()
    {
        Assert.Equal((234f, 468f), DisplayColumnGeometry.TabStops(float.NaN));
        Assert.Equal((234f, 468f), DisplayColumnGeometry.TabStops(0f));
    }
}
