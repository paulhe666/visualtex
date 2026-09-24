using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordDoubleClickRoutingTests
{
    [Fact]
    public void NativeOleOpensVisualTeXEditor()
    {
        var selection = FormulaSelection(FormulaOleContract.NativeOleMode);

        Assert.True(WordDoubleClickRouting.ShouldOpenVisualTeX(selection));
    }

    [Fact]
    public void CrossPlatformPictureOpensVisualTeXEditor()
    {
        var selection = FormulaSelection(FormulaOleContract.CrossPlatformPictureMode);

        Assert.True(WordDoubleClickRouting.ShouldOpenVisualTeX(selection));
    }

    [Fact]
    public void VisualTeXWordOmmlOpensVisualTeXEditor()
    {
        var selection = FormulaSelection(FormulaOleContract.WordOmmlMode);

        Assert.True(WordDoubleClickRouting.ShouldOpenVisualTeX(selection));
    }

    [Fact]
    public void OrdinaryWordSelectionDoesNothing()
    {
        Assert.False(WordDoubleClickRouting.ShouldOpenVisualTeX(null));
        Assert.False(WordDoubleClickRouting.ShouldOpenVisualTeX(new OfficeSelection()));
    }

    [Fact]
    public void DesktopHookAcceptsOnlyItsOwningWordProcess()
    {
        Assert.True(WordDoubleClickRouting.ForegroundProcessBelongsToOwner(105516u, 105516));
        Assert.False(WordDoubleClickRouting.ForegroundProcessBelongsToOwner(45744u, 105516));
        Assert.False(WordDoubleClickRouting.ForegroundProcessBelongsToOwner(0u, 105516));
        Assert.False(WordDoubleClickRouting.ForegroundProcessBelongsToOwner(105516u, 0));
    }

    [Theory]
    [InlineData(100, 200)]
    [InlineData(150, 230)]
    [InlineData(200, 260)]
    public void FormulaRectangleAcceptsPointsOnOrInsideItsBounds(int x, int y)
    {
        Assert.True(WordDoubleClickRouting.ScreenPointHitsFormulaRectangle(
            x,
            y,
            left: 100,
            top: 200,
            width: 100,
            height: 60));
    }

    [Theory]
    [InlineData(99, 230)]
    [InlineData(201, 230)]
    [InlineData(150, 199)]
    [InlineData(150, 261)]
    public void FormulaRectangleRejectsNearbyBlankPoints(int x, int y)
    {
        Assert.False(WordDoubleClickRouting.ScreenPointHitsFormulaRectangle(
            x,
            y,
            left: 100,
            top: 200,
            width: 100,
            height: 60));
    }

    [Fact]
    public void FormulaRectangleRejectsInvalidGeometry()
    {
        Assert.False(WordDoubleClickRouting.ScreenPointHitsFormulaRectangle(
            100,
            200,
            left: 100,
            top: 200,
            width: 0,
            height: 60));
        Assert.False(WordDoubleClickRouting.ScreenPointHitsFormulaRectangle(
            100,
            200,
            left: 100,
            top: 200,
            width: 100,
            height: -1));
    }

    private static OfficeSelection FormulaSelection(string objectMode)
    {
        var formulaId = Guid.NewGuid().ToString();
        return new OfficeSelection
        {
            Host = "word",
            FormulaId = formulaId,
            ObjectMode = objectMode,
            Metadata = new FormulaMetadata
            {
                FormulaId = formulaId,
                Latex = "x^2",
                Lines = new List<FormulaLine>
                {
                    new() { Id = Guid.NewGuid().ToString(), Latex = "x^2" },
                },
                CodeFormat = "latex",
                DisplayMode = "inline",
                CreatedAt = DateTimeOffset.UtcNow.ToString("O"),
                UpdatedAt = DateTimeOffset.UtcNow.ToString("O"),
            },
        };
    }
}
