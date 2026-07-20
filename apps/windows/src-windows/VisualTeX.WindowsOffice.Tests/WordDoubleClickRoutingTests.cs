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
    public void WordOmmlKeepsWordNativeEditor()
    {
        var selection = FormulaSelection(FormulaOleContract.WordOmmlMode);

        Assert.False(WordDoubleClickRouting.ShouldOpenVisualTeX(selection));
    }

    [Fact]
    public void OrdinaryWordSelectionDoesNothing()
    {
        Assert.False(WordDoubleClickRouting.ShouldOpenVisualTeX(null));
        Assert.False(WordDoubleClickRouting.ShouldOpenVisualTeX(new OfficeSelection()));
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
