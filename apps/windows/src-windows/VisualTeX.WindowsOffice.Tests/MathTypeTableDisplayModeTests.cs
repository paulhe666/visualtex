using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeTableDisplayModeTests
{
    [Theory]
    [InlineData("\t", true, true, true)]
    [InlineData("", true, true, false)]
    [InlineData(" ", true, true, false)]
    [InlineData("\t", false, true, false)]
    [InlineData("\t", true, false, false)]
    [InlineData("\t", false, false, false)]
    public void TableDisplayRequiresExactCenterTabAndBothTabStops(
        string prefix,
        bool hasCenterTab,
        bool hasRightTab,
        bool expected)
    {
        Assert.Equal(
            expected,
            WordEquationNumbering.IsTableMathTypeDisplayLayout(
                prefix,
                hasCenterTab,
                hasRightTab));
    }
}
