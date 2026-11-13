using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordCoordinateShiftsTests
{
    [Fact]
    public void SelectionKeepsCoveringTheSameTextAfterSourcesShrink()
    {
        // Selection [100, 400). Three LaTeX sources inside it become one-character
        // objects; one after it does too and must not move the selection.
        var shifts = new WordCoordinateShifts();
        shifts.Add(130, 1 - 30);   // [100, 130) starts exactly at the selection start
        shifts.Add(260, 1 - 60);   // [200, 260)
        shifts.Add(400, 1 - 20);   // [380, 400) ends exactly at the selection end
        shifts.Add(520, 1 - 20);   // [500, 520) after the selection

        Assert.Equal(100, shifts.Map(100));
        Assert.Equal(400 - 29 - 59 - 19, shifts.Map(400));
    }

    [Fact]
    public void SelectionGrowsWhenFormulasBecomeLongerSource()
    {
        var shifts = new WordCoordinateShifts();
        shifts.Add(51, 40 - 1);    // a one-character object at [50, 51) becomes 40 characters
        Assert.Equal(10, shifts.Map(10));
        Assert.Equal(90 + 39, shifts.Map(90));
    }
}
