using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeFieldLocatorTests
{
    [Theory]
    [InlineData(" MACROBUTTON MTPlaceRef (1)", "MTPlaceRef", true)]
    [InlineData("\tmacrobutton\tmtplaceref\tcustom-label", "MTPlaceRef", true)]
    [InlineData(" MACROBUTTON   MTEditEquationSection2 chapter", "MTEditEquationSection2", true)]
    [InlineData(" MACROBUTTON MTPlaceRef\u0013 SEQ MTEqn ", "MTPlaceRef", true)]
    [InlineData(" MACROBUTTON MTPlaceRefOther (1)", "MTPlaceRef", false)]
    [InlineData(" MACROBUTTON Other MTPlaceRef ", "MTPlaceRef", false)]
    [InlineData(" QUOTE \"MACROBUTTON MTPlaceRef\" ", "MTPlaceRef", false)]
    [InlineData(" EMBED Equation.DSMT4 ", "MTPlaceRef", false)]
    [InlineData(" MACROBUTTON\u0013 MTPlaceRef ", "MTPlaceRef", false)]
    [InlineData("", "MTPlaceRef", false)]
    [InlineData(null, "MTPlaceRef", false)]
    public void ValidatesMacroHeaderNotAnArgument(string? code, string macroName, bool expected)
        => Assert.Equal(expected, MathTypeFieldLocator.MatchesMacroCode(code, macroName));
}
