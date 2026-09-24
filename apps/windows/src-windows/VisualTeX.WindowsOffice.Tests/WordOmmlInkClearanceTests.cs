using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlInkClearanceTests
{
    private static string Math(string body) =>
        "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">" + body + "</m:oMath>";

    [Fact]
    public void OrdinarySingleLineDoesNotAcquireTallFormulaSpacing()
    {
        Assert.False(WordEquationNumbering.HasTallNativeOmmlStructure(Math("<m:r><m:t>x+y</m:t></m:r>")));
        Assert.False(WordEquationNumbering.HasTallNativeOmmlStructure(Math("<m:m><m:mr><m:e/></m:mr></m:m>")));
    }

    [Theory]
    [InlineData("<m:m><m:mr><m:e/></m:mr><m:mr><m:e/></m:mr></m:m>")]
    [InlineData("<m:eqArr><m:e/><m:e/></m:eqArr>")]
    [InlineData("<m:f><m:num/><m:den/></m:f>")]
    [InlineData("<m:nary><m:sub/><m:sup/><m:e/></m:nary>")]
    public void TallStructuresKeepDescentOnEditAndNumberRefresh(string formula)
        => Assert.True(WordEquationNumbering.HasTallNativeOmmlStructure(Math(formula)));
}
