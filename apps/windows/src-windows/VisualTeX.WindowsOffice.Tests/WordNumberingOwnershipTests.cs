using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordNumberingOwnershipTests
{
    [Fact]
    public void MathTypeOnlyDocumentDoesNotEnterVisualTeXReconciliation()
    {
        const string xml = "<document><body><p><OLEObject ProgID=\"Equation.DSMT4\"/>"
            + "<instrText>MACROBUTTON MTPlaceRef SEQ MTEqn</instrText></p></body></document>";
        Assert.True(WordEquationNumbering.IsUnmanagedEquationDocumentXml(xml));
    }

    [Theory]
    [InlineData("<OLEObject ProgID=\"VisualTeX.Formula.1\"/>")]
    [InlineData("<oMath><r>x</r></oMath>")]
    [InlineData("<bookmarkStart name=\"VTO_test\"/>")]
    [InlineData("<bookmarkStart name=\"VTOMML_test\"/>")]
    [InlineData("<instrText>REF VTEqNum_test</instrText>")]
    [InlineData("<instrText>SEQ VisualTeXEquation</instrText>")]
    [InlineData("<picture alt=\"visualtex:v1:deflate:cache\"/>")]
    public void MissingNumberAliasesNeverHideManagedContent(string content)
        => Assert.False(WordEquationNumbering.IsUnmanagedEquationDocumentXml(
            "<document><body>" + content + "</body></document>"));

    [Theory]
    [InlineData("")]
    [InlineData("<broken>")]
    [InlineData("<document/>")]
    public void UnknownInventoryCannotSkipMaintenance(string xml)
        => Assert.False(WordEquationNumbering.IsUnmanagedEquationDocumentXml(xml));
}
