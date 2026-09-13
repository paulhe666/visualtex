using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOwnedOmmlSeparatorTests
{
    [Theory]
    [InlineData(20, 20, 21, "\r", false, true)]
    [InlineData(20, 19, 20, "\r", false, false)]
    [InlineData(20, 21, 22, "\r", false, false)]
    [InlineData(20, 20, 20, "", false, false)]
    [InlineData(20, 20, 22, "\r\r", false, false)]
    [InlineData(20, 20, 21, "\r\a", true, false)]
    [InlineData(20, 20, 21, "\r", true, false)]
    [InlineData(20, 20, 21, "x", false, false)]
    [InlineData(20, 20, 21, null, false, false)]
    public void OnlyTheExactMarkedBodyParagraphCanBeReused(
        int tableEnd, int start, int end, string? text, bool withinTable, bool expected)
        => Assert.Equal(expected, WordEquationNumbering.IsExactOwnedNativeOmmlSeparator(
            tableEnd, start, end, text, withinTable));

    [Fact]
    public void SeparatorIdentityFitsWordAndParticipatesInOwnedRollback()
    {
        const string id = "11111111-2222-3333-4444-555555555555";
        var name = WordEquationNumbering.NativeOmmlOwnedSeparatorBookmarkName(id);
        Assert.Equal("VTEqSep_11111111222233334444555555555555", name);
        Assert.Equal(40, name.Length);
        Assert.Contains(name, WordBookmarkRecoverySnapshot.NamesForFormula(id));
    }
}
