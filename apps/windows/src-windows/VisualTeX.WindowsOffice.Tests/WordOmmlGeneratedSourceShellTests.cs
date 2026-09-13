using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlGeneratedSourceShellTests
{
    [Theory]
    [InlineData(227, 227, 227, 227, null, true)]
    [InlineData(227, 227, 227, 227, "", true)]
    [InlineData(227, 227, 228, 236, "p=q+r", false)]
    [InlineData(227, 227, 208, 216, "a=b+c", false)]
    [InlineData(227, 228, 227, 227, null, false)]
    [InlineData(227, 227, 227, 228, "", false)]
    [InlineData(227, 227, 228, 228, null, false)]
    [InlineData(227, 227, 227, 227, " ", false)]
    [InlineData(227, 227, 227, 227, "\r", false)]
    public void OnlyAnEmptyShellAtTheExactGeneratedSourcePointIsDiscardable(
        int contentStart, int contentEnd, int mathStart, int mathEnd, string? text, bool expected)
    {
        Assert.Equal(expected, WordEquationNumbering.IsEmptyGeneratedNativeMathShell(
            contentStart, contentEnd, mathStart, mathEnd, text));
    }
}
