using VisualTeX.WindowsOleBridge;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class ProtocolSecurityTests
{
    [Fact]
    public void PipeNameIsBoundToVisualTeXNamespace()
    {
        Assert.Equal(
            "VisualTeX.OfficeBridge.S-1-5-21-1",
            NamedPipeServer.NormalizePipeName(
                @"\\.\pipe\VisualTeX.OfficeBridge.S-1-5-21-1"));
        Assert.Throws<ArgumentException>(() =>
            NamedPipeServer.NormalizePipeName(@"Other\Nested"));
    }

    [Theory]
    [InlineData("same-token", "same-token", true)]
    [InlineData("same-token", "other-token", false)]
    [InlineData("short", "shorter", false)]
    public void PipeTokenComparisonRejectsMismatch(
        string left,
        string right,
        bool expected)
    {
        Assert.Equal(expected, NamedPipeServer.ConstantTimeEquals(left, right));
    }
}
