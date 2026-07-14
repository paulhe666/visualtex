using VisualTeX.WindowsOleBridge;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class ReplacementTransactionTests
{
    [Fact]
    public void FailedConfigurationKeepsOriginalAndDeletesCandidate()
    {
        var originalDeleted = false;
        var candidateDeleted = false;
        var candidate = new object();

        Assert.Throws<InvalidOperationException>(() =>
            ReplacementTransaction.Execute(
                () => candidate,
                _ => throw new InvalidOperationException("metadata write failed"),
                () => originalDeleted = true,
                _ => candidateDeleted = true));

        Assert.False(originalDeleted);
        Assert.True(candidateDeleted);
    }

    [Fact]
    public void SuccessfulConfigurationDeletesOriginalAfterCandidateIsReady()
    {
        var order = new List<string>();
        var candidate = ReplacementTransaction.Execute(
            () =>
            {
                order.Add("create");
                return new object();
            },
            _ => order.Add("configure"),
            () => order.Add("delete-original"),
            _ => order.Add("delete-candidate"));

        Assert.NotNull(candidate);
        Assert.Equal(new[] { "create", "configure", "delete-original" }, order);
    }
}
