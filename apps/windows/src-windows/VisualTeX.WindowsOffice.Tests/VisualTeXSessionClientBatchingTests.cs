using System;
using System.Linq;
using VisualTeX.WindowsOffice.Contracts;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class VisualTeXSessionClientBatchingTests
{
    [Fact]
    public void ChunkConverterSessionIds_PreservesAll257SessionsInOrder()
    {
        var sessions = Enumerable.Range(1, 257)
            .Select(index => Guid.NewGuid().ToString("D"))
            .ToArray();

        var chunks = VisualTeXSessionClient.ChunkConverterSessionIds(sessions);

        Assert.Equal(new[] { 256, 1 }, chunks.Select(chunk => chunk.Count).ToArray());
        Assert.Equal(sessions, chunks.SelectMany(chunk => chunk).ToArray());
    }

    [Fact]
    public void ChunkConverterSessionIds_PreservesAll1000SessionsInOrder()
    {
        var sessions = Enumerable.Range(1, 1000)
            .Select(index => Guid.NewGuid().ToString("D"))
            .ToArray();

        var chunks = VisualTeXSessionClient.ChunkConverterSessionIds(sessions);

        Assert.Equal(new[] { 256, 256, 256, 232 }, chunks.Select(chunk => chunk.Count).ToArray());
        Assert.Equal(sessions, chunks.SelectMany(chunk => chunk).ToArray());
    }

    // A Stop pressed before the companion check must surface as cancellation;
    // it used to be treated as a missing companion, start a VisualTeX process
    // and report "companion validation failed".
    [Fact]
    public async System.Threading.Tasks.Task EnsureHealthyAsync_CancelledCallerIsCancellationNotACompanionFailure()
    {
        using var client = new VisualTeXSessionClient();
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => client.EnsureHealthyAsync(cancellation.Token));
    }
}
