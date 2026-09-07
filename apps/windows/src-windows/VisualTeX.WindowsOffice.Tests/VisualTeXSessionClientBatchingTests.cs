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
}
