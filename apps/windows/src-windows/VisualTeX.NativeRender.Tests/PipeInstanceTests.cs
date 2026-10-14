using System.IO.Pipes;
using VisualTeX.NativeRenderService;
using Xunit;

namespace VisualTeX.NativeRender.Tests;

public sealed class PipeInstanceTests
{
    // An Office client that gives up mid-request (its own timeout, Word closing)
    // must not leave the single service instance deaf to every later request.
    [Fact]
    public async Task AbandonedRequestLeavesTheInstanceReadyForTheNextClient()
    {
        var name = "VisualTeX.PipeInstanceTest." + Guid.NewGuid().ToString("N");
        using var server = new NamedPipeServerStream(name, PipeDirection.InOut, 1,
            PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
        using (var abandoned = new NamedPipeClientStream(".", name, PipeDirection.InOut, PipeOptions.Asynchronous))
        {
            var accept = server.WaitForConnectionAsync();
            await abandoned.ConnectAsync(5000);
            await accept;
        }

        // The service then answers a client that is gone; the instance breaks.
        await Assert.ThrowsAnyAsync<IOException>(async () =>
        {
            for (var attempt = 0; attempt < 64; attempt++)
                await server.WriteAsync(new byte[64 * 1024]);
        });
        Assert.False(server.IsConnected);
        // Without a reset a broken instance cannot listen again.
        await Assert.ThrowsAnyAsync<IOException>(() => server.WaitForConnectionAsync());

        PipeInstance.Reset(server);

        using var next = new NamedPipeClientStream(".", name, PipeDirection.InOut, PipeOptions.Asynchronous);
        var acceptNext = server.WaitForConnectionAsync();
        await next.ConnectAsync(5000);
        await acceptNext.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.True(server.IsConnected);
    }

    [Fact]
    public void ResetIsHarmlessOnAnInstanceThatNeverConnected()
    {
        var name = "VisualTeX.PipeInstanceTest." + Guid.NewGuid().ToString("N");
        using var server = new NamedPipeServerStream(name, PipeDirection.InOut, 1,
            PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
        PipeInstance.Reset(server);
        PipeInstance.Reset(server);
    }
}
