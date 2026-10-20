using System.Text;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using Xunit;

namespace VisualTeX.NativeRender.Tests;

public sealed class NativeRenderProtocolTests
{
    private static NativeRenderRequest Request() => new()
    { Items = new() { new NativeRenderInput { Id = "formula", Mtef = new byte[] { 5, 1, 0 } } } };

    [Fact]
    public async Task DataOnlyRequestRoundTripsWithoutPaths()
    {
        using var stream = new MemoryStream();
        await NativeRenderProtocol.WriteAsync(stream, Request(), NativeRenderProtocol.MaxRequestBytes, default);
        stream.Position = 0;
        var request = await NativeRenderProtocol.ReadAsync<NativeRenderRequest>(stream, NativeRenderProtocol.MaxRequestBytes, default);
        NativeRenderProtocol.Validate(request);
        Assert.Equal(new byte[] { 5, 1, 0 }, request.Items.Single().Mtef);
        var json = JsonSerializer.Serialize(request);
        Assert.DoesNotContain("Path", json);
        Assert.DoesNotContain("Command", json);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(0)]
    [InlineData(int.MaxValue)]
    public async Task InvalidFrameSizeIsRejectedBeforeReadingBody(int size)
    {
        using var stream = new MemoryStream(BitConverter.GetBytes(size));
        await Assert.ThrowsAsync<InvalidDataException>(() => NativeRenderProtocol.ReadAsync<NativeRenderRequest>(stream, 1024, default));
    }

    [Fact]
    public async Task TruncatedFrameIsRejected()
    {
        using var stream = new MemoryStream(new byte[] { 8, 0, 0, 0, 1 });
        await Assert.ThrowsAsync<EndOfStreamException>(() => NativeRenderProtocol.ReadAsync<NativeRenderRequest>(stream, 1024, default));
    }

    [Fact]
    public void DuplicateAndOversizedInputsAreRejected()
    {
        var request = Request();
        request.Items.Add(request.Items[0]);
        Assert.Throws<InvalidDataException>(() => NativeRenderProtocol.Validate(request));
        request = Request();
        request.Items[0].Mtef = new byte[NativeRenderProtocol.MaxMtefBytes + 1];
        Assert.Throws<InvalidDataException>(() => NativeRenderProtocol.Validate(request));
        request = Request();
        request.Version++;
        Assert.Throws<InvalidDataException>(() => NativeRenderProtocol.Validate(request));
    }

    [Fact]
    public void OutputMustMatchInputsAndHaveFiniteDimensions()
    {
        var request = Request();
        var response = new NativeRenderResponse
        { Items = new() { new NativeRenderOutput { Id = "formula", Wmf = new byte[24], WidthPt = 12, HeightPt = 9 } } };
        NativeRenderProtocol.Validate(response, request);
        response.Items[0].WidthPt = float.PositiveInfinity;
        Assert.Throws<InvalidDataException>(() => NativeRenderProtocol.Validate(response, request));
        response.Items[0].WidthPt = 12;
        response.Items[0].Id = "not-requested";
        Assert.Throws<InvalidDataException>(() => NativeRenderProtocol.Validate(response, request));
    }

    [Fact]
    public void FailuresNameOnlyRequestedItemsWithoutOutput()
    {
        var request = Request();
        request.Items.Add(new NativeRenderInput { Id = "second", Mtef = new byte[] { 5, 1, 0 } });
        var response = new NativeRenderResponse
        {
            Items = new() { new NativeRenderOutput { Id = "formula", Wmf = new byte[24], WidthPt = 12, HeightPt = 9 } },
            Failures = new() { new NativeRenderFailure { Id = "second", Reason = "MathType native renderer returned status -4." } },
        };
        NativeRenderProtocol.Validate(response, request);
        response.Failures[0].Id = "formula";
        Assert.Throws<InvalidDataException>(() => NativeRenderProtocol.Validate(response, request));
        response.Failures[0].Id = "not-requested";
        Assert.Throws<InvalidDataException>(() => NativeRenderProtocol.Validate(response, request));
        response.Failures[0].Id = "second";
        response.Failures.Add(new NativeRenderFailure { Id = "second", Reason = "again" });
        Assert.Throws<InvalidDataException>(() => NativeRenderProtocol.Validate(response, request));
        response.Failures.RemoveAt(1);
        response.Failures[0].Reason = new string('x', NativeRenderProtocol.MaxFailureReasonLength + 1);
        Assert.Throws<InvalidDataException>(() => NativeRenderProtocol.Validate(response, request));
    }

    [Fact]
    public void ResponseFromAnOlderServiceHasNoFailures()
    {
        var response = JsonSerializer.Deserialize<NativeRenderResponse>("{\"Version\":1,\"Error\":\"\",\"Items\":[]}")!;
        NativeRenderProtocol.Validate(response, Request());
        Assert.Empty(response.Failures);
        Assert.Equal("No output.", NativeRenderProtocol.BoundFailureReason("  "));
        Assert.Equal(NativeRenderProtocol.MaxFailureReasonLength,
            NativeRenderProtocol.BoundFailureReason(new string('x', 500)).Length);
    }
}
