using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using VisualTeX.WindowsOffice.Contracts;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class VisualTeXRenderBatchTransportTests
{
    private static readonly TimeSpan TestTimeout = TimeSpan.FromSeconds(5);

    [Fact]
    public async Task RenderBatch_SubmitsOnceWaitsOnceAndReleasesWithInlineResultsInRequestOrder()
    {
        var requests = Requests("first", "second");
        var calls = new List<string>();
        var serverResult = Result("second", "first");
        serverResult.Items[0].Session.ExportResult = new OfficeExportDocument { Svg = "<svg></svg>" };

        var result = await VisualTeXSessionClient.SendRenderBatchAsync(
            requests, TestTimeout, CancellationToken.None, async (request, isWait, token) =>
            {
                calls.Add(request.Method + " " + request.RequestUri);
                if (request.Method == HttpMethod.Post)
                {
                    Assert.False(isWait);
                    using var body = JsonDocument.Parse(await request.Content!.ReadAsStringAsync());
                    var items = body.RootElement.GetProperty("items");
                    Assert.Equal(2, items.GetArrayLength());
                    Assert.Equal("first", items[0].GetProperty("id").GetString());
                    Assert.Equal("word", items[0].GetProperty("request").GetProperty("host").GetString());
                    return Created(request);
                }
                if (request.Method == HttpMethod.Get)
                {
                    Assert.True(isWait);
                    return JsonResult(request, serverResult);
                }
                Assert.False(isWait);
                Assert.False(token.IsCancellationRequested);
                serverResult.Items[0].Session.ExportResult = null;
                return Json(new { });
            });

        Assert.Equal(new[]
        {
            "POST /api/v1/app/converter/render-batches",
            $"GET /api/v1/app/converter/render-batches/{result.Id}/result",
            $"DELETE /api/v1/app/converter/render-batches/{result.Id}",
        }, calls);
        Assert.Equal(new[] { "first", "second" }, result.Items.Select(item => item.Id));
        Assert.Equal("<svg></svg>", result.Items[1].Session.ExportResult!.Svg);
    }

    [Fact]
    public async Task RenderBatch_HasNoHistorical256ItemLimit()
    {
        var ids = Enumerable.Range(0, 257).Select(index => "item-" + index).ToArray();
        var posts = 0;
        var gets = 0;
        var deletes = 0;
        var result = await VisualTeXSessionClient.SendRenderBatchAsync(
            Requests(ids), TestTimeout, CancellationToken.None, (request, _, _) =>
            {
                if (request.Method == HttpMethod.Post) { posts++; return Task.FromResult(Created(request)); }
                if (request.Method == HttpMethod.Get) { gets++; return Task.FromResult(JsonResult(request, Result(ids))); }
                deletes++;
                return Task.FromResult(Json(new { }));
            });

        Assert.Equal(257, result.Items.Count);
        Assert.Equal(1, posts);
        Assert.Equal(1, gets);
        Assert.Equal(1, deletes);
    }

    [Fact]
    public async Task RenderBatch_PreservesFailedAndCancelledItemDetailsForMaterialization()
    {
        var response = Result("failed", "cancelled");
        response.Items[0].Session.Status = "failed";
        response.Items[0].Session.Error = "invalid formula";
        response.Items[1].Session.Status = "cancelled";
        response.Items[1].Session.ExplicitCancel = true;

        var result = await RunWithResultAsync(Requests("failed", "cancelled"), response);

        Assert.Equal("failed", result.Items[0].Session.Status);
        Assert.Equal("invalid formula", result.Items[0].Session.Error);
        Assert.Equal("cancelled", result.Items[1].Session.Status);
        Assert.True(result.Items[1].Session.ExplicitCancel);
    }

    [Theory]
    [InlineData("duplicate")]
    [InlineData("unexpected")]
    [InlineData("missing")]
    [InlineData("wrong-job")]
    [InlineData("null-session")]
    public async Task RenderBatch_RejectsBrokenCorrelationAndStillReleases(string defect)
    {
        var response = defect switch
        {
            "duplicate" => Result("first", "first"),
            "unexpected" => Result("first", "other"),
            "missing" => Result("first"),
            _ => Result("first", "second"),
        };
        if (defect == "wrong-job") response.Id = "other-job";
        if (defect == "null-session") response.Items[1].Session = null!;
        var released = false;

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            RunWithResultAsync(Requests("first", "second"), response, () => released = true));

        Assert.True(released);
    }

    [Fact]
    public async Task RenderBatch_SnapshotsCallerCorrelationBeforeAwait()
    {
        var requests = Requests("original");
        var result = await VisualTeXSessionClient.SendRenderBatchAsync(
            requests, TestTimeout, CancellationToken.None, (request, _, _) =>
            {
                if (request.Method == HttpMethod.Post)
                {
                    requests[0].Id = "changed";
                    return Task.FromResult(Created(request));
                }
                return Task.FromResult(request.Method == HttpMethod.Get
                    ? JsonResult(request, Result("original")) : Json(new { }));
            });

        Assert.Equal("original", result.Items[0].Id);
    }

    [Fact]
    public async Task RenderBatch_RejectsDuplicateInputWithoutSending()
    {
        await Assert.ThrowsAsync<ArgumentException>(() =>
            VisualTeXSessionClient.SendRenderBatchAsync(
                Requests("same", "same"), TestTimeout, CancellationToken.None,
                (_, _, _) => throw new InvalidOperationException("No request should be sent.")));
    }

    [Fact]
    public async Task RenderBatch_UnknownEndpointDoesNotFallBackToDurableSessions()
    {
        var calls = new List<string>();
        await Assert.ThrowsAsync<HttpRequestException>(() =>
            VisualTeXSessionClient.SendRenderBatchAsync(
                Requests("first"), TestTimeout, CancellationToken.None, (request, _, _) =>
                {
                    calls.Add(request.Method + " " + request.RequestUri);
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound)
                    {
                        Content = new StringContent("endpoint unavailable"),
                    });
                }));

        Assert.Equal(2, calls.Count);
        Assert.Equal("POST /api/v1/app/converter/render-batches", calls[0]);
        Assert.StartsWith("DELETE /api/v1/app/converter/render-batches/", calls[1]);
    }

    [Fact]
    public async Task RenderBatch_CallerCancellationUsesIndependentCleanupToken()
    {
        using var cancellation = new CancellationTokenSource();
        var released = false;
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            VisualTeXSessionClient.SendRenderBatchAsync(
                Requests("first"), TestTimeout, cancellation.Token, (request, _, token) =>
                {
                    if (request.Method == HttpMethod.Post)
                        return Task.FromResult(Created(request));
                    if (request.Method == HttpMethod.Get)
                    {
                        cancellation.Cancel();
                        token.ThrowIfCancellationRequested();
                    }
                    Assert.False(token.IsCancellationRequested);
                    released = true;
                    return Task.FromResult(Json(new { }));
                }));

        Assert.True(released);
    }

    [Fact]
    public async Task RenderBatch_OperationTimeoutReleasesWithoutPolling()
    {
        var gets = 0;
        var released = false;
        await Assert.ThrowsAsync<TimeoutException>(() =>
            VisualTeXSessionClient.SendRenderBatchAsync(
                Requests("first"), TimeSpan.FromMilliseconds(50), CancellationToken.None,
                async (request, _, token) =>
                {
                    if (request.Method == HttpMethod.Post) return Created(request);
                    if (request.Method == HttpMethod.Get)
                    {
                        gets++;
                        await Task.Delay(Timeout.Infinite, token);
                    }
                    Assert.False(token.IsCancellationRequested);
                    released = true;
                    return Json(new { });
                }));

        Assert.Equal(1, gets);
        Assert.True(released);
    }

    [Fact]
    public async Task RenderBatch_AggregatesJobAndCleanupErrorsWithOriginalResult()
    {
        var response = Result("first");
        response.Error = "renderer stopped";
        response.Items[0].Session.Status = "failed";
        response.Items[0].Session.Error = "renderer stopped";
        var deletes = 0;

        var error = await Assert.ThrowsAsync<AggregateException>(() =>
            VisualTeXSessionClient.SendRenderBatchAsync(
                Requests("first"), TestTimeout, CancellationToken.None, (request, _, _) =>
                {
                    if (request.Method == HttpMethod.Post) return Task.FromResult(Created(request));
                    if (request.Method == HttpMethod.Get) return Task.FromResult(JsonResult(request, response));
                    deletes++;
                    throw new HttpRequestException("cleanup unavailable");
                }));

        Assert.Equal(2, error.InnerExceptions.Count);
        var primary = Assert.IsType<InvalidOperationException>(error.InnerExceptions[0]);
        Assert.Contains("renderer stopped", primary.Message);
        var details = Assert.IsType<RenderBatchResult>(primary.Data["RenderBatchResult"]);
        Assert.Equal("first", details.Items[0].Id);
        Assert.Equal("renderer stopped", details.Items[0].Session.Error);
        Assert.IsType<HttpRequestException>(error.InnerExceptions[1]);
        Assert.Equal(1, deletes);
    }

    [Fact]
    public async Task RenderBatch_SurfacesCleanupFailureWithoutRetry()
    {
        var deletes = 0;
        var error = await Assert.ThrowsAsync<HttpRequestException>(() =>
            VisualTeXSessionClient.SendRenderBatchAsync(
                Requests("first"), TestTimeout, CancellationToken.None, (request, _, _) =>
                {
                    if (request.Method == HttpMethod.Post) return Task.FromResult(Created(request));
                    if (request.Method == HttpMethod.Get) return Task.FromResult(JsonResult(request, Result("first")));
                    deletes++;
                    throw new HttpRequestException("cleanup unavailable");
                }));

        Assert.Equal("cleanup unavailable", error.Message);
        Assert.Equal(1, deletes);
    }

    [Fact]
    public async Task RenderBatch_LostSubmitResponseStillReleasesCallerOwnedId()
    {
        string? submittedId = null;
        string? releasedId = null;
        var error = await Assert.ThrowsAsync<HttpRequestException>(() =>
            VisualTeXSessionClient.SendRenderBatchAsync(
                Requests("first"), TestTimeout, CancellationToken.None, (request, _, _) =>
                {
                    if (request.Method == HttpMethod.Post)
                    {
                        using var body = JsonDocument.Parse(
                            request.Content!.ReadAsStringAsync().GetAwaiter().GetResult());
                        submittedId = body.RootElement.GetProperty("id").GetString();
                        throw new HttpRequestException("submit response lost");
                    }
                    Assert.Equal(HttpMethod.Delete, request.Method);
                    releasedId = request.RequestUri!.ToString().Split('/')[6];
                    return Task.FromResult(Json(new { }));
                }));

        Assert.Contains("submit response lost", error.Message);
        Assert.True(Guid.TryParse(submittedId, out _));
        Assert.Equal(submittedId, releasedId);
    }

    [Fact]
    public async Task RenderBatch_WrongSubmitEchoReleasesOriginalIdWithoutWaiting()
    {
        var calls = new List<string>();
        string? submittedId = null;
        await Assert.ThrowsAsync<InvalidDataException>(() =>
            VisualTeXSessionClient.SendRenderBatchAsync(
                Requests("first"), TestTimeout, CancellationToken.None, (request, _, _) =>
                {
                    calls.Add(request.Method.ToString());
                    if (request.Method == HttpMethod.Post)
                    {
                        using var body = JsonDocument.Parse(
                            request.Content!.ReadAsStringAsync().GetAwaiter().GetResult());
                        submittedId = body.RootElement.GetProperty("id").GetString();
                        return Task.FromResult(Json(new { id = "unexpected" }));
                    }
                    Assert.Equal(submittedId, request.RequestUri!.ToString().Split('/')[6]);
                    return Task.FromResult(Json(new { }));
                }));

        Assert.Equal(new[] { "POST", "DELETE" }, calls);
    }

    private static Task<RenderBatchResult> RunWithResultAsync(
        IReadOnlyList<RenderBatchItemRequest> requests,
        RenderBatchResult result,
        Action? onDelete = null) =>
        VisualTeXSessionClient.SendRenderBatchAsync(
            requests, TestTimeout, CancellationToken.None, (request, _, _) =>
            {
                if (request.Method == HttpMethod.Post) return Task.FromResult(Created(request));
                if (request.Method == HttpMethod.Get) return Task.FromResult(JsonResult(request, result));
                onDelete?.Invoke();
                return Task.FromResult(Json(new { }));
            });

    private static RenderBatchItemRequest[] Requests(params string[] ids) =>
        ids.Select(id => new RenderBatchItemRequest
        {
            Id = id,
            Request = new CreateVstoSessionRequest
            {
                Host = "word",
                Lines = new List<FormulaLine> { new() { Id = "line", Latex = "x" } },
            },
        }).ToArray();

    private static RenderBatchResult Result(params string[] ids) => new()
    {
        Id = "job",
        Items = ids.Select(id => new RenderBatchItemResult
        {
            Id = id,
            Session = new OfficeSessionDocument { Id = Guid.NewGuid().ToString("D"), Status = "committing" },
        }).ToArray(),
    };

    private static HttpResponseMessage Created(HttpRequestMessage request)
    {
        using var body = JsonDocument.Parse(request.Content!.ReadAsStringAsync().GetAwaiter().GetResult());
        var id = body.RootElement.GetProperty("id").GetString();
        Assert.True(Guid.TryParse(id, out _));
        return Json(new { id });
    }

    private static HttpResponseMessage JsonResult(HttpRequestMessage request, RenderBatchResult result)
    {
        if (result.Id == "job")
            result.Id = request.RequestUri!.ToString().Split('/')[6];
        return Json(result);
    }

    private static HttpResponseMessage Json<T>(T value) => new(HttpStatusCode.OK)
    {
        Content = new StringContent(JsonSerializer.Serialize(value, JsonOptions.Default),
            Encoding.UTF8, "application/json"),
    };
}
