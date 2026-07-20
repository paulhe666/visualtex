using VisualTeX.WindowsOleBridge;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class StaDispatcherTests
{
    [Fact]
    public async Task AllOfficeWorkRunsOnOneStaThread()
    {
        var root = Path.Combine(Path.GetTempPath(), "VisualTeX.Tests", Guid.NewGuid().ToString("N"));
        using var dispatcher = new OfficeStaDispatcher(new FileLogger(root));
        var tasks = Enumerable.Range(0, 20)
            .Select(_ => dispatcher.InvokeAsync(
                () => new
                {
                    ThreadId = Environment.CurrentManagedThreadId,
                    Apartment = Thread.CurrentThread.GetApartmentState(),
                },
                CancellationToken.None))
            .ToArray();
        var results = await Task.WhenAll(tasks);
        Assert.All(results, result => Assert.Equal(ApartmentState.STA, result.Apartment));
        Assert.Single(results.Select(result => result.ThreadId).Distinct());
    }
}
