using System.Net;
using Microsoft.Win32;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class VisualTeXCompanionEndpointTests
{
    [Theory]
    [InlineData(0, 43127, 43127)]
    [InlineData(43127, 43127, 43127)]
    [InlineData(38127, 43127, 38127)]
    [InlineData(51234, 43127, 51234)]
    [InlineData(-1, 43127, 43127)]
    [InlineData(65536, 43127, 43127)]
    public void PortVisualTeXBoundWinsOverTheStandardPort(int activePort, int registeredPort, int expected)
    {
        Assert.Equal(expected, CompanionEndpoint.ResolvePort(activePort, registeredPort));
    }

    [Fact]
    public void InstallJsonKeepsTheStandardPortWhileTheCompanionListensOnAnother()
    {
        var configuration = new CompanionConfiguration
        {
            AppDataRoot = Path.GetFullPath("normal-runtime-data"),
            RegisteredPort = 43127,
            CompanionPort = 38127,
        };

        Assert.Equal(38127, configuration.CompanionPort);
        Assert.Equal(43127, configuration.InstallJsonPort);
    }

    [Fact]
    public async Task RequestsFollowThePortTheCompanionMovedTo()
    {
        var recorder = new RecordingHandler();
        var port = 43127;
        using var client = new HttpClient(new CompanionPortHandler(recorder, () => port))
        {
            BaseAddress = new Uri("https://127.0.0.1:43127"),
        };

        using (await client.GetAsync("/health")) { }
        port = 38127;
        using (await client.GetAsync("/api/v1/sessions/abc?wait=25")) { }
        using (await client.PostAsync("/api/v1/sessions", new StringContent("{}"))) { }
        // an unusable port leaves the request where it was going
        port = 0;
        using (await client.GetAsync("/health")) { }

        Assert.Equal(
            new[]
            {
                "https://127.0.0.1:43127/health",
                "https://127.0.0.1:38127/api/v1/sessions/abc?wait=25",
                "https://127.0.0.1:38127/api/v1/sessions",
                "https://127.0.0.1:43127/health",
            },
            recorder.Requests);
    }

    [Fact]
    public void RunningVisualTeXWithoutListeningCompanionIsNamedWithItsReason()
    {
        var diagnostic = NotListening();

        VisualTeXSessionClient.DescribeStartFailure(
            diagnostic,
            visualTeXRunning: true,
            startError: " VisualTeX Office companion cannot listen on 127.0.0.1: port 43127 is in use ",
            occupiedByOther: null);

        Assert.Equal("CompanionNotListening", diagnostic.FailureType);
        Assert.Equal(
            "VisualTeX Office companion cannot listen on 127.0.0.1: port 43127 is in use",
            diagnostic.CompanionStartError);
        Assert.StartsWith(
            "VisualTeX is running, but its Office companion does not listen on 127.0.0.1:43127.",
            diagnostic.Message,
            StringComparison.Ordinal);
        Assert.Contains("port 43127 is in use", diagnostic.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("exit code", diagnostic.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RunningVisualTeXThatGaveNoReasonIsStillNotATimeout()
    {
        var diagnostic = NotListening();

        VisualTeXSessionClient.DescribeStartFailure(diagnostic, visualTeXRunning: true, startError: "", occupiedByOther: null);

        Assert.Equal("CompanionNotListening", diagnostic.FailureType);
        Assert.Equal(string.Empty, diagnostic.CompanionStartError);
        Assert.Equal(
            "VisualTeX is running, but its Office companion does not listen on 127.0.0.1:43127.",
            diagnostic.Message);
    }

    [Fact]
    public void VisualTeXThatNeverCameUpStaysATimeoutWithItsStage()
    {
        var diagnostic = NotListening();

        // a reason left by an earlier run says nothing about a VisualTeX that is not running
        VisualTeXSessionClient.DescribeStartFailure(diagnostic, visualTeXRunning: false, startError: "stale", occupiedByOther: null);

        Assert.Equal("PortNotListening", diagnostic.FailureType);
        Assert.Equal(string.Empty, diagnostic.CompanionStartError);
        Assert.Contains("did not pass stage 'port-listen' in time", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("No process is listening on 127.0.0.1:43127.", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PortHeldByAnotherProgramIsNamedWhenVisualTeXTookNoOtherPort()
    {
        var diagnostic = NotListening();
        diagnostic.PortListening = true;
        diagnostic.Stage = "https-handshake";
        diagnostic.FailureType = "TlsValidationFailed";
        const string occupied = "Port 43127 is occupied by other PID=4 at 'C:\\other.exe', not by 'C:\\VisualTeX\\visualtex.exe'.";

        VisualTeXSessionClient.DescribeStartFailure(diagnostic, visualTeXRunning: true, startError: "", occupiedByOther: occupied);

        Assert.Equal("PortOccupiedByOtherProcess", diagnostic.FailureType);
        Assert.Equal(occupied, diagnostic.Message);
    }

    [Theory]
    [InlineData("CompanionPort")]
    [InlineData("ProtocolVersion")]
    [InlineData("CompanionActivePort")]
    [InlineData("VisualTeXTestValueThatDoesNotExist")]
    public void RegistryNumbersAreReadAsTheRegistryHoldsThem(string name)
    {
        using var key = Registry.CurrentUser.OpenSubKey(@"Software\VisualTeX\OfficeIntegration");
        var expected = key?.GetValue(name) is int value ? value : 0;

        Assert.Equal(expected, CompanionEndpoint.ReadDword(name));
    }

    [Theory]
    [InlineData("ExecutablePath")]
    [InlineData("CertificateThumbprint")]
    [InlineData("CompanionStartError")]
    [InlineData("VisualTeXTestValueThatDoesNotExist")]
    public void RegistryTextsAreReadAsTheRegistryHoldsThem(string name)
    {
        using var key = Registry.CurrentUser.OpenSubKey(@"Software\VisualTeX\OfficeIntegration");
        var expected = (key?.GetValue(name) as string ?? string.Empty).Trim();

        Assert.Equal(expected, CompanionEndpoint.ReadString(name));
    }

    private static CompanionHealthDiagnostic NotListening() => new()
    {
        Stage = "port-listen",
        FailureType = "PortNotListening",
        Message = "No process is listening on 127.0.0.1:43127.",
        CompanionPort = 43127,
        PortListening = false,
    };

    private sealed class RecordingHandler : HttpMessageHandler
    {
        public List<string> Requests { get; } = new();

        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken)
        {
            Requests.Add(request.RequestUri!.AbsoluteUri);
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { RequestMessage = request });
        }
    }
}
