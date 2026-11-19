using System;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace VisualTeX.WindowsOffice.Contracts;

/// <summary>
/// Where VisualTeX's companion listens. The standard port can be taken or lie in a
/// range Windows reserved (WSL mirrored networking, Hyper-V); VisualTeX then binds
/// another one and names the port it listens on in the registry at every start.
/// </summary>
internal static class CompanionEndpoint
{
    internal const int StandardPort = 43127;
    private const string KeyPath = @"Software\VisualTeX\OfficeIntegration";
    private const string ActivePortValue = "CompanionActivePort";
    private const string StartErrorValue = "CompanionStartError";
    private static readonly IntPtr CurrentUser = new(unchecked((int)0x80000001));
    private const uint RrfRtRegSz = 0x00000002;
    private const uint RrfRtRegDword = 0x00000010;

    /// <summary>The port VisualTeX bound when it names a valid one, else the registered standard port.</summary>
    internal static int ResolvePort(int activePort, int registeredPort) =>
        activePort is > 0 and <= 65535 ? activePort : registeredPort;

    /// <summary>The port VisualTeX last bound; 0 when it never named one.</summary>
    internal static int ReadActivePort() => ReadDword(ActivePortValue);

    /// <summary>Why VisualTeX's companion does not listen although VisualTeX runs; empty when it gave no reason.</summary>
    internal static string ReadStartError() => ReadString(StartErrorValue);

    /// <summary>A number of VisualTeX's Office integration key; 0 when absent.</summary>
    internal static int ReadDword(string name)
    {
        try
        {
            var size = sizeof(int);
            return RegGetValueW(CurrentUser, KeyPath, name, RrfRtRegDword, IntPtr.Zero, out var value, ref size) == 0
                ? value
                : 0;
        }
        catch (Exception ignored)
        {
            IgnoredErrors.Ignore(ignored);
            return 0;
        }
    }

    /// <summary>A text of VisualTeX's Office integration key; empty when absent.</summary>
    internal static string ReadString(string name)
    {
        try
        {
            var size = 0;
            if (RegGetValueW(CurrentUser, KeyPath, name, RrfRtRegSz, IntPtr.Zero, null, ref size) != 0
                || size <= 0)
                return string.Empty;
            var buffer = new byte[size];
            return RegGetValueW(CurrentUser, KeyPath, name, RrfRtRegSz, IntPtr.Zero, buffer, ref size) == 0
                ? Encoding.Unicode.GetString(buffer, 0, Math.Min(size, buffer.Length)).TrimEnd('\0').Trim()
                : string.Empty;
        }
        catch (Exception ignored)
        {
            IgnoredErrors.Ignore(ignored);
            return string.Empty;
        }
    }

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    private static extern int RegGetValueW(
        IntPtr key, string subKey, string value, uint flags, IntPtr type, out int data, ref int size);

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    private static extern int RegGetValueW(
        IntPtr key, string subKey, string value, uint flags, IntPtr type, byte[]? data, ref int size);
}

/// <summary>
/// Sends every request to the companion's current port. The clients keep one base
/// address for their lifetime, while VisualTeX can come back on another port.
/// </summary>
internal sealed class CompanionPortHandler : DelegatingHandler
{
    private readonly Func<int> _port;

    internal CompanionPortHandler(HttpMessageHandler inner, Func<int> port) : base(inner)
    {
        _port = port;
    }

    protected override Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request,
        CancellationToken cancellationToken)
    {
        var port = _port();
        var address = request.RequestUri;
        if (address is { IsAbsoluteUri: true } && port is > 0 and <= 65535 && address.Port != port)
            request.RequestUri = new UriBuilder(address) { Port = port }.Uri;
        return base.SendAsync(request, cancellationToken);
    }
}
