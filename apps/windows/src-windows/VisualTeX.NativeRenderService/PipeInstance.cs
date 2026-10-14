using System.IO.Pipes;

namespace VisualTeX.NativeRenderService;

internal static class PipeInstance
{
    // Return the single pipe instance to the listening state after every request.
    // A client that gives up mid-request (its own timeout, a closed Office) leaves
    // the instance Broken: IsConnected is then false, yet WaitForConnection keeps
    // throwing "pipe is broken" until DisconnectNamedPipe resets it. Gating the
    // reset on IsConnected made one abandoned request stop the service from
    // answering anyone until it was restarted.
    internal static void Reset(NamedPipeServerStream endpoint)
    {
        try { endpoint.Disconnect(); }
        catch (ObjectDisposedException) { }
        catch (InvalidOperationException) { } // never connected, or already reset
        catch (IOException) { }
    }
}
