namespace VisualTeX.WordVsto;

// Data-only codec tests have no Word process or mouse hook. Keep the codec's
// diagnostic messages observable without loading the desktop hook in testhost.
internal static class WordDoubleClickHook
{
    internal static void TraceMessage(string message) => System.Diagnostics.Trace.WriteLine(message);
}
