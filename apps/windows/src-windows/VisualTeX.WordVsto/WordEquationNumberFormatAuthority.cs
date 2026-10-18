using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WordVsto;

/// <summary>
/// Resolves one document-wide equation-number format for every formula family.
///
/// The format is never stored in the document. The one chosen in this Word
/// session wins; otherwise the first proven equation-number series in document
/// order (Word-native or MathType) shows it, and a document without numbers
/// uses the user's default. This keeps OMML, MathType and VisualTeX counters
/// independent while preventing each host family from inventing its own
/// display format.
/// </summary>
internal static class WordEquationNumberFormatAuthority
{
    // One add-in operation (a redraw, a conversion, a bulk import, one formula
    // write) never changes the document's number format, but it used to ask
    // for it once per numbered formula. Inference scans every equation number
    // and field in the document, and every numbered write adds fields, so a
    // batch grew quadratically. Inside an operation scope the first answer is
    // reused for the same document. Scopes nest and live on the Word thread.
    [ThreadStatic]
    private static OperationScope? _scope;

    internal static IDisposable BeginOperationScope() => new OperationScope();

    private sealed class OperationScope : IDisposable
    {
        private readonly OperationScope? _outer;
        private readonly Dictionary<string, string> _resolved;
        private bool _disposed;

        internal OperationScope()
        {
            _outer = _scope;
            _resolved = _outer?._resolved
                ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _scope = this;
        }

        internal bool TryGet(string key, out string formatId)
        {
            if (_resolved.TryGetValue(key, out var value))
            {
                formatId = value;
                return true;
            }
            formatId = string.Empty;
            return false;
        }

        internal void Store(string key, string formatId) => _resolved[key] = formatId;

        internal void Clear() => _resolved.Clear();

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            if (ReferenceEquals(_scope, this)) _scope = _outer;
        }
    }

    // Callers that change the format (or the numbers it is inferred from) inside
    // an operation drop the cached answer so the next lookup sees the new state.
    internal static void InvalidateOperationScope() => _scope?.Clear();

    internal static string Resolve(Document document)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        var scope = _scope;
        if (scope is null) return ResolveUncached(document);
        string? key = null;
        try { key = document.FullName; }
        catch { }
        if (string.IsNullOrEmpty(key)) return ResolveUncached(document);
        if (scope.TryGet(key!, out var cached)) return cached;
        var resolved = ResolveUncached(document);
        scope.Store(key!, resolved);
        return resolved;
    }

    private static string ResolveUncached(Document document)
    {
        var stored =
            WordEquationNumbering.CaptureEquationNumberFormatVariable(
                document);
        if (stored.Exists)
            return EquationNumberFormat.Resolve(stored.Value).Id;

        bool inferredFromNumbers;
        string inferred;
        using (WordOperationMetrics.Measure("Numbering.Authority.InferFromNumbers"))
            inferredFromNumbers = TryInferFromNumbers(document, out inferred);
        if (inferredFromNumbers)
            return inferred;

        return EquationNumberFormat.Resolve(
            WordEquationNumbering
                .GetDefaultEquationNumberFormatId()).Id;
    }

    internal static bool TryInferFromNumbers(Document document, out string formatId)
    {
        formatId = string.Empty;
        var candidates =
            new List<(int Start, string FormatId)>(2);
        bool nativeFound;
        int nativeStart;
        string nativeFormat;
        using (WordOperationMetrics.Measure("Numbering.Authority.InferNative"))
            nativeFound = WordNativeNumberingAdoptionPlan.TryInferDocumentFormat(
                document,
                out nativeStart,
                out nativeFormat);
        if (nativeFound)
        {
            candidates.Add(
                (nativeStart, nativeFormat));
        }

        bool mathTypeFound;
        int mathTypeStart;
        string mathTypeFormat;
        // The earliest number decides, so a MathType number can only win
        // before the first native one: search no further than that.
        using (WordOperationMetrics.Measure("Numbering.Authority.InferMathType"))
            mathTypeFound = MathTypeEquationNumbering.TryInferDocumentFormat(
                document,
                nativeFound ? nativeStart : int.MaxValue,
                out mathTypeStart,
                out mathTypeFormat);
        if (mathTypeFound)
        {
            candidates.Add(
                (mathTypeStart, mathTypeFormat));
        }

        if (candidates.Count == 0)
            return false;
        var first =
            candidates
                .OrderBy(candidate => candidate.Start)
                .First();
        formatId = EquationNumberFormat.Resolve(
            first.FormatId).Id;
        return true;
    }
}
