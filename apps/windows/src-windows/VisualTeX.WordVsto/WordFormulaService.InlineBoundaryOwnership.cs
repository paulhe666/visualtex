using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    // Original prose is never a sacrificial OMML guard. Both insertion paths use
    // this same owner proof so table and body text follow identical cleanup rules.
    private sealed class InlineSourceBoundaryGuard : IDisposable
    {
        private readonly string expected;
        private readonly int hidden;
        internal Range Range { get; }

        internal InlineSourceBoundaryGuard(Range range, string expected, int hidden)
        {
            Range = range;
            this.expected = expected;
            this.hidden = hidden;
        }

        internal void VerifyAndRestoreVisibility()
        {
            if (!string.Equals(Range.Text ?? string.Empty, expected, StringComparison.Ordinal))
                throw new InvalidDataException("Native equation insertion changed the following original text character.");
            if (hidden is not 0 and not -1) return;
            var font = Range.Font;
            try { font.Hidden = hidden; }
            finally { Release(font); }
        }

        public void Dispose() => Release(Range);
    }

    private Range PrepareInlineInsertWithOwnedProse(
        Document document, Range insertion, string formulaId,
        out InlineSourceBoundaryGuard? boundary, bool createBookmark = true)
    {
        boundary = null;
        Range? original = null;
        Range? placeholder = null;
        Range? owner = null;
        string? expected = null;
        var hidden = 0;
        try
        {
            if (insertion.End < document.Content.End)
            {
                original = document.Range(insertion.End, insertion.End + 1);
                expected = original.Text ?? string.Empty;
                var font = original.Font;
                try { hidden = font.Hidden; }
                finally { Release(font); }
            }
            placeholder = PrepareInlineBaselineSentinelBeforeInsert(document, insertion,
                formulaId, createBookmark: createBookmark);
            if (expected is not null)
            {
                var start = placeholder.End + InlineMathGuard.Length + InlineBaselineSentinel.Length;
                owner = document.Range(start, start + 1);
                if (!string.Equals(owner.Text ?? string.Empty, expected, StringComparison.Ordinal))
                    throw new InvalidDataException("Preparing the native equation boundary changed the following source character.");
                boundary = new InlineSourceBoundaryGuard(owner, expected, hidden);
                owner = null;
            }
            var result = placeholder;
            placeholder = null;
            return result;
        }
        finally { Release(owner); Release(placeholder); Release(original); }
    }
}
