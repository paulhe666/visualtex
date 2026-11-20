using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static partial class WordVisualTeXParagraphNumbering
{
    internal enum BatchAttachReadiness
    {
        Ready,
        // The OLE owns its paragraph, but the paragraph lacks the justified
        // center/right tab layout of a numbered display (older writers).
        NeedsLayout,
        // Text or other content shares the paragraph; attaching would have to
        // overwrite it.
        OtherContent,
    }

    // Batch numbering preflight: the same paragraph checks Attach performs, but
    // reported instead of thrown, so one unsuitable formula is skipped rather
    // than rolling back the whole batch. Read-only; no OLE payload is opened.
    internal static BatchAttachReadiness CheckBatchAttachReadiness(
        Document document,
        WordFormulaHostDescriptor host)
    {
        Range? hostRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? before = null;
        Range? after = null;
        try
        {
            hostRange = WordFormulaHostSemanticReader.CreateRange(document, host.Range);
            paragraphs = hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                return BatchAttachReadiness.OtherContent;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            var editableEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            if (hostRange.Start < paragraphRange.Start || hostRange.End > editableEnd)
                return BatchAttachReadiness.OtherContent;
            before = document.Range(paragraphRange.Start, hostRange.Start);
            after = document.Range(hostRange.End, editableEnd);
            if (!ContainsOnlyScaffoldWhitespace(before.Text)
                || !ContainsOnlyScaffoldWhitespace(after.Text))
                return BatchAttachReadiness.OtherContent;
            try
            {
                ValidateDisplayParagraphGeometry(document, hostRange, paragraph, paragraphRange);
                return BatchAttachReadiness.Ready;
            }
            catch (InvalidDataException)
            {
                return BatchAttachReadiness.NeedsLayout;
            }
        }
        finally
        {
            Release(after);
            Release(before);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(hostRange);
        }
    }

    // Batch numbering builds each kind of number field once and copies it to
    // the other formulas. Fields.Add evaluates a new SEQ field at once, and Word
    // evaluates a SEQ field by counting every earlier SEQ of its series, so
    // building every number with Fields.Add costs O(n^2) over a batch (500
    // formulas: 23 s for the numbers, 11 s for the heading state). A
    // FormattedText copy keeps the source's cached results; the batch's final
    // pass evaluates every number. A copy also keeps the source's character
    // formatting, which a newly built field takes from the TAB before the
    // number, so sources are kept per TAB formatting.
    internal sealed class BodyFieldCopies : IDisposable
    {
        private readonly Dictionary<string, Range> _sources = new(StringComparer.Ordinal);

        internal static string PlaceRefKey(ResolvedNumberPlan plan, string format) =>
            "P\n" + string.Join("\n", (plan.Segments ?? BuildSegments(plan))
                .Select(segment => (segment.IsField ? "F " : "T ") + segment.Value)) + "\n" + format;

        internal static string HeadingStateKey(string instruction, string format) =>
            "H\n" + instruction + "\n" + format;

        // The visible character formatting of the TAB at position.
        internal static string FormatSignature(Document document, int position)
        {
            Range? character = null;
            Microsoft.Office.Interop.Word.Font? font = null;
            try
            {
                character = document.Range(position, position + 1);
                font = character.Font;
                return string.Join("|",
                    font.Name, font.NameAscii, font.NameFarEast, font.NameOther,
                    font.Size, font.Bold, font.Italic, font.Underline, font.Color,
                    font.Hidden, font.Position, font.Superscript, font.Subscript,
                    font.Spacing, font.Scaling, font.SmallCaps, font.AllCaps,
                    character.HighlightColorIndex, character.NoProofing);
            }
            finally { Release(font); Release(character); }
        }

        // The copied field, or null when no source of this kind exists yet.
        internal Field? TryCopy(Document document, string key, int position, Func<string, bool> isExpectedCode)
        {
            if (!_sources.TryGetValue(key, out var source)) return null;
            Range? insertion = null;
            Range? inserted = null;
            Fields? fields = null;
            Field? field = null;
            Range? code = null;
            try
            {
                var length = source.End - source.Start;
                insertion = document.Range(position, position);
                using (WordOperationMetrics.Measure("Numbering.Attach.CopyField"))
                    insertion.FormattedText = source.FormattedText;
                inserted = document.Range(position, position + length);
                fields = inserted.Fields;
                if (fields.Count < 1)
                    throw new InvalidOperationException("Word did not copy the VisualTeX number field.");
                field = fields[1];
                code = field.Code;
                if (code.Start != position + 1 || !isExpectedCode(code.Text ?? string.Empty))
                    throw new InvalidOperationException("The copied VisualTeX number field is not the expected field.");
                var result = field;
                field = null;
                return result;
            }
            finally
            {
                Release(code);
                Release(field);
                Release(fields);
                Release(inserted);
                Release(insertion);
            }
        }

        // Keeps the whole field (begin to end mark) as the source for its key.
        internal void Remember(Document document, string key, Field field)
        {
            if (_sources.ContainsKey(key)) return;
            Range? code = null;
            Range? result = null;
            try
            {
                code = field.Code;
                result = field.Result;
                _sources[key] = document.Range(code.Start - 1, ResolveOuterFieldEndExclusive(document, code, result));
            }
            finally { Release(result); Release(code); }
        }

        public void Dispose()
        {
            foreach (var source in _sources.Values) Release(source);
            _sources.Clear();
        }
    }

    private static bool IsSameInstruction(string code, string instruction) =>
        string.Equals(
            string.Join(" ", code.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)),
            string.Join(" ", instruction.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)),
            StringComparison.OrdinalIgnoreCase);
}
