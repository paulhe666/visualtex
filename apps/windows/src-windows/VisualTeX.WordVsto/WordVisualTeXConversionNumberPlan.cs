using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;
internal static partial class WordFormulaNumberingKernel
{
    // One synchronous preflight models the FINAL target-family order, including
    // equations which already exist. Reverse insertion cannot infer that order
    // from the partially materialized target bookmark collection.
    internal sealed class VisualTeXConversionNumberPlan : IDisposable
    {
        private readonly Dictionary<string, NumberPlan> _plans = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<(WordFormulaHostDescriptor Host, Range LiveRange, string OperationId)> _existingTables = new();
        internal NumberPlan For(string operationId) => _plans[operationId];

        // A target whose own transaction was rolled back (independent redraw
        // records) no longer owns the fields it created; forget them so the
        // final ordered field pass expects only materialized numbers.
        internal void Discard(string operationId)
        {
            if (!_plans.TryGetValue(operationId, out var plan)) return;
            plan.WrittenBodyLayout?.Dispose();
            plan.WrittenBodyLayout = null;
            plan.CreatedBodyFields?.Dispose();
            plan.CreatedBodyFields = null;
        }

        internal static VisualTeXConversionNumberPlan Capture(Document document,
            IEnumerable<(string OperationId, int Position)> newTargets, string formatId)
        {
            using var metric = WordOperationMetrics.Measure("FormatConversion.CaptureFinalVisualTeXNumberOrder");
            var result = new VisualTeXConversionNumberPlan();
            Range? content = null;
            try
            {
                var final = newTargets.ToList();
                content = document.Content;
                var existing = WordFormulaHostResolver.CaptureScopeKindIndex(document, content, WordFormulaHostKind.VisualTeX);
                foreach (var host in existing.Where(item => item.Display))
                {
                    var numbering = WordFormulaNumberingResolver.ResolveLocal(document, host);
                    if (!numbering.Numbered) continue;
                    var operationId = Guid.NewGuid().ToString("N");
                    final.Add((operationId, host.Range.Start));
                    if (host.WithinTable || numbering.ContainerKind == WordFormulaNumberingContainerKind.CanonicalBodyTable)
                        result._existingTables.Add((host, WordFormulaHostSemanticReader.CreateRange(document, host.Range), operationId));
                }
                var format = EquationNumberFormat.Resolve(formatId);
                var scopes = format.UsesHeading
                    ? WordEquationNumbering.CaptureHeadingScopesAtPositions(document, formatId, final.Select(item => item.Position).ToArray())
                    : new Dictionary<int, ResolvedEquationHeadingScope>();
                var ordinals = new Dictionary<int, int>();
                foreach (var item in final.OrderBy(item => item.Position))
                {
                    var scopeStart = int.MinValue;
                    var prefix = string.Empty;
                    var headingNumberText = string.Empty;
                    if (format.UsesHeading && scopes.TryGetValue(item.Position, out var scope))
                    {
                        scopeStart = scope.ScopeStart;
                        headingNumberText = scope.NumberText;
                        prefix = headingNumberText + format.Separator;
                    }
                    var ordinal = ordinals.TryGetValue(scopeStart, out var previous) ? previous + 1 : 1;
                    ordinals[scopeStart] = ordinal;
                    result._plans.Add(item.OperationId, new NumberPlan
                    {
                        Ordinal = ordinal,
                        Prefix = prefix,
                        BodyNumbering = WordVisualTeXParagraphNumbering.PrepareNumberPlan(format, headingNumberText),
                    });
                }
                return result;
            }
            catch { result.Dispose(); throw; }
            finally { Release(content); }
        }

        internal void ApplyExistingTables(Document document)
        {
            foreach (var item in _existingTables.OrderByDescending(item => item.LiveRange.Start))
            {
                item.Host.Range = new WordFormulaRangeAddress
                { StoryType = item.LiveRange.StoryType, Start = item.LiveRange.Start, End = item.LiveRange.End };
                var numbering = WordFormulaNumberingResolver.ResolveLocal(document, item.Host);
                var plan = For(item.OperationId);
                RewriteCanonicalNumberLabel(document, item.Host, numbering, plan.Ordinal, plan.Prefix);
            }
        }

        internal CreatedBodyFieldCursor CreateBodyFieldCursor() =>
            new(_plans.Values.Where(plan => plan.CreatedBodyFields is not null)
                .Select(plan => plan.CreatedBodyFields!));

        // The complete Word field stream remains the ordering authority. Sort
        // retained owners once, but compare LIVE code positions at every merge
        // step because earlier field evaluations can change later positions.
        internal sealed class CreatedBodyFieldCursor
        {
            private readonly WordVisualTeXParagraphNumbering.CreatedBodyFields[] _fields;
            private int _next;

            internal CreatedBodyFieldCursor(IEnumerable<WordVisualTeXParagraphNumbering.CreatedBodyFields> fields)
            {
                _fields = fields.Select(field => (Field: field, Position: CurrentPosition(field)))
                    .OrderBy(item => item.Position.Story).ThenBy(item => item.Position.Start)
                    .Select(item => item.Field).ToArray();
            }

            private static (WdStoryType Story, int Start) CurrentPosition(
                WordVisualTeXParagraphNumbering.CreatedBodyFields field)
            {
                Range? code = null;
                try { code = field.Outer.Code; return (code.StoryType, code.Start); }
                finally { Release(code); }
            }

            internal bool TryUpdate(Range currentCode)
            {
                if (_next == _fields.Length) return false;
                var expected = _fields[_next];
                Range? expectedCode = null;
                try
                {
                    expectedCode = expected.Outer.Code;
                    if (!expectedCode.InStory(currentCode))
                        throw new InvalidDataException("The new VisualTeX numbering field left the final field stream's story.");
                    var expectedStart = expectedCode.Start;
                    var currentStart = currentCode.Start;
                    if (expectedStart > currentStart) return false; // A genuinely existing body field.
                    if (expectedStart < currentStart)
                        throw new InvalidDataException("The ordered field stream omitted a newly created VisualTeX number.");
                    expected.Update();
                    _next++;
                    return true;
                }
                finally { Release(expectedCode); }
            }

            internal void EnsureComplete()
            {
                if (_next != _fields.Length)
                    throw new InvalidDataException("The ordered field stream did not evaluate every newly created VisualTeX number.");
            }
        }

        public void Dispose()
        {
            foreach (var item in _existingTables) Release(item.LiveRange);
            _existingTables.Clear();
            foreach (var plan in _plans.Values)
            {
                plan.WrittenBodyLayout?.Dispose();
                plan.WrittenBodyLayout = null;
                plan.CreatedBodyFields?.Dispose();
                plan.CreatedBodyFields = null;
            }
        }
    }
}
