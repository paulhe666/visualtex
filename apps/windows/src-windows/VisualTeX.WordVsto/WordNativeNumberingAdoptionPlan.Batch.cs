using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WordVsto;

internal sealed partial class WordNativeNumberingAdoptionPlan
{
    // Batch numbering of existing display equations. Every new number is an
    // insertion into Word's native SEQ series, but the whole batch shares ONE
    // capture of the existing native numbers, ONE plan and ONE ordered field
    // refresh: a per-equation insertion plan would capture and refresh the whole
    // document once per equation. Every existing series is adopted as "update
    // numbers" adopts it (literal #(n) labels become SEQ fields), and only the
    // earliest new number mapped to a series can take over its explicit start,
    // exactly as one insertion before that series would.
    private readonly Dictionary<int, (string SequenceName, int? Restart)> _batchInsertions = new();
    private WordNativeOmmlNumbering.NumberedOmmlBatch? _batchNumbers;

    internal static WordNativeNumberingAdoptionPlan PrepareBatchInsertion(
        Document document,
        IReadOnlyCollection<int> positions)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (positions is null || positions.Count == 0)
            throw new ArgumentException("A batch numbering plan requires at least one position.", nameof(positions));
        if (WordFormulaMutationTransaction.IsActiveOnCurrentThread)
            throw new InvalidOperationException("Native numbering adoption must be prepared before Undo.");

        var plan = new WordNativeNumberingAdoptionPlan(document, position: null);
        try
        {
            using (WordOperationMetrics.Measure("NativeNumbering.Batch.Capture"))
                plan.Capture();

            var ordered = positions.Distinct().OrderBy(position => position).ToArray();
            using (WordOperationMetrics.Measure("NativeNumbering.Batch.Plan"))
            {
                var defaultSequence = WordNativeOmmlNumbering.ResolveEquationSequenceName(document);
                var mapping = MapBatchInsertions(
                    plan._entries.Select(entry => entry.Start).ToArray(),
                    plan._entries.Select(SeriesKey).ToArray(),
                    ordered);

                // Every existing series is adopted; a series that new numbers
                // join is planned with its earliest new number as the insertion.
                var earliestByRun = mapping
                    .Where(pair => pair.Value.RunFirst >= 0)
                    .GroupBy(pair => pair.Value.RunFirst)
                    .ToDictionary(group => group.Key, group => group.Min(pair => pair.Key));
                for (var index = 0; index < plan._entries.Count;)
                {
                    var end = plan.RunEnd(index);
                    plan.PlanRun(
                        index,
                        end,
                        earliestByRun.TryGetValue(index, out var earliest) ? earliest : null);
                    index = end;
                }

                foreach (var pair in mapping)
                {
                    if (pair.Value.RunFirst < 0)
                    {
                        plan._batchInsertions[pair.Key] = (defaultSequence, null);
                        continue;
                    }
                    var first = plan._entries[pair.Value.RunFirst];
                    plan._batchInsertions[pair.Key] = (
                        first.Style.SequenceName,
                        pair.Value.TakesStart
                            ? ExplicitStartingValue(first.SequenceCode, first.Style.Ordinal)
                            : null);
                }

                plan._batchNumbers = WordNativeOmmlNumbering.NumberedOmmlBatch.Capture(document, ordered);
            }

            var extra = plan._changes
                .Where(change => change.Kind == ChangeKind.AdoptNumeral)
                .Select(change => change.Entry.Host)
                .ToArray();
            var changed = new HashSet<int>(plan._changes.Select(change => change.Entry.Start));
            var earliestChange = plan._changes.Count == 0
                ? ordered[0]
                : Math.Min(ordered[0], plan._changes.Min(change => change.Entry.Start));
            using (WordOperationMetrics.Measure("NativeNumbering.Batch.RefreshPrepare"))
                plan._refresh = WordNativeOmmlFieldRefresh.Prepare(
                    document,
                    earliestChange,
                    additionalNativeHosts: extra,
                    externallyRewrittenNumberStarts: changed);
            return plan;
        }
        catch
        {
            plan.Dispose();
            throw;
        }
    }

    // The same equality SameSeries applies (prefix, sequence name ignoring case).
    private static string SeriesKey(Entry entry) =>
        entry.Style.Prefix + "\u0001" + entry.Style.SequenceName.ToUpperInvariant();

    internal string PrepareBatchInsertionOmml(int position, string semanticOmml)
    {
        if (_batchNumbers is null || !_batchInsertions.TryGetValue(position, out var insertion))
            throw new InvalidOperationException("The equation is not part of this batch numbering plan.");
        return _batchNumbers.Prepare(position, semanticOmml, insertion.SequenceName, insertion.Restart);
    }

    // Called right after each new numbered equation is written, while its
    // address is exact; the refresh keeps its native OMath owner from then on.
    internal void AddBatchInserted(WordFormulaHostDescriptor inserted)
    {
        if (_refresh is null)
            throw new ObjectDisposedException(nameof(WordNativeNumberingAdoptionPlan));
        _refresh.AddInserted(inserted);
    }

    internal void ApplyBatch(int insertedCount)
    {
        if (_refresh is null)
            throw new ObjectDisposedException(nameof(WordNativeNumberingAdoptionPlan));
        using (WordOperationMetrics.Measure("NativeNumbering.Adoption.ApplyChanges"))
            foreach (var change in _changes.OrderBy(change => change.Entry.Start))
                ApplyChange(change);
        using (WordOperationMetrics.Measure("NativeNumbering.Adoption.RefreshApply"))
            _refresh.Apply(insertedCount);
    }
}
