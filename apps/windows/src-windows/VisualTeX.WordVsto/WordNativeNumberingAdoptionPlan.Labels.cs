using System.Globalization;
using System.Text.RegularExpressions;

namespace VisualTeX.WordVsto;

internal sealed partial class WordNativeNumberingAdoptionPlan
{
    internal sealed class NumberStyle
    {
        internal string Prefix = string.Empty;
        internal string Separator = string.Empty;
        internal int Ordinal;
        internal int HeadingLevel;
        internal string SequenceName = string.Empty;
        internal bool AutomaticHeading;
        internal string FormatId => HeadingLevel == 0 ? EquationNumberFormat.ContinuousId
            : HeadingLevel == 1 ? Separator == "." ? EquationNumberFormat.Heading1DotId : EquationNumberFormat.Heading1DashId
            : Separator == "." ? EquationNumberFormat.Heading2DotId : EquationNumberFormat.Heading2DashId;
        internal NumberStyle Copy() => (NumberStyle)MemberwiseClone();
    }

    internal static bool TryParseLabel(string text, out NumberStyle style)
    {
        style = new NumberStyle();
        // The structural #( ... ) proof happens separately. This parser accepts
        // numbering labels, not arbitrary mathematical delimiter expressions.
        var match = Regex.Match(text ?? string.Empty,
            @"^(?:(?<prefix>(?:[0-9]+|[A-Za-z]+)(?:\.[0-9]+)*)(?<separator>[.\-\u2010-\u2013\u2212]))?(?<ordinal>[0-9]+)$",
            RegexOptions.CultureInvariant);
        if (!match.Success || !int.TryParse(match.Groups["ordinal"].Value,
                NumberStyles.None, CultureInfo.InvariantCulture, out var ordinal) || ordinal < 1)
            return false;
        // A leading-zero picture is a distinct native number format, not ARABIC.
        if (match.Groups["ordinal"].Value.Length > 1 && match.Groups["ordinal"].Value[0] == '0') return false;
        style.Prefix = match.Groups["prefix"].Value;
        style.Separator = match.Groups["separator"].Value;
        style.Ordinal = ordinal;
        style.HeadingLevel = style.Prefix.Length == 0 ? 0 : style.Prefix.Contains(".") ? 2 : 1;
        return true;
    }

    /// <summary>
    /// The explicit starting value a series' first number hands over when a
    /// number is inserted before it: a literal label's own numeral (adoption
    /// turns it into SEQ \r N) or the N of its SEQ \r N. A plain SEQ has none.
    /// Its visible ordinal is only Word's count and can be a stale result, for
    /// example right after the number above it was removed.
    /// </summary>
    /// <summary>
    /// Batch numbering: the existing series (named by the index of its first
    /// entry) each new number joins, exactly as one insertion at that position
    /// would choose it, and whether that new number is the earliest one placed
    /// before the series and therefore takes over its explicit start.
    /// entrySeries holds one key per captured entry, in document order; equal
    /// adjacent keys form one series. With no entries the run index is -1.
    /// </summary>
    internal static Dictionary<int, (int RunFirst, bool TakesStart)> MapBatchInsertions(
        IReadOnlyList<int> entryStarts,
        IReadOnlyList<string> entrySeries,
        IEnumerable<int> positions)
    {
        if (entryStarts.Count != entrySeries.Count)
            throw new ArgumentException("Every captured entry needs one series key.", nameof(entrySeries));
        var result = new Dictionary<int, (int RunFirst, bool TakesStart)>();
        var seenRuns = new HashSet<int>();
        foreach (var position in positions.Distinct().OrderBy(position => position))
        {
            if (entryStarts.Count == 0)
            {
                result[position] = (-1, false);
                continue;
            }
            var nearest = -1;
            for (var index = 0; index < entryStarts.Count && entryStarts[index] < position; index++)
                nearest = index;
            if (nearest < 0)
                nearest = 0;
            var first = nearest;
            while (first > 0 && string.Equals(entrySeries[first - 1], entrySeries[nearest], StringComparison.Ordinal))
                first--;
            var earliestInRun = seenRuns.Add(first);
            result[position] = (first, earliestInRun && position <= entryStarts[first]);
        }
        return result;
    }

    internal static int? ExplicitStartingValue(string? sequenceCode, int visibleOrdinal)
    {
        if (sequenceCode is null) return visibleOrdinal;
        var match = Regex.Match(sequenceCode, @"(?:^|\s)\\r\s+(?<value>-?\d+)(?=\s|$)", RegexOptions.IgnoreCase);
        return match.Success && int.TryParse(match.Groups["value"].Value, NumberStyles.AllowLeadingSign,
            CultureInfo.InvariantCulture, out var value) ? value : null;
    }

}
