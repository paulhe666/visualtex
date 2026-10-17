using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

internal sealed partial class WordNativeNumberingAdoptionPlan
{
    /// <summary>
    /// How many Word-native #(...) numbers are not written exactly as
    /// BuildWordNativeNumberedOmml writes them for the document's format
    /// (docs/equation-format-v2.md): the number's own SEQ label, the format's
    /// prefix and separator, \s L, \r 1 on the first number of each
    /// chapter/section (or of the document), and the ARABIC/MERGEFORMAT
    /// switches. Field-free labels count as well. Read-only; PlanFormat is the
    /// writer that brings every number into this form.
    /// </summary>
    internal static int CountNonCanonicalNumbers(Document document, string formatId)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        using var plan = new WordNativeNumberingAdoptionPlan(document, position: null);
        plan.Capture();
        if (plan._entries.Count == 0) return 0;

        var format = EquationNumberFormat.Resolve(formatId);
        var level = format.UsesHeading ? format.HeadingLevel : 0;
        var scopes = WordEquationNumbering.CaptureHeadingScopesAtPositions(
            document, format.Id, plan._entries.Select(e => e.Start).ToArray());
        var counters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var count = 0;
        foreach (var entry in plan._entries)
        {
            // The same target PlanFormat computes for this entry.
            scopes.TryGetValue(entry.Start, out var scope);
            var prefix = level == 0 ? string.Empty : scope?.NumberText ?? "0";
            var automatic = level > 0 && scope is not null
                && scope.ScopeStart != int.MinValue && IsAutomaticHeading(document, scope.ScopeStart);
            var key = entry.Style.SequenceName + "|" + (level == 0 ? "continuous"
                : (scope?.ScopeStart ?? int.MinValue).ToString(CultureInfo.InvariantCulture));
            counters.TryGetValue(key, out var preceding);
            counters[key] = preceding + 1;
            if (!IsCanonicalNumber(entry, level, format.Separator, prefix, automatic,
                    preceding == 0 ? 1 : null))
                count++;
        }
        return count;
    }

    private static bool IsCanonicalNumber(
        Entry entry, int level, string separator, string prefix, bool automatic, int? restart)
    {
        if (entry.SequenceCode is null) return false;
        var name = entry.Style.SequenceName;
        var token = name.Any(char.IsWhiteSpace) ? "\"" + name + "\"" : name;
        var expected = "SEQ " + token
            + (restart.HasValue ? " \\r " + restart.Value.ToString(CultureInfo.InvariantCulture) : string.Empty)
            + (level > 0 ? " \\s " + level.ToString(CultureInfo.InvariantCulture) : string.Empty)
            + " \\* ARABIC \\* MERGEFORMAT";
        var actual = Regex.Replace(entry.SequenceCode.Trim(), @"\s+", " ");
        if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase)) return false;
        if (level == 0)
            return entry.Style.Prefix.Length == 0 && !entry.Style.AutomaticHeading;
        if (!string.Equals(entry.Style.Separator, separator, StringComparison.Ordinal)) return false;
        return automatic
            ? entry.Style.AutomaticHeading
            : !entry.Style.AutomaticHeading
                && string.Equals(entry.Style.Prefix, prefix, StringComparison.Ordinal);
    }
}
