using System.Linq;
using System.Text.RegularExpressions;

namespace VisualTeX.WindowsOffice.Contracts;

/// <summary>
/// The rows of a display formula written back as LaTeX code that reads back as the same formula. A formula's rows
/// are lines of its source, but inside $$...$$ a line break is only a space to TeX (and to the LaTeX redraw), and an
/// alignment point (&amp;) outside an alignment is an error. Several rows, or a row with alignment points, are
/// therefore written as one aligned block (gathered when no row has an alignment point) with the rows joined by \\.
/// Source that already has an environment or row breaks is returned as it is.
/// </summary>
public static class FormulaLatexRows
{
    private static readonly Regex AlignmentPoint = new(@"(^|[^\\])&");

    public static string ForDisplayCode(string? latex)
    {
        var source = (latex ?? string.Empty).Trim();
        if (source.Contains("\\begin{") || source.Contains("\\\\")) return source;
        var rows = source
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Replace('\v', '\n')
            .Split('\n')
            .Select(row => row.Trim())
            .Where(row => row.Length > 0)
            .ToList();
        var aligned = rows.Any(row => AlignmentPoint.IsMatch(row));
        if (rows.Count < 2 && !aligned) return source;
        var environment = aligned ? "aligned" : "gathered";
        return $"\\begin{{{environment}}}{string.Join(" \\\\ ", rows)}\\end{{{environment}}}";
    }
}
