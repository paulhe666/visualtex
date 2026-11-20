using System;

namespace VisualTeX.WindowsOffice.Contracts;

public readonly struct FormulaEquationTagSplit
{
    public FormulaEquationTagSplit(string latex, string? equationTag)
    {
        Latex = latex;
        EquationTag = equationTag;
    }

    public string Latex { get; }
    public string? EquationTag { get; }
}

public static class FormulaEquationTag
{
    public static FormulaEquationTagSplit Extract(string? source)
    {
        var trimmed = (source ?? string.Empty).Trim();
        if (trimmed.Length == 0) return new FormulaEquationTagSplit(string.Empty, null);

        for (var index = trimmed.LastIndexOf("\\tag", StringComparison.Ordinal);
             index >= 0;
             index = trimmed.LastIndexOf("\\tag", index - 1, StringComparison.Ordinal))
        {
            var cursor = index + "\\tag".Length;
            if (cursor < trimmed.Length && trimmed[cursor] == '*') cursor++;
            if (cursor < trimmed.Length
                && trimmed[cursor] != '{'
                && !char.IsWhiteSpace(trimmed[cursor]))
                continue;
            while (cursor < trimmed.Length && char.IsWhiteSpace(trimmed[cursor])) cursor++;
            if (cursor >= trimmed.Length) continue;

            if (trimmed[cursor] == '{')
            {
                var close = MatchingClosingBrace(trimmed, cursor);
                if (close < 0 || trimmed.Substring(close + 1).Trim().Length != 0) continue;
                var tag = trimmed.Substring(cursor + 1, close - cursor - 1).Trim();
                if (tag.Length == 0) continue;
                return new FormulaEquationTagSplit(trimmed.Substring(0, index).TrimEnd(), tag);
            }

            var bareTag = trimmed.Substring(cursor).Trim();
            if (bareTag.Length == 0 || bareTag.IndexOfAny(new[] { '{', '}' }) >= 0) continue;
            return new FormulaEquationTagSplit(trimmed.Substring(0, index).TrimEnd(), bareTag);
        }

        return new FormulaEquationTagSplit(trimmed, null);
    }

    public static string Attach(string? source, string? equationTag)
    {
        var body = Extract(source).Latex;
        var tag = equationTag?.Trim();
        return string.IsNullOrWhiteSpace(tag) ? body : $"{body}\\tag{{{tag}}}";
    }

    private static int MatchingClosingBrace(string source, int openIndex)
    {
        var depth = 0;
        for (var index = openIndex; index < source.Length; index++)
        {
            if (source[index] == '\\')
            {
                index++;
                continue;
            }
            if (source[index] == '{') depth++;
            else if (source[index] == '}')
            {
                depth--;
                if (depth == 0) return index;
            }
        }
        return -1;
    }
}
