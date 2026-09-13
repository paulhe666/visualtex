using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>Distinguishes document-level MathType section state from the formula
/// row. The state must survive conversions even when MathType puts both in one
/// paragraph. Detection is read-only; isolation is owned by the conversion undo.</summary>
internal static class MathTypeSourceHost
{
    internal static bool IsSectionStateCode(string? code)
        => Regex.IsMatch(code ?? string.Empty, @"^\s*MACROBUTTON\s+MTEditEquationSection2(?:\s|$)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    internal static bool IsScaffoldWhitespace(string? text)
    {
        foreach (var c in text ?? string.Empty)
            if (c is not ' ' and not '\t' and not '\u00a0') return false;
        return true;
    }

    // Returns the beginning of the TAB+equation+number row, never the beginning
    // of the section-state field. Prefix/suffix prose and unknown fields are not
    // absorbed into this range even when their field results look like numbers.
    internal static bool TryGetSectionPrefixSplit(Document document, Range paragraph, Range equation, out int split)
    {
        split = -1;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        Range? before = null;
        Range? between = null;
        Range? tab = null;
        Range? codeTerminator = null;
        try
        {
            if (paragraph.StoryType != equation.StoryType || equation.Start <= paragraph.Start
                || equation.End > paragraph.End) return false;
            var tabPosition = equation.Start - 1;
            tab = document.Range(tabPosition, equation.Start);
            if (tab.Text != "\t") return false;
            var stateStart = -1;
            var stateEnd = -1;
            fields = WordFormulaHost.GetLocalFields(paragraph);
            for (var i = 1; i <= fields.Count; i++)
            {
                Release(result); result = null;
                Release(code); code = null;
                Release(field); field = fields[i];
                code = field.Code;
                if (!IsSectionStateCode(code.Text)) continue;
                if (stateStart >= 0) return false;
                result = field.Result;
                stateStart = code.Start - 1;
                // A MACROBUTTON with no separator has Result.Start == Result.End
                // at the character AFTER its END. Adding one would claim the TAB
                // (or user text). Read the actual field delimiter instead.
                codeTerminator = document.Range(code.End, code.End + 1);
                stateEnd = codeTerminator.Text == "\u0015" ? code.End + 1
                    : codeTerminator.Text == "\u0014" ? result.End + 1 : -1;
                Release(codeTerminator); codeTerminator = null;
                if (stateStart < paragraph.Start || stateEnd > tabPosition || stateEnd <= stateStart) return false;
            }
            if (stateStart < 0) return false;
            before = document.Range(paragraph.Start, stateStart);
            between = document.Range(stateEnd, tabPosition);
            if (!IsScaffoldWhitespace(before.Text) || !IsScaffoldWhitespace(between.Text)
                || before.Fields.Count != 0 || between.Fields.Count != 0
                || before.InlineShapes.Count != 0 || between.InlineShapes.Count != 0
                || before.OMaths.Count != 0 || between.OMaths.Count != 0) return false;
            split = tabPosition;
            return true;
        }
        finally { Release(codeTerminator); Release(tab); Release(between); Release(before); Release(result); Release(code); Release(field); Release(fields); }
    }
    private static void Release(object? value) { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
}
