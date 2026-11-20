using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Word-native field-code discovery followed by a bounded physical-owner proof.
/// Do not marshal every unrelated EMBED field to find one MathType macro field.
/// TextRetrievalMode is changed only on a private range, never on Word's view.
/// </summary>
internal static class MathTypeFieldLocator
{
    internal const string PlaceRef = "MTPlaceRef";
    internal const string Section = "MTEditEquationSection2";

    internal static Field? Find(
        Range scope,
        string macroName,
        bool forward,
        int excludeStart = -1,
        int excludeEnd = -1,
        int? exactCodeStart = null)
    {
        if (scope is null) throw new ArgumentNullException(nameof(scope));
        if (macroName != PlaceRef && macroName != Section)
            throw new ArgumentException("Unsupported MathType field macro.", nameof(macroName));
        var start = scope.Start;
        var end = scope.End;
        Range? search = null;
        Find? find = null;
        TextRetrievalMode? retrieval = null;
        try
        {
            while (end > start)
            {
                // A Find object has its own search continuation state. Reacquire
                // it from each narrowed range instead of reusing a cursor after
                // rejecting a hit (for example an excluded source template).
                Release(find); find = null;
                Release(retrieval); retrieval = null;
                Release(search); search = scope.Duplicate;
                search.SetRange(start, end);
                retrieval = search.TextRetrievalMode;
                retrieval.IncludeFieldCodes = true;
                retrieval.IncludeHiddenText = true;
                find = search.Find;
                find.ClearFormatting();
                find.Text = macroName;
                find.Forward = forward;
                find.Wrap = WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = true;
                find.MatchWildcards = false;
                find.MatchSoundsLike = false;
                find.MatchAllWordForms = false;
                if (!find.Execute()) return null;
                var owner = ResolveOwner(search, macroName, start, end, forward,
                    excludeStart, excludeEnd, exactCodeStart, out var paragraphStart, out var paragraphEnd);
                if (owner is not null) return owner;

                // Word's text search snaps a start inside a field back to that
                // field's code, so advancing to only hit.End can repeat forever.
                // We have inspected every local macro in this paragraph; advance
                // across that proven paragraph, not an arbitrary character/token.
                if (forward)
                {
                    if (paragraphEnd <= start) throw new InvalidDataException("Word field search made no forward progress.");
                    start = paragraphEnd;
                }
                else
                {
                    if (paragraphStart >= end) throw new InvalidDataException("Word field search made no backward progress.");
                    end = paragraphStart;
                }
            }
            return null;
        }
        finally { Release(find); Release(retrieval); Release(search); }
    }

    // An already-owned newly materialized field must not depend on Find's text
    // search index becoming current within the same Word UI transaction. Query
    // only its known paragraph and prove the exact physical Code.Start instead.
    internal static Field? AtCodeStart(Range owner, string macroName, int codeStart)
    {
        Fields? fields = null;
        try
        {
            fields = WordFormulaHost.GetLocalFields(owner);
            var count = fields.Count;
            for (var index = 1; index <= count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    if (field.Type != WdFieldType.wdFieldMacroButton) continue;
                    code = field.Code;
                    if (code.Start != codeStart || !code.InStory(owner)
                        || code.Start < owner.Start || code.End > owner.End
                        || !MatchesMacroCode(code.Text, macroName)) continue;
                    var result = field;
                    field = null;
                    return result;
                }
                finally { Release(code); Release(field); }
            }
            return null;
        }
        finally { Release(fields); }
    }

    private static Field? ResolveOwner(
        Range hit, string macroName, int scopeStart, int scopeEnd, bool forward,
        int excludeStart, int excludeEnd, int? exactCodeStart,
        out int paragraphStart, out int paragraphEnd)
    {
        paragraphStart = paragraphEnd = -1;
        Range? paragraph = null;
        Fields? fields = null;
        Field? selected = null;
        var bestStart = forward ? int.MaxValue : int.MinValue;
        try
        {
            paragraph = hit.Duplicate;
            paragraph.Expand(WdUnits.wdParagraph);
            paragraphStart = paragraph.Start;
            paragraphEnd = paragraph.End;
            if (paragraphEnd <= scopeStart || paragraphStart >= scopeEnd)
                throw new InvalidDataException("Word field-code Find escaped the requested paragraph scope.");
            fields = WordFormulaHost.GetLocalFields(paragraph);
            var count = fields.Count;
            for (var index = 1; index <= count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    if (field.Type != WdFieldType.wdFieldMacroButton) continue;
                    code = field.Code;
                    var fieldStart = code.Start - 1;
                    if (!code.InStory(hit) || fieldStart < scopeStart || fieldStart >= scopeEnd
                        || fieldStart < paragraphStart || code.End > paragraphEnd
                        || (exactCodeStart.HasValue && code.Start != exactCodeStart.Value)
                        || (fieldStart >= excludeStart && fieldStart < excludeEnd)
                        || !MatchesMacroCode(code.Text, macroName))
                        continue;
                    if (forward ? fieldStart >= bestStart : fieldStart <= bestStart) continue;
                    Release(selected);
                    selected = field;
                    field = null;
                    bestStart = fieldStart;
                }
                finally { Release(code); Release(field); }
            }
            var result = selected;
            selected = null;
            return result;
        }
        finally { Release(selected); Release(fields); Release(paragraph); }
    }

    internal static bool MatchesMacroCode(string? code, string macroName)
    {
        if (string.IsNullOrEmpty(code) || string.IsNullOrEmpty(macroName)) return false;
        var index = 0;
        string ReadToken()
        {
            while (index < code!.Length && char.IsWhiteSpace(code[index])) index++;
            var start = index;
            while (index < code.Length && !char.IsWhiteSpace(code[index])
                && code[index] != '\u0013' && code[index] != '\u0014' && code[index] != '\u0015')
                index++;
            return code.Substring(start, index - start);
        }
        return string.Equals(ReadToken(), "MACROBUTTON", StringComparison.OrdinalIgnoreCase)
            && string.Equals(ReadToken(), macroName, StringComparison.OrdinalIgnoreCase);
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
