using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Operation-local, story-ordered snapshot of numbering/reference fields. Word's
/// document.Fields[index] can repeatedly walk unrelated EMBED fields and stall
/// in WWLIB on revised large documents. Native paragraph windows bound each
/// editable collection; text Find is not an authority for field existence.
/// No OLE payload, WordOpenXML, document view or persistent identity is touched.
/// </summary>
internal sealed class WordNumberingFieldCollection : IDisposable
{
    internal sealed class Entry : IDisposable
    {
        private Field? _field;
        private Range? _code;
        internal Field Field => _field!;
        internal Range Code => _code!;
        internal WdFieldType Type { get; }
        internal string Instruction { get; }

        internal Entry(Field field, Range code, WdFieldType type, string instruction)
        { _field = field; _code = code; Type = type; Instruction = instruction; }

        internal Field TakeField()
        {
            var code = _code; _code = null;
            Release(code);
            var field = _field!; _field = null;
            return field;
        }

        public void Dispose()
        {
            var code = _code; _code = null;
            var field = _field; _field = null;
            try { Release(code); }
            finally { Release(field); }
        }
    }

    private readonly List<Entry?> _entries;
    internal int Count => _entries.Count;
    private WordNumberingFieldCollection(IEnumerable<Entry> entries) =>
        _entries = entries.Cast<Entry?>().ToList();

    // Existing callers still receive exactly one owned Field reference. The
    // classified consumer instead takes the complete owning entry; its Code is
    // live while Type/Instruction preserve the same current-projection read.
    internal Field this[int index]
    {
        get
        {
            using var entry = TakeEntry(index);
            return entry.TakeField();
        }
    }

    internal Entry TakeEntry(int index)
    {
        if (index < 1 || index > _entries.Count) throw new ArgumentOutOfRangeException(nameof(index));
        var entry = _entries[index - 1] ?? throw new InvalidOperationException("A numbering-field snapshot entry was already consumed.");
        _entries[index - 1] = null;
        return entry;
    }

    internal static WordNumberingFieldCollection Capture(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        var captured = new SortedDictionary<int, Entry>();
        Range? scope = null;
        Range? window = null;
        try
        {
            scope = document.Content;
            var cursor = scope.Start;
            var scopeEnd = scope.End;
            while (cursor < scopeEnd)
            {
                Release(window);
                window = CaptureFieldWindow(scope, cursor, scopeEnd);
                CaptureUnitFields(window, window.Start, window.End, captured);
                cursor = window.End;
            }
            var result = new WordNumberingFieldCollection(captured.Values);
            captured.Clear();
            return result;
        }
        finally
        {
            foreach (var entry in captured.Values) entry.Dispose();
            Release(window); Release(scope);
        }
    }

    private static Range CaptureFieldWindow(Range story, int cursor, int end)
    {
        Range? window = null;
        Range? tail = null;
        Range? boundary = null;
        try
        {
            window = story.Duplicate;
            window.SetRange(cursor, cursor);
            window.MoveEnd(WdUnits.wdParagraph, 32);
            if (window.End > end) window.End = end;
            if (window.End <= cursor || !window.InStory(story))
                throw new InvalidDataException("Native field discovery did not advance its bounded paragraph window.");
            tail = window.Duplicate;
            tail.SetRange(window.End - 1, window.End);
            boundary = CaptureSearchUnit(tail);
            if (!boundary.InStory(story))
                throw new InvalidDataException("A native field discovery window escaped its story.");
            window.End = Math.Min(end, Math.Max(window.End, boundary.End));
            // A current REF with revised mathematical results can be invisible
            // to all Word Find spellings, including ^d. Scanning every global
            // Revision to repair that omission makes unrelated edits dominate
            // numbering cost. Native paragraph movement gives complete windows
            // without either assumption. Batch 32 paragraphs per native step;
            // promote a partial-table endpoint to the whole structural table.
            // No global Fields/Revisions index, OLE payload, WordOpenXML, layout
            // query or field-code view mutation participates in discovery.
            var result = window; window = null;
            return result;
        }
        finally { Release(boundary); Release(tail); Release(window); }
    }

    private static void CaptureUnitFields(
        Range unit,
        int start,
        int end,
        SortedDictionary<int, Entry> captured)
    {
        Fields? local = null;
        try
        {
            local = WordFormulaHost.GetLocalFields(unit);
            var count = local.Count;
            for (var index = 1; index <= count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = local[index];
                    var type = field.Type;
                    if (type is not (WdFieldType.wdFieldSequence or WdFieldType.wdFieldStyleRef
                        or WdFieldType.wdFieldMacroButton or WdFieldType.wdFieldRef or WdFieldType.wdFieldGoToButton))
                        continue;
                    if (!TryCaptureCurrentCode(field, out code, out var instruction)) continue;
                    var codeStart = code!.Start;
                    // The physical field begin marker owns discovery. Its code
                    // may span a window boundary; a whole-code containment test
                    // would drop that otherwise current multi-paragraph owner.
                    if (codeStart < start || codeStart > end || !code.InStory(unit)
                        || captured.ContainsKey(codeStart)) continue;
                    captured.Add(codeStart, new Entry(field, code, type, instruction));
                    field = null;
                    code = null;
                }
                finally { Release(code); Release(field); }
            }
        }
        finally { Release(local); }
    }

    internal static Range CaptureSearchUnit(Range match)
    {
        Range? unit = null;
        Tables? tables = null;
        try
        {
            unit = match.Duplicate;
            unit.Expand(WdUnits.wdParagraph);
            tables = match.Tables;
            for (var index = 1; index <= tables.Count; index++)
            {
                Table? table = null; Range? tableRange = null;
                try
                {
                    table = tables[index]; tableRange = table.Range;
                    if (!tableRange.InStory(match) || tableRange.Start > match.Start || tableRange.End < match.End)
                        continue;
                    if (tableRange.Start <= unit.Start && tableRange.End > unit.End)
                    {
                        Release(unit); unit = tableRange.Duplicate;
                    }
                }
                finally { Release(tableRange); Release(table); }
            }
            // Office 2021 Find may expand a partial-table search to the ENTIRE
            // table despite wdFindStop, returning a hit before the requested
            // cursor. Treat its actual structural search unit as one unit: inspect
            // every local field/OMath in that containing table once, then leave it.
            // Never advance by arbitrary characters, skip later cells, or query
            // layout-backed wdWithInTable. Body queries remain paragraph-bounded.
            var result = unit; unit = null;
            return result;
        }
        finally { Release(tables); Release(unit); }
    }

    internal static bool IsCurrentField(Field field)
    {
        if (field is null) return false;
        Range? code = null;
        try { return TryCaptureCurrentCode(field, out code, out _); }
        finally { Release(code); }
    }

    // Transfer the current-projection code owner and its text together. Callers
    // inspecting this same field must not discard that proof and immediately
    // read Code/Text again through COM. The caller releases currentCode.
    internal static bool TryCaptureCurrentCode(Field field, out Range? currentCode, out string instruction)
    {
        currentCode = null;
        instruction = string.Empty;
        Range? code = null;
        Range? result = null;
        try
        {
            code = field.Code;
            // Under Track Changes Word retains the replaced predecessor field in
            // Range.Fields, but its code/result are revision-deleted and Code.Text
            // is normally null. Treat numbering as a current-document projection:
            // inserted/current fields participate; deleted/moved-from/conflict-
            // deleted predecessors never do.
            instruction = code.Text ?? string.Empty;
            if (string.IsNullOrEmpty(instruction)
                || IsDeletedRevisionRange(code))
                return false;

            result = field.Result;
            if (IsDeletedRevisionRange(result)) return false;
            currentCode = code;
            code = null;
            return true;
        }
        catch (COMException)
        {
            return false;
        }
        finally
        {
            Release(result);
            Release(code);
        }
    }

    private static bool IsDeletedRevisionRange(Range range)
    {
        // Revisions is an intersecting context collection, not proof that the
        // queried content is deleted. In particular a native MACROBUTTON has
        // an empty Result at its end; querying that caret can report the deleted
        // predecessor immediately AFTER the new field. Empty content cannot be
        // deleted content, and partial/adjacent revisions cannot own a whole
        // current field. Require actual positive-span deletion coverage instead.
        var start=range.Start; var end=range.End;
        if (end<=start) return false;
        Revisions? revisions = null;
        var deleted=new List<(int Start,int End)>();
        try
        {
            revisions=range.Revisions;
            for(var index=1; index<=revisions.Count; index++)
            {
                Revision? revision=null; Range? span=null;
                try
                {
                    revision=revisions[index];
                    if (revision.Type is not (WdRevisionType.wdRevisionDelete
                        or WdRevisionType.wdRevisionMovedFrom or WdRevisionType.wdRevisionCellDeletion
                        or WdRevisionType.wdRevisionConflictDelete)) continue;
                    span=revision.Range;
                    if (!span.InStory(range)) continue;
                    var left=Math.Max(start,span.Start); var right=Math.Min(end,span.End);
                    if(right>left) deleted.Add((left,right));
                }
                finally { Release(span); Release(revision); }
            }
            var covered=start;
            foreach(var span in deleted.OrderBy(item=>item.Start))
            {
                if(span.Start>covered) return false;
                covered=Math.Max(covered,span.End);
                if(covered>=end) return true;
            }
            return false;
        }
        finally { Release(revisions); }
    }

    internal static bool HasVisualTeXNumbering(Document document)
    {
        using var fields = Capture(document);
        for (var index = 1; index <= fields.Count; index++)
        {
            using var entry = fields.TakeEntry(index);
            if (WordVisualTeXParagraphNumbering.IsPlaceRefCode(entry.Instruction)
                || entry.Type == WdFieldType.wdFieldSequence
                    && System.Text.RegularExpressions.Regex.IsMatch(entry.Instruction,
                        "^\\s*SEQ\\s+(?:VisualTeXEquation|\"VisualTeXEquation\")(?:\\s|$)",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase)) return true;
        }
        return false;
    }

    public void Dispose()
    {
        for (var index = 0; index < _entries.Count; index++)
        {
            var entry = _entries[index];
            _entries[index] = null;
            entry?.Dispose();
        }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
