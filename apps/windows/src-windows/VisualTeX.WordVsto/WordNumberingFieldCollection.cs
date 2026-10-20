using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
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

    // fromPosition: only fields that begin at or after it (a renumbering after
    // a change leaves every earlier field's value unchanged).
    internal static WordNumberingFieldCollection Capture(Document document, int fromPosition = 0)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        var captured = new SortedDictionary<int, Entry>();
        Range? scope = null;
        Range? window = null;
        try
        {
            scope = document.Content;
            var cursor = Math.Max(scope.Start, fromPosition);
            var scopeEnd = scope.End;
            var hasRevisions = HasRevisions(document);
            while (cursor < scopeEnd)
            {
                Release(window);
                window = CaptureFieldWindow(scope, cursor, scopeEnd);
                CaptureUnitFields(window, window.Start, window.End, captured, hasRevisions);
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

    // Word evaluates a SEQ field by counting every earlier SEQ of its series, so
    // one Update per number field costs O(n^2) in a long document (500 numbered
    // formulas: ~48 s), while one Range.Fields.Update evaluates them in a single
    // ordered pass (~2 s, same results). That pass counts each series from the
    // start of the updated range, so the range is always the whole main story:
    // starting at the first changed number would restart its series there.
    // Every field whose type is not SEQ, STYLEREF or MACROBUTTON is locked for
    // the pass and unlocked afterwards, so numbering never refreshes TOC, PAGE,
    // DATE, citations, REF or EMBED fields. Another SEQ (a caption series) is
    // only recomputed to its current value. Locking reads the field type alone.
    internal static void BulkUpdate(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        using var metric = WordOperationMetrics.Measure("Numbering.BulkUpdate");
        var scan = new BulkScan();
        var completed = false;
        Range? scope = null;
        Range? window = null;
        Range? target = null;
        Fields? fields = null;
        try
        {
            scope = document.Content;
            var start = scope.Start;
            var end = scope.End;
            using (WordOperationMetrics.Measure("Numbering.BulkUpdate.Lock"))
            {
                for (var cursor = start; cursor < end;)
                {
                    Release(window);
                    window = CaptureFieldWindow(scope, cursor, end);
                    LockUnitFields(window, scan);
                    cursor = window.End;
                }
            }
            target = document.Range(start, end);
            fields = target.Fields;
            int firstError;
            using (WordOperationMetrics.Measure("Numbering.BulkUpdate.Update"))
                firstError = fields.Update();
            // Self-check: re-evaluating the first, middle and last SEQ alone must
            // not change them. A stale result means Word skipped fields in the
            // bulk pass (seen for SEQ fields inside OMath, which therefore never
            // use this path); then every number field is updated one by one.
            var sequences = scan.Sequences;
            var exact = sequences.Count == 0
                || new[] { 0, sequences.Count / 2, sequences.Count - 1 }.Distinct()
                    .All(index => EvaluatesUnchanged(sequences[index]));
            WordDoubleClickHook.TraceMessage(
                $"[bulk] fields={scan.Seen} locked={scan.Locked.Count} types={scan.Types()} firstError={firstError} exact={exact}");
            if (!exact)
                using (WordOperationMetrics.Measure("Numbering.BulkUpdate.FieldByFieldFallback"))
                    foreach (var number in scan.Numbers)
                        number.Update();
            completed = true;
        }
        finally
        {
            // A failed operation is rolled back by its undo transaction; then
            // unlocking is best effort. After success every lock must lift.
            foreach (var field in scan.Locked)
            {
                try { field.Locked = false; }
                catch (COMException) when (!completed) { }
                finally { Release(field); }
            }
            foreach (var field in scan.Numbers) Release(field);
            Release(fields); Release(target); Release(window); Release(scope);
        }
    }

    private sealed class BulkScan
    {
        internal readonly List<Field> Locked = new();
        // Number fields in document order (kept for the exactness fallback).
        internal readonly List<Field> Numbers = new();
        internal readonly Dictionary<WdFieldType, int> TypeCounts = new();
        internal int Seen;
        internal int LastStart = int.MinValue;
        // The SEQ fields among Numbers, in document order.
        internal readonly List<Field> Sequences = new();
        internal string Types() =>
            string.Join(",", TypeCounts.OrderBy(pair => (int)pair.Key).Select(pair => $"{(int)pair.Key}x{pair.Value}"));
    }

    private static bool EvaluatesUnchanged(Field sequence)
    {
        Range? result = null;
        try
        {
            result = sequence.Result;
            var before = result.Text ?? string.Empty;
            sequence.Update();
            Release(result);
            result = sequence.Result;
            return string.Equals(before, result.Text ?? string.Empty, StringComparison.Ordinal);
        }
        finally { Release(result); }
    }

    // Windows can overlap at their boundary; a field already seen there (same
    // code start) is skipped, so nothing is locked or recorded twice.
    private static void LockUnitFields(Range unit, BulkScan scan)
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
                    code = field.Code;
                    var codeStart = code.Start;
                    if (codeStart <= scan.LastStart) continue;
                    scan.LastStart = codeStart;
                    scan.Seen++;
                    var type = field.Type;
                    scan.TypeCounts[type] = scan.TypeCounts.TryGetValue(type, out var seen) ? seen + 1 : 1;
                    if (type is WdFieldType.wdFieldSequence or WdFieldType.wdFieldStyleRef
                        or WdFieldType.wdFieldMacroButton)
                    {
                        if (type == WdFieldType.wdFieldSequence) scan.Sequences.Add(field);
                        scan.Numbers.Add(field);
                        field = null;
                        continue;
                    }
                    if (field.Locked) continue;
                    field.Locked = true;
                    scan.Locked.Add(field);
                    field = null;
                }
                finally { Release(code); Release(field); }
            }
        }
        finally { Release(local); }
    }

    internal static bool HasRevisions(Document document)
    {
        Revisions? revisions = null;
        try
        {
            revisions = document.Revisions;
            return revisions.Count > 0;
        }
        catch (COMException) { return true; }
        finally { Release(revisions); }
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
        SortedDictionary<int, Entry> captured,
        bool hasRevisions)
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
                    if (!TryCaptureCurrentCode(field, out code, out var instruction, hasRevisions)) continue;
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

    // True when the main story may hold a REF field. A whole-word, case-sensitive
    // search of the field codes takes milliseconds; a false positive (the word
    // REF in body text) only costs the caller's full reference handling. Tracked
    // changes can hide a REF from Find, so any revision answers true.
    internal static bool MayHaveReferenceFields(Document document)
    {
        Revisions? revisions = null;
        Range? search = null;
        TextRetrievalMode? retrieval = null;
        Find? find = null;
        try
        {
            revisions = document.Revisions;
            if (revisions.Count > 0) return true;
            search = document.Content;
            retrieval = search.TextRetrievalMode;
            retrieval.IncludeFieldCodes = true;
            retrieval.IncludeHiddenText = true;
            find = search.Find;
            find.ClearFormatting();
            find.Text = "REF";
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindStop;
            find.Format = false;
            find.MatchWholeWord = true;
            find.MatchWildcards = false;
            find.MatchCase = true;
            find.MatchSoundsLike = false;
            find.MatchAllWordForms = false;
            return find.Execute();
        }
        catch (COMException) { return true; }
        finally { Release(find); Release(retrieval); Release(search); Release(revisions); }
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
    // hasRevisions: false when the caller proved the document holds no tracked
    // change at all; then no field can be revision-deleted, and the two per-field
    // Revisions queries (the dominant cost of a field snapshot) are skipped.
    internal static bool TryCaptureCurrentCode(Field field, out Range? currentCode, out string instruction,
        bool hasRevisions = true)
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
                || hasRevisions && IsDeletedRevisionRange(code))
                return false;

            result = field.Result;
            if (hasRevisions && IsDeletedRevisionRange(result)) return false;
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
