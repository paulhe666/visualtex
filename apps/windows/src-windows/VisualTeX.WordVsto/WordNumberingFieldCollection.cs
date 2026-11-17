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
    // referencesOnly: only the REF and GOTOBUTTON fields are captured (the callers that refresh references need no
    // other field; reading the code of every number field of a long document cost 0.2 s a pass).
    internal static WordNumberingFieldCollection Capture(Document document, int fromPosition = 0, bool referencesOnly = false)
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
                CaptureUnitFields(window, window.Start, window.End, captured, hasRevisions, referencesOnly);
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
                catch (COMException ignored) when (!completed) { IgnoredErrors.Ignore(ignored); }
                finally { Release(field); }
            }
            foreach (var field in scan.Numbers) Release(field);
            Release(fields); Release(target); Release(window); Release(scope);
        }
    }

    /// <summary>
    /// After a number was inserted or removed: evaluates every VisualTeX number field from the paragraph holding
    /// fromPosition to the end of the main story in ONE Range.Fields.Update. The same Word pass renumbers a field in
    /// about 2 ms, where one Field.Update per number costs 13 ms (33 ms with the screen painting) and grows with the
    /// document: 100 formulas behind an insertion took 2.7 s. Measured on a 100-formula document: a range that starts
    /// at a later formula gives the same numbers as the whole story, Word counts a SEQ from the document's start.
    /// Every field that is not SEQ, STYLEREF or MACROBUTTON is locked for the pass (and unlocked after), as in
    /// BulkUpdate: numbering never refreshes TOC, PAGE, DATE or citations. (A range that holds nothing but number
    /// fields and the formulas' own EMBED objects is not visited at all: one read of its text shows that.) Returns the number of SEQ
    /// fields in the range, or -1 when the check below found a number the pass left stale (the caller then updates
    /// the fields one by one). Self-check: the visible number of the first formula behind the change must have
    /// changed in the pass, or, when it did not, re-evaluating it on its own must not change it.
    /// </summary>
    internal static int BulkUpdateFrom(Document document, int fromPosition, bool nativeOnly = false, int headingLevel = 0)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        using var metric = WordOperationMetrics.Measure("Numbering.BulkUpdateFrom");
        var locked = new List<Field>();
        var samples = new SampleSlots();
        var completed = false;
        var sequences = 0;
        Range? scope = null;
        Range? window = null;
        Range? target = null;
        Fields? fields = null;
        try
        {
            scope = document.Content;
            var end = scope.End;
            var start = Math.Max(scope.Start, Math.Min(fromPosition, end));
            using (WordOperationMetrics.Measure("Numbering.BulkUpdateFrom.Start"))
                start = ParagraphStartAt(document, start);
            // A number by chapter restarts at every heading of its level, so a change reaches the next such heading and
            // no further.
            if (headingLevel > 0)
                using (WordOperationMetrics.Measure("Numbering.BulkUpdateFrom.Scope"))
                    end = Math.Min(end, NextHeadingStart(document, start, headingLevel));
            target = document.Range(start, end);
            bool onlyNumberFields;
            using (WordOperationMetrics.Measure("Numbering.BulkUpdateFrom.Scan"))
                onlyNumberFields = HasOnlyNumberFields(target, nativeOnly);
            // One read of the range's text (field codes included) shows whether any field but the number fields is in
            // it; only then is every field visited to lock the others.
            using (WordOperationMetrics.Measure("Numbering.BulkUpdateFrom.Lock"))
            {
                for (var cursor = start; cursor < end && !(onlyNumberFields && samples.Complete);)
                {
                    Release(window);
                    window = CaptureFieldWindow(scope, cursor, end);
                    if (onlyNumberFields) CollectSequenceSamples(window, samples, ref sequences, nativeOnly);
                    else LockForeignFields(window, locked, samples, ref sequences, nativeOnly);
                    cursor = window.End;
                }
            }
            fields = target.Fields;
            int firstError;
            var screenBefore = TryReadScreenUpdating(document);
            using (WordOperationMetrics.Measure("Numbering.BulkUpdateFrom.Update"))
                firstError = fields.Update();
            bool exact;
            using (WordOperationMetrics.Measure("Numbering.BulkUpdateFrom.Check"))
                exact = samples.Chosen is null || SampleShowsThePass(samples);
            WordDoubleClickHook.TraceMessage(
                $"[bulk] from={start} sequences={sequences} locked={locked.Count} firstError={firstError} exact={exact} screenUpdating={screenBefore}");
            completed = true;
            return exact ? sequences : -1;
        }
        finally
        {
            // A failed operation is rolled back by its undo transaction; then
            // unlocking is best effort. After success every lock must lift.
            using (WordOperationMetrics.Measure("Numbering.BulkUpdateFrom.Unlock"))
                foreach (var field in locked)
                {
                    try { field.Locked = false; }
                    catch (COMException ignored) when (!completed) { IgnoredErrors.Ignore(ignored); }
                    finally { Release(field); }
                }
            samples.Dispose();
            Release(fields); Release(target); Release(window); Release(scope);
        }
    }

    /// <summary>
    /// The REF fields of the main story, in document order (the caller owns and releases the fields). Only a field's
    /// type is read to find them, and the code of a REF alone (about 2 ms): the snapshot of Capture reads the code of
    /// every field, number fields and all.
    /// </summary>
    internal static List<Field> CaptureReferenceFields(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        var found = new List<Field>();
        var starts = new HashSet<int>();
        Range? scope = null;
        Range? window = null;
        try
        {
            scope = document.Content;
            var end = scope.End;
            for (var cursor = scope.Start; cursor < end;)
            {
                Release(window);
                window = CaptureFieldWindow(scope, cursor, end);
                Fields? local = null;
                try
                {
                    local = WordFormulaHost.GetLocalFields(window);
                    var count = local.Count;
                    for (var index = 1; index <= count; index++)
                    {
                        Field? field = null;
                        Range? code = null;
                        try
                        {
                            field = local[index];
                            if (field.Type != WdFieldType.wdFieldRef) continue;
                            code = field.Code;
                            // Windows can overlap at their boundary: a field seen there is kept once.
                            if (!starts.Add(code.Start)) continue;
                            found.Add(field);
                            field = null;
                        }
                        finally { Release(code); Release(field); }
                    }
                }
                finally { Release(local); }
                cursor = window.End;
            }
            return found;
        }
        catch
        {
            foreach (var field in found) Release(field);
            throw;
        }
        finally { Release(window); Release(scope); }
    }

    private static bool? TryReadScreenUpdating(Document document)
    {
        try { return document.Application.ScreenUpdating; }
        catch (COMException) { return null; }
    }

    // True when the fields of the range are all SEQ, STYLEREF, MACROBUTTON or EMBED (a VisualTeX formula's own
    // OLE object): read from the range's text with its field codes (WordFieldCode.AllFieldsAre). Any other field,
    // or a code that cannot be read, answers false.
    private static readonly string[] NumberFieldKeywords = { "SEQ", "MACROBUTTON", "EMBED", "STYLEREF" };

    // Word's own equation numbers (a native OMML number is SEQ and STYLEREF only). The numbers of the other kinds
    // are not: a VisualTeX OLE number (SEQ VisualTeXEquation, MACROBUTTON, EMBED) and a MathType one (SEQ MTEqn,
    // MTSec, MTChap) stay as they are when native equations are renumbered.
    private static readonly string[] NativeNumberFieldKeywords = { "SEQ", "STYLEREF" };
    private static readonly string[] OtherKindSequenceMarkers = { "VisualTeX", "MTEqn", "MTSec", "MTChap" };

    internal static bool MentionsOtherKind(string? text)
    {
        if (string.IsNullOrEmpty(text)) return false;
        foreach (var marker in OtherKindSequenceMarkers)
            if (text!.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0) return true;
        return false;
    }

    private static bool HasOnlyNumberFields(Range target, bool nativeOnly)
    {
        try
        {
            var text = ReadTextWithFieldCodes(target);
            if (!nativeOnly) return WordFieldCode.AllFieldsAre(text, NumberFieldKeywords);
            return WordFieldCode.AllFieldsAre(text, NativeNumberFieldKeywords) && !MentionsOtherKind(text);
        }
        catch (COMException) { return false; }
    }

    // The text of a range with its field codes and hidden text. Word's text retrieval mode outlives the Range it was
    // read from: a later Range.Text of the add-in (an agent's context text, a reference's number), of a macro or of
    // another add-in follows it, so the flags are put back as they were.
    internal static string ReadTextWithFieldCodes(Range range)
    {
        TextRetrievalMode? retrieval = null;
        try
        {
            retrieval = range.TextRetrievalMode;
            var codes = retrieval.IncludeFieldCodes;
            var hidden = retrieval.IncludeHiddenText;
            retrieval.IncludeFieldCodes = true;
            retrieval.IncludeHiddenText = true;
            try { return range.Text ?? string.Empty; }
            finally
            {
                try { retrieval.IncludeFieldCodes = codes; retrieval.IncludeHiddenText = hidden; }
                catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
            }
        }
        finally { Release(retrieval); }
    }

    // The SEQ whose result is compared before and after the pass, for the exactness check. It is the visible number
    // of the first formula FOLLOWING the change: a VisualTeX OLE number holds 4 SEQ fields and the 4th is the visible
    // one, so the 2nd formula's is the 8th SEQ of the range; a native equation number is one SEQ, so the 2nd of the
    // range. The first formula of the range is usually the one just written, which was updated on its own and shows
    // nothing of the pass. A range with a single OLE formula keeps its 4th SEQ.
    private sealed class SampleSlots : IDisposable
    {
        private Field? _primary;
        private string _primaryBefore = string.Empty;
        private Field? _fallback;
        private string _fallbackBefore = string.Empty;
        internal bool Complete => _primary is not null;
        internal Field? Chosen => _primary ?? _fallback;
        internal string ChosenBefore => _primary is not null ? _primaryBefore : _fallbackBefore;

        internal void Offer(ref Field? field, int sequences, bool nativeOnly)
        {
            if (field is null) return;
            if (sequences == (nativeOnly ? 2 : 8))
            {
                _primaryBefore = ReadResultText(field);
                _primary = field;
                field = null;
            }
            else if (!nativeOnly && sequences == 4)
            {
                _fallbackBefore = ReadResultText(field);
                _fallback = field;
                field = null;
            }
        }

        public void Dispose()
        {
            Release(_primary); _primary = null;
            Release(_fallback); _fallback = null;
        }
    }

    private static string ReadResultText(Field field)
    {
        Range? result = null;
        try { result = field.Result; return result.Text ?? string.Empty; }
        finally { Release(result); }
    }

    // The pass reached the sample when its result changed (the numbers behind a change all move by one); a sample
    // that did not change is evaluated on its own: it must not change then.
    private static bool SampleShowsThePass(SampleSlots samples)
    {
        var field = samples.Chosen!;
        return !string.Equals(ReadResultText(field), samples.ChosenBefore, StringComparison.Ordinal)
            || EvaluatesUnchanged(field);
    }

    // The sample SEQ fields, when no field needs locking.
    private static void CollectSequenceSamples(Range unit, SampleSlots samples, ref int sequences, bool nativeOnly)
    {
        Fields? local = null;
        try
        {
            local = WordFormulaHost.GetLocalFields(unit);
            var count = local.Count;
            for (var index = 1; index <= count && !samples.Complete; index++)
            {
                Field? field = null;
                try
                {
                    field = local[index];
                    if (field.Type != WdFieldType.wdFieldSequence) continue;
                    sequences++;
                    samples.Offer(ref field, sequences, nativeOnly);
                }
                finally { Release(field); }
            }
        }
        finally { Release(local); }
    }

    /// <summary>
    /// The start of the first paragraph behind <paramref name="from"/> in the built-in heading style of
    /// <paramref name="level"/> (a SEQ with the \s switch of that level starts again there), or int.MaxValue when there
    /// is none, when it cannot be told or when it starts where the search began. Only the style of exactly that level
    /// counts (found by style in a few milliseconds): whether Word also restarts at the higher levels, or at a heading
    /// of another style, only leaves the answer later than it is, so more numbers are updated, never fewer.
    /// </summary>
    internal static int NextHeadingStart(Document document, int from, int level)
    {
        Range? content = null;
        Range? search = null;
        Find? find = null;
        Styles? styles = null;
        Style? style = null;
        try
        {
            content = document.Content;
            search = document.Range(from, content.End);
            find = search.Find;
            find.ClearFormatting();
            find.Text = string.Empty;
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindStop;
            find.Format = true;
            find.MatchWildcards = false;
            styles = document.Styles;
            style = styles[(WdBuiltinStyle)((int)WdBuiltinStyle.wdStyleHeading1 - (level - 1))];
            object findStyle = style;
            find.set_Style(ref findStyle);
            return find.Execute() && search.Start > from ? search.Start : int.MaxValue;
        }
        catch (COMException) { return int.MaxValue; }
        finally { Release(style); Release(styles); Release(find); Release(search); Release(content); }
    }

    // The start of the paragraph that holds position (a field range must not begin inside a paragraph's fields).
    private static int ParagraphStartAt(Document document, int position)
    {
        Range? point = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        try
        {
            point = document.Range(position, position);
            paragraphs = point.Paragraphs;
            paragraph = paragraphs[1];
            range = paragraph.Range;
            return range.Start;
        }
        finally { Release(range); Release(paragraph); Release(paragraphs); Release(point); }
    }

    // Locks the fields of one window that are not number fields; keeps two SEQ fields (every 4th one is a
    // visible number in a PlaceRef: the first of those and one about 128 SEQs later) for the exactness check
    // (each check is one Field.Update of its own, about 15 ms).
    private static void LockForeignFields(Range unit, List<Field> locked, SampleSlots samples, ref int sequences,
        bool nativeOnly = false)
    {
        Fields? local = null;
        try
        {
            local = WordFormulaHost.GetLocalFields(unit);
            var count = local.Count;
            for (var index = 1; index <= count; index++)
            {
                Field? field = null;
                try
                {
                    field = local[index];
                    var type = field.Type;
                    if (type == WdFieldType.wdFieldStyleRef) continue;
                    if (type == WdFieldType.wdFieldMacroButton && !nativeOnly) continue;
                    if (type == WdFieldType.wdFieldSequence && !(nativeOnly && IsOtherKindSequence(field)))
                    {
                        sequences++;
                        samples.Offer(ref field, sequences, nativeOnly);
                        continue;
                    }
                    if (field.Locked) continue;
                    field.Locked = true;
                    locked.Add(field);
                    field = null;
                }
                finally { Release(field); }
            }
        }
        finally { Release(local); }
    }

    private static bool IsOtherKindSequence(Field sequence)
    {
        Range? code = null;
        try
        {
            code = sequence.Code;
            return MentionsOtherKind(code.Text);
        }
        finally { Release(code); }
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

    [ThreadStatic] private static int _noRevisionsProven;

    private sealed class NoRevisionsScope : IDisposable
    {
        private bool _disposed;
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _noRevisionsProven--;
        }
    }

    /// <summary>
    /// For a read that changes nothing: when the document holds no tracked change at all, field
    /// reads on this thread skip the two per-field Revisions queries (TryCaptureCurrentCode) until the
    /// scope ends. Null when the document has revisions (nothing changes then). Listing reference
    /// targets spent about 30 ms a native Word equation in those queries.
    /// </summary>
    internal static IDisposable? AssumeNoRevisionsWhileReading(Document document)
    {
        if (HasRevisions(document)) return null;
        _noRevisionsProven++;
        return new NoRevisionsScope();
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
        bool hasRevisions,
        bool referencesOnly = false)
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
                    if (referencesOnly
                        ? type is not (WdFieldType.wdFieldRef or WdFieldType.wdFieldGoToButton)
                        : type is not (WdFieldType.wdFieldSequence or WdFieldType.wdFieldStyleRef
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

    /// <summary>True when the range holds at most <paramref name="limit"/> equations (one OMaths.Count).</summary>
    internal static bool EquationCountAtMost(Range range, int limit)
    {
        OMaths? maths = null;
        try
        {
            maths = range.OMaths;
            return maths.Count <= limit;
        }
        catch (COMException) { return false; }
        finally { Release(maths); }
    }

    private static bool HasAnyBookmark(Document document)
    {
        Bookmarks? bookmarks = null;
        var changed = false;
        var previous = false;
        try
        {
            bookmarks = document.Bookmarks;
            previous = bookmarks.ShowHidden;
            bookmarks.ShowHidden = true;
            changed = true;
            return bookmarks.Count > 0;
        }
        catch (COMException) { return true; }
        finally
        {
            if (changed && bookmarks is not null)
            {
                try { bookmarks.ShowHidden = previous; }
                catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
            }
            Release(bookmarks);
        }
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
        var restoreRetrieval = false;
        var previousCodes = false;
        var previousHidden = false;
        try
        {
            revisions = document.Revisions;
            if (revisions.Count > 0) return true;
            // A REF field refers to a bookmark; a document without one (hidden _Ref bookmarks counted) holds no
            // reference that works, and the search of all field codes below (70-135 ms at 200 equations) is not needed.
            if (!HasAnyBookmark(document)) return false;
            search = document.Content;
            retrieval = search.TextRetrievalMode;
            previousCodes = retrieval.IncludeFieldCodes;
            previousHidden = retrieval.IncludeHiddenText;
            retrieval.IncludeFieldCodes = true;
            retrieval.IncludeHiddenText = true;
            restoreRetrieval = true;
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
        finally
        {
            // (the text retrieval mode outlives this range, see ReadTextWithFieldCodes)
            if (restoreRetrieval && retrieval is not null)
            {
                try { retrieval.IncludeFieldCodes = previousCodes; retrieval.IncludeHiddenText = previousHidden; }
                catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
            }
            Release(find); Release(retrieval); Release(search); Release(revisions);
        }
    }

    // WPS's Range.Find does not search inside equations, nor the field codes there
    // ("SEQ", "#" are never found). Numbering discovery walks the equations instead:
    // the search unit of the first equation that starts in [start, end), or null.
    internal static Range? NextEquationUnit(Range scope, int start, int end)
    {
        Range? range = null;
        OMaths? maths = null;
        try
        {
            range = scope.Duplicate;
            range.SetRange(start, end);
            maths = range.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? exact = null;
                try
                {
                    math = maths[index];
                    exact = math.Range;
                    if (exact.Start < start || exact.Start >= end) continue;
                    return CaptureSearchUnit(exact);
                }
                finally { Release(exact); Release(math); }
            }
            return null;
        }
        finally { Release(maths); Release(range); }
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
        if (_noRevisionsProven > 0) hasRevisions = false;
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

    /// <summary>
    /// Every field whose code begins in <paramref name="scope"/> (one story), in story order, each
    /// given once to <paramref name="visit"/> (released afterwards; return false to stop). The same
    /// bounded paragraph windows as <see cref="Capture"/>: Fields[index] on a whole document walks
    /// the collection from its start at every index, so a full pass costs O(n^2) COM work (a
    /// 400-formula document: about 12 s a pass, three passes to list reference targets).
    /// No type or revision filter: the callers see exactly what scope.Fields held.
    /// </summary>
    internal static void ForEachFieldInScope(Range scope, Func<Field, bool> visit)
    {
        if (scope is null) throw new ArgumentNullException(nameof(scope));
        if (visit is null) throw new ArgumentNullException(nameof(visit));
        using var metric = WordOperationMetrics.Measure("Fields.ScopePass");
        var start = scope.Start;
        var end = scope.End;
        var lastCodeStart = int.MinValue;
        Range? window = null;
        try
        {
            for (var cursor = start; cursor < end;)
            {
                Release(window);
                window = CaptureFieldWindow(scope, cursor, end);
                Fields? local = null;
                try
                {
                    local = WordFormulaHost.GetLocalFields(window);
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
                            // (windows can overlap at a boundary: a field already given is skipped)
                            if (codeStart <= lastCodeStart || codeStart < start || codeStart > end)
                                continue;
                            lastCodeStart = codeStart;
                            if (!visit(field)) return;
                        }
                        finally { Release(code); Release(field); }
                    }
                }
                finally { Release(local); }
                cursor = window.End;
            }
        }
        finally { Release(window); }
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
