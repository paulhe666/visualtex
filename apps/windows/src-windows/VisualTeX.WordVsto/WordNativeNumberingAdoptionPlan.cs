using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// A native numbered OMath is not necessarily a SEQ field: Word's #(label)
/// equation syntax also admits literal labels. Read that structure once before
/// Undo, adopt only the numeral into real Word fields, and let the native field
/// engine own subsequent counting. No formula ID or whole-paragraph rewrite.
/// </summary>
internal sealed partial class WordNativeNumberingAdoptionPlan : IDisposable
{
    private sealed class Entry
    {
        internal OMath Host = null!;
        internal int Start;
        internal string Xml = string.Empty;
        internal string Label = string.Empty;
        internal NumberStyle Style = null!;
        internal string? SequenceCode;
    }

    private enum ChangeKind { AdoptNumeral, Prefix, TransferReset, Format }
    private sealed class Change
    {
        internal Entry Entry = null!;
        internal ChangeKind Kind;
        internal NumberStyle Style = null!;
        internal int? Restart;
        internal Range? PreparedNumber;
        internal Range? PreparedPrefix;
        internal readonly List<(string Name, int Kind, int Start, int End)> Aliases = new();
    }

    private readonly Document _document;
    private readonly List<Entry> _entries = new();
    private readonly List<Change> _changes = new();
    private readonly List<MathTypeLocalEditScope> _recovery = new();
    private WordNativeOmmlFieldRefresh? _refresh;
    private Document? _formatStaging;
    private WordNativeOmmlLiteralTextStaging? _literalStaging;
    private NumberStyle? _insertionStyle;
    private int? _insertionRestart;
    private readonly int? _insertionPosition;
    internal int Count => _refresh?.Count ?? 0;

    private WordNativeNumberingAdoptionPlan(Document document, int? position)
    { _document = document; _insertionPosition = position; }

    internal static bool TryInferDocumentFormat(
        Document document,
        out int start,
        out string formatId)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        start = int.MaxValue;
        formatId = string.Empty;
        using var plan =
            new WordNativeNumberingAdoptionPlan(
                document,
                position: null);
        // Only the earliest proven number is needed: stop at it instead of
        // exporting the XML of every numbered equation in the document.
        plan._stopAtFirstEntry = true;
        plan.Capture();
        if (plan._entries.Count == 0)
            return false;

        // A legacy document has no persisted VisualTeX format variable. In that
        // case the earliest proven Word-native #(number) host is the document's
        // migration authority, rather than allowing each later insertion path to
        // inherit a different local/default preference.
        var first = plan._entries[0];
        start = first.Start;
        formatId = first.Style.FormatId;
        return true;
    }

    internal static WordNativeNumberingAdoptionPlan Prepare(Document document, int? insertionPosition = null, string? numberFormat = null)
    {
        if (WordFormulaMutationTransaction.IsActiveOnCurrentThread)
            throw new InvalidOperationException("Native numbering adoption must be prepared before Undo.");
        var plan = new WordNativeNumberingAdoptionPlan(document, insertionPosition);
        try
        {
            using (WordOperationMetrics.Measure("NativeNumbering.Adoption.Capture"))
                plan.Capture();
            using var planMetric = WordOperationMetrics.Measure("NativeNumbering.Adoption.Plan");
            if (insertionPosition.HasValue && numberFormat is not null)
                plan.PlanInsertion(
                    insertionPosition.Value,
                    EquationNumberFormat.Resolve(numberFormat).Id);
            else if (numberFormat is not null)
                plan.PlanFormat(numberFormat);
            else if (insertionPosition.HasValue)
                plan.PlanInsertion(insertionPosition.Value);
            else
            {
                for (var index = 0; index < plan._entries.Count;)
                {
                    var end = plan.RunEnd(index);
                    plan.PlanRun(index, end, null);
                    index = end;
                }
            }
            var extra = plan._changes.Where(c => c.Kind == ChangeKind.AdoptNumeral || c.Kind == ChangeKind.Format && c.Entry.SequenceCode is null)
                .Select(c => c.Entry.Host).ToArray();
            var changed = new HashSet<int>(plan._changes.Select(c => c.Entry.Start));
            var earliest = insertionPosition ?? 0;
            if (plan._changes.Count != 0) earliest = Math.Min(earliest, plan._changes.Min(c => c.Entry.Start));
            using var refreshMetric = WordOperationMetrics.Measure("NativeNumbering.Adoption.RefreshPrepare");
            plan._refresh = WordNativeOmmlFieldRefresh.Prepare(document, earliest,
                refreshLiteralPrefixes: !insertionPosition.HasValue && numberFormat is null,
                additionalNativeHosts: extra, externallyRewrittenNumberStarts: changed);
            return plan;
        }
        catch { plan.Dispose(); throw; }
    }

    private void PlanFormat(string formatId)
    {
        var format = EquationNumberFormat.Resolve(formatId);
        var level = format.UsesHeading ? format.HeadingLevel : 0;
        var separator = format.Separator;
        var scopes = WordEquationNumbering.CaptureHeadingScopesAtPositions(
            _document, format.Id, _entries.Select(e => e.Start).ToArray());
        var counters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in _entries)
        {
            var target = entry.Style.Copy();
            target.HeadingLevel = level; target.Separator = separator;
            scopes.TryGetValue(entry.Start, out var scope);
            target.Prefix = level == 0 ? string.Empty : scope?.NumberText ?? "0";
            target.AutomaticHeading = level > 0 && scope is not null
                && scope.ScopeStart != int.MinValue && IsAutomaticHeading(_document, scope.ScopeStart);
            var key = target.SequenceName + "|" + (level == 0 ? "continuous" : (scope?.ScopeStart ?? int.MinValue).ToString(CultureInfo.InvariantCulture));
            counters.TryGetValue(key, out var preceding);
            target.Ordinal = checked(preceding + 1); counters[key] = target.Ordinal;
            AddChange(entry, ChangeKind.Format, target, preceding == 0 ? 1 : (int?)null);
        }
        PrepareFormatNumberStaging();
    }

    private void PrepareFormatNumberStaging()
    {
        var changes = _changes.Where(change => change.Kind == ChangeKind.Format).ToArray();
        if (changes.Length == 0) return;
        if (WordFormulaMutationTransaction.IsActiveOnCurrentThread)
            throw new InvalidOperationException("Native number-format staging must be prepared before Undo.");

        _formatStaging = _document.Application.Documents.Add(Visible: false);
        _formatStaging.TrackRevisions = false;
        Range? content = null;
        try
        {
            content = _formatStaging.Content;
            content.Text = string.Concat(Enumerable.Repeat("\uE101\r", changes.Length));
        }
        finally { Release(content); }

        try
        {
            const string semantic =
                "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x</m:t></m:r></m:oMath>";
            for (var index = changes.Length - 1; index >= 0; index--)
            {
                Range? slot = null;
                Range? inserted = null;
                OMaths? maths = null;
                OMath? math = null;
                Range? exact = null;
                Range? number = null;
                try
                {
                    slot = _formatStaging.Range(index * 2, index * 2 + 1);
                    if (!string.Equals(slot.Text, "\uE101", StringComparison.Ordinal))
                        throw new InvalidDataException("Prepared native-number slot lost its exact boundary.");
                    var change = changes[index];
                    var style = change.Style;
                    var prepared = WordOmmlConverter.BuildWordNativeNumberedOmml(
                        semantic,
                        style.SequenceName,
                        numberBookmarkName: null,
                        style.HeadingLevel,
                        style.HeadingLevel > 0 ? style.Separator : string.Empty,
                        style.Ordinal.ToString(CultureInfo.InvariantCulture),
                        style.HeadingLevel > 0 && !style.AutomaticHeading ? style.Prefix : null,
                        change.Restart);
                    inserted = WordOmmlConverter.ReplaceWithPreparedOmmlDirect(
                        _formatStaging,
                        slot,
                        prepared,
                        display: true,
                        mathFontName: _formatStaging.OMathFontName);
                    maths = inserted.OMaths;
                    if (maths.Count != 1)
                        throw new InvalidDataException("Prepared native-number slot did not materialize one OMath.");
                    math = maths[1];
                    exact = math.Range;
                    number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
                    change.PreparedNumber = number.Duplicate;
                }
                finally
                {
                    Release(number); Release(exact); Release(math); Release(maths);
                    Release(inserted); Release(slot);
                }
            }
        }
        catch
        {
            foreach (var change in changes)
            {
                Release(change.PreparedNumber);
                change.PreparedNumber = null;
            }
            try { _formatStaging.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(_formatStaging); _formatStaging = null;
            throw;
        }
        finally { _document.Activate(); }
    }

    private int RunEnd(int first)
    {
        var end = first + 1;
        while (end < _entries.Count && SameSeries(_entries[first], _entries[end])) end++;
        return end;
    }

    private static bool SameSeries(Entry a, Entry b) =>
        string.Equals(a.Style.Prefix, b.Style.Prefix, StringComparison.Ordinal)
        && string.Equals(a.Style.SequenceName, b.Style.SequenceName, StringComparison.OrdinalIgnoreCase);

    private void PlanInsertion(
        int position,
        string? authoritativeFormatId = null)
    {
        if (_entries.Count == 0) return;

        if (authoritativeFormatId is null)
        {
            PlanLegacyInsertion(
                position);
            return;
        }

        var nearest =
            _entries.FindLastIndex(
                entry => entry.Start < position);
        if (nearest < 0)
            nearest = 0;

        var first = nearest;
        while (first > 0
               && SameSeries(
                   _entries[first - 1],
                   _entries[nearest]))
            first--;
        var end =
            RunEnd(first);

        // Keep native counter/reset adoption tied to the existing Word sequence,
        // but do not let that local host decide the visible format. The document
        // format authority is shared by OMML/MathType/VisualTeX; counters remain
        // independent through their own sequence names.
        PlanRun(
            first,
            end,
            position);

        var format =
            EquationNumberFormat.Resolve(
                authoritativeFormatId);
        var style =
            _entries[first].Style.Copy();
        style.HeadingLevel =
            format.UsesHeading
                ? format.HeadingLevel
                : 0;
        style.Separator =
            format.UsesHeading
                ? format.Separator
                : string.Empty;
        style.Prefix =
            string.Empty;
        style.AutomaticHeading =
            false;

        if (format.UsesHeading)
        {
            var scope =
                WordEquationNumbering
                    .ResolveHeadingScopeAtPosition(
                        _document,
                        position,
                        format.Id);
            var literalPrefix = WordNativeOmmlNumbering.ResolveLiteralHeadingPrefix(_document, scope);
            style.Prefix = scope.NumberText;
            style.AutomaticHeading = literalPrefix is null;
        }

        _insertionStyle =
            style;

        // Only an insertion before the first member of an existing sequence
        // inherits/transfers that series' explicit starting value. Heading
        // changes are owned by SEQ \s and must not manufacture a chapter-local
        // \r reset from the legacy local style.
        _insertionRestart =
            position <= _entries[first].Start
                ? _entries[first].Style.Ordinal
                : null;
    }

    private void PlanLegacyInsertion(int position)
    {
        var nearest = _entries.FindLastIndex(e => e.Start < position);
        if (nearest < 0) nearest = 0;
        var style = _entries[nearest].Style.Copy();
        if (style.HeadingLevel > 0)
        {
            var current = WordEquationNumbering.ResolveHeadingScopeAtPosition(_document, position, style.FormatId);
            var source = WordEquationNumbering.ResolveHeadingScopeAtPosition(_document, _entries[nearest].Start, style.FormatId);
            if (current.ScopeStart != int.MinValue && current.ScopeStart != source.ScopeStart)
            {
                // Carry the series' punctuation across a chapter boundary, not
                // the previous chapter's prefix/counter. An existing following
                // native label in this chapter is a better context than the old one.
                var following = _entries.FindIndex(e => e.Start >= position
                    && string.Equals(e.Style.Prefix, current.NumberText, StringComparison.Ordinal));
                if (following >= 0) nearest = following;
                else
                {
                    style.Prefix = current.NumberText;
                    style.AutomaticHeading = IsAutomaticHeading(_document, current.ScopeStart);
                    _insertionStyle = style; _insertionRestart = 1;
                    return;
                }
            }
        }
        var first = nearest;
        while (first > 0 && SameSeries(_entries[first - 1], _entries[nearest])) first--;
        var end = RunEnd(first);
        _insertionStyle = _entries[first].Style.Copy();
        if (position <= _entries[first].Start)
            _insertionRestart = _entries[first].Style.Ordinal;
        PlanRun(first, end, position);
    }

    private void PlanRun(int first, int end, int? insertion)
    {
        var anchor = _entries[first];
        var hasLiteral = _entries.Skip(first).Take(end - first).Any(e => e.SequenceCode is null);
        for (var index = first; index < end; index++)
        {
            var entry = _entries[index];
            if (entry.SequenceCode is null)
            {
                var restart = index == first && (!insertion.HasValue || insertion.Value > entry.Start)
                    ? entry.Style.Ordinal : (int?)null;
                AddChange(entry, ChangeKind.AdoptNumeral, entry.Style.Copy(), restart);
            }
            else
            {
                if (index == first && insertion.HasValue && insertion.Value <= entry.Start
                    && Regex.IsMatch(entry.SequenceCode, @"(?:^|\s)\\r\s+-?\d+(?=\s|$)", RegexOptions.IgnoreCase))
                    AddChange(entry, ChangeKind.TransferReset, entry.Style.Copy(), null);
                if (hasLiteral && !string.Equals(entry.Style.Separator, anchor.Style.Separator, StringComparison.Ordinal))
                {
                    var target = entry.Style.Copy(); target.Separator = anchor.Style.Separator;
                    AddChange(entry, ChangeKind.Prefix, target, null);
                }
            }
        }
    }

    private void AddChange(Entry entry, ChangeKind kind, NumberStyle style, int? restart)
    {
        var change = new Change { Entry = entry, Kind = kind, Style = style, Restart = restart };
        Range? exact = null; Range? number = null; Range? ordinal = null; Range? prefix = null;
        Bookmarks? bookmarks = null;
        try
        {
            exact = entry.Host.Range;
            if (!_changes.Any(c => c.Entry == entry)) _recovery.Add(MathTypeLocalEditScope.Capture(_document, exact));
            if (kind != ChangeKind.TransferReset)
            {
                number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
                GetParts(entry.Host, entry.Style, out ordinal, out prefix);
                bookmarks = _document.Bookmarks;
                foreach (var name in MathTypeLocalFragment.BookmarkNames(entry.Xml))
                {
                    Bookmark? bookmark = null; Range? range = null;
                    try
                    {
                        if (!bookmarks.Exists(name)) continue;
                        bookmark = bookmarks[name]; range = bookmark.Range;
                        if (!range.InStory(number) || range.Start < number.Start || range.End > number.End) continue;
                        if (range.Start == number.Start && range.End == number.End)
                            change.Aliases.Add((name, 0, 0, 0));
                        else if (range.Start >= ordinal.Start && range.End <= ordinal.End)
                            change.Aliases.Add((name, 1, range.Start - ordinal.Start, range.End - ordinal.Start));
                        else if (range.Start >= prefix.Start && range.End <= prefix.End)
                            change.Aliases.Add((name, 2, range.Start - prefix.Start, range.End - prefix.Start));
                        else throw new InvalidDataException("A native number alias crosses an unsupported field boundary; no numbering was changed.");
                    }
                    finally { Release(range); Release(bookmark); }
                }
            }
            if (kind == ChangeKind.Prefix || kind == ChangeKind.AdoptNumeral && prefix!.End > prefix.Start)
            {
                var text = style.AutomaticHeading ? style.Separator : style.Prefix + style.Separator;
                _literalStaging ??= new WordNativeOmmlLiteralTextStaging(_document);
                change.PreparedPrefix = _literalStaging.Prepare(text, WordCharacterFormatting.Capture(prefix!));
            }
            _changes.Add(change);
        }
        finally { Release(bookmarks); Release(prefix); Release(ordinal); Release(number); Release(exact); }
    }

    internal string PrepareInsertionOmml(string semanticOmml)
    {
        if (!_insertionPosition.HasValue) throw new InvalidOperationException("This is not an insertion plan.");
        if (_insertionStyle is null)
            return WordNativeOmmlNumbering.PrepareNumberedOmml(_document, _insertionPosition.Value, semanticOmml);
        var s = _insertionStyle;
        return WordOmmlConverter.BuildWordNativeNumberedOmml(semanticOmml, s.SequenceName, null,
            s.HeadingLevel, s.Separator, (_insertionRestart ?? checked(s.Ordinal + 1)).ToString(CultureInfo.InvariantCulture),
            s.HeadingLevel > 0 && !s.AutomaticHeading ? s.Prefix : null, _insertionRestart);
    }

    internal void Apply(WordFormulaHostDescriptor? inserted = null)
    {
        if (_refresh is null) throw new ObjectDisposedException(nameof(WordNativeNumberingAdoptionPlan));
        if (inserted is not null) _refresh.AddInserted(inserted);
        using (WordOperationMetrics.Measure("NativeNumbering.Adoption.ApplyChanges"))
            foreach (var change in _changes.OrderBy(c => c.Entry.Start)) ApplyChange(change);
        using (WordOperationMetrics.Measure("NativeNumbering.Adoption.RefreshApply"))
            _refresh.Apply(inserted is null ? 0 : 1);
    }

    internal void VerifyRollback()
    {
        foreach (var recovery in _recovery) recovery.VerifyRollback();
        _refresh?.VerifyRollback();
    }

    private void ApplyChange(Change change)
    {
        var host = change.Entry.Host;
        Range? exact = null; Range? number = null; Range? ordinal = null; Range? prefix = null;
        Fields? fields = null; Field? sequence = null; Range? code = null;
        Bookmarks? bookmarks = null;
        try
        {
            exact = host.Range; number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
            GetParts(host, change.Entry.Style, out ordinal, out prefix);
            if (change.Kind == ChangeKind.TransferReset)
            {
                fields = number.Fields;
                for (var index = 1; index <= fields.Count; index++)
                {
                    Release(sequence); sequence = fields[index];
                    if (!WordNumberingFieldCollection.IsCurrentField(sequence)
                        || sequence.Type != WdFieldType.wdFieldSequence) continue;
                    code = sequence.Code;
                    if (!string.Equals(code.Text, change.Entry.SequenceCode, StringComparison.Ordinal))
                        throw new InvalidDataException("The native sequence reset changed after preflight.");
                    var match = Regex.Match(code.Text, @"(?:^|\s)\\r\s+-?\d+(?=\s|$)", RegexOptions.IgnoreCase);
                    if (!match.Success) throw new InvalidDataException("The prepared native sequence reset disappeared.");
                    // Transfer the reset to a new preceding number without ever
                    // assigning Field.Code.Text or replacing the mathematical body.
                    code.SetRange(code.Start + match.Index, code.Start + match.Index + match.Length);
                    code.Text = " ";
                    return;
                }
                throw new InvalidDataException("The reset owner has no native SEQ.");
            }
            if (change.Kind == ChangeKind.AdoptNumeral
                && (HasAnyCurrentField(number) || number.Text != change.Entry.Label))
                throw new InvalidDataException("The field-free native label changed after preflight.");
            bookmarks = _document.Bookmarks;
            foreach (var alias in change.Aliases)
            {
                Bookmark? bookmark = null;
                try
                {
                    if (!bookmarks.Exists(alias.Name)) throw new InvalidDataException("A native number alias disappeared before adoption.");
                    bookmark = bookmarks[alias.Name]; bookmark.Delete();
                }
                finally { Release(bookmark); }
            }
            if (change.Kind == ChangeKind.Format)
            {
                ReplaceNumberFieldsFromPrepared(
                    host,
                    number,
                    change.PreparedNumber
                    ?? throw new InvalidDataException("The prepared native number-format owner is missing."));
            }
            else if (change.Kind == ChangeKind.AdoptNumeral)
            {
                var s = change.Style;
                var token = s.SequenceName.Any(char.IsWhiteSpace) ? "\"" + s.SequenceName + "\"" : s.SequenceName;
                var instruction = token + (change.Restart.HasValue ? " \\r " + change.Restart.Value.ToString(CultureInfo.InvariantCulture) : string.Empty)
                    + (s.HeadingLevel > 0 ? " \\s " + s.HeadingLevel.ToString(CultureInfo.InvariantCulture) : string.Empty)
                    + " \\* ARABIC";
                fields = ordinal.Fields;
                sequence = fields.Add(ordinal, WdFieldType.wdFieldSequence, instruction, true);
                sequence.Update();
                if (change.PreparedPrefix is not null)
                {
                    // Adopting a literal ordinal creates the same field/text
                    // seam as a format change. Preserve the literal prefix by
                    // the shared prepared-run transport, not math-input repair.
                    Release(prefix); prefix = null; Release(ordinal); ordinal = null;
                    GetParts(host, change.Style, out ordinal, out prefix);
                    prefix.FormattedText = change.PreparedPrefix.FormattedText;
                }
            }
            else
            {
                var before = change.Entry.Style.AutomaticHeading ? change.Entry.Style.Separator
                    : change.Entry.Style.Prefix + change.Entry.Style.Separator;
                if (prefix.Text != before) throw new InvalidDataException("The prepared native numbering prefix changed.");
                var prepared = change.PreparedPrefix
                    ?? throw new InvalidDataException("The native prefix reconciliation has no prepared literal owner.");
                prefix.FormattedText = prepared.FormattedText;
            }
            Release(prefix); prefix = null; Release(ordinal); ordinal = null;
            Release(number); number = null; Release(exact); exact = host.Range;
            number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
            GetParts(host, change.Style, out ordinal, out prefix);
            foreach (var alias in change.Aliases)
            {
                Range? target = null; Bookmark? bookmark = null;
                try
                {
                    target = (alias.Kind == 0 ? number : alias.Kind == 1 ? ordinal : prefix).Duplicate;
                    if (alias.Kind != 0)
                    {
                        var origin = target.Start; var length = target.End - origin;
                        target.SetRange(origin + Math.Min(alias.Start, length), origin + Math.Min(alias.End, length));
                    }
                    bookmark = bookmarks.Add(alias.Name, target);
                }
                finally { Release(bookmark); Release(target); }
            }
        }
        finally { Release(code); Release(sequence); Release(fields); Release(bookmarks); Release(prefix); Release(ordinal); Release(number); Release(exact); }
    }

    private static void ReplaceNumberFieldsFromPrepared(
        OMath host,
        Range number,
        Range preparedNumber)
    {
        Range? liveHost = null;
        Range? liveNumber = null;
        Fields? fields = null;
        Field? field = null;
        Field? sequence = null;
        Field? styleRef = null;
        try
        {
            // The complete native field topology, including literal m:nor runs,
            // was materialized in a hidden Word document BEFORE the target undo
            // record. Copying its exact number argument is one target-document
            // write: no field is incrementally rebuilt while Word's math parser
            // can reclassify adjacent punctuation, and no foreign document is
            // mutated inside Undo.
            number.FormattedText = preparedNumber.FormattedText;

            liveHost = host.Range;
            liveNumber = WordNativeOmmlNumbering.ResolveNumberContentRange(liveHost);
            fields = liveNumber.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(field);
                field = fields[index];
                if (!WordNumberingFieldCollection.IsCurrentField(field))
                    continue;
                if (field.Type == WdFieldType.wdFieldSequence)
                {
                    if (sequence is not null)
                        throw new InvalidDataException("Prepared native number contains multiple current SEQ fields.");
                    sequence = field; field = null;
                }
                else if (field.Type == WdFieldType.wdFieldStyleRef)
                {
                    if (styleRef is not null)
                        throw new InvalidDataException("Prepared native number contains multiple current STYLEREF fields.");
                    styleRef = field; field = null;
                }
                else
                    throw new InvalidDataException("Prepared native number contains an unsupported current field.");
            }
            if (sequence is null)
                throw new InvalidDataException("Prepared native number lost its current SEQ field.");
            styleRef?.Update();
            sequence.Update();
        }
        finally
        {
            Release(styleRef); Release(sequence); Release(field); Release(fields);
            Release(liveNumber); Release(liveHost);
        }
    }

    private static void GetParts(OMath host, NumberStyle style, out Range ordinal, out Range prefix)
    {
        ordinal = null!; prefix = null!;
        Range? exact = null; Range? number = null; Fields? fields = null;
        Field? field = null; Range? code = null; Range? result = null;
        try
        {
            exact = host.Range; number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
            fields = number.Fields;
            var prefixStart = number.Start; var sequenceStart = -1;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(result); result = null; Release(code); code = null; Release(field); field = fields[index];
                if (!WordNumberingFieldCollection.IsCurrentField(field))
                    continue;
                code = field.Code; result = field.Result;
                if (!code.InStory(number) || code.Start <= number.Start || result.End > number.End)
                    throw new InvalidDataException("Native number fields escaped their argument.");
                if (field.Type == WdFieldType.wdFieldSequence)
                {
                    if (ordinal is not null) throw new InvalidDataException("Multiple native ordinal fields.");
                    ordinal = result.Duplicate; sequenceStart = code.Start - 1;
                }
                else if (field.Type == WdFieldType.wdFieldStyleRef) prefixStart = result.End + 1;
                else throw new InvalidDataException("Unrecognized field in a native equation number.");
            }
            if (ordinal is null)
            {
                if (HasAnyCurrentField(number)
                    || number.Text != style.Prefix + style.Separator + style.Ordinal.ToString(CultureInfo.InvariantCulture))
                    throw new InvalidDataException("Native literal number no longer matches its semantic plan.");
                ordinal = number.Duplicate; ordinal.Start = ordinal.End - style.Ordinal.ToString(CultureInfo.InvariantCulture).Length;
                sequenceStart = ordinal.Start;
            }
            if (prefixStart > sequenceStart) throw new InvalidDataException("Native prefix overlaps its ordinal field.");
            prefix = number.Duplicate; prefix.SetRange(prefixStart, sequenceStart);
        }
        catch { Release(ordinal); Release(prefix); ordinal = null!; prefix = null!; throw; }
        finally { Release(result); Release(code); Release(field); Release(fields); Release(number); Release(exact); }
    }

    private bool _stopAtFirstEntry;

    private void Capture()
    {
        var defaultSequence = WordNativeOmmlNumbering.ResolveEquationSequenceName(_document);
        Range? scope = null; Range? search = null; Range? paragraph = null;
        TextRetrievalMode? retrieval = null; Find? find = null; OMaths? maths = null;
        try
        {
            scope = _document.Content; var cursor = scope.Start; var end = scope.End; var seen = new HashSet<int>();
            while (cursor < end)
            {
                Release(maths); maths = null; Release(paragraph); paragraph = null;
                Release(find); find = null; Release(retrieval); retrieval = null; Release(search); search = scope.Duplicate;
                search.SetRange(cursor, end); retrieval = search.TextRetrievalMode;
                retrieval.IncludeFieldCodes = true; retrieval.IncludeHiddenText = true;
                find = search.Find; find.ClearFormatting(); find.Text = "#"; find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop; find.Format = false; find.MatchWildcards = false;
                find.MatchCase = false; find.MatchWholeWord = false; find.MatchSoundsLike = false; find.MatchAllWordForms = false;
                bool found;
                using (WordOperationMetrics.Measure("NativeNumbering.Capture.Find"))
                    found = find.Execute();
                if (!found) break;
                paragraph = WordNumberingFieldCollection.CaptureSearchUnit(search);
                if (paragraph.End <= cursor || paragraph.Start >= end || !paragraph.InStory(scope))
                    throw new InvalidDataException("Native-number discovery repeated or escaped its bounded paragraph.");
                cursor = paragraph.End; maths = paragraph.OMaths; var count = maths.Count;
                for (var index = 1; index <= count; index++)
                {
                    OMath? host = null; Range? exact = null; Range? number = null; Fields? fields = null;
                    try
                    {
                        host = maths[index]; if (host.Type != WdOMathType.wdOMathDisplay) continue;
                        exact = host.Range;
                        if (exact.Start < paragraph.Start || exact.End > paragraph.End || !seen.Add(exact.Start)) continue;
                        string xml;
                        using (WordOperationMetrics.Measure("NativeNumbering.Capture.WordOpenXml"))
                            xml = exact.WordOpenXML;
                        if (!WordOmmlConverter.HasWordNativeEquationNumberHost(xml)) continue;
                        number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
                        var label = WordOmmlConverter.ReadWordNativeEquationNumberText(xml);
                        if (!TryParseLabel(label, out var style)) continue;
                        fields = number.Fields; string? sequenceCode = null; var supported = true;
                        for (var child = 1; child <= fields.Count; child++)
                        {
                            Field? field = null; Range? code = null;
                            try
                            {
                                field = fields[child];
                                if (!WordNumberingFieldCollection.IsCurrentField(field))
                                    continue;
                                code = field.Code;
                                if (field.Type == WdFieldType.wdFieldSequence)
                                {
                                    if (sequenceCode is not null) { supported = false; break; }
                                    sequenceCode = code.Text ?? string.Empty;
                                    var match = Regex.Match(sequenceCode, "^\\s*SEQ\\s+(?:\"(?<q>[^\"]+)\"|(?<u>[^\\s\\\\]+))", RegexOptions.IgnoreCase);
                                    if (!match.Success) { supported = false; break; }
                                    style.SequenceName = match.Groups["q"].Success ? match.Groups["q"].Value : match.Groups["u"].Value;
                                }
                                else if (field.Type == WdFieldType.wdFieldStyleRef) style.AutomaticHeading = true;
                                else supported = false;
                            }
                            finally { Release(code); Release(field); }
                        }
                        if (!supported || style.AutomaticHeading && sequenceCode is null) continue;
                        if (sequenceCode is null) style.SequenceName = defaultSequence;
                        _entries.Add(new Entry { Host = host, Start = exact.Start, Xml = xml, Label = label, Style = style, SequenceCode = sequenceCode });
                        host = null;
                        // Find walks forward, so the first entry is the earliest.
                        if (_stopAtFirstEntry) return;
                    }
                    finally { Release(fields); Release(number); Release(exact); Release(host); }
                }
            }
            _entries.Sort((a, b) => a.Start.CompareTo(b.Start));
        }
        finally { Release(maths); Release(paragraph); Release(find); Release(retrieval); Release(search); Release(scope); }
    }

    private static bool HasAnyCurrentField(Range range)
    {
        Fields? fields = null;
        Field? field = null;
        try
        {
            fields = range.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(field);
                field = fields[index];
                if (WordNumberingFieldCollection.IsCurrentField(field))
                    return true;
            }
            return false;
        }
        finally
        {
            Release(field);
            Release(fields);
        }
    }

    private static bool IsAutomaticHeading(Document document, int position)
    {
        Range? range = null; ListFormat? list = null;
        try { range = document.Range(position, position); list = range.ListFormat; return !string.IsNullOrWhiteSpace(list.ListString); }
        finally { Release(list); Release(range); }
    }

    public void Dispose()
    {
        _refresh?.Dispose(); _refresh = null;
        foreach (var recovery in _recovery) recovery.Dispose(); _recovery.Clear();
        foreach (var change in _changes)
        {
            Release(change.PreparedNumber);
            change.PreparedNumber = null;
            Release(change.PreparedPrefix);
            change.PreparedPrefix = null;
        }
        _literalStaging?.Dispose(); _literalStaging = null;
        if (_formatStaging is not null)
        {
            try { _formatStaging.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(_formatStaging);
            _formatStaging = null;
        }
        foreach (var entry in _entries) Release(entry.Host); _entries.Clear(); _changes.Clear();
    }

    private static void Release(object? value)
    { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
}
