using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Operation-local native numbering updates. Candidate discovery and XML proof
/// precede the Word undo transaction; applying the plan performs only bounded
/// native field updates. Retained native OMath owners supply fresh exact ranges
/// after preceding insertions; raw Range boundary gravity is not object identity.
/// No durable identity, document OLE index, or heading walk is involved.
/// </summary>
internal sealed class WordNativeOmmlFieldRefresh : IDisposable
{
    private readonly Document _document;
    private readonly List<OMath> _hosts = new();
    private readonly Dictionary<OMath, LiteralPrefixUpdate> _literalPrefixes = new();
    private WordNativeOmmlLiteralTextStaging? _literalStaging;
    private readonly List<MathTypeLocalEditScope> _literalRecovery = new();
    private readonly HashSet<string> _numberAliases = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<Field> _references = new();
    private int _insertedCount;
    internal int Count => _hosts.Count;

    private sealed class LiteralPrefixUpdate
    {
        internal string Before = string.Empty;
        internal string After = string.Empty;
        internal Range? PreparedText;
        internal readonly List<(string Name, bool WholeNumber)> Aliases = new();

        internal void Apply(Document document, OMath host)
        {
            Range? prefix = null;
            Bookmarks? bookmarks = null;
            try
            {
                prefix = GetLiteralPrefix(host, out _)
                    ?? throw new InvalidDataException("The prepared literal native number changed topology.");
                if (!string.Equals(prefix.Text ?? string.Empty, Before, StringComparison.Ordinal))
                    throw new InvalidDataException("The literal native number prefix changed after preflight.");
                bookmarks = document.Bookmarks;
                foreach (var alias in Aliases)
                {
                    Bookmark? bookmark = null;
                    try
                    {
                        if (!bookmarks.Exists(alias.Name)) throw new InvalidDataException("A captured number reference disappeared before refresh.");
                        bookmark = bookmarks[alias.Name]; bookmark.Delete();
                    }
                    finally { Release(bookmark); }
                }
                // Native literal semantics are part of the prepared run, not
                // a formatting repair applied after Word parses ordinary text.
                // This writes only the target; its private source is immutable
                // for the complete undo transaction.
                var prepared = PreparedText
                    ?? throw new InvalidDataException("The native literal update has no prepared text owner.");
                prefix.FormattedText = prepared.FormattedText;
            }
            finally { Release(bookmarks); Release(prefix); }
        }

        internal void RestoreAliases(Document document, OMath host)
        {
            Range? exact = null;
            Range? number = null;
            Range? prefix = null;
            Bookmarks? bookmarks = null;
            try
            {
                exact = host.Range;
                number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
                prefix = GetLiteralPrefix(host, out _)
                    ?? throw new InvalidDataException("The refreshed native prefix lost its sequence boundary.");
                if (!string.Equals(prefix.Text ?? string.Empty, After, StringComparison.Ordinal))
                    throw new InvalidDataException("Word changed the prepared literal chapter prefix.");
                bookmarks = document.Bookmarks;
                foreach (var alias in Aliases)
                {
                    Bookmark? bookmark = null;
                    try { bookmark = bookmarks.Add(alias.Name, alias.WholeNumber ? number : prefix); }
                    finally { Release(bookmark); }
                }
            }
            finally { Release(bookmarks); Release(prefix); Release(number); Release(exact); }
        }
    }

    private static Range? GetLiteralPrefix(OMath host, out int level)
    {
        level = 0;
        Range? exact = null;
        Range? number = null;
        Fields? fields = null;
        Field? field = null;
        Field? sequence = null;
        Range? code = null;
        try
        {
            exact = host.Range; fields = exact.Fields;
            // Range.Fields is revision-inclusive. After a tracked replacement
            // Word keeps the deleted predecessor SEQ in the same collection, so
            // raw Fields.Count is not the current numbering topology. Project
            // every local field through the same current-field predicate used by
            // document-wide numbering discovery.
            var currentCount = 0;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(field); field = fields[index];
                if (!WordNumberingFieldCollection.IsCurrentField(field))
                    continue;
                currentCount++;
                if (field.Type == WdFieldType.wdFieldSequence && sequence is null)
                {
                    sequence = field;
                    field = null;
                }
            }
            if (currentCount != 1 || sequence is null) return null;
            code = sequence.Code;
            var instruction = code.Text ?? string.Empty;
            if (!instruction.TrimStart().StartsWith("SEQ ", StringComparison.OrdinalIgnoreCase)) return null;
            var restart = System.Text.RegularExpressions.Regex.Match(instruction, @"(?:^|\s)\\s\s+([12])(?:\s|$)");
            if (!restart.Success) return null;
            level = int.Parse(restart.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
            var begin = code.Start - 1;
            if (begin < number.Start || begin > number.End || !code.InStory(number))
                throw new InvalidDataException("The native sequence does not start inside its number argument.");
            number.End = begin;
            var result = number; number = null;
            return result;
        }
        finally { Release(code); Release(sequence); Release(field); Release(fields); Release(number); Release(exact); }
    }

    private LiteralPrefixUpdate? PrepareLiteralPrefix(Document document, OMath host)
    {
        Range? exact = null;
        Range? prefix = null;
        Range? number = null;
        Bookmarks? bookmarks = null;
        try
        {
            prefix = GetLiteralPrefix(host, out var level);
            if (prefix is null) return null;
            var before = prefix.Text ?? string.Empty;
            // Retain the equation's own separator/scope, not an unrelated current
            // create preference. Unrecognized user-written suffixes are not ours
            // to reinterpret; their native fields still update normally.
            var separator = before.EndsWith("-", StringComparison.Ordinal) ? "-"
                : before.EndsWith(".", StringComparison.Ordinal) ? "." : null;
            if (separator is null) return null;
            var formatId = level == 1
                ? separator == "." ? EquationNumberFormat.Heading1DotId : EquationNumberFormat.Heading1DashId
                : separator == "." ? EquationNumberFormat.Heading2DotId : EquationNumberFormat.Heading2DashId;
            exact = host.Range;
            var scope = WordEquationNumbering.ResolveHeadingScopeAtPosition(document, exact.Start, formatId);
            // Without a matching heading there is no chapter text to invent.
            // Preserve the current prefix rather than silently changing it to0.
            if (scope.ScopeStart == int.MinValue) return null;
            var after = scope.NumberText + separator;
            if (string.Equals(before, after, StringComparison.Ordinal)) return null;
            var plan = new LiteralPrefixUpdate { Before = before, After = after };
            number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
            bookmarks = document.Bookmarks;
            foreach (var name in MathTypeLocalFragment.BookmarkNames(exact.WordOpenXML))
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    if (!bookmarks.Exists(name)) continue;
                    bookmark = bookmarks[name]; range = bookmark.Range;
                    if (!range.InStory(number)) continue;
                    if (range.Start == number.Start && range.End == number.End)
                        plan.Aliases.Add((name, true));
                    else if (range.Start == prefix.Start && range.End == prefix.End)
                        plan.Aliases.Add((name, false));
                }
                finally { Release(range); Release(bookmark); }
            }
            _literalRecovery.Add(MathTypeLocalEditScope.Capture(document, exact));
            _literalStaging ??= new WordNativeOmmlLiteralTextStaging(document);
            plan.PreparedText = _literalStaging.Prepare(after, WordCharacterFormatting.Capture(prefix));
            return plan;
        }
        finally { Release(bookmarks); Release(number); Release(prefix); Release(exact); }
    }

    private WordNativeOmmlFieldRefresh(Document document) { _document = document; }

    internal static WordNativeOmmlFieldRefresh Prepare(Document document, int earliestChange, bool refreshLiteralPrefixes = false,
        IReadOnlyList<OMath>? additionalNativeHosts = null, ISet<int>? externallyRewrittenNumberStarts = null)
    {
        var plan = new WordNativeOmmlFieldRefresh(document);
        Range? scope = null;
        Range? search = null;
        Range? paragraph = null;
        Find? find = null;
        TextRetrievalMode? retrieval = null;
        OMaths? maths = null;
        try
        {
            if (WordFormulaMutationTransaction.IsActiveOnCurrentThread)
                throw new InvalidOperationException("Native-number discovery must precede the target undo transaction.");
            scope = document.Content;
            var cursor = Math.Max(scope.Start, earliestChange);
            var end = scope.End;
            var captured = new HashSet<int>();
            while (cursor < end)
            {
                Release(maths); maths = null;
                Release(paragraph); paragraph = null;
                Release(find); find = null;
                Release(retrieval); retrieval = null;
                Release(search); search = scope.Duplicate;
                search.SetRange(cursor, end);
                retrieval = search.TextRetrievalMode;
                retrieval.IncludeFieldCodes = true;
                retrieval.IncludeHiddenText = true;
                find = search.Find;
                find.ClearFormatting();
                find.Text = "SEQ";
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchWholeWord = true;
                find.MatchWildcards = false;
                find.MatchCase = false;
                find.MatchSoundsLike = false;
                find.MatchAllWordForms = false;
                if (!find.Execute()) break;
                paragraph = WordNumberingFieldCollection.CaptureSearchUnit(search);
                if (paragraph.End <= cursor || paragraph.Start >= end)
                    throw new InvalidDataException("Native-number search escaped or repeated its bounded suffix.");
                // Inspect every OMath in this found paragraph before moving past
                // it; restarting inside a field can snap Find to that same field.
                maths = paragraph.OMaths;
                for (var index = 1; index <= maths.Count; index++)
                {
                    OMath? math = null;
                    Range? exact = null;
                    try
                    {
                        math = maths[index];
                        if (math.Type != WdOMathType.wdOMathDisplay) continue;
                        exact = math.Range.Duplicate;
                        if (exact.Start < earliestChange || exact.End > end || !exact.InStory(scope)
                            || captured.Contains(exact.Start)) continue;
                        if (!WordNativeOmmlNumbering.TryResolveUnownedNativeStructure(document, exact, out _)) continue;
                        captured.Add(exact.Start);
                        plan.CaptureNumberAliases(exact);
                        var prefix = refreshLiteralPrefixes && !(externallyRewrittenNumberStarts?.Contains(exact.Start) ?? false)
                            ? plan.PrepareLiteralPrefix(document, math) : null;
                        if (prefix is not null) plan._literalPrefixes.Add(math, prefix);
                        plan._hosts.Add(math);
                        math = null;
                    }
                    finally { Release(exact); Release(math); }
                }
                cursor = paragraph.End;
            }
            // Field-free native #(label) hosts are structurally proved by the
            // adoption preflight. Retain an independent native owner before any
            // writes so their newly materialized SEQs join the SAME ordered field
            // and reference update, including a predecessor before the insertion.
            if (additionalNativeHosts is not null)
            {
                foreach (var source in additionalNativeHosts)
                {
                    Range? exact = null; OMaths? owners = null; OMath? owner = null;
                    try
                    {
                        exact = source.Range;
                        if (!captured.Add(exact.Start)) continue;
                        if (!exact.InStory(scope) || !WordOmmlConverter.HasWordNativeEquationNumberHost(exact.WordOpenXML))
                            throw new InvalidDataException("A prepared native number lost its structural ownership.");
                        owners = exact.OMaths;
                        if (owners.Count != 1) throw new InvalidDataException("A prepared native number no longer owns one OMath.");
                        owner = owners[1];
                        plan.CaptureNumberAliases(exact);
                        plan._hosts.Add(owner); owner = null;
                    }
                    finally { Release(owner); Release(owners); Release(exact); }
                }
            }
            plan.CaptureReferences();
            return plan;
        }
        catch { plan.Dispose(); throw; }
        finally
        {
            Release(maths); Release(paragraph); Release(find);
            Release(retrieval); Release(search); Release(scope);
        }
    }

    private void CaptureNumberAliases(Range exact)
    {
        Range? number = null;
        Bookmarks? bookmarks = null;
        try
        {
            number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
            bookmarks = _document.Bookmarks;
            foreach (var name in MathTypeLocalFragment.BookmarkNames(exact.WordOpenXML))
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    if (!bookmarks.Exists(name)) continue;
                    bookmark = bookmarks[name]; range = bookmark.Range;
                    if (range.InStory(number) && range.Start >= number.Start && range.End <= number.End)
                        _numberAliases.Add(name);
                }
                finally { Release(range); Release(bookmark); }
            }
        }
        finally { Release(bookmarks); Release(number); }
    }

    private void CaptureReferences()
    {
        if (_numberAliases.Count == 0) return;
        using var fields = WordNumberingFieldCollection.Capture(_document);
        var seen = new HashSet<int>();
        for (var index = 1; index <= fields.Count; index++)
        {
            using var entry = fields.TakeEntry(index);
            if (entry.Type != WdFieldType.wdFieldRef) continue;
            if (!WordEquationReferenceFields.TryReadReferenceBookmark(entry.Instruction, out var name)
                || !_numberAliases.Contains(name)
                || !seen.Add(entry.Code.Start))
                continue;
            _references.Add(entry.TakeField());
        }
    }

    internal void AddInserted(WordFormulaHostDescriptor host)
    {
        if (host.Kind != WordFormulaHostKind.Omml || !host.Numbering.Numbered)
            throw new ArgumentException("The numbering plan requires one newly inserted numbered OMath.");
        Range? range = null;
        Range? exact = null;
        OMaths? maths = null;
        OMath? math = null;
        try
        {
            range = WordFormulaHostSemanticReader.CreateRange(_document, host.Range);
            maths = range.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException("The inserted numbering owner is not one native OMath.");
            math = maths[1]; exact = math.Range;
            if (exact.Start != range.Start || exact.End != range.End || !exact.InStory(range))
                throw new InvalidDataException("The inserted native numbering owner escaped its exact host range.");
            _hosts.Add(math); math = null;
            _insertedCount++;
        }
        finally { Release(math); Release(maths); Release(exact); Release(range); }
    }

    internal void Apply(int expectedInserted)
    {
        if (_insertedCount != expectedInserted)
            throw new InvalidDataException("The native-number update plan is missing converted hosts.");
        WordFormulaMutationTransaction.TraceBoundary(_document.Application, "native-number-plan-before-update");
        // A raw range starting exactly where a preceding paragraph replacement
        // ends can expand/collapse through Word's boundary gravity. Reacquire the
        // actual OMath.Range from its native owner, never reinterpret that gravity
        // as formula identity. Sorting uses current positions once; field updates
        // reacquire them again because earlier result widths can change.
        int Position(OMath host)
        {
            Range? current = null;
            try { current = host.Range; return current.Start; }
            finally { Release(current); }
        }
        foreach (var host in _hosts.OrderBy(Position))
        {
            Range? range = null;
            try
            {
                _literalPrefixes.TryGetValue(host, out var prefix);
                prefix?.Apply(_document, host);
                range = host.Range;
                var address = new WordFormulaRangeAddress
                { StoryType = range.StoryType, Start = range.Start, End = range.End };
                WordFormulaMutationTransaction.TraceBoundary(_document.Application,
                    $"native-number-update-host start={address.Start} end={address.End}");
                _ = WordNativeOmmlNumbering.RefreshCanonicalNativeFields(host,
                    new WordFormulaHostDescriptor
                    {
                        Kind = WordFormulaHostKind.Omml, DisplayMode = "block", Range = address,
                        Numbering = new WordFormulaNumberingDescriptor
                        {
                            Numbered = true, ContainerKind = WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                            ContainerRange = address,
                        },
                    });
                prefix?.RestoreAliases(_document, host);
            }
            finally { Release(range); }
        }
        // These are preflight-proven references to actual native number aliases,
        // not only VTEqNum_* names and not every unrelated REF in the document.
        foreach (var reference in _references)
        {
            // Native number references can themselves return OMath. Reapplying
            // a body Font snapshot (notably Subscript/Superscript) to that math
            // result is invalid in Word. Let the original REF field and its own
            // formatting switches perform the native result update instead.
            reference.Update();
        }
        WordFormulaMutationTransaction.TraceBoundary(_document.Application, "native-number-plan-after-update");
    }

    internal void VerifyRollback()
    {
        foreach (var recovery in _literalRecovery) recovery.VerifyRollback();
    }

    public void Dispose()
    {
        foreach (var recovery in _literalRecovery) recovery.Dispose();
        _literalRecovery.Clear();
        foreach (var range in _hosts) Release(range);
        _hosts.Clear();
        foreach (var prefix in _literalPrefixes.Values) Release(prefix.PreparedText);
        _literalPrefixes.Clear();
        _literalStaging?.Dispose(); _literalStaging = null;
        foreach (var reference in _references) Release(reference);
        _references.Clear();
        _numberAliases.Clear();
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
