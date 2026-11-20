using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
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
    // The bookmark each REF found from the REF side (CaptureTailReferences) points at: a REF whose bookmark is gone when
    // the plan is applied (its equation was removed or replaced) is left as it is, Word would write an error into it.
    private readonly Dictionary<Field, string> _tailReferenceNames = new();
    private int _insertedCount;
    private int _earliestChange;
    // >= 0: the equations from this position on are renumbered in one Range.Fields.Update (see Prepare), not one by one.
    private int _bulkFrom = -1;
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
            var scope = WordEquationHeadings.ResolveHeadingScopeAtPosition(document, exact.Start, formatId);
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

    private static bool RangeHasAnyEquation(Document document, int start, int end)
    {
        Range? range = null;
        OMaths? maths = null;
        try
        {
            range = document.Range(start, end);
            maths = range.OMaths;
            return maths.Count > 0;
        }
        catch
        {
            // Unable to tell: keep the exact search.
            return true;
        }
        finally
        {
            Release(maths);
            Release(range);
        }
    }

    // Start of the paragraph holding the first equation in [start, end), or end
    // when there is none. Unable to tell: start, i.e. the exact search goes on.
    private static int NextEquationParagraphStart(Document document, int start, int end)
    {
        if (start >= end) return end;
        Range? range = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? exact = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? owner = null;
        try
        {
            range = document.Range(start, end);
            maths = range.OMaths;
            if (maths.Count == 0) return end;
            math = maths[1];
            exact = math.Range;
            paragraphs = exact.Paragraphs;
            if (paragraphs.Count == 0) return start;
            paragraph = paragraphs[1];
            owner = paragraph.Range;
            return Math.Min(end, Math.Max(start, owner.Start));
        }
        catch
        {
            return start;
        }
        finally
        {
            Release(owner);
            Release(paragraph);
            Release(paragraphs);
            Release(exact);
            Release(math);
            Release(maths);
            Release(range);
        }
    }

    internal static WordNativeOmmlFieldRefresh Prepare(Document document, int earliestChange, bool refreshLiteralPrefixes = false,
        IReadOnlyList<OMath>? additionalNativeHosts = null, ISet<int>? externallyRewrittenNumberStarts = null)
    {
        var plan = new WordNativeOmmlFieldRefresh(document) { _earliestChange = earliestChange };
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
            // Number aliases (bookmarks inside a number) matter only to the REF
            // fields that display them; reading them costs one more WordOpenXML
            // export per equation. A document without any REF needs none.
            bool captureAliases;
            using (WordOperationMetrics.Measure("NativeNumbering.RefreshPrepare.References"))
                captureAliases = WordNumberingFieldCollection.MayHaveReferenceFields(document);
            // Only equations can be native numbers. When no equation follows the
            // change, skip the SEQ search: in a document numbered by VisualTeX
            // OLE every following PlaceRef SEQ would otherwise be one more hit.
            if (cursor < end)
                using (WordOperationMetrics.Measure("NativeNumbering.RefreshPrepare.AnyEquation"))
                    if (!RangeHasAnyEquation(document, cursor, end))
                        cursor = end;
            // No literal chapter prefix to rewrite, no tracked change, not WPS: the equations behind the change are
            // not visited one by one (22 ms each, and 30 ms more for the bookmarks of one that is referenced). Apply
            // renumbers them all in one Range.Fields.Update, which gives the same numbers in about 150 ms (measured on
            // 100 equations). The equations that need a check or a write of their own (_hosts) are still handled
            // singly; the REF fields that point into the equations behind the change are found from the REF fields.
            if (cursor < end && !OfficeHostEnvironment.IsWps && !refreshLiteralPrefixes
                && !document.TrackRevisions && !WordNumberingFieldCollection.HasRevisions(document))
            {
                plan._bulkFrom = cursor;
                if (captureAliases)
                    using (WordOperationMetrics.Measure("NativeNumbering.RefreshPrepare.TailReferences"))
                        plan.CaptureTailReferences(document, cursor);
                cursor = end;
            }
            // The equations behind the change are visited by index in one window (a Find for "SEQ" costs 40 ms a
            // numbered equation in a document of equations), unless the document holds very many of them.
            var byIndex = false;
            if (!OfficeHostEnvironment.IsWps && cursor < end)
            {
                Range? tail = null;
                try
                {
                    tail = document.Range(cursor, end);
                    byIndex = WordNumberingFieldCollection.EquationCountAtMost(tail, WordNativeNumberingAdoptionPlan.EquationVisitLimit);
                }
                finally { Release(tail); }
            }
            while (cursor < end)
            {
                Release(maths); maths = null;
                Release(paragraph); paragraph = null;
                Release(find); find = null;
                Release(retrieval); retrieval = null;
                Release(search); search = null;
                if (OfficeHostEnvironment.IsWps)
                {
                    paragraph = WordNumberingFieldCollection.NextEquationUnit(scope, cursor, end);
                    if (paragraph is null) break;
                }
                else if (byIndex)
                {
                    paragraph = document.Range(cursor, end);
                }
                else
                {
                    search = scope.Duplicate;
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
                }
                if (paragraph.End <= cursor || paragraph.Start >= end)
                    throw new InvalidDataException("Native-number search escaped or repeated its bounded suffix.");
                // Inspect every OMath in this found paragraph before moving past
                // it; restarting inside a field can snap Find to that same field.
                maths = paragraph.OMaths;
                if (maths.Count == 0)
                {
                    // A SEQ paragraph without an equation (a VisualTeX OLE number,
                    // a caption) holds no native number. Resume at the paragraph of
                    // the next equation rather than at the next SEQ: in a document
                    // numbered by VisualTeX OLE every PlaceRef SEQ was one more hit.
                    // Every paragraph skipped this way has no equation either.
                    cursor = NextEquationParagraphStart(document, paragraph.End, end);
                    continue;
                }
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
                        if (captureAliases) plan.CaptureNumberAliases(exact);
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
                        if (!exact.InStory(scope) || !WordOmmlEquationNumbers.HasWordNativeEquationNumberHost(
                                WordOmmlConverter.ReadLocalEquationXml(exact)))
                            throw new InvalidDataException("A prepared native number lost its structural ownership.");
                        owners = exact.OMaths;
                        if (owners.Count != 1) throw new InvalidDataException("A prepared native number no longer owns one OMath.");
                        owner = owners[1];
                        if (captureAliases) plan.CaptureNumberAliases(exact);
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

    // The names of the bookmarks in an equation. Word: read from the range's own bookmarks, hidden ones (_Ref...)
    // included, which costs a COM call or two; an export of the equation's XML costs about 30 ms. WPS: the XML.
    private static IReadOnlyList<string> BookmarkNamesIn(Range exact)
    {
        if (OfficeHostEnvironment.IsWps) return MathTypeLocalFragment.BookmarkNames(exact.WordOpenXML);
        Bookmarks? inside = null;
        var hiddenShown = false;
        var previous = false;
        try
        {
            inside = exact.Bookmarks;
            // (hidden bookmarks, the _Ref ones, are listed only while ShowHidden is on; put it back as it was)
            previous = inside.ShowHidden;
            inside.ShowHidden = true;
            hiddenShown = true;
            var count = inside.Count;
            var names = new List<string>(count);
            for (var index = 1; index <= count; index++)
            {
                Bookmark? mark = null;
                try { mark = inside[index]; names.Add(mark.Name); }
                finally { Release(mark); }
            }
            return names;
        }
        finally
        {
            if (hiddenShown && inside is not null)
            {
                try { inside.ShowHidden = previous; }
                catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
            }
            Release(inside);
        }
    }

    private void CaptureNumberAliases(Range exact)
    {
        Range? number = null;
        Bookmarks? bookmarks = null;
        try
        {
            var names = BookmarkNamesIn(exact);
            if (names.Count == 0) return;
            number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
            bookmarks = _document.Bookmarks;
            foreach (var name in names)
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

    // The REF fields whose bookmark lies in an equation from `from` on: the references to the numbers that the
    // renumbering changes (a REF before the change can point to an equation behind it).
    private void CaptureTailReferences(Document document, int from)
    {
        var references = WordNumberingFieldCollection.CaptureReferenceFields(document);
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var field in references)
            {
                Range? code = null;
                Bookmark? bookmark = null;
                Range? target = null;
                OMaths? maths = null;
                try
                {
                    code = field.Code;
                    if (!WordEquationReferenceFields.TryReadReferenceBookmark(code.Text, out var name)
                        || !bookmarks.Exists(name)) continue;
                    bookmark = bookmarks[name];
                    target = bookmark.Range;
                    if (target.Start < from) continue;
                    maths = target.OMaths;
                    if (maths.Count == 0) continue;
                    _references.Add(field);
                    _tailReferenceNames[field] = name;
                }
                finally { Release(maths); Release(target); Release(bookmark); Release(code); }
            }
        }
        finally
        {
            // A field kept in _references is owned by the plan; every other one is released here.
            foreach (var field in references)
                if (!_references.Contains(field)) Release(field);
            Release(bookmarks);
        }
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
        // Without any tracked change no field can be revision-deleted; skip the
        // two Revisions queries per field (about 26 ms per equation).
        var hasRevisions = WordNumberingFieldCollection.HasRevisions(_document);
        if (OfficeHostEnvironment.IsWps)
        {
            ApplyInWps(hasRevisions);
            return;
        }
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
                    },
                    hasRevisions: hasRevisions,
                    // the bulk pass below updates these fields too (the check of the structure stays)
                    updateFields: _bulkFrom < 0 || hasRevisions);
                prefix?.RestoreAliases(_document, host);
            }
            finally { Release(range); }
        }
        // The equations behind the change (not visited singly, see Prepare): every number field from the change on in
        // one Range.Fields.Update, checked on samples; when Word left a number stale, field by field.
        if (_bulkFrom >= 0)
        {
            using (WordOperationMetrics.Measure("NativeNumbering.RefreshApply.Bulk"))
            {
                var updated = hasRevisions ? -1 : WordNumberingFieldCollection.BulkUpdateFrom(_document, _bulkFrom, nativeOnly: true);
                if (updated < 0)
                    using (WordOperationMetrics.Measure("NativeNumbering.RefreshApply.BulkFallback"))
                        UpdateTailFieldByField(_bulkFrom);
            }
        }
        // These are preflight-proven references to actual native number aliases,
        // not only VTEqNum_* names and not every unrelated REF in the document.
        foreach (var reference in _references)
        {
            if (_tailReferenceNames.TryGetValue(reference, out var bookmarkName) && !BookmarkExists(bookmarkName)) continue;
            // Native number references can themselves return OMath. Reapplying
            // a body Font snapshot (notably Subscript/Superscript) to that math
            // result is invalid in Word. Let the original REF field and its own
            // formatting switches perform the native result update instead.
            reference.Update();
            WordEquationReferenceFields.FlattenWpsMathResult(reference);
        }
        WordFormulaMutationTransaction.TraceBoundary(_document.Application, "native-number-plan-after-update");
    }

    private bool BookmarkExists(string name)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = _document.Bookmarks;
            return bookmarks.Exists(name);
        }
        finally { Release(bookmarks); }
    }

    // The field-by-field counterpart of the bulk pass: the STYLEREF and SEQ fields of each display equation from
    // the position on, equation by equation, the chapter state before the SEQ that reads it.
    private void UpdateTailFieldByField(int from)
    {
        Range? content = null;
        Range? tail = null;
        OMaths? maths = null;
        try
        {
            content = _document.Content;
            tail = _document.Range(Math.Max(content.Start, Math.Min(from, content.End - 1)), content.End);
            maths = tail.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                Fields? fields = null;
                try
                {
                    math = maths[index];
                    if (math.Type != WdOMathType.wdOMathDisplay) continue;
                    range = math.Range;
                    fields = range.Fields;
                    UpdateFieldsOfType(fields, WdFieldType.wdFieldStyleRef);
                    UpdateFieldsOfType(fields, WdFieldType.wdFieldSequence);
                }
                finally { Release(fields); Release(range); Release(math); }
            }
        }
        finally { Release(maths); Release(tail); Release(content); }
    }

    private static void UpdateFieldsOfType(Fields fields, WdFieldType type)
    {
        for (var index = 1; index <= fields.Count; index++)
        {
            Field? field = null;
            try
            {
                field = fields[index];
                if (field.Type == type) field.Update();
            }
            finally { Release(field); }
        }
    }

    // A WPS OMath is a tracked range, not the equation: one taken before a write at its
    // start grows over what is written there (the equation inserted before it). The
    // numbered equations are found again after the write, in document order, and the
    // references to their numbers are found from the REF fields, because WPS's Find does
    // not search field codes. Literal chapter prefixes (number refresh only, no write
    // before the hosts) keep the hosts taken in preflight.
    private void ApplyInWps(bool hasRevisions)
    {
        WordFormulaMutationTransaction.TraceBoundary(_document.Application, "native-number-plan-before-update");
        foreach (var item in _literalPrefixes)
            item.Value.Apply(_document, item.Key);
        var numbers = new List<(int Start, int End)>();
        Range? scope = null;
        try
        {
            scope = _document.Content;
            var cursor = Math.Max(scope.Start, _earliestChange);
            while (cursor < scope.End)
            {
                Range? unit = null;
                OMaths? maths = null;
                try
                {
                    unit = WordNumberingFieldCollection.NextEquationUnit(scope, cursor, scope.End);
                    if (unit is null) break;
                    maths = unit.OMaths;
                    for (var index = 1; index <= maths.Count; index++)
                    {
                        OMath? math = null;
                        Range? exact = null;
                        try
                        {
                            math = maths[index];
                            if (math.Type != WdOMathType.wdOMathDisplay) continue;
                            exact = math.Range.Duplicate;
                            if (exact.Start < cursor
                                || !WordNativeOmmlNumbering.TryResolveUnownedNativeStructure(_document, exact, out _))
                                continue;
                            var address = new WordFormulaRangeAddress
                            { StoryType = exact.StoryType, Start = exact.Start, End = exact.End };
                            WordFormulaMutationTransaction.TraceBoundary(_document.Application,
                                $"native-number-update-host start={address.Start} end={address.End}");
                            var refreshed = WordNativeOmmlNumbering.RefreshCanonicalNativeFields(math,
                                new WordFormulaHostDescriptor
                                {
                                    Kind = WordFormulaHostKind.Omml, DisplayMode = "block", Range = address,
                                    Numbering = new WordFormulaNumberingDescriptor
                                    {
                                        Numbered = true, ContainerKind = WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                                        ContainerRange = address,
                                    },
                                },
                                hasRevisions: hasRevisions);
                            var number = refreshed.Numbering.NumberRange!;
                            numbers.Add((number.Start, number.End));
                        }
                        finally { Release(exact); Release(math); }
                    }
                    // The updates may have changed the paragraph's length.
                    if (unit.End <= cursor)
                        throw new InvalidDataException("Native-number refresh did not advance past its paragraph.");
                    cursor = unit.End;
                }
                finally { Release(maths); Release(unit); }
            }
        }
        finally { Release(scope); }
        foreach (var item in _literalPrefixes)
            item.Value.RestoreAliases(_document, item.Key);
        UpdateNumberReferencesInWps(numbers);
        WordFormulaMutationTransaction.TraceBoundary(_document.Application, "native-number-plan-after-update");
    }

    // Every REF whose bookmark lies inside one of the refreshed numbers.
    private void UpdateNumberReferencesInWps(IReadOnlyList<(int Start, int End)> numbers)
    {
        if (numbers.Count == 0) return;
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        try
        {
            fields = _document.Fields;
            bookmarks = _document.Bookmarks;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Bookmark? bookmark = null;
                Range? target = null;
                try
                {
                    field = fields[index];
                    if (field.Type != WdFieldType.wdFieldRef) continue;
                    code = field.Code;
                    if (!WordEquationReferenceFields.TryReadReferenceBookmark(code.Text ?? string.Empty, out var name)
                        || !bookmarks.Exists(name))
                        continue;
                    bookmark = bookmarks[name];
                    target = bookmark.Range;
                    if (target.StoryType != WdStoryType.wdMainTextStory
                        || !numbers.Any(number => target.Start >= number.Start && target.End <= number.End))
                        continue;
                    field.Update();
                    WordEquationReferenceFields.FlattenWpsMathResult(field);
                }
                finally { Release(target); Release(bookmark); Release(code); Release(field); }
            }
        }
        finally { Release(bookmarks); Release(fields); }
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
        _tailReferenceNames.Clear();
        _numberAliases.Clear();
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
