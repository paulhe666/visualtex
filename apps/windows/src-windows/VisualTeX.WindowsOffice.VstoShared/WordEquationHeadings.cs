using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordEquationNumbering;
using static VisualTeX.WordVsto.WordNumberedDisplayLayout;
using static VisualTeX.WordVsto.WordNumberingFields;
using static VisualTeX.WordVsto.WordNumberingState;

namespace VisualTeX.WordVsto;

/// <summary>
/// Headings that scope equation numbers: heading anchors, heading numbers and the heading scope of a position.
/// </summary>
internal static class WordEquationHeadings
{
    internal sealed class HeadingNumberAnchor
    {
        public HeadingNumberAnchor(int position, string numberText)
        {
            Position = position;
            NumberText = numberText;
        }

        public int Position { get; }
        public string NumberText { get; }
    }

    private sealed class HeadingParagraphCandidate
    {
        public HeadingParagraphCandidate(
            int position,
            int outlineLevel,
            string listNumber,
            string textNumber)
        {
            Position = position;
            OutlineLevel = outlineLevel;
            ListNumber = listNumber;
            TextNumber = textNumber;
        }

        public int Position { get; }
        public int OutlineLevel { get; }
        public string ListNumber { get; }
        public string TextNumber { get; }
        public string ExplicitNumber =>
            !string.IsNullOrWhiteSpace(ListNumber) ? ListNumber : TextNumber;
    }

    internal static IReadOnlyList<HeadingNumberAnchor> GetHeadingNumberAnchorsForFormatBatch(
        Document document,
        int targetLevel,
        IReadOnlyList<NativeEquationCaptionEntry> captions)
    {
        var captionPositions = captions
            .Select(caption => caption.Position)
            .ToHashSet();
        var candidates = new List<HeadingParagraphCandidate>();
        Range? content = null;
        try
        {
            content = document.Content;
            ParagraphFormat? aggregateFormat = null;
            try
            {
                aggregateFormat = content.ParagraphFormat;
                var level = (int)aggregateFormat.OutlineLevel;
                // A defined aggregate value applies to every paragraph. If all
                // are body text (or deeper than the requested level), there can
                // be no heading anchor. Mixed values still use native Find.
                if (level > targetLevel && level <= (int)WdOutlineLevel.wdOutlineLevelBodyText)
                    return Array.Empty<HeadingNumberAnchor>();
            }
            finally { Release(aggregateFormat); }
            for (var outlineLevel = 1; outlineLevel <= targetLevel; outlineLevel++)
            {
                var cursor = content.Start;
                while (cursor < content.End)
                {
                    Range? searchRange = null;
                    Find? find = null;
                    ParagraphFormat? findParagraphFormat = null;
                    Paragraphs? foundParagraphs = null;
                    Paragraph? foundParagraph = null;
                    Range? range = null;
                    ListFormat? listFormat = null;
                    Frames? frames = null;
                    try
                    {
                        searchRange = document.Range(cursor, content.End);
                        find = searchRange.Find;
                        find.ClearFormatting();
                        find.Text = "^p";
                        find.MatchWildcards = false;
                        find.MatchCase = false;
                        find.MatchWholeWord = false;
                        find.MatchSoundsLike = false;
                        find.MatchAllWordForms = false;
                        find.Forward = true;
                        find.Wrap = WdFindWrap.wdFindStop;
                        find.Format = true;
                        findParagraphFormat = find.ParagraphFormat;
                        findParagraphFormat.OutlineLevel = (WdOutlineLevel)outlineLevel;
                        if (!find.Execute()) break;

                        // Find narrows searchRange to the matching paragraph mark.
                        // Resolve only that paragraph; Word's native formatter has
                        // already skipped hundreds of formula/table/caption rows.
                        foundParagraphs = searchRange.Paragraphs;
                        if (foundParagraphs.Count == 0)
                        {
                            cursor = Math.Max(cursor + 1, searchRange.End);
                            continue;
                        }
                        foundParagraph = foundParagraphs[1];
                        range = foundParagraph.Range;
                        cursor = Math.Max(cursor + 1, range.End);

                        if (captionPositions.Contains(range.Start)) continue;
                        if (IsWithinTableLayoutFree(range))
                            continue;
                        try
                        {
                            frames = range.Frames;
                            if (frames.Count > 0) continue;
                        }
                        catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                        if (IsVisualTeXStructuralHeadingParagraph(range))
                            continue;

                        var listNumber = string.Empty;
                        try
                        {
                            listFormat = range.ListFormat;
                            listNumber = NormalizeHeadingNumberText(
                                listFormat.ListString,
                                outlineLevel);
                        }
                        catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                        candidates.Add(new HeadingParagraphCandidate(
                            range.Start,
                            outlineLevel,
                            listNumber,
                            ParseHeadingNumberFromText(range.Text, outlineLevel)));
                    }
                    finally
                    {
                        Release(frames);
                        Release(listFormat);
                        Release(range);
                        Release(foundParagraph);
                        Release(foundParagraphs);
                        Release(findParagraphFormat);
                        Release(find);
                        Release(searchRange);
                    }
                }
            }
        }
        catch
        {
            // Formatting-only Find is unavailable on a few compatibility-mode
            // builds. Preserve the existing exhaustive paragraph scan there.
            return GetHeadingNumberAnchorsExhaustive(document, targetLevel);
        }
        finally { Release(content); }

        candidates.Sort((left, right) => left.Position.CompareTo(right.Position));
        var result = new List<HeadingNumberAnchor>();
        var usesExplicitNumbering = candidates.Any(candidate =>
            !string.IsNullOrWhiteSpace(candidate.ExplicitNumber));
        var counters = new int[10];
        foreach (var candidate in candidates)
        {
            var numberText = candidate.ExplicitNumber;
            if (string.IsNullOrWhiteSpace(numberText))
            {
                if (usesExplicitNumbering) continue;
                counters[candidate.OutlineLevel]++;
                for (var deeper = candidate.OutlineLevel + 1;
                     deeper < counters.Length;
                     deeper++)
                    counters[deeper] = 0;
                var parts = new List<string>();
                for (var level = 1; level <= candidate.OutlineLevel; level++)
                    parts.Add(Math.Max(1, counters[level]).ToString());
                numberText = string.Join(".", parts);
            }

            if (candidate.OutlineLevel < targetLevel)
            {
                var missingLevels = targetLevel - candidate.OutlineLevel;
                numberText += string.Concat(
                    Enumerable.Repeat(".0", missingLevels));
            }
            result.Add(new HeadingNumberAnchor(candidate.Position, numberText));
        }
        return result;
    }

    internal static IReadOnlyList<HeadingNumberAnchor> GetHeadingNumberAnchors(
        Document document, int targetLevel) =>
        GetHeadingNumberAnchorsForFormatBatch(document, targetLevel, Array.Empty<NativeEquationCaptionEntry>());

    private static IReadOnlyList<HeadingNumberAnchor> GetHeadingNumberAnchorsExhaustive(
        Document document,
        int targetLevel)
    {
        var candidates = new List<HeadingParagraphCandidate>();
        Paragraphs? paragraphs = null;
        try
        {
            paragraphs = document.Paragraphs;
            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Paragraph? paragraph = null;
                Range? range = null;
                ListFormat? listFormat = null;
                Frames? frames = null;
                try
                {
                    paragraph = paragraphs[index];
                    var outlineLevel = (int)paragraph.OutlineLevel;
                    if (outlineLevel < 1 || outlineLevel > targetLevel) continue;

                    range = paragraph.Range;
                    // Numbering tables and clipped native-caption paragraphs can
                    // inherit Heading 1 from the source formula paragraph. They
                    // are implementation scaffolds, never chapter boundaries.
                    if (IsWithinTableLayoutFree(range))
                        continue;
                    try
                    {
                        frames = range.Frames;
                        if (frames.Count > 0) continue;
                    }
                    catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                    if (IsVisualTeXStructuralHeadingParagraph(range))
                        continue;

                    var listNumber = string.Empty;
                    try
                    {
                        listFormat = range.ListFormat;
                        listNumber = NormalizeHeadingNumberText(
                            listFormat.ListString,
                            outlineLevel);
                    }
                    catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                    candidates.Add(new HeadingParagraphCandidate(
                        range.Start,
                        outlineLevel,
                        listNumber,
                        ParseHeadingNumberFromText(range.Text, outlineLevel)));
                }
                finally
                {
                    Release(frames);
                    Release(listFormat);
                    Release(range);
                    Release(paragraph);
                }
            }
        }
        finally { Release(paragraphs); }

        var result = new List<HeadingNumberAnchor>();
        var usesExplicitNumbering = candidates.Any(candidate =>
            !string.IsNullOrWhiteSpace(candidate.ExplicitNumber));
        var counters = new int[10];
        foreach (var candidate in candidates)
        {
            var numberText = candidate.ExplicitNumber;
            if (string.IsNullOrWhiteSpace(numberText))
            {
                // In a manually/automatically numbered document, an unnumbered
                // Heading 1 is usually the document title, not "chapter 1".
                if (usesExplicitNumbering) continue;
                counters[candidate.OutlineLevel]++;
                for (var deeper = candidate.OutlineLevel + 1;
                     deeper < counters.Length;
                     deeper++)
                    counters[deeper] = 0;
                var parts = new List<string>();
                for (var level = 1; level <= candidate.OutlineLevel; level++)
                    parts.Add(Math.Max(1, counters[level]).ToString());
                numberText = string.Join(".", parts);
            }

            if (candidate.OutlineLevel < targetLevel)
            {
                var missingLevels = targetLevel - candidate.OutlineLevel;
                numberText += string.Concat(
                    Enumerable.Repeat(".0", missingLevels));
            }
            result.Add(new HeadingNumberAnchor(candidate.Position, numberText));
        }
        return result;
    }

    internal static string NormalizeHeadingNumberText(string? value, int outlineLevel)
    {
        var text = (value ?? string.Empty)
            .Replace("\t", string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace("\a", string.Empty)
            .Trim();
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;

        // Word's ListString is presentation text, not a semantic chapter number.
        // Multilevel lists may therefore expose labels such as "第3章",
        // "第3章 第2节" or "Chapter 3 Section 2". Never copy those decorations
        // into an equation-number prefix. Accept only Arabic-number components
        // surrounded by a small, explicit set of list-label decorations and
        // normalize every hierarchy separator to '.'. Any unexpected text fails
        // closed so the ordinary heading-text/fallback counter path can take over.
        var normalized = text
            .Replace('０', '0')
            .Replace('１', '1')
            .Replace('２', '2')
            .Replace('３', '3')
            .Replace('４', '4')
            .Replace('５', '5')
            .Replace('６', '6')
            .Replace('７', '7')
            .Replace('８', '8')
            .Replace('９', '9');
        var numericMatches = Regex.Matches(
            normalized,
            @"\d+",
            RegexOptions.CultureInvariant);
        if (numericMatches.Count == 0
            || numericMatches.Count > Math.Max(1, outlineLevel))
            return string.Empty;

        var remainder = Regex.Replace(
            normalized,
            @"\d+",
            string.Empty,
            RegexOptions.CultureInvariant);
        remainder = Regex.Replace(
            remainder,
            @"(?i:\b(?:chapter|section|part)\b)|[第章节篇部卷]",
            string.Empty,
            RegexOptions.CultureInvariant);
        remainder = Regex.Replace(
            remainder,
            @"[\s.．、,:：;；\-–—()（）\[\]【】{}]+",
            string.Empty,
            RegexOptions.CultureInvariant);
        if (!string.IsNullOrEmpty(remainder)) return string.Empty;

        return string.Join(
            ".",
            numericMatches
                .Cast<Match>()
                .Select(match => match.Value));
    }

    private static bool IsVisualTeXStructuralHeadingParagraph(Range range)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        ParagraphFormat? format = null;
        try
        {
            // A legacy terminal typing paragraph may still carry Heading/list
            // properties even though VisualTeX owns it only as document structure.
            // Never allow that bookmark to become a chapter or section anchor.
            bookmarks = range.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmark);
                bookmark = bookmarks[index];
                var name = bookmark.Name ?? string.Empty;
                if (string.Equals(
                        name,
                        CompactTypingTailBookmarkName,
                        StringComparison.Ordinal)
                    || name.StartsWith(
                        NativeCaptionBookmarkPrefix,
                        StringComparison.Ordinal)
                    || name.StartsWith(
                        NativeDisplayAnchorBookmarkPrefix,
                        StringComparison.Ordinal))
                    return true;
            }

            // Interrupted/legacy cleanup can lose the bookmark while leaving the
            // unmistakable empty 1 pt structural paragraph behind. Excluding only
            // that empty exact-height form avoids hiding any real heading text.
            if (!IsNumberingParagraphAdornment(range.Text)) return false;
            font = range.Font;
            format = range.ParagraphFormat;
            return font.Size > 0f
                && font.Size <= CompactTypingTailFontSizePoints + 0.25f
                && format.LineSpacingRule == WdLineSpacing.wdLineSpaceExactly
                && format.LineSpacing > 0f
                && format.LineSpacing
                    <= CompactTypingTailLineSpacingPoints + 0.25f;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(format);
            Release(font);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static string ParseHeadingNumberFromText(string? value, int outlineLevel)
    {
        var text = (value ?? string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace("\a", string.Empty)
            .Trim();
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;

        // Support common manually typed headings such as "8. 多元微积分",
        // "8-1 小节" and "第 8 章" when Word's ListString is empty.
        var match = Regex.Match(
            text,
            @"^(?:第\s*)?(?<number>\d+(?:\s*[.．\-–—]\s*\d+)*)(?:\s*[章节篇部]|\s*[.．、:：)）\-–—]|\s+)",
            RegexOptions.CultureInvariant);
        if (!match.Success) return string.Empty;
        var number = Regex.Replace(
            match.Groups["number"].Value,
            @"\s*[.．\-–—]\s*",
            ".",
            RegexOptions.CultureInvariant);
        var parts = number
            .Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim())
            .Where(part => part.Length > 0)
            .ToArray();
        if (parts.Length == 0 || parts.Length > Math.Max(1, outlineLevel))
            return string.Empty;
        return string.Join(".", parts);
    }

    internal static IReadOnlyDictionary<int, ResolvedEquationHeadingScope>
        CaptureHeadingScopesAtPositions(
            Document document,
            string? formatId,
            IEnumerable<int> formulaPositions)
    {
        var positions = formulaPositions.Distinct().ToArray();
        var result = new Dictionary<int, ResolvedEquationHeadingScope>();
        if (positions.Length == 0) return result;

        var format = EquationNumberFormat.Resolve(formatId);
        if (!format.UsesHeading)
        {
            foreach (var position in positions)
                result[position] = new ResolvedEquationHeadingScope(
                    int.MinValue,
                    int.MinValue,
                    string.Empty);
            return result;
        }

        // Expensive Word paragraph/outline discovery happens exactly once for the
        // whole conversion batch. ResolveEquationNumberScope itself is pure list
        // lookup, and the paragraph end for a real heading is read once per unique
        // heading rather than once per numbered formula.
        var anchors = GetHeadingNumberAnchors(document, format.HeadingLevel);
        var scopeEnds = new Dictionary<int, int>();
        foreach (var formulaPosition in positions)
        {
            var scope = ResolveEquationNumberScope(formulaPosition, format, anchors);
            var numberText = scope.Prefix;
            if (!string.IsNullOrEmpty(format.Separator)
                && numberText.EndsWith(format.Separator, StringComparison.Ordinal))
                numberText = numberText.Substring(
                    0,
                    numberText.Length - format.Separator.Length);

            if (scope.ScopePosition == int.MinValue)
            {
                result[formulaPosition] = new ResolvedEquationHeadingScope(
                    int.MinValue,
                    int.MinValue,
                    numberText);
                continue;
            }

            if (!scopeEnds.TryGetValue(scope.ScopePosition, out var scopeEnd))
            {
                Range? probe = null;
                Paragraphs? paragraphs = null;
                Paragraph? paragraph = null;
                Range? paragraphRange = null;
                try
                {
                    probe = document.Range(scope.ScopePosition, scope.ScopePosition);
                    paragraphs = probe.Paragraphs;
                    if (paragraphs.Count == 0)
                        scopeEnd = scope.ScopePosition;
                    else
                    {
                        paragraph = paragraphs[1];
                        paragraphRange = paragraph.Range;
                        scopeEnd = paragraphRange.End;
                    }
                    scopeEnds[scope.ScopePosition] = scopeEnd;
                }
                finally
                {
                    Release(paragraphRange);
                    Release(paragraph);
                    Release(paragraphs);
                    Release(probe);
                }
            }

            result[formulaPosition] = new ResolvedEquationHeadingScope(
                scope.ScopePosition,
                scopeEnd,
                numberText);
        }
        return result;
    }

    // For one interactive formula, an explicit nearby heading is authoritative:
    // once any explicit numbering exists, the full resolver ignores unnumbered
    // headings. Resolve the last explicit heading at each relevant outline level
    // with a backward native Find instead of rebuilding all future anchors.
    // When none is explicit, keep the full resolver for its synthetic counters.
    private static bool TryResolveExplicitHeadingScopeAtPosition(
        Document document, int formulaPosition, int targetLevel,
        out (int ScopeStart, int ScopeEnd, string NumberText) scope)
    {
        scope = (int.MinValue, int.MinValue, string.Empty);
        Range? content = null;
        Range? point = null;
        try
        {
            content = document.Content;
            point = content.Duplicate;
            var position = Math.Max(content.Start, Math.Min(formulaPosition, content.End));
            point.SetRange(position, position);
            point.Expand(WdUnits.wdParagraph);
            var searchEnd = Math.Min(content.End, point.End);
            for (var level = 1; level <= targetLevel; level++)
            {
                var end = searchEnd;
                while (end > content.Start)
                {
                    Range? search = null;
                    Find? find = null;
                    ParagraphFormat? format = null;
                    Paragraphs? paragraphs = null;
                    Paragraph? paragraph = null;
                    Range? range = null;
                    ListFormat? list = null;
                    Frames? frames = null;
                    try
                    {
                        search = content.Duplicate;
                        search.SetRange(content.Start, end);
                        find = search.Find;
                        find.ClearFormatting();
                        find.Text = "^p";
                        find.MatchWildcards = false;
                        find.MatchCase = false;
                        find.MatchWholeWord = false;
                        find.MatchSoundsLike = false;
                        find.MatchAllWordForms = false;
                        find.Forward = false;
                        find.Wrap = WdFindWrap.wdFindStop;
                        find.Format = true;
                        format = find.ParagraphFormat;
                        format.OutlineLevel = (WdOutlineLevel)level;
                        if (!find.Execute()) break;
                        paragraphs = search.Paragraphs;
                        if (paragraphs.Count == 0) return false;
                        paragraph = paragraphs[1];
                        range = paragraph.Range;
                        if (range.Start >= end || range.Start < content.Start) return false;
                        end = range.Start;
                        if (range.Start > position) continue;
                        if (scope.ScopeStart != int.MinValue && range.Start <= scope.ScopeStart) break;
                        // The candidate is one exact paragraph: structural table
                        // ownership proves the same exclusion without forcing a
                        // whole-document pagination after native Undo invalidates
                        // Word's layout cache. Do not query wdWithInTable here.
                        if (RangeIsWhollyWithinTable(range)) continue;
                        try { frames = range.Frames; if (frames.Count > 0) continue; }
                        catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
                        if (IsVisualTeXStructuralHeadingParagraph(range)) continue;
                        var number = string.Empty;
                        try
                        {
                            list = range.ListFormat;
                            number = NormalizeHeadingNumberText(list.ListString, level);
                        }
                        catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
                        if (string.IsNullOrWhiteSpace(number)) number = ParseHeadingNumberFromText(range.Text, level);
                        if (string.IsNullOrWhiteSpace(number)) continue;
                        if (level < targetLevel)
                            number += string.Concat(Enumerable.Repeat(".0", targetLevel - level));
                        scope = (range.Start, range.End, number);
                        break;
                    }
                    finally
                    {
                        Release(frames); Release(list); Release(range);
                        Release(paragraph); Release(paragraphs); Release(format); Release(find); Release(search);
                    }
                }
            }
            return scope.ScopeStart != int.MinValue;
        }
        catch (COMException) { return false; }
        finally { Release(point); Release(content); }
    }

    internal static (int ScopeStart, int ScopeEnd, string NumberText) ResolveHeadingScopeAtPosition(
        Document document,
        int formulaPosition,
        string? formatId)
    {
        var format = EquationNumberFormat.Resolve(formatId);
        if (!format.UsesHeading)
            return (int.MinValue, int.MinValue, string.Empty);

        if (TryResolveExplicitHeadingScopeAtPosition(document, formulaPosition, format.HeadingLevel, out var explicitScope))
            return explicitScope;

        var anchors = GetHeadingNumberAnchors(document, format.HeadingLevel);
        var scope = ResolveEquationNumberScope(formulaPosition, format, anchors);
        var numberText = scope.Prefix;
        if (!string.IsNullOrEmpty(format.Separator)
            && numberText.EndsWith(format.Separator, StringComparison.Ordinal))
            numberText = numberText.Substring(0, numberText.Length - format.Separator.Length);

        if (scope.ScopePosition == int.MinValue)
            return (int.MinValue, int.MinValue, numberText);

        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            probe = document.Range(scope.ScopePosition, scope.ScopePosition);
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count == 0)
                return (scope.ScopePosition, scope.ScopePosition, numberText);
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return (scope.ScopePosition, paragraphRange.End, numberText);
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    internal static (int ScopePosition, string Prefix) ResolveEquationNumberScope(
        int formulaPosition,
        EquationNumberFormat format,
        IReadOnlyList<HeadingNumberAnchor> anchors)
    {
        if (!format.UsesHeading) return (0, string.Empty);
        HeadingNumberAnchor? selected = null;
        for (var index = 0; index < anchors.Count; index++)
        {
            var anchor = anchors[index];
            if (anchor.Position > formulaPosition) break;
            selected = anchor;
        }
        var headingText = selected?.NumberText
            ?? string.Join(".", Enumerable.Repeat("0", format.HeadingLevel));
        return (
            selected?.Position ?? int.MinValue,
            headingText + format.Separator);
    }
}
