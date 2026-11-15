using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using static VisualTeX.WordVsto.WordFormulaService;
using static VisualTeX.WordVsto.WordFormulaHostSetup;

namespace VisualTeX.WordVsto;

/// <summary>
/// The font size a formula takes from the text around it: font-size indexes built from WordOpenXML and styles, the
/// source contexts of a LaTeX redraw, and the nearest visible text.
/// </summary>
internal static class WordSourceFontSize
{
    internal static double?[]? TryBuildWordOpenXmlFontSizeIndex(
        Range scope,
        string sourceText,
        out int reconstructedLength,
        out int mismatchIndex)
    {
        reconstructedLength = -1;
        mismatchIndex = -1;
        const string packageNamespace =
            "http://schemas.microsoft.com/office/2006/xmlPackage";
        const string wordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        try
        {
            var package = XDocument.Parse(
                scope.WordOpenXML,
                LoadOptions.PreserveWhitespace);
            XNamespace packageNs = packageNamespace;
            XNamespace word = wordNamespace;
            var documentPart = package
                .Descendants(packageNs + "part")
                .FirstOrDefault(part => string.Equals(
                    (string?)part.Attribute(packageNs + "name"),
                    "/word/document.xml",
                    StringComparison.OrdinalIgnoreCase));
            var body = documentPart?
                .Descendants(word + "body")
                .FirstOrDefault();
            if (body is null) return null;

            var stylesPart = package
                .Descendants(packageNs + "part")
                .FirstOrDefault(part => string.Equals(
                    (string?)part.Attribute(packageNs + "name"),
                    "/word/styles.xml",
                    StringComparison.OrdinalIgnoreCase));
            var stylesRoot = stylesPart?
                .Descendants(word + "styles")
                .FirstOrDefault();
            var documentDefaultFontSize = ReadWordOpenXmlFontSize(
                stylesRoot?
                    .Element(word + "docDefaults")?
                    .Element(word + "rPrDefault")?
                    .Element(word + "rPr"),
                word);
            var styleFontSizes = BuildWordOpenXmlStyleFontSizeIndex(
                stylesRoot,
                word,
                out var defaultParagraphStyleId);

            var reconstructed = new StringBuilder(sourceText.Length + 16);
            var fontSizes = new List<double?>(sourceText.Length + 16);
            double? AppendParagraph(XElement paragraph, string terminator)
            {
                var paragraphProperties = paragraph.Element(word + "pPr");
                var paragraphStyleId =
                    (string?)paragraphProperties?
                        .Element(word + "pStyle")?
                        .Attribute(word + "val")
                    ?? defaultParagraphStyleId;
                var paragraphFontSize = ReadWordOpenXmlFontSize(
                    paragraphProperties?.Element(word + "rPr"),
                    word)
                    ?? ResolveWordOpenXmlStyleFontSize(
                        styleFontSizes,
                        paragraphStyleId)
                    ?? documentDefaultFontSize;
                foreach (var token in paragraph.Descendants())
                {
                    string? tokenText = null;
                    if (token.Name == word + "t"
                        || token.Name == word + "delText")
                    {
                        tokenText = token.Value;
                    }
                    else if (token.Name == word + "tab")
                    {
                        tokenText = "\t";
                    }
                    else if (token.Name == word + "br"
                        || token.Name == word + "cr")
                    {
                        tokenText = "\v";
                    }
                    else if (token.Name == word + "noBreakHyphen")
                    {
                        tokenText = "\u2011";
                    }
                    else if (token.Name == word + "softHyphen")
                    {
                        tokenText = "\u00AD";
                    }
                    if (tokenText is null) continue;

                    var run = token.Ancestors(word + "r").FirstOrDefault();
                    var runProperties = run?.Element(word + "rPr");
                    var runStyleId =
                        (string?)runProperties?
                            .Element(word + "rStyle")?
                            .Attribute(word + "val");
                    var runFontSize = ReadWordOpenXmlFontSize(
                        runProperties,
                        word)
                        ?? ResolveWordOpenXmlStyleFontSize(
                            styleFontSizes,
                            runStyleId)
                        ?? paragraphFontSize;
                    reconstructed.Append(tokenText);
                    for (var index = 0; index < tokenText.Length; index++)
                        fontSizes.Add(runFontSize);
                }
                reconstructed.Append(terminator);
                for (var index = 0; index < terminator.Length; index++)
                    fontSizes.Add(paragraphFontSize);
                return paragraphFontSize;
            }

            // Range.Text ends every table cell and every table row with CR+BEL
            // (one story position each, see BuildWordStoryOffsetIndex). Rebuild
            // that exact sequence so a document with ordinary tables keeps the
            // bulk XML size index instead of falling back to one COM font read
            // per formula. Nested tables or unknown block containers are not
            // modelled: they leave the text unmatched and keep the COM fallback.
            var modelled = true;
            void AppendBlocks(IEnumerable<XElement> blocks)
            {
                foreach (var block in blocks)
                {
                    if (!modelled) return;
                    if (block.Name == word + "p")
                    {
                        AppendParagraph(block, "\r");
                    }
                    else if (block.Name == word + "tbl")
                    {
                        foreach (var row in block.Elements(word + "tr"))
                        {
                            double? lastSize = documentDefaultFontSize;
                            foreach (var cell in row.Elements(word + "tc"))
                            {
                                // Markers such as bookmarks carry no text; any
                                // other container holding paragraphs (nested
                                // table, content control) is not modelled.
                                var children = cell.Elements()
                                    .Where(element => element.Name == word + "p"
                                        || element.Descendants(word + "p").Any())
                                    .ToList();
                                if (children.Count == 0
                                    || children.Any(element => element.Name != word + "p"))
                                {
                                    modelled = false;
                                    return;
                                }
                                var paragraphs = children;
                                for (var index = 0; index < paragraphs.Count; index++)
                                    lastSize = AppendParagraph(
                                        paragraphs[index],
                                        index == paragraphs.Count - 1 ? "\r\a" : "\r");
                            }
                            reconstructed.Append("\r\a");
                            fontSizes.Add(lastSize);
                            fontSizes.Add(lastSize);
                        }
                    }
                    else if (block.Name == word + "sdt")
                    {
                        var content = block.Element(word + "sdtContent");
                        if (content is not null) AppendBlocks(content.Elements());
                    }
                    else if (block.Name == word + "customXml")
                    {
                        AppendBlocks(block.Elements());
                    }
                    else
                    {
                        // Anything else keeps the earlier flat rule: each
                        // contained paragraph ends with one CR (markers such as
                        // sectPr or bookmarks contain none and add nothing).
                        foreach (var paragraph in block.Descendants(word + "p"))
                            AppendParagraph(paragraph, "\r");
                    }
                }
            }
            AppendBlocks(body.Elements());
            if (!modelled) return null;

            var reconstructedText = reconstructed.ToString();
            reconstructedLength = reconstructedText.Length;
            var sharedLength = Math.Min(
                reconstructedText.Length,
                sourceText.Length);
            var firstMismatch = 0;
            while (firstMismatch < sharedLength
                && reconstructedText[firstMismatch] == sourceText[firstMismatch])
                firstMismatch++;
            mismatchIndex = firstMismatch == sharedLength
                && reconstructedText.Length == sourceText.Length
                    ? -1
                    : firstMismatch;
            if (mismatchIndex < 0)
                return fontSizes.ToArray();

            // A Word Open XML fragment can include one mandatory trailing empty
            // paragraph that is outside the requested story range. Accept only that
            // exact, unambiguous difference; every other mismatch falls back to COM.
            if (reconstructedText.Length == sourceText.Length + 1
                && reconstructedText[reconstructedText.Length - 1] == '\r'
                && string.Equals(
                    reconstructedText.Substring(0, sourceText.Length),
                    sourceText,
                    StringComparison.Ordinal))
            {
                fontSizes.RemoveAt(fontSizes.Count - 1);
                return fontSizes.ToArray();
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    private static double? ReadWordOpenXmlFontSize(
        XElement? runProperties,
        XNamespace word)
    {
        if (runProperties is null) return null;
        var value = (string?)runProperties
            .Element(word + "sz")?
            .Attribute(word + "val")
            ?? (string?)runProperties
                .Element(word + "szCs")?
                .Attribute(word + "val");
        if (!double.TryParse(
                value,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out var halfPoints)
            || halfPoints <= 0
            || double.IsNaN(halfPoints)
            || double.IsInfinity(halfPoints))
            return null;
        return halfPoints / 2d;
    }

    private static IReadOnlyDictionary<string, double?>
        BuildWordOpenXmlStyleFontSizeIndex(
            XElement? stylesRoot,
            XNamespace word,
            out string? defaultParagraphStyleId)
    {
        defaultParagraphStyleId = null;
        var definitions = new Dictionary<string, (double? Direct, string? BasedOn)>(
            StringComparer.Ordinal);
        if (stylesRoot is null)
            return new Dictionary<string, double?>(StringComparer.Ordinal);

        foreach (var style in stylesRoot.Elements(word + "style"))
        {
            var styleId = (string?)style.Attribute(word + "styleId");
            if (string.IsNullOrWhiteSpace(styleId))
                continue;
            var styleType = (string?)style.Attribute(word + "type");
            var isDefault = (string?)style.Attribute(word + "default");
            if (defaultParagraphStyleId is null
                && string.Equals(styleType, "paragraph", StringComparison.OrdinalIgnoreCase)
                && (string.Equals(isDefault, "1", StringComparison.Ordinal)
                    || string.Equals(isDefault, "true", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(isDefault, "on", StringComparison.OrdinalIgnoreCase)))
                defaultParagraphStyleId = styleId;

            definitions[styleId] = (
                ReadWordOpenXmlFontSize(style.Element(word + "rPr"), word),
                (string?)style
                    .Element(word + "basedOn")?
                    .Attribute(word + "val"));
        }

        var resolved = new Dictionary<string, double?>(StringComparer.Ordinal);
        double? Resolve(string styleId, HashSet<string> visiting)
        {
            if (resolved.TryGetValue(styleId, out var cached))
                return cached;
            if (!definitions.TryGetValue(styleId, out var definition))
                return null;
            if (!visiting.Add(styleId))
                return null;
            var value = definition.Direct;
            if (!value.HasValue
                && !string.IsNullOrWhiteSpace(definition.BasedOn))
                value = Resolve(definition.BasedOn!, visiting);
            visiting.Remove(styleId);
            resolved[styleId] = value;
            return value;
        }

        foreach (var styleId in definitions.Keys)
            _ = Resolve(styleId, new HashSet<string>(StringComparer.Ordinal));
        return resolved;
    }

    private static double? ResolveWordOpenXmlStyleFontSize(
        IReadOnlyDictionary<string, double?> styleFontSizes,
        string? styleId)
    {
        if (string.IsNullOrWhiteSpace(styleId))
            return null;
        return styleFontSizes.TryGetValue(styleId!, out var value)
            ? value
            : null;
    }

    internal static IReadOnlyDictionary<string, LatexRedrawSourceContext>
        BuildLatexRedrawSourceContexts(
            string sourceText,
            IReadOnlyList<WordLatexRedrawTarget> targets)
    {
        var contexts = new Dictionary<string, LatexRedrawSourceContext>(
            targets.Count,
            StringComparer.Ordinal);
        var formulaMask = new bool[sourceText.Length];
        foreach (var target in targets)
        {
            var start = Math.Max(0, Math.Min(sourceText.Length, target.RelativeStart));
            var end = Math.Max(start, Math.Min(
                sourceText.Length,
                target.RelativeStart + target.SourceLength));
            for (var index = start; index < end; index++)
                formulaMask[index] = true;
        }

        foreach (var target in targets)
        {
            var start = Math.Max(0, Math.Min(sourceText.Length, target.RelativeStart));
            var end = Math.Max(start, Math.Min(
                sourceText.Length,
                target.RelativeStart + target.SourceLength));
            var cellStart = FindSourceCellStart(sourceText, start);
            var cellEnd = FindSourceCellEnd(sourceText, end);
            var paragraphStart = Math.Max(
                cellStart,
                FindSourceParagraphStart(sourceText, start));
            var paragraphEnd = Math.Min(
                cellEnd,
                FindSourceParagraphEnd(sourceText, end));
            var previousInParagraph = FindPreviousVisibleSourcePosition(
                sourceText,
                start - 1,
                paragraphStart,
                formulaMask);
            var nextInParagraph = FindNextVisibleSourcePosition(
                sourceText,
                end,
                paragraphEnd,
                formulaMask);
            var hasVisibleSurroundingText =
                previousInParagraph >= 0 || nextInParagraph >= 0;
            var display = string.Equals(
                target.DisplayMode,
                "block",
                StringComparison.Ordinal);

            int fontContextPosition;
            if (!display || hasVisibleSurroundingText)
            {
                fontContextPosition = previousInParagraph >= 0
                    ? previousInParagraph
                    : nextInParagraph;
            }
            else
            {
                // A display formula that is the only visible content in a Word
                // table cell must inherit from that cell, not from text in a
                // neighboring cell. Word's story text concatenates cells with \a,
                // so an unrestricted previous/next search silently crosses table
                // geometry and can pick an unrelated font/size.
                var previousParagraph = FindPreviousVisibleSourcePosition(
                    sourceText,
                    paragraphStart - 1,
                    cellStart,
                    formulaMask);
                var nextParagraph = FindNextVisibleSourcePosition(
                    sourceText,
                    paragraphEnd,
                    cellEnd,
                    formulaMask);
                fontContextPosition = previousParagraph >= 0
                    ? previousParagraph
                    : nextParagraph;
            }

            contexts[target.Id] = new LatexRedrawSourceContext
            {
                HasVisibleSurroundingText = hasVisibleSurroundingText,
                FontContextRelativePosition = fontContextPosition,
            };
        }
        return contexts;
    }

    private static int FindSourceCellStart(string sourceText, int position)
    {
        for (var index = Math.Min(position, sourceText.Length) - 1;
             index >= 0;
             index--)
        {
            if (sourceText[index] == '\a') return index + 1;
        }
        return 0;
    }

    private static int FindSourceCellEnd(string sourceText, int position)
    {
        for (var index = Math.Max(0, position);
             index < sourceText.Length;
             index++)
        {
            if (sourceText[index] == '\a') return index;
        }
        return sourceText.Length;
    }

    internal static int FindSourceParagraphStart(string sourceText, int position)
    {
        for (var index = Math.Min(position, sourceText.Length) - 1;
             index >= 0;
             index--)
        {
            if (IsSourceParagraphBoundary(sourceText[index]))
                return index + 1;
        }
        return 0;
    }

    internal static int FindSourceParagraphEnd(string sourceText, int position)
    {
        for (var index = Math.Max(0, position);
             index < sourceText.Length;
             index++)
        {
            if (IsSourceParagraphBoundary(sourceText[index]))
                return index;
        }
        return sourceText.Length;
    }

    private static int FindPreviousVisibleSourcePosition(
        string sourceText,
        int startPosition,
        int minimumPosition,
        IReadOnlyList<bool> formulaMask)
    {
        for (var index = Math.Min(startPosition, sourceText.Length - 1);
             index >= Math.Max(0, minimumPosition);
             index--)
        {
            if (char.IsLowSurrogate(sourceText[index])
                && index > minimumPosition
                && char.IsHighSurrogate(sourceText[index - 1]))
                index--;
            if (formulaMask[index]) continue;
            if (char.IsHighSurrogate(sourceText[index])
                && index + 1 < sourceText.Length
                && char.IsLowSurrogate(sourceText[index + 1])
                && formulaMask[index + 1])
                continue;
            if (IsVisibleSourceCharacter(sourceText[index])) return index;
        }
        return -1;
    }

    private static int FindNextVisibleSourcePosition(
        string sourceText,
        int startPosition,
        int maximumPosition,
        IReadOnlyList<bool> formulaMask)
    {
        var maximum = Math.Min(sourceText.Length, maximumPosition);
        for (var index = Math.Max(0, startPosition);
             index < maximum;
             index++)
        {
            if (char.IsLowSurrogate(sourceText[index])
                && index > 0
                && char.IsHighSurrogate(sourceText[index - 1]))
                continue;
            if (formulaMask[index]) continue;
            if (char.IsHighSurrogate(sourceText[index])
                && index + 1 < sourceText.Length
                && char.IsLowSurrogate(sourceText[index + 1])
                && formulaMask[index + 1])
                continue;
            if (IsVisibleSourceCharacter(sourceText[index])) return index;
            if (char.IsHighSurrogate(sourceText[index])
                && index + 1 < maximum
                && char.IsLowSurrogate(sourceText[index + 1]))
                index++;
        }
        return -1;
    }

    private static bool IsSourceParagraphBoundary(char character) =>
        character is '\r' or '\n' or '\v' or '\a';

    private static bool IsVisibleSourceCharacter(char character)
    {
        if (character is '\r' or '\n' or '\t' or '\v' or '\a'
            or '\u0001' or '\u200B' or '\u200C')
            return false;
        return !char.IsWhiteSpace(character);
    }

    private static bool TryResolveSourceFormulaFontSizeFromContext(
        Document document,
        WordLatexRedrawPlan plan,
        WordLatexRedrawTarget target,
        int sourceStart,
        int sourceEnd,
        int contextRelativePosition,
        out float fontSizePt)
    {
        fontSizePt = FormulaFontSize.DefaultPt;
        if (contextRelativePosition < 0
            || contextRelativePosition >= plan.SourceText.Length)
            return false;

        var targetEnd = target.RelativeStart + target.SourceLength;
        int wordPosition;
        if (contextRelativePosition < target.RelativeStart)
        {
            wordPosition = sourceStart - CountWordStoryCharacters(
                plan.SourceText,
                contextRelativePosition,
                target.RelativeStart);
        }
        else if (contextRelativePosition >= targetEnd)
        {
            wordPosition = sourceEnd + CountWordStoryCharacters(
                plan.SourceText,
                targetEnd,
                contextRelativePosition);
        }
        else
        {
            return false;
        }

        if (wordPosition < plan.ScopeStart || wordPosition >= plan.ScopeEnd)
            return false;

        Range? probe = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            probe = document.Range(wordPosition, wordPosition + 1);
            var expectedText = ReadSourceCharacterAt(
                plan.SourceText,
                contextRelativePosition);
            var probeText = probe.Text ?? string.Empty;
            if (!string.Equals(
                    probeText,
                    expectedText,
                    StringComparison.Ordinal)
                || !ContainsVisibleBodyText(probeText))
                return false;
            font = probe.Font;
            return TryResolveWordFontSize(font.Size, out fontSizePt);
        }
        catch (COMException)
        {
            return false;
        }
        finally
        {
            Release(font);
            Release(probe);
        }
    }

    internal static int[] BuildWordStoryOffsetIndex(string sourceText)
    {
        var offsets = new int[sourceText.Length + 1];
        var sourceIndex = 0;
        var wordOffset = 0;
        while (sourceIndex < sourceText.Length)
        {
            offsets[sourceIndex] = wordOffset;
            if (char.IsHighSurrogate(sourceText[sourceIndex])
                && sourceIndex + 1 < sourceText.Length
                && char.IsLowSurrogate(sourceText[sourceIndex + 1]))
            {
                // Word counts a valid UTF-16 surrogate pair as one story
                // character. The intermediate boundary is not a valid formula
                // boundary, but mapping it to the pair start keeps the table total.
                offsets[sourceIndex + 1] = wordOffset;
                sourceIndex += 2;
                wordOffset++;
                offsets[sourceIndex] = wordOffset;
                continue;
            }
            if (sourceText[sourceIndex] == '\r'
                && sourceIndex + 1 < sourceText.Length
                && sourceText[sourceIndex + 1] == '\a')
            {
                // Range.Text serializes every Word table cell/row terminator as
                // CR + BEL (\r\a), but Word's story coordinates count that pair as
                // one structural character. Treating both UTF-16 characters as two
                // story positions shifts every formula after the first table cell
                // and eventually sends relocation into a cross-table Word.Find.
                offsets[sourceIndex + 1] = wordOffset;
                sourceIndex += 2;
                wordOffset++;
                offsets[sourceIndex] = wordOffset;
                continue;
            }
            sourceIndex++;
            wordOffset++;
            offsets[sourceIndex] = wordOffset;
        }
        return offsets;
    }

    private static int CountWordStoryCharacters(
        string sourceText,
        int startPosition,
        int endPosition)
    {
        var start = Math.Max(0, Math.Min(sourceText.Length, startPosition));
        var end = Math.Max(start, Math.Min(sourceText.Length, endPosition));
        var count = 0;
        for (var index = start; index < end; index++, count++)
        {
            if (char.IsHighSurrogate(sourceText[index])
                && index + 1 < end
                && char.IsLowSurrogate(sourceText[index + 1]))
            {
                index++;
                continue;
            }
            if (sourceText[index] == '\r'
                && index + 1 < end
                && sourceText[index + 1] == '\a')
                index++;
        }
        return count;
    }

    private static string ReadSourceCharacterAt(string sourceText, int position)
    {
        if (position < 0 || position >= sourceText.Length) return string.Empty;
        if (char.IsHighSurrogate(sourceText[position])
            && position + 1 < sourceText.Length
            && char.IsLowSurrogate(sourceText[position + 1]))
            return sourceText.Substring(position, 2);
        return sourceText[position].ToString();
    }

    internal static float ResolveSourceFormulaFontSize(
        Document document,
        Range source,
        bool display)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            paragraphs = source.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
            }

            var paragraphStart = paragraphRange?.Start ?? source.Start;
            var paragraphBodyEnd = Math.Max(
                paragraphStart,
                (paragraphRange?.End ?? source.End) - 1);

            // Inline formulas should inherit the prose beside them. Generated
            // documents often deliberately make the raw $...$ source smaller than
            // the surrounding text, so the source run itself is only a fallback.
            if (!display || HasVisibleSurroundingText(source))
            {
                if (TryResolveNearbyVisibleFontSize(
                        document,
                        source.Start - 1,
                        out var previousInline,
                        minimumPosition: paragraphStart,
                        step: -1))
                    return previousInline;
                if (TryResolveNearbyVisibleFontSize(
                        document,
                        source.End,
                        out var nextInline,
                        maximumPosition: paragraphBodyEnd,
                        step: 1))
                    return nextInline;
            }

            // A display formula normally occupies its own paragraph. In that case
            // inherit from the nearest visible prose outside the source paragraph,
            // preferring the preceding paragraph as Word users generally expect.
            if (TryResolveNearbyVisibleFontSize(
                    document,
                    paragraphStart - 1,
                    out var previousParagraph,
                    minimumPosition: 0,
                    step: -1))
                return previousParagraph;
            if (TryResolveNearbyVisibleFontSize(
                    document,
                    paragraphRange?.End ?? source.End,
                    out var nextParagraph,
                    maximumPosition: int.MaxValue,
                    step: 1))
                return nextParagraph;

            font = source.Font;
            if (TryResolveWordFontSize(font.Size, out var direct)) return direct;
            Release(font);
            font = null;

            if (paragraphRange is not null)
            {
                if (paragraphRange.End > paragraphRange.Start)
                    paragraphRange.End -= 1;
                font = paragraphRange.Font;
                if (TryResolveWordFontSize(font.Size, out var paragraphSize))
                    return paragraphSize;
            }
            return FormulaFontSize.DefaultPt;
        }
        finally
        {
            Release(font);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool TryResolveNearbyVisibleFontSize(
        Document document,
        int startPosition,
        out float fontSizePt,
        int minimumPosition = 0,
        int maximumPosition = int.MaxValue,
        int step = 1)
    {
        fontSizePt = FormulaFontSize.DefaultPt;
        if (step is not (-1 or 1)) return false;

        Range? content = null;
        try
        {
            content = document.Content;
            var contentStart = content.Start;
            var contentEnd = Math.Max(contentStart, content.End - 1);
            var lowerBound = Math.Max(contentStart, minimumPosition);
            var upperBound = Math.Min(contentEnd, maximumPosition);
            if (upperBound < lowerBound
                || (step < 0 && startPosition < lowerBound)
                || (step > 0 && startPosition > upperBound))
                return false;
            var position = startPosition;
            const int maximumProbeCharacters = 256;
            for (var probeIndex = 0;
                 probeIndex < maximumProbeCharacters
                 && position >= lowerBound
                 && position <= upperBound;
                 probeIndex++, position += step)
            {
                Range? probe = null;
                Microsoft.Office.Interop.Word.Font? font = null;
                try
                {
                    probe = document.Range(position, Math.Min(position + 1, content.End));
                    if (!ContainsVisibleBodyText(probe.Text)) continue;
                    font = probe.Font;
                    if (TryResolveWordFontSize(font.Size, out fontSizePt))
                        return true;
                }
                catch (COMException ignored)
                {
                    // Keep probing neighboring Word story coordinates.
                    IgnoredErrors.Ignore(ignored);
                }
                finally
                {
                    Release(font);
                    Release(probe);
                }
            }
            return false;
        }
        finally { Release(content); }
    }
}
