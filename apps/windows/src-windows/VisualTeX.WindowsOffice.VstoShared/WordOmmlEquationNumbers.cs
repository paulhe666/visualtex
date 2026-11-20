using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using static VisualTeX.WordVsto.WordOmmlConverter;

namespace VisualTeX.WordVsto;

/// <summary>
/// Equation numbers inside OMML: building numbered OMML (VisualTeX's hash sequence, Word's native form), and
/// reading, stripping and replacing the number.
/// </summary>
internal static class WordOmmlEquationNumbers
{
    private const string VisualTeXEquationNumberPlaceholderPrefix = "981730";
    private const string VisualTeXNativeNumberBookmarkPrefix = "VTEqNum_";

    internal static string BuildImmutableHashSequenceNumberedOmml(
        string semanticOmml,
        string sequenceName,
        string numberBookmarkName,
        string visibleBookmarkName,
        string captionBookmarkName,
        string prefix,
        int restartHeadingLevel,
        string initialSequenceResult)
    {
        if (string.IsNullOrWhiteSpace(semanticOmml))
            throw new ArgumentException(
                "The semantic OMML payload must not be empty.",
                nameof(semanticOmml));
        sequenceName = (sequenceName ?? string.Empty).Trim();
        if (sequenceName.Length == 0
            || sequenceName.Length > 80
            || sequenceName.Any(character =>
                character is '\r' or '\n' or '\u0013' or '\u0014' or '\u0015' or '"'))
            throw new ArgumentException(
                "The Word SEQ identifier is invalid.",
                nameof(sequenceName));
        var sequenceToken = Regex.IsMatch(
                sequenceName,
                @"^[A-Za-z][A-Za-z0-9_]{0,39}$",
                RegexOptions.CultureInvariant)
            ? sequenceName
            : """ + sequenceName + """;
        if (restartHeadingLevel < 0 || restartHeadingLevel > 9)
            throw new ArgumentOutOfRangeException(
                nameof(restartHeadingLevel),
                "A Word heading reset level must be between 0 and 9.");
        if (string.IsNullOrWhiteSpace(initialSequenceResult)
            || initialSequenceResult.Any(character =>
                character is '\r' or '\n' or '\u0013' or '\u0014' or '\u0015'))
            throw new ArgumentException(
                "The initial SEQ result is invalid.",
                nameof(initialSequenceResult));
        prefix ??= string.Empty;
        if (prefix.Any(character =>
                character is '\r' or '\n' or '\u0013' or '\u0014' or '\u0015'))
            throw new ArgumentException(
                "The equation-number prefix contains an invalid Word field-control character.",
                nameof(prefix));

        var normalizedFormulaId = ValidateManagedEquationBookmarkName(
            numberBookmarkName,
            VisualTeXNativeNumberBookmarkPrefix,
            nameof(numberBookmarkName));
        if (!string.Equals(
                normalizedFormulaId,
                ValidateManagedEquationBookmarkName(
                    visibleBookmarkName,
                    "VTEq_",
                    nameof(visibleBookmarkName)),
                StringComparison.OrdinalIgnoreCase)
            || !string.Equals(
                normalizedFormulaId,
                ValidateManagedEquationBookmarkName(
                    captionBookmarkName,
                    "VTEqCap_",
                    nameof(captionBookmarkName)),
                StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(
                "All VisualTeX equation-number bookmarks must carry the same FormulaId.");

        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;
        var equation = XElement.Parse(
            ExtractSingleOMath(semanticOmml),
            LoadOptions.PreserveWhitespace);
        var formulaNodes = equation
            .Elements()
            .Select(element => new XElement(element))
            .Cast<object>()
            .ToList();
        if (formulaNodes.Count == 0)
            throw new InvalidDataException(
                "A numbered OMML display equation must contain a nonempty mathematical body.");

        var representativeRunProperties = equation
            .Descendants(word + "rPr")
            .FirstOrDefault(properties => properties.Element(word + "sz") is not null)
            ?? equation.Descendants(word + "rPr").FirstOrDefault();
        XElement WrapperRunProperties(bool noProof = false, bool italic = false)
        {
            var properties = new XElement(word + "rPr");
            foreach (var name in new[] { word + "rFonts", word + "sz", word + "szCs" })
            {
                var property = representativeRunProperties?.Element(name);
                if (property is not null)
                    properties.Add(new XElement(property));
            }
            if (italic)
                properties.Add(new XElement(word + "i"));
            if (noProof)
                properties.Add(new XElement(word + "noProof"));
            return properties;
        }
        XElement WrapperControlProperties() =>
            new(
                math + "ctrlPr",
                // Match Word's own professional #({SEQ ...}) serialization. The
                // control run is italic, while the visible numeric result is not.
                // Marking field-control runs with m:nor makes Word treat the hidden
                // instruction as ordinary mathematical content and expands the
                // equation-number slot leftward instead of keeping it at the right
                // margin.
                WrapperRunProperties(italic: true));
        XElement FieldBoundaryRun(XElement content) =>
            new(
                math + "r",
                WrapperRunProperties(italic: true),
                content);
        XElement FieldInstructionRun(XElement content) =>
            new(
                math + "r",
                WrapperRunProperties(),
                content);
        XElement FieldResultRun(XElement content) =>
            new(
                math + "r",
                WrapperRunProperties(noProof: true),
                content);
        XElement TextRun(string text) =>
            new(
                math + "r",
                WrapperRunProperties(),
                new XElement(math + "t", text));
        XElement HashSeparatorRun() =>
            // Keep the hash itself Word-canonical. Word is allowed to merge this run
            // into the preceding simple formula run during normalization; the native
            // number parser therefore recognizes both a standalone '#' token and a
            // formula-tail run whose final character is '#'. Adding m:nor here makes
            // Word treat the separator as ordinary math and destroys right-margin
            // #() label geometry.
            TextRun("#");

        var sequenceSwitch = restartHeadingLevel > 0
            ? $" \\s {restartHeadingLevel}"
            : string.Empty;
        var sequenceInstruction =
            // Word itself adds MERGEFORMAT when a SEQ field is inserted with
            // PreserveFormatting=true before professional BuildUp. Keep the same
            // canonical field instruction so the #() number remains a right-margin
            // label after save/reopen instead of behaving like ordinary math text.
            $" SEQ {sequenceName}{sequenceSwitch} \\* ARABIC \\* MERGEFORMAT ";
        var fieldElements = new List<object>
        {
            // Keep the number alias outermost: VTEqNum must cover the complete
            // chapter prefix plus the live SEQ result read by ordinary body REF.
            new XElement(
                word + "bookmarkStart",
                new XAttribute(word + "id", "31801"),
                new XAttribute(word + "name", numberBookmarkName)),
            new XElement(
                word + "bookmarkStart",
                new XAttribute(word + "id", "31802"),
                new XAttribute(word + "name", visibleBookmarkName)),
            new XElement(
                word + "bookmarkStart",
                new XAttribute(word + "id", "31803"),
                new XAttribute(word + "name", captionBookmarkName)),
        };
        if (!string.IsNullOrEmpty(prefix))
            fieldElements.Add(TextRun(prefix));
        fieldElements.Add(FieldBoundaryRun(new XElement(
            word + "fldChar",
            // Word's own Ctrl+F9/BuildUp #(SEQ) field is not imported as dirty.
            // Setting w:dirty=true on an XML-inserted field makes interactive Word
            // show the "fields may refer to other files" update-confirmation dialog
            // even in a brand-new document. VisualTeX explicitly updates this SEQ
            // after insertion, so the dirty flag is both unnecessary and harmful.
            new XAttribute(word + "fldCharType", "begin"))));
        fieldElements.Add(FieldInstructionRun(new XElement(
            word + "instrText",
            new XAttribute(XNamespace.Xml + "space", "preserve"),
            sequenceInstruction)));
        fieldElements.Add(FieldBoundaryRun(new XElement(
            word + "fldChar",
            new XAttribute(word + "fldCharType", "separate"))));
        fieldElements.Add(FieldResultRun(new XElement(math + "t", initialSequenceResult)));
        fieldElements.Add(FieldBoundaryRun(new XElement(
            word + "fldChar",
            new XAttribute(word + "fldCharType", "end"))));
        fieldElements.Add(new XElement(
            word + "bookmarkEnd",
            new XAttribute(word + "id", "31803")));
        fieldElements.Add(new XElement(
            word + "bookmarkEnd",
            new XAttribute(word + "id", "31802")));
        fieldElements.Add(new XElement(
            word + "bookmarkEnd",
            new XAttribute(word + "id", "31801")));

        var delimiter = new XElement(
            math + "d",
            new XElement(
                math + "dPr",
                WrapperControlProperties()),
            new XElement(math + "e", fieldElements));
        var equationBody = new XElement(math + "e", formulaNodes);
        equationBody.Add(
            HashSeparatorRun(),
            delimiter);
        return new XElement(
                math + "oMath",
                new XElement(
                    math + "eqArr",
                    new XElement(
                        math + "eqArrPr",
                        new XElement(
                            math + "maxDist",
                            new XAttribute(math + "val", "1")),
                        WrapperControlProperties()),
                    equationBody))
            .ToString(SaveOptions.DisableFormatting);
    }

    internal static string BuildWordNativeNumberedOmml(
        string semanticOmml,
        string sequenceName,
        string? numberBookmarkName,
        int restartHeadingLevel,
        string headingSeparator,
        string initialSequenceResult = "1",
        string? literalHeadingPrefix = null,
        int? restartSequenceValue = null)
    {
        if (string.IsNullOrWhiteSpace(semanticOmml))
            throw new ArgumentException(
                "The semantic OMML payload must not be empty.",
                nameof(semanticOmml));
        if (string.IsNullOrWhiteSpace(sequenceName)
            || sequenceName.Length > 80
            || sequenceName.Any(character =>
                character is '\r' or '\n' or '\u0013' or '\u0014' or '\u0015'))
            throw new ArgumentException(
                "The Word SEQ identifier is invalid.",
                nameof(sequenceName));
        if (sequenceName.IndexOf('"') >= 0
            || sequenceName.IndexOf('\\') >= 0)
            throw new ArgumentException(
                "The Word SEQ identifier contains field-syntax characters.",
                nameof(sequenceName));
        var sequenceToken =
            sequenceName.Any(char.IsWhiteSpace)
                ? '"' + sequenceName + '"'
                : sequenceName;
        if (restartHeadingLevel < 0 || restartHeadingLevel > 9)
            throw new ArgumentOutOfRangeException(
                nameof(restartHeadingLevel));
        if (string.IsNullOrWhiteSpace(initialSequenceResult)
            || initialSequenceResult.Any(character =>
                character is '\r' or '\n' or '\u0013' or '\u0014' or '\u0015'))
            throw new ArgumentException(
                "The initial SEQ result is invalid.",
                nameof(initialSequenceResult));

        if (!string.IsNullOrWhiteSpace(numberBookmarkName))
        {
            _ = ValidateManagedEquationBookmarkName(
                numberBookmarkName,
                VisualTeXNativeNumberBookmarkPrefix,
                nameof(numberBookmarkName));
        }

        headingSeparator ??= string.Empty;
        if (headingSeparator.Any(character =>
                character is '\r' or '\n' or '\u0013' or '\u0014' or '\u0015'))
            throw new ArgumentException(
                "The heading separator contains an invalid Word field-control character.",
                nameof(headingSeparator));
        if (literalHeadingPrefix is not null
            && (restartHeadingLevel == 0 || string.IsNullOrWhiteSpace(literalHeadingPrefix)
                || literalHeadingPrefix.Length > 80 || literalHeadingPrefix.Any(char.IsControl)))
            throw new ArgumentException("A literal chapter prefix requires a valid heading scope and plain text.",
                nameof(literalHeadingPrefix));

        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;
        var equation = XElement.Parse(
            ExtractSingleOMath(semanticOmml),
            LoadOptions.PreserveWhitespace);
        var formulaNodes = equation
            .Elements()
            .Select(element => new XElement(element))
            .Cast<object>()
            .ToList();
        if (formulaNodes.Count == 0)
            throw new InvalidDataException(
                "A numbered Word equation requires a nonempty mathematical body.");

        var representativeRunProperties = equation
            .Descendants(word + "rPr")
            .FirstOrDefault(properties => properties.Element(word + "sz") is not null)
            ?? equation.Descendants(word + "rPr").FirstOrDefault();

        XElement RunProperties(bool noProof = false, bool italic = false)
        {
            var properties = new XElement(word + "rPr");
            foreach (var name in new[]
                     {
                         word + "rFonts",
                         word + "sz",
                         word + "szCs",
                     })
            {
                var property = representativeRunProperties?.Element(name);
                if (property is not null)
                    properties.Add(new XElement(property));
            }
            if (italic)
                properties.Add(new XElement(word + "i"));
            if (noProof)
                properties.Add(new XElement(word + "noProof"));
            return properties;
        }

        XElement ControlProperties() =>
            new(
                math + "ctrlPr",
                RunProperties(italic: true));

        XElement BoundaryRun(XElement content) =>
            new(
                math + "r",
                RunProperties(italic: true),
                content);

        XElement InstructionRun(string instruction) =>
            new(
                math + "r",
                RunProperties(),
                new XElement(
                    word + "instrText",
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    instruction));

        XElement ResultRun(string result) =>
            new(
                math + "r",
                RunProperties(noProof: true),
                new XElement(math + "t", result));

        XElement TextRun(string text) =>
            new(
                math + "r",
                RunProperties(),
                new XElement(math + "t", text));

        IEnumerable<object> NativeField(string instruction, string result)
        {
            yield return BoundaryRun(
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "begin")));
            yield return InstructionRun(instruction);
            yield return BoundaryRun(
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "separate")));
            yield return ResultRun(result);
            yield return BoundaryRun(
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "end")));
        }

        var numberNodes = new List<object>();
        if (!string.IsNullOrWhiteSpace(numberBookmarkName))
        {
            numberNodes.Add(
                new XElement(
                    word + "bookmarkStart",
                    new XAttribute(word + "id", "31901"),
                    new XAttribute(word + "name", numberBookmarkName)));
        }

        if (restartHeadingLevel > 0)
        {
            // STYLEREF reads an automatic paragraph number, not a number typed
            // into heading text. Preserve a literal heading as native text;
            // never manufacture a false STYLEREF result or change user lists.
            if (literalHeadingPrefix is null)
                numberNodes.AddRange(NativeField(
                    $" STYLEREF {restartHeadingLevel} \\w \\* MERGEFORMAT ", "1"));
            else
            {
                var prefixRun = TextRun(literalHeadingPrefix);
                prefixRun.AddFirst(new XElement(math + "rPr", new XElement(math + "nor")));
                numberNodes.Add(prefixRun);
            }
            if (!string.IsNullOrEmpty(headingSeparator))
            {
                // A numbering delimiter is literal text, not a mathematical
                // operator. BuildUp must not replace a requested hyphen with
                // U+2212 or reclassify a letter-based chapter label as variables.
                var separatorRun = TextRun(headingSeparator);
                separatorRun.AddFirst(new XElement(math + "rPr", new XElement(math + "nor")));
                numberNodes.Add(separatorRun);
            }
        }

        if (restartSequenceValue.HasValue && restartSequenceValue.Value < 1)
            throw new ArgumentOutOfRangeException(nameof(restartSequenceValue));
        var sequenceSwitch = (restartSequenceValue.HasValue
            ? $" \\r {restartSequenceValue.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}"
            : string.Empty) + (restartHeadingLevel > 0
            ? $" \\s {restartHeadingLevel}"
            : string.Empty);
        numberNodes.AddRange(
            NativeField(
                $" SEQ {sequenceToken}{sequenceSwitch} \\* ARABIC \\* MERGEFORMAT ",
                initialSequenceResult));
        if (!string.IsNullOrWhiteSpace(numberBookmarkName))
        {
            numberNodes.Add(
                new XElement(
                    word + "bookmarkEnd",
                    new XAttribute(word + "id", "31901")));
        }

        var delimiter = new XElement(
            math + "d",
            new XElement(
                math + "dPr",
                ControlProperties()),
            new XElement(math + "e", numberNodes));

        var equationBody = new XElement(
            math + "e",
            formulaNodes);
        equationBody.Add(
            TextRun("#"),
            delimiter);

        return new XElement(
                math + "oMath",
                new XElement(
                    math + "eqArr",
                    new XElement(
                        math + "eqArrPr",
                        new XElement(
                            math + "maxDist",
                            new XAttribute(math + "val", "1")),
                        ControlProperties()),
                    equationBody))
            .ToString(SaveOptions.DisableFormatting);
    }

    private static string ValidateManagedEquationBookmarkName(
        string? bookmarkName,
        string expectedPrefix,
        string parameterName)
    {
        var value = bookmarkName ?? string.Empty;
        if (!value.StartsWith(expectedPrefix, StringComparison.Ordinal)
            || value.Length != expectedPrefix.Length + 32)
            throw new ArgumentException(
                $"The VisualTeX bookmark must use the {expectedPrefix}<32-hex-FormulaId> form.",
                parameterName);
        var identifier = value.Substring(expectedPrefix.Length);
        if (!Guid.TryParseExact(identifier, "N", out var parsed))
            throw new ArgumentException(
                $"The VisualTeX bookmark must use the {expectedPrefix}<32-hex-FormulaId> form.",
                parameterName);
        return parsed.ToString("N");
    }

    internal static bool HasVisualTeXNativeEquationNumber(string wordOpenXml)
    {
        var equation = XElement.Parse(
            ExtractSingleOMath(wordOpenXml),
            LoadOptions.PreserveWhitespace);
        return TryResolveVisualTeXNativeEquationNumber(
            equation,
            out _,
            out _,
            out _);
    }

    internal static bool HasVisualTeXDirectSequenceEquationNumber(
        string wordOpenXml,
        string? formulaId = null)
    {
        var equation = XElement.Parse(
            ExtractSingleOMath(wordOpenXml),
            LoadOptions.PreserveWhitespace);
        return TryResolveVisualTeXDirectSequenceEquationNumber(
            equation,
            formulaId);
    }

    /// <summary>
    /// Detects Word's native display-equation-number host independently of
    /// VisualTeX ownership. Word linear input such as "#(2)" is normalized to
    /// one m:eqArr/m:e whose mathematical body is followed by a top-level '#'
    /// run and one parenthesized m:d number. Pure Word documents have no
    /// VisualTeX bookmark, metadata or SEQ field, so semantic readers must use
    /// the host topology rather than ownership markers.
    /// </summary>
    internal static bool HasWordNativeEquationNumberHost(string wordOpenXml)
    {
        var equation = XElement.Parse(
            ExtractSingleOMath(wordOpenXml),
            LoadOptions.PreserveWhitespace);
        return TryResolveWordNativeEquationNumberHost(
            equation,
            out _,
            out _,
            out _);
    }

    internal static string ReadWordNativeEquationNumberText(string wordOpenXml)
    {
        var equation = XElement.Parse(ExtractSingleOMath(wordOpenXml), LoadOptions.PreserveWhitespace);
        if (!TryResolveWordNativeEquationNumberHost(equation, out _, out _, out var delimiter)
            || delimiter is null) return string.Empty;
        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;
        // Field-code retrieval/view flags must not redefine the numbering label.
        // Only current text and cached field RESULTS participate, never instrText,
        // deleted old numerals or move-from provenance.
        var text = new StringBuilder();
        var fieldResults = new List<bool>();
        foreach (var node in delimiter.Descendants().Where(IsCurrentOmmlContent))
        {
            if (node.Name == word + "fldChar")
            {
                var kind = (string?)node.Attribute(word + "fldCharType");
                if (kind == "begin") fieldResults.Add(false);
                else if (kind == "separate")
                {
                    if (fieldResults.Count == 0) throw new InvalidDataException("A native number field separator has no owner.");
                    fieldResults[fieldResults.Count - 1] = true;
                }
                else if (kind == "end")
                {
                    if (fieldResults.Count == 0) throw new InvalidDataException("A native number field end has no owner.");
                    fieldResults.RemoveAt(fieldResults.Count - 1);
                }
                continue;
            }
            // Word may export field instructions as m:t, not w:instrText.
            // Field-state semantics therefore outrank the XML text-node name.
            // Nested results in an outer instruction remain instruction data.
            if ((node.Name == math + "t" || node.Name == word + "t")
                && fieldResults.All(isResult => isResult)) text.Append(node.Value);
        }
        if (fieldResults.Count != 0) throw new InvalidDataException("A native number field is unclosed.");
        return text.ToString();
    }

    /// <summary>
    /// Returns only the mathematical body of a Word-native numbered OMath.
    /// This is a semantic projection: it never claims ownership of the number
    /// and never mutates SEQ/REF/bookmark state in the document.
    /// </summary>
    internal static string StripWordNativeEquationNumberHost(
        string wordOpenXml)
    {
        var equation = XElement.Parse(
            ExtractSingleOMath(wordOpenXml),
            LoadOptions.PreserveWhitespace);
        var removed = 0;
        while (TryResolveWordNativeEquationNumberHost(
                   equation,
                   out var body,
                   out var separatorIndex,
                   out _)
               && body is not null
               && separatorIndex >= 0)
        {
            equation = StripEquationNumberWrapperBody(
                body,
                separatorIndex,
                "The Word-native equation-number wrapper contains no formula body.");
            removed++;
            if (removed > 8)
                throw new InvalidDataException(
                    "The Word-native equation-number wrapper is recursively malformed.");
        }

        return equation.ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>
    /// Replaces only the mathematical body of one Word-native numbered OMath.
    /// The original m:eqArr host, '#'-separator, number delimiter, field codes,
    /// bookmarks, and control properties are retained verbatim. This is the
    /// required edit path for a formula whose numbering is owned by Word.
    /// </summary>
    internal static string ReplaceWordNativeEquationNumberHostBody(
        string numberedWordOpenXml,
        string semanticOmml)
    {
        if (string.IsNullOrWhiteSpace(numberedWordOpenXml))
            throw new ArgumentException(
                "A numbered Word OMML host is required.",
                nameof(numberedWordOpenXml));
        if (string.IsNullOrWhiteSpace(semanticOmml))
            throw new ArgumentException(
                "Semantic OMML is required.",
                nameof(semanticOmml));

        var numberedEquation = XElement.Parse(
            ExtractSingleOMath(numberedWordOpenXml),
            LoadOptions.PreserveWhitespace);
        if (!TryResolveWordNativeEquationNumberHost(
                numberedEquation,
                out var numberedBody,
                out var separatorIndex,
                out _)
            || numberedBody is null
            || separatorIndex < 0)
        {
            throw new InvalidDataException(
                "The source OMML is not one Word-native numbered equation host.");
        }

        var semanticEquation = XElement.Parse(
            ExtractSingleOMath(semanticOmml),
            LoadOptions.PreserveWhitespace);
        if (TryResolveWordNativeEquationNumberHost(
                semanticEquation,
                out _,
                out _,
                out _))
        {
            semanticEquation = XElement.Parse(
                StripWordNativeEquationNumberHost(semanticOmml),
                LoadOptions.PreserveWhitespace);
        }

        XNamespace math = MathNamespace;
        var replacementBody = semanticEquation
            .Elements()
            .Where(element => element.Name != math + "ctrlPr")
            .Select(element => new XElement(element))
            .ToArray();
        if (replacementBody.Length == 0)
            throw new InvalidDataException(
                "The replacement OMML contains no mathematical body.");

        var numberedChildren = numberedBody.Elements().ToArray();
        if (separatorIndex >= numberedChildren.Length)
            throw new InvalidDataException(
                "The Word-native equation-number separator is outside its equation body.");

        var separator = NormalizeWordNativeNumberSeparator(
            numberedChildren[separatorIndex]);
        var preservedSuffix = numberedChildren
            .Skip(separatorIndex + 1)
            .Select(element => new XElement(element))
            .ToArray();

        numberedBody.ReplaceNodes(
            replacementBody
                .Concat(new[] { separator })
                .Concat(preservedSuffix));

        return numberedEquation.ToString(SaveOptions.DisableFormatting);
    }

    internal static string StripVisualTeXNativeEquationNumber(string wordOpenXml) =>
        StripVisualTeXNativeEquationNumberCore(
            wordOpenXml,
            allowUnboundDirectSequence: false);

    internal static string StripVisualTeXNativeEquationNumberForManagedRepair(
        string wordOpenXml) =>
        StripVisualTeXNativeEquationNumberCore(
            wordOpenXml,
            allowUnboundDirectSequence: true);

    // Compatibility name used by the managed-numbering repair and its XML
    // regression tests. Direct SEQ numbering is recognized by Word's native
    // host topology, including equations without VisualTeX aliases.
    internal static string StripManagedVisualTeXNativeEquationNumber(
        string wordOpenXml) =>
        StripVisualTeXNativeEquationNumberForManagedRepair(wordOpenXml);

    private static string StripVisualTeXNativeEquationNumberCore(
        string wordOpenXml,
        bool allowUnboundDirectSequence)
    {
        var equation = XElement.Parse(
            ExtractSingleOMath(wordOpenXml),
            LoadOptions.PreserveWhitespace);
        var removed = 0;
        while (TryResolveVisualTeXNativeEquationNumber(
                   equation,
                   out var body,
                   out var separatorIndex,
                   out _,
                   allowUnboundDirectSequence)
               && body is not null
               && separatorIndex >= 0)
        {
            equation = StripEquationNumberWrapperBody(
                body,
                separatorIndex,
                "The generated VisualTeX equation-number wrapper contains no formula body.");
            removed++;
            if (removed > 8)
                throw new InvalidDataException(
                    "The VisualTeX native equation-number wrapper is recursively malformed.");
        }
        return equation.ToString(SaveOptions.DisableFormatting);
    }

    // Word's current mathematical text is not necessarily a direct child of
    // m:r. With tracked edits, Word emits m:r/w:ins/m:t (and moveTo); deleted
    // content and historical property snapshots are not the current equation.
    // All native-number readers share this view. It only selects source nodes:
    // it never accepts revisions or rewrites the serialized history.
    private static bool IsCurrentOmmlContent(XElement element)
    {
        XNamespace word = WordNamespace;
        return !element.AncestorsAndSelf().Any(ancestor =>
            ancestor.Name == word + "del"
            || ancestor.Name == word + "moveFrom"
            || (ancestor.Name.Namespace == word
                && ancestor.Name.LocalName.EndsWith("PrChange", StringComparison.Ordinal)));
    }

    private static IEnumerable<XElement> CurrentOmmlTextElements(
        XElement container,
        XName name) =>
        container.Descendants(name).Where(IsCurrentOmmlContent);

    private static IEnumerable<XElement> CurrentMathRunTextElements(XElement run)
    {
        XNamespace math = MathNamespace;
        return CurrentOmmlTextElements(run, math + "t").Where(text =>
            ReferenceEquals(text.Ancestors(math + "r").FirstOrDefault(), run));
    }

    private static string CurrentMathRunText(XElement run) =>
        string.Concat(CurrentMathRunTextElements(run).Select(text => text.Value));

    private static XElement StripEquationNumberWrapperBody(
        XElement body,
        int separatorIndex,
        string emptyBodyMessage)
    {
        XNamespace math = MathNamespace;
        var bodyElements = body.Elements().ToArray();
        if (separatorIndex < 0 || separatorIndex >= bodyElements.Length)
            throw new InvalidDataException(
                "The equation-number wrapper separator is outside its equation body.");

        var formulaElements = bodyElements
            .Take(separatorIndex)
            .Select(element => new XElement(element))
            .ToList();
        var separator = bodyElements[separatorIndex];
        var separatorText = CurrentMathRunText(separator);
        if (separatorText.Length > 1
            && separatorText.EndsWith("#", StringComparison.Ordinal))
        {
            var mergedTail = new XElement(separator);
            var tailTexts = CurrentMathRunTextElements(mergedTail).ToArray();
            if (tailTexts.Length > 0)
            {
                var lastText = tailTexts[tailTexts.Length - 1];
                if (lastText.Value.EndsWith("#", StringComparison.Ordinal))
                    lastText.Value =
                        lastText.Value.Substring(
                            0,
                            lastText.Value.Length - 1);
                if (tailTexts.Any(text => !string.IsNullOrEmpty(text.Value)))
                    formulaElements.Add(mergedTail);
            }
        }

        if (formulaElements.Count == 0)
            throw new InvalidDataException(emptyBodyMessage);

        return new XElement(math + "oMath", formulaElements);
    }

    private static XElement NormalizeWordNativeNumberSeparator(
        XElement separator)
    {
        XNamespace math = MathNamespace;
        var separatorText = CurrentMathRunText(separator);
        if (string.Equals(
                separatorText,
                "#",
                StringComparison.Ordinal))
            return new XElement(separator);

        if (!separatorText.EndsWith(
                "#",
                StringComparison.Ordinal))
        {
            throw new InvalidDataException(
                "The Word-native equation-number separator no longer ends with '#'.");
        }

        // Word can merge the final formula text and '#' into one math run.
        // During an edit, only the '#' belongs to the numbering host. Keep the
        // original run properties and revision wrappers. Change only the
        // current formula-text prefix; historical/deleted text is untouched.
        // Rebuilding a bare m:r here loses Word rPr and the separator's author.
        var normalized = new XElement(separator);
        var texts = CurrentMathRunTextElements(normalized).ToArray();
        for (var index = 0; index < texts.Length; index++)
            texts[index].Value = index == texts.Length - 1 ? "#" : string.Empty;
        return normalized;
    }

    private static bool TryResolveWordNativeEquationNumberHost(
        XElement equation,
        out XElement? body,
        out int separatorIndex,
        out XElement? numberDelimiter)
    {
        body = null;
        separatorIndex = -1;
        numberDelimiter = null;
        XNamespace math = MathNamespace;

        // A Word-native numbered display equation is a host wrapper around one
        // equation array. Refuse to strip a '#' construct from an OMath that has
        // any other mathematical root sibling: that is ordinary formula syntax,
        // not the native numbering host.
        var rootChildren = equation.Elements()
            .Where(element => element.Name != math + "ctrlPr")
            .ToArray();
        if (rootChildren.Length != 1
            || rootChildren[0].Name != math + "eqArr")
            return false;

        var equationArray = rootChildren[0];
        var entries = equationArray.Elements(math + "e").ToArray();
        if (entries.Length != 1)
            return false;

        var candidateBody = entries[0];
        var children = candidateBody.Elements().ToArray();
        var found = false;
        for (var index = 0; index + 1 < children.Length; index++)
        {
            var separator = children[index];
            if (separator.Name != math + "r")
                continue;
            var separatorText = CurrentMathRunText(separator);
            if (!separatorText.EndsWith("#", StringComparison.Ordinal))
                continue;

            var delimiter = children[index + 1];
            if (delimiter.Name != math + "d"
                || !IsWordNativeEquationNumberDelimiter(delimiter))
                continue;

            // Word may append m:ctrlPr to m:e after normalizing the native
            // '#(...)' host. Any later mathematical node means the delimiter is
            // still part of the formula, so do not reinterpret it as numbering.
            if (children
                .Skip(index + 2)
                .Any(element => element.Name != math + "ctrlPr"))
                continue;

            if (found)
            {
                // Multiple plausible native-number suffixes make the host
                // ambiguous. Semantic extraction must never guess.
                body = null;
                separatorIndex = -1;
                numberDelimiter = null;
                return false;
            }

            body = candidateBody;
            separatorIndex = index;
            numberDelimiter = delimiter;
            found = true;
        }

        return found;
    }

    private static bool IsWordNativeEquationNumberDelimiter(
        XElement delimiter)
    {
        XNamespace math = MathNamespace;
        var entries = delimiter.Elements(math + "e").ToArray();
        if (entries.Length != 1)
            return false;

        var properties = delimiter.Element(math + "dPr");
        var begin = (string?)properties
            ?.Element(math + "begChr")
            ?.Attribute(math + "val");
        var end = (string?)properties
            ?.Element(math + "endChr")
            ?.Attribute(math + "val");

        // Missing dPr means Word's default parenthesis delimiter. If Word emits
        // explicit characters, require the same native '#(...)' shape observed
        // for built-in numbered equations.
        if (!string.IsNullOrEmpty(begin)
            && !string.Equals(begin, "(", StringComparison.Ordinal))
            return false;
        if (!string.IsNullOrEmpty(end)
            && !string.Equals(end, ")", StringComparison.Ordinal))
            return false;

        return true;
    }

    private static bool TryResolveVisualTeXNativeEquationNumber(
        XElement equation,
        out XElement? body,
        out int separatorIndex,
        out XElement? numberDelimiter,
        bool allowUnboundDirectSequence = false)
    {
        body = null;
        separatorIndex = -1;
        numberDelimiter = null;
        XNamespace math = MathNamespace;
        // VisualTeX numbers an equation through its one root equation array.
        // Word also allows several arrays side by side; those are never ours.
        var equationArrays = equation.Elements(math + "eqArr").Take(2).ToArray();
        if (equationArrays.Length != 1) return false;
        var entries = equationArrays[0].Elements(math + "e").ToArray();
        if (entries.Length != 1) return false;
        var candidateBody = entries[0];
        var children = candidateBody.Elements().ToArray();
        for (var index = 0; index + 1 < children.Length; index++)
        {
            var candidateSeparator = children[index];
            if (candidateSeparator.Name != math + "r") continue;
            var separatorText = CurrentMathRunText(candidateSeparator);
            if (!separatorText.EndsWith("#", StringComparison.Ordinal)) continue;
            var candidateNumber = children[index + 1];
            if (candidateNumber.Name != math + "d") continue;
            var numberText = string.Concat(
                candidateNumber.Descendants(math + "t").Select(text => text.Value));
            var generatedPlaceholder = numberText.IndexOf(
                VisualTeXEquationNumberPlaceholderPrefix,
                StringComparison.Ordinal) >= 0;
            XNamespace word = WordNamespace;
            var fieldCode = string.Concat(
                candidateNumber
                    .Descendants(word + "instrText")
                    .Select(text => text.Value));
            var generatedReference = numberText.IndexOf(
                    "REF " + VisualTeXNativeNumberBookmarkPrefix,
                    StringComparison.OrdinalIgnoreCase) >= 0
                || fieldCode.IndexOf(
                    "REF " + VisualTeXNativeNumberBookmarkPrefix,
                    StringComparison.OrdinalIgnoreCase) >= 0;
            // Arbitrary user SEQ names are valid native numbering only inside
            // Word's complete terminal '#(...)' host. A matching field alone
            // must not make us discard following math or other root siblings.
            var generatedDirectSequence =
                TryResolveWordNativeEquationNumberHost(
                    equation,
                    out _,
                    out var nativeSeparatorIndex,
                    out var nativeNumber)
                && nativeSeparatorIndex == index
                && ReferenceEquals(nativeNumber, candidateNumber)
                && IsVisualTeXDirectSequenceDelimiter(
                    candidateNumber,
                    formulaId: null);
            if (!generatedPlaceholder && !generatedReference && !generatedDirectSequence)
                continue;
            body = candidateBody;
            separatorIndex = index;
            numberDelimiter = candidateNumber;
            return true;
        }
        return false;
    }

    private static bool TryResolveVisualTeXDirectSequenceEquationNumber(
        XElement equation,
        string? formulaId)
    {
        return TryResolveWordNativeEquationNumberHost(
                equation,
                out _,
                out _,
                out var delimiter)
            && delimiter is not null
            && IsVisualTeXDirectSequenceDelimiter(delimiter, formulaId);
    }

    private static bool IsVisualTeXDirectSequenceDelimiter(
        XElement delimiter,
        string? formulaId)
    {
        XNamespace word = WordNamespace;
        XNamespace math = MathNamespace;
        // Word accepts w:instrText when the prepared OMath is imported, but its
        // native math normalizer commonly serializes that field instruction back as
        // m:r/m:t while COM still exposes the same live Field.Code.Text. Recognize
        // both spellings; never rewrite either representation in place.
        var fieldCode = string.Concat(
            CurrentOmmlTextElements(delimiter, word + "instrText").Select(text => text.Value))
            + string.Concat(
                CurrentOmmlTextElements(delimiter, math + "t").Select(text => text.Value));
        if (!Regex.IsMatch(
                fieldCode,
                @"\bSEQ\s+(?:""[^""]+""|[^\s\\]+)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            return false;
        if (Regex.IsMatch(
                fieldCode,
                @"\bREF\s+VTEqNum_",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            return false;

        // Native Word equation numbering is structural. FormulaId/bookmark
        // ownership is intentionally irrelevant.
        return true;
    }
}
