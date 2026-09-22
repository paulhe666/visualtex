using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Pure Word-native numbering adapter for display OMML.
///
/// OMML has no VisualTeX-owned persistent identity. The document truth is the
/// exact OMath plus its local OOXML/fields. Numbered display equations use one
/// wdOMathDisplay whose body is Word's native m:eqArr + # + STYLEREF/SEQ form.
/// No VTOMML/VTEqNum/table/shape/frame metadata is created for OMML.
/// </summary>
internal static class WordNativeOmmlNumbering
{
    private static readonly string[] LegacyOmmlBookmarkPrefixes =
    {
        "VTOMML_",
        "VTEqNum_",
        "VTEq_",
        "VTEqCap_",
        "VTBL_",
        "VTEqAnc_",
        "VTEqSep_",
    };

    // Compatibility entry point. FormulaId is deliberately ignored for OMML.
    internal static bool TryResolve(
        Document document,
        Range hostRange,
        string? formulaId,
        out WordFormulaNumberingDescriptor descriptor) =>
        TryResolveUnownedNativeStructure(
            document,
            hostRange,
            out descriptor);

    internal static bool TryResolveUnownedNativeStructure(
        Document document,
        Range hostRange,
        out WordFormulaNumberingDescriptor descriptor)
    {
        descriptor = null!;
        if (document is null || hostRange is null)
            return false;

        OMaths? maths = null;
        OMath? math = null;
        Range? exact = null;
        Fields? fields = null;
        Field? field = null;
        Field? sequence = null;
        Field? styleRef = null;
        Range? code = null;
        Range? sequenceResult = null;
        Range? styleResult = null;
        Range? numberRange = null;
        try
        {
            maths = hostRange.OMaths;
            if (maths.Count != 1)
                return false;

            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay)
                return false;

            exact = math.Range.Duplicate;
            if (exact.StoryType != hostRange.StoryType
                || exact.Start != hostRange.Start
                || exact.End != hostRange.End)
                return false;

            var xml = exact.WordOpenXML ?? string.Empty;
            if (!WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                    xml,
                    formulaId: null))
                return false;

            fields = exact.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code.Duplicate;
                var instruction = code.Text ?? string.Empty;

                if (IsReferenceField(instruction))
                    return false;

                if (IsSequenceField(instruction))
                {
                    if (sequence is not null)
                        return false;
                    sequence = field;
                    field = null;
                }
                else if (instruction.TrimStart().StartsWith(
                             "STYLEREF ",
                             StringComparison.OrdinalIgnoreCase))
                {
                    if (styleRef is not null)
                        return false;
                    styleRef = field;
                    field = null;
                }
            }

            if (sequence is null)
                return false;

            sequenceResult = sequence.Result.Duplicate;
            if (styleRef is not null)
                styleResult = styleRef.Result.Duplicate;

            var start = styleResult is null
                ? sequenceResult.Start
                : Math.Min(styleResult.Start, sequenceResult.Start);
            var end = sequenceResult.End;
            numberRange = CreateStoryRange(
                document,
                exact.StoryType,
                start,
                end);

            descriptor = new WordFormulaNumberingDescriptor
            {
                Numbered = true,
                FormulaId = null,
                Position = "right",
                ContainerKind =
                    WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                ContainerRange = Address(exact),
                NumberRange = Address(numberRange),
            };
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(numberRange);
            Release(styleResult);
            Release(sequenceResult);
            Release(code);
            Release(styleRef);
            Release(sequence);
            Release(field);
            Release(fields);
            Release(exact);
            Release(math);
            Release(maths);
        }
    }

    // Retained as a compatibility surface for callers migrated in stages.
    // There is no adoption or document mutation anymore.
    internal static WordFormulaHostDescriptor AdoptUnownedNativeStructure(
        Document document,
        WordFormulaHostDescriptor host,
        string formulaId)
    {
        if (host.Kind != WordFormulaHostKind.Omml)
            throw new ArgumentException(
                "Only OMML can resolve Word-native numbering.");

        Range? hostRange = null;
        try
        {
            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            if (!TryResolveUnownedNativeStructure(
                    document,
                    hostRange,
                    out var numbering))
                throw new InvalidDataException(
                    "The OMML host is not one exact native #(SEQ) equation.");

            host.FormulaId = null;
            numbering.FormulaId = null;
            return WithNumbering(host, numbering);
        }
        finally
        {
            Release(hostRange);
        }
    }

    internal static WordFormulaHostDescriptor Attach(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (host.Kind != WordFormulaHostKind.Omml)
            throw new ArgumentException(
                "Word-native # numbering is only valid for OMML hosts.");
        if (!host.Display)
            throw new InvalidOperationException(
                "Only display OMML can use Word-native equation numbering.");

        if (host.Numbering.Numbered
            && host.Numbering.ContainerKind ==
                WordFormulaNumberingContainerKind.CanonicalNativeOmml)
        {
            // Fresh pure-OMML insertion already materialized the final native
            // #(SEQ) host. Do not rediscover the same container before field
            // refresh; one final local resolve after Word updates the fields is
            // sufficient.
            return RefreshCanonicalNativeFields(
                document,
                host);
        }

        var current =
            WordFormulaNumberingResolver.ResolveLocal(
                document,
                host);
        if (current.ContainerKind ==
            WordFormulaNumberingContainerKind.CanonicalNativeOmml)
            return RefreshCanonicalNativeFields(
                document,
                host);
        if (current.Numbered)
            throw new InvalidDataException(
                "OMML must be detached from its previous numbering topology before native numbering is attached.");

        return Rebuild(
            document,
            host,
            numbered: true);
    }

    internal static WordFormulaHostDescriptor Detach(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (host.Kind != WordFormulaHostKind.Omml)
            throw new ArgumentException(
                "Word-native # numbering detach received a non-OMML host.");
        return Rebuild(
            document,
            host,
            numbered: false);
    }

    internal static WordFormulaHostDescriptor RebuildNumbered(
        Document document,
        WordFormulaHostDescriptor host) =>
        Rebuild(
            document,
            host,
            numbered: true);

    // Legacy call surface retained for old compatibility tests. Pure OMML ignores
    // FormulaId and never serializes it into Word.
    internal static WordFormulaHostDescriptor RebuildNumbered(
        Document document,
        WordFormulaHostDescriptor host,
        string formulaId) =>
        RebuildNumbered(
            document,
            host);

    // Native Word paste already owns the resulting OMath structure. Do not re-key,
    // add bookmarks, or mutate its SEQ fields merely because it was pasted.
    internal static WordFormulaHostDescriptor RekeyPastedNativeHost(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (host.Kind != WordFormulaHostKind.Omml)
            throw new ArgumentException(
                "Native OMML paste validation received a non-OMML host.");

        Range? exact = null;
        try
        {
            exact =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            if (!TryResolveUnownedNativeStructure(
                    document,
                    exact,
                    out var numbering))
                throw new InvalidDataException(
                    "The pasted OMML is not one Word-native #(SEQ) equation.");
            host.FormulaId = null;
            numbering.FormulaId = null;
            return WithNumbering(host, numbering);
        }
        finally
        {
            Release(exact);
        }
    }

    internal static string PrepareNumberedOmml(
        Document document,
        int position,
        string semanticOmml)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(semanticOmml))
            throw new ArgumentException(
                "Numbered OMML preparation requires semantic OMML.",
                nameof(semanticOmml));

        var format =
            EquationNumberFormat.Resolve(
                WordEquationNumbering
                    .GetEquationNumberFormatId(document));
        var nativeHeadingLevel = 0;
        if (format.UsesHeading)
        {
            var scopes =
                WordEquationNumbering
                    .CaptureHeadingScopesAtPositions(
                        document,
                        format.Id,
                        new[] { position });
            if (scopes.TryGetValue(
                    position,
                    out var scope)
                && scope.ScopeStart != int.MinValue)
            {
                nativeHeadingLevel =
                    format.HeadingLevel;
            }
        }

        return WordOmmlConverter.BuildWordNativeNumberedOmml(
            semanticOmml,
            ResolveEquationSequenceName(document),
            numberBookmarkName: null,
            nativeHeadingLevel,
            nativeHeadingLevel > 0
                ? format.Separator
                : string.Empty);
    }

    internal static void UpdateNativeFields(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (host.Kind != WordFormulaHostKind.Omml)
            return;

        Range? hostRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            fields = hostRange.Fields;

            // Word heading state must be current before the sequence field reads
            // it. Never rewrite Field.Code.Text inside an OMath.
            for (var pass = 0; pass < 2; pass++)
            {
                for (var index = 1; index <= fields.Count; index++)
                {
                    Release(code); code = null;
                    Release(field); field = fields[index];
                    code = field.Code.Duplicate;
                    var instruction = code.Text ?? string.Empty;
                    var isStyleRef =
                        instruction.TrimStart().StartsWith(
                            "STYLEREF ",
                            StringComparison.OrdinalIgnoreCase);
                    var isSequence =
                        IsSequenceField(
                            instruction);

                    if ((pass == 0 && isStyleRef)
                        || (pass == 1 && isSequence))
                        field.Update();
                }
            }
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(hostRange);
        }
    }

    private static WordFormulaHostDescriptor RefreshCanonicalNativeFields(
        Document document,
        WordFormulaHostDescriptor host)
    {
        Range? liveRange = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? exact = null;
        Fields? fields = null;
        Field? field = null;
        Field? sequence = null;
        Field? styleRef = null;
        Range? code = null;
        Range? sequenceResult = null;
        Range? styleResult = null;
        Range? numberRange = null;
        try
        {
            liveRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            maths = liveRange.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    "The canonical native OMML host no longer contains exactly one equation.");

            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay)
                throw new InvalidDataException(
                    "Canonical native numbering requires wdOMathDisplay.");

            exact = math.Range.Duplicate;
            if (!WordFormulaHostSemanticReader.SameAddress(
                    exact,
                    host.Range))
                throw new InvalidDataException(
                    "The canonical native OMML host moved before field refresh.");

            fields = exact.Fields;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code.Duplicate;
                var instruction =
                    code.Text
                    ?? string.Empty;

                if (IsReferenceField(instruction))
                    throw new InvalidDataException(
                        "Canonical native OMML numbering must not contain a REF field.");

                if (IsSequenceField(instruction))
                {
                    if (sequence is not null)
                        throw new InvalidDataException(
                            "Canonical native OMML contains more than one SEQ field.");
                    sequence = field;
                    field = null;
                    continue;
                }

                if (instruction.TrimStart().StartsWith(
                        "STYLEREF ",
                        StringComparison.OrdinalIgnoreCase))
                {
                    if (styleRef is not null)
                        throw new InvalidDataException(
                            "Canonical native OMML contains more than one STYLEREF field.");
                    styleRef = field;
                    field = null;
                }
            }

            if (sequence is null)
                throw new InvalidDataException(
                    "Canonical native OMML contains no SEQ field.");

            // Heading state must be current before SEQ \s reads it.
            styleRef?.Update();
            sequence.Update();

            sequenceResult =
                sequence.Result.Duplicate;
            if (styleRef is not null)
                styleResult =
                    styleRef.Result.Duplicate;

            // Word can change the field-result width while updating; reacquire
            // the exact live OMath boundary after the updates.
            Release(exact);
            exact =
                math.Range.Duplicate;

            var numberStart =
                styleResult is null
                    ? sequenceResult.Start
                    : Math.Min(
                        styleResult.Start,
                        sequenceResult.Start);
            numberRange =
                CreateStoryRange(
                    document,
                    exact.StoryType,
                    numberStart,
                    sequenceResult.End);

            var refreshed =
                new WordFormulaHostDescriptor
                {
                    Kind = WordFormulaHostKind.Omml,
                    Range = Address(exact),
                    DisplayMode = "block",
                    FormulaId = null,
                    Metadata = null,
                    MetadataAuthoritative = false,
                    WithinTable = host.WithinTable,
                    SourceMathMl = host.SourceMathMl,
                    Latex = host.Latex,
                };
            var descriptor =
                new WordFormulaNumberingDescriptor
                {
                    Numbered = true,
                    FormulaId = null,
                    Position = "right",
                    ContainerKind =
                        WordFormulaNumberingContainerKind
                            .CanonicalNativeOmml,
                    ContainerRange = Address(exact),
                    NumberRange = Address(numberRange),
                };
            return WithNumbering(
                refreshed,
                descriptor);
        }
        finally
        {
            Release(numberRange);
            Release(styleResult);
            Release(sequenceResult);
            Release(code);
            Release(styleRef);
            Release(sequence);
            Release(field);
            Release(fields);
            Release(exact);
            Release(math);
            Release(maths);
            Release(liveRange);
        }
    }

    internal static string ReadVisibleNumber(
        Document document,
        WordFormulaNumberingDescriptor numbering)
    {
        if (numbering.NumberRange is null)
            return string.Empty;

        Range? range = null;
        try
        {
            range =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    numbering.NumberRange);
            return (range.Text ?? string.Empty)
                .Trim()
                .Trim('(', ')')
                .Trim();
        }
        finally
        {
            Release(range);
        }
    }

    private static WordFormulaHostDescriptor Rebuild(
        Document document,
        WordFormulaHostDescriptor host,
        bool numbered)
    {
        Range? range = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? exact = null;
        Range? replacement = null;
        try
        {
            range =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            maths = range.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    "The OMML numbering source no longer contains exactly one equation.");

            math = maths[1];
            exact = math.Range.Duplicate;
            if (!WordFormulaHostSemanticReader.SameAddress(
                    exact,
                    host.Range))
                throw new InvalidDataException(
                    "The OMML numbering source moved before rebuild.");
            if (math.Type != WdOMathType.wdOMathDisplay)
                throw new InvalidDataException(
                    "Word-native equation numbering requires wdOMathDisplay.");

            var semanticOmml =
                ExtractSemanticOmml(exact);

            // A touched legacy OMML host is migrated to the pure Word model. This
            // is bounded to the exact local equation and its deterministic start
            // boundary; no global bookmark search participates.
            RemoveLegacyOmmlBookmarks(
                document,
                exact);

            var prepared =
                numbered
                    ? PrepareNumberedOmml(
                        document,
                        exact.Start,
                        semanticOmml)
                    : semanticOmml;

            replacement =
                WordOmmlConverter.ReplaceWithPreparedOmmlDirect(
                    document,
                    exact,
                    prepared,
                    display: true,
                    mathFontName: document.OMathFontName);
            WordFormulaHostLayout.ConfigureDisplayParagraph(
                replacement);

            var refreshed =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    replacement,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "The rebuilt OMML host cannot be resolved locally.");
            refreshed.FormulaId = null;

            if (!numbered)
            {
                refreshed.Numbering =
                    new WordFormulaNumberingDescriptor
                    {
                        Numbered = false,
                        FormulaId = null,
                        ContainerKind =
                            WordFormulaNumberingContainerKind.None,
                    };
                return refreshed;
            }

            return RefreshCanonicalNativeFields(
                document,
                refreshed);
        }
        finally
        {
            Release(replacement);
            Release(exact);
            Release(math);
            Release(maths);
            Release(range);
        }
    }

    private static string ExtractSemanticOmml(
        Range exact)
    {
        var xml = exact.WordOpenXML ?? string.Empty;
        if (WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                xml,
                formulaId: null))
            return WordOmmlConverter
                .StripManagedVisualTeXNativeEquationNumber(
                    xml);

        return WordOmmlConverter.ExtractSingleOMath(
            xml);
    }

    internal static string ResolveEquationSequenceName(
        Document document)
    {
        CaptionLabels? labels = null;
        CaptionLabel? label = null;
        try
        {
            labels = document.Application.CaptionLabels;
            label = labels[WdCaptionLabelID.wdCaptionEquation];
            var name = (label.Name ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(name)
                ? "Equation"
                : name;
        }
        catch
        {
            return "Equation";
        }
        finally
        {
            Release(label);
            Release(labels);
        }
    }

    internal static bool IsEquationSequenceFieldCode(
        Document document,
        string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;

        var name =
            ResolveEquationSequenceName(
                document);
        var code =
            instruction!.TrimStart();
        return code.StartsWith(
                   "SEQ " + name + " ",
                   StringComparison.OrdinalIgnoreCase)
            || code.StartsWith(
                   "SEQ \"" + name + "\" ",
                   StringComparison.OrdinalIgnoreCase);
    }

    internal static int ResolveNativeCrossReferenceItem(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (host.Kind != WordFormulaHostKind.Omml)
            throw new ArgumentException(
                "Native Word equation cross-reference items are only valid for OMML.");

        Range? exact = null;
        Fields? localFields = null;
        Field? localField = null;
        Range? localCode = null;
        Fields? documentFields = null;
        Field? candidate = null;
        Range? candidateCode = null;
        try
        {
            exact =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            localFields = exact.Fields;

            int? sequenceCodeStart = null;
            for (var index = 1;
                 index <= localFields.Count;
                 index++)
            {
                Release(localCode); localCode = null;
                Release(localField); localField = localFields[index];
                localCode = localField.Code.Duplicate;
                if (!IsEquationSequenceFieldCode(
                        document,
                        localCode.Text))
                    continue;
                if (sequenceCodeStart.HasValue)
                    throw new InvalidDataException(
                        "The native OMML host contains more than one Equation SEQ field.");
                sequenceCodeStart =
                    localCode.Start;
            }

            if (!sequenceCodeStart.HasValue)
                throw new InvalidDataException(
                    "The native OMML host contains no Word Equation SEQ field.");

            documentFields =
                document.Fields;
            var equationOrdinal = 0;
            for (var index = 1;
                 index <= documentFields.Count;
                 index++)
            {
                Release(candidateCode); candidateCode = null;
                Release(candidate); candidate = documentFields[index];
                candidateCode = candidate.Code.Duplicate;
                if (!IsEquationSequenceFieldCode(
                        document,
                        candidateCode.Text))
                    continue;

                equationOrdinal++;
                if (candidateCode.Start ==
                    sequenceCodeStart.Value)
                    return equationOrdinal;
            }

            throw new InvalidDataException(
                "The native OMML Equation SEQ field is not present in Word's document field order.");
        }
        finally
        {
            Release(candidateCode);
            Release(candidate);
            Release(documentFields);
            Release(localCode);
            Release(localField);
            Release(localFields);
            Release(exact);
        }
    }

    private static void RemoveLegacyOmmlBookmarks(
        Document document,
        Range exact)
    {
        Range? probe = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        try
        {
            var start =
                Math.Max(
                    document.Content.Start,
                    exact.Start - 1);
            var end =
                Math.Min(
                    document.Content.End,
                    exact.End + 1);
            probe = document.Range(start, end);
            bookmarks = probe.Bookmarks;

            for (var index = bookmarks.Count;
                 index >= 1;
                 index--)
            {
                Release(bookmarkRange); bookmarkRange = null;
                Release(bookmark); bookmark = bookmarks[index];
                var name = bookmark.Name ?? string.Empty;
                if (!LegacyOmmlBookmarkPrefixes.Any(prefix =>
                        name.StartsWith(
                            prefix,
                            StringComparison.OrdinalIgnoreCase)))
                    continue;

                bookmarkRange = bookmark.Range.Duplicate;
                var locallyOwned =
                    bookmarkRange.StoryType == exact.StoryType
                    && bookmarkRange.Start >= exact.Start - 1
                    && bookmarkRange.End <= exact.End + 1;
                if (locallyOwned)
                    bookmark.Delete();
            }
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(probe);
        }
    }

    internal static string DescribeNativeState(
        Document document,
        Range range,
        string formulaId = "")
    {
        OMaths? maths = null;
        OMath? math = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            maths = range.OMaths;
            var type = "none";
            if (maths.Count == 1)
            {
                math = maths[1];
                type = math.Type.ToString();
            }

            var codes = new List<string>();
            var results = new List<string>();
            fields = range.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code.Duplicate;
                var result = field.Result.Duplicate;
                try
                {
                    codes.Add(
                        (code.Text ?? string.Empty)
                            .Replace("\r", "\\r")
                            .Replace("\n", "\\n"));
                    results.Add(
                        $"{result.Start}:{result.End}=[{(result.Text ?? string.Empty).Trim()}]");
                }
                finally
                {
                    Release(result);
                }
            }

            var xml = range.WordOpenXML ?? string.Empty;
            var direct = false;
            try
            {
                direct =
                    WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                        xml,
                        formulaId: null);
            }
            catch { }

            return
                $"range={range.StoryType}:{range.Start}:{range.End} "
                + $"type={type} fields={fields.Count} "
                + $"codes=[{string.Join(" | ", codes)}] "
                + $"results=[{string.Join(" | ", results)}] "
                + $"direct={direct} "
                + $"eqArr={xml.IndexOf("<m:eqArr", StringComparison.OrdinalIgnoreCase) >= 0}";
        }
        catch (Exception error)
        {
            return "diagnostic-failed=" + error.Message;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(math);
            Release(maths);
        }
    }

    // Legacy compatibility helpers only. New OMML code must not create or depend
    // on these bookmarks.
    internal static string ReferenceBookmarkName(
        string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var parsed))
            throw new InvalidDataException(
                "Legacy equation reference requires a UUID FormulaId.");
        return "VTEqNum_" + parsed.ToString("N");
    }

    internal static bool TryParseReferenceBookmarkName(
        string? name,
        out string formulaId)
    {
        formulaId = string.Empty;
        const string prefix = "VTEqNum_";
        if (string.IsNullOrWhiteSpace(name)
            || !name!.StartsWith(
                prefix,
                StringComparison.OrdinalIgnoreCase))
            return false;

        var suffix =
            name.Substring(prefix.Length);
        if (!Guid.TryParseExact(
                suffix,
                "N",
                out var parsed))
            return false;

        formulaId = parsed.ToString("D");
        return true;
    }

    private static bool IsSequenceField(
        string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;
        return instruction!.TrimStart().StartsWith(
            "SEQ ",
            StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsReferenceField(
        string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return false;
        return instruction!.TrimStart().StartsWith(
            "REF ",
            StringComparison.OrdinalIgnoreCase);
    }

    private static WordFormulaHostDescriptor WithNumbering(
        WordFormulaHostDescriptor host,
        WordFormulaNumberingDescriptor numbering) =>
        new()
        {
            Kind = host.Kind,
            Range = host.Range,
            DisplayMode = host.DisplayMode,
            FormulaId = host.Kind == WordFormulaHostKind.Omml
                ? null
                : host.FormulaId,
            Metadata = host.Metadata,
            MetadataAuthoritative =
                host.MetadataAuthoritative,
            Numbering = numbering,
            WithinTable = host.WithinTable,
            SourceMathMl = host.SourceMathMl,
            Latex = host.Latex,
        };

    private static Range CreateStoryRange(
        Document document,
        WdStoryType storyType,
        int start,
        int end)
    {
        if (storyType == WdStoryType.wdMainTextStory)
            return document.Range(start, end);

        Range? story = null;
        Range? result = null;
        try
        {
            story = document.StoryRanges[storyType];
            result = story.Duplicate;
            result.SetRange(start, end);
            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(result);
            Release(story);
        }
    }

    private static WordFormulaRangeAddress Address(
        Range range) =>
        new()
        {
            StoryType = range.StoryType,
            Start = range.Start,
            End = range.End,
        };

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value))
            return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
