using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
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

            // Numbering is a live OMath/Field topology, not an XML export result.
            // Word can omit an exact OMath at a serialization boundary; that
            // must never turn an ordinary edit into a number-attachment edit.
            numberRange = TryResolveNumberContentRange(math, exact);
            if (numberRange is null)
                return false;

            fields = exact.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                if (!WordNumberingFieldCollection.TryCaptureCurrentCode(field, out code, out var instruction))
                    continue;

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

            if (sequenceResult.Start < numberRange.Start || sequenceResult.End > numberRange.End
                || (styleResult is not null && (styleResult.Start < numberRange.Start || styleResult.End > numberRange.End)))
                return false;

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
        using var operationMetric =
            WordOperationMetrics.Measure(
                "FormatConversion.NativeOmmlNumbering.Detach");
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
                WordEquationNumberFormatAuthority.Resolve(
                    document));
        var nativeHeadingLevel = format.UsesHeading ? format.HeadingLevel : 0;
        string? literalHeadingPrefix = null;
        if (format.UsesHeading)
        {
            var scope = WordEquationNumbering.ResolveHeadingScopeAtPosition(document, position, format.Id);
            literalHeadingPrefix = ResolveLiteralHeadingPrefix(document, scope);
        }

        return WordOmmlConverter.BuildWordNativeNumberedOmml(
            semanticOmml,
            ResolveEquationSequenceName(document),
            numberBookmarkName: null,
            nativeHeadingLevel,
            nativeHeadingLevel > 0
                ? format.Separator
                : string.Empty,
            literalHeadingPrefix: literalHeadingPrefix);
    }

    // Batch writers (LaTeX redraw, format conversion) number many new equations
    // in one operation. Resolve the number format, the equation label and the
    // heading scopes of every target once (CaptureHeadingScopesAtPositions is
    // the full resolver, run for all positions together) instead of rescanning
    // the document's headings for each formula; then build each numbered
    // equation exactly as PrepareNumberedOmml does.
    internal sealed class NumberedOmmlBatch
    {
        private readonly Document _document;
        private readonly EquationNumberFormat _format;
        private readonly string _sequenceName;
        private readonly IReadOnlyDictionary<int, ResolvedEquationHeadingScope> _scopes;
        private readonly Dictionary<int, string?> _literalPrefixes = new();

        private NumberedOmmlBatch(
            Document document,
            EquationNumberFormat format,
            string sequenceName,
            IReadOnlyDictionary<int, ResolvedEquationHeadingScope> scopes)
        {
            _document = document;
            _format = format;
            _sequenceName = sequenceName;
            _scopes = scopes;
        }

        internal static NumberedOmmlBatch Capture(Document document, IEnumerable<int> positions)
        {
            var format = EquationNumberFormat.Resolve(
                WordEquationNumberFormatAuthority.Resolve(document));
            var scopes = format.UsesHeading
                ? WordEquationNumbering.CaptureHeadingScopesAtPositions(document, format.Id, positions)
                : new Dictionary<int, ResolvedEquationHeadingScope>();
            return new NumberedOmmlBatch(
                document, format, ResolveEquationSequenceName(document), scopes);
        }

        // A batch numbering plan supplies the series' own sequence name and, for
        // the one equation that takes over an explicit start, its \r value.
        internal string Prepare(int position, string semanticOmml,
            string? sequenceName = null, int? restart = null)
        {
            if (string.IsNullOrWhiteSpace(semanticOmml))
                throw new ArgumentException(
                    "Numbered OMML preparation requires semantic OMML.",
                    nameof(semanticOmml));
            var nativeHeadingLevel = _format.UsesHeading ? _format.HeadingLevel : 0;
            string? literalHeadingPrefix = null;
            if (_format.UsesHeading)
            {
                var scope = _scopes.TryGetValue(position, out var captured)
                    ? (captured.ScopeStart, captured.ScopeEnd, captured.NumberText)
                    : WordEquationNumbering.ResolveHeadingScopeAtPosition(_document, position, _format.Id);
                var key = scope.ScopeStart;
                if (!_literalPrefixes.TryGetValue(key, out literalHeadingPrefix)
                    || key == int.MinValue)
                {
                    literalHeadingPrefix = ResolveLiteralHeadingPrefix(_document, scope);
                    if (key != int.MinValue) _literalPrefixes[key] = literalHeadingPrefix;
                }
            }
            return WordOmmlConverter.BuildWordNativeNumberedOmml(
                semanticOmml,
                sequenceName ?? _sequenceName,
                numberBookmarkName: null,
                nativeHeadingLevel,
                nativeHeadingLevel > 0
                    ? _format.Separator
                    : string.Empty,
                restart?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "1",
                literalHeadingPrefix: literalHeadingPrefix,
                restartSequenceValue: restart);
        }
    }

    internal static string? ResolveLiteralHeadingPrefix(
        Document document, (int ScopeStart, int ScopeEnd, string NumberText) scope)
    {
        // The document numbering authority supplies 0 / 0.0 before the first
        // heading. Only a real automatic list heading is readable by STYLEREF.
        if (scope.ScopeStart == int.MinValue) return scope.NumberText;
        Range? heading = null;
        ListFormat? list = null;
        try
        {
            heading = document.Range(scope.ScopeStart, scope.ScopeEnd);
            list = heading.ListFormat;
            return list.ListType == WdListType.wdListNoNumbering ? scope.NumberText : null;
        }
        finally { Release(list); Release(heading); }
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
                    if (!WordNumberingFieldCollection.TryCaptureCurrentCode(field, out code, out var instruction))
                        continue;
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

    internal static WordFormulaHostDescriptor RefreshCanonicalNativeFields(
        Document document,
        WordFormulaHostDescriptor host)
    {
        Range? liveRange = null;
        OMaths? maths = null;
        OMath? math = null;
        try
        {
            liveRange = WordFormulaHostSemanticReader.CreateRange(document, host.Range);
            maths = liveRange.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    "The canonical native OMML host no longer contains exactly one equation.");
            math = maths[1];
            return RefreshCanonicalNativeFields(math, host);
        }
        finally { Release(math); Release(maths); Release(liveRange); }
    }

    // The batch already retains each exact native owner across preceding edits.
    // Borrow it for the same topology/field checks, rather than resolving its
    // numeric range back into another OMath collection in the same UI turn.
    // hasRevisions: false when the caller proved the document has no tracked
    // change (see WordNumberingFieldCollection.TryCaptureCurrentCode).
    internal static WordFormulaHostDescriptor RefreshCanonicalNativeFields(
        OMath math,
        WordFormulaHostDescriptor host,
        bool hasRevisions = true)
    {
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
                if (!WordNumberingFieldCollection.TryCaptureCurrentCode(field, out code, out var instruction, hasRevisions))
                    continue;

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

            numberRange = TryResolveNumberContentRange(math, exact)
                ?? throw new InvalidDataException("The equation has no native #(number) host.");
            if (sequenceResult.Start < numberRange.Start || sequenceResult.End > numberRange.End
                || (styleResult is not null && (styleResult.Start < numberRange.Start || styleResult.End > numberRange.End)))
                throw new InvalidDataException("The native number fields escaped their exact delimiter argument.");

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
        }
    }

    // Resolve Word's native number ARGUMENT, not an arithmetic span starting at
    // the first field result. A literal chapter prefix is valid current content
    // before SEQ and must be part of native REF bookmarks as well. Using OMath's
    // structural COM model avoids XML serialization inside the undo transaction.
    internal static Range ResolveNumberContentRange(Range equation) =>
        TryResolveNumberContentRange(equation)
        ?? throw new InvalidDataException("The equation has no native #(number) host.");

    internal static Range? TryResolveNumberContentRange(Range equation)
    {
        OMaths? maths = null;
        OMath? math = null;
        try
        {
            maths = equation.OMaths;
            if (maths.Count != 1) return null;
            math = maths[1];
            return TryResolveNumberContentRange(math, equation);
        }
        finally { Release(math); Release(maths); }
    }

    private static Range? TryResolveNumberContentRange(OMath math, Range equation)
    {
        OMathFunctions? roots = null;
        OMathFunction? arrayFunction = null;
        OMathEqArray? array = null;
        OMathArgs? entries = null;
        OMath? body = null;
        OMathFunctions? functions = null;
        OMathFunction? delimiterFunction = null;
        OMathDelim? delimiter = null;
        OMathArgs? arguments = null;
        OMath? number = null;
        Range? range = null;
        Range? arrayRange = null;
        Range? bodyRange = null;
        Range? delimiterRange = null;
        Range? separator = null;
        try
        {
            roots = math.Functions;
            if (roots.Count != 1) return null;
            arrayFunction = roots[1];
            if (arrayFunction.Type != WdOMathFunctionType.wdOMathFunctionEqArray)
                return null;
            arrayRange = arrayFunction.Range;
            if (arrayRange.Start != equation.Start || arrayRange.End != equation.End) return null;
            array = arrayFunction.EqArray;
            entries = array.E;
            if (entries.Count != 1) return null;
            body = entries.Item(1);
            bodyRange = body.Range;
            functions = body.Functions;
            if (functions.Count < 1) return null;
            delimiterFunction = functions[functions.Count];
            if (delimiterFunction.Type != WdOMathFunctionType.wdOMathFunctionDelim)
                return null;
            delimiterRange = delimiterFunction.Range;
            if (delimiterRange.End != bodyRange.End || delimiterRange.Start <= bodyRange.Start) return null;
            separator = bodyRange.Duplicate;
            separator.SetRange(delimiterRange.Start - 1, delimiterRange.Start);
            if (separator.Text != "#") return null;
            delimiter = delimiterFunction.Delim;
            if (delimiter.BegChar != '(' || delimiter.EndChar != ')') return null;
            arguments = delimiter.E;
            if (arguments.Count != 1) return null;
            number = arguments.Item(1);
            range = number.Range;
            return range.Duplicate;
        }
        finally
        {
            Release(separator); Release(delimiterRange); Release(bodyRange); Release(arrayRange);
            Release(range); Release(number); Release(arguments); Release(delimiter);
            Release(delimiterFunction); Release(functions); Release(body); Release(entries);
            Release(array); Release(arrayFunction); Release(roots);
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
        using var operationMetric =
            WordOperationMetrics.Measure(
                "FormatConversion.NativeOmmlRebuild."
                + (numbered ? "numbered" : "unnumbered"));
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
                WordFormulaHostLookup.ResolveLocal(
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

    // The semantic OMML of one live equation, read before Undo opens.
    internal static string ReadSemanticOmml(Range exact) => ExtractSemanticOmml(exact);

    // Batch numbering writer: the numbered OMML was prepared before Undo by a
    // batch plan; replace exactly this display OMath and leave the field values
    // to the batch's single ordered refresh.
    internal static WordFormulaHostDescriptor ReplaceWithPreparedNumbered(
        Document document,
        WordFormulaHostDescriptor host,
        string preparedNumberedOmml)
    {
        using var operationMetric =
            WordOperationMetrics.Measure("DisplayNumbering.Omml.Replace");
        Range? range = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? exact = null;
        Range? replacement = null;
        try
        {
            range = WordFormulaHostSemanticReader.CreateRange(document, host.Range);
            maths = range.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    "The equation to number no longer contains exactly one equation.");
            math = maths[1];
            exact = math.Range.Duplicate;
            if (!WordFormulaHostSemanticReader.SameAddress(exact, host.Range))
                throw new InvalidDataException("The equation to number moved before it was numbered.");
            if (math.Type != WdOMathType.wdOMathDisplay)
                throw new InvalidDataException(
                    "Word-native equation numbering requires wdOMathDisplay.");

            RemoveLegacyOmmlBookmarks(document, exact);
            replacement = WordOmmlConverter.ReplaceWithPreparedOmmlDirect(
                document,
                exact,
                preparedNumberedOmml,
                display: true,
                mathFontName: document.OMathFontName);
            WordFormulaHostLayout.ConfigureDisplayParagraph(replacement);

            var numbered = WordFormulaHostLookup.ResolveLocal(
                    document,
                    replacement,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "The numbered equation cannot be resolved locally.");
            numbered.FormulaId = null;
            numbered.Numbering = new WordFormulaNumberingDescriptor
            {
                Numbered = true,
                FormulaId = null,
                Position = "right",
                ContainerKind = WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                ContainerRange = numbered.Range,
            };
            return numbered;
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
        var xml = WordOmmlConverter.ReadLocalEquationXml(exact);
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
