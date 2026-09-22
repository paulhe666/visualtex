namespace VisualTeX.WordVsto;

/// <summary>
/// Local postcondition checks for one host mutation. No document-wide
/// cardinality or unrelated identity checks are permitted here.
/// </summary>
internal static class WordFormulaMutationValidator
{
    internal static WordFormulaHostDescriptor ValidateInsertedHost(
        Microsoft.Office.Interop.Word.Document document,
        WordFormulaHostWriteRequest request,
        WordFormulaHostDescriptor inserted)
    {
        var localRange = WordFormulaHostSemanticReader.CreateRange(
            document,
            inserted.Range);
        try
        {
            var resolved = WordFormulaHostResolver.ResolveLocal(
                    document,
                    localRange,
                    request.Kind)
                ?? throw new InvalidDataException(
                    "The newly inserted formula host cannot be resolved locally.");

            if (resolved.Range.StoryType != inserted.Range.StoryType
                || resolved.Range.Start != inserted.Range.Start
                || resolved.Range.End != inserted.Range.End)
                throw new InvalidDataException(
                    "The newly inserted formula host moved before validation.");

            if (!string.Equals(
                    resolved.DisplayMode,
                    request.DisplayMode,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "The newly inserted formula host has the wrong inline/display mode.");

            var payload = WordFormulaHostSemanticReader.Read(
                document,
                resolved);
            switch (request.Kind)
            {
                case WordFormulaHostKind.VisualTeX:
                    ValidateVisualTeX(request, resolved, payload);
                    break;
                case WordFormulaHostKind.Omml:
                    ValidateOmml(request, payload);
                    break;
                default:
                    throw new NotSupportedException(
                        $"Mutation validation is not owned by this core for {request.Kind}.");
            }

            return resolved;
        }
        finally
        {
            Release(localRange);
        }
    }

    private static void ValidateVisualTeX(
        WordFormulaHostWriteRequest request,
        WordFormulaHostDescriptor resolved,
        WordFormulaSemanticPayload payload)
    {
        var expected = request.Metadata
            ?? throw new InvalidDataException(
                "VisualTeX validation requires target metadata.");
        var actual = payload.Metadata
            ?? throw new InvalidDataException(
                "The inserted VisualTeX host returned no embedded metadata.");

        if (!string.Equals(
                actual.FormulaId,
                expected.FormulaId,
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException(
                "The inserted VisualTeX FormulaId differs from the requested host.");

        if (!string.Equals(
                actual.Latex,
                expected.Latex,
                StringComparison.Ordinal))
            throw new InvalidDataException(
                "The inserted VisualTeX source differs from the requested formula.");

        if (!string.Equals(
                actual.DisplayMode,
                request.DisplayMode,
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException(
                "The inserted VisualTeX metadata has the wrong display mode.");

        if (!string.IsNullOrWhiteSpace(resolved.FormulaId)
            && !string.Equals(
                resolved.FormulaId,
                actual.FormulaId,
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException(
                "The VisualTeX host identity and embedded payload disagree.");
    }

    private static void ValidateOmml(
        WordFormulaHostWriteRequest request,
        WordFormulaSemanticPayload payload)
    {
        if (string.IsNullOrWhiteSpace(request.MathMl))
            throw new InvalidDataException(
                "OMML validation requires the requested MathML.");
        if (string.IsNullOrWhiteSpace(payload.WordOpenXml))
            throw new InvalidDataException(
                "The inserted OMML host returned no local WordOpenXML.");

        var display =
            string.Equals(
                request.DisplayMode,
                "block",
                StringComparison.OrdinalIgnoreCase);
        var comparison =
            WordNativeOmmlSemanticComparer.CompareMathMlToWordOpenXml(
                request.MathMl!,
                payload.WordOpenXml!,
                display,
                request.AdditionalValidOmmlContentSignatures);
        if (comparison.Equivalent)
            return;

        var expectedRawLatex =
            MathMlToLatexConverter.Convert(
                request.MathMl!);
        var actualRawLatex =
            payload.Latex;
        var expectedLatex =
            NormalizeLatexFragment(
                expectedRawLatex);
        var actualLatexForDiagnostic =
            NormalizeLatexFragment(
                actualRawLatex);
        var expectedWordNativeLatex =
            NormalizeWordNativeLatexSemantics(
                expectedRawLatex);
        var actualWordNativeLatex =
            NormalizeWordNativeLatexSemantics(
                actualRawLatex);

        throw new InvalidDataException(
            "The inserted Word equation differs from the requested mathematical content after Word-native semantic canonicalization. "
            + $"display={request.DisplayMode}; expectedLatex=[{expectedLatex}]; "
            + $"actualLatex=[{actualLatexForDiagnostic}]; "
            + $"expectedWordNativeLatex=[{expectedWordNativeLatex}]; "
            + $"actualWordNativeLatex=[{actualWordNativeLatex}]; "
            + $"expectedOmmlSignature=[{comparison.ExpectedOmmlSignature}]; "
            + $"actualOmmlSignature=[{comparison.ActualOmmlSignature}]; "
            + $"expectedMathMlSignature=[{comparison.ExpectedMathMlSignature}]; "
            + $"actualMathMlSignature=[{comparison.ActualMathMlSignature}]; "
            + $"semanticExtractionError=[{comparison.SemanticExtractionError ?? string.Empty}].");
    }

    private static string NormalizeLatexFragment(
        string? latex)
    {
        if (string.IsNullOrWhiteSpace(
                latex))
            return string.Empty;

        return string.Concat(
            latex!.Where(character =>
                !char.IsWhiteSpace(
                    character)));
    }

    internal static string NormalizeWordNativeLatexSemantics(
        string? latex)
    {
        if (string.IsNullOrWhiteSpace(latex))
            return string.Empty;

        var source = latex!;
        var normalized =
            new System.Text.StringBuilder(
                source.Length);
        for (var index = 0;
             index < source.Length;)
        {
            if (char.IsWhiteSpace(source[index]))
            {
                index++;
                continue;
            }

            if (source[index] == '\\'
                && index + 1 < source.Length
                && char.IsLetter(source[index + 1]))
            {
                var commandEnd = index + 2;
                while (commandEnd < source.Length
                       && char.IsLetter(source[commandEnd]))
                    commandEnd++;

                var command =
                    source.Substring(
                        index,
                        commandEnd - index);
                if (string.Equals(
                        command,
                        @"\left",
                        StringComparison.Ordinal)
                    || string.Equals(
                        command,
                        @"\right",
                        StringComparison.Ordinal))
                {
                    index = commandEnd;
                    // TeX uses an immediately following '.' as an invisible
                    // delimiter in \left. / \right. pairs. It is structural
                    // layout syntax, not a literal decimal point, so drop it
                    // together with the sizing command.
                    while (index < source.Length
                           && char.IsWhiteSpace(source[index]))
                        index++;
                    if (index < source.Length
                        && source[index] == '.')
                        index++;
                    continue;
                }
                if (string.Equals(
                        command,
                        @"\mid",
                        StringComparison.Ordinal))
                {
                    normalized.Append('∣');
                    index = commandEnd;
                    continue;
                }

                normalized.Append(command);
                index = commandEnd;
                continue;
            }

            normalized.Append(source[index]);
            index++;
        }

        return normalized.ToString();
    }

    private static void Release(object? value)
    {
        if (value is null
            || !System.Runtime.InteropServices.Marshal.IsComObject(value))
            return;
        try
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(value);
        }
        catch { }
    }
}
