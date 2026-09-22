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

        var expectedOmml =
            WordOmmlConverter.TransformMathMlToOmml(request.MathMl!);
        var expectedSignature =
            WordOmmlConverter.ComputeImportedOmmlContentSignature(
                expectedOmml);
        var actualSignature =
            WordOmmlConverter.ComputeImportedOmmlContentSignature(
                payload.WordOpenXml!);

        if (string.Equals(
                expectedSignature,
                actualSignature,
                StringComparison.Ordinal)
            || request.AdditionalValidOmmlContentSignatures.Any(
                signature => string.Equals(
                    signature,
                    actualSignature,
                    StringComparison.Ordinal)))
            return;

        if (request.RequiredOmmlLatexFragments.Count > 0)
        {
            var actualLatex =
                NormalizeLatexFragment(
                    payload.Latex);
            var containsEveryRequiredFragment =
                request.RequiredOmmlLatexFragments
                    .Select(
                        NormalizeLatexFragment)
                    .Where(fragment =>
                        fragment.Length > 0)
                    .All(fragment =>
                        actualLatex.IndexOf(
                            fragment,
                            StringComparison.Ordinal)
                        >= 0);
            if (containsEveryRequiredFragment)
                return;
        }

        // Word is allowed to canonicalize equivalent mathematical structures
        // when BuildUp materializes the native equation tree. A common example
        // is ordinary parenthesis runs becoming Word's fenced form, which reads
        // back as \left(...\right). Raw OMML hashes intentionally remain the
        // fast path above, but a hash mismatch is not itself a semantic failure.
        // Compare the canonical MathML semantics before rejecting the mutation.
        if (!string.IsNullOrWhiteSpace(
                payload.MathMl))
        {
            var expectedSemanticSignature =
                MathTypeMtefCodec.SemanticSignature(
                    request.MathMl!);
            var actualSemanticSignature =
                MathTypeMtefCodec.SemanticSignature(
                    payload.MathMl!);
            if (string.Equals(
                    expectedSemanticSignature,
                    actualSemanticSignature,
                    StringComparison.Ordinal))
                return;
        }

        var expectedLatex =
            NormalizeLatexFragment(
                MathMlToLatexConverter.Convert(
                    request.MathMl!));
        var actualLatexForDiagnostic =
            NormalizeLatexFragment(
                payload.Latex);

        throw new InvalidDataException(
            "The inserted Word equation differs from the requested mathematical content and from every valid Word-native adjacent-inline semantic outcome. "
            + $"display={request.DisplayMode}; expectedLatex=[{expectedLatex}]; "
            + $"actualLatex=[{actualLatexForDiagnostic}]; "
            + $"expectedSignature=[{expectedSignature}]; actualSignature=[{actualSignature}].");
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
