using VisualTeX.WindowsOffice.Contracts;

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
        using var operationMetric =
            WordOperationMetrics.Measure(
                "FormatConversion.ValidateInsertedHost."
                + request.Kind);
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

            if (request.Kind == WordFormulaHostKind.Omml)
            {
                // Pure OMML trusts Word's final native shape. Do not compare
                // inline/display mode, semantics, fingerprints or numbering
                // containers against the request. The only postcondition is that
                // one non-empty physical OMath still owns this exact range.
                ValidateOmml(localRange);
                return resolved;
            }

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
        Microsoft.Office.Interop.Word.Range hostRange)
    {
        Microsoft.Office.Interop.Word.OMaths? maths = null;
        Microsoft.Office.Interop.Word.OMath? math = null;
        Microsoft.Office.Interop.Word.Range? exact = null;
        try
        {
            maths = hostRange.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    "Word did not retain exactly one native OMath.");
            math = maths[1];
            exact = math.Range.Duplicate;
            if (exact.End <= exact.Start
                || exact.StoryType != hostRange.StoryType
                || exact.Start != hostRange.Start
                || exact.End != hostRange.End)
                throw new InvalidDataException(
                    "Word returned an invalid native OMath range.");
        }
        finally
        {
            Release(exact);
            Release(math);
            Release(maths);
        }
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
