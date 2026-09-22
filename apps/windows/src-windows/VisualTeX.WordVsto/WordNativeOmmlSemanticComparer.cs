using System.Xml.Linq;

namespace VisualTeX.WordVsto;

internal sealed class WordNativeOmmlSemanticComparison
{
    internal bool Equivalent { get; set; }
    internal bool ExactOmmlMatch { get; set; }
    internal string ExpectedOmmlSignature { get; set; } = string.Empty;
    internal string ActualOmmlSignature { get; set; } = string.Empty;
    internal string ExpectedMathMlSignature { get; set; } = string.Empty;
    internal string ActualMathMlSignature { get; set; } = string.Empty;
    internal string? SemanticExtractionError { get; set; }
}

/// <summary>
/// Single authority for deciding whether Word's materialized OMML still
/// represents the requested mathematics. XML-content hashes remain a fast
/// path, but Word is free to canonicalize equivalent native structures during
/// BuildUp, so the final decision is made from canonical MathML semantics.
/// </summary>
internal static class WordNativeOmmlSemanticComparer
{
    internal static WordNativeOmmlSemanticComparison CompareOmml(
        string expectedWordOpenXml,
        string actualWordOpenXml,
        bool display,
        IReadOnlyCollection<string>? additionalValidOmmlContentSignatures = null)
    {
        if (string.IsNullOrWhiteSpace(expectedWordOpenXml))
            throw new ArgumentException(
                "Expected OMML is required.",
                nameof(expectedWordOpenXml));
        if (string.IsNullOrWhiteSpace(actualWordOpenXml))
            throw new ArgumentException(
                "Actual WordOpenXML is required.",
                nameof(actualWordOpenXml));

        // Equation numbering is host structure, not mathematics.
        // Normalize it here so every caller reaches the same semantic authority
        // instead of maintaining its own SEQ/bookmark-specific stripping rule.
        var expectedSemanticOmml =
            display
                ? WordOmmlConverter.StripWordNativeEquationNumberHost(
                    expectedWordOpenXml)
                : expectedWordOpenXml;
        var actualSemanticOmml =
            display
                ? WordOmmlConverter.StripWordNativeEquationNumberHost(
                    actualWordOpenXml)
                : actualWordOpenXml;

        var expectedOmmlSignature =
            WordOmmlConverter.ComputeImportedOmmlContentSignature(
                expectedSemanticOmml);
        var actualOmmlSignature =
            WordOmmlConverter.ComputeImportedOmmlContentSignature(
                actualSemanticOmml);
        var exactOmmlMatch =
            string.Equals(
                expectedOmmlSignature,
                actualOmmlSignature,
                StringComparison.Ordinal)
            || (additionalValidOmmlContentSignatures?.Any(
                    signature => string.Equals(
                        signature,
                        actualOmmlSignature,
                        StringComparison.Ordinal))
                ?? false);

        if (exactOmmlMatch)
        {
            return new WordNativeOmmlSemanticComparison
            {
                Equivalent = true,
                ExactOmmlMatch = true,
                ExpectedOmmlSignature = expectedOmmlSignature,
                ActualOmmlSignature = actualOmmlSignature,
            };
        }

        try
        {
            var expectedMathMl =
                WordOmmlConverter.TransformOmmlToMathMl(
                    expectedSemanticOmml,
                    display);
            var actualMathMl =
                WordOmmlConverter.TransformOmmlToMathMl(
                    actualSemanticOmml,
                    display);
            var expectedMathMlSignature =
                ComputeMathMlSemanticSignature(
                    expectedMathMl);
            var actualMathMlSignature =
                ComputeMathMlSemanticSignature(
                    actualMathMl);

            return new WordNativeOmmlSemanticComparison
            {
                Equivalent = string.Equals(
                    expectedMathMlSignature,
                    actualMathMlSignature,
                    StringComparison.Ordinal),
                ExactOmmlMatch = false,
                ExpectedOmmlSignature = expectedOmmlSignature,
                ActualOmmlSignature = actualOmmlSignature,
                ExpectedMathMlSignature = expectedMathMlSignature,
                ActualMathMlSignature = actualMathMlSignature,
            };
        }
        catch (Exception error)
        {
            return new WordNativeOmmlSemanticComparison
            {
                Equivalent = false,
                ExactOmmlMatch = false,
                ExpectedOmmlSignature = expectedOmmlSignature,
                ActualOmmlSignature = actualOmmlSignature,
                SemanticExtractionError =
                    error.GetType().Name + ": " + error.Message,
            };
        }
    }

    internal static WordNativeOmmlSemanticComparison CompareMathMlToWordOpenXml(
        string expectedMathMl,
        string actualWordOpenXml,
        bool display,
        IReadOnlyCollection<string>? additionalValidOmmlContentSignatures = null)
    {
        if (string.IsNullOrWhiteSpace(expectedMathMl))
            throw new ArgumentException(
                "Expected MathML is required.",
                nameof(expectedMathMl));

        var expectedOmml =
            WordOmmlConverter.TransformMathMlToOmml(
                expectedMathMl);
        return CompareOmml(
            expectedOmml,
            actualWordOpenXml,
            display,
            additionalValidOmmlContentSignatures);
    }

    internal static string ComputeMathMlSemanticSignature(
        string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            return string.Empty;

        var normalized =
            XDocument.Parse(
                mathMl,
                LoadOptions.PreserveWhitespace);
        foreach (var token in
                 normalized
                     .Descendants()
                     .Where(element =>
                         element.Name.LocalName == "mi"))
        {
            var variant =
                ((string?)token.Attribute(
                    "mathvariant")
                 ?? string.Empty)
                .Trim()
                .ToLowerInvariant();
            var normalizedVariant =
                variant switch
                {
                    "bold-italic" =>
                        "bold",
                    "sans-serif-bold-italic" =>
                        "bold-sans-serif",
                    "sans-serif-italic" =>
                        "sans-serif",
                    _ =>
                        variant,
                };
            if (!string.Equals(
                    variant,
                    normalizedVariant,
                    StringComparison.Ordinal))
            {
                token.SetAttributeValue(
                    "mathvariant",
                    normalizedVariant);
            }
        }

        return MathTypeMtefCodec.SemanticSignature(
            normalized.Root?.ToString(
                SaveOptions.DisableFormatting)
            ?? mathMl);
    }
}
