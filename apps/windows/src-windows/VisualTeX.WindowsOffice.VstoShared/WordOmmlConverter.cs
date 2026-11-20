using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordOmmlEquationNumbers;
using static VisualTeX.WordVsto.WordOmmlFromMathMl;
using static VisualTeX.WordVsto.WordOmmlToMathMl;

namespace VisualTeX.WordVsto;

/// <summary>
/// Converting between MathML and Word's OMML: the two transforms (the PowerPoint add-in calls them by reflection),
/// the OMML namespaces and reading one equation's OMML. The other parts are WordOmmlFromMathMl, WordOmmlToMathMl,
/// WordOmmlInsertion, WordOmmlEquationNumbers, WordOmmlFingerprints, WordOmmlDocuments and WordOmmlBatchSources.
/// </summary>
internal static partial class WordOmmlConverter
{
    internal const string MathNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/math";
    internal const string WordNamespace =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    internal const string FormulaBookmarkName = "VisualTeXFormula";
    private const string VisualTeXEquationSequenceName = "VisualTeXEquation";

    internal static string TransformMathMlToOmml(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException("VisualTeX did not provide MathML for the Word OMML formula.");
        ValidateMathMlForOmml(mathMl);
        mathMl = NormalizeMathTypeBinomialPiles(mathMl);
        mathMl = NormalizeMathJaxBinomialDelimiters(mathMl);
        mathMl = NormalizePlainBracketPairs(mathMl);
        mathMl = NormalizeFencedMathMlTables(mathMl);
        mathMl = NormalizeNestedEmptyBaseScripts(mathMl);
        mathMl = NormalizeMathMlPrescripts(mathMl);
        mathMl = NormalizeMathMlAccents(mathMl);
        mathMl = NormalizeNaryArguments(mathMl);
        mathMl = NormalizeMathMlTokenStyleAuthority(mathMl);
        mathMl = MathMlColorContract.MaterializeInheritedColors(mathMl);
        mathMl = ProjectMathMlSpacesForOmml(mathMl).MathMl;
        var placeholderResult = ReplaceExtendedIntegralsWithOfficePlaceholders(mathMl);
        mathMl = placeholderResult.MathMl;
        var display = IsBlockMathMl(mathMl);
        // The projection is VisualTeX's own: the installed Office's MML2OMML.XSL
        // changes with Office updates and is never read.
        var omml = ExtractSingleOMath(MathMlOmmlProjection.Project(mathMl));
        omml = RestoreExtendedIntegralCharacters(omml, placeholderResult.NaryCharacters);
        omml = NormalizeExplicitUprightRuns(omml, mathMl);
        omml = NormalizeAppliedFunctionStructures(omml, mathMl);
        omml = NormalizeExplicitTableColumnAlignment(omml, mathMl);
        omml = NormalizeOmmlPlaceholderVisibility(omml);
        omml = NormalizeOptionalPrescriptSlots(omml);
        omml = NormalizeDisplayNaryOmml(omml, display);
        omml = NormalizeGeneratedOmmlArgumentOrder(omml);
        return omml;
    }

    internal static string TransformOmmlToMathMl(string wordOpenXml, bool display)
    {
        using var operationMetric = VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("WordOmmlConverter.TransformOmmlToMathMl");
        var omml = StripVisualTeXNativeEquationNumber(wordOpenXml);
        // Remove only the empty zero-extent representation of a missing native
        // prescript half before reverse conversion; real phantom content is retained.
        var reverseDocument = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace reverseMath = MathNamespace;
        foreach (var script in reverseDocument.Descendants(reverseMath + "sPre"))
            foreach (var slot in script.Elements().Where(element => element.Name == reverseMath + "sub" || element.Name == reverseMath + "sup"))
            {
                var slotChildren = slot.Elements().Take(2).ToArray();
                var phantom = slotChildren.Length == 1 ? slotChildren[0] : null;
                if (phantom is null || !IsHiddenPhantom(phantom) || HasVisibleMathContent(phantom)) continue;
                var properties = phantom.Element(reverseMath + "phantPr");
                if (new[] { "zeroAsc", "zeroDesc", "zeroWid" }.All(name => IsMathBooleanEnabled(properties, reverseMath + name)))
                    slot.RemoveNodes();
            }
        omml = reverseDocument.Root!.ToString(SaveOptions.DisableFormatting);
        // The projection is VisualTeX's own: the installed Office's OMML2MML.XSL
        // changes with Office updates and is never read. The result uses the one
        // default MathML namespace the renderer and MathType bridge accept.
        var canonicalRoot = CanonicalizeMathMlElement(OmmlMathMlProjection.Project(omml));
        RestoreMergedNumericPunctuationTokens(canonicalRoot);
        RestoreMathSymbolTextTokens(canonicalRoot);
        RestoreWordCanonicalAsteriskOperators(canonicalRoot);
        RestoreOmmlNoBarFractionSemantics(omml, canonicalRoot);
        RestoreOmmlMatrixColumnAlignment(omml, canonicalRoot);
        RestoreOmmlAccentSemantics(omml, canonicalRoot);
        RestoreOmmlLimitBaseSemantics(omml, canonicalRoot);
        RestoreOmmlScriptBaseRunSemantics(omml, canonicalRoot);
        RestoreOmmlFunctionApplicationRuns(omml, canonicalRoot);
        RestoreOmmlExplicitSpaceTokens(canonicalRoot);
        canonicalRoot.SetAttributeValue("display", display ? "block" : "inline");
        return canonicalRoot.ToString(SaveOptions.DisableFormatting);
    }

    internal static string CombineInlineOmml(
        params string[] equations)
    {
        if (equations is null || equations.Length == 0)
            throw new ArgumentException(
                "At least one OMML equation is required.",
                nameof(equations));

        XNamespace math = MathNamespace;
        var combined = new XElement(math + "oMath");
        foreach (var source in equations)
        {
            var equation = XElement.Parse(
                ExtractSingleOMath(source),
                LoadOptions.PreserveWhitespace);
            foreach (var child in equation.Elements())
                combined.Add(new XElement(child));
        }

        if (!combined.Elements().Any())
            throw new InvalidDataException(
                "The combined inline OMML equation has no mathematical content.");

        return combined.ToString(
            SaveOptions.DisableFormatting);
    }

    internal static string ExtractSingleOMath(string omml)
    {
        if (string.IsNullOrWhiteSpace(omml))
            throw new InvalidDataException("Office produced an empty OMML transformation.");
        var document = FormulaXml.Parse(omml, "OMML", LoadOptions.None);
        XNamespace math = MathNamespace;
        // This API promises one complete equation, not the first equation in a
        // larger range. Counting also detects an illegally nested oMath root.
        // A tracked-deleted or moved-away copy (Track Changes) is history, not
        // a second current equation.
        XNamespace word = WordNamespace;
        var equations = document.Descendants(math + "oMath")
            .Where(equation => !equation.Ancestors().Any(ancestor =>
                ancestor.Name == word + "del" || ancestor.Name == word + "moveFrom"))
            .Take(2).ToArray();
        if (equations.Length == 0)
            throw new InvalidDataException("Office MathML conversion did not produce an m:oMath node.");
        if (equations.Length != 1)
            throw new InvalidDataException("Expected exactly one OMML equation; the input contains multiple equations.");
        return equations[0].ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>
    /// The WordOpenXML of one live equation, for readers that parse exactly one
    /// m:oMath from it. Some Word builds serialize an exact OMath range
    /// incompletely at a field/run boundary (a numbered equation holds SEQ or
    /// STYLEREF fields), sometimes only on the first export after an edit: the
    /// package then has no m:oMath. Export the same range once more, then a
    /// range two characters wider, which still holds only this equation. Only a
    /// package with exactly one equation is accepted.
    /// </summary>
    internal static string ReadLocalEquationXml(Range exactEquation)
    {
        if (exactEquation is null) throw new ArgumentNullException(nameof(exactEquation));
        for (var attempt = 1; attempt <= 2; attempt++)
        {
            var exact = TryReadExactlyOneEquationXml(exactEquation);
            if (exact is not null) return exact;
            WordDoubleClickHook.TraceMessage(
                $"omml-xml-incomplete-export range={exactEquation.Start}:{exactEquation.End} attempt={attempt}");
        }

        Range? probe = null;
        try
        {
            probe = exactEquation.Duplicate;
            try { probe.MoveStart(WdUnits.wdCharacter, -2); } catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
            try { probe.MoveEnd(WdUnits.wdCharacter, 2); } catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
            var expanded = TryReadExactlyOneEquationXml(probe);
            if (expanded is not null) return expanded;
        }
        finally { Release(probe); }

        throw new InvalidDataException(
            "Word could not export one local OMML equation from the resolved host.");
    }

    private static string? TryReadExactlyOneEquationXml(Range range)
    {
        string xml;
        try { xml = range.WordOpenXML ?? string.Empty; }
        catch (COMException) { return null; }
        return CountOMathElements(xml) == 1 ? xml : null;
    }

    // Streams the package and stops at a second current equation (copies inside
    // tracked deletions or moves are history); callers parse the XML themselves.
    internal static int CountOMathElements(string xml)
    {
        if (string.IsNullOrWhiteSpace(xml)) return 0;
        var count = 0;
        var revisionDepths = new Stack<int>();
        try
        {
            using var reader = XmlReader.Create(
                new StringReader(xml),
                new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
            while (reader.Read())
            {
                var revision = reader.NamespaceURI == WordNamespace
                    && reader.LocalName is "del" or "moveFrom";
                if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (revision && revisionDepths.Count > 0 && revisionDepths.Peek() == reader.Depth)
                        revisionDepths.Pop();
                    continue;
                }
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (revision)
                {
                    if (!reader.IsEmptyElement) revisionDepths.Push(reader.Depth);
                    continue;
                }
                if (revisionDepths.Count == 0
                    && reader.LocalName == "oMath"
                    && reader.NamespaceURI == MathNamespace
                    && ++count > 1)
                    break;
            }
        }
        catch (XmlException) { return 0; }
        return count;
    }

    internal static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
    }
}
