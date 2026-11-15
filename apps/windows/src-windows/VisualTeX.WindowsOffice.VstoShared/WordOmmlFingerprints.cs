using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using static VisualTeX.WordVsto.WordOmmlConverter;
using static VisualTeX.WordVsto.WordOmmlEquationNumbers;
using static VisualTeX.WordVsto.WordOmmlFromMathMl;
using static VisualTeX.WordVsto.WordOmmlToMathMl;

namespace VisualTeX.WordVsto;

/// <summary>
/// Fingerprints and content signatures of OMML: the current, expected-version and legacy fingerprints, and the
/// canonical form of imported OMML they compare.
/// </summary>
internal static class WordOmmlFingerprints
{
    // v2 uses the SAME mathematical-content model as import verification. XML
    // namespace aliases and equivalent default properties are not content edits.
    // The prefix is part of the persisted value; never silently reinterpret v1.
    internal static string ComputeOmmlFingerprint(string wordOpenXml) =>
        OmmlFingerprintFormat.CanonicalPrefix + ComputeImportedOmmlContentSignature(wordOpenXml);

    internal static string ComputeOmmlFingerprintForExpectedVersion(
        string wordOpenXml, string expectedFingerprint) =>
        OmmlFingerprintFormat.GetVersion(expectedFingerprint) == 1
            ? ComputeLegacyOmmlFingerprint(wordOpenXml)
            : ComputeOmmlFingerprint(wordOpenXml);

    internal static bool MatchesStoredOmmlFingerprint(string wordOpenXml, string? expectedFingerprint) =>
        !string.IsNullOrWhiteSpace(expectedFingerprint)
        && string.Equals(ComputeOmmlFingerprintForExpectedVersion(wordOpenXml, expectedFingerprint!),
            expectedFingerprint, StringComparison.OrdinalIgnoreCase);

    internal static string UpgradeVerifiedOmmlFingerprint(string wordOpenXml, string storedFingerprint)
    {
        _ = storedFingerprint;
        ExtractSingleOMath(wordOpenXml);
        return ComputeOmmlFingerprint(wordOpenXml);
    }

    // Frozen v1 reader for existing documents. Do not add new normalizations here:
    // a legacy hash cannot be upgraded without the captured equation it describes.
    internal static string ComputeLegacyOmmlFingerprint(string wordOpenXml)
    {
        var normalized = StripVisualTeXNativeEquationNumber(wordOpenXml);
        var document = XDocument.Parse(normalized, LoadOptions.PreserveWhitespace);
        XNamespace word = WordNamespace;
        XNamespace math = MathNamespace;

        // Word stores the visible math size and other proof/font state in
        // ordinary run properties. These are presentation state, not formula
        // content, and Word may add them while importing the same OMML.
        document.Descendants(word + "rPr").Remove();
        document.Descendants(math + "ctrlPr").Remove();
        document.Descendants(word + "bookmarkStart").Remove();
        document.Descendants(word + "bookmarkEnd").Remove();
        NormalizeMathRunGrouping(document, math);

        // Word stores the visible math size in ordinary run properties. Font
        // size is presentation state, not formula content: changing 14 pt to
        // 18 pt must not force an OMML -> MathML -> LaTeX source refresh.
        document
            .Descendants()
            .Where(element => element.Name == word + "sz" || element.Name == word + "szCs")
            .Remove();

        normalized = document.Root?.ToString(SaveOptions.DisableFormatting) ?? normalized;
        using var hash = SHA256.Create();
        var bytes = hash.ComputeHash(Encoding.UTF8.GetBytes(normalized));
        return string.Concat(bytes.Select(value => value.ToString("x2")));
    }

    // Shared content model for first-import verification and v2 persistence.
    // Changes that alter persisted v2 digests need a version/reader review;
    // legacy identity remains isolated in ComputeLegacyOmmlFingerprint.
    // Namespace spelling, empty property containers, on/off
    // lexical forms and Word's default integral glyph are equivalent encodings,
    // not a reason to bind an identity to different mathematical content.
    internal static string ComputeImportedOmmlContentSignature(string wordOpenXml)
    {
        var document = XDocument.Parse(StripVisualTeXNativeEquationNumber(wordOpenXml), LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;
        document.Descendants(word + "rPr").Remove();
        document.Descendants(math + "ctrlPr").Remove();
        document.Descendants(word + "bookmarkStart").Remove();
        document.Descendants(word + "bookmarkEnd").Remove();
        NormalizeImportedMathOnOffProperties(document, math);
        // The Office MathML stylesheet emits a legacy scrLvl=0 hint for
        // display-style fraction arguments (e.g. \\dfrac). Word 2021 drops this
        // hint when importing the same numerator/denominator. Match only this
        // observed neutral encoding; preserve nonzero hints, valid argSz
        // properties, and every operand/structural node for strict comparison.
        foreach (var hint in document.Descendants(math + "scrLvl").Where(element =>
                     element.Parent?.Name == math + "argPr"
                     && (string?)element.Attribute(math + "val") == "0"
                     && !element.HasElements
                     && string.IsNullOrWhiteSpace(element.Value)
                     && element.Attributes().All(attribute =>
                         attribute.IsNamespaceDeclaration || attribute.Name == math + "val")).ToArray())
            hint.Remove();
        foreach (var nary in document.Descendants(math + "nary"))
        {
            var properties = nary.Element(math + "naryPr");
            if (properties is null) { properties = new XElement(math + "naryPr"); nary.AddFirst(properties); }
            if (properties.Element(math + "chr") is null)
                properties.AddFirst(new XElement(math + "chr", new XAttribute(math + "val", "∫")));
        }
        // ISO/IEC 29500 m:fPr/m:type defaults to bar both when the element
        // is absent and when val is absent. Word omits this default on import.
        // Keep every nondefault type: a stack or skewed fraction is not bar.
        foreach (var fraction in document.Descendants(math + "f"))
        {
            var properties = fraction.Element(math + "fPr");
            if (properties is null) { properties = new XElement(math + "fPr"); fraction.AddFirst(properties); }
            var type = properties.Element(math + "type");
            if (type is null) properties.AddFirst(new XElement(math + "type", new XAttribute(math + "val", "bar")));
            else if (type.Attribute(math + "val") is null) type.SetAttributeValue(math + "val", "bar");
        }
        // ISO/IEC 29500 m:accPr/m:chr defaults to the combining circumflex.
        // Word commonly removes an explicit U+0302 after importing \hat{x}; the
        // omission and explicit value are the same mathematical accent. Preserve
        // every nondefault accent character.
        foreach (var accent in document.Descendants(math + "acc"))
        {
            var properties = accent.Element(math + "accPr");
            if (properties is null)
            {
                properties = new XElement(math + "accPr");
                accent.AddFirst(properties);
            }
            var character = properties.Element(math + "chr");
            if (character is null)
                properties.AddFirst(new XElement(
                    math + "chr",
                    new XAttribute(math + "val", "\u0302")));
            else if (character.Attribute(math + "val") is null)
                character.SetAttributeValue(math + "val", "\u0302");
        }
        // ISO/IEC 29500 m:mPr/m:baseJc defaults to center. Word removes the
        // explicit center during matrix import; keep nondefault top/bot intact.
        // This is a representational equivalence, not permission to change a
        // matrix's cells, row order, alignment or any mathematical contents.
        foreach (var matrix in document.Descendants(math + "m"))
        {
            var justification = matrix.Element(math + "mPr")?.Element(math + "baseJc");
            if (justification is not null
                && ((string?)justification.Attribute(math + "val") is null or "center"))
                justification.Remove();
        }
        NormalizeImportedMatrixColumnGroups(document, math);
        // A separator is rendered only between distinct delimiter arguments.
        // A single argument (including a multirow equation array) has none;
        // Word drops sepChr there. Beginning/end delimiters are always retained.
        foreach (var delimiter in document.Descendants(math + "d"))
            if (delimiter.Elements(math + "e").Count() == 1)
                delimiter.Element(math + "dPr")?.Elements(math + "sepChr").Remove();
        NormalizeLegacyMathAlphabetSpelling(document, math);
        NormalizeMathScriptRunBoundaries(document, math);
        foreach (var text in document.Descendants(math + "t"))
        {
            var properties = text.Parent?.Element(math + "rPr");
            var literal = properties?.Elements().Any(e => (e.Name == math + "nor" || e.Name == math + "lit")
                && (string?)e.Attribute(math + "val") == "1") == true;
            if (!literal) text.Value = NormalizeImportedMathTextSpelling(text.Value);
        }
        NormalizeImportedDefaultProperties(document, math);
        NormalizeImportedPlainStyleDefaults(document, math);
        // The MathML stylesheet can leave a zero-length normal-text run between
        // operands. Word removes that empty run and joins its neighbours. Remove
        // only truly empty text/format-only runs in the import comparison: a real
        // space, alignment/break control, object, or empty argument container must
        // remain observable in both import verification and v2 identity.
        foreach (var run in document.Descendants(math + "r").Where(run =>
                     !run.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)
                     && run.Elements(math + "t").Any()
                     && run.Elements(math + "t").All(text => !text.HasElements && text.Value.Length == 0)
                     && run.Elements().All(child => child.Name == math + "t"
                         || (child.Name == math + "rPr"
                             && !child.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)
                             && child.Elements().All(property => property.Name == math + "nor"
                                 || property.Name == math + "lit" || property.Name == math + "sty"
                                 || property.Name == math + "scr")))).ToArray())
            run.Remove();
        foreach (var property in document.Descendants().Where(e => e.Name.Namespace == math
                     && e.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal)
                     && !e.HasElements && !e.Attributes().Any(a => !a.IsNamespaceDeclaration)
                     && string.IsNullOrWhiteSpace(e.Value)).ToArray())
            property.Remove();
        NormalizeImportedMathRunGrouping(document, math);
        var canonical = SerializeCanonicalOmmlElement(
            document.Root ?? throw new InvalidDataException("The imported OMML has no equation root."));
        using var hash = SHA256.Create();
        return string.Concat(hash.ComputeHash(Encoding.UTF8.GetBytes(canonical))
            .Select(value => value.ToString("x2")));
    }

    private static string SerializeCanonicalOmmlElement(XElement root)
    {
        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;
        var canonical = new StringBuilder();
        void AppendToken(string token) { canonical.Append(token.Length).Append(':').Append(token); }
        void AppendElement(XElement element)
        {
            canonical.Append('('); AppendToken(element.Name.ToString());
            foreach (var attribute in element.Attributes().Where(a => !a.IsNamespaceDeclaration)
                         .OrderBy(a => a.Name.ToString(), StringComparer.Ordinal))
            { canonical.Append('@'); AppendToken(attribute.Name.ToString()); AppendToken(attribute.Value); }
            foreach (var node in element.Nodes())
            {
                if (node is XElement child) AppendElement(child);
                else if (node is XText text && (element.Name == math + "t"
                    || element.Name == word + "t" || element.Name == word + "instrText"
                    || element.Name == word + "delText" || element.Name == word + "delInstrText"
                    || !string.IsNullOrWhiteSpace(text.Value)))
                { canonical.Append('#'); AppendToken(text.Value); }
            }
            canonical.Append(')');
        }
        AppendElement(root);
        return canonical.ToString();
    }

    private static void NormalizeImportedMathRunGrouping(XDocument document, XNamespace math)
    {
        // This is intentionally separate from the frozen v1 grouping routine.
        // m:r can contain Word controls, field code and annotated text. Only
        // plain, unannotated mathematical text with known style properties can
        // be concatenated without erasing a control or an extension payload.
        bool PlainRun(XElement run)
        {
            if (run.Name != math + "r" || run.Attributes().Any(a => !a.IsNamespaceDeclaration)
                || run.Nodes().OfType<XText>().Any(t => !string.IsNullOrWhiteSpace(t.Value))
                || run.Elements(math + "rPr").Count() > 1) return false;
            var texts = run.Elements(math + "t").ToArray();
            if (texts.Length == 0 || texts.Any(t => t.HasElements
                || t.Attributes().Any(a => !a.IsNamespaceDeclaration))) return false;
            if (run.Elements().Any(e => e.Name != math + "t" && e.Name != math + "rPr")) return false;
            var properties = run.Element(math + "rPr");
            return properties is null || (!properties.Attributes().Any(a => !a.IsNamespaceDeclaration)
                && !properties.Nodes().OfType<XText>().Any(t => !string.IsNullOrWhiteSpace(t.Value))
                && properties.Elements().All(p =>
                    (p.Name == math + "sty" || p.Name == math + "scr"
                        || p.Name == math + "nor" || p.Name == math + "lit")
                    && !p.HasElements && string.IsNullOrWhiteSpace(p.Value)
                    && p.Attributes().All(a => a.IsNamespaceDeclaration || a.Name == math + "val")));
        }

        foreach (var run in document.Descendants(math + "r").Where(PlainRun).ToArray())
        {
            var texts = run.Elements(math + "t").ToArray();
            texts[0].Value = string.Concat(texts.Select(t => t.Value));
            foreach (var extra in texts.Skip(1)) extra.Remove();
        }
        foreach (var parent in document.Root?.DescendantsAndSelf().ToArray() ?? Array.Empty<XElement>())
        {
            XElement? previous = null;
            string? previousKey = null;
            foreach (var run in parent.Elements().ToArray())
            {
                if (!PlainRun(run)) { previous = null; previousKey = null; continue; }
                var properties = run.Element(math + "rPr");
                var key = properties is null ? string.Empty : SerializeCanonicalOmmlElement(properties);
                if (previous is not null && string.Equals(previousKey, key, StringComparison.Ordinal))
                {
                    previous.Element(math + "t")!.Value += run.Element(math + "t")!.Value;
                    run.Remove();
                }
                else { previous = run; previousKey = key; }
            }
        }
    }

    private static void NormalizeImportedMatrixColumnGroups(
        XDocument document,
        XNamespace math)
    {
        foreach (var matrix in document.Descendants(math + "m"))
        {
            var columns = matrix.Element(math + "mPr")?.Element(math + "mcs");
            if (columns is null) continue;
            var groups = columns.Elements(math + "mc").ToArray();
            if (groups.Length == 0) continue;

            var expanded = new List<XElement>();
            var safe = true;
            foreach (var group in groups)
            {
                if (group.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)
                    || group.Elements().Count() != 1
                    || group.Element(math + "mcPr") is not XElement properties
                    || group.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
                {
                    safe = false;
                    break;
                }

                if (properties.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)
                    || properties.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
                {
                    safe = false;
                    break;
                }

                var propertyElements = properties.Elements().ToArray();
                if (propertyElements.Length != 2
                    || propertyElements.Count(element => element.Name == math + "count") != 1
                    || propertyElements.Count(element => element.Name == math + "mcJc") != 1)
                {
                    safe = false;
                    break;
                }

                var count = properties.Element(math + "count");
                var justification = properties.Element(math + "mcJc");
                var countValue = (string?)count?.Attribute(math + "val");
                var justificationValue = (string?)justification?.Attribute(math + "val");
                if (count is null
                    || justification is null
                    || count.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != math + "val")
                    || justification.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != math + "val")
                    || count.HasElements
                    || justification.HasElements
                    || !string.IsNullOrWhiteSpace(count.Value)
                    || !string.IsNullOrWhiteSpace(justification.Value)
                    || !int.TryParse(countValue, NumberStyles.None, CultureInfo.InvariantCulture, out var repetitions)
                    || repetitions <= 0
                    || repetitions > 256
                    || string.IsNullOrWhiteSpace(justificationValue))
                {
                    safe = false;
                    break;
                }

                for (var index = 0; index < repetitions; index++)
                {
                    expanded.Add(
                        new XElement(
                            math + "mc",
                            new XElement(
                                math + "mcPr",
                                new XElement(
                                    math + "count",
                                    new XAttribute(math + "val", "1")),
                                new XElement(
                                    math + "mcJc",
                                    new XAttribute(math + "val", justificationValue)))));
                }
            }

            if (safe)
                columns.ReplaceNodes(expanded);
        }
    }

    private static void NormalizeImportedMathOnOffProperties(
        XDocument document,
        XNamespace math)
    {
        // OfficeMath CT_OnOff uses two independent defaults:
        //   1) when an element is present without m:val, that val means true;
        //   2) when the element itself is absent, the property's semantic default
        //      depends on the property (and for m:grow, on its parent object).
        // Word routinely removes properties that equal their semantic default while
        // materializing OMML. Canonicalize those representation-only omissions, but
        // keep every nondefault value so a real mathematical/layout change remains
        // a different content signature.
        var onOffNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "degHide", "subHide", "supHide", "grow", "nor", "lit", "aln", "diff", "noBreak",
            "opEmu", "transp", "zeroAsc", "zeroDesc", "zeroWid", "show", "plcHide",
            "hideTop", "hideBot", "hideLeft", "hideRight", "strikeH", "strikeV",
            "strikeBLTR", "strikeTLBR",
        };
        var defaultFalseNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "degHide", "subHide", "supHide", "nor", "lit", "aln", "diff", "noBreak",
            "opEmu", "transp", "zeroAsc", "zeroDesc", "zeroWid", "plcHide",
            "hideTop", "hideBot", "hideLeft", "hideRight", "strikeH", "strikeV",
            "strikeBLTR", "strikeTLBR",
        };

        foreach (var property in document.Descendants()
                     .Where(element => element.Name.Namespace == math
                         && onOffNames.Contains(element.Name.LocalName))
                     .ToArray())
        {
            var normalizedValue = ReadMathBooleanValue(property);
            property.SetAttributeValue(math + "val", normalizedValue ? "1" : "0");

            bool? semanticDefault = null;
            if (defaultFalseNames.Contains(property.Name.LocalName))
            {
                semanticDefault = false;
            }
            else if (property.Name == math + "show")
            {
                // Phantom content is shown when m:show is omitted.
                semanticDefault = true;
            }
            else if (property.Name == math + "grow")
            {
                // ISO/IEC 29500 gives m:grow a context-sensitive default:
                // delimiter objects grow by default, n-ary operators do not.
                if (property.Parent?.Name == math + "dPr")
                    semanticDefault = true;
                else if (property.Parent?.Name == math + "naryPr")
                    semanticDefault = false;
            }

            if (!semanticDefault.HasValue) continue;
            if (normalizedValue == semanticDefault.Value)
                property.Remove();
        }
    }

    private static void NormalizeImportedDefaultProperties(XDocument document, XNamespace math)
    {
        void NormalizeScalar(XElement? property, string elementDefault, string attributeDefault)
        {
            if (property is null || property.HasElements || !string.IsNullOrWhiteSpace(property.Value)
                || property.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != math + "val"))
                return;
            var value = (string?)property.Attribute(math + "val") ?? attributeDefault;
            if (value == elementDefault) property.Remove();
            else property.SetAttributeValue(math + "val", value);
        }
        // ISO 29500 distinguishes omission of the element from omission of val:
        // begChr/endChr absent = parentheses; present without val = NO delimiter.
        // sty has the same italic default in both cases. Preserve every explicit
        // nondefault value and unknown extension attribute.
        foreach (var delimiter in document.Descendants(math + "d"))
        {
            var properties = delimiter.Element(math + "dPr");
            NormalizeScalar(properties?.Element(math + "begChr"), "(", string.Empty);
            NormalizeScalar(properties?.Element(math + "endChr"), ")", string.Empty);
        }
        foreach (var run in document.Descendants(math + "r"))
            NormalizeScalar(run.Element(math + "rPr")?.Element(math + "sty"), "i", "i");
        foreach (var text in document.Descendants(math + "t"))
        {
            var space = text.Attribute(XNamespace.Xml + "space");
            // A preservation annotation on unpadded text changes no characters.
            // Keep real leading/trailing whitespace and every character untouched.
            if (space?.Value is "preserve" or "default"
                && text.Value.Length > 0
                && text.Value.Trim(' ', '\t', '\r', '\n') == text.Value)
                space.Remove();
        }
    }

    private static void NormalizeImportedPlainStyleDefaults(
        XDocument document,
        XNamespace math)
    {
        foreach (var run in document.Descendants(math + "r").ToArray())
        {
            var properties = run.Element(math + "rPr");
            var style = properties?.Element(math + "sty");
            if (style is null
                || !string.Equals(
                    (string?)style.Attribute(math + "val"),
                    "p",
                    StringComparison.Ordinal))
                continue;

            var text = string.Concat(run.Elements(math + "t").Select(item => item.Value));
            if (string.IsNullOrEmpty(text) || ContainsUnicodeLetter(text))
                continue;

            // Word commonly splits a mixed math run such as "E=m" into
            // E / = / m and adds an explicit plain style to the operator. It also
            // writes m:sty="p" on ordinary digits such as a superscript 2. For
            // characters with no Unicode letter semantics this is only an explicit
            // spelling of their normal upright math appearance, not a change of
            // formula content. Keep plain styling on any run containing a letter:
            // \mathrm{x}, upright Greek, script alphabets, etc. remain semantic.
            style.Remove();
            if (properties is not null
                && !properties.HasElements
                && !properties.HasAttributes)
                properties.Remove();
        }
    }

    internal static bool ContainsUnicodeLetter(string value)
    {
        for (var index = 0; index < value.Length;)
        {
            var category = CharUnicodeInfo.GetUnicodeCategory(value, index);
            if (category is UnicodeCategory.UppercaseLetter
                or UnicodeCategory.LowercaseLetter
                or UnicodeCategory.TitlecaseLetter
                or UnicodeCategory.ModifierLetter
                or UnicodeCategory.OtherLetter
                or UnicodeCategory.LetterNumber)
                return true;
            index += char.IsSurrogatePair(value, index) ? 2 : 1;
        }
        return false;
    }

    // Frozen v1 serialization behavior. New content comparisons must use the
    // control-preserving NormalizeImportedMathRunGrouping instead.
    private static void NormalizeMathRunGrouping(XDocument document, XNamespace math)
    {
        foreach (var run in document.Descendants(math + "r").ToArray())
        {
            var texts = run.Elements(math + "t").ToArray();
            if (texts.Length > 1)
            {
                var combined = string.Concat(texts.Select(text => text.Value));
                texts[0].Value = combined;
                foreach (var extra in texts.Skip(1)) extra.Remove();
            }
            var properties = run.Element(math + "rPr");
            if (properties is not null && !properties.HasElements && !properties.HasAttributes)
                properties.Remove();
        }

        foreach (var parent in document.Root?.DescendantsAndSelf().ToArray()
                     ?? Array.Empty<XElement>())
        {
            XElement? previousRun = null;
            string? previousKey = null;
            foreach (var child in parent.Elements().ToArray())
            {
                if (child.Name != math + "r")
                {
                    previousRun = null;
                    previousKey = null;
                    continue;
                }
                var key = child.Element(math + "rPr")?.ToString(SaveOptions.DisableFormatting)
                    ?? string.Empty;
                if (previousRun is not null
                    && string.Equals(previousKey, key, StringComparison.Ordinal))
                {
                    var previousText = previousRun.Element(math + "t");
                    var currentText = child.Element(math + "t");
                    if (previousText is not null && currentText is not null)
                    {
                        previousText.Value += currentText.Value;
                        child.Remove();
                        continue;
                    }
                }
                previousRun = child;
                previousKey = key;
            }
        }
    }
}
