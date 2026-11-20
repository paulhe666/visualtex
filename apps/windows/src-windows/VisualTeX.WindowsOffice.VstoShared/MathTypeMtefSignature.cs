using System.Globalization;
using System.Text;
using System.Xml.Linq;
using static VisualTeX.WordVsto.MathTypeMtefCharacters;
using static VisualTeX.WordVsto.MathTypeMtefFences;
using static VisualTeX.WordVsto.MathTypeMtefGlyphs;
using static VisualTeX.WordVsto.MathTypeMtefTemplates;

namespace VisualTeX.WordVsto;

/// <summary>
/// The MathML that semantic signatures compare: MathML as MathType stores and exports it, normalized (stored
/// glyphs, prescripts, tab piles), and its canonical form.
/// </summary>
internal static class MathTypeMtefSignature
{
    // Glyphs whose MathType form differs from the MathML token but not in
    // meaning: a SimSun substitute shape (U+25B4 is stored as U+25B2), and CJK
    // or full-width characters, which MathType keeps only in its text face
    // (typeface 12) whatever their MathML token class.
    internal static void NormalizeMathTypeStoredGlyphs(XElement root)
    {
        MergeMathTypeAlternates(root, keepNativeTokens: false);
        // MathType has one hbar glyph per weight (MT Extra's, slanted by design):
        // italic, upright or unstyled hbar is that glyph, bold-italic its bold.
        foreach (var hbar in root.DescendantsAndSelf().Where(e => e.Name.LocalName == "mi"
                     && e.Value.Trim() == "ℏ").ToArray())
        {
            var variant = MathMlTokenStyleContract.ResolveDeclaredVariant(hbar);
            if (variant is "italic" or "normal") hbar.Attribute("mathvariant")?.Remove();
            else if (variant == "bold-italic") hbar.SetAttributeValue("mathvariant", "bold");
        }
        foreach (var token in root.DescendantsAndSelf().Where(e => !e.HasElements
                     && e.Name.LocalName is "mi" or "mo" or "mtext").ToArray())
        {
            if (token.Value.Any(FarEastGlyphSubstitutes.ContainsKey))
                token.Value = new string(token.Value
                    .Select(c => FarEastGlyphSubstitutes.TryGetValue(c, out var covered) ? covered : c).ToArray());
            // A no-break space of text is stored as MathType's text space U+0020.
            if (token.Name.LocalName == "mtext" && token.Value.Contains(' '))
                token.Value = token.Value.Replace(' ', ' ');
            var value = token.Value.Trim();
            if (token.Name.LocalName != "mtext" && value.Length == 1 && IsFarEastTextScalar(value[0]))
            {
                token.Name = token.Name.Namespace + "mtext";
                token.Attribute("mathvariant")?.Remove();
            }
            else if (token.Name.LocalName == "mo" && value.Length == 1 && char.IsLetter(value[0])
                // a modifier letter (U+02C7 caron of \check) is an accent mark
                && CharUnicodeInfo.GetUnicodeCategory(value[0]) != UnicodeCategory.ModifierLetter)
            {
                // A letter is an identifier whichever token carries it; MathType
                // reads it back as an upright mi.
                token.Name = token.Name.Namespace + "mi";
                if (token.Attribute("mathvariant") is null) token.SetAttributeValue("mathvariant", "normal");
            }
            else if (token.Name.LocalName == "mtext" && value.Length == 1
                && value[0] is '∋' or '∣' or '◯' or '♡')
                token.Name = token.Name.Namespace + "mo";
        }
        // MathType keeps consecutive text characters as one run, so CJK
        // punctuation written as separate tokens reads back as one mtext.
        // Merge equally attributed neighbours within a sequence only; fraction
        // and script arguments stay separate.
        foreach (var sequence in root.DescendantsAndSelf().Where(e => e.Name.LocalName
                     is "math" or "mrow" or "mstyle" or "mtd" or "msqrt" or "mphantom" or "mpadded" or "menclose").ToArray())
        {
            XElement? previous = null;
            foreach (var child in sequence.Elements().ToArray())
            {
                if (previous is not null && child.Name.LocalName == "mtext" && previous.Name.LocalName == "mtext"
                    && !child.HasElements && !previous.HasElements
                    && string.Join("|", child.Attributes().Select(a => a.ToString()).OrderBy(a => a, StringComparer.Ordinal))
                        == string.Join("|", previous.Attributes().Select(a => a.ToString()).OrderBy(a => a, StringComparer.Ordinal)))
                {
                    previous.Value += child.Value;
                    child.Remove();
                    continue;
                }
                previous = child;
            }
        }
    }

    internal static void NormalizeMathMlPrescripts(XElement root)
    {
        var scriptNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "msub", "msup", "msubsup",
        };
        // Only containers whose children form one visual expression may donate
        // the following atom as a base. Fraction, root and script children are
        // separate arguments, so crossing one of those boundaries is invalid.
        var sequenceNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "math", "mrow", "mstyle", "mtd", "msqrt", "mphantom",
            "mpadded", "menclose",
        };
        var baseNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "mi", "mn", "mrow", "mstyle", "mfrac", "msqrt", "mroot",
            "msub", "msup", "msubsup", "mmultiscripts", "mover", "munder",
            "munderover", "mfenced",
        };

        foreach (var script in root.DescendantsAndSelf()
                     .Where(element => scriptNames.Contains(element.Name.LocalName))
                     .Reverse()
                     .ToArray())
        {
            if (script.Parent is not XElement parent
                || !sequenceNames.Contains(parent.Name.LocalName))
                continue;
            var arguments = script.Elements().ToArray();
            if (arguments.Length != (script.Name.LocalName == "msubsup" ? 3 : 2)
                || !IsEmptyScriptBase(arguments[0]))
                continue;
            var basis = script.ElementsAfterSelf().FirstOrDefault();
            if (basis is null
                || !baseNames.Contains(basis.Name.LocalName)
                || string.IsNullOrWhiteSpace(basis.Value))
                continue;

            var ns = script.Name.Namespace;
            var sub = script.Name.LocalName == "msup"
                ? new XElement(ns + "none")
                : new XElement(arguments[1]);
            var sup = script.Name.LocalName == "msub"
                ? new XElement(ns + "none")
                : new XElement(arguments[script.Name.LocalName == "msubsup" ? 2 : 1]);
            if (string.IsNullOrWhiteSpace(sub.Value)
                && string.IsNullOrWhiteSpace(sup.Value))
                continue;

            // TeX writes a left script as an empty-base script followed by its
            // visible atom (for example {}^{59}_{27}\mathrm{Co}). A null LINE
            // prevents accidental binding to the previous sibling, but it still
            // gives MathType an independent empty base whose vertical metrics can
            // differ at the start of a formula. MTEF has a real preceding-script
            // template; express the source as mmultiscripts so all occurrences
            // attach to, and align against, their actual base atom.
            var native = new XElement(
                ns + "mmultiscripts",
                script.Attributes()
                    .Where(attribute => !attribute.IsNamespaceDeclaration)
                    .Select(attribute => new XAttribute(attribute)),
                new XElement(basis),
                new XElement(ns + "mprescripts"),
                sub,
                sup);
            script.ReplaceWith(native);
            basis.Remove();
        }
    }

    internal static void NormalizeMathTypeExportedTabPiles(XElement root)
    {
        foreach (var table in root.DescendantsAndSelf()
                     .Where(element => element.Name.LocalName == "mtable")
                     .ToArray())
        {
            var sourceRows = table.Elements()
                .Where(row => row.Name.LocalName is "mtr" or "mlabeledtr")
                .ToArray();
            if (sourceRows.Length == 0) continue;

            var allowsEmptyComTabMarkers =
                IsMathTypeComExportedTabPile(table, sourceRows);
            var rebuiltRows = new List<XElement>();
            var sawTab = false;
            var supportedShape = true;
            foreach (var sourceRow in sourceRows)
            {
                var sourceCells = sourceRow.Elements()
                    .Where(cell => cell.Name.LocalName == "mtd")
                    .ToArray();
                if (sourceCells.Length != 1)
                {
                    supportedShape = false;
                    break;
                }

                if (sourceCells[0].Elements().Any(element =>
                        element.Name.LocalName is "maligngroup" or "malignmark"))
                {
                    var markedRow = RebuildMathTypeAlignmentMarkedRow(
                        sourceCells[0],
                        out var markedSawTab);
                    sawTab |= markedSawTab;
                    rebuiltRows.Add(markedRow);
                    continue;
                }

                var rebuiltRow = new XElement("mtr");
                var rebuiltCell = new XElement("mtd");
                foreach (var node in sourceCells[0].Nodes())
                {
                    if (node is XElement marker
                        && IsMathTypeExportedTabMarker(
                            marker,
                            allowsEmptyComTabMarkers))
                    {
                        sawTab = true;
                        NormalizeMathTypeExportedTabCell(rebuiltCell);
                        rebuiltRow.Add(rebuiltCell);
                        rebuiltCell = new XElement("mtd");
                        continue;
                    }
                    if (node is XText whitespace && string.IsNullOrWhiteSpace(whitespace.Value))
                        continue;
                    rebuiltCell.Add(CloneNode(node));
                }
                NormalizeMathTypeExportedTabCell(rebuiltCell);
                rebuiltRow.Add(rebuiltCell);
                rebuiltRows.Add(rebuiltRow);
            }
            if (!supportedShape || !sawTab) continue;

            while (rebuiltRows.Count > 1
                && !MathTypeTabRowHasMeaningfulContent(rebuiltRows[rebuiltRows.Count - 1]))
                rebuiltRows.RemoveAt(rebuiltRows.Count - 1);

            table.RemoveNodes();
            foreach (var row in rebuiltRows) table.Add(row);
            var columnCount = rebuiltRows
                .Select(row => row.Elements().Count(cell => cell.Name.LocalName == "mtd"))
                .DefaultIfEmpty(0)
                .Max();
            if (columnCount > 0)
            {
                table.SetAttributeValue(
                    "columnalign",
                    string.Join(" ", Enumerable.Range(0, columnCount)
                        .Select(column => (column & 1) == 0 ? "right" : "left")));
            }
        }
    }

    private static XElement RebuildMathTypeAlignmentMarkedRow(
        XElement sourceCell,
        out bool sawTab)
    {
        sawTab = false;
        var row = new XElement("mtr");
        var cell = new XElement("mtd");
        var emittedAnyCell = false;

        bool CellHasContent() => cell.Nodes().Any(node =>
            node is XElement
            || node is XText text && !string.IsNullOrWhiteSpace(text.Value));

        void EmitCell()
        {
            NormalizeMathTypeExportedTabCell(cell);
            row.Add(cell);
            cell = new XElement("mtd");
            emittedAnyCell = true;
        }

        foreach (var node in sourceCell.Nodes())
        {
            if (node is XText whitespace && string.IsNullOrWhiteSpace(whitespace.Value))
                continue;
            if (node is XElement element)
            {
                if (element.Name.LocalName == "maligngroup")
                    continue;
                if (element.Name.LocalName == "malignmark")
                {
                    EmitCell();
                    continue;
                }
                if (IsMathTypeExportedTabMarker(element, allowEmptyComMarker: false))
                {
                    sawTab = true;
                    // MathType emits one leading U+0009 after <maligngroup/> to
                    // enter the first ruler group. It is positioning metadata,
                    // not an empty TeX column. Later tabs terminate the current
                    // right-hand cell and start the next right/left pair.
                    if (!emittedAnyCell && !CellHasContent())
                        continue;
                    EmitCell();
                    continue;
                }
            }
            cell.Add(CloneNode(node));
        }
        EmitCell();
        return row;
    }

    private static bool IsMathTypeComExportedTabPile(
        XElement table,
        XElement[] rows)
    {
        var columnAlign = ((string?)table.Attribute("columnalign") ?? string.Empty)
            .Trim();
        if (!string.Equals(columnAlign, "left", StringComparison.OrdinalIgnoreCase)
            || rows.Length < 2
            || MathTypeTabRowHasMeaningfulContent(rows[rows.Length - 1]))
            return false;

        var nonTrailingRows = rows.Take(rows.Length - 1).ToArray();
        if (nonTrailingRows.Any(row => row.Elements()
                .Count(cell => cell.Name.LocalName == "mtd") != 1))
            return false;
        return nonTrailingRows.Any(row => row.Descendants()
            .Any(element => element.Name.LocalName == "mtext"
                && !element.HasElements
                && element.Value.Length == 0));
    }

    private static bool IsMathTypeExportedTabMarker(
        XElement element,
        bool allowEmptyComMarker)
    {
        if (element.Name.LocalName is not ("mtext" or "mi")) return false;
        var value = element.Value;
        if (value.Length > 0 && value.All(character => character == '\t'))
            return true;
        return allowEmptyComMarker
            && element.Name.LocalName == "mtext"
            && !element.HasElements
            && value.Length == 0;
    }

    private static bool MathTypeTabRowHasMeaningfulContent(XElement row)
    {
        foreach (var element in row.Descendants())
        {
            if (element.Name.LocalName is "mrow" or "mstyle" or "mpadded" or "mtd")
                continue;
            if (!string.IsNullOrWhiteSpace(element.Value.Replace("\t", string.Empty)))
                return true;
            if (element.HasElements) return true;
        }
        return false;
    }

    private static void NormalizeMathTypeExportedTabCell(XElement cell)
    {
        var elements = cell.Elements().ToList();
        for (var index = 0; index < elements.Count; index++)
        {
            var first = elements[index];
            if (first.Name.LocalName != "mtext" || string.IsNullOrWhiteSpace(first.Value))
                continue;
            var mergedText = first.Value;
            var cursor = index + 1;
            var consumed = new List<XElement>();
            while (cursor + 1 < elements.Count
                && IsMathTypeExportedTextSpace(elements[cursor])
                && elements[cursor + 1].Name.LocalName == "mtext"
                && !string.IsNullOrWhiteSpace(elements[cursor + 1].Value))
            {
                mergedText += " " + elements[cursor + 1].Value;
                consumed.Add(elements[cursor]);
                consumed.Add(elements[cursor + 1]);
                cursor += 2;
            }
            if (consumed.Count == 0) continue;
            first.Value = mergedText;
            foreach (var item in consumed) item.Remove();
            elements = cell.Elements().ToList();
        }

        // MathType can terminate an ordinary text run before punctuation even
        // though the source was one \text{...} node (for example it exports
        // "High word." as mtext("High word") + mtext(".")). Rejoin only an
        // immediately adjacent punctuation-only mtext fragment inside this
        // already-confirmed tab-aligned export shape.
        elements = cell.Elements().ToList();
        for (var index = 0; index + 1 < elements.Count; index++)
        {
            var current = elements[index];
            var next = elements[index + 1];
            if (current.Name.LocalName != "mtext"
                || next.Name.LocalName != "mtext"
                || current.HasElements
                || next.HasElements
                || next.Value.Length == 0
                || !next.Value.All(char.IsPunctuation))
                continue;
            current.Value += next.Value;
            next.Remove();
            elements = cell.Elements().ToList();
            index--;
        }

        elements = cell.Elements().ToList();
        for (var index = 0; index < elements.Count; index++)
        {
            if (!IsSingleUnstyledLatinIdentifier(elements[index])) continue;
            var end = index;
            var name = new StringBuilder();
            while (end < elements.Count && IsSingleUnstyledLatinIdentifier(elements[end]))
            {
                name.Append(elements[end].Value);
                end++;
            }
            if (name.Length < 2
                || end >= elements.Count
                || elements[end].Name.LocalName != "mo"
                || elements[end].Value.Trim() != "(")
            {
                index = Math.Max(index, end - 1);
                continue;
            }

            var function = new XElement(
                "mi",
                new XAttribute("mathvariant", "normal"),
                name.ToString());
            elements[index].ReplaceWith(function);
            for (var remove = index + 1; remove < end; remove++)
                elements[remove].Remove();
            elements = cell.Elements().ToList();
        }
    }

    private static bool IsMathTypeExportedTextSpace(XElement element)
    {
        if (element.HasElements) return false;
        if (element.Name.LocalName == "mi" && element.Value.Length == 0)
        {
            // MathType's IDataObject MathML exporter represents a word-space
            // inside text in a native tab-pile as an empty <mi/> rather than
            // mspace/whitespace text. This helper is used only after the table
            // has already been proven to be a MathType tab-pile export.
            return true;
        }
        return element.Name.LocalName is "mi" or "mtext"
            && element.Value.Length > 0
            && element.Value.All(char.IsWhiteSpace);
    }

    private static bool IsSingleUnstyledLatinIdentifier(XElement element)
    {
        if (element.Name.LocalName != "mi"
            || element.Attribute("mathvariant") is not null)
            return false;
        var value = element.Value;
        return value.Length == 1 && value[0] <= 0x7F && char.IsLetter(value[0]);
    }

    internal static string CanonicalizeMathMl(XElement element, string? inheritedMathVariant)
    {
        var local = element.Name.LocalName;
        var variant = MathMlTokenStyleContract.ResolveDeclaredVariant(element) ?? inheritedMathVariant;
        string Children(string? childVariant = null) =>
            CanonicalizeElementSequence(
                element.Elements()
                    .Where(child => child.Name.LocalName is not ("annotation" or "annotation-xml"))
                    .ToArray(),
                childVariant ?? variant);
        var children = element.Elements()
            .Where(child => child.Name.LocalName is not ("annotation" or "annotation-xml"))
            .ToArray();
        switch (local)
        {
            case "math":
            case "mpadded":
                return Children();
            case "mrow":
            {
                if (TryCanonicalizeMathJaxFencedSuperscriptRow(element, variant, out var fencedSuperscript))
                    return fencedSuperscript;
                if (TryGetMathJaxFenceRow(element, out var open, out var close, out var fenceChildren))
                {
                    if (TryCanonicalizeMathJaxBinomialFence(
                            open,
                            close,
                            fenceChildren,
                            variant,
                            out var binomial))
                        return binomial;
                    var normalizedOpen = NormalizeFence(open);
                    var normalizedClose = NormalizeFence(close);
                    var fencedChildrenSignature =
                        CanonicalizeElementSequence(fenceChildren, variant);
                    return CanonicalFenceSignature(
                        normalizedOpen,
                        normalizedClose,
                        fencedChildrenSignature);
                }
                return Children();
            }
            case "mstyle":
                return Children(variant);
            case "semantics":
            case "maction":
                return children.Length == 0
                    ? string.Empty
                    : CanonicalizeMathMl(children[0], variant);
            case "annotation":
            case "annotation-xml":
            case "none":
            case "mspace":
                return string.Empty;
            case "mi":
            {
                var value = element.Value.Trim();
                if (value.Length == 0) return string.Empty;
                var primeSignature = CanonicalPrimeSignature(value);
                if (primeSignature is not null) return primeSignature;
                if ((value is "ℵ" or "ℜ" or "ℑ")
                    && string.Equals(variant, "normal", StringComparison.OrdinalIgnoreCase))
                {
                    // Production MathJax marks these fixed letterlike symbols as
                    // upright, while MathType's direct MTEF readback omits an
                    // explicit mathvariant. The glyph identity already fixes their
                    // presentation, so the attribute is semantically redundant.
                    return CanonicalToken("mi", value, string.Empty);
                }
                // MathType 7 sometimes exports standalone mathematical glyphs
                // such as infinity as <mi> even though MathJax/VisualTeX uses
                // <mo>.  That token-tag difference is not a semantic change.
                // Keep mtext distinct so an actual Symbol -> Text regression is
                // still rejected by the round-trip validator.
                if (IsMathematicalSymbolToken(value))
                    return CanonicalStyledOperator(value,
                        MathMlTokenStyleContract.ResolveEffectiveVariant(element, variant));
                // Intrinsic Unicode alphabet variants fold before MathML's
                // implicit single-mi default; a declaration remains authoritative.
                if (string.IsNullOrWhiteSpace(variant) && value.Length == 1
                    && TryNormalizeLetterlikeForWrite(value[0], out var letter, out var letterVariant))
                    return CanonicalToken("mi", ((char)letter).ToString(), letterVariant);
                return CanonicalToken("mi", value,
                    MathMlTokenStyleContract.ResolveEffectiveVariant(element, variant));
            }
            case "mn":
            {
                // Numbers default to upright, not to identifier italic. Their
                // explicit variant is part of the contract too: ignoring it
                // makes a bold/monospace digit loss invisible to validation.
                var numberVariant = (variant ?? string.Empty).Trim().ToLowerInvariant();
                var numberPrefix = numberVariant.Length == 0 || numberVariant is "normal" or "upright"
                    ? "n("
                    : "n[" + numberVariant + "](";
                var number = element.Value.Trim();
                if (number.Length == 0) return numberPrefix + ")";
                // MTEF stores a number as one CHAR record per digit and its
                // read-back groups adjacent digits into one <mn>, while MathJax
                // keeps digits written apart ("1 0 0", as formula OCR emits
                // them) as separate <mn> tokens. The same digits in the same
                // style are the same number either way: compare them one
                // character at a time, as CanonicalToken does for letter runs.
                var digits = new StringBuilder();
                for (var index = 0; index < number.Length; index++)
                {
                    var length = char.IsSurrogatePair(number, index) ? 2 : 1;
                    digits.Append(numberPrefix).Append(number, index, length).Append(')');
                    index += length - 1;
                }
                return digits.ToString();
            }
            case "mo":
            {
                var value = element.Value.Trim();
                // MathJax represents the invisible delimiters from \\left. and
                // \\right. as empty <mo> nodes. MathType/MTEF correctly stores no
                // character for them. They are presentation scaffolding only and
                // must not make semantic round-trip validation fail.
                if (value.Length == 0) return string.Empty;
                // MathJax inserts U+2061 FUNCTION APPLICATION between a named
                // function and its argument. It has no visible glyph and should
                // never become a literal MathType character or a fake '?'.
                if (value == "⁡") return string.Empty;
                var primeSignature = CanonicalPrimeSignature(value);
                if (primeSignature is not null) return primeSignature;
                if (value.Length > 1 && value.All(char.IsLetter))
                    return CanonicalToken("mi", value, MathMlTokenStyleContract.ResolveEffectiveVariant(element, variant));
                // A multi-word name (MathJax's \liminf is "lim inf") is stored as
                // one function run per word with a native space between them.
                var words = value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                if (words.Length > 1 && words.All(word => word.All(char.IsLetter)))
                    return string.Concat(words.Select(word => CanonicalToken("mi",
                        word, MathMlTokenStyleContract.ResolveEffectiveVariant(element, variant))));
                // CHAR storage splits a pure operator run into scalar glyphs.
                // Use the same run grammar as font binding, retaining order,
                // every glyph's variant and the distinction from prose/functions.
                if (IsScalarOperatorRun(value))
                {
                    return string.Concat(EnumerateBmpScalars(value).Select(scalar =>
                        CanonicalStyledOperator(char.ConvertFromUtf32(scalar), variant)));
                }
                return CanonicalStyledOperator(value, variant);
            }
            case "mtext":
            {
                var value = element.Value.Trim();
                // MathJax may serialize TeX spacing commands as whitespace-only
                // <mtext> (commonly NBSP), while MathType normalizes the same
                // presentation-only gap to <mspace> on MTEF read-back. Neither
                // carries mathematical semantics, so treating empty/whitespace
                // mtext as text() makes an otherwise exact equation fail the
                // standalone-MTEF round-trip gate. Keep real textual content
                // strict, but canonicalize pure spacing exactly like mspace.
                if (value.Length == 0) return string.Empty;
                if (value.All(char.IsLetter))
                    return CanonicalToken("mi", value, variant ?? "normal");
                var textVariant = MathMlTokenStyleContract.ResolveEffectiveVariant(element, variant);
                return "text" + (textVariant is null or "normal" or "upright" ? string.Empty : "[" + textVariant + "]")
                    + "(" + value + ")";
            }
            case "mfrac":
            {
                var fractionKind = CanonicalFractionKind(element);
                if (fractionKind == "nobarfrac" && children.Length == 2)
                    return CanonicalBinary("stack", children, variant);
                return CanonicalBinary(fractionKind, children, variant);
            }
            case "msqrt":
                return "sqrt(" + Children() + ")";
            case "mroot":
                return CanonicalBinary("root", children, variant);
            case "msub":
                return CanonicalBinary("sub", children, variant);
            case "msup":
                return CanonicalBinary("sup", children, variant);
            case "msubsup":
            {
                if (children.Length < 3)
                    return string.Concat(children.Select(child => CanonicalizeMathMl(child, variant)));
                var baseSignature = CanonicalizeScriptBase(children[0], variant);
                var subSignature = CanonicalizeMathMl(children[1], variant);
                var supSignature = CanonicalizeMathMl(children[2], variant);
                if (subSignature.Length == 0 && supSignature.Length > 0)
                    return "sup(" + baseSignature + "," + supSignature + ")";
                if (supSignature.Length == 0 && subSignature.Length > 0)
                    return "sub(" + baseSignature + "," + subSignature + ")";
                if (subSignature.Length == 0 && supSignature.Length == 0)
                    return baseSignature;
                return "subsup(" + baseSignature + "," + subSignature + "," + supSignature + ")";
            }
            case "munder":
            {
                if (children.Length == 1)
                {
                    if (TryCanonicalizeWrappedHorizontalBrace(
                        children[0], variant, top: false, out var wrappedBrace))
                        return wrappedBrace;
                    return CanonicalizeMathMl(children[0], variant);
                }
                if (children.Length >= 2)
                {
                    var under = children[1].Value.Trim();
                    if (under is "_" or "\u00AF" or "\u203E" or "\u2015" or "\u02C9")
                        return "underbar(" + CanonicalizeMathMl(children[0], variant) + ")";
                    if (IsHorizontalBraceMarker(under, top: false))
                        return "hbrace-bottom(" + CanonicalizeMathMl(children[0], variant) + ")";
                }
                return CanonicalBinary("sub", children, variant);
            }
            case "mover":
            {
                if (children.Length == 1)
                {
                    if (TryCanonicalizeWrappedHorizontalBrace(
                        children[0], variant, top: true, out var wrappedBrace))
                        return wrappedBrace;
                    return CanonicalizeMathMl(children[0], variant);
                }
                if (children.Length < 2) return Children();
                var over = NormalizeAccentMark(children[1].Value.Trim());
                // MathType has no U+23DC arc; it stores \overparen as U+2322.
                if (over == "⏜") over = "⌢";
                if (IsHorizontalBraceMarker(over, top: true))
                    return "hbrace-top(" + CanonicalizeMathMl(children[0], variant) + ")";
                var isAccent = ResolveOverAccent(element, children[1]);
                // MathType's overbar template is also the representation used
                // for TeX overline (whose MathML may explicitly set accent=false).
                if (over == "¯")
                    return "accent(overbar," + CanonicalizeMathMl(children[0], variant) + ")";
                var accentGlyph = over is "→" or "←" or "↔"
                    or "^" or "~" or "." or "˙" or "¨"
                    or "⌢" or "⏜" or "ˇ" or "˘" or "´" or "`" or "˚";
                return isAccent && accentGlyph
                    ? "accent(" + over + "," + CanonicalizeMathMl(children[0], variant) + ")"
                    : CanonicalBinary("sup", children, variant);
            }
            case "munderover":
            {
                // as msubsup: an empty limit is no limit (a big operator with an
                // upper limit only is written with an empty lower slot)
                if (children.Length >= 3)
                {
                    var underSignature = CanonicalizeMathMl(children[1], variant);
                    var overSignature = CanonicalizeMathMl(children[2], variant);
                    if (underSignature.Length == 0 && overSignature.Length > 0)
                        return "sup(" + CanonicalizeScriptBase(children[0], variant) + "," + overSignature + ")";
                    if (overSignature.Length == 0 && underSignature.Length > 0)
                        return "sub(" + CanonicalizeScriptBase(children[0], variant) + "," + underSignature + ")";
                }
                return CanonicalTernary("subsup", children, variant);
            }
            case "mfenced":
            {
                if (TryCanonicalizeMathTypeBinomialFence(element, variant, out var binomial))
                    return binomial;
                var open = NormalizeFence((string?)element.Attribute("open") ?? "(");
                var close = NormalizeFence((string?)element.Attribute("close") ?? ")");
                var fencedChildren = CanonicalizeMfencedChildren(element, children, variant);
                // Word OMML materializes a one-sided system/cases delimiter as
                // a real delimiter object (m:d / MathML mfenced with an empty
                // opposite fence), while MathJax can represent the same visible
                // mathematics as a loose <mo> delimiter followed by the table.
                // Canonicalize only one-sided fences to that explicit token form;
                // paired fences remain structurally strict.
                return CanonicalFenceSignature(
                    open,
                    close,
                    fencedChildren);
            }
            case "mtable":
            {
                // EmitFraction writes every zero-rule fraction as a centered
                // two-line PILE, because MTEF has no rule-less fraction, and a
                // rewritten native PILE can come back as a one-column MATRIX.
                // A zero-rule fraction and any two-row, one-column table are
                // therefore one "stack" wherever they appear, not only inside a
                // recognized binomial fence: Word OMML writes \binom with fixed
                // literal parentheses, which no fence rule accepts.
                if (TryCanonicalizeTwoLineStack(element, variant, out var stacked))
                    return stacked;
                return "table(" + string.Join(";", children.Select(row =>
                    CanonicalizeMathMl(row, variant))) + ")";
            }
            case "mtr":
            case "mlabeledtr":
                return "row(" + string.Join(",", children.Select(cell =>
                    CanonicalizeMathMl(cell, variant))) + ")";
            case "mtd":
                return "cell(" + Children() + ")";
            case "menclose":
            {
                var notation = new HashSet<string>(
                    ((string?)element.Attribute("notation") ?? string.Empty)
                        .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
                        .Select(value => value.Trim().ToLowerInvariant()),
                    StringComparer.Ordinal);
                if (notation.Contains("radical") && notation.Count == 1)
                    return "sqrt(" + Children() + ")";

                var borders = new List<string>();
                if (notation.Contains("box"))
                {
                    borders.AddRange(new[] { "top", "bottom", "left", "right" });
                }
                else
                {
                    if (notation.Contains("top")) borders.Add("top");
                    if (notation.Contains("bottom")) borders.Add("bottom");
                    if (notation.Contains("left")) borders.Add("left");
                    if (notation.Contains("right")) borders.Add("right");
                }

                var strikeParts = new List<string>();
                if (notation.Contains("horizontalstrike")) strikeParts.Add("h");
                if (notation.Contains("verticalstrike")) strikeParts.Add("v");
                if (notation.Contains("updiagonalstrike")) strikeParts.Add("up");
                if (notation.Contains("downdiagonalstrike")) strikeParts.Add("down");

                var otherVisibleNotations = notation
                    .Where(value => value is not (
                        "box" or "top" or "bottom" or "left" or "right"
                        or "horizontalstrike" or "verticalstrike"
                        or "updiagonalstrike" or "downdiagonalstrike"))
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .ToArray();
                var content = Children();
                if (borders.Count == 4 && strikeParts.Count == 0
                    && otherVisibleNotations.Length == 0)
                    return "box(" + content + ")";
                if (borders.Count == 0 && strikeParts.Count > 0
                    && otherVisibleNotations.Length == 0)
                    return "strike[" + string.Join("+", strikeParts) + "](" + content + ")";
                if (borders.Count > 0 || strikeParts.Count > 0 || otherVisibleNotations.Length > 0)
                {
                    return "enclose[borders=" + string.Join("+", borders)
                        + ";strikes=" + string.Join("+", strikeParts)
                        + ";other=" + string.Join("+", otherVisibleNotations)
                        + "](" + content + ")";
                }
                return content;
            }
            case "mphantom":
                // MathType has no invisible content: a phantom is only its room,
                // which EmitPhantomSpace writes as native spaces (not compared).
                return string.Empty;
            case "mmultiscripts":
            {
                if (children.Length == 0) return string.Empty;
                var result = CanonicalizeScriptBase(children[0], variant);
                var precedes = false;
                for (var index = 1; index < children.Length;)
                {
                    if (children[index].Name.LocalName == "mprescripts")
                    {
                        if (precedes) throw new InvalidDataException("Duplicate MathML prescript separator.");
                        precedes = true; index++; continue;
                    }
                    if (index + 1 >= children.Length || children[index + 1].Name.LocalName == "mprescripts")
                        throw new InvalidDataException("MathML multiscripts require complete sub/sup pairs.");
                    var subSignature = children[index].Name.LocalName == "none" ? string.Empty : CanonicalizeMathMl(children[index], variant);
                    var supSignature = children[index + 1].Name.LocalName == "none" ? string.Empty : CanonicalizeMathMl(children[index + 1], variant);
                    if (subSignature.Length > 0 || supSignature.Length > 0)
                        result = precedes ? "pre(" + result + "," + subSignature + "," + supSignature + ")"
                            : subSignature.Length > 0 && supSignature.Length > 0 ? "subsup(" + result + "," + subSignature + "," + supSignature + ")"
                            : subSignature.Length > 0 ? "sub(" + result + "," + subSignature + ")"
                            : "sup(" + result + "," + supSignature + ")";
                    index += 2;
                }
                return result;
            }
            default:
                return children.Length > 0
                    ? Children()
                    : element.Value.Trim();
        }
    }

    private static string? CanonicalPrimeSignature(string value) => value switch
    {
        "'" => "o(′)",
        "′" => "o(′)",
        "″" => "o(′)o(′)",
        "‴" => "o(′)o(′)o(′)",
        "⁗" => "o(′)o(′)o(′)o(′)",
        _ => null,
    };

    private static bool IsReplacementGlyphOnly(string value) =>
        value.Length > 0 && value.All(character => character == '\uFFFD');

    private static string CanonicalStyledOperator(string value, string? variant)
    {
        var style=(variant ?? string.Empty).Trim().ToLowerInvariant();
        return "o" + (style.Length==0 || style is "normal" or "upright" ? string.Empty : "["+style+"]")
            + "("+NormalizeOperatorToken(value)+")";
    }

    internal static bool IsMathematicalSymbolToken(string value) =>
        value.Length > 0 && value.All(character => IsMathematicalSymbolScalar(character));

    internal static bool IsScalarOperatorRun(string value) =>
        value.Length > 1 && value.All(character =>
            IsMathematicalSymbolScalar(character) || character is ':' or '-');

    internal static bool IsHorizontalBraceMarker(string value, bool top) =>
        value == (top ? "⏞" : "⏟")
        || value == (top ? "\uFE37" : "\uFE38")
        || IsReplacementGlyphOnly(value);

    private static bool TryCanonicalizeWrappedHorizontalBrace(
        XElement candidate,
        string? variant,
        bool top,
        out string signature)
    {
        signature = string.Empty;
        if (candidate.Name.LocalName != (top ? "mover" : "munder")) return false;
        var inner = candidate.Elements().ToArray();
        if (inner.Length < 2 || inner[1].Name.LocalName != "mo") return false;
        var marker = inner[1].Value.Trim();
        var stretchy = string.Equals(
            (string?)inner[1].Attribute("stretchy"),
            "true",
            StringComparison.OrdinalIgnoreCase);
        // MathType 7 can export its private horizontal-brace glyph as replacement
        // characters. The extra one-child mover/munder wrapper plus a stretchy
        // operator is the stable structural signal; ordinary overset/underset
        // expressions do not have this wrapper.
        if (!stretchy && !IsHorizontalBraceMarker(marker, top))
            return false;
        signature = (top ? "hbrace-top(" : "hbrace-bottom(")
            + CanonicalizeMathMl(inner[0], variant) + ")";
        return true;
    }

    internal static bool IsTrueMathMlBoolean(string? value) =>
        string.Equals(value?.Trim(), "true", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value?.Trim(), "1", StringComparison.Ordinal);

    internal static bool IsFalseMathMlBoolean(string? value) =>
        string.Equals(value?.Trim(), "false", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value?.Trim(), "0", StringComparison.Ordinal);

    private static string CanonicalFractionKind(XElement fraction)
    {
        var lineThickness = ((string?)fraction.Attribute("linethickness") ?? string.Empty).Trim();
        if (IsZeroMathMlLength(lineThickness))
            return "nobarfrac";

        var bevelled = ((string?)fraction.Attribute("bevelled") ?? string.Empty).Trim();
        if (string.Equals(bevelled, "true", StringComparison.OrdinalIgnoreCase)
            || string.Equals(bevelled, "1", StringComparison.Ordinal))
            return "bevelledfrac";

        return "frac";
    }

    internal static bool IsZeroMathMlLength(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        var trimmed = value.Trim();
        var index = 0;
        if (trimmed[index] is '+' or '-') index++;
        var sawDigit = false;
        while (index < trimmed.Length
               && (char.IsDigit(trimmed[index]) || trimmed[index] == '.'))
        {
            sawDigit |= char.IsDigit(trimmed[index]);
            index++;
        }
        if (!sawDigit) return false;
        if (!double.TryParse(
                trimmed.Substring(0, index),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out var numeric))
            return false;
        return numeric == 0d;
    }

    private static string CanonicalizeMfencedChildren(
        XElement fenced,
        XElement[] children,
        string? variant)
    {
        if (children.Length == 0) return string.Empty;
        if (children.Length == 1)
            return CanonicalizeMathMl(children[0], variant);

        var rawSeparators = fenced.Attribute("separators") is null
            ? ","
            : ((string?)fenced.Attribute("separators") ?? string.Empty);
        var separators = rawSeparators
            .Where(character => !char.IsWhiteSpace(character))
            .Select(character => NormalizeMfencedSeparator(character.ToString()))
            .Where(separator => separator.Length > 0)
            .ToArray();

        var signature = new StringBuilder();
        for (var index = 0; index < children.Length; index++)
        {
            if (index > 0 && separators.Length > 0)
            {
                var separator = separators[Math.Min(index - 1, separators.Length - 1)];
                signature.Append("o(");
                signature.Append(separator);
                signature.Append(')');
            }
            signature.Append(CanonicalizeMathMl(children[index], variant));
        }
        return signature.ToString();
    }

    private static string NormalizeMfencedSeparator(string value) => value switch
    {
        // Word canonicalizes the ordinary set-builder vertical separator into
        // U+2223 DIVIDES when it materializes m:d/m:sepChr. Inside mfenced this
        // is the same visible separator as MathJax's ASCII vertical line.
        "∣" => "|",
        _ => NormalizeOperatorToken(value),
    };

    private static bool TryCanonicalizeTwoLineStack(
        XElement table,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        var rows = table.Elements()
            .Where(row => row.Name.LocalName is "mtr" or "mlabeledtr")
            .ToArray();
        if (rows.Length != 2) return false;
        var cells = rows.Select(row => row.Elements()
                .Where(cell => cell.Name.LocalName == "mtd")
                .ToArray())
            .ToArray();
        if (cells.Any(row => row.Length != 1)) return false;
        string Line(XElement cell) => CanonicalizeElementSequence(
            cell.Elements()
                .Where(child => child.Name.LocalName is not ("annotation" or "annotation-xml"))
                .ToArray(),
            variant);
        signature = "stack(" + Line(cells[0][0]) + "," + Line(cells[1][0]) + ")";
        return true;
    }

    private static string CanonicalBinary(
        string name,
        XElement[] children,
        string? variant)
    {
        if (children.Length < 2)
            return string.Concat(children.Select(child => CanonicalizeMathMl(child, variant)));
        return name + "("
            + (name is "sub" or "sup"
                ? CanonicalizeScriptBase(children[0], variant)
                : CanonicalizeMathMl(children[0], variant)) + ","
            + CanonicalizeMathMl(children[1], variant) + ")";
    }

    private static string CanonicalTernary(
        string name,
        XElement[] children,
        string? variant)
    {
        if (children.Length < 3)
            return string.Concat(children.Select(child => CanonicalizeMathMl(child, variant)));
        return name + "("
            + (name == "subsup"
                ? CanonicalizeScriptBase(children[0], variant)
                : CanonicalizeMathMl(children[0], variant)) + ","
            + CanonicalizeMathMl(children[1], variant) + ","
            + CanonicalizeMathMl(children[2], variant) + ")";
    }

    private static string CanonicalizeScriptBase(XElement element, string? variant)
    {
        if (element.Name.LocalName == "mo")
        {
            // MathType's BigOp templates export the ordinary set-union/
            // intersection code points while MathJax represents \bigcup/\bigcap
            // with U+22C3/U+22C2. They are the same semantic operator when used
            // as the base of a limit/script structure.
            var value = element.Value.Trim();
            var effectiveVariant = MathMlTokenStyleContract.ResolveEffectiveVariant(element,
                MathMlTokenStyleContract.ResolveDeclaredVariant(element) ?? variant);
            if (value == "∪") return CanonicalStyledOperator("⋃", effectiveVariant);
            if (value == "∩") return CanonicalStyledOperator("⋂", effectiveVariant);
        }
        return CanonicalizeMathMl(element, variant);
    }

    private static string CanonicalToken(string kind, string value, string? mathVariant)
    {
        var token = value.Trim();
        var variant = (mathVariant ?? string.Empty).Trim().ToLowerInvariant();
        if (variant.Contains("sans-serif-bold-italic")) variant = "sans-serif-bold-italic";
        else if (variant.Contains("sans-serif-italic")) variant = "sans-serif-italic";
        else if (variant.Contains("bold-sans-serif")) variant = "bold-sans-serif";
        else if (variant.Contains("bold-script")) variant = "bold-script";
        else if (variant.Contains("bold-fraktur")) variant = "bold-fraktur";
        else if (variant.Contains("bold-italic")) variant = "bold-italic";
        else if (variant.Contains("double-struck")) variant = "double-struck";
        else if (variant.Contains("fraktur")) variant = "fraktur";
        else if (variant.Contains("script")) variant = "script";
        else if (variant.Contains("monospace")) variant = "monospace";
        else if (variant.Contains("sans-serif")) variant = "sans-serif";
        else if (variant.Contains("normal") || variant.Contains("upright")) variant = "normal";
        else if (variant.Contains("bold")) variant = "bold";
        else if (variant.Contains("italic")) variant = string.Empty;
        else variant = string.Empty;
        if (kind == "mi" && variant.Length == 0 && token.Length == 1)
        {
            switch (token[0])
            {
                case 'ℂ': token = "C"; variant = "double-struck"; break;
                case 'ℍ': token = "H"; variant = "double-struck"; break;
                case 'ℕ': token = "N"; variant = "double-struck"; break;
                case 'ℙ': token = "P"; variant = "double-struck"; break;
                case 'ℚ': token = "Q"; variant = "double-struck"; break;
                case 'ℝ': token = "R"; variant = "double-struck"; break;
                case 'ℤ': token = "Z"; variant = "double-struck"; break;
                case 'ℱ': token = "F"; variant = "script"; break;
                case 'ℒ': token = "L"; variant = "script"; break;
                case 'ℛ': token = "R"; variant = "script"; break;
            }
        }
        // Native font preferences must be materialized by the MTEF reader;
        // they cannot override MathML single-mi defaults or explicit italics.
        // Upright ordinary literals (punctuation, digits, mixed text) use MTEF's
        // native Text face and read back as mtext. Preserve their exact text, not
        // the exporter-specific mi tag. Actual mathematical symbols and styled
        // identifiers remain distinct, so this cannot conceal symbol/style loss.
        if (kind == "mi" && variant == "normal"
            && token.Any(character => !char.IsLetter(character))
            && token.All(character => !IsMathematicalSymbolScalar(character)))
            return "text(" + token + ")";
        var explicitlyItalic = string.Equals(mathVariant?.Trim(), "italic", StringComparison.OrdinalIgnoreCase);
        if (kind == "mi" && variant.Length == 0 && !explicitlyItalic
            && token.Length > 1 && token.All(char.IsLetter))
            variant = "normal";
        // MTEF stores styled identifiers as individual CHAR records; exporters
        // may group or split an uninterrupted same-style alphabetic run. Compare
        // its ordered characters with the exact style retained on EVERY token.
        // This covers monospace/script/bold as well as upright runs without
        // equating fonts, symbols, digit tokens or different character contents.
        if (kind == "mi" && (variant.Length > 0 || explicitlyItalic)
            && token.Length > 1 && token.All(char.IsLetter))
            return string.Concat(token.Select(character =>
                "mi[" + variant + "](" + character + ")"));
        return kind + "[" + variant + "](" + token + ")";
    }
}
