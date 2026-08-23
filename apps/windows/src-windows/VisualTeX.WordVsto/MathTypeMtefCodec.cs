using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace VisualTeX.WordVsto;

/// <summary>
/// Minimal, dependency-free MTEF v5 structural codec used to update an existing
/// MathType OLE object without starting or installing MathType.  The original
/// MathType header, font definitions, equation preferences and initial size are
/// preserved byte-for-byte; only the root equation object list is replaced.
/// </summary>
internal static class MathTypeMtefCodec
{
    private const byte MtefVersion5 = 5;
    private const byte RecordEnd = 0;
    private const byte RecordLine = 1;
    private const byte RecordChar = 2;
    private const byte RecordTemplate = 3;
    private const byte RecordPile = 4;
    private const byte RecordMatrix = 5;
    private const byte RecordEmbellishment = 6;
    private const byte RecordFontStyleDef = 8;
    private const byte RecordSize = 9;
    private const byte RecordFull = 10;
    private const byte RecordSub = 11;
    private const byte RecordSub2 = 12;
    private const byte RecordSym = 13;
    private const byte RecordSubSym = 14;
    private const byte RecordColor = 15;
    private const byte RecordColorDef = 16;
    private const byte RecordFontDef = 17;
    private const byte RecordEqnPrefs = 18;
    private const byte RecordEncodingDef = 19;

    private const byte LineNull = 0x01;
    private const byte CharHasEmbellishment = 0x01;
    private const byte CharFunctionStart = 0x02;
    private const byte CharEncoded8 = 0x04;

    private const byte TypefaceText = 1;
    private const byte TypefaceFunction = 2;
    private const byte TypefaceVariable = 3;
    private const byte TypefaceLowerGreek = 4;
    private const byte TypefaceUpperGreek = 5;
    private const byte TypefaceSymbol = 6;
    private const byte TypefaceVector = 7;
    private const byte TypefaceNumber = 8;
    private const byte TypefaceMtExtra = 11;
    // MathType 7 persists expanding fence characters ((), [], {}, |, ...)
    // with typeface 22 and MTCode, not Symbol style. Using Symbol here makes
    // the same ASCII code point resolve to a different fence glyph.
    private const byte TypefaceFence = 22;
    private const byte TypefaceSpace = 24;

    private const byte TemplateAngle = 0;
    private const byte TemplateParen = 1;
    private const byte TemplateBrace = 2;
    private const byte TemplateBracket = 3;
    private const byte TemplateBar = 4;
    private const byte TemplateDoubleBar = 5;
    private const byte TemplateFloor = 6;
    private const byte TemplateCeiling = 7;
    private const byte TemplateRoot = 10;
    private const byte TemplateFraction = 11;
    private const byte TemplateUnderbar = 12;
    private const byte TemplateOverbar = 13;
    private const byte TemplateArrow = 14;
    private const byte TemplateIntegral = 15;
    private const byte TemplateSum = 16;
    private const byte TemplateProduct = 17;
    private const byte TemplateCoproduct = 18;
    private const byte TemplateUnion = 19;
    private const byte TemplateIntersection = 20;
    private const byte TemplateLimit = 23;
    private const byte TemplateHorizontalBrace = 24;
    private const byte TemplateHorizontalBracket = 25;
    private const byte TemplateSub = 27;
    private const byte TemplateSup = 28;
    private const byte TemplateSubSup = 29;
    private const byte TemplateVector = 31;
    private const byte TemplateTilde = 32;
    private const byte TemplateHat = 33;
    private const byte TemplateArc = 34;
    private const byte TemplateStrike = 36;
    private const byte TemplateBox = 37;

    private const byte EmbellDot = 2;
    private const byte EmbellDoubleDot = 3;
    private const byte EmbellTilde = 8;
    private const byte EmbellHat = 9;
    private const byte EmbellRightArrow = 11;
    private const byte EmbellLeftArrow = 12;
    private const byte EmbellBothArrow = 13;
    private const byte EmbellOverbar = 17;

    internal sealed class RewriteResult
    {
        public byte[] EquationNative { get; set; } = Array.Empty<byte>();
        public byte[] Mtef { get; set; } = Array.Empty<byte>();
        public int StructureOffset { get; set; }
    }

    internal static string ReadEquationNativeMathMl(byte[] equationNative)
    {
        if (equationNative is null || equationNative.Length < 40)
            throw new InvalidDataException("MathType Equation Native stream is too short.");
        var headerLength = BitConverter.ToUInt16(equationNative, 0);
        if (headerLength != 28 || headerLength >= equationNative.Length)
            throw new InvalidDataException(
                $"Unsupported MathType OLE native header length: {headerLength}.");
        var objectLength = checked((int)BitConverter.ToUInt32(equationNative, 8));
        if (objectLength <= 0 || headerLength + objectLength > equationNative.Length)
            throw new InvalidDataException(
                $"Invalid MathType OLE MTEF length: {objectLength}.");
        var mtef = new byte[objectLength];
        Buffer.BlockCopy(equationNative, headerLength, mtef, 0, objectLength);
        if (mtef[0] != MtefVersion5)
            throw new InvalidDataException(
                $"VisualTeX currently reads MathType OLE directly for MTEF v5 only, actual={mtef[0]}.");

        var structureOffset = FindRootStructureOffset(mtef);
        var isEmptyEquation =
            structureOffset == mtef.Length - 1
            && mtef[structureOffset] == RecordEnd;
        var content = isEmptyEquation
            ? Array.Empty<XNode>()
            : new MtefStructureReader(mtef, structureOffset).ReadRoot();
        XNamespace mathMlNamespace = "http://www.w3.org/1998/Math/MathML";
        var math = new XElement(
            mathMlNamespace + "math",
            content);

        // MtefStructureReader deliberately builds compact XElement trees using
        // local names (mi/mo/mrow/...). Once those nodes are attached below a
        // namespaced <math>, LINQ to XML serializes them with xmlns="" unless we
        // promote their XName as well. That produced visually plausible MathML
        // which our LaTeX reader tolerated, but standards-based MathML->OMML XSLT
        // saw an empty equation. Normalize every parser-owned descendant into the
        // MathML namespace at the direct MTEF read boundary.
        foreach (var element in math.Descendants().ToArray())
        {
            if (element.Name.Namespace == XNamespace.None)
                element.Name = mathMlNamespace + element.Name.LocalName;
        }
        foreach (var marker in math.DescendantsAndSelf()
                     .Attributes("data-mtef-run")
                     .ToArray())
            marker.Remove();
        return new XDocument(math).ToString(SaveOptions.DisableFormatting);
    }

    private const string StandaloneMtefPrefixBase64 =
        "BQEABwhEU01UNwAAE1dpbkFsbEJhc2ljQ29kZVBhZ2VzABEFVGltZXMgTmV3IFJvbWFuABEDU3ltYm9sABEFQ291cmllciBOZXcAEQRNVCBFeHRyYQATV2luQWxsQ29kZVBhZ2VzABEGy87M5QASAAghL0WPRC9BUPQQD0dfQVDyHx5BUPQVD0EA9EX0JfSPQl9BAPQQD0NfQQD0j0X0Kl9I9I9BAPQQD0D0j0F/SPQQD0EqX0RfRfRfRfRfQQ8MAQABAAECAgICAAIAAQEBAAMAAQAEAAUACg==";

    private static readonly byte[] StandaloneEquationNativeHeader =
    {
        0x1C, 0x00,             // EQNOLEFILEHDR.cbHdr = 28
        0x00, 0x00, 0x02, 0x00, // Equation Native format version
        0x42, 0xC2,             // MathType native clipboard format used by DSMT7
        0x00, 0x00, 0x00, 0x00, // cbObject, filled below
        0x00, 0x00, 0x00, 0x00,
        0xFC, 0xDE, 0x56, 0x0A,
        0x2D, 0xDF, 0xD4, 0x00,
        0x0C, 0x00, 0x85, 0x09,
    };

    internal static RewriteResult CreateEquationNative(
        string mathMl,
        bool inline)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException("MathType creation requires MathML.");

        var prefix = Convert.FromBase64String(StandaloneMtefPrefixBase64);
        var seedMtef = new byte[prefix.Length + 4];
        Buffer.BlockCopy(prefix, 0, seedMtef, 0, prefix.Length);
        seedMtef[prefix.Length] = RecordLine;
        seedMtef[prefix.Length + 1] = 0;
        seedMtef[prefix.Length + 2] = RecordEnd;
        seedMtef[prefix.Length + 3] = RecordEnd;
        if (FindRootStructureOffset(seedMtef) != prefix.Length)
            throw new InvalidDataException(
                "VisualTeX's standalone MathType MTEF prefix is internally inconsistent.");

        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        var math = document.Root?.DescendantsAndSelf()
            .FirstOrDefault(element => element.Name.LocalName == "math")
            ?? throw new InvalidDataException("MathML has no <math> root.");
        var generated = BuildRootStructure(math, seedMtef);
        var mtef = new byte[prefix.Length + generated.Length];
        Buffer.BlockCopy(prefix, 0, mtef, 0, prefix.Length);
        Buffer.BlockCopy(generated, 0, mtef, prefix.Length, generated.Length);

        var equationNative = new byte[StandaloneEquationNativeHeader.Length + mtef.Length];
        Buffer.BlockCopy(
            StandaloneEquationNativeHeader,
            0,
            equationNative,
            0,
            StandaloneEquationNativeHeader.Length);
        Buffer.BlockCopy(
            BitConverter.GetBytes((uint)mtef.Length),
            0,
            equationNative,
            8,
            sizeof(uint));
        Buffer.BlockCopy(
            mtef,
            0,
            equationNative,
            StandaloneEquationNativeHeader.Length,
            mtef.Length);

        _ = inline; // Word paragraph/layout owns inline vs display positioning.
        return new RewriteResult
        {
            EquationNative = equationNative,
            Mtef = mtef,
            StructureOffset = prefix.Length,
        };
    }

    internal static RewriteResult RewriteEquationNative(
        byte[] equationNative,
        string mathMl,
        bool inline)
    {
        if (equationNative is null || equationNative.Length < 40)
            throw new InvalidDataException("MathType Equation Native stream is too short.");
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException("MathType rewrite requires MathML.");

        var headerLength = BitConverter.ToUInt16(equationNative, 0);
        if (headerLength != 28 || headerLength >= equationNative.Length)
            throw new InvalidDataException(
                $"Unsupported MathType OLE native header length: {headerLength}.");
        var objectLength = BitConverter.ToUInt32(equationNative, 8);
        if (objectLength == 0 || headerLength + objectLength > equationNative.Length)
            throw new InvalidDataException(
                $"Invalid MathType OLE MTEF length: {objectLength}.");

        var sourceMtef = new byte[objectLength];
        Buffer.BlockCopy(equationNative, headerLength, sourceMtef, 0, sourceMtef.Length);
        if (sourceMtef[0] != MtefVersion5)
            throw new InvalidDataException(
                $"VisualTeX currently preserves MathType OLE only for MTEF v5, actual={sourceMtef[0]}.");

        var structureOffset = FindRootStructureOffset(sourceMtef);
        if (structureOffset <= 0 || structureOffset > sourceMtef.Length)
            throw new InvalidDataException("Could not locate the MathType root equation structure.");

        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        var math = document.Root?.DescendantsAndSelf()
            .FirstOrDefault(element => element.Name.LocalName == "math")
            ?? throw new InvalidDataException("MathML has no <math> root.");

        using var stream = new MemoryStream();
        stream.Write(sourceMtef, 0, structureOffset);
        // Preserve MathType's equation-options byte exactly as stored in the
        // source MTEF. Word controls whether this OLE sits inline or as a display
        // object; rewriting MathType's own inline bit caused genuine MathType 7
        // character styles (notably MT Extra blackboard-bold glyphs) to be
        // reinterpreted on reopen even though the root MTEF bytes were correct.
        // The `inline` parameter remains part of the public rewrite contract for
        // callers, but must not mutate the MathType-native header.
        _ = inline;
        var generated = BuildRootStructure(math, sourceMtef);
        stream.Write(generated, 0, generated.Length);
        var rewrittenMtef = stream.ToArray();

        var rewrittenNative = new byte[headerLength + rewrittenMtef.Length];
        Buffer.BlockCopy(equationNative, 0, rewrittenNative, 0, headerLength);
        Buffer.BlockCopy(rewrittenMtef, 0, rewrittenNative, headerLength, rewrittenMtef.Length);
        var objectLengthBytes = BitConverter.GetBytes((uint)rewrittenMtef.Length);
        Buffer.BlockCopy(objectLengthBytes, 0, rewrittenNative, 8, objectLengthBytes.Length);

        return new RewriteResult
        {
            EquationNative = rewrittenNative,
            Mtef = rewrittenMtef,
            StructureOffset = structureOffset,
        };
    }

    internal static string SemanticSignature(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)) return string.Empty;
        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        var math = document.Root?.DescendantsAndSelf()
            .FirstOrDefault(element => element.Name.LocalName == "math")
            ?? document.Root
            ?? throw new InvalidDataException("MathML has no root element.");
        return CanonicalizeMathMl(math, inheritedMathVariant: null);
    }

    private static string CanonicalizeMathMl(XElement element, string? inheritedMathVariant)
    {
        var local = element.Name.LocalName;
        var ownVariant = ((string?)element.Attribute("mathvariant"))?.Trim();
        var variant = string.IsNullOrWhiteSpace(ownVariant) ? inheritedMathVariant : ownVariant;
        string Children(string? childVariant = null) => string.Concat(
            element.Elements()
                .Where(child => child.Name.LocalName is not ("annotation" or "annotation-xml"))
                .Select(child => CanonicalizeMathMl(child, childVariant ?? variant)));
        var children = element.Elements()
            .Where(child => child.Name.LocalName is not ("annotation" or "annotation-xml"))
            .ToArray();
        switch (local)
        {
            case "math":
            case "mpadded":
            {
                if (TryCanonicalizeSplitFencedSuperscriptSequence(children, variant, out var splitFencedSuperscript))
                    return splitFencedSuperscript;
                if (TryCanonicalizeLooseBinomialSequence(children, variant, out var looseBinomialSequence))
                    return looseBinomialSequence;
                return Children();
            }
            case "mrow":
            {
                if (TryCanonicalizeSplitFencedSuperscriptSequence(children, variant, out var splitFencedSuperscript))
                    return splitFencedSuperscript;
                if (TryCanonicalizeLooseBinomialSequence(children, variant, out var looseBinomialSequence))
                    return looseBinomialSequence;
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
                    return "fence(" + NormalizeFence(open) + "," + NormalizeFence(close) + ","
                        + string.Concat(fenceChildren.Select(child => CanonicalizeMathMl(child, variant))) + ")";
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
                // MathType 7 sometimes exports standalone mathematical glyphs
                // such as infinity as <mi> even though MathJax/VisualTeX uses
                // <mo>.  That token-tag difference is not a semantic change.
                // Keep mtext distinct so an actual Symbol -> Text regression is
                // still rejected by the round-trip validator.
                if (IsMathematicalSymbolToken(value))
                    return "o(" + NormalizeOperatorToken(value) + ")";
                return CanonicalToken("mi", value, variant);
            }
            case "mn":
                return "n(" + element.Value.Trim() + ")";
            case "mo":
            {
                var value = element.Value.Trim();
                // MathJax inserts U+2061 FUNCTION APPLICATION between a named
                // function and its argument. It has no visible glyph and should
                // never become a literal MathType character or a fake '?'.
                if (value == "⁡") return string.Empty;
                var primeSignature = CanonicalPrimeSignature(value);
                if (primeSignature is not null) return primeSignature;
                if (value.Length > 1 && value.All(char.IsLetter))
                    return CanonicalToken("mi", value, "normal");
                return "o(" + NormalizeOperatorToken(value) + ")";
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
                    return CanonicalToken("mi", value, "normal");
                return "text(" + value + ")";
            }
            case "mfrac":
                return CanonicalBinary("frac", children, variant);
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
                if (IsHorizontalBraceMarker(over, top: true))
                    return "hbrace-top(" + CanonicalizeMathMl(children[0], variant) + ")";
                if (over == "¯")
                    return "accent(overbar," + CanonicalizeMathMl(children[0], variant) + ")";
                var accent = over is "→" or "←" or "↔"
                    or "^" or "~" or "." or "˙" or "¨"
                    or "⌢" or "⏜" or "ˇ" or "˘" or "´" or "`" or "˚";
                return accent
                    ? "accent(" + over + "," + CanonicalizeMathMl(children[0], variant) + ")"
                    : CanonicalBinary("sup", children, variant);
            }
            case "munderover":
                return CanonicalTernary("subsup", children, variant);
            case "mfenced":
            {
                if (TryCanonicalizeMathTypeBinomialFence(element, variant, out var binomial))
                    return binomial;
                var open = NormalizeFence((string?)element.Attribute("open") ?? "(");
                var close = NormalizeFence((string?)element.Attribute("close") ?? ")");
                return "fence(" + open + "," + close + "," + Children() + ")";
            }
            case "mtable":
                return "table(" + string.Join(";", children.Select(row =>
                    CanonicalizeMathMl(row, variant))) + ")";
            case "mtr":
            case "mlabeledtr":
                return "row(" + string.Join(",", children.Select(cell =>
                    CanonicalizeMathMl(cell, variant))) + ")";
            case "mtd":
                return "cell(" + Children() + ")";
            case "menclose":
            {
                var notation = ((string?)element.Attribute("notation") ?? string.Empty)
                    .Trim()
                    .ToLowerInvariant();
                if (notation.Contains("radical")) return "sqrt(" + Children() + ")";
                if (notation.Contains("box")) return "box(" + Children() + ")";
                var strikeParts = new List<string>();
                if (notation.Contains("horizontalstrike")) strikeParts.Add("h");
                if (notation.Contains("updiagonalstrike")) strikeParts.Add("up");
                if (notation.Contains("downdiagonalstrike")) strikeParts.Add("down");
                if (strikeParts.Count > 0)
                    return "strike[" + string.Join("+", strikeParts) + "](" + Children() + ")";
                return Children();
            }
            case "mphantom":
                return "phantom(" + Children() + ")";
            case "mmultiscripts":
            {
                if (children.Length >= 3
                    && children.All(child => child.Name.LocalName != "mprescripts"))
                {
                    var baseSignature = CanonicalizeMathMl(children[0], variant);
                    var subSignature = children[1].Name.LocalName == "none"
                        ? string.Empty
                        : CanonicalizeMathMl(children[1], variant);
                    var supSignature = children[2].Name.LocalName == "none"
                        ? string.Empty
                        : CanonicalizeMathMl(children[2], variant);
                    if (subSignature.Length > 0 && supSignature.Length > 0)
                        return "subsup(" + baseSignature + "," + subSignature + "," + supSignature + ")";
                    if (subSignature.Length > 0)
                        return "sub(" + baseSignature + "," + subSignature + ")";
                    if (supSignature.Length > 0)
                        return "sup(" + baseSignature + "," + supSignature + ")";
                    return baseSignature;
                }
                return "multi(" + Children() + ")";
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

    private static bool IsMathematicalSymbolToken(string value) =>
        value.Length > 0 && value.All(character =>
        {
            var category = char.GetUnicodeCategory(character);
            return category == System.Globalization.UnicodeCategory.MathSymbol
                || category == System.Globalization.UnicodeCategory.CurrencySymbol
                || category == System.Globalization.UnicodeCategory.ModifierSymbol
                || category == System.Globalization.UnicodeCategory.OtherSymbol;
        });

    private static bool IsHorizontalBraceMarker(string value, bool top) =>
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
            if (value == "∪") return "o(⋃)";
            if (value == "∩") return "o(⋂)";
        }
        return CanonicalizeMathMl(element, variant);
    }

    private static string CanonicalToken(string kind, string value, string? mathVariant)
    {
        var token = value.Trim();
        var variant = (mathVariant ?? string.Empty).Trim().ToLowerInvariant();
        if (variant.Contains("double-struck")) variant = "double-struck";
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
        if (kind == "mi" && variant.Length == 0
            && token.Length == 1
            && IsUpperGreek(token[0]))
            variant = "normal";
        if (kind == "mi" && variant.Length == 0
            && token.Length > 1 && token.All(char.IsLetter))
            variant = "normal";
        // MathType coalesces consecutive upright identifier runs (for example
        // <mi mathvariant="normal">R</mi><mi mathvariant="normal">e</mi>) into
        // one <mtext>Re</mtext> token on MTEF read-back. Token boundaries inside
        // the same upright alphabetic run are presentational, not semantic.
        if (kind == "mi" && variant == "normal"
            && token.Length > 1 && token.All(char.IsLetter))
            return string.Concat(token.Select(character =>
                "mi[normal](" + character + ")"));
        return kind + "[" + variant + "](" + token + ")";
    }

    internal static int FindRootStructureOffset(byte[] mtef)
    {
        if (mtef is null || mtef.Length < 16 || mtef[0] != MtefVersion5)
            throw new InvalidDataException("Invalid or unsupported MTEF stream.");
        var position = FindEquationOptionsOffset(mtef) + 1;
        var sawPreferences = false;
        var preferencesOffset = -1;
        var sawInitialSize = false;
        while (position < mtef.Length)
        {
            var record = mtef[position];
            switch (record)
            {
                case RecordEncodingDef:
                    position = SkipNullTerminated(mtef, position + 1);
                    break;
                case RecordFontDef:
                    position++;
                    _ = ReadUnsigned(mtef, ref position);
                    position = SkipNullTerminated(mtef, position);
                    break;
                case RecordFontStyleDef:
                    position++;
                    _ = ReadUnsigned(mtef, ref position);
                    Require(mtef, position, 1);
                    position++;
                    break;
                case RecordColorDef:
                    position = SkipColorDefinition(mtef, position);
                    break;
                case RecordEqnPrefs:
                    preferencesOffset = position;
                    position = SkipEquationPreferences(mtef, position);
                    sawPreferences = true;
                    break;
                case >= 100:
                    // MTEF v5 explicitly reserves record ids >= 100 for forward
                    // compatible extensions. MathType versions are free to emit
                    // them before the root; readers must skip the length-prefixed
                    // payload instead of treating the record as the initial SIZE.
                    position = SkipFutureRecord(mtef, position);
                    break;
                default:
                    if (!sawPreferences)
                        throw new InvalidDataException(
                            $"Unexpected MTEF record {record} before equation preferences at offset {position}.");
                    if (record is RecordLine or RecordPile or RecordMatrix)
                    {
                        if (!sawInitialSize)
                            WordDoubleClickHook.TraceMessage(
                                $"mathtype-mtef-root-without-initial-size offset={position} record={record}");
                        return position;
                    }
                    if (record == RecordSize
                        || record >= RecordFull && record <= RecordSubSym)
                    {
                        // The published layout contains one initial size record,
                        // but accepting repeated size changes here is harmless and
                        // makes the reader tolerant of vendor/version prefixes while
                        // keeping all of them in the immutable source prefix.
                        position = SkipInitialSizeRecord(mtef, position);
                        sawInitialSize = true;
                        break;
                    }
                    if (record == RecordColor)
                    {
                        position++;
                        _ = ReadUnsigned(mtef, ref position);
                        break;
                    }
                    // MathType 7.x point releases can produce a prefix boundary
                    // that differs from the canonical DSMT7 seed. Before treating
                    // the current byte as authoritative, scan the post-preference
                    // region for a LINE/PILE/MATRIX that parses as exactly one
                    // complete top-level equation. A nested structure cannot pass
                    // that whole-equation check because parent records remain.
                    if (TryRecoverRootStructureOffset(
                            mtef,
                            preferencesOffset >= 0 ? preferencesOffset + 2 : 0,
                            position,
                            out var recoveredRoot))
                        return recoveredRoot;
                    // A genuine empty MathType equation is encoded as the normal
                    // prefix/initial-size state followed directly by the final
                    // equation END. MathType 7.8.x therefore legitimately has a
                    // zero at the same offset where non-empty DSMT7 equations put
                    // their root LINE. Treat only the terminal END as empty; any
                    // other zero remains a hard parse failure.
                    if (record == RecordEnd && position == mtef.Length - 1)
                    {
                        WordDoubleClickHook.TraceMessage(
                            $"mathtype-mtef-empty-equation endOffset={position}");
                        return position;
                    }
                    throw new InvalidDataException(
                        $"Unsupported MathType root record {record} at offset {position}.");
            }
        }
        throw new InvalidDataException("MTEF has no root equation structure.");
    }

    private static bool TryRecoverRootStructureOffset(
        byte[] data,
        int lowerBound,
        int expectedOffset,
        out int rootOffset)
    {
        rootOffset = -1;
        const int forwardSearchRadius = 256;
        var start = Math.Max(0, lowerBound);
        var end = Math.Min(data.Length - 2, expectedOffset + forwardSearchRadius);
        var candidates = new List<int>();
        for (var position = start; position <= end; position++)
        {
            var record = data[position];
            if (record is not RecordLine and not RecordPile and not RecordMatrix)
                continue;
            if (!RootConsumesWholeEquation(data, position))
                continue;
            candidates.Add(position);
        }
        if (candidates.Count == 0) return false;
        rootOffset = candidates
            .OrderBy(position => Math.Abs(position - expectedOffset))
            .ThenBy(position => position)
            .First();
        return true;
    }

    private static bool RootConsumesWholeEquation(byte[] data, int rootOffset)
    {
        try
        {
            var parser = new MtefStructureReader(data, rootOffset);
            _ = parser.ReadRoot().ToArray();
            return parser.HasOnlyEquationEndRemaining();
        }
        catch (Exception error) when (
            error is InvalidDataException
            || error is EndOfStreamException
            || error is ArgumentOutOfRangeException)
        {
            return false;
        }
    }

    private static int FindEquationOptionsOffset(byte[] mtef)
    {
        Require(mtef, 0, 6);
        var position = 5; // version/platform/product/version/subversion
        position = SkipNullTerminated(mtef, position);
        Require(mtef, position, 1);
        return position;
    }

    private static int SkipEquationPreferences(byte[] data, int position)
    {
        Require(data, position, 2);
        position += 2; // record type + options
        position = SkipDimensionArray(data, position);
        position = SkipDimensionArray(data, position);
        Require(data, position, 1);
        var styleCount = data[position++];
        for (var index = 0; index < styleCount; index++)
        {
            var fontIndex = ReadUnsigned(data, ref position);
            if (fontIndex != 0)
            {
                Require(data, position, 1);
                position++; // character style
            }
        }
        return position;
    }

    private static int SkipDimensionArray(byte[] data, int position)
    {
        Require(data, position, 1);
        var count = data[position++];
        var nibbleIndex = 0;
        byte ReadNibble()
        {
            Require(data, position, 1);
            var value = nibbleIndex == 0
                ? (byte)(data[position] >> 4)
                : (byte)(data[position] & 0x0F);
            if (nibbleIndex == 0) nibbleIndex = 1;
            else
            {
                nibbleIndex = 0;
                position++;
            }
            return value;
        }

        for (var dimension = 0; dimension < count; dimension++)
        {
            _ = ReadNibble(); // units
            while (ReadNibble() != 0x0F) { }
        }
        if (nibbleIndex != 0) position++; // padded low nibble
        return position;
    }

    private static int SkipColorDefinition(byte[] data, int position)
    {
        Require(data, position, 2);
        var options = data[position + 1];
        position += 2;
        var cmyk = (options & 0x01) != 0;
        var named = (options & 0x04) != 0;
        Require(data, position, cmyk ? 8 : 6);
        position += cmyk ? 8 : 6;
        if (named) position = SkipNullTerminated(data, position);
        return position;
    }

    private static int SkipFutureRecord(byte[] data, int position)
    {
        Require(data, position, 1);
        if (data[position] < 100)
            throw new InvalidDataException(
                $"MTEF record at offset {position} is not a future-extension record.");
        position++;
        var length = ReadUnsigned(data, ref position);
        Require(data, position, length);
        return position + length;
    }

    private static int SkipInitialSizeRecord(byte[] data, int position)
    {
        Require(data, position, 1);
        var record = data[position];
        if (record >= RecordFull && record <= RecordSubSym)
            return position + 1;
        if (record != RecordSize)
            throw new InvalidDataException(
                $"Expected initial MTEF size record at offset {position}, actual={record}.");
        Require(data, position, 2);
        var form = data[position + 1];
        if (form == 100 || form == 101)
        {
            Require(data, position, form == 100 ? 5 : 4);
            return position + (form == 100 ? 5 : 4);
        }
        Require(data, position, 3);
        return position + 3;
    }

    private static int SkipNullTerminated(byte[] data, int position)
    {
        while (position < data.Length && data[position] != 0) position++;
        if (position >= data.Length)
            throw new EndOfStreamException("MTEF null-terminated field is truncated.");
        return position + 1;
    }

    private static int ReadUnsigned(byte[] data, ref int position)
    {
        Require(data, position, 1);
        var value = data[position++];
        if (value != 0xFF) return value;
        Require(data, position, 2);
        var expanded = data[position] | (data[position + 1] << 8);
        position += 2;
        return expanded;
    }

    private static byte[] BuildRootStructure(XElement math, byte[] sourceMtef)
    {
        var preparedMath = new XElement(math);
        MaterializeInheritedMathVariants(preparedMath, inheritedMathVariant: null);
        NormalizeMathJaxFenceRows(preparedMath);
        var output = new List<byte> { RecordLine, 0 };
        // Preserve MathType's root-level formatting state (notably COLOR_DEF /
        // COLOR) from the source equation. MathType 7 writes these records even
        // when the visible formula is plain black; dropping them produced OLEs
        // that our own parser could read but the native MathType server could
        // reinterpret with the wrong expandable-fence family.
        CopyRootLeadingFormattingRecords(sourceMtef, output);
        EmitExplicitFontDefinitions(preparedMath, sourceMtef, output);
        EmitContainerChildren(preparedMath, output, inheritedMathVariant: null);
        output.Add(RecordEnd); // root line
        output.Add(RecordEnd); // equation
        return output.ToArray();
    }

    private static void CopyRootLeadingFormattingRecords(
        byte[] sourceMtef,
        List<byte> output)
    {
        var root = FindRootStructureOffset(sourceMtef);
        if (root >= sourceMtef.Length
            || sourceMtef[root] == RecordEnd)
            return;
        Require(sourceMtef, root, 2);
        if (sourceMtef[root] != RecordLine)
            return;
        var cursor = root + 2;
        while (cursor < sourceMtef.Length)
        {
            var start = cursor;
            var record = sourceMtef[cursor];
            if (record == RecordColorDef)
                cursor = SkipColorDefinition(sourceMtef, cursor);
            else if (record == 15) // COLOR
            {
                cursor++;
                _ = ReadUnsigned(sourceMtef, ref cursor);
            }
            else
                break;
            for (var index = start; index < cursor; index++)
                output.Add(sourceMtef[index]);
        }
    }

    private static void MaterializeInheritedMathVariants(
        XElement element,
        string? inheritedMathVariant)
    {
        var ownVariant = ((string?)element.Attribute("mathvariant"))?.Trim();
        var variant = string.IsNullOrWhiteSpace(ownVariant)
            ? inheritedMathVariant
            : ownVariant;
        if (element.Name.LocalName is "mi" or "mn" or "mo" or "mtext")
        {
            if (element.Attribute("mathvariant") is null
                && !string.IsNullOrWhiteSpace(variant))
                element.SetAttributeValue("mathvariant", variant);
            return;
        }
        foreach (var child in element.Elements())
            MaterializeInheritedMathVariants(child, variant);
    }

    private static void NormalizeMathJaxFenceRows(XElement root)
    {
        var rows = root.DescendantsAndSelf()
            .Where(element => element.Name.LocalName == "mrow")
            .Reverse()
            .ToArray();
        foreach (var row in rows)
        {
            if (!TryGetMathJaxFenceRow(row, out var open, out var close, out var children))
                continue;
            var fenced = new XElement(
                "mfenced",
                new XAttribute("open", NormalizeFence(open)),
                new XAttribute("close", NormalizeFence(close)),
                children.Select(child => new XElement(child)));
            row.ReplaceNodes(fenced);
        }
    }

    private static bool TryCanonicalizeLooseBinomialSequence(
        XElement[] elements,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        if (elements.Length < 3) return false;
        var builder = new StringBuilder();
        var replaced = false;
        for (var index = 0; index < elements.Length;)
        {
            if (index + 2 < elements.Length
                && TryGetLooseFenceToken(elements[index], out var open)
                && TryGetLooseFenceToken(elements[index + 2], out var close)
                && NormalizeFence(open) == "("
                && NormalizeFence(close) == ")")
            {
                if (TryCanonicalizeLooseBinomialBody(
                        elements[index + 1],
                        variant,
                        out var binomialBody))
                {
                    builder.Append(binomialBody);
                    index += 3;
                    replaced = true;
                    continue;
                }
            }
            builder.Append(CanonicalizeMathMl(elements[index], variant));
            index++;
        }
        if (!replaced) return false;
        signature = builder.ToString();
        return true;
    }

    private static bool TryCanonicalizeLooseBinomialBody(
        XElement body,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        if (body.Name.LocalName == "mfrac")
        {
            var thickness = ((string?)body.Attribute("linethickness") ?? string.Empty).Trim();
            var fractionChildren = body.Elements().ToArray();
            if (!thickness.StartsWith("0", StringComparison.Ordinal)
                || fractionChildren.Length != 2)
                return false;
            signature = "binom("
                + CanonicalizeMathMl(fractionChildren[0], variant)
                + ","
                + CanonicalizeMathMl(fractionChildren[1], variant)
                + ")";
            return true;
        }

        body = UnwrapSingleTransparentContainer(body);
        if (body.Name.LocalName != "mtable"
            || !string.Equals(
                (string?)body.Attribute("data-mtef-pile"),
                "true",
                StringComparison.OrdinalIgnoreCase))
            return false;
        var rows = body.Elements()
            .Where(row => row.Name.LocalName is "mtr" or "mlabeledtr")
            .ToArray();
        if (rows.Length != 2) return false;
        var cells = rows.Select(row => row.Elements()
                .Where(cell => cell.Name.LocalName == "mtd")
                .ToArray())
            .ToArray();
        if (cells.Any(row => row.Length != 1)) return false;
        var top = string.Concat(cells[0][0].Elements()
            .Select(child => CanonicalizeMathMl(child, variant)));
        var bottom = string.Concat(cells[1][0].Elements()
            .Select(child => CanonicalizeMathMl(child, variant)));
        signature = "binom(" + top + "," + bottom + ")";
        return true;
    }

    private static XElement UnwrapSingleTransparentContainer(XElement element)
    {
        while (element.Name.LocalName is "mrow" or "mstyle" or "mpadded")
        {
            var children = element.Elements().ToArray();
            if (children.Length != 1) break;
            element = children[0];
        }
        return element;
    }

    private static bool TryCanonicalizeMathJaxBinomialFence(
        string open,
        string close,
        XElement[] children,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        if (NormalizeFence(open) != "(" || NormalizeFence(close) != ")") return false;
        if (children.Length != 1 || children[0].Name.LocalName != "mfrac") return false;
        var lineThickness = (children[0].Attribute("linethickness")?.Value ?? string.Empty).Trim();
        if (lineThickness.Length == 0 || !lineThickness.StartsWith("0", StringComparison.Ordinal)) return false;
        var fractionChildren = children[0].Elements().ToArray();
        if (fractionChildren.Length != 2) return false;
        var top = CanonicalizeMathMl(fractionChildren[0], variant);
        var bottom = CanonicalizeMathMl(fractionChildren[1], variant);
        signature = "binom(" + top + "," + bottom + ")";
        return true;
    }

    private static bool TryCanonicalizeMathTypeBinomialFence(
        XElement fenced,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        var open = NormalizeFence((string?)fenced.Attribute("open") ?? "(");
        var close = NormalizeFence((string?)fenced.Attribute("close") ?? ")");
        if (open != "(" || close != ")") return false;
        var table = fenced.Elements().SingleOrDefault();
        // Word's OMML round-trip commonly materializes a native no-bar
        // fraction as mfenced -> mrow -> mfrac(linethickness=0). Treat that
        // as the same binomial semantics before checking MathType's PILE form.
        if (table is not null
            && TryCanonicalizeLooseBinomialBody(table, variant, out signature))
            return true;
        if (table is not null)
            table = UnwrapSingleTransparentContainer(table);
        if (table?.Name.LocalName == "mfrac")
            return TryCanonicalizeMathJaxBinomialFence(
                open,
                close,
                new[] { table },
                variant,
                out signature);
        if (table is null || table.Name.LocalName != "mtable") return false;
        // A genuine MathType binomial reaches us through a PILE record. An
        // explicit 2x1 matrix inside parentheses is visually similar, but is a
        // MATRIX record and must remain a matrix. The old shape-only heuristic
        // collapsed both into binom(...), which made strict VisualTeX→MathType
        // round-trip validation reject explicit column matrices such as the
        // binomial-theorem formula from 文档1.
        if (!string.Equals(
                (string?)table.Attribute("data-mtef-pile"),
                "true",
                StringComparison.OrdinalIgnoreCase))
            return false;
        var rows = table.Elements()
            .Where(row => row.Name.LocalName is "mtr" or "mlabeledtr")
            .ToArray();
        if (rows.Length != 2) return false;
        var cells = rows.Select(row => row.Elements()
                .Where(cell => cell.Name.LocalName == "mtd")
                .ToArray())
            .ToArray();
        if (cells.Any(row => row.Length != 1)) return false;
        var top = string.Concat(cells[0][0].Elements()
            .Select(child => CanonicalizeMathMl(child, variant)));
        var bottom = string.Concat(cells[1][0].Elements()
            .Select(child => CanonicalizeMathMl(child, variant)));
        signature = "binom(" + top + "," + bottom + ")";
        return true;
    }

    private static bool TryCanonicalizeSplitFencedSuperscriptSequence(
        XElement[] elements,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        if (elements.Length < 3) return false;
        if (!TryGetLooseFenceToken(elements[0], out var open)) return false;

        for (var scriptIndex = 2; scriptIndex < elements.Length; scriptIndex++)
        {
            var script = elements[scriptIndex];
            if (script.Name.LocalName != "msup") continue;
            var scriptChildren = script.Elements().ToArray();
            if (scriptChildren.Length < 2) continue;
            if (!TryGetLooseFenceToken(scriptChildren[0], out var close)) continue;
            if (!AreMatchingFences(open, close)) continue;

            var inner = string.Concat(elements
                .Skip(1)
                .Take(scriptIndex - 1)
                .Select(child => CanonicalizeMathMl(child, variant)));
            if (inner.Length == 0) continue;
            var exponentSignature = CanonicalizeMathMl(scriptChildren[1], variant);
            if (exponentSignature.Length == 0) continue;

            var baseSignature = "o(" + NormalizeFence(open) + ")"
                + inner
                + "o(" + NormalizeFence(close) + ")";
            var tail = string.Concat(elements
                .Skip(scriptIndex + 1)
                .Select(child => CanonicalizeMathMl(child, variant)));
            signature = "sup(" + baseSignature + "," + exponentSignature + ")" + tail;
            return true;
        }
        return false;
    }

    private static bool TryGetLooseFenceToken(XElement candidate, out string fence)
    {
        fence = string.Empty;
        XElement? token = candidate.Name.LocalName == "mo" ? candidate : null;
        if (token is null && candidate.Name.LocalName == "mrow")
        {
            var nested = candidate.Elements().ToArray();
            if (nested.Length == 1 && nested[0].Name.LocalName == "mo")
                token = nested[0];
        }
        if (token is null) return false;
        fence = NormalizeFence(token.Value.Trim());
        return fence is "(" or ")" or "[" or "]" or "{" or "}"
            or "⟨" or "⟩" or "⌈" or "⌉" or "⌊" or "⌋" or "|" or "‖";
    }

    private static bool AreMatchingFences(string open, string close) =>
        (NormalizeFence(open), NormalizeFence(close)) switch
        {
            ("(", ")") => true,
            ("[", "]") => true,
            ("{", "}") => true,
            ("⟨", "⟩") => true,
            ("⌈", "⌉") => true,
            ("⌊", "⌋") => true,
            ("|", "|") => true,
            ("‖", "‖") => true,
            _ => false,
        };

    private static bool TryCanonicalizeMathJaxFencedSuperscriptRow(
        XElement row,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        if (row.Name.LocalName != "mrow") return false;
        var elements = row.Elements().ToArray();
        if (elements.Length < 3) return false;
        if (!TryGetMathJaxFenceToken(elements[0], "OPEN", out var open)) return false;

        var script = elements[elements.Length - 1];
        if (script.Name.LocalName != "msup") return false;
        var scriptChildren = script.Elements().ToArray();
        if (scriptChildren.Length < 2) return false;
        if (!TryGetMathJaxFenceToken(scriptChildren[0], "CLOSE", out var close)) return false;

        var inner = string.Concat(elements
            .Skip(1)
            .Take(elements.Length - 2)
            .Select(child => CanonicalizeMathMl(child, variant)));
        var baseSignature = "o(" + NormalizeFence(open) + ")"
            + inner
            + "o(" + NormalizeFence(close) + ")";
        var exponentSignature = CanonicalizeMathMl(scriptChildren[1], variant);
        signature = exponentSignature.Length == 0
            ? baseSignature
            : "sup(" + baseSignature + "," + exponentSignature + ")";
        return true;
    }

    private static bool TryGetMathJaxFenceRow(
        XElement row,
        out string open,
        out string close,
        out XElement[] children)
    {
        open = string.Empty;
        close = string.Empty;
        children = Array.Empty<XElement>();
        if (row.Name.LocalName != "mrow") return false;
        var elements = row.Elements().ToArray();
        if (elements.Length < 3) return false;
        if (!TryGetMathJaxFenceToken(elements[0], "OPEN", out open)) return false;
        if (!TryGetMathJaxFenceToken(elements[elements.Length - 1], "CLOSE", out close)) return false;
        children = elements.Skip(1).Take(elements.Length - 2).ToArray();
        return children.Length > 0;
    }

    private static bool TryGetMathJaxFenceToken(
        XElement candidate,
        string expectedClass,
        out string fence)
    {
        fence = string.Empty;
        XElement? token = null;
        if (candidate.Name.LocalName == "mo")
            token = candidate;
        else if (candidate.Name.LocalName == "mrow")
        {
            var nested = candidate.Elements().ToArray();
            if (nested.Length == 1 && nested[0].Name.LocalName == "mo")
                token = nested[0];
        }
        if (token is null) return false;

        var candidateClass = ((string?)candidate.Attribute("data-mjx-texclass") ?? string.Empty).Trim();
        var tokenClass = ((string?)token.Attribute("data-mjx-texclass") ?? string.Empty).Trim();
        var markerMatches = string.Equals(candidateClass, expectedClass, StringComparison.OrdinalIgnoreCase)
            || string.Equals(tokenClass, expectedClass, StringComparison.OrdinalIgnoreCase);
        if (!markerMatches) return false;
        fence = token.Value.Trim();
        return true;
    }

    private static void EmitExplicitFontDefinitions(
        XElement math,
        byte[] sourceMtef,
        List<byte> output)
    {
        var specs = new[]
        {
            (Variant: "script", Encoding: "EuclidMath1", Font: "Euclid Math One"),
            (Variant: "double-struck", Encoding: "EuclidMath2", Font: "Euclid Math Two"),
            (Variant: "fraktur", Encoding: "EuclidFraktur", Font: "Euclid Fraktur"),
        };
        var existing = CountPrefixDefinitions(sourceMtef);
        var created = 0;
        foreach (var spec in specs)
        {
            var tokens = math.DescendantsAndSelf()
                .Where(element => element.Name.LocalName == "mi"
                    && string.Equals(
                        ((string?)element.Attribute("mathvariant"))?.Trim(),
                        spec.Variant,
                        StringComparison.OrdinalIgnoreCase)
                    && !(string.Equals(spec.Variant, "double-struck", StringComparison.OrdinalIgnoreCase)
                        && EnumerateBmpScalars(element.Value)
                            .All(scalar => TryMtExtraDoubleStruck(scalar, out _, out _))))
                .ToArray();
            if (tokens.Length == 0) continue;

            created++;
            var encodingIndex = 4 + existing.EncodingDefinitions + created;
            var fontDefinitionIndex = existing.FontDefinitions + created;
            var fontStyleIndex = existing.FontStyleDefinitions + created;

            output.Add(RecordEncodingDef);
            WriteNullTerminatedAscii(output, spec.Encoding);
            output.Add(RecordFontDef);
            WriteUnsigned(output, encodingIndex);
            WriteNullTerminatedAscii(output, spec.Font);
            output.Add(RecordFontStyleDef);
            WriteUnsigned(output, fontDefinitionIndex);
            output.Add(0); // plain character style; the explicit font carries the alphabet style

            foreach (var token in tokens)
                token.SetAttributeValue("data-mtef-explicit-typeface", -fontStyleIndex);
        }
    }

    private static (int EncodingDefinitions, int FontDefinitions, int FontStyleDefinitions)
        CountPrefixDefinitions(byte[] mtef)
    {
        var root = FindRootStructureOffset(mtef);
        var position = FindEquationOptionsOffset(mtef) + 1;
        var encodingCount = 0;
        var fontCount = 0;
        var styleCount = 0;
        while (position < root)
        {
            var record = mtef[position];
            switch (record)
            {
                case RecordEncodingDef:
                    encodingCount++;
                    position = SkipNullTerminated(mtef, position + 1);
                    break;
                case RecordFontDef:
                    fontCount++;
                    position++;
                    _ = ReadUnsigned(mtef, ref position);
                    position = SkipNullTerminated(mtef, position);
                    break;
                case RecordFontStyleDef:
                    styleCount++;
                    position++;
                    _ = ReadUnsigned(mtef, ref position);
                    position++;
                    break;
                case RecordColorDef:
                    position = SkipColorDefinition(mtef, position);
                    break;
                case RecordEqnPrefs:
                    position = SkipEquationPreferences(mtef, position);
                    break;
                default:
                    position = SkipInitialSizeRecord(mtef, position);
                    return (encodingCount, fontCount, styleCount);
            }
        }
        return (encodingCount, fontCount, styleCount);
    }

    private static void WriteNullTerminatedAscii(List<byte> output, string value)
    {
        output.AddRange(System.Text.Encoding.ASCII.GetBytes(value));
        output.Add(0);
    }

    private static void WriteUnsigned(List<byte> output, int value)
    {
        if (value < 0)
            throw new ArgumentOutOfRangeException(nameof(value));
        if (value < 255)
        {
            output.Add((byte)value);
            return;
        }
        output.Add(0xFF);
        output.Add((byte)(value & 0xFF));
        output.Add((byte)((value >> 8) & 0xFF));
    }

    private static void EmitNode(XNode node, List<byte> output)
    {
        if (node is XText text)
        {
            EmitText(text.Value, TypefaceText, output);
            return;
        }
        if (node is not XElement element) return;
        var local = element.Name.LocalName;
        switch (local)
        {
            case "math":
            case "mrow":
            case "semantics":
            case "mpadded":
            case "mphantom":
                EmitContainerChildren(element, output, inheritedMathVariant: null);
                break;
            case "mstyle":
                EmitContainerChildren(
                    element,
                    output,
                    ((string?)element.Attribute("mathvariant"))?.Trim());
                break;
            case "mi":
                EmitIdentifier(element, output);
                break;
            case "mn":
                EmitText(element.Value, TypefaceNumber, output);
                break;
            case "mo":
                if (TryResolveNativeIntegral(element, out var integral))
                    EmitIntegralTemplate(integral, main: null, lower: null, upper: null, output);
                else if (TryResolveNativeBigOperator(element, out var standaloneBigOperator))
                    EmitBigOperatorTemplate(standaloneBigOperator, main: null, lower: null, upper: null, output);
                else if (IsNamedLimitOperator(element))
                    EmitFunctionRun(element.Value.Trim(), output);
                else
                    EmitOperator(element.Value, output);
                break;
            case "mtext":
                EmitText(element.Value, TypefaceText, output);
                break;
            case "mspace":
                EmitCharacter(' ', TypefaceSpace, output);
                break;
            case "mfrac":
                EmitFraction(element, output);
                break;
            case "msqrt":
                EmitSquareRoot(element, output);
                break;
            case "mroot":
                EmitNthRoot(element, output);
                break;
            case "msup":
                EmitScript(element, TemplateSup, output);
                break;
            case "msub":
                EmitScript(element, TemplateSub, output);
                break;
            case "msubsup":
                EmitScript(element, TemplateSubSup, output);
                break;
            case "mfenced":
                EmitFenced(element, output);
                break;
            case "mover":
                EmitOver(element, output);
                break;
            case "munder":
                EmitUnder(element, output);
                break;
            case "munderover":
                EmitUnderOver(element, output);
                break;
            case "mtable":
                EmitMatrix(element, output);
                break;
            case "mtr":
            case "mlabeledtr":
            case "mtd":
                EmitContainerChildren(element, output, inheritedMathVariant: null);
                break;
            case "menclose":
                EmitEnclose(element, output);
                break;
            case "mmultiscripts":
                EmitMultiScripts(element, output);
                break;
            case "none":
            case "annotation":
            case "annotation-xml":
                break;
            default:
                // Unknown pure presentation wrappers are flattened. Unknown leaf
                // tokens are kept as text so common MathML extensions remain
                // editable instead of producing structurally invalid MTEF.
                if (element.Elements().Any())
                    EmitContainerChildren(element, output, inheritedMathVariant: null);
                else if (!string.IsNullOrEmpty(element.Value))
                    EmitText(element.Value, TypefaceText, output);
                break;
        }
    }

    private static IEnumerable<XNode> SignificantChildren(XElement element) =>
        element.Nodes().Where(node =>
            node is XElement
            || node is XText text && !string.IsNullOrWhiteSpace(text.Value));

    private static void EmitContainerChildren(
        XElement element,
        List<byte> output,
        string? inheritedMathVariant)
    {
        var children = new List<XNode>();
        foreach (var child in SignificantChildren(element))
        {
            if (child is XElement childElement
                && (childElement.Name.LocalName == "annotation"
                    || childElement.Name.LocalName == "annotation-xml"))
                continue;
            if (string.IsNullOrWhiteSpace(inheritedMathVariant)
                || child is not XElement styledChild)
            {
                children.Add(child);
                continue;
            }
            var clone = new XElement(styledChild);
            ApplyInheritedMathVariant(clone, inheritedMathVariant!);
            children.Add(clone);
        }

        for (var index = 0; index < children.Count; index++)
        {
            if (children[index] is XElement bigOperatorHead
                && TryResolveBigOperatorHead(bigOperatorHead, out var resolvedHead))
            {
                var mainNodes = new List<XNode>();
                var cursor = index + 1;
                while (cursor < children.Count && !IsBigOperatorMainBoundary(children[cursor]))
                {
                    mainNodes.Add(CloneNode(children[cursor]));
                    cursor++;
                }
                if (mainNodes.Count > 0)
                {
                    var main = new XElement("mrow", mainNodes);
                    EmitResolvedBigOperator(resolvedHead, main, output);
                    index = cursor - 1;
                    continue;
                }
            }
            EmitNode(children[index], output);
        }
    }

    private readonly struct ResolvedBigOperatorHead
    {
        public ResolvedBigOperatorHead(
            NativeIntegral integral,
            NativeBigOperator bigOperator,
            XElement? lower,
            XElement? upper)
        {
            Integral = integral;
            BigOperator = bigOperator;
            Lower = lower;
            Upper = upper;
        }

        public NativeIntegral Integral { get; }
        public NativeBigOperator BigOperator { get; }
        public XElement? Lower { get; }
        public XElement? Upper { get; }
        public bool IsIntegral => Integral.VariationKind != 0;
    }

    private static bool TryResolveBigOperatorHead(
        XElement element,
        out ResolvedBigOperatorHead resolved)
    {
        resolved = default;
        var local = element.Name.LocalName;
        var children = element.Elements().ToArray();
        XElement? baseElement = element;
        XElement? lower = null;
        XElement? upper = null;
        if (local is "msub" or "msup" or "msubsup" or "munder" or "mover" or "munderover")
        {
            if (children.Length < 2) return false;
            baseElement = children[0];
            if (local is "msub" or "munder") lower = children[1];
            else if (local is "msup" or "mover") upper = children[1];
            else
            {
                if (children.Length < 3) return false;
                lower = children[1];
                upper = children[2];
            }
        }
        else if (local != "mo")
        {
            return false;
        }

        if (TryResolveNativeIntegral(baseElement, out var integral))
        {
            resolved = new ResolvedBigOperatorHead(integral, default, lower, upper);
            return true;
        }
        if (TryResolveNativeBigOperator(
                baseElement,
                out var bigOperator,
                allowAmbiguousUnionIntersection: local != "mo"))
        {
            resolved = new ResolvedBigOperatorHead(default, bigOperator, lower, upper);
            return true;
        }
        return false;
    }

    private static void EmitResolvedBigOperator(
        ResolvedBigOperatorHead resolved,
        XElement main,
        List<byte> output)
    {
        if (resolved.IsIntegral)
            EmitIntegralTemplate(
                resolved.Integral,
                main,
                resolved.Lower,
                resolved.Upper,
                output);
        else
            EmitBigOperatorTemplate(
                resolved.BigOperator,
                main,
                resolved.Lower,
                resolved.Upper,
                output);
    }

    private static bool IsBigOperatorMainBoundary(XNode node)
    {
        if (node is not XElement element || element.Name.LocalName != "mo") return false;
        var value = element.Value.Trim();
        return value is "+" or "-" or "−" or "±" or "∓"
            or "=" or "≠" or "<" or ">" or "≤" or "≥" or "≈" or "≃" or "∼"
            or "≡" or "∝" or "∈" or "∉" or "⊂" or "⊆" or "⊃" or "⊇"
            or "," or ";";
    }

    private static XNode CloneNode(XNode node) => node switch
    {
        XElement element => new XElement(element),
        XText text => new XText(text.Value),
        _ => new XText(node.ToString()),
    };

    private static void ApplyInheritedMathVariant(XElement element, string mathVariant)
    {
        if (element.Name.LocalName is "mi" or "mn" or "mo" or "mtext")
        {
            if (element.Attribute("mathvariant") is null)
                element.SetAttributeValue("mathvariant", mathVariant);
            return;
        }
        foreach (var descendant in element.Descendants())
        {
            if (descendant.Name.LocalName is not ("mi" or "mn" or "mo" or "mtext"))
                continue;
            if (descendant.Attribute("mathvariant") is null)
                descendant.SetAttributeValue("mathvariant", mathVariant);
        }
    }

    private static void EmitFraction(XElement element, List<byte> output)
    {
        var children = element.Elements().ToArray();
        if (children.Length < 2)
            throw new InvalidDataException("MathML mfrac requires numerator and denominator.");

        var lineThickness = ((string?)element.Attribute("linethickness") ?? string.Empty).Trim();
        if (lineThickness.Length > 0 && lineThickness.StartsWith("0", StringComparison.Ordinal))
        {
            // MathJax represents \binom and related stacked expressions as an
            // mfrac with a zero rule. MathType's native equivalent is a centered
            // PILE, not the ordinary fraction template (which always draws a rule).
            output.AddRange(new byte[]
            {
                RecordPile,
                0,
                2, // horizontal alignment: centered
                1, // vertical alignment: centered
            });
            EmitLine(children[0], output);
            EmitLine(children[1], output);
            output.Add(RecordEnd);
            return;
        }

        output.AddRange(new byte[] { RecordTemplate, 0, TemplateFraction, 0, 0 });
        EmitLine(children[0], output);
        EmitLine(children[1], output);
        output.Add(RecordEnd);
        output.Add(RecordFull);
    }

    private static void EmitSquareRoot(XElement element, List<byte> output)
    {
        output.AddRange(new byte[] { RecordTemplate, 0, TemplateRoot, 0, 0 });
        EmitLineContents(SignificantChildren(element), output);
        output.Add(RecordSub);
        output.AddRange(new byte[] { RecordLine, LineNull });
        output.Add(RecordEnd);
        output.Add(RecordFull);
    }

    private static void EmitNthRoot(XElement element, List<byte> output)
    {
        var children = element.Elements().ToArray();
        if (children.Length < 2)
            throw new InvalidDataException("MathML mroot requires radicand and index.");
        output.AddRange(new byte[] { RecordTemplate, 0, TemplateRoot, 1, 0 });
        EmitLine(children[0], output);
        output.Add(RecordSub);
        EmitLine(children[1], output);
        output.Add(RecordEnd);
        output.Add(RecordFull);
    }

    private static void EmitScript(XElement element, byte selector, List<byte> output)
    {
        var children = element.Elements().ToArray();
        var required = selector == TemplateSubSup ? 3 : 2;
        if (children.Length < required)
            throw new InvalidDataException($"MathML {element.Name.LocalName} has too few children.");
        if (TryResolveNativeIntegral(children[0], out var integral))
        {
            EmitIntegralTemplate(
                integral,
                main: null,
                lower: selector is TemplateSub or TemplateSubSup ? children[1] : null,
                upper: selector == TemplateSup ? children[1] : selector == TemplateSubSup ? children[2] : null,
                output);
            return;
        }
        if (TryResolveNativeBigOperator(
                children[0],
                out var bigOperator,
                allowAmbiguousUnionIntersection: true))
        {
            EmitBigOperatorTemplate(
                bigOperator,
                main: null,
                lower: selector is TemplateSub or TemplateSubSup ? children[1] : null,
                upper: selector == TemplateSup ? children[1] : selector == TemplateSubSup ? children[2] : null,
                output);
            return;
        }
        if (IsNamedLimitOperator(children[0]))
        {
            EmitLimitTemplate(
                children[0],
                selector is TemplateSub or TemplateSubSup ? children[1] : null,
                selector == TemplateSup ? children[1] : selector == TemplateSubSup ? children[2] : null,
                output);
            return;
        }
        EmitNode(children[0], output);
        output.AddRange(new byte[] { RecordTemplate, 0, selector, 0, 0 });
        output.Add(RecordSub);
        if (selector == TemplateSub)
        {
            EmitLine(children[1], output);
            output.AddRange(new byte[] { RecordLine, LineNull });
        }
        else if (selector == TemplateSup)
        {
            output.AddRange(new byte[] { RecordLine, LineNull });
            EmitLine(children[1], output);
        }
        else
        {
            EmitLine(children[1], output);
            EmitLine(children[2], output);
        }
        output.Add(RecordEnd);
        output.Add(RecordFull);
    }

    private static void EmitOver(XElement element, List<byte> output)
    {
        var children = element.Elements().ToArray();
        if (children.Length < 2)
        {
            EmitContainerChildren(element, output, inheritedMathVariant: null);
            return;
        }
        if (TryResolveNativeIntegral(children[0], out var integral))
        {
            EmitIntegralTemplate(integral, main: null, lower: null, upper: children[1], output);
            return;
        }
        if (TryResolveNativeBigOperator(
                children[0],
                out var bigOperator,
                allowAmbiguousUnionIntersection: true))
        {
            EmitBigOperatorTemplate(
                bigOperator,
                main: null,
                lower: null,
                upper: children[1],
                output);
            return;
        }
        if (TryGetAnnotatedHorizontalBraceBody(children[0], top: true, out var annotatedBraceBody))
        {
            EmitHorizontalFenceTemplate(
                TemplateHorizontalBrace,
                top: true,
                annotatedBraceBody,
                children[1],
                output);
            return;
        }
        var over = NormalizeAccentMark(children[1].Value.Trim());
        var embellishmentBody = UnwrapSingleTokenAccentBody(children[0]);
        if (embellishmentBody is not null
            && TryEmitSingleCharacterEmbellishment(embellishmentBody, over, output)) return;
        switch (over)
        {
            case "\u00AF":
            case "\u203E":
            case "\u2015":
            case "\u02C9":
                EmitSingleSlotTemplate(TemplateOverbar, 0, children[0], output);
                return;
            case "→":
                EmitVectorTemplate(0x02, 0x20D7, children[0], output);
                return;
            case "←":
                EmitVectorTemplate(0x01, 0x20D6, children[0], output);
                return;
            case "↔":
                EmitVectorTemplate(0x03, 0x20E1, children[0], output);
                return;
            case "^":
            case "ˆ":
                EmitSingleSlotTemplate(TemplateHat, 0, children[0], output);
                return;
            case "~":
            case "˜":
                EmitSingleSlotTemplate(TemplateTilde, 0, children[0], output);
                return;
            case "⌢":
            case "⏜":
                EmitSingleSlotTemplate(TemplateArc, 0, children[0], output);
                return;
            case "⏞":
                EmitHorizontalFenceTemplate(
                    TemplateHorizontalBrace,
                    top: true,
                    children[0],
                    annotation: null,
                    output);
                return;
        }
        EmitNode(children[0], output);
        EmitTrailingScript(children[1], isSubscript: false, output);
    }

    private static void EmitUnder(XElement element, List<byte> output)
    {
        var children = element.Elements().ToArray();
        if (children.Length < 2)
        {
            EmitContainerChildren(element, output, inheritedMathVariant: null);
            return;
        }
        if (TryResolveNativeIntegral(children[0], out var integral))
        {
            EmitIntegralTemplate(integral, main: null, lower: children[1], upper: null, output);
            return;
        }
        if (TryResolveNativeBigOperator(
                children[0],
                out var bigOperator,
                allowAmbiguousUnionIntersection: true))
        {
            EmitBigOperatorTemplate(
                bigOperator,
                main: null,
                lower: children[1],
                upper: null,
                output);
            return;
        }
        if (TryGetAnnotatedHorizontalBraceBody(children[0], top: false, out var annotatedBraceBody))
        {
            EmitHorizontalFenceTemplate(
                TemplateHorizontalBrace,
                top: false,
                annotatedBraceBody,
                children[1],
                output);
            return;
        }
        if (IsNamedLimitOperator(children[0]))
        {
            EmitLimitTemplate(children[0], children[1], upper: null, output);
            return;
        }
        var under = children[1].Value.Trim();
        if (under is "_" or "¯" or "‾" or "―" or "ˉ" or "̅" or "̲")
        {
            EmitSingleSlotTemplate(TemplateUnderbar, 0, children[0], output);
            return;
        }
        if (under == "⏟")
        {
            EmitHorizontalFenceTemplate(
                TemplateHorizontalBrace,
                top: false,
                children[0],
                annotation: null,
                output);
            return;
        }
        EmitNode(children[0], output);
        EmitTrailingScript(children[1], isSubscript: true, output);
    }

    private static void EmitUnderOver(XElement element, List<byte> output)
    {
        var children = element.Elements().ToArray();
        if (children.Length < 3)
        {
            EmitContainerChildren(element, output, inheritedMathVariant: null);
            return;
        }
        if (TryResolveNativeIntegral(children[0], out var integral))
        {
            EmitIntegralTemplate(
                integral,
                main: null,
                lower: children[1],
                upper: children[2],
                output);
            return;
        }
        if (TryResolveNativeBigOperator(
                children[0],
                out var bigOperator,
                allowAmbiguousUnionIntersection: true))
        {
            EmitBigOperatorTemplate(
                bigOperator,
                main: null,
                lower: children[1],
                upper: children[2],
                output);
            return;
        }
        if (IsNamedLimitOperator(children[0]))
        {
            EmitLimitTemplate(children[0], children[1], children[2], output);
            return;
        }
        EmitNode(children[0], output);
        output.AddRange(new byte[] { RecordTemplate, 0, TemplateSubSup, 0, 0, RecordSub });
        EmitLine(children[1], output);
        EmitLine(children[2], output);
        output.Add(RecordEnd);
        output.Add(RecordFull);
    }

    private readonly struct NativeIntegral
    {
        public NativeIntegral(int variationKind, int integralCount)
        {
            VariationKind = variationKind;
            IntegralCount = integralCount;
        }

        public int VariationKind { get; }
        public int IntegralCount { get; }
    }

    private static bool TryResolveNativeIntegral(
        XElement element,
        out NativeIntegral integral)
    {
        integral = default;
        if (element.Name.LocalName != "mo") return false;
        integral = element.Value.Trim() switch
        {
            "∫" => new NativeIntegral(0x01, 1),
            "∬" => new NativeIntegral(0x02, 2),
            "∭" => new NativeIntegral(0x03, 3),
            "∮" => new NativeIntegral(0x05, 1),
            "∲" => new NativeIntegral(0x08, 1),
            "∳" => new NativeIntegral(0x0C, 1),
            _ => default,
        };
        return integral.VariationKind != 0;
    }

    private static void EmitIntegralTemplate(
        NativeIntegral integral,
        XElement? main,
        XElement? lower,
        XElement? upper,
        List<byte> output)
    {
        // MathType 7 uses selector 15 for the integral family. Genuine native
        // equations encode lower/upper presence in 0x10/0x20 and the integral
        // kind in the low nibble (1=single, 2=double, 3=triple, 5=contour).
        // A contour template also requires MathType Extra's loop adornment
        // character before the ordinary integral glyph. Variation 4 with only
        // the integral character makes MathPage access invalid native state.
        var variation = integral.VariationKind
            | (lower is null ? 0 : 0x10)
            | (upper is null ? 0 : 0x20);
        EmitTemplateHeader(TemplateIntegral, variation, output);
        if (main is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(main, output);
        // MathType stores integral limits in SUB size. SIZE records are stateful;
        // omitting this transition makes the lower/upper slots inherit FULL and
        // renders limits visibly too large even though the MTEF remains readable.
        if (lower is not null || upper is not null)
            output.Add(RecordSub);
        if (lower is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(lower, output);
        if (upper is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(upper, output);
        output.Add(RecordSym);
        if (integral.VariationKind == 0x05)
        {
            // Exact bytes emitted by genuine MathType 7 for the closed-loop
            // adornment: fnMTEXTRA, MTCode U+EE11, legacy position 0xD1.
            EmitScalar(
                0xEE11,
                TypefaceMtExtra,
                output,
                includeEncoded8: true,
                encoded8Override: 0xD1);
        }
        for (var index = 0; index < integral.IntegralCount; index++)
        {
            EmitScalar(
                0x222B,
                TypefaceSymbol,
                output,
                includeEncoded8: true,
                encoded8Override: 0xF2);
        }
        output.Add(RecordEnd);
        // SYM is a persistent MTEF typesize state. Genuine MathType BigOp
        // equations normally place the complete integrand in the template's
        // main slot, so there is no ordinary sibling text after the symbol.
        // VisualTeX can encounter presentation MathML where the integrand is a
        // following sibling; always restore FULL after the template so that
        // subsequent ordinary characters cannot inherit symbol size.
        output.Add(RecordFull);
    }

    private readonly struct NativeBigOperator
    {
        public NativeBigOperator(
            byte selector,
            int mtCode,
            int typeface,
            byte encoded8)
        {
            Selector = selector;
            MtCode = mtCode;
            Typeface = typeface;
            Encoded8 = encoded8;
        }

        public byte Selector { get; }
        public int MtCode { get; }
        public int Typeface { get; }
        public byte Encoded8 { get; }
    }

    private static bool TryResolveNativeBigOperator(
        XElement element,
        out NativeBigOperator bigOperator,
        bool allowAmbiguousUnionIntersection = false)
    {
        bigOperator = default;
        if (element.Name.LocalName != "mo") return false;
        var value = element.Value.Trim();
        var explicitBigOperator = string.Equals(
                ((string?)element.Attribute("data-mjx-texclass"))?.Trim(),
                "OP",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                ((string?)element.Attribute("movablelimits"))?.Trim(),
                "true",
                StringComparison.OrdinalIgnoreCase);
        bigOperator = value switch
        {
            "∑" => new NativeBigOperator(TemplateSum, 0x2211, TypefaceSymbol, 0xE5),
            "∏" => new NativeBigOperator(TemplateProduct, 0x220F, TypefaceSymbol, 0xD5),
            "∐" => new NativeBigOperator(TemplateCoproduct, 0x2210, TypefaceMtExtra, 0x43),
            "⋃" => new NativeBigOperator(TemplateUnion, 0x222A, TypefaceMtExtra, 0x55),
            "⋂" => new NativeBigOperator(TemplateIntersection, 0x2229, TypefaceMtExtra, 0x49),
            "∪" when allowAmbiguousUnionIntersection || explicitBigOperator
                => new NativeBigOperator(TemplateUnion, 0x222A, TypefaceMtExtra, 0x55),
            "∩" when allowAmbiguousUnionIntersection || explicitBigOperator
                => new NativeBigOperator(TemplateIntersection, 0x2229, TypefaceMtExtra, 0x49),
            _ => default,
        };
        return bigOperator.Selector != 0;
    }

    private static void EmitBigOperatorTemplate(
        NativeBigOperator bigOperator,
        XElement? main,
        XElement? lower,
        XElement? upper,
        List<byte> output)
    {
        // MathType 7 native BigOp storage observed from genuine Equation.DSMT4:
        // variation 0x70 and line slots main/lower/upper followed by SYM + CHAR.
        EmitTemplateHeader(bigOperator.Selector, 0x70, output);
        if (main is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(main, output);
        // Sum/product/union/intersection limits use the same script-size state as
        // native MathType. Without SUB, both limits are interpreted at FULL size.
        if (lower is not null || upper is not null)
            output.Add(RecordSub);
        if (lower is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(lower, output);
        if (upper is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(upper, output);
        output.Add(RecordSym);
        EmitScalar(
            bigOperator.MtCode,
            bigOperator.Typeface,
            output,
            includeEncoded8: true,
            encoded8Override: bigOperator.Encoded8);
        output.Add(RecordEnd);
        // Do not leak the SYM typesize used by the large operator glyph into
        // following siblings. MathType treats FULL/SUB/SYM as size state.
        output.Add(RecordFull);
    }

    private static bool IsNamedLimitOperator(XElement element)
    {
        if (element.Name.LocalName is not ("mi" or "mo")) return false;
        var value = element.Value.Trim();
        if (value.Length <= 1 || !value.All(char.IsLetter)) return false;
        var movableLimits = string.Equals(
            (string?)element.Attribute("movablelimits"),
            "true",
            StringComparison.OrdinalIgnoreCase);
        var texClass = ((string?)element.Attribute("data-mjx-texclass") ?? string.Empty).Trim();
        var variant = ((string?)element.Attribute("mathvariant") ?? string.Empty).Trim();
        return movableLimits
            || string.Equals(texClass, "OP", StringComparison.OrdinalIgnoreCase)
            || variant.IndexOf("normal", StringComparison.OrdinalIgnoreCase) >= 0
            || variant.IndexOf("upright", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static void EmitLimitTemplate(
        XElement operatorElement,
        XElement? lower,
        XElement? upper,
        List<byte> output)
    {
        // Genuine MathType 7 stores named operators such as max/min/lim/sup/inf
        // as function-style CHAR records followed by the ordinary sub/sup
        // template. Writing selector 23 (tmLIM) here produced Equation Native
        // streams that our parser accepted but MathType could hang while opening.
        var name = operatorElement.Value.Trim();
        EmitFunctionRun(name, output);
        if (lower is null && upper is null) return;

        var selector = lower is not null && upper is not null
            ? TemplateSubSup
            : lower is not null
                ? TemplateSub
                : TemplateSup;
        output.AddRange(new byte[]
        {
            RecordTemplate,
            0,
            selector,
            0,
            0,
            RecordSub,
        });
        if (lower is null)
            output.AddRange(new byte[] { RecordLine, LineNull });
        else
            EmitLine(lower, output);
        if (upper is null)
            output.AddRange(new byte[] { RecordLine, LineNull });
        else
            EmitLine(upper, output);
        output.Add(RecordEnd);
        output.Add(RecordFull);
    }

    private static void EmitFunctionRun(string name, List<byte> output)
    {
        var functionScalars = EnumerateBmpScalars(name).ToArray();
        for (var index = 0; index < functionScalars.Length; index++)
        {
            EmitScalar(
                functionScalars[index],
                TypefaceFunction,
                output,
                functionStart: index == 0);
        }
    }

    private static void EmitTrailingScript(
        XElement script,
        bool isSubscript,
        List<byte> output)
    {
        output.AddRange(new byte[]
        {
            RecordTemplate,
            0,
            isSubscript ? TemplateSub : TemplateSup,
            0,
            0,
            RecordSub,
        });
        if (isSubscript)
        {
            EmitLine(script, output);
            output.AddRange(new byte[] { RecordLine, LineNull });
        }
        else
        {
            output.AddRange(new byte[] { RecordLine, LineNull });
            EmitLine(script, output);
        }
        output.Add(RecordEnd);
        output.Add(RecordFull);
    }

    private static void EmitSingleSlotTemplate(
        byte selector,
        int variation,
        XElement body,
        List<byte> output)
    {
        EmitTemplateHeader(selector, variation, output);
        EmitLine(body, output);
        output.Add(RecordEnd);
        output.Add(RecordFull);
    }

    private static void EmitVectorTemplate(
        int variation,
        int combiningArrow,
        XElement body,
        List<byte> output)
    {
        // Genuine MathType 7 tmVEC HatBox records are not a generic one-slot
        // template. After the body line MathType persists the selected combining
        // arrow as an explicit MTCode character using its expanding-glyph
        // typeface. Omitting this trailing character makes Word accept the CFB
        // initially but reject/rematerialize its OLE presentation for multi-token
        // vectors such as \overrightarrow{AB} / \overleftrightarrow{AB}.
        EmitTemplateHeader(TemplateVector, variation, output);
        EmitColor(0, output);
        EmitLine(body, output);
        EmitScalar(
            combiningArrow,
            TypefaceFence,
            output,
            includeEncoded8: false);
        output.Add(RecordEnd);
        output.Add(RecordFull);
    }

    private static bool TryGetAnnotatedHorizontalBraceBody(
        XElement candidate,
        bool top,
        out XElement body)
    {
        body = null!;
        var current = candidate;
        while (current.Name.LocalName == "mrow")
        {
            var wrapped = current.Elements().ToArray();
            if (wrapped.Length != 1) break;
            current = wrapped[0];
        }
        if (current.Name.LocalName != (top ? "mover" : "munder")) return false;
        var inner = current.Elements().ToArray();
        if (inner.Length < 2 || inner[1].Name.LocalName != "mo") return false;
        var marker = inner[1].Value.Trim();
        var expected = top ? "⏞" : "⏟";
        if (!IsHorizontalBraceMarker(marker, top)) return false;
        body = inner[0];
        return true;
    }

    private static void EmitHorizontalFenceTemplate(
        byte selector,
        bool top,
        XElement body,
        XElement? annotation,
        List<byte> output)
    {
        EmitTemplateHeader(selector, top ? 1 : 0, output);

        // Match MathType 7's native HorizontalBrace/HorizontalBracket layout.
        // A real Equation.DSMT4 stores the main body at full size, switches to
        // SUB for the annotation, restores FULL, then writes the private fence
        // glyph using the expanding-fence typeface. Writing the public Unicode
        // U+23DE/U+23DF operators here produces an MTEF stream that VisualTeX can
        // parse itself but that MathType may reject while activating the OLE.
        EmitColor(0, output);
        EmitMathTypeColoredLine(body, output, resetColorAfter: false);
        output.Add(RecordSub);
        EmitColor(0, output);
        if (annotation is null)
            output.AddRange(new byte[] { RecordLine, LineNull });
        else
            EmitMathTypeColoredLine(annotation, output, resetColorAfter: false);
        output.Add(RecordFull);
        EmitScalar(
            top ? 0xFE37 : 0xFE38,
            TypefaceFence,
            output,
            includeEncoded8: false);
        output.Add(RecordEnd);
    }

    private static void EmitTemplateHeader(byte selector, int variation, List<byte> output)
    {
        output.Add(RecordTemplate);
        output.Add(0);
        output.Add(selector);
        if (variation < 0x80)
        {
            output.Add((byte)variation);
        }
        else
        {
            output.Add((byte)(0x80 | (variation & 0x7F)));
            output.Add((byte)(variation >> 8));
        }
        output.Add(0);
    }

    private static XElement? UnwrapSingleTokenAccentBody(XElement body)
    {
        var current = body;
        while (current.Name.LocalName == "mrow")
        {
            var children = current.Elements().ToArray();
            if (children.Length != 1) break;
            current = children[0];
        }
        return current.Name.LocalName is "mi" or "mn" or "mo" ? current : null;
    }

    private static string NormalizeAccentMark(string mark) => mark switch
    {
        "\u0305" or "\u203E" or "\u2015" or "\u02C9" => "¯",
        "\u0302" or "ˆ" => "^",
        "\u0303" or "˜" => "~",
        "\u20D7" => "→",
        "\u20D6" => "←",
        "\u20E1" => "↔",
        "." or "\u0307" => "˙",
        "\u0308" => "¨",
        "\u030C" => "ˇ",
        "\u0306" => "˘",
        "\u0301" => "´",
        "\u0300" => "`",
        "\u030A" => "˚",
        _ => mark,
    };

    private static bool TryEmitSingleCharacterEmbellishment(
        XElement body,
        string over,
        List<byte> output)
    {
        var code = NormalizeAccentMark(over) switch
        {
            "." or "˙" => EmbellDot,
            "¨" => EmbellDoubleDot,
            "~" or "˜" => EmbellTilde,
            "^" or "ˆ" => EmbellHat,
            "→" => EmbellRightArrow,
            "←" => EmbellLeftArrow,
            "↔" => EmbellBothArrow,
            "\u00AF" or "\u203E" or "\u2015" or "\u02C9" => EmbellOverbar,
            _ => (byte)0,
        };
        if (code == 0 || body.Name.LocalName is not ("mi" or "mn" or "mo"))
            return false;
        var scalars = EnumerateBmpScalars(body.Value).ToArray();
        if (scalars.Length != 1) return false;
        var typeface = ResolveTokenTypeface(body, scalars[0]);
        EmitScalar(
            scalars[0],
            typeface,
            output,
            includeEncoded8: body.Name.LocalName == "mo" && scalars[0] <= 0xFF,
            embellishments: new[] { code });
        return true;
    }

    private static void EmitMatrix(XElement element, List<byte> output)
    {
        var rows = element.Elements()
            .Where(row => row.Name.LocalName is "mtr" or "mlabeledtr")
            .ToArray();
        if (rows.Length == 0)
        {
            output.AddRange(new byte[] { RecordMatrix, 0, 1, 2, 1, 1, 1, 0, 0 });
            output.AddRange(new byte[] { RecordLine, LineNull, RecordEnd, RecordFull });
            return;
        }
        var cells = rows.Select(row => row.Elements()
                .Where(cell => cell.Name.LocalName == "mtd")
                .ToArray())
            .ToArray();
        var columnCount = Math.Max(1, cells.Max(row => row.Length));
        if (rows.Length > byte.MaxValue || columnCount > byte.MaxValue)
            throw new InvalidDataException("MathType MTEF matrix exceeds 255 rows or columns.");

        output.AddRange(new byte[]
        {
            RecordMatrix,
            0,
            1, // vertical alignment: center line
            1, // horizontal cell justification: centered (MathType 7 native value)
            1, // vertical cell justification: baseline center
            (byte)rows.Length,
            (byte)columnCount,
        });
        for (var index = 0; index < PartitionByteCount(rows.Length + 1); index++) output.Add(0);
        for (var index = 0; index < PartitionByteCount(columnCount + 1); index++) output.Add(0);
        for (var row = 0; row < rows.Length; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                var isLastCell = row == rows.Length - 1 && column == columnCount - 1;
                if (column < cells[row].Length)
                    EmitMathTypeColoredLine(cells[row][column], output, resetColorAfter: !isLastCell);
                else output.AddRange(new byte[] { RecordLine, LineNull });
            }
        }
        output.Add(RecordEnd);
    }

    private static int PartitionByteCount(int partitionCount) =>
        Math.Max(1, (partitionCount + 3) / 4);

    private static void EmitEnclose(XElement element, List<byte> output)
    {
        var notation = ((string?)element.Attribute("notation") ?? string.Empty)
            .ToLowerInvariant();
        var significant = SignificantChildren(element).ToArray();
        var body = significant.Length == 1 && significant[0] is XElement singleElement
            ? new XElement(singleElement)
            : new XElement("mrow", significant);
        if (notation.Contains("radical"))
        {
            EmitSquareRoot(new XElement("msqrt", new XElement(body)), output);
            return;
        }
        if (notation.Contains("box"))
        {
            EmitSingleSlotTemplate(TemplateBox, 0x1E, body, output);
            return;
        }
        var strikeVariation = 0;
        if (notation.Contains("horizontalstrike")) strikeVariation |= 0x01;
        if (notation.Contains("updiagonalstrike")) strikeVariation |= 0x02;
        if (notation.Contains("downdiagonalstrike")) strikeVariation |= 0x04;
        if (strikeVariation != 0)
        {
            EmitSingleSlotTemplate(TemplateStrike, strikeVariation, body, output);
            return;
        }
        EmitNode(body, output);
    }

    private static void EmitMultiScripts(XElement element, List<byte> output)
    {
        var children = element.Elements().ToArray();
        if (children.Length == 0) return;
        EmitNode(children[0], output);
        var index = 1;
        while (index < children.Length
            && children[index].Name.LocalName != "mprescripts")
        {
            var sub = children[index];
            var sup = index + 1 < children.Length ? children[index + 1] : null;
            if (sub.Name.LocalName == "none" && (sup is null || sup.Name.LocalName == "none"))
            {
                index += 2;
                continue;
            }
            if (sub.Name.LocalName == "none" && sup is not null)
                EmitTrailingScript(sup, isSubscript: false, output);
            else if (sup is null || sup.Name.LocalName == "none")
                EmitTrailingScript(sub, isSubscript: true, output);
            else
            {
                output.AddRange(new byte[] { RecordTemplate, 0, TemplateSubSup, 0, 0, RecordSub });
                EmitLine(sub, output);
                EmitLine(sup, output);
                output.Add(RecordEnd);
                output.Add(RecordFull);
            }
            index += 2;
        }
        // Prescripts are uncommon in Office equations. Keep them editable rather
        // than rejecting the whole formula; they are appended as ordinary scripts
        // until a dedicated preceding-script writer is needed.
        if (index < children.Length && children[index].Name.LocalName == "mprescripts")
        {
            index++;
            while (index < children.Length)
            {
                var sub = children[index];
                var sup = index + 1 < children.Length ? children[index + 1] : null;
                if (sub.Name.LocalName != "none") EmitTrailingScript(sub, true, output);
                if (sup is not null && sup.Name.LocalName != "none") EmitTrailingScript(sup, false, output);
                index += 2;
            }
        }
    }

    private static void EmitFenced(XElement element, List<byte> output)
    {
        var open = NormalizeFence((string?)element.Attribute("open") ?? "(");
        var close = NormalizeFence((string?)element.Attribute("close") ?? ")");
        var selector = SelectFenceTemplate(open, close);
        if (selector is null)
        {
            if (!string.IsNullOrEmpty(open)) EmitOperator(open, output);
            foreach (var child in SignificantChildren(element)) EmitNode(child, output);
            if (!string.IsNullOrEmpty(close)) EmitOperator(close, output);
            return;
        }
        var variation = (string.IsNullOrEmpty(open) ? 0 : 1)
            | (string.IsNullOrEmpty(close) ? 0 : 2);
        EmitTemplateHeader(selector.Value, variation, output);
        EmitColor(0, output);
        EmitLineContents(SignificantChildren(element), output);
        if (!string.IsNullOrEmpty(open)) EmitFenceCharacter(open, output);
        if (!string.IsNullOrEmpty(close)) EmitFenceCharacter(close, output);
        output.Add(RecordEnd);
    }

    private static void EmitColor(int index, List<byte> output)
    {
        output.Add(RecordColor);
        WriteUnsigned(output, index);
    }

    private static void EmitMathTypeColoredLine(
        XElement element,
        List<byte> output,
        bool resetColorAfter)
    {
        output.AddRange(new byte[] { RecordLine, 0 });
        EmitColor(1, output);
        EmitNode(element, output);
        output.Add(RecordEnd);
        if (resetColorAfter) EmitColor(0, output);
    }

    private static void EmitFenceCharacter(string value, List<byte> output)
    {
        foreach (var scalar in EnumerateBmpScalars(value))
            EmitScalar(scalar, TypefaceFence, output, includeEncoded8: false);
    }

    private static string NormalizeFence(string value) => value switch
    {
        "." => string.Empty,
        "〈" or "〈" or "⟨" => "⟨",
        "〉" or "〉" or "⟩" => "⟩",
        _ => value,
    };

    private static byte? SelectFenceTemplate(string open, string close)
    {
        var probe = !string.IsNullOrEmpty(open) ? open : close;
        if ((open is "⟨" or "〈" || string.IsNullOrEmpty(open))
            && (close is "⟩" or "〉" || string.IsNullOrEmpty(close))
            && probe is "⟨" or "〈" or "⟩" or "〉") return TemplateAngle;
        if ((open == "(" || string.IsNullOrEmpty(open))
            && (close == ")" || string.IsNullOrEmpty(close))
            && probe is "(" or ")") return TemplateParen;
        if ((open == "[" || string.IsNullOrEmpty(open))
            && (close == "]" || string.IsNullOrEmpty(close))
            && probe is "[" or "]") return TemplateBracket;
        if ((open == "{" || string.IsNullOrEmpty(open))
            && (close == "}" || string.IsNullOrEmpty(close))
            && probe is "{" or "}") return TemplateBrace;
        if ((open == "|" || string.IsNullOrEmpty(open))
            && (close == "|" || string.IsNullOrEmpty(close))
            && probe == "|") return TemplateBar;
        if ((open is "‖" or "||" || string.IsNullOrEmpty(open))
            && (close is "‖" or "||" || string.IsNullOrEmpty(close))
            && probe is "‖" or "||") return TemplateDoubleBar;
        if ((open == "⌊" || string.IsNullOrEmpty(open))
            && (close == "⌋" || string.IsNullOrEmpty(close))
            && probe is "⌊" or "⌋") return TemplateFloor;
        if ((open == "⌈" || string.IsNullOrEmpty(open))
            && (close == "⌉" || string.IsNullOrEmpty(close))
            && probe is "⌈" or "⌉") return TemplateCeiling;
        return null;
    }

    private static void EmitLine(XElement element, List<byte> output)
    {
        output.AddRange(new byte[] { RecordLine, 0 });
        EmitNode(element, output);
        output.Add(RecordEnd);
    }

    private static void EmitLineContents(IEnumerable<XNode> nodes, List<byte> output)
    {
        output.AddRange(new byte[] { RecordLine, 0 });
        foreach (var node in nodes) EmitNode(node, output);
        output.Add(RecordEnd);
    }

    private static void EmitIdentifier(XElement element, List<byte> output)
    {
        var value = element.Value;
        if (string.IsNullOrEmpty(value)) return;
        var scalars = EnumerateBmpScalars(value).ToArray();
        var variant = ((string?)element.Attribute("mathvariant") ?? string.Empty)
            .Trim()
            .ToLowerInvariant();
        var isFunction = scalars.Length > 1 && scalars.All(IsLetterScalar);
        var explicitTypeface = int.TryParse(
            (string?)element.Attribute("data-mtef-explicit-typeface"),
            out var parsedExplicitTypeface)
            ? parsedExplicitTypeface
            : (int?)null;
        foreach (var originalScalar in scalars)
        {
            // Word's OMML reverse transform represents a single prime in a
            // superscript as ASCII apostrophe, while MathJax/MathType use the
            // mathematical prime U+2032. Normalize before typeface/glyph lookup
            // so OMML→MathType writes MathType's native Symbol prime rather than
            // an italic variable apostrophe.
            var scalar = originalScalar == '\'' ? 0x2032 : originalScalar;
            var effectiveVariant = variant;
            if (string.IsNullOrWhiteSpace(effectiveVariant)
                && TryNormalizeLetterlikeForWrite(
                    originalScalar,
                    out var normalizedScalar,
                    out var normalizedVariant))
            {
                scalar = normalizedScalar;
                effectiveVariant = normalizedVariant;
            }

            if (scalar == 0x210F)
            {
                // Genuine MathType 7 persists \hbar / U+210F as an MT Extra
                // character, with the legacy 8-bit position 'h'. Writing the
                // same MTCode using Variable/Text typeface leaves MathType with
                // no glyph in that font and it renders an unknown-symbol box.
                // Keep this before mathvariant handling: MathJax commonly marks
                // the hbar <mi> as normal/upright, but MathType's native MTEF
                // representation is still fnMTEXTRA + encoded8 0x68.
                EmitScalar(
                    0x210F,
                    TypefaceMtExtra,
                    output,
                    includeEncoded8: true,
                    encoded8Override: 0x68);
                continue;
            }

            if (effectiveVariant.Contains("double-struck")
                && TryMtExtraDoubleStruck(scalar, out var mtExtraCode, out var mtExtraPosition))
            {
                // MathType 7 stores the five standard blackboard-bold set letters
                // in its built-in Extra Math style (fnMTEXTRA), not as an explicit
                // Euclid Math Two font. This exactly matches native MathType MTEF.
                EmitScalar(
                    mtExtraCode,
                    TypefaceMtExtra,
                    output,
                    includeEncoded8: true,
                    encoded8Override: mtExtraPosition);
                continue;
            }

            if (explicitTypeface is null
                && !effectiveVariant.Contains("bold")
                && !effectiveVariant.Contains("double-struck")
                && !effectiveVariant.Contains("script")
                && !effectiveVariant.Contains("fraktur")
                && TryResolveLegacyMathTypeGlyph(
                    scalar,
                    out var nativeMtCode,
                    out var nativeTypeface,
                    out var nativeEncoded8))
            {
                EmitScalar(
                    nativeMtCode,
                    nativeTypeface,
                    output,
                    includeEncoded8: true,
                    encoded8Override: nativeEncoded8);
                continue;
            }

            if (explicitTypeface is null
                && IsMathematicalSymbolScalar(scalar))
            {
                // MathType's built-in Symbol style is an 8-bit Adobe Symbol
                // encoding, not a generic Unicode font.  If a mathematical
                // symbol has no legacy Symbol position, keep its Unicode MTCode
                // in the text font instead of emitting a non-existent Symbol
                // glyph (which MathType renders as '?'/unknown-symbol).
                EmitScalar(
                    NormalizeMathTypeMtCode(scalar),
                    TypefaceText,
                    output);
                continue;
            }

            var typeface = explicitTypeface
                ?? ResolveTokenTypeface(element, scalar, effectiveVariant, isFunction);
            var mtCode = explicitTypeface is < 0
                ? ExplicitVariantMtCode(effectiveVariant, scalar)
                : effectiveVariant.Contains("double-struck")
                    && TryStandardDoubleStruckMtCode(scalar, out var standardDoubleStruck)
                        ? standardDoubleStruck
                        : scalar;
            EmitScalar(
                mtCode,
                typeface,
                output,
                includeEncoded8: explicitTypeface is < 0 && scalar <= 0xFF,
                encoded8Override: explicitTypeface is < 0 && scalar <= 0xFF
                    ? (byte)scalar
                    : null);
        }
    }

    private static bool CanUseStandardDoubleStruckMtCode(int scalar) =>
        TryStandardDoubleStruckMtCode(scalar, out _);

    private static bool TryNormalizeLetterlikeForWrite(
        int scalar,
        out int normalizedScalar,
        out string mathVariant)
    {
        switch (scalar)
        {
            case 0x2102: normalizedScalar = 'C'; mathVariant = "double-struck"; return true;
            case 0x210B: normalizedScalar = 'H'; mathVariant = "script"; return true;
            case 0x210C: normalizedScalar = 'H'; mathVariant = "fraktur"; return true;
            case 0x210D: normalizedScalar = 'H'; mathVariant = "double-struck"; return true;
            case 0x2110: normalizedScalar = 'I'; mathVariant = "script"; return true;
            case 0x2111: normalizedScalar = 'I'; mathVariant = "fraktur"; return true;
            case 0x2112: normalizedScalar = 'L'; mathVariant = "script"; return true;
            case 0x2115: normalizedScalar = 'N'; mathVariant = "double-struck"; return true;
            case 0x2119: normalizedScalar = 'P'; mathVariant = "double-struck"; return true;
            case 0x211A: normalizedScalar = 'Q'; mathVariant = "double-struck"; return true;
            case 0x211B: normalizedScalar = 'R'; mathVariant = "script"; return true;
            case 0x211C: normalizedScalar = 'R'; mathVariant = "fraktur"; return true;
            case 0x211D: normalizedScalar = 'R'; mathVariant = "double-struck"; return true;
            case 0x2124: normalizedScalar = 'Z'; mathVariant = "double-struck"; return true;
            case 0x212C: normalizedScalar = 'B'; mathVariant = "script"; return true;
            case 0x212D: normalizedScalar = 'C'; mathVariant = "fraktur"; return true;
            case 0x212F: normalizedScalar = 'e'; mathVariant = "script"; return true;
            case 0x2130: normalizedScalar = 'E'; mathVariant = "script"; return true;
            case 0x2131: normalizedScalar = 'F'; mathVariant = "script"; return true;
            case 0x2133: normalizedScalar = 'M'; mathVariant = "script"; return true;
            case 0x2134: normalizedScalar = 'o'; mathVariant = "script"; return true;
            default:
                normalizedScalar = scalar;
                mathVariant = string.Empty;
                return false;
        }
    }

    private static bool TryMtExtraDoubleStruck(
        int scalar,
        out int mtCode,
        out byte fontPosition)
    {
        switch (scalar)
        {
            case 'R': mtCode = 0x211D; fontPosition = 0xA1; return true;
            case 'Z': mtCode = 0x2124; fontPosition = 0xA2; return true;
            case 'C': mtCode = 0x2102; fontPosition = 0xA3; return true;
            case 'Q': mtCode = 0x211A; fontPosition = 0xA4; return true;
            case 'N': mtCode = 0x2115; fontPosition = 0xA5; return true;
            default:
                mtCode = 0;
                fontPosition = 0;
                return false;
        }
    }

    private static bool TryStandardDoubleStruckMtCode(int scalar, out int mtCode)
    {
        mtCode = scalar switch
        {
            'C' => 0x2102,
            'H' => 0x210D,
            'N' => 0x2115,
            'P' => 0x2119,
            'Q' => 0x211A,
            'R' => 0x211D,
            'Z' => 0x2124,
            _ => 0,
        };
        return mtCode != 0;
    }

    private static int ExplicitVariantMtCode(string variant, int scalar)
    {
        if (variant.Contains("double-struck"))
        {
            // MathType's EuclidMath2 encoding uses its own BMP PUA MTCode
            // range, even for the familiar Unicode letterlike symbols such as
            // ℝ. These values are the mappings shipped in MathType 7's own
            // AMS/Desire2Learn translators.
            if (scalar is >= 'A' and <= 'Z') return 0xF080 + scalar - 'A';
            if (scalar is >= 'a' and <= 'z') return 0xF09A + scalar - 'a';
            if (scalar is >= '0' and <= '9') return 0xF0C0 + scalar - '0';
            return scalar;
        }
        if (variant.Contains("script"))
        {
            return scalar switch
            {
                'B' => 0x212C,
                'E' => 0x2130,
                'F' => 0x2131,
                'H' => 0x210B,
                'I' => 0x2110,
                'L' => 0x2112,
                'M' => 0x2133,
                'R' => 0x211B,
                'e' => 0x212F,
                'g' => 0x210A,
                'o' => 0x2134,
                _ => scalar,
            };
        }
        if (variant.Contains("fraktur"))
        {
            return scalar switch
            {
                'C' => 0x212D,
                'H' => 0x210C,
                'I' => 0x2111,
                'R' => 0x211C,
                'Z' => 0x2128,
                _ => scalar,
            };
        }
        return scalar;
    }

    private static int ResolveTokenTypeface(
        XElement element,
        int scalar,
        string? normalizedVariant = null,
        bool? functionHint = null)
    {
        var variant = normalizedVariant
            ?? (((string?)element.Attribute("mathvariant") ?? string.Empty)
                .Trim()
                .ToLowerInvariant());
        var tokenKind = element.Name.LocalName;

        // MTEF typeface is semantic, not merely visual.  In particular MathJax
        // can attach mathvariant="normal" to operator tokens such as infinity,
        // relations, arrows and set symbols.  Treating that visual hint as Text
        // changes the token to mtext on read-back (for example \infty becomes
        // \text{∞}) and makes otherwise-valid MathType MTEF fail semantic
        // round-trip validation.  Preserve the MathML token class first; apply
        // mathvariant only inside identifier-like tokens.
        if (tokenKind == "mtext") return TypefaceText;
        if (tokenKind == "mn") return TypefaceNumber;
        if (tokenKind == "mo")
            return functionHint == true ? TypefaceFunction : TypefaceSymbol;
        if (IsMathematicalSymbolScalar(scalar)) return TypefaceSymbol;

        if (variant.Contains("normal") || variant.Contains("upright"))
            return functionHint == true ? TypefaceFunction : TypefaceText;
        if (variant.Contains("bold")) return TypefaceVector;
        if (variant.Contains("italic"))
            return IsLowerGreek(scalar)
                ? TypefaceLowerGreek
                : IsUpperGreek(scalar)
                    ? TypefaceUpperGreek
                    : TypefaceVariable;
        if (functionHint == true) return TypefaceFunction;
        if (IsLowerGreek(scalar)) return TypefaceLowerGreek;
        if (IsUpperGreek(scalar)) return TypefaceUpperGreek;
        return TypefaceVariable;
    }

    private static bool IsMathematicalSymbolScalar(int scalar)
    {
        if (scalar < 0 || scalar > char.MaxValue) return false;
        if (scalar is 0x2020 or 0x2021
            or 0x2032 or 0x2033 or 0x2034
            // Unicode classifies floor/ceiling as opening/closing punctuation,
            // not MathSymbol. MathJax emits the common shorthand
            // \lfloor x\rfloor / \lceil x\rceil as standalone <mo> tokens, so
            // an MTEF Text-font fallback must still read them back as operators.
            // Otherwise they become mtext and semantic round-trip validation
            // rejects an otherwise valid MathType equation.
            or 0x2308 or 0x2309 or 0x230A or 0x230B
            or 0x2329 or 0x232A
            or 0x27E8 or 0x27E9)
            return true;
        var character = (char)scalar;
        var category = char.GetUnicodeCategory(character);
        return category == System.Globalization.UnicodeCategory.MathSymbol
            || category == System.Globalization.UnicodeCategory.CurrencySymbol
            || category == System.Globalization.UnicodeCategory.ModifierSymbol
            || category == System.Globalization.UnicodeCategory.OtherSymbol;
    }

    private static string NormalizeOperatorToken(string value) => value switch
    {
        "〈" or "〈" or "⟨" => "⟨",
        "〉" or "〉" or "⟩" => "⟩",
        // MathJax emits the mathematical minus U+2212 while Word OMML commonly
        // serializes the same binary subtraction operator as ASCII hyphen-minus.
        // They are semantically identical in an <mo> token and must compare equal
        // across VisualTeX/OMML/MathType round trips.
        "−" => "-",
        _ => value,
    };

    private static int NormalizeMathTypeMtCode(int scalar) => scalar switch
    {
        // MathType 7/Adobe Symbol use the historical angle-bracket MTCode pair.
        0x27E8 => 0x2329,
        0x27E9 => 0x232A,
        _ => scalar,
    };

    private static bool TryResolveLegacyMathTypeGlyph(
        int scalar,
        out int mtCode,
        out int typeface,
        out byte encoded8)
    {
        mtCode = NormalizeMathTypeMtCode(scalar);
        typeface = TypefaceSymbol;
        encoded8 = 0;

        if (mtCode == 0x2223)
        {
            // \mid uses the ordinary vertical-bar position in Symbol while its
            // MTCode remains U+2223 so semantic readback is not degraded to '|'.
            encoded8 = 0x7C;
            return true;
        }

        if (!TryGetAdobeSymbolEncoded8(mtCode, out encoded8)) return false;
        if (IsLowerGreek(mtCode)) typeface = TypefaceLowerGreek;
        else if (IsUpperGreek(mtCode)) typeface = TypefaceUpperGreek;
        return true;
    }

    private static bool TryGetAdobeSymbolEncoded8(int scalar, out byte encoded8)
    {
        // Static Adobe Symbol encoding used by MathType's built-in Lower Greek,
        // Upper Greek and Symbol styles.  Keeping this table in VisualTeX makes
        // standalone MTEF generation independent of MathType/MathPage at runtime.
        switch (scalar)
        {
            case 0x00AC: encoded8 = 0xD8; return true;
            case 0x00B0: encoded8 = 0xB0; return true;
            case 0x00B1: encoded8 = 0xB1; return true;
            case 0x00B5: encoded8 = 0x6D; return true;
            case 0x00D7: encoded8 = 0xB4; return true;
            case 0x00F7: encoded8 = 0xB8; return true;
            case 0x0192: encoded8 = 0xA6; return true;
            case 0x0391: encoded8 = 0x41; return true;
            case 0x0392: encoded8 = 0x42; return true;
            case 0x0393: encoded8 = 0x47; return true;
            case 0x0394: encoded8 = 0x44; return true;
            case 0x0395: encoded8 = 0x45; return true;
            case 0x0396: encoded8 = 0x5A; return true;
            case 0x0397: encoded8 = 0x48; return true;
            case 0x0398: encoded8 = 0x51; return true;
            case 0x0399: encoded8 = 0x49; return true;
            case 0x039A: encoded8 = 0x4B; return true;
            case 0x039B: encoded8 = 0x4C; return true;
            case 0x039C: encoded8 = 0x4D; return true;
            case 0x039D: encoded8 = 0x4E; return true;
            case 0x039E: encoded8 = 0x58; return true;
            case 0x039F: encoded8 = 0x4F; return true;
            case 0x03A0: encoded8 = 0x50; return true;
            case 0x03A1: encoded8 = 0x52; return true;
            case 0x03A3: encoded8 = 0x53; return true;
            case 0x03A4: encoded8 = 0x54; return true;
            case 0x03A5: encoded8 = 0x55; return true;
            case 0x03A6: encoded8 = 0x46; return true;
            case 0x03A7: encoded8 = 0x43; return true;
            case 0x03A8: encoded8 = 0x59; return true;
            case 0x03A9: encoded8 = 0x57; return true;
            case 0x03B1: encoded8 = 0x61; return true;
            case 0x03B2: encoded8 = 0x62; return true;
            case 0x03B3: encoded8 = 0x67; return true;
            case 0x03B4: encoded8 = 0x64; return true;
            case 0x03B5: encoded8 = 0x65; return true;
            case 0x03B6: encoded8 = 0x7A; return true;
            case 0x03B7: encoded8 = 0x68; return true;
            case 0x03B8: encoded8 = 0x71; return true;
            case 0x03B9: encoded8 = 0x69; return true;
            case 0x03BA: encoded8 = 0x6B; return true;
            case 0x03BB: encoded8 = 0x6C; return true;
            case 0x03BC: encoded8 = 0x6D; return true;
            case 0x03BD: encoded8 = 0x6E; return true;
            case 0x03BE: encoded8 = 0x78; return true;
            case 0x03BF: encoded8 = 0x6F; return true;
            case 0x03C0: encoded8 = 0x70; return true;
            case 0x03C1: encoded8 = 0x72; return true;
            case 0x03C2: encoded8 = 0x56; return true;
            case 0x03C3: encoded8 = 0x73; return true;
            case 0x03C4: encoded8 = 0x74; return true;
            case 0x03C5: encoded8 = 0x75; return true;
            case 0x03C6: encoded8 = 0x66; return true;
            case 0x03C7: encoded8 = 0x63; return true;
            case 0x03C8: encoded8 = 0x79; return true;
            case 0x03C9: encoded8 = 0x77; return true;
            case 0x03D1: encoded8 = 0x4A; return true;
            case 0x03D2: encoded8 = 0xA1; return true;
            case 0x03D5: encoded8 = 0x6A; return true;
            case 0x03D6: encoded8 = 0x76; return true;
            case 0x2022: encoded8 = 0xB7; return true;
            case 0x2026: encoded8 = 0xBC; return true;
            case 0x2032: encoded8 = 0xA2; return true;
            case 0x2033: encoded8 = 0xB2; return true;
            case 0x2044: encoded8 = 0xA4; return true;
            case 0x2111: encoded8 = 0xC1; return true;
            case 0x2118: encoded8 = 0xC3; return true;
            case 0x211C: encoded8 = 0xC2; return true;
            case 0x2126: encoded8 = 0x57; return true;
            case 0x2135: encoded8 = 0xC0; return true;
            case 0x2190: encoded8 = 0xAC; return true;
            case 0x2191: encoded8 = 0xAD; return true;
            case 0x2192: encoded8 = 0xAE; return true;
            case 0x2193: encoded8 = 0xAF; return true;
            case 0x2194: encoded8 = 0xAB; return true;
            case 0x21B5: encoded8 = 0xBF; return true;
            case 0x21D0: encoded8 = 0xDC; return true;
            case 0x21D1: encoded8 = 0xDD; return true;
            case 0x21D2: encoded8 = 0xDE; return true;
            case 0x21D3: encoded8 = 0xDF; return true;
            case 0x21D4: encoded8 = 0xDB; return true;
            case 0x2200: encoded8 = 0x22; return true;
            case 0x2202: encoded8 = 0xB6; return true;
            case 0x2203: encoded8 = 0x24; return true;
            case 0x2205: encoded8 = 0xC6; return true;
            case 0x2206: encoded8 = 0x44; return true;
            case 0x2207: encoded8 = 0xD1; return true;
            case 0x2208: encoded8 = 0xCE; return true;
            case 0x2209: encoded8 = 0xCF; return true;
            case 0x220B: encoded8 = 0x27; return true;
            case 0x220F: encoded8 = 0xD5; return true;
            case 0x2211: encoded8 = 0xE5; return true;
            case 0x2212: encoded8 = 0x2D; return true;
            case 0x2217: encoded8 = 0x2A; return true;
            case 0x221A: encoded8 = 0xD6; return true;
            case 0x221D: encoded8 = 0xB5; return true;
            case 0x221E: encoded8 = 0xA5; return true;
            case 0x2220: encoded8 = 0xD0; return true;
            case 0x2227: encoded8 = 0xD9; return true;
            case 0x2228: encoded8 = 0xDA; return true;
            case 0x2229: encoded8 = 0xC7; return true;
            case 0x222A: encoded8 = 0xC8; return true;
            case 0x222B: encoded8 = 0xF2; return true;
            case 0x2234: encoded8 = 0x5C; return true;
            case 0x223C: encoded8 = 0x7E; return true;
            case 0x2245: encoded8 = 0x40; return true;
            case 0x2248: encoded8 = 0xBB; return true;
            case 0x2260: encoded8 = 0xB9; return true;
            case 0x2261: encoded8 = 0xBA; return true;
            case 0x2264: encoded8 = 0xA3; return true;
            case 0x2265: encoded8 = 0xB3; return true;
            case 0x2282: encoded8 = 0xCC; return true;
            case 0x2283: encoded8 = 0xC9; return true;
            case 0x2284: encoded8 = 0xCB; return true;
            case 0x2286: encoded8 = 0xCD; return true;
            case 0x2287: encoded8 = 0xCA; return true;
            case 0x2295: encoded8 = 0xC5; return true;
            case 0x2297: encoded8 = 0xC4; return true;
            case 0x22A5: encoded8 = 0x5E; return true;
            case 0x22C5: encoded8 = 0xD7; return true;
            case 0x2320: encoded8 = 0xF3; return true;
            case 0x2321: encoded8 = 0xF5; return true;
            case 0x2329: encoded8 = 0xE1; return true;
            case 0x232A: encoded8 = 0xF1; return true;
            case 0x25CA: encoded8 = 0xE0; return true;
            case 0x2660: encoded8 = 0xAA; return true;
            case 0x2663: encoded8 = 0xA7; return true;
            case 0x2665: encoded8 = 0xA9; return true;
            case 0x2666: encoded8 = 0xA8; return true;
            default:
                encoded8 = 0;
                return false;
        }
    }

    private static void EmitOperator(string value, List<byte> output)
    {
        if (value == "⁡") return;
        foreach (var scalar in EnumerateBmpScalars(value))
        {
            if (IsWhiteSpaceScalar(scalar))
            {
                EmitScalar(scalar, TypefaceSpace, output);
                continue;
            }

            if (TryResolveLegacyMathTypeGlyph(
                    scalar,
                    out var mtCode,
                    out var typeface,
                    out var encoded8))
            {
                EmitScalar(
                    mtCode,
                    typeface,
                    output,
                    includeEncoded8: true,
                    encoded8Override: encoded8);
                continue;
            }

            if (scalar <= 0xFF)
            {
                EmitScalar(
                    scalar,
                    TypefaceSymbol,
                    output,
                    includeEncoded8: true,
                    encoded8Override: (byte)scalar);
                continue;
            }

            // Do not pretend every Unicode operator lives in Adobe Symbol.
            // Times New Roman/MathType Text is Unicode-capable and is the safe
            // fallback for symbols not present in the legacy Symbol encoding.
            EmitScalar(NormalizeMathTypeMtCode(scalar), TypefaceText, output);
        }
    }

    private static void EmitText(string value, int typeface, List<byte> output)
    {
        foreach (var scalar in EnumerateBmpScalars(value))
        {
            if (IsWhiteSpaceScalar(scalar)) EmitScalar(scalar, TypefaceSpace, output);
            else EmitScalar(scalar, typeface, output);
        }
    }

    private static IEnumerable<int> EnumerateBmpScalars(string value)
    {
        foreach (var character in value)
        {
            if (char.IsSurrogate(character))
                throw new InvalidDataException(
                    "MathType MTEF v5 writer does not yet support non-BMP Unicode characters.");
            yield return character;
        }
    }

    private static bool IsLetterScalar(int scalar) =>
        scalar <= char.MaxValue && char.IsLetter((char)scalar);

    private static bool IsWhiteSpaceScalar(int scalar) =>
        scalar <= char.MaxValue && char.IsWhiteSpace((char)scalar);

    private static void EmitCharacter(char character, int typeface, List<byte> output) =>
        EmitScalar(character, typeface, output);

    private static void EmitScalar(
        int scalar,
        int typeface,
        List<byte> output,
        bool includeEncoded8 = false,
        IReadOnlyList<byte>? embellishments = null,
        byte? encoded8Override = null,
        bool functionStart = false)
    {
        if (scalar < 0 || scalar > 0xFFFF)
            throw new InvalidDataException(
                $"MathType MTEF v5 MTCode writer does not yet support non-BMP scalar U+{scalar:X}.");
        var hasEmbellishments = embellishments is { Count: > 0 };
        var options = (byte)(includeEncoded8 ? CharEncoded8 : 0);
        if (hasEmbellishments) options |= CharHasEmbellishment;
        if (functionStart) options |= CharFunctionStart;
        output.Add(RecordChar);
        output.Add(options);
        WriteSigned(output, typeface);
        output.Add((byte)(scalar & 0xFF));
        output.Add((byte)((scalar >> 8) & 0xFF));
        if (includeEncoded8) output.Add(encoded8Override ?? (byte)scalar);
        if (!hasEmbellishments) return;
        foreach (var embellishment in embellishments!)
            output.AddRange(new byte[] { RecordEmbellishment, 0, embellishment });
        output.Add(RecordEnd);
    }

    private static void WriteSigned(List<byte> output, int value)
    {
        if (value is >= -128 and < 127)
        {
            output.Add(unchecked((byte)(value + 128)));
            return;
        }
        if (value < short.MinValue || value > short.MaxValue - 1)
            throw new ArgumentOutOfRangeException(nameof(value));
        var raw = value + 32768;
        output.Add(0xFF);
        output.Add((byte)(raw & 0xFF));
        output.Add((byte)((raw >> 8) & 0xFF));
    }

    private static bool IsLowerGreek(int scalar) =>
        scalar is >= 0x03B1 and <= 0x03C9
        || scalar is 0x03D1 or 0x03D5 or 0x03D6 or 0x03F1 or 0x03F5;

    private static bool IsUpperGreek(int scalar) => scalar is >= 0x0391 and <= 0x03A9;

    private sealed class MtefStructureReader
    {
        private sealed class FontDefinition
        {
            public int EncodingIndex { get; set; }
            public string Name { get; set; } = string.Empty;
        }

        private sealed class FontStyleDefinition
        {
            public int FontDefinitionIndex { get; set; }
            public byte CharacterStyle { get; set; }
        }

        private readonly byte[] _data;
        private readonly List<string> _encodingDefinitions = new()
        {
            "MTCode",
            "Unknown",
            "Symbol",
            "MTExtra",
        };
        private readonly List<FontDefinition> _fontDefinitions = new();
        private readonly List<FontStyleDefinition> _fontStyleDefinitions = new();
        private int _position;

        internal MtefStructureReader(byte[] data, int position)
        {
            _data = data;
            ScanPrefixDefinitions(position);
            _position = position;
        }

        private void ScanPrefixDefinitions(int rootPosition)
        {
            var cursor = FindEquationOptionsOffset(_data) + 1;
            while (cursor < rootPosition)
            {
                var record = _data[cursor];
                switch (record)
                {
                    case RecordEncodingDef:
                        cursor++;
                        _encodingDefinitions.Add(ReadNullTerminatedString(_data, ref cursor));
                        break;
                    case RecordFontDef:
                    {
                        cursor++;
                        var encoding = ReadUnsigned(_data, ref cursor);
                        _fontDefinitions.Add(new FontDefinition
                        {
                            EncodingIndex = encoding,
                            Name = ReadNullTerminatedString(_data, ref cursor),
                        });
                        break;
                    }
                    case RecordFontStyleDef:
                    {
                        cursor++;
                        var font = ReadUnsigned(_data, ref cursor);
                        Require(_data, cursor, 1);
                        _fontStyleDefinitions.Add(new FontStyleDefinition
                        {
                            FontDefinitionIndex = font,
                            CharacterStyle = _data[cursor++],
                        });
                        break;
                    }
                    case RecordColorDef:
                        cursor = SkipColorDefinition(_data, cursor);
                        break;
                    case RecordEqnPrefs:
                        cursor = SkipEquationPreferences(_data, cursor);
                        break;
                    default:
                        // The initial SIZE record marks the end of the immutable
                        // definition prefix and immediately precedes the root object.
                        cursor = SkipInitialSizeRecord(_data, cursor);
                        return;
                }
            }
        }

        private static string ReadNullTerminatedString(byte[] data, ref int position)
        {
            var start = position;
            while (position < data.Length && data[position] != 0) position++;
            if (position >= data.Length)
                throw new EndOfStreamException("MTEF string is not null terminated.");
            var value = System.Text.Encoding.Default.GetString(data, start, position - start);
            position++;
            return value;
        }

        private void ReadDefinitionRecord()
        {
            var record = Peek();
            switch (record)
            {
                case RecordEncodingDef:
                    _position++;
                    _encodingDefinitions.Add(ReadNullTerminatedString(_data, ref _position));
                    return;
                case RecordFontDef:
                    _position++;
                    _fontDefinitions.Add(new FontDefinition
                    {
                        EncodingIndex = ReadUnsigned(_data, ref _position),
                        Name = ReadNullTerminatedString(_data, ref _position),
                    });
                    return;
                case RecordFontStyleDef:
                    _position++;
                    _fontStyleDefinitions.Add(new FontStyleDefinition
                    {
                        FontDefinitionIndex = ReadUnsigned(_data, ref _position),
                        CharacterStyle = ReadByte(),
                    });
                    return;
                default:
                    throw new InvalidDataException(
                        $"MTEF record {record} is not a font/encoding definition.");
            }
        }

        internal IEnumerable<XNode> ReadRoot()
        {
            SkipFormattingRecords();
            if (Peek() == RecordLine)
                return ReadLine().Nodes().ToArray();
            if (Peek() == RecordPile)
                return new XNode[] { ReadPile() };
            if (Peek() == RecordMatrix)
                return new XNode[] { ReadMatrix() };
            throw new InvalidDataException(
                $"Unsupported MathType root record {Peek()} at offset {_position}.");
        }

        internal bool HasOnlyEquationEndRemaining()
        {
            SkipFormattingRecords();
            return _position == _data.Length - 1
                && Peek() == RecordEnd;
        }

        private XElement ReadLine()
        {
            Expect(RecordLine);
            var options = ReadByte();
            SkipNudge(options);
            if ((options & 0x04) != 0) Skip(2);
            if ((options & 0x02) != 0) SkipRuler();
            var row = new XElement("mrow");
            if ((options & LineNull) != 0) return row;
            ReadObjectList(row);
            return row;
        }

        private XElement ReadPile()
        {
            Expect(RecordPile);
            var options = ReadByte();
            SkipNudge(options);
            Skip(2); // horizontal and vertical alignment
            if ((options & 0x02) != 0) SkipRuler();
            // Preserve the MTEF container kind. Both PILE and MATRIX map well
            // to MathML <mtable>, but a two-row/one-column PILE inside ordinary
            // parentheses is MathType's native binomial representation whereas
            // a MATRIX with the same visible shape is an explicit column matrix.
            // Keeping this private data marker lets semantic round-trip logic and
            // MathML→LaTeX conversion distinguish the two without changing the
            // visible MathML structure.
            var table = new XElement(
                "mtable",
                new XAttribute("data-mtef-pile", "true"));
            while (true)
            {
                SkipFormattingRecords();
                if (Peek() == RecordEnd)
                {
                    _position++;
                    break;
                }
                var line = ReadLine();
                table.Add(new XElement("mtr", new XElement("mtd", line.Nodes())));
            }
            return table;
        }

        private XElement ReadMatrix()
        {
            Expect(RecordMatrix);
            var options = ReadByte();
            SkipNudge(options);
            Skip(3); // vertical alignment, horizontal and vertical cell justification
            var rowCount = ReadByte();
            var columnCount = ReadByte();
            if (rowCount == 0 || columnCount == 0)
                throw new InvalidDataException("MathType MATRIX has zero rows or columns.");
            Skip(PartitionByteCount(rowCount + 1));
            Skip(PartitionByteCount(columnCount + 1));
            var table = new XElement("mtable");
            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var row = new XElement("mtr");
                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    SkipFormattingRecords();
                    if (Peek() != RecordLine)
                        throw new InvalidDataException(
                            $"MathType MATRIX cell {rowIndex},{columnIndex} is not a LINE record.");
                    var cell = ReadLine();
                    row.Add(new XElement("mtd", cell.Nodes()));
                }
                table.Add(row);
            }
            SkipFormattingRecords();
            if (Peek() != RecordEnd)
                throw new InvalidDataException(
                    $"MathType MATRIX has no END record at offset {_position}.");
            _position++;
            return table;
        }

        private void ReadObjectList(XElement row)
        {
            while (true)
            {
                SkipFormattingRecords();
                var record = Peek();
                if (record == RecordEnd)
                {
                    _position++;
                    return;
                }
                if (record == RecordChar)
                {
                    AppendNode(row, ReadCharacter());
                    continue;
                }
                if (record == RecordTemplate)
                {
                    var template = ReadTemplate();
                    if (template.ScriptKind != 0)
                    {
                        var baseNode = row.Nodes().LastOrDefault();
                        if (baseNode is null)
                            throw new InvalidDataException(
                                $"MathType script template at offset {_position} has no preceding base object.");
                        baseNode.Remove();
                        row.Add(BuildScriptNode(baseNode, template));
                    }
                    else if (template.Node is not null)
                    {
                        row.Add(template.Node);
                    }
                    continue;
                }
                if (record == RecordLine)
                {
                    row.Add(ReadLine());
                    continue;
                }
                if (record == RecordPile)
                {
                    row.Add(ReadPile());
                    continue;
                }
                if (record == RecordMatrix)
                {
                    row.Add(ReadMatrix());
                    continue;
                }
                if (record >= 100)
                {
                    SkipFutureRecord();
                    continue;
                }
                throw new InvalidDataException(
                    $"Unsupported MathType MTEF record {record} at offset {_position}.");
            }
        }

        private XElement ReadCharacter()
        {
            Expect(RecordChar);
            var options = ReadByte();
            SkipNudge(options);
            var typeface = ReadSigned();
            int scalar = -1;
            if ((options & 0x20) == 0)
            {
                Require(_data, _position, 2);
                scalar = _data[_position] | (_data[_position + 1] << 8);
                _position += 2;
            }
            int encoded = -1;
            if ((options & CharEncoded8) != 0) encoded = ReadByte();
            if ((options & 0x10) != 0)
            {
                Require(_data, _position, 2);
                encoded = _data[_position] | (_data[_position + 1] << 8);
                _position += 2;
            }
            if (scalar < 0) scalar = encoded;
            if (scalar < 0 || scalar > char.MaxValue)
                scalar = '?';

            var embellishments = (options & CharHasEmbellishment) != 0
                ? ReadEmbellishmentList()
                : Array.Empty<byte>();

            var text = ((char)scalar).ToString();
            string? explicitMathVariant = null;
            if (typeface < 0)
            {
                explicitMathVariant = ResolveExplicitMathVariant(
                    -typeface,
                    scalar,
                    encoded,
                    out var explicitText);
                if (!string.IsNullOrEmpty(explicitText)) text = explicitText;
            }

            XElement token;
            if (typeface < 0)
            {
                token = new XElement("mi", text);
                if (!string.IsNullOrWhiteSpace(explicitMathVariant))
                    token.SetAttributeValue("mathvariant", explicitMathVariant);
            }
            else if (TryNormalizeLetterlikeScalar(
                         scalar,
                         out var normalizedLetterlikeText,
                         out var normalizedLetterlikeVariant))
            {
                token = new XElement(
                    "mi",
                    new XAttribute("mathvariant", normalizedLetterlikeVariant),
                    normalizedLetterlikeText);
            }
            else if (typeface == TypefaceNumber)
                token = new XElement("mn", text);
            else if (typeface == TypefaceSpace)
                token = new XElement("mspace", new XAttribute("width", "0.2em"));
            else if (typeface == TypefaceSymbol)
                token = new XElement("mo", text);
            else if (typeface == TypefaceText)
                token = IsMathematicalSymbolScalar(scalar)
                    ? new XElement("mo", text)
                    : new XElement("mtext", text);
            else if (typeface == TypefaceFunction)
            {
                if (text.All(char.IsLetter))
                {
                    token = new XElement(
                        "mi",
                        new XAttribute("mathvariant", "normal"),
                        new XAttribute("data-mtef-run", "function"),
                        text);
                }
                else if (text.All(char.IsWhiteSpace))
                    token = new XElement("mspace", new XAttribute("width", "0.2em"));
                else
                    token = new XElement("mo", text);
            }
            else if (typeface == TypefaceVector)
                token = new XElement("mi", new XAttribute("mathvariant", "bold"), text);
            else
                token = new XElement("mi", text);

            foreach (var embellishment in embellishments)
            {
                var mark = EmbellishmentMark(embellishment);
                if (mark.Length == 0) continue;
                token = new XElement(
                    "mover",
                    new XAttribute("accent", "true"),
                    token,
                    new XElement("mo", mark));
            }
            return token;
        }

        private string? ResolveExplicitMathVariant(
            int styleIndex,
            int scalar,
            int encoded,
            out string text)
        {
            text = encoded is >= 0x21 and <= 0x7E
                ? ((char)encoded).ToString()
                : ((char)scalar).ToString();
            string? fontName = null;
            byte characterStyle = 0;
            if (styleIndex > 0 && styleIndex <= _fontStyleDefinitions.Count)
            {
                var style = _fontStyleDefinitions[styleIndex - 1];
                characterStyle = style.CharacterStyle;
                if (style.FontDefinitionIndex > 0
                    && style.FontDefinitionIndex <= _fontDefinitions.Count)
                    fontName = _fontDefinitions[style.FontDefinitionIndex - 1].Name;
            }

            var font = fontName ?? string.Empty;
            if (font.IndexOf("Euclid Math One", StringComparison.OrdinalIgnoreCase) >= 0)
                return "script";
            if (font.IndexOf("Euclid Math Two", StringComparison.OrdinalIgnoreCase) >= 0)
                return "double-struck";
            if (font.IndexOf("Fraktur", StringComparison.OrdinalIgnoreCase) >= 0)
                return "fraktur";

            if (TryNormalizeLetterlikeScalar(scalar, out var letterlikeText, out var letterlikeVariant))
            {
                text = letterlikeText;
                return letterlikeVariant;
            }

            return characterStyle switch
            {
                1 => "bold",
                2 => "italic",
                3 => "bold-italic",
                _ => "normal",
            };
        }

        private static bool TryNormalizeLetterlikeScalar(
            int scalar,
            out string text,
            out string mathVariant)
        {
            text = string.Empty;
            mathVariant = string.Empty;
            switch (scalar)
            {
                case 0x2102: text = "C"; mathVariant = "double-struck"; return true;
                case 0x210B: text = "H"; mathVariant = "script"; return true;
                case 0x210C: text = "H"; mathVariant = "fraktur"; return true;
                case 0x210D: text = "H"; mathVariant = "double-struck"; return true;
                case 0x2110: text = "I"; mathVariant = "script"; return true;
                case 0x2111: text = "I"; mathVariant = "fraktur"; return true;
                case 0x2112: text = "L"; mathVariant = "script"; return true;
                case 0x2115: text = "N"; mathVariant = "double-struck"; return true;
                case 0x2119: text = "P"; mathVariant = "double-struck"; return true;
                case 0x211A: text = "Q"; mathVariant = "double-struck"; return true;
                case 0x211B: text = "R"; mathVariant = "script"; return true;
                case 0x211C: text = "R"; mathVariant = "fraktur"; return true;
                case 0x211D: text = "R"; mathVariant = "double-struck"; return true;
                case 0x2124: text = "Z"; mathVariant = "double-struck"; return true;
                case 0x2128: text = "Z"; mathVariant = "fraktur"; return true;
                case 0x212C: text = "B"; mathVariant = "script"; return true;
                case 0x212D: text = "C"; mathVariant = "fraktur"; return true;
                case 0x212F: text = "e"; mathVariant = "script"; return true;
                case 0x2130: text = "E"; mathVariant = "script"; return true;
                case 0x2131: text = "F"; mathVariant = "script"; return true;
                case 0x2133: text = "M"; mathVariant = "script"; return true;
                case 0x2134: text = "o"; mathVariant = "script"; return true;
                default: return false;
            }
        }

        private static void AppendNode(XElement row, XElement node)
        {
            var previous = row.Elements().LastOrDefault();
            var mergeableRun = node.Name.LocalName is "mn" or "mtext"
                || node.Name.LocalName == "mi"
                    && string.Equals(
                        (string?)node.Attribute("data-mtef-run"),
                        "function",
                        StringComparison.Ordinal)
                    && string.Equals(
                        (string?)previous?.Attribute("data-mtef-run"),
                        "function",
                        StringComparison.Ordinal);
            if (previous is not null
                && mergeableRun
                && previous.Name.LocalName == node.Name.LocalName
                && string.Equals(
                    (string?)previous.Attribute("mathvariant") ?? string.Empty,
                    (string?)node.Attribute("mathvariant") ?? string.Empty,
                    StringComparison.OrdinalIgnoreCase)
                && !previous.HasElements
                && !node.HasElements)
            {
                previous.Value += node.Value;
                return;
            }
            row.Add(node);
        }

        private static string EmbellishmentMark(byte embellishment) => embellishment switch
        {
            EmbellDot => ".",
            EmbellDoubleDot => "¨",
            EmbellTilde => "~",
            EmbellHat => "^",
            EmbellRightArrow => "→",
            EmbellLeftArrow => "←",
            EmbellBothArrow => "↔",
            EmbellOverbar => "¯",
            _ => string.Empty,
        };

        private sealed class ParsedTemplate
        {
            internal byte ScriptKind { get; set; }
            internal XElement? Node { get; set; }
            internal XElement? Subscript { get; set; }
            internal XElement? Superscript { get; set; }
        }

        private ParsedTemplate ReadTemplate()
        {
            Expect(RecordTemplate);
            var options = ReadByte();
            SkipNudge(options);
            var selector = ReadByte();
            var variationFirst = ReadByte();
            var variation = (variationFirst & 0x80) == 0
                ? variationFirst
                : (variationFirst & 0x7F) | (ReadByte() << 8);
            _ = ReadByte(); // template-specific options

            switch (selector)
            {
                case TemplateFraction:
                {
                    var numerator = ReadNextLineSlot();
                    var denominator = ReadNextLineSlot();
                    ConsumeTemplateEnd();
                    return new ParsedTemplate
                    {
                        Node = new XElement(
                            "mfrac",
                            CollapseRow(numerator),
                            CollapseRow(denominator)),
                    };
                }
                case TemplateRoot:
                {
                    var radicand = ReadNextLineSlot();
                    var index = ReadNextLineSlot();
                    ConsumeTemplateEnd();
                    return new ParsedTemplate
                    {
                        Node = variation == 1 && index.HasElements
                            ? new XElement(
                                "mroot",
                                CollapseRow(radicand),
                                CollapseRow(index))
                            : new XElement("msqrt", radicand.Nodes()),
                    };
                }
                case TemplateSub:
                case TemplateSup:
                case TemplateSubSup:
                {
                    var subscript = ReadNextLineSlot();
                    var superscript = ReadNextLineSlot();
                    ConsumeTemplateEnd();
                    return new ParsedTemplate
                    {
                        ScriptKind = selector,
                        Subscript = subscript,
                        Superscript = superscript,
                    };
                }
                case TemplateAngle:
                case TemplateParen:
                case TemplateBrace:
                case TemplateBracket:
                case TemplateBar:
                case TemplateDoubleBar:
                case TemplateFloor:
                case TemplateCeiling:
                {
                    var main = ReadNextLineSlot();
                    // MathType may include explicit left/right fence CHAR records.
                    // They are presentation metadata; consume all remaining
                    // subobjects until this template's END and reconstruct from
                    // selector/variation instead.
                    while (true)
                    {
                        SkipFormattingRecords();
                        if (Peek() == RecordEnd)
                        {
                            _position++;
                            break;
                        }
                        SkipOneObject();
                    }
                    var (open, close) = FenceCharacters(selector, variation);
                    return new ParsedTemplate
                    {
                        Node = new XElement(
                            "mfenced",
                            new XAttribute("open", open),
                            new XAttribute("close", close),
                            CollapseRow(main)),
                    };
                }
                case TemplateUnderbar:
                case TemplateOverbar:
                case TemplateArrow:
                case TemplateVector:
                case TemplateTilde:
                case TemplateHat:
                case TemplateArc:
                case TemplateStrike:
                case TemplateBox:
                {
                    var main = ReadNextLineSlot();
                    ConsumeRemainingTemplateObjects();
                    var body = CollapseRow(main);
                    XElement node;
                    if (selector == TemplateUnderbar)
                        node = new XElement("munder", body, new XElement("mo", "_"));
                    else if (selector == TemplateOverbar)
                        node = new XElement("mover", body, new XElement("mo", "¯"));
                    else if (selector is TemplateArrow or TemplateVector)
                    {
                        var arrow = (variation & 0x03) switch
                        {
                            1 => "←",
                            3 => "↔",
                            _ => "→",
                        };
                        node = new XElement("mover", body, new XElement("mo", arrow));
                    }
                    else if (selector == TemplateTilde)
                        node = new XElement("mover", body, new XElement("mo", "~"));
                    else if (selector == TemplateHat)
                        node = new XElement("mover", body, new XElement("mo", "^"));
                    else if (selector == TemplateArc)
                        node = new XElement("mover", body, new XElement("mo", "⌢"));
                    else if (selector == TemplateBox)
                        node = new XElement("menclose", new XAttribute("notation", "box"), body);
                    else
                    {
                        var strikeNotations = new List<string>();
                        if ((variation & 0x01) != 0) strikeNotations.Add("horizontalstrike");
                        if ((variation & 0x02) != 0) strikeNotations.Add("updiagonalstrike");
                        if ((variation & 0x04) != 0) strikeNotations.Add("downdiagonalstrike");
                        if (strikeNotations.Count == 0) strikeNotations.Add("horizontalstrike");
                        node = new XElement(
                            "menclose",
                            new XAttribute("notation", string.Join(" ", strikeNotations)),
                            body);
                    }
                    return new ParsedTemplate { Node = node };
                }
                case TemplateHorizontalBrace:
                case TemplateHorizontalBracket:
                {
                    var slots = ReadTemplateLineSlotsUntilEnd();
                    var main = slots.Count > 0
                        ? CollapseRow(slots[0])
                        : new XElement("mrow");
                    var annotation = slots.Count > 1 && slots[1].HasElements
                        ? CollapseRow(slots[1])
                        : null;
                    var top = (variation & 0x01) != 0;
                    var mark = selector == TemplateHorizontalBrace
                        ? (top ? "⏞" : "⏟")
                        : (top ? "⎴" : "⎵");
                    XNode decorated = top
                        ? new XElement("mover", main, new XElement("mo", mark))
                        : new XElement("munder", main, new XElement("mo", mark));
                    if (annotation is not null)
                        decorated = top
                            ? new XElement("mover", decorated, annotation)
                            : new XElement("munder", decorated, annotation);
                    return new ParsedTemplate { Node = (XElement)decorated };
                }
                case TemplateIntegral:
                case TemplateSum:
                case TemplateProduct:
                case TemplateCoproduct:
                case TemplateUnion:
                case TemplateIntersection:
                case 21: // integral-family variants in MathType's BigOp group
                case 22:
                    return ReadBigOperatorTemplate(selector, variation);
                case TemplateLimit:
                    return ReadLimitTemplate();
                default:
                    throw new InvalidDataException(
                        $"Unsupported MathType template selector {selector} at offset {_position}.");
            }
        }

        private ParsedTemplate ReadBigOperatorTemplate(byte selector, int variation)
        {
            var main = ReadNextLineSlot();
            // MathType 7's persisted BigOp objects are observed in Word as
            // main, lower, upper, operator. Keep the empirically verified order:
            // swapping these two slots makes i=1/n reopen with inverted limits.
            var lower = ReadNextLineSlot();
            var upper = ReadNextLineSlot();
            SkipFormattingRecords();
            XElement? serializedOperator = null;
            if (Peek() == RecordChar)
            {
                serializedOperator = ReadCharacter();
                while (serializedOperator.Name.LocalName == "mover")
                    serializedOperator = serializedOperator.Elements().FirstOrDefault()
                        ?? new XElement("mo", DefaultBigOperatorCharacter(selector));
            }
            ConsumeRemainingTemplateObjects();
            // MathType stores the large glyph in a font/MTCode-specific CHAR. For
            // the dedicated BigOp selectors the selector itself is the durable
            // semantic identity; using the decoded CHAR can leak a private glyph
            // or U+FFFD into VisualTeX. Generic BigOp selectors still need the
            // serialized character because their selector does not name the op.
            var operatorNode = selector == TemplateIntegral
                ? new XElement("mo", IntegralOperatorCharacter(variation))
                : selector is >= TemplateSum and <= TemplateIntersection
                    ? new XElement("mo", DefaultBigOperatorCharacter(selector))
                    : serializedOperator ?? new XElement("mo", DefaultBigOperatorCharacter(selector));
            XNode decorated = DecorateWithLimits(operatorNode, lower, upper);
            var row = new XElement("mrow", decorated);
            foreach (var node in main.Nodes().ToArray())
            {
                node.Remove();
                row.Add(node);
            }
            return new ParsedTemplate { Node = row };
        }

        private ParsedTemplate ReadLimitTemplate()
        {
            var main = ReadNextLineSlot();
            var lower = ReadNextLineSlot();
            var upper = ReadNextLineSlot();
            ConsumeRemainingTemplateObjects();
            XNode baseNode = CollapseRow(main);
            return new ParsedTemplate
            {
                Node = (XElement)DecorateWithLimits(baseNode, lower, upper),
            };
        }

        private static XNode DecorateWithLimits(
            XNode baseNode,
            XElement lower,
            XElement upper)
        {
            var hasLower = lower.HasElements;
            var hasUpper = upper.HasElements;
            if (hasLower && hasUpper)
                return new XElement(
                    "msubsup",
                    baseNode,
                    CollapseRow(lower),
                    CollapseRow(upper));
            if (hasLower)
                return new XElement("msub", baseNode, CollapseRow(lower));
            if (hasUpper)
                return new XElement("msup", baseNode, CollapseRow(upper));
            return baseNode;
        }

        private static string IntegralOperatorCharacter(int variation) => (variation & 0x0F) switch
        {
            0x02 => "∬",
            0x03 => "∭",
            // MathType 7 persists the contour-integral BigOp as kind 5:
            // an MT Extra contour adornment followed by the ordinary integral.
            // Keep kind 4 readable for documents produced by older VisualTeX builds.
            0x04 => "∮",
            0x05 => "∮",
            0x08 => "∲",
            0x0C => "∳",
            _ => "∫",
        };

        private static string DefaultBigOperatorCharacter(byte selector) => selector switch
        {
            TemplateIntegral => "∫",
            TemplateSum => "∑",
            TemplateProduct => "∏",
            TemplateCoproduct => "∐",
            TemplateUnion => "⋃",
            TemplateIntersection => "⋂",
            21 => "∫",
            22 => "∫",
            _ => "∑",
        };

        private List<XElement> ReadTemplateLineSlotsUntilEnd()
        {
            var slots = new List<XElement>();
            while (true)
            {
                SkipFormattingRecords();
                if (Peek() == RecordEnd)
                {
                    _position++;
                    return slots;
                }
                if (Peek() == RecordLine)
                {
                    slots.Add(ReadLine());
                    continue;
                }
                SkipOneObject();
            }
        }

        private void ConsumeRemainingTemplateObjects()
        {
            while (true)
            {
                SkipFormattingRecords();
                if (Peek() == RecordEnd)
                {
                    _position++;
                    return;
                }
                SkipOneObject();
            }
        }

        private static XNode BuildScriptNode(XNode baseNode, ParsedTemplate template)
        {
            var sub = template.Subscript is null
                ? new XElement("mrow")
                : CollapseRow(template.Subscript);
            var sup = template.Superscript is null
                ? new XElement("mrow")
                : CollapseRow(template.Superscript);
            if (template.ScriptKind == TemplateSub)
                return new XElement("msub", baseNode, sub);
            if (template.ScriptKind == TemplateSup)
                return new XElement("msup", baseNode, sup);
            return new XElement("msubsup", baseNode, sub, sup);
        }

        private static XNode CollapseRow(XElement row)
        {
            var nodes = row.Nodes().ToArray();
            if (nodes.Length == 1)
            {
                nodes[0].Remove();
                return nodes[0];
            }
            return new XElement("mrow", nodes);
        }

        private XElement ReadNextLineSlot()
        {
            SkipFormattingRecords();
            if (Peek() != RecordLine)
                throw new InvalidDataException(
                    $"Expected MathType template LINE slot at offset {_position}, actual={Peek()}.");
            return ReadLine();
        }

        private void ConsumeTemplateEnd()
        {
            SkipFormattingRecords();
            if (Peek() != RecordEnd)
                throw new InvalidDataException(
                    $"Expected MathType template END at offset {_position}, actual={Peek()}.");
            _position++;
        }

        private void SkipOneObject()
        {
            var record = Peek();
            switch (record)
            {
                case RecordChar:
                    _ = ReadCharacter();
                    return;
                case RecordLine:
                    _ = ReadLine();
                    return;
                case RecordTemplate:
                    _ = ReadTemplate();
                    return;
                case RecordPile:
                    _ = ReadPile();
                    return;
                case RecordMatrix:
                    _ = ReadMatrix();
                    return;
                default:
                    SkipFormattingRecords();
                    if (Peek() == record)
                        throw new InvalidDataException(
                            $"Cannot skip MathType MTEF record {record} at offset {_position}.");
                    return;
            }
        }

        private void SkipFormattingRecords()
        {
            while (_position < _data.Length)
            {
                var record = _data[_position];
                if (record >= RecordFull && record <= RecordSubSym)
                {
                    _position++;
                    continue;
                }
                if (record == RecordSize)
                {
                    _position = SkipInitialSizeRecord(_data, _position);
                    continue;
                }
                if (record == 15)
                {
                    _position++;
                    _ = ReadUnsigned(_data, ref _position);
                    continue;
                }
                if (record == RecordColorDef)
                {
                    _position = SkipColorDefinition(_data, _position);
                    continue;
                }
                if (record is RecordEncodingDef or RecordFontDef or RecordFontStyleDef)
                {
                    ReadDefinitionRecord();
                    continue;
                }
                if (record >= 100)
                {
                    SkipFutureRecord();
                    continue;
                }
                break;
            }
        }

        private void SkipFutureRecord()
        {
            _position++; // record type
            var length = ReadUnsigned(_data, ref _position);
            Skip(length);
        }

        private byte[] ReadEmbellishmentList()
        {
            var embellishments = new List<byte>();
            while (true)
            {
                var record = Peek();
                if (record == RecordEnd)
                {
                    _position++;
                    return embellishments.ToArray();
                }
                if (record != RecordEmbellishment)
                    throw new InvalidDataException(
                        $"Unexpected MTEF embellishment record {record} at offset {_position}.");
                _position++;
                var options = ReadByte();
                SkipNudge(options);
                embellishments.Add(ReadByte());
            }
        }

        private void SkipRuler()
        {
            Expect(7);
            var count = ReadByte();
            Skip(count * 3);
        }

        private void SkipNudge(byte options)
        {
            if ((options & 0x08) == 0) return;
            var dx = ReadByte();
            var dy = ReadByte();
            if (dx == 128 && dy == 128) Skip(4);
        }

        private int ReadSigned()
        {
            var first = ReadByte();
            if (first != 0xFF) return first - 128;
            Require(_data, _position, 2);
            var raw = _data[_position] | (_data[_position + 1] << 8);
            _position += 2;
            return raw - 32768;
        }

        private byte ReadByte()
        {
            Require(_data, _position, 1);
            return _data[_position++];
        }

        private byte Peek()
        {
            Require(_data, _position, 1);
            return _data[_position];
        }

        private void Expect(byte expected)
        {
            var actual = ReadByte();
            if (actual != expected)
                throw new InvalidDataException(
                    $"Expected MTEF record {expected}, actual={actual} at offset {_position - 1}.");
        }

        private void Skip(int count)
        {
            Require(_data, _position, count);
            _position += count;
        }

        private static (string Open, string Close) FenceCharacters(byte selector, int variation)
        {
            var presentLeft = (variation & 0x01) != 0;
            var presentRight = (variation & 0x02) != 0;
            var pair = selector switch
            {
                TemplateAngle => ("⟨", "⟩"),
                TemplateParen => ("(", ")"),
                TemplateBrace => ("{", "}"),
                TemplateBracket => ("[", "]"),
                TemplateBar => ("|", "|"),
                TemplateDoubleBar => ("‖", "‖"),
                TemplateFloor => ("⌊", "⌋"),
                TemplateCeiling => ("⌈", "⌉"),
                _ => (string.Empty, string.Empty),
            };
            return (presentLeft ? pair.Item1 : string.Empty, presentRight ? pair.Item2 : string.Empty);
        }
    }

    private static void Require(byte[] data, int offset, int count)
    {
        if (offset < 0 || count < 0 || offset + count > data.Length)
            throw new EndOfStreamException(
                $"MTEF stream is truncated at offset {offset}, need {count} bytes.");
    }
}
