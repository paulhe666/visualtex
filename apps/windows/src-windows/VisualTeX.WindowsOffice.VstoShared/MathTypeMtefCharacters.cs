using System.Globalization;
using System.Text;
using System.Xml.Linq;
using static VisualTeX.WordVsto.MathTypeMtefCodec;
using static VisualTeX.WordVsto.MathTypeMtefGlyphs;
using static VisualTeX.WordVsto.MathTypeMtefSignature;

namespace VisualTeX.WordVsto;

/// <summary>
/// Emitting MathML tokens as MTEF characters: identifiers, operators and text, their typefaces and encodings, and
/// math alphanumerics.
/// </summary>
internal static class MathTypeMtefCharacters
{
    internal static void EmitIdentifier(
        XElement element,
        List<byte> output,
        IReadOnlyList<byte>? embellishments = null)
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
        for (var scalarIndex = 0; scalarIndex < scalars.Length; scalarIndex++)
        {
            var originalScalar = scalars[scalarIndex];
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

            if (explicitTypeface is not null
                && scalars.Length == 1
                && TryEmitPreservedCharacter(element, explicitTypeface.Value, output, embellishments))
                continue;

            if (explicitTypeface is not null && scalars.Length == 1
                && ResolveNativeGlyphAuthority(element) is { } nativeGlyph)
            {
                EmitScalar(nativeGlyph.MtCode, explicitTypeface.Value, output,
                    includeEncoded8: true, embellishments: embellishments,
                    encoded8Override: nativeGlyph.Encoded8);
                continue;
            }

            if (scalar == 0x210F && explicitTypeface is null)
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
                    embellishments: embellishments,
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
                    embellishments: embellishments,
                    encoded8Override: mtExtraPosition);
                continue;
            }

            if (explicitTypeface is null
                && !effectiveVariant.Contains("bold")
                && !effectiveVariant.Contains("double-struck")
                && !effectiveVariant.Contains("script")
                && !effectiveVariant.Contains("fraktur")
                // An ASCII punctuation glyph in a literal/upright identifier
                // is not a Symbol operator just because that font contains it.
                // Keep token semantics separate from the font encoding table.
                && (scalar > 0x7F || IsMathematicalSymbolScalar(scalar))
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
                    embellishments: embellishments,
                    encoded8Override: nativeEncoded8);
                continue;
            }

            if (explicitTypeface is null
                && scalars.Length == 1
                && embellishments is null or { Count: 0 }
                && IsMathematicalSymbolToken(((char)scalar).ToString())
                && (IsMathematicalSymbolScalar(scalar)
                    || ResolveTokenTypeface(element, scalar, effectiveVariant, isFunction) == TypefaceText))
            {
                // A symbol identifier (\diamondsuit, \heartsuit, \$, currency
                // signs) in the Text face reads back as text. The operator path
                // has MathType's own form for these glyphs (MT Extra, SimSun,
                // Function), which reads back as the same operator.
                EmitOperator(((char)scalar).ToString(), output);
                continue;
            }

            if (explicitTypeface is null
                && scalar < 0x80 && !char.IsLetterOrDigit((char)scalar)
                && IsMathematicalSymbolScalar(scalar))
            {
                // MathType keeps ASCII punctuation in its Function face, which
                // also reads back as an operator (the Text face would be mi).
                EmitScalar(scalar, TypefaceFunction, output, embellishments: embellishments);
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
                    output,
                    embellishments: embellishments);
                continue;
            }

            var typeface = explicitTypeface
                ?? ResolveTokenTypeface(element, scalar, effectiveVariant, isFunction);
            var dotlessJMtCode = 0;
            byte dotlessJEncoded8 = 0;
            // A font position is meaningful only in its declared font. PUA and
            // encoded8 aliases from Euclid Math One must not be copied into an
            // explicitly requested Times/Arial/Courier character record.
            var usesEuclidMathOne = string.Equals(
                (string?)element.Attribute("data-mtef-explicit-font"),
                "Euclid Math One", StringComparison.OrdinalIgnoreCase);
            var hasDotlessJ = explicitTypeface is < 0 && usesEuclidMathOne
                && TryEuclidMathOneDotlessJ(
                    scalar,
                    out dotlessJMtCode,
                    out dotlessJEncoded8);
            var mtCode = hasDotlessJ
                ? dotlessJMtCode
                : explicitTypeface is < 0
                    ? ExplicitVariantMtCode(effectiveVariant, scalar)
                    : effectiveVariant.Contains("double-struck")
                        && TryStandardDoubleStruckMtCode(scalar, out var standardDoubleStruck)
                            ? standardDoubleStruck
                            : scalar;
            var explicitEncoded8 = hasDotlessJ
                ? (byte?)dotlessJEncoded8
                : explicitTypeface is < 0 && usesEuclidMathOne
                    && TryEuclidMathOneGreekVariantEncoded8(scalar, out var greekVariantEncoded8)
                        ? (byte?)greekVariantEncoded8
                        : explicitTypeface is < 0 && scalar <= 0xFF
                            ? (byte?)scalar
                            : null;
            EmitScalar(
                mtCode,
                typeface,
                output,
                includeEncoded8: explicitEncoded8 is not null,
                embellishments: embellishments,
                encoded8Override: explicitEncoded8,
                functionStart:
                    isFunction
                    && typeface == TypefaceFunction
                    && scalarIndex == 0);
        }
    }

    private static bool CanUseStandardDoubleStruckMtCode(int scalar) =>
        TryStandardDoubleStruckMtCode(scalar, out _);

    internal static bool TryNormalizeLetterlikeForWrite(
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
            case 0x2112: normalizedScalar = 'L'; mathVariant = "script"; return true;
            case 0x2115: normalizedScalar = 'N'; mathVariant = "double-struck"; return true;
            case 0x2119: normalizedScalar = 'P'; mathVariant = "double-struck"; return true;
            case 0x211A: normalizedScalar = 'Q'; mathVariant = "double-struck"; return true;
            case 0x211B: normalizedScalar = 'R'; mathVariant = "script"; return true;
            case 0x211D: normalizedScalar = 'R'; mathVariant = "double-struck"; return true;
            case 0x2124: normalizedScalar = 'Z'; mathVariant = "double-struck"; return true;
            case 0x2128: normalizedScalar = 'Z'; mathVariant = "fraktur"; return true;
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

    internal static bool TryEuclidMathOneGreekVariantEncoded8(
        int scalar,
        out byte encoded8)
    {
        switch (scalar)
        {
            case 0x03F1: encoded8 = 0xF1; return true; // \varrho
            case 0x03F5: encoded8 = 0xF2; return true; // \epsilon
            case 0x03F0: encoded8 = 0xF9; return true; // \varkappa
            default:
                encoded8 = 0;
                return false;
        }
    }

    internal static bool TryEuclidMathOneDotlessJ(
        int scalar,
        out int mtCode,
        out byte encoded8)
    {
        if (scalar == 0x0237) // \\jmath
        {
            mtCode = 0xED02;
            encoded8 = 0xF8;
            return true;
        }
        mtCode = 0;
        encoded8 = 0;
        return false;
    }

    internal static bool TryEuclidMathOneOperatorEncoded8(
        int scalar,
        out byte encoded8)
    {
        switch (scalar)
        {
            case 0x2216: encoded8 = 0x82; return true; // \\setminus
            case 0x224D: encoded8 = 0xA9; return true; // \\asymp
            case 0x22A2: encoded8 = 0x90; return true; // \\vdash
            case 0x22A3: encoded8 = 0x94; return true; // \\dashv
            case 0x22A8: encoded8 = 0x91; return true; // \\models
            case 0x21BC: encoded8 = 0xB4; return true; // \\leftharpoonup
            case 0x21C1: encoded8 = 0xB7; return true; // \\rightharpoondown
            case 0x2296: encoded8 = 0x21; return true; // \\ominus
            case 0x2298: encoded8 = 0x25; return true; // \\oslash
            case 0x22A4: encoded8 = 0x95; return true; // \\top
            case 0x22C6: encoded8 = 0xE5; return true; // \\star
            case 0x2240: encoded8 = 0xAA; return true; // \\wr
            default:
                encoded8 = 0;
                return false;
        }
    }

    internal static bool TryEuclidMathTwoOperatorEncoded8(
        int scalar,
        out byte encoded8)
    {
        switch (scalar)
        {
            // MathType's own TeX translator and Euclid Math Two font agree on
            // the private MTCode pair E938/E939 at legacy positions B0/B1 for
            // \preceq/\succeq. MathJax exposes the public Unicode pair
            // U+2AAF/U+2AB0; accept both at the codec boundary.
            case 0x2AAF:
            case 0xE938: encoded8 = 0xB0; return true; // \preceq
            case 0x2AB0:
            case 0xE939: encoded8 = 0xB1; return true; // \succeq
            case 0x227C: encoded8 = 0xB0; return true; // \preccurlyeq
            case 0x227D: encoded8 = 0xB1; return true; // \succcurlyeq
            case 0x228F: encoded8 = 0xF0; return true; // \\sqsubset
            case 0x2291: encoded8 = 0xF4; return true; // \\sqsubseteq
            case 0x2290: encoded8 = 0xF1; return true; // \\sqsupset
            case 0x2292: encoded8 = 0xF5; return true; // \\sqsupseteq
            case 0x228E: encoded8 = 0xE2; return true; // \\uplus
            case 0x2293: encoded8 = 0xF3; return true; // \\sqcap
            case 0x2294: encoded8 = 0xF2; return true; // \\sqcup
        }
        encoded8 = 0;
        return false;
    }

    internal static bool TryMtExtraDoubleStruck(
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
        if (tokenKind == "mtext")
            return IsFarEastTextScalar(scalar) ? TypefaceFarEastText : TypefaceText;
        if (tokenKind == "mn") return TypefaceNumber;
        if (tokenKind == "mo")
            return functionHint == true ? TypefaceFunction : TypefaceSymbol;
        if (IsMathematicalSymbolScalar(scalar)) return TypefaceSymbol;
        if (IsFarEastTextScalar(scalar)) return TypefaceFarEastText;

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

    internal static bool IsMathematicalSymbolScalar(int scalar)
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

    internal static bool IsFarEastTextScalar(int scalar) =>
        scalar is >= 0x1100 and <= 0x11FF   // Hangul Jamo
            or >= 0x2E80 and <= 0x2FFF     // CJK radicals and description characters
            or >= 0x3000 and <= 0x303F     // CJK symbols and punctuation
            or >= 0x3040 and <= 0x30FF     // Hiragana and Katakana
            or >= 0x3100 and <= 0x31BF     // Bopomofo and related extensions
            or >= 0x31C0 and <= 0x31EF     // CJK strokes
            or >= 0x31F0 and <= 0x31FF     // Katakana phonetic extensions
            or >= 0x3200 and <= 0x33FF     // Enclosed/compatibility CJK
            or >= 0x3400 and <= 0x4DBF     // CJK Unified Ideographs Extension A
            or >= 0x4E00 and <= 0x9FFF     // CJK Unified Ideographs
            or >= 0xA000 and <= 0xA4CF     // Yi syllables and radicals
            or >= 0xA960 and <= 0xA97F     // Hangul Jamo Extended-A
            or >= 0xAC00 and <= 0xD7AF     // Hangul syllables/Jamo Extended-B
            or >= 0xF900 and <= 0xFAFF     // CJK Compatibility Ideographs
            or >= 0xFE10 and <= 0xFE1F     // Vertical forms
            or >= 0xFE30 and <= 0xFE6F     // CJK compatibility/small forms
            or >= 0xFF00 and <= 0xFFEF;    // Halfwidth and fullwidth forms

    // The standalone prefix binds Text/Function/Variable/Number/Vector to
    // Times New Roman with encoding WinAllBasicCodePages (Windows 1250-1258).
    // MathType renders an MTCode outside those code pages as U+FFFD there,
    // e.g. the geometry triangle U+25B3. SimSun with WinAllCodePages (also
    // 874/932/936/949/950) renders it; verified in MathType 7 with U+25B3,
    // U+2312, U+224C, U+2235, U+2460 and U+2103. Text goes to typeface 12
    // (the prefix's SimSun slot, read back as mtext); operators get an
    // explicit SimSun style so they read back as mo.
    // SimSun has the large geometric shapes but not the small/medium forms
    // MathJax uses for \blacktriangle (U+25B4), \square (U+25FB) and
    // \blacksquare (U+25FC); MathType draws those as its missing-glyph
    // diamond. Write the covered shape instead; it reads back as the same command.
    internal static readonly IReadOnlyDictionary<char, char> FarEastGlyphSubstitutes =
        new Dictionary<char, char>
        {
            ['\u25B4'] = '\u25B2', ['\u25B5'] = '\u25B3', ['\u25BE'] = '\u25BC', ['\u25BF'] = '\u25BD',
            ['\u25FB'] = '\u25A1', ['\u25FD'] = '\u25A1', ['\u25FC'] = '\u25A0', ['\u25FE'] = '\u25A0',
        };

    internal static bool NeedsFarEastCodePage(int scalar) =>
        scalar >= 0x80 && scalar <= char.MaxValue && !IsWinAllBasicCodePageScalar(scalar);

    internal static bool IsFarEastCodePageOperator(int scalar)
    {
        if (IsFarEastTextScalar(scalar) || !NeedsFarEastCodePage(scalar)
            || char.GetUnicodeCategory((char)scalar) is UnicodeCategory.Format or UnicodeCategory.Control)
            return false;
        var character = new List<byte>();
        EmitOperator(((char)scalar).ToString(), character);
        return character.Count == 5 && character[0] == RecordChar
            && (character[2] == TypefaceText + 128 || character[2] == TypefaceFunction + 128);
    }

    private static bool IsWinAllBasicCodePageScalar(int scalar)
    {
        var low = 0;
        var high = WinAllBasicCodePageRanges.Length - 1;
        while (low <= high)
        {
            var middle = (low + high) / 2;
            var range = WinAllBasicCodePageRanges[middle];
            if (scalar < range.First) high = middle - 1;
            else if (scalar > range.Last) low = middle + 1;
            else return true;
        }
        return false;
    }

    // Non-ASCII BMP scalars that round-trip through at least one of the
    // Windows code pages 1250-1258 (generated from the Windows code pages).
    private static readonly (int First, int Last)[] WinAllBasicCodePageRanges =
    {
        (0x0081, 0x0081), (0x0083, 0x0083), (0x0088, 0x0088), (0x008A, 0x008A), (0x008C, 0x0090), (0x0098, 0x0098),
        (0x009A, 0x009A), (0x009C, 0x0107), (0x010C, 0x0113), (0x0116, 0x011B), (0x011E, 0x011F), (0x0122, 0x0123),
        (0x012A, 0x012B), (0x012E, 0x0131), (0x0136, 0x0137), (0x0139, 0x013E), (0x0141, 0x0148), (0x014C, 0x014D),
        (0x0150, 0x015B), (0x015E, 0x0165), (0x016A, 0x016B), (0x016E, 0x0173), (0x0178, 0x017E), (0x0192, 0x0192),
        (0x01A0, 0x01A1), (0x01AF, 0x01B0), (0x02C6, 0x02C7), (0x02D8, 0x02D9), (0x02DB, 0x02DD), (0x0300, 0x0301),
        (0x0303, 0x0303), (0x0309, 0x0309), (0x0323, 0x0323), (0x0384, 0x0386), (0x0388, 0x038A), (0x038C, 0x038C),
        (0x038E, 0x03A1), (0x03A3, 0x03CE), (0x0401, 0x040C), (0x040E, 0x044F), (0x0451, 0x045C), (0x045E, 0x045F),
        (0x0490, 0x0491), (0x05B0, 0x05C3), (0x05D0, 0x05EA), (0x05F0, 0x05F4), (0x060C, 0x060C), (0x061B, 0x061B),
        (0x061F, 0x061F), (0x0621, 0x063A), (0x0640, 0x0652), (0x0679, 0x0679), (0x067E, 0x067E), (0x0686, 0x0686),
        (0x0688, 0x0688), (0x0691, 0x0691), (0x0698, 0x0698), (0x06A9, 0x06A9), (0x06AF, 0x06AF), (0x06BA, 0x06BA),
        (0x06BE, 0x06BE), (0x06C1, 0x06C1), (0x06D2, 0x06D2), (0x200C, 0x200F), (0x2013, 0x2015), (0x2018, 0x201A),
        (0x201C, 0x201E), (0x2020, 0x2022), (0x2026, 0x2026), (0x2030, 0x2030), (0x2039, 0x203A), (0x20AA, 0x20AC),
        (0x2116, 0x2116), (0x2122, 0x2122), (0xF88D, 0xF896), (0xF8F9, 0xF8FD),
    };

    internal static string NormalizeOperatorToken(string value) => value switch
    {
        "〈" or "〈" or "⟨" => "⟨",
        "〉" or "〉" or "⟩" => "⟩",
        // MathJax uses U+25C3/U+25B9 for TeX triangleleft/triangleright,
        // while MathType persists the same glyphs as U+22B2/U+22B3.
        "◃" or "⊲" => "⊲",
        "▹" or "⊳" => "⊳",
        // MathJax emits the mathematical minus U+2212 while Word OMML commonly
        // serializes the same binary subtraction operator as ASCII hyphen-minus.
        // They are semantically identical in an <mo> token and must compare equal
        // across VisualTeX/OMML/MathType round trips.
        "−" => "-",
        // MathType 7 may expose TeX \\mid as U+2223 DIVIDES in either <mi>
        // or <mo>, while Word commonly materializes the same relation as the
        // ASCII vertical line. At the semantic-token layer these are the same
        // vertical relation. Absolute-value/norm bars remain structurally
        // distinct because fence recognition canonicalizes them as fence(...)
        // before ordinary operator-token comparison.
        "∣" => "|",
        // MathType 7 stores TeX \\sim as ASCII '~' in its Function typeface,
        // while MathJax represents the same operator as U+223C. This is an
        // operator-token equivalence only; accent \\tilde is represented by
        // mover/template structure and never reaches this normalization.
        "~" => "∼",
        _ => value,
    };

    private static int NormalizeMathTypeMtCode(int scalar) => scalar switch
    {
        // MathType 7/Adobe Symbol use the historical angle-bracket MTCode pair.
        0x27E8 => 0x2329,
        0x27E9 => 0x232A,
        _ => scalar,
    };

    internal static bool TryResolveLegacyMathTypeGlyph(
        int scalar,
        out int mtCode,
        out int typeface,
        out byte encoded8)
    {
        mtCode = NormalizeMathTypeMtCode(scalar);
        typeface = TypefaceSymbol;
        encoded8 = 0;

        // MathType 7.9's importer keeps ASCII punctuation, brackets included,
        // in its Function face (Times New Roman); only + - = < > take their
        // Symbol positions, and the hyphen is stored as MINUS SIGN.
        if (mtCode == '-') mtCode = 0x2212;
        if (mtCode < 0x80 && mtCode is not ('+' or '=' or '<' or '>')) return false;
        if (mtCode == 0x2213)
        {
            // Genuine MathType 7 persists \mp / MINUS-OR-PLUS in MT Extra,
            // not Adobe Symbol: fnMTEXTRA + MTCode U+2213 + encoded8 0x6D.
            // Falling back to Unicode text changes the glyph metrics/baseline.
            typeface = TypefaceMtExtra;
            encoded8 = 0x6D;
            return true;
        }
        if (mtCode == 0x21A6)
        {
            // Genuine MathType 7 persists \mapsto in MT Extra:
            // fnMTEXTRA + MTCode U+21A6 + encoded8 0x61.
            typeface = TypefaceMtExtra;
            encoded8 = 0x61;
            return true;
        }
        if (mtCode is 0x2016 or 0x2225)
        {
            // Genuine MathType 7 persists \parallel in MT Extra:
            // fnMTEXTRA + MTCode U+2225 + encoded8 0x50.
            // Fixed norm bars use the same double-line glyph, retaining U+2016
            // as their MTCode. Unicode Text fallback renders them as boxes.
            typeface = TypefaceMtExtra;
            encoded8 = 0x50;
            return true;
        }

        if (mtCode == 0x2113)
        {
            // Genuine MathType 7 persists \\ell as MT Extra U+2113 / 0x6C.
            typeface = TypefaceMtExtra;
            encoded8 = 0x6C;
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
            // These ASCII punctuation glyphs retain their encoded positions in
            // Adobe Symbol. Letters, quotes, dollar, caret, backslash and the
            // Latin-1 range do NOT generally have identity mappings.
            case '!' or '#' or '%' or '&' or '(' or ')' or '*' or '+' or ','
                or '-' or '.' or '/' or ':' or ';' or '<' or '=' or '>' or '?'
                or '[' or ']' or '_' or '{' or '|' or '}':
                encoded8 = (byte)scalar; return true;
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
            case 0x03C6: encoded8 = 0x6A; return true;
            case 0x03C7: encoded8 = 0x63; return true;
            case 0x03C8: encoded8 = 0x79; return true;
            case 0x03C9: encoded8 = 0x77; return true;
            case 0x03D1: encoded8 = 0x4A; return true;
            case 0x03D2: encoded8 = 0xA1; return true;
            case 0x03D5: encoded8 = 0x66; return true;
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

    internal static void EmitOperator(XElement element, List<byte> output)
    {
        var explicitTypeface = int.TryParse(
            (string?)element.Attribute("data-mtef-explicit-typeface"),
            out var parsedExplicitTypeface)
            ? parsedExplicitTypeface
            : (int?)null;
        if (explicitTypeface is null)
        {
            EmitOperator(element.Value, output);
            return;
        }

        if (element.Value.Length == 1
            && TryEmitPreservedCharacter(element, explicitTypeface.Value, output))
            return;
        if (ResolveNativeGlyphAuthority(element) is { } nativeGlyph)
        {
            EmitScalar(nativeGlyph.MtCode, explicitTypeface.Value, output,
                includeEncoded8: true, encoded8Override: nativeGlyph.Encoded8);
            return;
        }
        foreach (var scalar in EnumerateBmpScalars(element.Value))
        {
            if (TryEuclidMathOneOperatorEncoded8(scalar, out var encoded8)
                || TryEuclidMathTwoOperatorEncoded8(scalar, out encoded8))
            {
                var nativeScalar = scalar switch
                {
                    0x2AAF => 0xE938, // MathJax \preceq -> MathType private MTCode
                    0x2AB0 => 0xE939, // MathJax \succeq -> MathType private MTCode
                    _ => scalar,
                };
                EmitScalar(
                    nativeScalar,
                    explicitTypeface.Value,
                    output,
                    includeEncoded8: true,
                    encoded8Override: encoded8);
                continue;
            }

            EmitScalar(
                scalar,
                explicitTypeface.Value,
                output,
                includeEncoded8: scalar <= 0xFF,
                encoded8Override: scalar <= 0xFF ? (byte?)scalar : null);
        }
    }

    private static bool TryEmitMathTypeSpecialOperator(int scalar, List<byte> output)
    {
        switch (scalar)
        {
            case 0x2020: // \\dagger: genuine MathType uses Function typeface.
                EmitScalar(0x2020, TypefaceFunction, output, includeEncoded8: false);
                return true;
            case 0x2021: // \\ddagger: genuine MathType uses Function typeface.
                EmitScalar(0x2021, TypefaceFunction, output, includeEncoded8: false);
                return true;
            case 0x2218: // Circle MTCode and the legacy Symbol glyph slot are distinct.
                EmitScalar(0x2218, TypefaceSymbol, output, includeEncoded8: true, encoded8Override: 0xB0);
                return true;
            case 0x2022: // \\bullet: genuine MathType uses Function, no encoded8.
                EmitScalar(0x2022, TypefaceFunction, output, includeEncoded8: false);
                return true;
            case 0x22C4: // \\diamond: Symbol position 0xE0.
                EmitScalar(0x22C4, TypefaceSymbol, output, includeEncoded8: true, encoded8Override: 0xE0);
                return true;
            case 0x22B2: // MathType-side \\triangleleft scalar.
            case 0x25C3: // MathJax-side \\triangleleft scalar.
                EmitScalar(0x22B2, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x3C);
                return true;
            case 0x22B3: // MathType-side \\triangleright scalar.
            case 0x25B9: // MathJax-side \\triangleright scalar.
                EmitScalar(0x22B3, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x3E);
                return true;
            case 0x2661: // \\heartsuit: built-in MathType typeface 12.
                EmitScalar(0x2661, TypefaceFarEastText, output, includeEncoded8: false);
                return true;
            case 0x2662: // \\diamondsuit: MT Extra glyph position 0x6E.
                EmitScalar(0xFFFD, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x6E);
                return true;
            case 0x2299: // \\odot: MT Extra position 0x65.
                EmitScalar(0x2299, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x65);
                return true;
            case 0x25EF: // \\bigcirc: built-in typeface 12.
                EmitScalar(0x25EF, TypefaceFarEastText, output, includeEncoded8: false);
                return true;
            case 0x226A: // \\ll: MT Extra position 0x3D.
                EmitScalar(0x226A, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x3D);
                return true;
            case 0x226B: // \\gg: MT Extra position 0x3F.
                EmitScalar(0x226B, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x3F);
                return true;
            case 0x2243: // \\simeq
                EmitScalar(0x2243, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x3B);
                return true;
            case 0x2250: // \\doteq
                EmitScalar(0x2250, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x42);
                return true;
            case 0x227A: // \\prec
                EmitScalar(0x227A, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x70);
                return true;
            case 0x227B: // \\succ
                EmitScalar(0x227B, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x66);
                return true;
            case 0x2195: // \\updownarrow
                EmitScalar(0x2195, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x62);
                return true;
            case 0x21D5: // \\Updownarrow
                EmitScalar(0x21D5, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x63);
                return true;
            case 0x21BD: // \\leftharpoondown
                EmitScalar(0x21BD, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x87);
                return true;
            case 0x21C0: // \\rightharpoonup
                EmitScalar(0x21C0, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x86);
                return true;
            case 0x2197: // \\nearrow
                EmitScalar(0x2197, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x5A);
                return true;
            case 0x2198: // \\searrow
                EmitScalar(0x2198, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x5D);
                return true;
            case 0x2199: // \\swarrow
                EmitScalar(0x2199, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x5B);
                return true;
            case 0x2196: // \\nwarrow
                EmitScalar(0x2196, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x5E);
                return true;
            case 0x2A3F: // \\amalg: MathType stores ordinary amalg as MT Extra U+2210 / 0x43.
                EmitScalar(0x2210, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x43);
                return true;
            case 0x220B: // \\ni: built-in typeface 12, no encoded8.
                EmitScalar(0x220B, TypefaceFarEastText, output, includeEncoded8: false);
                return true;
            case 0x2223: // \\mid: MathType's importer uses built-in typeface 12 too.
                EmitScalar(0x2223, TypefaceFarEastText, output, includeEncoded8: false);
                return true;
            case 0x22EF: // \\cdots
                EmitScalar(0x22EF, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x4C);
                return true;
            case 0x22EE: // \\vdots
                EmitScalar(0x22EE, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x4D);
                return true;
            case 0x22F1: // \\ddots
                EmitScalar(0x22F1, TypefaceMtExtra, output, includeEncoded8: true, encoded8Override: 0x4F);
                return true;
            default:
                return false;
        }
    }

    internal static void EmitOperator(string value, List<byte> output)
    {
        if (value == "⁡") return;
        foreach (var scalar in EnumerateBmpScalars(value))
        {
            // Genuine legacy MathType stores a literal prime as ASCII apostrophe
            // in Function (Times), not byte 0x27 in Symbol (the contains glyph).
            // Keep that representation; U+2032 still follows its Symbol mapping.
            if (scalar == '\'')
            {
                EmitScalar(scalar, TypefaceFunction, output, includeEncoded8: false);
                continue;
            }
            if (scalar is 0x223C or '~')
            {
                // Genuine MathType 7 represents TeX \sim as ASCII '~' in the
                // Function typeface, not U+223C in Adobe Symbol.
                EmitScalar('~', TypefaceFunction, output, includeEncoded8: false);
                continue;
            }
            if (TryEmitMathTypeSpecialOperator(scalar, output))
                continue;
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

            // Code-point identity is not an Adobe Symbol font-position map.
            // In particular Latin-1 diaeresis at 0xA8 would display a diamond.
            // Only the verified encoding table above may choose Symbol.
            // Do not pretend every Unicode operator lives in Adobe Symbol.
            // Times New Roman/MathType Text is Unicode-capable and is the safe
            // fallback for symbols not present in the legacy Symbol encoding.
            // Non-letters use its Function face so they read back as operators.
            EmitScalar(
                NormalizeMathTypeMtCode(scalar),
                IsFarEastTextScalar(scalar) ? TypefaceFarEastText
                    : IsMathematicalSymbolScalar(scalar) || !IsLetterScalar(scalar)
                        && char.GetUnicodeCategory((char)scalar) != UnicodeCategory.Format ? TypefaceFunction : TypefaceText,
                output);
        }
    }

    internal static bool IsVerifiedMtExtraOperator(int scalar, int encoded)
    {
        // MT Extra also contains alphabetic identifiers (ell, hbar, blackboard
        // letters). Only invert an already verified operator CHAR mapping;
        // the font name alone must never turn all of its glyphs into operators.
        if (!IsMathematicalSymbolScalar(scalar)) return false;
        var character = new List<byte>();
        EmitOperator(((char)scalar).ToString(), character);
        return character.Count == 6 && character[0] == RecordChar
            && character[1] == CharEncoded8 && character[2] == TypefaceMtExtra + 128
            && (character[3] | character[4] << 8) == scalar
            && (encoded < 0 || encoded == character[5]);
    }

    internal static void EmitText(string value, int typeface, List<byte> output)
    {
        foreach (var scalar in EnumerateBmpScalars(value))
        {
            // MathType keeps the word spaces of text in the Text face, where
            // they have the font's width; in its space face they have none.
            // It stores them as U+0020 (as its own TeX import of \text{a b}
            // does): its renderer writes a Text-face U+00A0 into the preview
            // WMF as byte 0xA0, which Word draws as a wide "?" (issue #35).
            if (typeface == TypefaceText && scalar == 0xA0)
                EmitScalar(0x20, typeface, output);
            else if (IsWhiteSpaceScalar(scalar)
                && !(typeface == TypefaceText && scalar == 0x20))
                EmitScalar(scalar, TypefaceSpace, output);
            else EmitScalar(
                scalar,
                IsFarEastTextScalar(scalar) || typeface == TypefaceText && NeedsFarEastCodePage(scalar)
                    ? TypefaceFarEastText : typeface,
                output);
        }
    }

    // Unicode Mathematical Alphanumeric Symbols (U+1D400-U+1D7FF) are base
    // letters, Greek letters and digits in a mathematical alphabet. MTEF v5
    // stores 16-bit MTCode with the alphabet as typeface or character style,
    // so they are written as their base character with that mathvariant.
    private static readonly string[] LatinMathAlphabets =
    {
        "bold", "italic", "bold-italic", "script", "bold-script", "fraktur", "double-struck",
        "bold-fraktur", "sans-serif", "bold-sans-serif", "sans-serif-italic", "sans-serif-bold-italic",
        "monospace",
    };

    private static readonly string[] GreekMathAlphabets =
        { "bold", "italic", "bold-italic", "bold-sans-serif", "sans-serif-bold-italic" };

    private static readonly string[] DigitMathAlphabets =
        { "bold", "double-struck", "sans-serif", "bold-sans-serif", "monospace" };

    // The 58 letters of each Greek mathematical alphabet, in Unicode order.
    private static readonly int[] GreekMathAlphabetLetters = Enumerable.Range(0x0391, 17)
        .Concat(new[] { 0x03F4 })
        .Concat(Enumerable.Range(0x03A3, 7))
        .Concat(new[] { 0x2207 })
        .Concat(Enumerable.Range(0x03B1, 25))
        .Concat(new[] { 0x2202, 0x03F5, 0x03D1, 0x03F0, 0x03D5, 0x03F1, 0x03D6 })
        .ToArray();

    internal static bool TryDecomposeMathAlphanumeric(int scalar, out int baseScalar, out string variant)
    {
        if (scalar >= 0x1D400 && scalar < 0x1D400 + 13 * 52)
        {
            var index = scalar - 0x1D400;
            variant = LatinMathAlphabets[index / 52];
            var letter = index % 52;
            baseScalar = letter < 26 ? 'A' + letter : 'a' + letter - 26;
            return true;
        }
        if (scalar is 0x1D6A4 or 0x1D6A5)
        {
            variant = "italic";
            baseScalar = scalar == 0x1D6A4 ? 0x0131 : 0x0237;
            return true;
        }
        if (scalar >= 0x1D6A8 && scalar < 0x1D6A8 + 5 * 58)
        {
            var index = scalar - 0x1D6A8;
            variant = GreekMathAlphabets[index / 58];
            baseScalar = GreekMathAlphabetLetters[index % 58];
            return true;
        }
        if (scalar is 0x1D7CA or 0x1D7CB)
        {
            variant = "bold";
            baseScalar = scalar == 0x1D7CA ? 0x03DC : 0x03DD;
            return true;
        }
        if (scalar >= 0x1D7CE && scalar < 0x1D7CE + 5 * 10)
        {
            var index = scalar - 0x1D7CE;
            variant = DigitMathAlphabets[index / 10];
            baseScalar = '0' + index % 10;
            return true;
        }
        baseScalar = scalar;
        variant = string.Empty;
        return false;
    }

    internal static void DecomposeMathematicalAlphanumerics(XElement math)
    {
        var tokens = math.DescendantsAndSelf()
            .Where(element => element.Name.LocalName is "mi" or "mn" or "mo" or "mtext" or "ms"
                && !element.HasElements
                && element.Value.IndexOf((char)0xD835) >= 0)
            .ToList();
        foreach (var token in tokens)
        {
            var text = token.Value;
            var segments = new List<(StringBuilder Text, string? Variant)>();
            for (var index = 0; index < text.Length;)
            {
                var pair = char.IsHighSurrogate(text[index]) && index + 1 < text.Length
                    && char.IsLowSurrogate(text[index + 1]);
                var scalar = pair ? char.ConvertToUtf32(text, index) : text[index];
                string? variant = null;
                var written = text.Substring(index, pair ? 2 : 1);
                if (TryDecomposeMathAlphanumeric(scalar, out var baseScalar, out var alphabet))
                {
                    variant = alphabet;
                    written = char.ConvertFromUtf32(baseScalar);
                }
                if (segments.Count > 0 && segments[segments.Count - 1].Variant == variant)
                    segments[segments.Count - 1].Text.Append(written);
                else
                    segments.Add((new StringBuilder(written), variant));
                index += pair ? 2 : 1;
            }
            if (segments.All(segment => segment.Variant is null)) continue;

            var replacements = segments.Select(segment =>
            {
                var replacement = new XElement(token.Name, token.Attributes(), segment.Text.ToString());
                if (segment.Variant is not null) replacement.SetAttributeValue("mathvariant", segment.Variant);
                return replacement;
            }).ToList();
            if (replacements.Count == 1) token.ReplaceWith(replacements[0]);
            else if (token.Parent?.Name.LocalName is "mrow" or "math" or "mstyle" or "mtd" or "msqrt")
                token.ReplaceWith(replacements);
            else token.ReplaceWith(new XElement(token.Name.Namespace + "mrow", replacements));
        }
    }

    internal static IEnumerable<int> EnumerateBmpScalars(string value)
    {
        foreach (var character in value)
        {
            if (char.IsSurrogate(character))
                throw new InvalidDataException(
                    "MathType MTEF v5 writer does not yet support non-BMP Unicode characters.");
            yield return character;
        }
    }

    internal static bool IsLetterScalar(int scalar) =>
        scalar <= char.MaxValue && char.IsLetter((char)scalar);

    private static bool IsWhiteSpaceScalar(int scalar) =>
        scalar <= char.MaxValue && char.IsWhiteSpace((char)scalar);

    internal static void EmitCharacter(char character, int typeface, List<byte> output) =>
        EmitScalar(character, typeface, output);

    // The semantic MTCode and the font's encoded position are independent.
    // Keep presence and width from a native CHAR instead of coercing every
    // font position to encoded8 or inventing a missing MTCode on rewrite.
    internal static bool TryEmitPreservedCharacter(
        XElement element,
        int typeface,
        List<byte> output,
        IReadOnlyList<byte>? embellishments = null)
    {
        var omitMtCode = IsTrueMathMlBoolean((string?)element.Attribute("data-mtef-no-mtcode"));
        var codeAttribute = element.Attribute("data-mtef-mtcode");
        if (codeAttribute is null && !omitMtCode) return false;

        int ReadPosition(string name, int maximum)
        {
            if (!int.TryParse((string?)element.Attribute(name), NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out var value) || value < 0 || value > maximum)
                throw new InvalidDataException($"Invalid preserved MathType character attribute {name}.");
            return value;
        }

        var scalar = omitMtCode ? 0 : ReadPosition("data-mtef-mtcode", ushort.MaxValue);
        byte? encoded8 = element.Attribute("data-mtef-encoded8") is null
            ? null : (byte?)ReadPosition("data-mtef-encoded8", byte.MaxValue);
        ushort? encoded16 = element.Attribute("data-mtef-encoded16") is null
            ? null : (ushort?)ReadPosition("data-mtef-encoded16", ushort.MaxValue);
        EmitScalar(scalar, typeface, output, includeEncoded8: encoded8.HasValue,
            embellishments: embellishments, encoded8Override: encoded8,
            functionStart: IsTrueMathMlBoolean((string?)element.Attribute("data-mtef-function-start")),
            encoded16Override: encoded16, omitMtCode: omitMtCode);
        return true;
    }

    internal static void EmitScalar(
        int scalar,
        int typeface,
        List<byte> output,
        bool includeEncoded8 = false,
        IReadOnlyList<byte>? embellishments = null,
        byte? encoded8Override = null,
        bool functionStart = false,
        ushort? encoded16Override = null,
        bool omitMtCode = false)
    {
        if (includeEncoded8 && encoded16Override.HasValue)
            throw new InvalidDataException("MathType CHAR cannot have both 8-bit and 16-bit font positions.");
        if (omitMtCode && !includeEncoded8 && !encoded16Override.HasValue)
            throw new InvalidDataException("MathType CHAR without MTCode requires a font position.");
        if (scalar < 0 || scalar > 0xFFFF)
            throw new InvalidDataException(
                $"MathType MTEF v5 MTCode writer does not yet support non-BMP scalar U+{scalar:X}.");
        var hasEmbellishments = embellishments is { Count: > 0 };
        var options = (byte)(includeEncoded8 ? CharEncoded8 : 0);
        if (encoded16Override.HasValue) options |= CharEncoded16;
        if (omitMtCode) options |= CharNoMtCode;
        if (hasEmbellishments) options |= CharHasEmbellishment;
        if (functionStart) options |= CharFunctionStart;
        output.Add(RecordChar);
        output.Add(options);
        WriteSigned(output, typeface);
        if (!omitMtCode)
        {
            output.Add((byte)(scalar & 0xFF));
            output.Add((byte)((scalar >> 8) & 0xFF));
        }
        if (includeEncoded8) output.Add(encoded8Override ?? (byte)scalar);
        if (encoded16Override.HasValue)
        {
            output.Add((byte)(encoded16Override.Value & 0xFF));
            output.Add((byte)(encoded16Override.Value >> 8));
        }
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

    internal static bool IsLowerGreek(int scalar) =>
        scalar is >= 0x03B1 and <= 0x03C9
        || scalar is 0x03D1 or 0x03D5 or 0x03D6 or 0x03F1 or 0x03F5;

    internal static bool IsUpperGreek(int scalar) => scalar is >= 0x0391 and <= 0x03A9;
}
