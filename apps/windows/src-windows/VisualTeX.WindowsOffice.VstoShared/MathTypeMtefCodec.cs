using System.Xml.Linq;
using static VisualTeX.WordVsto.MathTypeMtefAssembly;
using static VisualTeX.WordVsto.MathTypeMtefCharacters;
using static VisualTeX.WordVsto.MathTypeMtefFences;
using static VisualTeX.WordVsto.MathTypeMtefGlyphs;
using static VisualTeX.WordVsto.MathTypeMtefPrefix;
using static VisualTeX.WordVsto.MathTypeMtefSignature;

namespace VisualTeX.WordVsto;

/// <summary>
/// Minimal, dependency-free MTEF v5 structural codec used to update an existing
/// MathType OLE object without starting or installing MathType.  The original
/// MathType header, font definitions, equation preferences and initial size are
/// preserved byte-for-byte; only the root equation object list is replaced.
/// </summary>
internal static class MathTypeMtefCodec
{
    internal const byte MtefVersion5 = 5;
    internal const byte RecordEnd = 0;
    internal const byte RecordLine = 1;
    internal const byte RecordChar = 2;
    internal const byte RecordTemplate = 3;
    internal const byte RecordPile = 4;
    internal const byte RecordMatrix = 5;
    internal const byte RecordEmbellishment = 6;
    internal const byte RecordRuler = 7;
    internal const byte RecordFontStyleDef = 8;
    internal const byte RecordSize = 9;
    internal const byte RecordFull = 10;
    internal const byte RecordSub = 11;
    internal const byte RecordSub2 = 12;
    internal const byte RecordSym = 13;
    internal const byte RecordSubSym = 14;
    internal const byte RecordColor = 15;
    internal const byte RecordColorDef = 16;
    internal const byte RecordFontDef = 17;
    internal const byte RecordEqnPrefs = 18;
    internal const byte RecordEncodingDef = 19;

    internal const byte LineNull = 0x01;
    internal const byte CharHasEmbellishment = 0x01;
    internal const byte CharFunctionStart = 0x02;
    internal const byte CharEncoded8 = 0x04;
    internal const byte CharEncoded16 = 0x10;
    internal const byte CharNoMtCode = 0x20;

    internal const byte TypefaceText = 1;
    internal const byte TypefaceFunction = 2;
    internal const byte TypefaceVariable = 3;
    internal const byte TypefaceLowerGreek = 4;
    internal const byte TypefaceUpperGreek = 5;
    internal const byte TypefaceSymbol = 6;
    internal const byte TypefaceVector = 7;
    internal const byte TypefaceNumber = 8;
    internal const byte TypefaceMtExtra = 11;
    // MTEF v5 built-in typeface 12 is fnTEXT_FE (Far-East Text). MathType also
    // persists a few native glyphs such as reverse membership (\\ni) and
    // \\bigcirc in this slot; those are distinguished by MTCode when reading.
    internal const byte TypefaceFarEastText = 12;
    internal const string FarEastOperatorMarker = "data-mtef-fe-operator";
    internal const byte TypefaceMarker = 23;
    // MathType 7 persists expanding fence characters ((), [], {}, |, ...)
    // with typeface 22 and MTCode, not Symbol style. Using Symbol here makes
    // the same ASCII code point resolve to a different fence glyph.
    internal const byte TypefaceFence = 22;
    internal const byte TypefaceSpace = 24;

    internal const int MathTypeAlignmentMarkerMtCode = 0xEF00;
    internal const string VisualTexMtefRulerStopsAttribute =
        "data-visualtex-mtef-ruler-stops";

    internal const byte TemplateAngle = 0;
    internal const byte TemplateParen = 1;
    internal const byte TemplateBrace = 2;
    internal const byte TemplateBracket = 3;
    internal const byte TemplateBar = 4;
    internal const byte TemplateDoubleBar = 5;
    internal const byte TemplateFloor = 6;
    internal const byte TemplateCeiling = 7;
    internal const byte TemplateDoubleBracket = 8;
    internal const byte TemplateInterval = 9;
    internal const byte TemplateRoot = 10;
    internal const byte TemplateFraction = 11;
    internal const byte TemplateUnderbar = 12;
    internal const byte TemplateOverbar = 13;
    internal const byte TemplateArrow = 14;
    internal const byte TemplateIntegral = 15;
    internal const byte TemplateSum = 16;
    internal const byte TemplateProduct = 17;
    internal const byte TemplateCoproduct = 18;
    internal const byte TemplateUnion = 19;
    internal const byte TemplateIntersection = 20;
    internal const byte TemplateIntegralOperator = 21;
    internal const byte TemplateSumOperator = 22;
    internal const byte TemplateLimit = 23;
    internal const byte TemplateHorizontalBrace = 24;
    internal const byte TemplateHorizontalBracket = 25;
    internal const byte TemplateSub = 27;
    internal const byte TemplateSup = 28;
    internal const byte TemplateSubSup = 29;
    internal const byte TemplateDirac = 30;
    internal const byte TemplateVector = 31;
    internal const byte TemplateTilde = 32;
    internal const byte TemplateHat = 33;
    internal const byte TemplateArc = 34;
    internal const byte TemplateStrike = 36;
    internal const byte TemplateBox = 37;

    internal const byte EmbellDot = 2;
    internal const byte EmbellDoubleDot = 3;
    internal const byte EmbellPrime = 5;
    internal const byte EmbellDoublePrime = 6;
    internal const byte EmbellBackPrime = 7;
    internal const byte EmbellTriplePrime = 18;
    internal const byte EmbellTilde = 8;
    internal const byte EmbellHat = 9;
    internal const byte EmbellRightArrow = 11;
    internal const byte EmbellLeftArrow = 12;
    internal const byte EmbellBothArrow = 13;
    internal const byte EmbellOverbar = 17;
    // What MathType's TeX import stores for \dddot, \breve, \ddddot and
    // \underline on one character.
    internal const byte EmbellTripleDot = 4;
    internal const byte EmbellBreve = 20;
    internal const byte EmbellQuadrupleDot = 24;
    internal const byte EmbellUnderbar = 29;
    // Read-only: MathType's other embellishment palette entries (MTEF v5
    // emb_NOT, harpoons, mid bar, frown, slashes and the marks below).
    internal const byte EmbellNot = 10;
    internal const byte EmbellRightHarpoon = 14;
    internal const byte EmbellLeftHarpoon = 15;
    internal const byte EmbellMidbar = 16;
    internal const byte EmbellFrown = 19;
    internal const byte EmbellCrossBars = 21;
    internal const byte EmbellUpBar = 22;
    internal const byte EmbellDownBar = 23;
    internal const byte EmbellUnderDot = 25;
    internal const byte EmbellUnderQuadrupleDot = 28;
    internal const byte EmbellUnderTilde = 30;
    internal const byte EmbellUnderFrown = 31;
    internal const byte EmbellUnderSmile = 32;
    internal const byte EmbellUnderRightArrow = 33;
    internal const byte EmbellUnderLeftArrow = 34;
    internal const byte EmbellUnderBothArrow = 35;
    internal const byte EmbellUnderRightHarpoon = 36;
    internal const byte EmbellUnderLeftHarpoon = 37;

    internal sealed class RewriteResult
    {
        public byte[] EquationNative { get; set; } = Array.Empty<byte>();
        public byte[] Mtef { get; set; } = Array.Empty<byte>();
        public int StructureOffset { get; set; }
    }

    internal static string ReadEquationNativeMathMl(byte[] equationNative)
    {
        var mtef = ExtractNativeMtef(equationNative);
        return ReadMtefMathMl(mtef);
    }

    internal static byte[] ExtractNativeMtef(byte[] equationNative)
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

        return mtef;
    }

    // Whether two Equation Native streams hold the same equation. MathType
    // fills EQNOLEFILEHDR.cf and reserved2..reserved4 from memory it never
    // initializes, so each save of an unchanged equation writes a different
    // header. Office asks MathType to save whenever it serializes an object
    // that is open in MathType's own window (PowerPoint SaveCopyAs, Word
    // WordOpenXML). Identity is the MTEF payload the header describes.
    // Streams whose header cannot be read are compared byte for byte.
    internal static bool EquationNativeContentEquals(byte[] left, byte[] right)
    {
        if (left is null || right is null) return ReferenceEquals(left, right);
        if (!TryLocateNativeMtef(left, out var leftOffset, out var leftLength)
            || !TryLocateNativeMtef(right, out var rightOffset, out var rightLength))
            return BytesEqual(left, 0, right, 0, left.Length, right.Length);
        return BytesEqual(left, leftOffset, right, rightOffset, leftLength, rightLength);
    }

    private static bool TryLocateNativeMtef(byte[] equationNative, out int offset, out int length)
    {
        offset = length = 0;
        if (equationNative.Length < 12) return false;
        var headerLength = BitConverter.ToUInt16(equationNative, 0);
        if (headerLength < 12 || headerLength >= equationNative.Length) return false;
        var objectLength = BitConverter.ToUInt32(equationNative, 8);
        if (objectLength == 0 || objectLength > (uint)(equationNative.Length - headerLength)) return false;
        offset = headerLength;
        length = (int)objectLength;
        return true;
    }

    private static bool BytesEqual(byte[] left, int leftOffset, byte[] right, int rightOffset, int leftLength, int rightLength)
    {
        if (leftLength != rightLength) return false;
        for (var index = 0; index < leftLength; index++)
            if (left[leftOffset + index] != right[rightOffset + index]) return false;
        return true;
    }

    // MathType's Function style says a name is applied to what follows; MTEF
    // has no record for the application itself. MathML states it with U+2061,
    // which is what gives \ln and Word's m:func (and their spacing) instead of
    // upright letters set against the argument.
    private static void MarkFunctionApplications(XElement math)
    {
        foreach (var name in math.Descendants()
                     .Where(e => (string?)e.Attribute("data-mtef-run") == "function"
                         && e.Name.LocalName == "mi")
                     .ToArray())
        {
            var owner = name;
            if (name.Parent is { } script
                && script.Name.LocalName is "msub" or "msup" or "msubsup"
                && script.Elements().First() == name)
                owner = script;
            if (owner.Parent is null
                || owner.Parent.Name.LocalName is not ("math" or "mrow" or "mtd" or "msqrt" or "mstyle"
                    or "menclose" or "mpadded" or "mphantom" or "mfenced"))
                continue;
            if (owner.Parent.Name.LocalName == "mfenced" && owner.Parent.Elements().Count() != 1
                && owner.Parent.Attribute("separators") is null)
                continue;
            var argument = owner.ElementsAfterSelf().FirstOrDefault();
            if (argument is null || argument.Name.LocalName is "mspace" or "mtext") continue;
            if (argument.Name.LocalName == "mo" && argument.Value.Trim() is not ("(" or "[" or "{")) continue;
            if ((string?)argument.Attribute("data-mtef-run") == "function") continue;
            owner.AddAfterSelf(new XElement("mo", ((char)0x2061).ToString()));
        }
    }

    private static string ReadMtefMathMl(byte[] mtef)
    {
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
        MarkFunctionApplications(math);

        // MtefStructureReader deliberately builds compact XElement trees using
        // local names (mi/mo/mrow/...). Once those nodes are attached below a
        // namespaced <math>, LINQ to XML serializes them with xmlns="" unless we
        // promote their XName as well. That produced visually plausible MathML
        // which our LaTeX reader tolerated, but standards-based MathML->OMML XSLT
        // saw an empty equation. Normalize every parser-owned descendant into the
        // MathML namespace at the direct MTEF read boundary.
        // Sibling native script/embellishment records can build nested XML
        // without recursively nested native containers. Bound that resulting
        // tree as well, before serialization or downstream recursive rules.
        var pending = new Stack<(XElement Element, int Depth)>();
        pending.Push((math, 1));
        var elements = 1;
        while (pending.Count != 0)
        {
            var (element, depth) = pending.Pop();
            if (depth > FormulaXml.MaximumDepth)
                throw new InvalidDataException($"MathType formula exceeds the nesting limit of {FormulaXml.MaximumDepth} levels.");
            if (element.Name.Namespace == XNamespace.None)
                element.Name = mathMlNamespace + element.Name.LocalName;
            element.Attribute("data-mtef-run")?.Remove();
            for (var node = element.LastNode; node is not null; node = node.PreviousNode)
            {
                if (node is not XElement child) continue;
                if (++elements > FormulaXml.MaximumElements)
                    throw new InvalidDataException($"MathType formula exceeds the limit of {FormulaXml.MaximumElements} elements.");
                pending.Push((child, depth + 1));
            }
        }
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

        var document = FormulaXml.Parse(mathMl, "MathML");
        var math = document.Root?.DescendantsAndSelf()
            .FirstOrDefault(element => element.Name.LocalName == "math")
            ?? throw new InvalidDataException("MathML has no <math> root.");
        NormalizeMathMlPrescripts(math);
        DecomposeMathematicalAlphanumerics(math);
        var generated = BuildRootStructure(math, seedMtef, out var prefixDefinitions);
        var mtef = AssembleRewrittenMtef(
            seedMtef,
            prefix.Length,
            prefixDefinitions,
            generated,
            out var structureOffset);

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
        return ApplyNativeTextStyle(new RewriteResult
        {
            EquationNative = equationNative,
            Mtef = mtef,
            StructureOffset = structureOffset,
        }, math);
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

        var document = FormulaXml.Parse(mathMl, "MathML");
        var math = document.Root?.DescendantsAndSelf()
            .FirstOrDefault(element => element.Name.LocalName == "math")
            ?? throw new InvalidDataException("MathML has no <math> root.");
        NormalizeMathMlPrescripts(math);
        DecomposeMathematicalAlphanumerics(math);

        // Preserve MathType's equation-options byte exactly as stored in the
        // source MTEF. Word controls whether this OLE sits inline or as a display
        // object; rewriting MathType's own inline bit caused genuine MathType 7
        // character styles (notably MT Extra blackboard-bold glyphs) to be
        // reinterpreted on reopen even though the root MTEF bytes were correct.
        // The `inline` parameter remains part of the public rewrite contract for
        // callers, but must not mutate the MathType-native header.
        _ = inline;
        var generated = BuildRootStructure(math, sourceMtef, out var prefixDefinitions);
        var rewrittenMtef = AssembleRewrittenMtef(
            sourceMtef,
            structureOffset,
            prefixDefinitions,
            generated,
            out var rewrittenStructureOffset);

        var rewrittenNative = new byte[headerLength + rewrittenMtef.Length];
        Buffer.BlockCopy(equationNative, 0, rewrittenNative, 0, headerLength);
        Buffer.BlockCopy(rewrittenMtef, 0, rewrittenNative, headerLength, rewrittenMtef.Length);
        var objectLengthBytes = BitConverter.GetBytes((uint)rewrittenMtef.Length);
        Buffer.BlockCopy(objectLengthBytes, 0, rewrittenNative, 8, objectLengthBytes.Length);

        return ApplyNativeTextStyle(new RewriteResult
        {
            EquationNative = rewrittenNative,
            Mtef = rewrittenMtef,
            StructureOffset = rewrittenStructureOffset,
        }, math);
    }

    private static RewriteResult ApplyNativeTextStyle(RewriteResult source, XElement math)
    {
        // The native Text logical face is declared by EQN_PREFS. Its font and
        // character bits are visible formatting, not a private token annotation.
        // Keep all literal CHARs in that role (including '=' and whitespace).
        var variants = math.DescendantsAndSelf().Where(e => e.Name.LocalName == "mtext" && e.Value.Length > 0)
            .Select(e => MathMlTokenStyleContract.ResolveEffectiveVariant(e) ?? "normal")
            .Distinct(StringComparer.Ordinal).ToArray();
        if (variants.Length == 0) return source;
        // One logical face cannot express multiple text fonts. Leave the native
        // source binding intact; the strict styled-text comparison detects any
        // remaining mismatch instead of pretending that a style was preserved.
        if (variants.Length != 1) return source;
        var variant = variants[0];
        var font = variant.Contains("sans-serif") ? "Arial"
            : variant == "monospace" ? "Courier New" : "Times New Roman";
        if (variant is not ("normal" or "upright" or "bold" or "italic" or "bold-italic"
            or "sans-serif" or "bold-sans-serif" or "sans-serif-italic" or "sans-serif-bold-italic" or "monospace"))
            return source;
        var bits = (byte)((variant.Contains("bold") ? 1 : 0) | (variant.Contains("italic") ? 2 : 0));
        var layout = ReadPrefixLayout(source.Mtef);
        var fontRecords = layout.Records.Concat(ReadRootLeadingFormattingRecords(source.Mtef))
            .Where(r => r.Type is RecordFontDef or RecordFontStyleDef or RecordEncodingDef).ToArray();
        var fonts = fontRecords.Where(r => r.Type == RecordFontDef).Select(r =>
        {
            var position = r.Start + 1;
            var encoding = ReadUnsigned(source.Mtef, ref position);
            return (Encoding: encoding, Name: ReadPrefixNullTerminatedString(source.Mtef, ref position));
        }).ToArray();
        var bindings = ReadNativeCharacterStyleBindings(source.Mtef);
        if (!variant.Contains("sans-serif") && variant != "monospace"
            && bindings.TryGetValue(TypefaceText, out var inheritedText)
            && inheritedText.FontDefinition > 0 && inheritedText.FontDefinition <= fonts.Length)
            font = fonts[inheritedText.FontDefinition - 1].Name;
        if (bindings.TryGetValue(TypefaceText, out var binding)
            && binding.FontDefinition > 0 && binding.FontDefinition <= fonts.Length
            && fonts[binding.FontDefinition - 1].Name == font && binding.CharacterStyle == bits)
            return source;

        var fontIndex = Array.FindIndex(fonts, f => f.Name == font) + 1;
        var definitions = fontRecords.SelectMany(r => source.Mtef.Skip(r.Start).Take(r.End - r.Start)).ToList();
        if (fontIndex == 0)
        {
            fontIndex = fonts.Length + 1;
            var latin = Array.FindIndex(fonts, f => f.Name == "Times New Roman");
            var encoding = latin >= 0 ? fonts[latin].Encoding : 0;
            if (encoding == 0)
            {
                encoding = 5 + fontRecords.Count(r => r.Type == RecordEncodingDef);
                definitions.Add(RecordEncodingDef);
                WriteNullTerminatedAscii(definitions, "WindowsANSI");
            }
            definitions.Add(RecordFontDef);
            WriteUnsigned(definitions, encoding);
            WriteNullTerminatedAscii(definitions, font);
        }
        var prefs = layout.Preferences;
        var cursor = SkipDimensionArray(source.Mtef, prefs + 2);
        cursor = SkipDimensionArray(source.Mtef, cursor);
        Require(source.Mtef, cursor, 1);
        if (source.Mtef[cursor++] < TypefaceText) throw new InvalidDataException("Native equation has no Text style slot.");
        var oldStart = cursor;
        var oldFont = ReadUnsigned(source.Mtef, ref cursor);
        if (oldFont != 0) { Require(source.Mtef, cursor, 1); cursor++; }
        var newBinding = new List<byte>();
        WriteUnsigned(newBinding, fontIndex); newBinding.Add(bits);

        // Move font definitions, in their original index order, ahead of the
        // preference that references them. Other prefix/root records are exact.
        var skipped = fontRecords.ToDictionary(r => r.Start, r => r.End);
        var bytes = new List<byte>(source.Mtef.Length + definitions.Count + newBinding.Count);
        for (var position = 0; position < source.Mtef.Length;)
        {
            if (position == prefs) bytes.AddRange(definitions);
            if (skipped.TryGetValue(position, out var end)) { position = end; continue; }
            if (position == oldStart) { bytes.AddRange(newBinding); position = cursor; continue; }
            bytes.Add(source.Mtef[position++]);
        }
        var mtef = bytes.ToArray();
        var native = source.EquationNative.Take(28).Concat(mtef).ToArray();
        BitConverter.GetBytes((uint)mtef.Length).CopyTo(native, 8);
        return new RewriteResult { Mtef = mtef, EquationNative = native, StructureOffset = FindRootStructureOffset(mtef) };
    }

    internal static string SemanticSignature(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)) return string.Empty;
        var document = FormulaXml.Parse(mathMl, "MathML");
        var math = document.Root?.DescendantsAndSelf()
            .FirstOrDefault(element => element.Name.LocalName == "math")
            ?? document.Root
            ?? throw new InvalidDataException("MathML has no root element.");
        var normalized = new XElement(math);
        NormalizeMathMlPrescripts(normalized);
        DecomposeMathematicalAlphanumerics(normalized);
        NormalizeMathTypeStoredGlyphs(normalized);
        NormalizeMathTypeExportedTabPiles(normalized);
        MaterializeLooseFenceGroups(normalized);
        var structure = CanonicalizeMathMl(normalized, inheritedMathVariant: null);
        var tokenColors = normalized.DescendantsAndSelf()
            .Where(element => element.Name.LocalName is "mi" or "mn" or "mo" or "mtext"
                && !string.IsNullOrWhiteSpace(element.Value))
            .Select(element => MathMlColorContract.Effective(element))
            .Select(color => string.Equals(color, "#000000", StringComparison.OrdinalIgnoreCase)
                ? null
                : color)
            .ToArray();
        return tokenColors.Any(color => color is not null)
            ? structure + "|colors(" + string.Join(",", tokenColors.Select(color => color ?? "-")) + ")"
            : structure;
    }

    internal static int ReadUnsigned(byte[] data, ref int position)
    {
        Require(data, position, 1);
        var value = data[position++];
        if (value != 0xFF) return value;
        Require(data, position, 2);
        var expanded = data[position] | (data[position + 1] << 8);
        position += 2;
        return expanded;
    }

    internal static void WriteUnsigned(List<byte> output, int value)
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

    internal static void Require(byte[] data, int offset, int count)
    {
        if (offset < 0 || count < 0 || offset + count > data.Length)
            throw new EndOfStreamException(
                $"MTEF stream is truncated at offset {offset}, need {count} bytes.");
    }
}
