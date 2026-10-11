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
internal static partial class MathTypeMtefCodec
{
    private const byte MtefVersion5 = 5;
    private const byte RecordEnd = 0;
    private const byte RecordLine = 1;
    private const byte RecordChar = 2;
    private const byte RecordTemplate = 3;
    private const byte RecordPile = 4;
    private const byte RecordMatrix = 5;
    private const byte RecordEmbellishment = 6;
    private const byte RecordRuler = 7;
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
    private const byte CharEncoded16 = 0x10;
    private const byte CharNoMtCode = 0x20;

    private const byte TypefaceText = 1;
    private const byte TypefaceFunction = 2;
    private const byte TypefaceVariable = 3;
    private const byte TypefaceLowerGreek = 4;
    private const byte TypefaceUpperGreek = 5;
    private const byte TypefaceSymbol = 6;
    private const byte TypefaceVector = 7;
    private const byte TypefaceNumber = 8;
    private const byte TypefaceMtExtra = 11;
    // MTEF v5 built-in typeface 12 is fnTEXT_FE (Far-East Text). MathType also
    // persists a few native glyphs such as reverse membership (\\ni) and
    // \\bigcirc in this slot; those are distinguished by MTCode when reading.
    private const byte TypefaceFarEastText = 12;
    private const string FarEastOperatorMarker = "data-mtef-fe-operator";
    private const byte TypefaceMarker = 23;
    // MathType 7 persists expanding fence characters ((), [], {}, |, ...)
    // with typeface 22 and MTCode, not Symbol style. Using Symbol here makes
    // the same ASCII code point resolve to a different fence glyph.
    private const byte TypefaceFence = 22;
    private const byte TypefaceSpace = 24;

    private const int MathTypeAlignmentMarkerMtCode = 0xEF00;
    private const string VisualTexMtefRulerStopsAttribute =
        "data-visualtex-mtef-ruler-stops";

    private const byte TemplateAngle = 0;
    private const byte TemplateParen = 1;
    private const byte TemplateBrace = 2;
    private const byte TemplateBracket = 3;
    private const byte TemplateBar = 4;
    private const byte TemplateDoubleBar = 5;
    private const byte TemplateFloor = 6;
    private const byte TemplateCeiling = 7;
    private const byte TemplateInterval = 9;
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
    private const byte TemplateDirac = 30;
    private const byte TemplateVector = 31;
    private const byte TemplateTilde = 32;
    private const byte TemplateHat = 33;
    private const byte TemplateArc = 34;
    private const byte TemplateStrike = 36;
    private const byte TemplateBox = 37;

    private const byte EmbellDot = 2;
    private const byte EmbellDoubleDot = 3;
    private const byte EmbellPrime = 5;
    private const byte EmbellDoublePrime = 6;
    private const byte EmbellBackPrime = 7;
    private const byte EmbellTriplePrime = 18;
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

    private static void NormalizeMathTypeExportedTabPiles(XElement root)
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

    private static string CanonicalizeMathMl(XElement element, string? inheritedMathVariant)
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
                var intrinsicVariant = string.IsNullOrWhiteSpace(variant) && value.Length == 1
                    && TryNormalizeLetterlikeForWrite(value[0], out _, out _);
                return CanonicalToken("mi", value, intrinsicVariant ? null
                    : MathMlTokenStyleContract.ResolveEffectiveVariant(element, variant));
            }
            case "mn":
            {
                // Numbers default to upright, not to identifier italic. Their
                // explicit variant is part of the contract too: ignoring it
                // makes a bold/monospace digit loss invisible to validation.
                var numberVariant = (variant ?? string.Empty).Trim().ToLowerInvariant();
                return numberVariant.Length == 0 || numberVariant is "normal" or "upright"
                    ? "n(" + element.Value.Trim() + ")"
                    : "n[" + numberVariant + "](" + element.Value.Trim() + ")";
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
                return CanonicalTernary("subsup", children, variant);
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
                return "phantom(" + Children() + ")";
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

    private static bool IsMathematicalSymbolToken(string value) =>
        value.Length > 0 && value.All(character => IsMathematicalSymbolScalar(character));

    private static bool IsScalarOperatorRun(string value) =>
        value.Length > 1 && value.All(character =>
            IsMathematicalSymbolScalar(character) || character is ':' or '-');

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

    private static bool IsTrueMathMlBoolean(string? value) =>
        string.Equals(value?.Trim(), "true", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value?.Trim(), "1", StringComparison.Ordinal);

    private static bool IsFalseMathMlBoolean(string? value) =>
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

    private static bool IsZeroMathMlLength(string value)
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

    internal static int FindRootStructureOffset(byte[] mtef)
        => ReadPrefixLayout(mtef).Root;

    private sealed class MtefPrefixLayout
    {
        internal MtefPrefixLayout(int root, int preferences, int colors,
            IReadOnlyList<(byte Type, int Start, int End)> records)
        { Root = root; Preferences = preferences; Colors = colors; Records = records; }
        internal int Root { get; }
        internal int Preferences { get; }
        internal int Colors { get; }
        internal IReadOnlyList<(byte Type, int Start, int End)> Records { get; }
    }

    // The sole prefix boundary parser. Consumers inspect its typed record spans;
    // they must not independently guess where preferences, fonts, sizes or the
    // root begin. This also preserves unknown length-delimited vendor records.
    private static MtefPrefixLayout ReadPrefixLayout(byte[] mtef)
    {
        if (mtef is null || mtef.Length < 16 || mtef[0] != MtefVersion5)
            throw new InvalidDataException("Invalid or unsupported MTEF stream.");
        var position = FindEquationOptionsOffset(mtef) + 1;
        var sawPreferences = false;
        var preferencesOffset = -1;
        var sawInitialSize = false;
        var colorDefinitions = 0;
        var records = new List<(byte Type, int Start, int End)>();
        while (position < mtef.Length)
        {
            var recordStart = position;
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
                    colorDefinitions++;
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
                            System.Diagnostics.Trace.WriteLine(
                                $"mathtype-mtef-root-without-initial-size offset={position} record={record}");
                        return new MtefPrefixLayout(position, preferencesOffset, colorDefinitions, records);
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
                    // Never search arbitrary preference/extension bytes for a
                    // plausible LINE. A malformed boundary must fail before any
                    // Word mutation, not silently adopt a nested sub-expression.
                    // A genuine empty MathType equation is encoded as the normal
                    // prefix/initial-size state followed directly by the final
                    // equation END. MathType 7.8.x therefore legitimately has a
                    // zero at the same offset where non-empty DSMT7 equations put
                    // their root LINE. Treat only the terminal END as empty; any
                    // other zero remains a hard parse failure.
                    if (record == RecordEnd && position == mtef.Length - 1)
                    {
                        System.Diagnostics.Trace.WriteLine(
                            $"mathtype-mtef-empty-equation endOffset={position}");
                        return new MtefPrefixLayout(position, preferencesOffset, colorDefinitions, records);
                    }
                    throw new InvalidDataException(
                        $"Unsupported MathType root record {record} at offset {position}.");
            }
            if (position <= recordStart || position > mtef.Length)
                throw new InvalidDataException($"Invalid MTEF prefix extent at {recordStart}.");
            records.Add((record, recordStart, position));
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

    internal static string PrepareMathMlForMathTypeInterop(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException("MathType conversion requires MathML.");
        var document = FormulaXml.Parse(mathMl, "MathML");
        var math = document.Root?.DescendantsAndSelf()
            .FirstOrDefault(element => element.Name.LocalName == "math")
            ?? throw new InvalidDataException("MathML has no <math> root.");
        DecomposeMathematicalAlphanumerics(math);
        return PrepareMathForMathType(math)
            .ToString(SaveOptions.DisableFormatting);
    }

    private static XElement PrepareMathForMathType(XElement math)
    {
        var preparedMath = new XElement(math);
        MaterializeMfencedSeparators(preparedMath);
        MaterializeInheritedMathVariants(preparedMath, inheritedMathVariant: null);
        NormalizeMathJaxFenceRows(preparedMath);
        MaterializeLooseFenceGroups(preparedMath);
        return preparedMath;
    }

    private static void MaterializeMfencedSeparators(XElement root)
    {
        // MathML's mfenced implicitly inserts operators between operands; a
        // ParBox has a single LINE slot and therefore needs those real CHARs.
        // Expand before resolving inherited styles so separators obey the same
        // token rules as explicit mo elements. One operand needs no expansion.
        foreach (var fence in root.DescendantsAndSelf()
                     .Where(element => element.Name.LocalName == "mfenced").Reverse().ToArray())
        {
            var children = fence.Elements().ToArray();
            if (children.Length < 2) continue;
            var separators = ((string?)fence.Attribute("separators") ?? ",")
                .Where(character => !char.IsWhiteSpace(character)).ToArray();
            var body = new XElement(fence.Name.Namespace + "mrow");
            for (var index = 0; index < children.Length; index++)
            {
                if (index > 0 && separators.Length > 0)
                    body.Add(new XElement(fence.Name.Namespace + "mo",
                        separators[Math.Min(index - 1, separators.Length - 1)].ToString()));
                body.Add(children[index]);
            }
            fence.ReplaceNodes(body);
            fence.SetAttributeValue("separators", string.Empty);
        }
    }

    private sealed class MtefColorPlan
    {
        internal Dictionary<string, int> Indexes { get; } =
            new(StringComparer.OrdinalIgnoreCase);
        internal List<byte> InsertedDefinitions { get; } = new();
        internal int InsertedCount { get; set; }
        internal int PrefixDefinitionCount { get; set; }
    }

    private static string? ReadRgbColorDefinition(
        byte[] data,
        (byte Type, int Start, int End) record)
    {
        if (record.Type != RecordColorDef) return null;
        Require(data, record.Start, 8);
        var options = data[record.Start + 1];
        if ((options & 0x01) != 0) return null; // CMYK is preserved but not reused for RGB input.
        var red = data[record.Start + 2] | (data[record.Start + 3] << 8);
        var green = data[record.Start + 4] | (data[record.Start + 5] << 8);
        var blue = data[record.Start + 6] | (data[record.Start + 7] << 8);
        if (red > 1000 || green > 1000 || blue > 1000) return null;
        static int ToByte(int component) =>
            Math.Min(255, Math.Max(0, (int)Math.Round(component * 255d / 1000d,
                MidpointRounding.AwayFromZero)));
        return $"#{ToByte(red):X2}{ToByte(green):X2}{ToByte(blue):X2}";
    }

    private static void AppendRgbColorDefinition(string color, List<byte> output)
    {
        var normalized = MathMlColorContract.Normalize(color)
            ?? throw new InvalidDataException($"Unsupported MathType formula color '{color}'.");
        static ushort ToMtef(byte component) => checked((ushort)Math.Round(
            component * 1000d / 255d,
            MidpointRounding.AwayFromZero));
        var red = ToMtef(byte.Parse(normalized.Substring(1, 2), NumberStyles.HexNumber,
            CultureInfo.InvariantCulture));
        var green = ToMtef(byte.Parse(normalized.Substring(3, 2), NumberStyles.HexNumber,
            CultureInfo.InvariantCulture));
        var blue = ToMtef(byte.Parse(normalized.Substring(5, 2), NumberStyles.HexNumber,
            CultureInfo.InvariantCulture));
        output.AddRange(new byte[]
        {
            RecordColorDef,
            0,
            (byte)(red & 0xFF), (byte)(red >> 8),
            (byte)(green & 0xFF), (byte)(green >> 8),
            (byte)(blue & 0xFF), (byte)(blue >> 8),
        });
    }

    private static MtefColorPlan BuildColorPlan(
        XElement preparedMath,
        byte[] sourceMtef,
        IReadOnlyList<(byte Type, int Start, int End)> rootFormattingRecords,
        bool requiresNativeBlack)
    {
        var plan = new MtefColorPlan();
        var prefix = ReadPrefixLayout(sourceMtef);
        var prefixColors = prefix.Records.Where(record => record.Type == RecordColorDef).ToArray();
        var rootColors = rootFormattingRecords.Where(record => record.Type == RecordColorDef).ToArray();
        plan.PrefixDefinitionCount = prefixColors.Length;

        var prefixValues = prefixColors.Select(record => ReadRgbColorDefinition(sourceMtef, record)).ToArray();
        var rootValues = rootColors.Select(record => ReadRgbColorDefinition(sourceMtef, record)).ToArray();
        for (var index = 0; index < prefixValues.Length; index++)
            if (prefixValues[index] is { } color && !plan.Indexes.ContainsKey(color))
                plan.Indexes[color] = index + 1;

        var requested = preparedMath.DescendantsAndSelf()
            .Select(MathMlColorContract.Explicit)
            .Where(color => color is not null)
            .Cast<string>()
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (requiresNativeBlack && prefixValues.Length + rootValues.Length == 0)
            requested.Insert(0, "#000000");

        var newColors = new List<string>();
        foreach (var color in requested)
        {
            if (plan.Indexes.ContainsKey(color)
                || rootValues.Any(existing => string.Equals(existing, color,
                    StringComparison.OrdinalIgnoreCase))
                || newColors.Contains(color, StringComparer.OrdinalIgnoreCase))
                continue;
            newColors.Add(color);
        }
        plan.InsertedCount = newColors.Count;
        for (var index = 0; index < newColors.Count; index++)
        {
            var color = newColors[index];
            plan.Indexes[color] = prefixValues.Length + index + 1;
            AppendRgbColorDefinition(color, plan.InsertedDefinitions);
        }
        for (var index = 0; index < rootValues.Length; index++)
            if (rootValues[index] is { } color && !plan.Indexes.ContainsKey(color))
                plan.Indexes[color] = prefixValues.Length + plan.InsertedCount + index + 1;
        return plan;
    }

    // MTEF logical sizes and colors are stream state, not properties local to a TMPL.
    // Carry that state with this one emission buffer so nested templates can
    // restore their parent size without global/thread-local mutable state.
    private sealed class MtefWriter : List<byte>
    {
        private readonly IReadOnlyDictionary<string, int> _colorIndexes;
        internal byte ContentSize { get; private set; } = RecordFull;
        internal byte SymbolSize => ContentSize == RecordFull ? RecordSym : RecordSubSym;
        internal int CurrentColorIndex { get; private set; }
        internal int? ExplicitColorIndex { get; private set; }

        internal MtefWriter(
            IReadOnlyDictionary<string, int>? colorIndexes = null,
            int currentColorIndex = 0,
            int? explicitColorIndex = null)
        {
            _colorIndexes = colorIndexes ?? new Dictionary<string, int>();
            CurrentColorIndex = currentColorIndex;
            ExplicitColorIndex = explicitColorIndex;
        }

        // Scratch emission is inspected as one standalone OBJECT and then copied
        // into this writer. Stream state already exists in the destination; if
        // it is inherited here, a native COLOR transition precedes the scratch
        // CHAR and prevents the compact EMBELL representation from being used.
        // Keep the palette available for an explicitly colored scratch node,
        // but start its detached byte stream from neutral state.
        internal MtefWriter CreateScratch() => new(_colorIndexes);

        internal IDisposable EnterExplicitColor(string color)
        {
            var normalized = MathMlColorContract.Normalize(color)
                ?? throw new InvalidDataException($"Unsupported MathType formula color '{color}'.");
            if (!_colorIndexes.TryGetValue(normalized, out var index))
                throw new InvalidDataException(
                    $"MathType formula color {normalized} has no MTEF palette definition.");
            return new ColorScope(this, index);
        }

        internal void EmitNativeColor(int nativeDefaultIndex)
        {
            var index = ExplicitColorIndex ?? nativeDefaultIndex;
            Add(RecordColor);
            WriteUnsigned(this, index);
            CurrentColorIndex = index;
        }

        internal IDisposable EnterScriptSize()
        {
            var parent = ContentSize;
            ContentSize = parent == RecordFull ? RecordSub : RecordSub2;
            Add(ContentSize);
            return new SizeScope(this, parent);
        }

        internal void RestoreContentSize() => Add(ContentSize);

        private sealed class SizeScope : IDisposable
        {
            private MtefWriter? _writer;
            private readonly byte _parent;
            internal SizeScope(MtefWriter writer, byte parent)
            { _writer = writer; _parent = parent; }
            public void Dispose()
            {
                var writer = _writer;
                if (writer is null) return;
                _writer = null;
                writer.ContentSize = _parent;
                writer.Add(_parent);
            }
        }

        private sealed class ColorScope : IDisposable
        {
            private MtefWriter? _writer;
            private readonly int _previousColorIndex;
            private readonly int? _previousExplicitColorIndex;

            internal ColorScope(MtefWriter writer, int index)
            {
                _writer = writer;
                _previousColorIndex = writer.CurrentColorIndex;
                _previousExplicitColorIndex = writer.ExplicitColorIndex;
                writer.ExplicitColorIndex = index;
                if (writer.CurrentColorIndex != index)
                    writer.EmitNativeColor(index);
            }

            public void Dispose()
            {
                var writer = _writer;
                if (writer is null) return;
                _writer = null;
                writer.ExplicitColorIndex = _previousExplicitColorIndex;
                if (writer.CurrentColorIndex != _previousColorIndex)
                {
                    writer.Add(RecordColor);
                    WriteUnsigned(writer, _previousColorIndex);
                    writer.CurrentColorIndex = _previousColorIndex;
                }
            }
        }
    }

    private static byte[] BuildRootStructure(
        XElement math,
        byte[] sourceMtef,
        out byte[] prefixDefinitions)
    {
        var prefixLayout = ReadPrefixLayout(sourceMtef);
        var rootFormattingRecords = ReadRootLeadingFormattingRecords(sourceMtef);
        var retainedColors = prefixLayout.Colors
            + rootFormattingRecords.Count(record => record.Type == RecordColorDef);
        var preparedMath = PrepareMathForMathType(math);
        prefixDefinitions = Array.Empty<byte>();
        var topLevelElements = SignificantChildren(preparedMath)
            .OfType<XElement>()
            .Where(element => element.Name.LocalName is not ("annotation" or "annotation-xml"))
            .ToArray();
        var hasExplicitColor = preparedMath.DescendantsAndSelf()
            .Any(element => MathMlColorContract.Explicit(element) is not null);
        if (!hasExplicitColor
            && topLevelElements.Length == 1
            && topLevelElements[0].Name.LocalName == "mtable")
        {
            // Font/encoding definitions are global MTEF records. MathType accepts
            // them inside an ordinary root LINE, but putting them inside a root
            // PILE makes its native editor preserve the glyphs while degrading
            // fnMARKER tabs into empty text. Build the aligned table on a clone,
            // emit any missing definitions into the global prefix instead, then
            // keep the PILE object list identical to MathType's native layout.
            var alignedTable = new XElement(topLevelElements[0]);
            // A genuine native aligned PILE does not reference COLOR 1 at all.
            // Older code injected a default black COLOR_DEF whenever the source
            // had no palette, even when the generated root had no COLOR record;
            // that needlessly rewrote otherwise byte-identical MathType equations.
            // Only nested constructs whose emitters actually use COLOR 1 need the
            // synthetic definition. The top-level aligned table itself does not.
            var alignedPrefixDefinitions = new List<byte>();
            if (retainedColors == 0
                && MathTypeEmissionRequiresColorDefinition(
                    preparedMath,
                    topLevelTableIsAlignedPile: true))
            {
                alignedPrefixDefinitions.AddRange(
                    new byte[] { RecordColorDef, 0, 0, 0, 0, 0, 0, 0 });
            }
            EmitExplicitFontDefinitions(
                alignedTable,
                sourceMtef,
                alignedPrefixDefinitions);
            var alignedRoot = new MtefWriter();
            if (TryEmitAlignedPile(alignedTable, sourceMtef, alignedRoot))
            {
                prefixDefinitions = alignedPrefixDefinitions.ToArray();
                // MathType 7's own multi-point aligned equations use PILE as the
                // top-level equation structure. Keeping a synthetic root LINE
                // around that PILE causes MathType to export fnMARKER tabs as
                // literal text instead of alignment boundaries.
                alignedRoot.Add(RecordEnd); // equation
                return alignedRoot.ToArray();
            }
        }

        // The ordinary template/matrix writer uses a native palette entry for
        // several structures. Add that black entry only when the original had no
        // palette, then append the exact RGB colors requested by the editor.
        var colorPlan = BuildColorPlan(
            preparedMath,
            sourceMtef,
            rootFormattingRecords,
            MathTypeEmissionRequiresColorDefinition(
                preparedMath,
                topLevelTableIsAlignedPile: false));
        prefixDefinitions = colorPlan.InsertedDefinitions.ToArray();

        var rootFormatting = new List<byte>();
        // Once MathML carries explicit colors, those attributes are the source
        // of truth for the rewritten stream. A leading MTEF COLOR is state for
        // the first object, not an independent document default. Re-copying it
        // makes a later uncolored sibling inherit that first object's color.
        // Keep definitions and all other native formatting, but let EmitNode
        // recreate the COLOR transitions from the actual MathML scopes.
        var retainedRootFormatting = hasExplicitColor
            ? rootFormattingRecords.Where(record => record.Type != RecordColor).ToArray()
            : rootFormattingRecords;
        CopyRootLeadingFormattingRecords(
            sourceMtef,
            retainedRootFormatting,
            rootFormatting,
            colorPlan.PrefixDefinitionCount,
            colorPlan.InsertedCount,
            out var currentColorIndex);

        var output = new MtefWriter(colorPlan.Indexes, currentColorIndex)
        {
            RecordLine,
            0,
        };
        // Preserve MathType's root-level formatting state (notably COLOR_DEF /
        // COLOR) from the source equation. MathType 7 writes these records even
        // when the visible formula is plain black; dropping them produced OLEs
        // that our own parser could read but the native MathType server could
        // reinterpret with the wrong expandable-fence family.
        output.AddRange(rootFormatting);
        EmitExplicitFontDefinitions(preparedMath, sourceMtef, output);
        EmitNode(preparedMath, output);
        output.Add(RecordEnd); // root line
        output.Add(RecordEnd); // equation
        return output.ToArray();
    }

    private static byte[] AssembleRewrittenMtef(
        byte[] sourceMtef,
        int sourceStructureOffset,
        byte[] prefixDefinitions,
        byte[] rootStructure,
        out int rewrittenStructureOffset)
    {
        if (prefixDefinitions.Length == 0)
        {
            var direct = new byte[sourceStructureOffset + rootStructure.Length];
            Buffer.BlockCopy(sourceMtef, 0, direct, 0, sourceStructureOffset);
            Buffer.BlockCopy(
                rootStructure,
                0,
                direct,
                sourceStructureOffset,
                rootStructure.Length);
            rewrittenStructureOffset = sourceStructureOffset;
            return direct;
        }

        // MathType's global ENCODING_DEF/FONT_DEF/FONT_STYLE_DEF records belong
        // before the initial SIZE record. Putting them after SIZE or inside a root
        // PILE makes MathType preserve the styled glyphs but stop interpreting
        // fnMARKER/U+0009 as alignment tabs. Insert only the newly required
        // definitions at this prefix boundary and keep every original prefix byte.
        var insertionOffset = FindInitialSizeRecordOffset(
            sourceMtef,
            sourceStructureOffset);
        var rewritten = new byte[
            sourceStructureOffset + prefixDefinitions.Length + rootStructure.Length];
        Buffer.BlockCopy(sourceMtef, 0, rewritten, 0, insertionOffset);
        Buffer.BlockCopy(
            prefixDefinitions,
            0,
            rewritten,
            insertionOffset,
            prefixDefinitions.Length);
        Buffer.BlockCopy(
            sourceMtef,
            insertionOffset,
            rewritten,
            insertionOffset + prefixDefinitions.Length,
            sourceStructureOffset - insertionOffset);
        rewrittenStructureOffset = sourceStructureOffset + prefixDefinitions.Length;
        Buffer.BlockCopy(
            rootStructure,
            0,
            rewritten,
            rewrittenStructureOffset,
            rootStructure.Length);
        return rewritten;
    }

    private static int FindInitialSizeRecordOffset(byte[] mtef, int rootOffset)
    {
        var prefix = ReadPrefixLayout(mtef);
        if (prefix.Root != rootOffset)
            throw new InvalidDataException("MTEF rewrite root does not match its validated prefix.");
        foreach (var record in prefix.Records)
            if (record.Type == RecordSize || record.Type >= RecordFull && record.Type <= RecordSubSym)
                return record.Start;
        return prefix.Root;
    }

    private static void CopyRootLeadingFormattingRecords(
        byte[] sourceMtef,
        IReadOnlyList<(byte Type, int Start, int End)> records,
        List<byte> output,
        int prefixDefinitionCount,
        int insertedDefinitionCount,
        out int currentColorIndex)
    {
        currentColorIndex = ReadPrefixCurrentColorIndex(sourceMtef);
        foreach (var record in records)
        {
            if (record.Type == RecordColor && insertedDefinitionCount != 0)
            {
                var position = record.Start + 1;
                var index = ReadUnsigned(sourceMtef, ref position);
                var remapped = index > prefixDefinitionCount
                    ? index + insertedDefinitionCount
                    : index;
                output.Add(RecordColor);
                WriteUnsigned(output, remapped);
                currentColorIndex = remapped;
                continue;
            }
            for (var index = record.Start; index < record.End; index++)
                output.Add(sourceMtef[index]);
            if (record.Type == RecordColor)
            {
                var position = record.Start + 1;
                currentColorIndex = ReadUnsigned(sourceMtef, ref position);
            }
        }
    }

    private static void CopyRootLeadingFormattingRecords(
        byte[] sourceMtef,
        List<byte> output)
    {
        var records = ReadRootLeadingFormattingRecords(sourceMtef);
        CopyRootLeadingFormattingRecords(
            sourceMtef,
            records,
            output,
            ReadPrefixLayout(sourceMtef).Colors,
            insertedDefinitionCount: 0,
            out _);
    }

    private static int ReadPrefixCurrentColorIndex(byte[] sourceMtef)
    {
        var color = 0;
        foreach (var record in ReadPrefixLayout(sourceMtef).Records)
        {
            if (record.Type != RecordColor) continue;
            var position = record.Start + 1;
            color = ReadUnsigned(sourceMtef, ref position);
        }
        return color;
    }

    private static IReadOnlyList<(byte Type, int Start, int End)>
        ReadRootLeadingFormattingRecords(byte[] sourceMtef)
    {
        var result = new List<(byte Type, int Start, int End)>();
        var root = FindRootStructureOffset(sourceMtef);
        if (root >= sourceMtef.Length || sourceMtef[root] == RecordEnd)
            return result;
        Require(sourceMtef, root, 2);
        var cursor = root + 1;
        var rootRecord = sourceMtef[root];
        var options = sourceMtef[cursor++];
        SkipMtefNudge(sourceMtef, ref cursor, options);
        if (rootRecord == RecordLine)
        {
            if ((options & 0x04) != 0)
            {
                Require(sourceMtef, cursor, 2);
                cursor += 2;
            }
        }
        else if (rootRecord == RecordPile)
        {
            Require(sourceMtef, cursor, 2);
            cursor += 2; // horizontal + vertical alignment
        }
        else
        {
            return result;
        }
        if ((options & 0x02) != 0)
        {
            Require(sourceMtef, cursor, 2);
            if (sourceMtef[cursor] != RecordRuler) return result;
            var rulerCount = sourceMtef[cursor + 1];
            Require(sourceMtef, cursor + 2, rulerCount * 3);
            cursor += 2 + rulerCount * 3;
        }

        while (cursor < sourceMtef.Length)
        {
            var start = cursor;
            var type = sourceMtef[cursor];
            switch (type)
            {
                case RecordEncodingDef:
                    cursor = SkipNullTerminated(sourceMtef, cursor + 1);
                    break;
                case RecordFontDef:
                    cursor++;
                    _ = ReadUnsigned(sourceMtef, ref cursor);
                    cursor = SkipNullTerminated(sourceMtef, cursor);
                    break;
                case RecordFontStyleDef:
                    cursor++;
                    _ = ReadUnsigned(sourceMtef, ref cursor);
                    Require(sourceMtef, cursor, 1);
                    cursor++;
                    break;
                case RecordColorDef:
                    cursor = SkipColorDefinition(sourceMtef, cursor);
                    break;
                case RecordColor:
                    cursor++;
                    _ = ReadUnsigned(sourceMtef, ref cursor);
                    break;
                default:
                    return result;
            }
            result.Add((type, start, cursor));
        }
        return result;
    }

    private static void SkipMtefNudge(
        byte[] data,
        ref int position,
        byte options)
    {
        if ((options & 0x08) == 0) return;
        Require(data, position, 2);
        var dx = data[position++];
        var dy = data[position++];
        if (dx == 128 && dy == 128)
        {
            Require(data, position, 4);
            position += 4;
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
            // Recognizing a source fence does not establish a native template.
            // Keep unsupported endpoint tokens on their original emission path,
            // including verified CHAR font/encoding bindings and local attributes.
            if (!SelectFenceTemplate(open, close).HasValue) continue;
            var fenced = new XElement(
                "mfenced",
                new XAttribute("open", NormalizeFence(open)),
                new XAttribute("close", NormalizeFence(close)),
                new XAttribute("separators", string.Empty),
                children.Select(child => new XElement(child)));
            row.ReplaceNodes(fenced);
        }
    }

    private static void MaterializeLooseFenceGroups(XElement root)
    {
        // OMML's reverse projection can express a native group as sibling
        // fences, with a trailing script attached to the closing fence. The
        // semantic reader already recognizes that group. Encode the same group
        // as one native fence OBJECT before the trailing-script writer binds it.
        // Explicitly fixed fence CHARs are excluded by the shared predicates.
        foreach (var container in root.DescendantsAndSelf()
                     .Where(element => element.Name.LocalName is "math" or "mrow" or "mstyle")
                     .Reverse().ToArray())
        {
            var children = container.Elements().ToArray();
            for (var index = 0; index < children.Length; index++)
            {
                if (!TryGetLooseOpeningFenceToken(children[index], out var open))
                    continue;
                var stack = new List<string> { open };
                for (var cursor = index + 1; cursor < children.Length; cursor++)
                {
                    var candidate = children[cursor];
                    var scriptChildren = candidate.Elements().ToArray();
                    var isScript = candidate.Name.LocalName is "msub" or "msup" or "msubsup";
                    var requiredSlots = candidate.Name.LocalName == "msubsup" ? 3 : 2;
                    var closingCandidate = isScript && scriptChildren.Length == requiredSlots
                        ? scriptChildren[0] : candidate;
                    if (TryGetLooseClosingFenceToken(closingCandidate, out var close))
                    {
                        if (!AreMatchingFences(stack[stack.Count - 1], close))
                            break;
                        stack.RemoveAt(stack.Count - 1);
                        if (stack.Count > 0) continue;
                        var inner = new XElement(container.Name.Namespace + "mrow",
                            children.Skip(index + 1).Take(cursor - index - 1)
                                .Select(element => new XElement(element)));
                        MaterializeLooseFenceGroups(inner);
                        var fenced = new XElement(container.Name.Namespace + "mfenced",
                            new XAttribute("open", NormalizeFence(open)),
                            new XAttribute("close", NormalizeFence(close)),
                            new XAttribute("separators", string.Empty), inner);
                        XElement replacement = fenced;
                        if (isScript)
                        {
                            replacement = new XElement(candidate);
                            replacement.Elements().First().ReplaceWith(fenced);
                        }
                        children[index].ReplaceWith(replacement);
                        foreach (var removed in children.Skip(index + 1).Take(cursor - index))
                            removed.Remove();
                        children = container.Elements().ToArray();
                        break;
                    }
                    if (TryGetLooseOpeningFenceToken(candidate, out var nestedOpen))
                        stack.Add(nestedOpen);
                }
            }
        }
    }

    private static string CanonicalizeElementSequence(
        XElement[] elements,
        string? variant)
    {
        if (elements.Length == 0)
            return string.Empty;
        if (TryCanonicalizeSplitFencedScriptSequence(
                elements,
                variant,
                out var splitFencedScript))
            return splitFencedScript;
        if (TryCanonicalizeLooseBinomialSequence(
                elements,
                variant,
                out var looseBinomialSequence))
            return looseBinomialSequence;
        if (TryCanonicalizeLooseFenceSequence(
                elements,
                variant,
                out var looseFenceSequence))
            return looseFenceSequence;
        return string.Concat(
            elements.Select(
                child => CanonicalizeMathMl(
                    child,
                    variant)));
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

    private static bool TryCanonicalizeLooseFenceSequence(
        XElement[] elements,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        if (elements.Length < 2) return false;

        var builder = new StringBuilder();
        var replaced = false;
        for (var index = 0; index < elements.Length;)
        {
            if (TryGetLooseOpeningFenceToken(
                    elements[index],
                    out var open)
                && TryFindMatchingLooseFence(
                    elements,
                    index,
                    open,
                    out var closeIndex,
                    out var close))
            {
                var innerElements =
                    elements
                        .Skip(index + 1)
                        .Take(closeIndex - index - 1)
                        .ToArray();
                innerElements =
                    StripRedundantNestedLooseFenceLayers(
                        innerElements,
                        open,
                        close);
                var innerSignature =
                    CanonicalizeElementSequence(
                        innerElements,
                        variant);

                builder.Append(
                    CanonicalFenceSignature(
                        open,
                        close,
                        innerSignature));
                index = closeIndex + 1;
                replaced = true;
                continue;
            }

            builder.Append(
                CanonicalizeMathMl(
                    elements[index],
                    variant));
            index++;
        }

        if (!replaced) return false;
        signature = builder.ToString();
        return true;
    }

    private static bool TryFindMatchingLooseFence(
        XElement[] elements,
        int openIndex,
        string open,
        out int closeIndex,
        out string close)
    {
        closeIndex = -1;
        close = string.Empty;
        var normalizedOpen = NormalizeFence(open);
        var expectedClose =
            MatchingClosingFence(
                normalizedOpen);
        if (expectedClose.Length == 0)
            return false;

        var depth = 0;
        for (var index = openIndex + 1;
             index < elements.Length;
             index++)
        {
            if (TryGetLooseOpeningFenceToken(
                    elements[index],
                    out var nestedOpen)
                && string.Equals(
                    NormalizeFence(nestedOpen),
                    normalizedOpen,
                    StringComparison.Ordinal))
            {
                depth++;
                continue;
            }

            if (!TryGetLooseClosingFenceToken(
                    elements[index],
                    out var candidateClose)
                || !string.Equals(
                    NormalizeFence(candidateClose),
                    expectedClose,
                    StringComparison.Ordinal))
                continue;

            if (depth > 0)
            {
                depth--;
                continue;
            }

            closeIndex = index;
            close = candidateClose;
            return true;
        }

        return false;
    }

    private static bool TryGetLooseOpeningFenceToken(
        XElement candidate,
        out string fence)
    {
        if (!TryGetLooseFenceToken(
                candidate,
                out fence))
            return false;

        var normalized =
            NormalizeFence(
                fence);
        if (normalized is "(" or "[" or "{" or "⟨" or "⌈" or "⌊")
            return true;

        if (normalized is "|" or "‖")
            return HasExplicitFenceRole(
                candidate,
                "OPEN");

        return false;
    }

    private static bool TryGetLooseClosingFenceToken(
        XElement candidate,
        out string fence)
    {
        if (!TryGetLooseFenceToken(
                candidate,
                out fence))
            return false;

        var normalized =
            NormalizeFence(
                fence);
        if (normalized is ")" or "]" or "}" or "⟩" or "⌉" or "⌋")
            return true;

        if (normalized is "|" or "‖")
            return HasExplicitFenceRole(
                candidate,
                "CLOSE");

        return false;
    }

    private static bool HasExplicitFenceRole(
        XElement candidate,
        string expectedRole)
    {
        XElement token = candidate;
        if (candidate.Name.LocalName == "mrow")
        {
            var children = candidate.Elements().ToArray();
            if (children.Length != 1
                || children[0].Name.LocalName != "mo")
                return false;
            token = children[0];
        }

        var candidateClass =
            ((string?)candidate.Attribute(
                "data-mjx-texclass")
             ?? string.Empty).Trim();
        var tokenClass =
            ((string?)token.Attribute(
                "data-mjx-texclass")
             ?? string.Empty).Trim();
        if (string.Equals(
                candidateClass,
                expectedRole,
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                tokenClass,
                expectedRole,
                StringComparison.OrdinalIgnoreCase))
            return true;

        return string.Equals(
                (string?)token.Attribute(
                    "fence"),
                "true",
                StringComparison.OrdinalIgnoreCase)
            && string.Equals(
                (string?)token.Attribute(
                    "stretchy"),
                "true",
                StringComparison.OrdinalIgnoreCase);
    }

    private static string MatchingClosingFence(
        string open) =>
        NormalizeFence(
            open) switch
        {
            "(" => ")",
            "[" => "]",
            "{" => "}",
            "⟨" => "⟩",
            "⌈" => "⌉",
            "⌊" => "⌋",
            "|" => "|",
            "‖" => "‖",
            _ => string.Empty,
        };

    private static XElement[] StripRedundantNestedLooseFenceLayers(
        XElement[] elements,
        string outerOpen,
        string outerClose)
    {
        var normalizedOuterOpen =
            NormalizeFence(
                outerOpen);
        var normalizedOuterClose =
            NormalizeFence(
                outerClose);
        var current =
            elements;
        while (current.Length >= 2
               && TryGetLooseOpeningFenceToken(
                    current[0],
                    out var nestedOpen)
               && TryFindMatchingLooseFence(
                    current,
                    0,
                    nestedOpen,
                    out var nestedCloseIndex,
                    out var nestedClose)
               && nestedCloseIndex ==
                    current.Length - 1
               && string.Equals(
                    NormalizeFence(
                        nestedOpen),
                    normalizedOuterOpen,
                    StringComparison.Ordinal)
               && string.Equals(
                    NormalizeFence(
                        nestedClose),
                    normalizedOuterClose,
                    StringComparison.Ordinal))
        {
            current =
                current
                    .Skip(1)
                    .Take(
                        current.Length - 2)
                    .ToArray();
        }

        return current;
    }

    private static string CanonicalFenceSignature(
        string open,
        string close,
        string innerSignature)
    {
        var normalizedOpen =
            NormalizeFence(
                open);
        var normalizedClose =
            NormalizeFence(
                close);

        if (normalizedOpen.Length > 0
            && normalizedClose.Length == 0)
        {
            return "o("
                + normalizedOpen
                + ")"
                + innerSignature;
        }

        if (normalizedOpen.Length == 0
            && normalizedClose.Length > 0)
        {
            return innerSignature
                + "o("
                + normalizedClose
                + ")";
        }

        return "fence("
            + normalizedOpen
            + ","
            + normalizedClose
            + ","
            + innerSignature
            + ")";
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

    private static bool TryCanonicalizeSplitFencedScriptSequence(
        XElement[] elements,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        if (elements.Length < 3) return false;
        if (!TryGetLooseOpeningFenceToken(
                elements[0],
                out var open))
            return false;

        for (var scriptIndex = 2;
             scriptIndex < elements.Length;
             scriptIndex++)
        {
            var script = elements[scriptIndex];
            var scriptKind = script.Name.LocalName;
            if (scriptKind is not ("msub" or "msup" or "msubsup"))
                continue;

            var scriptChildren =
                script.Elements().ToArray();
            var minimumChildren =
                scriptKind == "msubsup"
                    ? 3
                    : 2;
            if (scriptChildren.Length < minimumChildren)
                continue;
            if (!TryGetLooseClosingFenceToken(
                    scriptChildren[0],
                    out var close))
                continue;
            if (!AreMatchingFences(
                    open,
                    close))
                continue;

            var innerElements =
                elements
                    .Skip(1)
                    .Take(scriptIndex - 1)
                    .ToArray();
            innerElements =
                StripRedundantNestedLooseFenceLayers(
                    innerElements,
                    open,
                    close);
            var innerSignature =
                CanonicalizeElementSequence(
                    innerElements,
                    variant);
            if (innerSignature.Length == 0)
                continue;

            var baseSignature =
                CanonicalFenceSignature(
                    open,
                    close,
                    innerSignature);
            string scriptedSignature;
            if (scriptKind == "msub")
            {
                var subSignature =
                    CanonicalizeMathMl(
                        scriptChildren[1],
                        variant);
                if (subSignature.Length == 0)
                    continue;
                scriptedSignature =
                    "sub("
                    + baseSignature
                    + ","
                    + subSignature
                    + ")";
            }
            else if (scriptKind == "msup")
            {
                var supSignature =
                    CanonicalizeMathMl(
                        scriptChildren[1],
                        variant);
                if (supSignature.Length == 0)
                    continue;
                scriptedSignature =
                    "sup("
                    + baseSignature
                    + ","
                    + supSignature
                    + ")";
            }
            else
            {
                var subSignature =
                    CanonicalizeMathMl(
                        scriptChildren[1],
                        variant);
                var supSignature =
                    CanonicalizeMathMl(
                        scriptChildren[2],
                        variant);
                if (subSignature.Length == 0
                    && supSignature.Length == 0)
                    continue;
                if (subSignature.Length == 0)
                {
                    scriptedSignature =
                        "sup("
                        + baseSignature
                        + ","
                        + supSignature
                        + ")";
                }
                else if (supSignature.Length == 0)
                {
                    scriptedSignature =
                        "sub("
                        + baseSignature
                        + ","
                        + subSignature
                        + ")";
                }
                else
                {
                    scriptedSignature =
                        "subsup("
                        + baseSignature
                        + ","
                        + subSignature
                        + ","
                        + supSignature
                        + ")";
                }
            }

            var tail =
                CanonicalizeElementSequence(
                    elements
                        .Skip(scriptIndex + 1)
                        .ToArray(),
                    variant);
            signature =
                scriptedSignature
                + tail;
            return true;
        }

        return false;
    }

    // Presentation MathML gives explicit fixed-glyph attributes precedence over
    // TeX spacing classes. OPEN/CLOSE on \\lvert, \\rvert and \\langle describe
    // spacing, not an expandable fence template. MTEF CHAR does not retain those
    // MathJax annotations, so canonicalization and emission must agree here.
    private static bool IsExplicitlyFixedFence(XElement token) =>
        IsFalseMathMlBoolean((string?)token.Attribute("fence"))
        || IsFalseMathMlBoolean((string?)token.Attribute("stretchy"));

    // The source's attachment semantics decide decoration versus an upper
    // limit before a big-operator shortcut is selected. Overline and horizontal
    // braces have native decoration templates even with accent=false.
    private static bool IsOverDecoration(XElement mover, XElement overscript) =>
        ResolveOverAccent(mover, overscript)
        || NormalizeAccentMark(overscript.Value.Trim()) == "¯"
        || IsHorizontalBraceMarker(overscript.Value.Trim(), top: true);

    private static bool ResolveOverAccent(XElement mover, XElement overscript)
    {
        var own = (string?)mover.Attribute("accent");
        if (IsTrueMathMlBoolean(own)) return true;
        if (IsFalseMathMlBoolean(own)) return false;

        // MathML 3, 3.4.5: an omitted accent inherits the core mo's accent,
        // including its operator-dictionary default. Absence is not false.
        var core = overscript;
        while (core.Name.LocalName is "mrow" or "mstyle" or "mpadded" or "semantics")
        {
            var contents = core.Elements()
                .Where(child => child.Name.LocalName is not ("annotation" or "annotation-xml"))
                .ToArray();
            if (contents.Length != 1) return false;
            core = contents[0];
        }
        if (core.Name.LocalName != "mo") return false;
        var mark = (string?)core.Attribute("accent");
        if (IsTrueMathMlBoolean(mark)) return true;
        if (IsFalseMathMlBoolean(mark)) return false;
        return NormalizeAccentMark(core.Value.Trim()) is "→" or "←" or "↔"
            or "^" or "~" or "˙" or "¨" or "¯"
            or "⌢" or "⏜" or "ˇ" or "˘" or "´" or "`" or "˚";
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
        if (token is null || IsExplicitlyFixedFence(token)) return false;
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
        if (!AreMatchingFences(open, close)) return false;

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
        // Explicit source structure is independent of MTEF template support.
        // Unsupported templates still use EmitFenced\'s existing fallback;
        // semantic comparison must not erase a source fence to excuse that loss.
        if (MathMlExplicitFenceContract.TryGetRow(row, out open, out close, out children))
            return true;
        if (row.Name.LocalName != "mrow") return false;
        var elements = row.Elements().ToArray();
        if (elements.Length < 3) return false;
        if (!TryGetMathJaxFenceToken(elements[0], "OPEN", out open)) return false;
        if (!TryGetMathJaxFenceToken(elements[elements.Length - 1], "CLOSE", out close)) return false;
        // MathJax represents \left\{ ... \right. with an explicit empty
        // closing fence. It is a one-sided delimiter template, not an empty
        // operator after an ordinary brace glyph. Use the same recognition for
        // MTEF geometry and semantic comparison.
        var oneSided = (open.Length == 0) != (close.Length == 0);
        if (!AreMatchingFences(open, close)
            && !(oneSided && SelectFenceTemplate(open, close).HasValue)) return false;
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
        fence = NormalizeFence(token.Value.Trim());
        // This rule must run before OPEN/CLOSE spacing-class recognition and
        // must be identical to the loose-sequence canonicalizer's rule.
        if (IsExplicitlyFixedFence(token)) return false;
        if (markerMatches) return true;
        if (fence.Length == 0
            && string.Equals((string?)token.Attribute("fence"), "true", StringComparison.OrdinalIgnoreCase))
            return true;

        // Word's OMML->MathML transform emits structurally complete delimiter
        // rows without MathJax's data-mjx-texclass markers. MathType 7 otherwise
        // imports only the first opening delimiter into a big-operator main slot
        // and leaves the remaining operand as root siblings. Accept an unmarked
        // token only when its glyph unambiguously matches the requested side;
        // the caller also verifies that the two delimiters form a valid pair.
        return string.Equals(expectedClass, "OPEN", StringComparison.OrdinalIgnoreCase)
            ? fence is "(" or "[" or "{" or "⟨" or "⌈" or "⌊" or "|" or "‖"
            : string.Equals(expectedClass, "CLOSE", StringComparison.OrdinalIgnoreCase)
                && fence is ")" or "]" or "}" or "⟩" or "⌉" or "⌋" or "|" or "‖";
    }

    // Exact native glyph contracts, shared by font binding, CHAR emission and
    // decoding. Do not infer a glyph from a font's Unicode coverage. The corner
    // positions are Design Science's eumat1 AFM/post names, cross-checked with
    // lcdf-typetools texglyphlist and native WMF pixels. E92F is MathType's own
    // shortparallel MTCode; its EuclidMath1 position 89 renders the short bars.
    private sealed class NativeGlyphAuthority
    {
        internal NativeGlyphAuthority(int scalar, bool alternate, string font,
            string encoding, int mtCode, byte encoded8, string tokenClass, string glyphName)
        {
            Scalar = scalar; Alternate = alternate; Font = font; Encoding = encoding;
            MtCode = mtCode; Encoded8 = encoded8; TokenClass = tokenClass; GlyphName = glyphName;
        }
        internal int Scalar { get; }
        internal bool Alternate { get; }
        internal string Font { get; }
        internal string Encoding { get; }
        internal int MtCode { get; }
        internal byte Encoded8 { get; }
        internal string TokenClass { get; }
        internal string GlyphName { get; }
    }

    private static void EmitMathSpace(XElement element, List<byte> output)
    {
        var space = MathSpaceContract.ParseWidth((string?)element.Attribute("width")).NativeSpace;
        if (space is null)
        {
            // Preserve compatibility for unproven widths; this is not a claim
            // that ASCII space faithfully represents an arbitrary MathML gap.
            EmitCharacter(' ', TypefaceSpace, output);
            return;
        }
        // An edited width is authoritative over stale native CHAR metadata.
        // For an unchanged known space, preserve its original encoding flags.
        if ((int?)element.Attribute("data-mtef-explicit-typeface") == TypefaceSpace
            && (int?)element.Attribute("data-mtef-mtcode") == space.MtCode
            && !IsTrueMathMlBoolean((string?)element.Attribute("data-mtef-no-mtcode"))
            && TryEmitPreservedCharacter(element, TypefaceSpace, output))
            return;
        EmitScalar(space.MtCode, TypefaceSpace, output);
    }

    private static readonly NativeGlyphAuthority[] NativeGlyphAuthorities =
    {
        // MT Extra's planckover2pi slot is the TeX hbar glyph. Bind it through
        // the same native font/style path as other alternate glyphs; selecting
        // the built-in Extra face directly discards the identifier's style.
        new(0x210F, true, "MT Extra", "MTExtra", 0x210F, 0x68, "mi", "planckover2pi"),
        new(0x231C, false, "Euclid Math One", "EuclidMath1", 0x231C, 0x84, "mo", "rightanglenw"),
        new(0x231D, false, "Euclid Math One", "EuclidMath1", 0x231D, 0x86, "mo", "rightanglene"),
        new(0x231E, false, "Euclid Math One", "EuclidMath1", 0x231E, 0x85, "mo", "rightanglesw"),
        new(0x231F, false, "Euclid Math One", "EuclidMath1", 0x231F, 0x87, "mo", "rightanglese"),
        new(0x2225, true, "Euclid Math One", "EuclidMath1", 0xE92F, 0x89, "mo", "parallelvar"),
        new(0x2226, false, "Euclid Math One", "EuclidMath1", 0x2226, 0x8D, "mo", "notparallelto"),
        new(0x2226, true, "Euclid Math One", "EuclidMath1", 0xEA2F, 0x8B, "mo", "notparalleltovar"),
        // Vendor Texvc input + Euclid AFM/post/cmap authority; alternate slots
        // remain distinct from their ordinary Unicode counterparts.
        new(0x2223, true, "Euclid Math One", "EuclidMath1", 0xE92E, 0x88, "mo", "dividesvar"),
        new(0x2224, false, "Euclid Math One", "EuclidMath1", 0x2224, 0x8C, "mo", "doesnotdivide"),
        new(0x2224, true, "Euclid Math One", "EuclidMath1", 0xEA2E, 0x8A, "mo", "doesnotdividevar"),
        new(0x2216, true, "Euclid Math One", "EuclidMath1", 0x2216, 0x80, "mo", "integerdivide"),
        new(0x228A, false, "Euclid Math Two", "EuclidMath2", 0x228A, 0xD6, "mo", "subsetnotequal"),
        new(0x228A, true, "Euclid Math Two", "EuclidMath2", 0xEA42, 0xD8, "mo", "notsubsetoreql"),
        new(0x228B, false, "Euclid Math Two", "EuclidMath2", 0x228B, 0xD7, "mo", "supersetnotequal"),
        new(0x228B, true, "Euclid Math Two", "EuclidMath2", 0xEA43, 0xD9, "mo", "notsupersetoreql"),
        new(0x2288, false, "Euclid Math Two", "EuclidMath2", 0x2288, 0xDA, "mo", "notreflexsubset"),
        new(0x2288, true, "Euclid Math Two", "EuclidMath2", 0xEA0C, 0xE0, "mo", "notreflexsubsetdbl"),
        new(0x2289, false, "Euclid Math Two", "EuclidMath2", 0x2289, 0xDB, "mo", "notreflexsuperset"),
        new(0x2289, true, "Euclid Math Two", "EuclidMath2", 0xEA0B, 0xE1, "mo", "notreflexsupersetdbl"),
        new(0x2ACB, false, "Euclid Math Two", "EuclidMath2", 0xEA44, 0xDC, "mo", "subsetornotdbleql"),
        new(0x2ACB, true, "Euclid Math Two", "EuclidMath2", 0xEA46, 0xDE, "mo", "notsubsetordblequal"),
        new(0x2ACC, false, "Euclid Math Two", "EuclidMath2", 0xEA45, 0xDD, "mo", "supersetornotdbleql"),
        new(0x2ACC, true, "Euclid Math Two", "EuclidMath2", 0xEA47, 0xDF, "mo", "notsupersetordblequal"),
        new(0x2220, false, "Symbol", "Symbol", 0x2220, 0xD0, "mi", "angle"),
        new(0x2221, false, "MT Extra", "MTExtra", 0x2221, 0x52, "mi", "measuredangle"),
        new(0x2222, false, "MT Extra", "MTExtra", 0x2222, 0x53, "mi", "sphericalangle"),
    };

    private static NativeGlyphAuthority? ResolveNativeGlyphAuthority(XElement token)
    {
        if (token.Name.LocalName is not ("mo" or "mi") || token.Value.Length != 1)
            return null;
        var alternate = (string?)token.Attribute("data-mjx-alternate") == "1";
        return NativeGlyphAuthorities.FirstOrDefault(glyph =>
            glyph.Scalar == token.Value[0] && glyph.Alternate == alternate);
    }

    private static NativeGlyphAuthority? ResolveNativeGlyphAuthority(
        string? font, string? encoding, int scalar, int encoded, bool hasMtCode) =>
        NativeGlyphAuthorities.FirstOrDefault(glyph =>
            string.Equals(font, glyph.Font, StringComparison.OrdinalIgnoreCase)
            && string.Equals(encoding, glyph.Encoding, StringComparison.OrdinalIgnoreCase)
            && (encoded < 0 || encoded == glyph.Encoded8)
            && (!hasMtCode || scalar == glyph.MtCode
                || scalar == glyph.Scalar && encoded == glyph.Encoded8));

    private static string? NativeGlyphCharacterVariant(byte style, NativeGlyphAuthority glyph) =>
        style switch
        { 1 => "bold", 2 => "italic", 3 => "bold-italic", _ => glyph.TokenClass == "mi" ? "normal" : null };

    private static void EmitExplicitFontDefinitions(
        XElement math,
        byte[] sourceMtef,
        List<byte> output)
    {
        var definitionsStart = output.Count;
        var existing = CountAvailableDefinitions(sourceMtef);
        var nativeStyles = ReadNativeCharacterStyleBindings(sourceMtef);
        // MTEF persists scalar CHARs. Split only the already-canonical equivalent
        // operator runs before binding fonts, so every glyph receives its style.
        foreach (var token in math.DescendantsAndSelf().Where(e => e.Name.LocalName == "mo"
                     && IsScalarOperatorRun(e.Value)).ToArray())
            token.ReplaceWith(EnumerateBmpScalars(token.Value).Select(scalar =>
                new XElement(token.Name, token.Attributes(), char.ConvertFromUtf32(scalar))));
        // An operator outside Times New Roman's code pages (e.g. U+25B3) has no
        // styled identifier form; bind it to SimSun/WinAllCodePages below.
        foreach (var token in math.DescendantsAndSelf().Where(e => e.Value.Trim().Length == 1
                     && (e.Name.LocalName == "mo"
                         || e.Name.LocalName == "mi" && IsMathematicalSymbolScalar(e.Value.Trim()[0]))
                     && IsFarEastCodePageOperator(e.Value.Trim()[0])
                     && MathMlTokenStyleContract.ResolveDeclaredVariant(e) is null
                     && ResolveNativeGlyphAuthority(e) is null
                     && !HasValidPreservedExplicitTypeface(e, existing.FontStyleDefinitions, nativeStyles)).ToArray())
        {
            // A symbol mi keeps its implicit italic, as SemanticSignature does.
            token.SetAttributeValue(FarEastOperatorMarker, token.Name.LocalName == "mi" ? "italic" : "normal");
            token.Name = token.Name.Namespace + "mo";
            token.Attribute("mathvariant")?.Remove();
        }
        // A single mathematical-symbol mi inherits italic when no variant was
        // supplied. Inherited variants have already been materialized. Resolve
        // this before changing token class or binding a native font, so a plain
        // native glyph cannot masquerade as an implicit-italic MathML token.
        foreach (var token in math.DescendantsAndSelf().Where(e => e.Name.LocalName == "mi"
                     && e.Value.Trim().Length == 1
                     && (IsMathematicalSymbolToken(e.Value.Trim())
                         || ResolveNativeGlyphAuthority(e) is not null
                         || IsLowerGreek(e.Value.Trim()[0]) || IsUpperGreek(e.Value.Trim()[0]))
                     && MathMlTokenStyleContract.ResolveDeclaredVariant(e) is null))
        {
            if (TryNormalizeLetterlikeForWrite(token.Value.Trim()[0], out _, out _)) continue;
            var effective = MathMlTokenStyleContract.ResolveEffectiveVariant(token);
            var scalar = token.Value.Trim()[0];
            var greekFace = IsLowerGreek(scalar) ? TypefaceLowerGreek
                : IsUpperGreek(scalar) ? TypefaceUpperGreek : 0;
            if (greekFace != 0 && nativeStyles.TryGetValue(greekFace, out var binding)
                && TryNativeCharacterStyle(effective, out var requestedBits)
                && requestedBits == binding.CharacterStyle)
                continue;
            token.SetAttributeValue("mathvariant", effective);
        }
        foreach (var token in math.DescendantsAndSelf())
        {
            if (!int.TryParse((string?)token.Attribute("data-mtef-explicit-typeface"), out var face)
                || !nativeStyles.TryGetValue(face, out var actual)) continue;
            var requested = MathMlTokenStyleContract.ResolveEffectiveVariant(token);
            if (TryNativeCharacterStyle(requested, out var bits) && bits != actual.CharacterStyle)
            {
                // A new font/style must not inherit an old font-position byte.
                foreach (var attribute in token.Attributes().Where(a =>
                             a.Name.LocalName.StartsWith("data-mtef-", StringComparison.Ordinal)).ToArray())
                    attribute.Remove();
            }
        }
        foreach (var token in math.DescendantsAndSelf().Where(e => e.Name.LocalName == "mi"
                     && ResolveNativeGlyphAuthority(e) is not null
                     && !HasValidPreservedExplicitTypeface(e, existing.FontStyleDefinitions, nativeStyles)))
            token.Name = token.Name.Namespace + ResolveNativeGlyphAuthority(token)!.TokenClass;
        // A mathematical glyph has the same native character mapping whether
        // a frontend calls it mi or mo. Keep explicit styles on that mapping.
        foreach (var token in math.DescendantsAndSelf().Where(e => e.Name.LocalName == "mi"
                     && IsMathematicalSymbolToken(e.Value.Trim())
                     && ResolveNativeGlyphAuthority(e)?.TokenClass != "mi"
                     && ((string?)e.Attribute("mathvariant") ?? string.Empty) is "bold" or "italic" or "bold-italic"))
            token.Name = token.Name.Namespace + "mo";
        var specs = new[]
        {
            (Variant: "normal", Encoding: "WindowsANSI", Font: "Times New Roman", CharacterStyle: (byte)0),
            (Variant: "bold", Encoding: "WindowsANSI", Font: "Times New Roman", CharacterStyle: (byte)1),
            (Variant: "italic", Encoding: "WindowsANSI", Font: "Times New Roman", CharacterStyle: (byte)2),
            (Variant: "script", Encoding: "EuclidMath1", Font: "Euclid Math One", CharacterStyle: (byte)0),
            (Variant: "bold-script", Encoding: "EuclidMath1", Font: "Euclid Math One", CharacterStyle: (byte)1),
            (Variant: "double-struck", Encoding: "EuclidMath2", Font: "Euclid Math Two", CharacterStyle: (byte)0),
            (Variant: "fraktur", Encoding: "EuclidFraktur", Font: "Euclid Fraktur", CharacterStyle: (byte)0),
            (Variant: "bold-fraktur", Encoding: "EuclidFraktur", Font: "Euclid Fraktur", CharacterStyle: (byte)1),
            (Variant: "sans-serif", Encoding: "WindowsANSI", Font: "Arial", CharacterStyle: (byte)0),
            (Variant: "bold-sans-serif", Encoding: "WindowsANSI", Font: "Arial", CharacterStyle: (byte)1),
            (Variant: "sans-serif-italic", Encoding: "WindowsANSI", Font: "Arial", CharacterStyle: (byte)2),
            (Variant: "sans-serif-bold-italic", Encoding: "WindowsANSI", Font: "Arial", CharacterStyle: (byte)3),
            (Variant: "monospace", Encoding: "WindowsANSI", Font: "Courier New", CharacterStyle: (byte)0),
            (Variant: "bold-italic", Encoding: "WindowsANSI", Font: "Times New Roman", CharacterStyle: (byte)3),
            (Variant: "native-glyph", Encoding: "Symbol", Font: "Symbol", CharacterStyle: (byte)0),
            (Variant: "native-glyph", Encoding: "MTExtra", Font: "MT Extra", CharacterStyle: (byte)0),
            // Font is the prefix's own typeface-12 FONT_DEF (SimSun under its
            // localized name); MathType rejects the English name "SimSun".
            (Variant: "far-east-operator", Encoding: "WinAllCodePages", Font: "", CharacterStyle: (byte)0),
            (Variant: "far-east-operator-italic", Encoding: "WinAllCodePages", Font: "", CharacterStyle: (byte)2),
        };
        var createdEncodings = 0;
        var createdFonts = 0;
        var createdStyles = 0;
        foreach (var spec in specs)
        {
            var tokens = math.DescendantsAndSelf()
                .Where(element => ResolveNativeGlyphAuthority(element) is { } glyph
                    ? spec.CharacterStyle == 0 && glyph.Font == spec.Font && glyph.Encoding == spec.Encoding
                    : MatchesExplicitFontSpec(element, spec.Variant))
                .Where(element => !HasValidPreservedExplicitTypeface(
                    element,
                    existing.FontStyleDefinitions, nativeStyles))
                .ToArray();
            if (tokens.Length == 0) continue;

            if (spec.Variant.StartsWith("far-east-operator", StringComparison.Ordinal))
            {
                // Without a typeface-12 font the operator keeps its old Times binding.
                if (!nativeStyles.TryGetValue(TypefaceFarEastText, out var farEast) || farEast.FontDefinition <= 0)
                    continue;
                var farEastStyle = FindPrefixFontStyleIndex(sourceMtef, farEast.FontDefinition, spec.CharacterStyle);
                if (farEastStyle is null)
                {
                    farEastStyle = existing.FontStyleDefinitions + ++createdStyles;
                    output.Add(RecordFontStyleDef);
                    WriteUnsigned(output, farEast.FontDefinition);
                    output.Add(spec.CharacterStyle);
                }
                foreach (var token in tokens)
                    token.SetAttributeValue("data-mtef-explicit-typeface", -farEastStyle.Value);
                continue;
            }

            var existingStyleIndex = FindPrefixFontStyleIndex(
                sourceMtef,
                spec.Font,
                spec.CharacterStyle);
            if (existingStyleIndex is not null)
            {
                foreach (var token in tokens)
                {
                    token.SetAttributeValue("data-mtef-explicit-typeface", -existingStyleIndex.Value);
                    token.SetAttributeValue("data-mtef-explicit-font", spec.Font);
                }
                continue;
            }

            var encodingIndex = 4 + existing.EncodingDefinitions + ++createdEncodings;
            var fontDefinitionIndex = existing.FontDefinitions + ++createdFonts;
            var fontStyleIndex = existing.FontStyleDefinitions + ++createdStyles;

            output.Add(RecordEncodingDef);
            WriteNullTerminatedAscii(output, spec.Encoding);
            output.Add(RecordFontDef);
            WriteUnsigned(output, encodingIndex);
            WriteNullTerminatedAscii(output, spec.Font);
            output.Add(RecordFontStyleDef);
            WriteUnsigned(output, fontDefinitionIndex);
            output.Add(spec.CharacterStyle);

            foreach (var token in tokens)
            {
                token.SetAttributeValue("data-mtef-explicit-typeface", -fontStyleIndex);
                token.SetAttributeValue("data-mtef-explicit-font", spec.Font);
            }
        }
        MaterializeExplicitCharacterStyles(math, sourceMtef, output, definitionsStart);
    }

    // Resolve font/style through the native glyph's actual encoding. A bold
    // symbol is NOT an identifier in the Vector face, nor a Unicode scalar in
    // an arbitrary Times font. Retain the canonical CHAR's MTCode/font-position
    // and bind a FONT_STYLE_DEF to the SAME FONT_DEF with the requested bits.
    // This applies to Symbol, MT Extra, Euclid and text fallback characters.
    private static void MaterializeExplicitCharacterStyles(
        XElement math, byte[] sourceMtef, List<byte> output, int definitionsStart)
    {
        // FUNCTION APPLICATION (U+2061) is stored as no CHAR at all, so a style
        // on it has nothing to bind to and no visible effect.
        var tokens = math.DescendantsAndSelf().Where(e => (e.Name.LocalName is "mo" or "mi")
            && e.Value.Length == 1
            && e.Value[0] != (char)0x2061
            && (((string?)e.Attribute("mathvariant") ?? string.Empty) is "bold" or "italic" or "bold-italic"
                || TryGetBasicStyledLegacyIdentifierGlyph(e, out _, out _, out _))).ToArray();
        if (tokens.Length == 0) return;
        var fonts = new List<string>();
        var styles = new List<(int Font, byte Bits)>();
        var builtin = ReadNativeCharacterStyleBindings(sourceMtef)
            .Where(item => item.Key > 0).ToDictionary(item => item.Key, item => item.Value.FontDefinition);
        void Definition(byte[] data, ref int position)
        {
            var record = data[position++];
            if (record == RecordEncodingDef) position = SkipNullTerminated(data, position);
            else if (record == RecordFontDef)
            {
                _ = ReadUnsigned(data, ref position);
                fonts.Add(ReadPrefixNullTerminatedString(data, ref position));
            }
            else if (record == RecordFontStyleDef)
            {
                var index = ReadUnsigned(data, ref position);
                Require(data,position,1); styles.Add((index,data[position++]));
            }
            else throw new InvalidDataException("Unexpected record in native character font definitions.");
        }
        foreach(var record in ReadPrefixLayout(sourceMtef).Records.Concat(ReadRootLeadingFormattingRecords(sourceMtef)))
        {
            var cursor=record.Start;
            if(record.Type is RecordEncodingDef or RecordFontDef or RecordFontStyleDef)
                Definition(sourceMtef,ref cursor);

        }
        var added=output.Skip(definitionsStart).ToArray();
        for(var cursor=0;cursor<added.Length;) Definition(added,ref cursor);
        int StyledTypeface(int typeface, byte bits)
        {
            var font=0;
            if(typeface<0 && -typeface<=styles.Count) font=styles[-typeface-1].Font;
            else builtin.TryGetValue(typeface,out font);
            if(font==0)
            {
                var family=typeface is TypefaceSymbol or TypefaceLowerGreek or TypefaceUpperGreek ? "Symbol"
                    : typeface==TypefaceMtExtra ? "MT Extra" : "Times New Roman";
                font=fonts.FindIndex(name=>string.Equals(name,family,StringComparison.OrdinalIgnoreCase))+1;
            }
            if(font<1 || font>fonts.Count) throw new InvalidDataException("No declared font owns the styled native operator.");
            var style=styles.FindIndex(s=>s.Font==font && s.Bits==bits)+1;
            if(style==0)
            {
                output.Add(RecordFontStyleDef); WriteUnsigned(output,font); output.Add(bits);
                styles.Add((font,bits)); style=styles.Count;
            }
            return -style;
        }
        foreach(var token in tokens)
        {
            var variant=MathMlTokenStyleContract.ResolveEffectiveVariant(token)!;
            var bits=(byte)(variant=="bold" ? 1 : variant=="italic" ? 2 : 3);
            // A decoded native CHAR already owns its exact explicit style.
            // Preserve its encoded8/encoded16/no-MTCode contract byte-for-byte.
            if (HasValidPreservedExplicitTypeface(token, styles.Count, ReadNativeCharacterStyleBindings(sourceMtef))
                && int.TryParse((string?)token.Attribute("data-mtef-explicit-typeface"), out var retained)
                && (retained > 0 || styles[-retained - 1].Bits == bits))
                continue;
            var character=new List<byte>();
            // Bind style to the exact CHAR selected by this token's encoder.
            // Identifier glyphs must not bypass explicit style merely because
            // their native glyph is supplied by a legacy font map.
            if (TryGetBasicStyledLegacyIdentifierGlyph(token, out var mappedCode, out var mappedFace, out var mappedPosition))
                EmitScalar(mappedCode, mappedFace, character, includeEncoded8: true, encoded8Override: mappedPosition);
            else if (token.Name.LocalName == "mi") EmitIdentifier(token,character);
            else EmitOperator(token,character);
            if(character.Count<5 || character[0]!=RecordChar || (character[1] & 0x29)!=0)
                throw new InvalidDataException("Styled native operator must resolve to one undecorated CHAR.");
            var cursor=2;
            var lead=character[cursor++];
            var typeface=lead==255 ? (character[cursor++] | character[cursor++]<<8)-32768 : lead-128;
            var scalar=character[cursor++] | character[cursor++]<<8;
            byte? encoded=(character[1]&CharEncoded8)!=0 ? character[cursor++] : null;
            if(cursor!=character.Count) throw new InvalidDataException("Styled operator CHAR has unsupported trailing data.");
            var styled=StyledTypeface(typeface,bits);
            token.SetAttributeValue("data-mtef-explicit-typeface",styled);
            token.SetAttributeValue("data-mtef-mtcode",scalar);
            token.SetAttributeValue("data-mtef-encoded8",encoded);
            token.SetAttributeValue("data-mtef-explicit-font",fonts[styles[-styled-1].Font-1]);
            // A large-operator template can use a different encoded glyph than
            // its fixed CHAR form. Bind the template glyph to its own native
            // font instead of reusing the ordinary character's font position.
            if(TryResolveNativeBigOperator(token,out var big,allowAmbiguousUnionIntersection:true))
                token.SetAttributeValue("data-mtef-template-typeface",StyledTypeface(big.Typeface,bits));
            if(TryResolveNativeIntegral(token,out _))
            {
                token.SetAttributeValue("data-mtef-template-typeface",StyledTypeface(TypefaceSymbol,bits));
                token.SetAttributeValue("data-mtef-adornment-typeface",StyledTypeface(TypefaceMtExtra,bits));
            }
        }
    }

    private static string NativeCharacterVariant(byte style) => style switch
    { 1 => "bold", 2 => "italic", 3 => "bold-italic", _ => "normal" };

    private static string NativeTextFontVariant(string font, byte style) =>
        string.Equals(font, "Courier New", StringComparison.OrdinalIgnoreCase) ? "monospace"
        : string.Equals(font, "Arial", StringComparison.OrdinalIgnoreCase) ? style switch
        { 1 => "bold-sans-serif", 2 => "sans-serif-italic", 3 => "sans-serif-bold-italic", _ => "sans-serif" }
        : NativeCharacterVariant(style);

    private static bool TryNativeCharacterStyle(string? variant, out byte style)
    {
        style = variant switch
        { "normal" or "upright" => 0, "bold" => 1, "italic" => 2, "bold-italic" => 3, _ => byte.MaxValue };
        return style != byte.MaxValue;
    }

    private static Dictionary<int, (int FontDefinition, byte CharacterStyle)> ReadNativeCharacterStyleBindings(byte[] data)
    {
        var result = new Dictionary<int, (int FontDefinition, byte CharacterStyle)>();
        var explicitStyle = 0;
        foreach (var record in ReadPrefixLayout(data).Records.Concat(ReadRootLeadingFormattingRecords(data)))
        {
            var cursor = record.Start + 1;
            if (record.Type == RecordFontStyleDef)
            {
                var font = ReadUnsigned(data, ref cursor);
                Require(data, cursor, 1);
                result[-++explicitStyle] = (font, data[cursor]);
            }
            else if (record.Type == RecordEqnPrefs)
            {
                cursor++;
                cursor = SkipDimensionArray(data, cursor);
                cursor = SkipDimensionArray(data, cursor);
                Require(data, cursor, 1);
                var count = data[cursor++];
                for (var face = 1; face <= count; face++)
                {
                    var font = ReadUnsigned(data, ref cursor);
                    if (font == 0) continue;
                    Require(data, cursor, 1);
                    result[face] = (font, data[cursor++]);
                }
            }
        }
        return result;
    }

    private static bool HasValidPreservedExplicitTypeface(
        XElement element,
        int availableFontStyleDefinitions,
        IReadOnlyDictionary<int, (int FontDefinition, byte CharacterStyle)> nativeStyles)
    {
        return int.TryParse(
                (string?)element.Attribute("data-mtef-explicit-typeface"),
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out var typeface)
            && (typeface < 0 && -typeface <= availableFontStyleDefinitions
                || typeface > 0 && nativeStyles.TryGetValue(typeface, out var nativeStyle)
                    && TryNativeCharacterStyle(MathMlTokenStyleContract.ResolveEffectiveVariant(element), out var bits)
                    && bits == nativeStyle.CharacterStyle)
            && (int.TryParse(
                (string?)element.Attribute("data-mtef-mtcode"),
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out _)
                || IsTrueMathMlBoolean((string?)element.Attribute("data-mtef-no-mtcode")));
    }

    private static bool TryGetBasicStyledLegacyIdentifierGlyph(
        XElement token, out int mtCode, out int typeface, out byte encoded8)
    {
        mtCode = typeface = 0; encoded8 = 0;
        if (token.Name.LocalName != "mi" || token.HasElements || token.Value.Length != 1
            || MathMlTokenStyleContract.ResolveEffectiveVariant(token) is not ("bold" or "italic" or "bold-italic")
            || (string?)token.Attribute("data-mjx-alternate") == "1"
            || ResolveNativeGlyphAuthority(token) is not null)
            return false;
        var scalar = token.Value[0];
        // Alphabet-specific Greek and explicit alternate glyphs keep their
        // established authorities. This is the existing native symbol tuple,
        // selected before adding weight/slant, not a Unicode-font fallback.
        if (TryEuclidMathOneGreekVariantEncoded8(scalar, out _)
            || !(scalar > 0x7F || IsMathematicalSymbolScalar(scalar)))
            return false;
        return TryResolveLegacyMathTypeGlyph(scalar, out mtCode, out typeface, out encoded8)
            && typeface is TypefaceSymbol or TypefaceMtExtra;
    }

    private static bool MatchesExplicitFontSpec(XElement element, string variant)
    {
        if (variant.StartsWith("far-east-operator", StringComparison.Ordinal)
            || element.Attribute(FarEastOperatorMarker) is not null)
            return (string?)element.Attribute(FarEastOperatorMarker) switch
            {
                "normal" => variant == "far-east-operator",
                "italic" => variant == "far-east-operator-italic",
                _ => false,
            };
        if (TryGetBasicStyledLegacyIdentifierGlyph(element, out _, out _, out _)) return false;
        var local = element.Name.LocalName;
        var scalars = EnumerateBmpScalars(element.Value).ToArray();
        if (scalars.Length == 0) return false;
        if (local == "mo" && scalars.Length > 1 && scalars.All(IsLetterScalar)
            && variant is "bold" or "italic" or "bold-italic")
            return MathMlTokenStyleContract.ResolveEffectiveVariant(element) == variant;

        if (variant == "normal")
        {
            // Explicit upright alphabetic identifiers need their own font style:
            // the built-in Text face reads back as prose and merges neighboring
            // identifiers across math spaces. Greek faces also inherit slant
            // from preferences. Bind Latin/Greek identifiers to the existing
            // Times New Roman contract with style bits zero; keep prose, mixed
            // literals, default Greek and symbol/operator mappings unchanged.
            return local == "mi"
                && string.Equals(((string?)element.Attribute("mathvariant"))?.Trim(),
                    "normal", StringComparison.OrdinalIgnoreCase)
                && scalars.All(scalar => scalar <= 0x7F && IsLetterScalar(scalar)
                    || IsLowerGreek(scalar) || IsUpperGreek(scalar)
                    || TryEuclidMathOneGreekVariantEncoded8(scalar, out _));
        }
        if (local == "mn")
            return string.Equals(((string?)element.Attribute("mathvariant"))?.Trim(), variant, StringComparison.OrdinalIgnoreCase);
        // Built-in Greek faces inherit equation preferences: UpperGreek is
        // commonly upright, so it cannot encode an explicit italic request.
        // Use the same explicit Times font/style contract as upright Greek.
        // Leave unstyled Greek and ordinary Latin identifier preferences alone.
        if (variant is "bold" or "italic")
            return variant == "italic" && local == "mi"
                && string.Equals(((string?)element.Attribute("mathvariant"))?.Trim(),
                    variant, StringComparison.OrdinalIgnoreCase)
                && scalars.All(scalar => IsLowerGreek(scalar) || IsUpperGreek(scalar));
        if (local == "mi")
        {
            if (string.Equals(
                    ((string?)element.Attribute("mathvariant"))?.Trim(),
                    variant,
                    StringComparison.OrdinalIgnoreCase))
            {
                // The standard blackboard set letters use MathType's built-in
                // MT Extra style, so they do not require an explicit EuclidMath2
                // prefix even when MathML calls them double-struck.
                return !string.Equals(variant, "double-struck", StringComparison.OrdinalIgnoreCase)
                    || !scalars.All(scalar => TryMtExtraDoubleStruck(scalar, out _, out _));
            }

            if (string.Equals(variant, "script", StringComparison.OrdinalIgnoreCase)
                && (string.IsNullOrWhiteSpace((string?)element.Attribute("mathvariant"))
                    || string.Equals((string?)element.Attribute("mathvariant"), "italic", StringComparison.OrdinalIgnoreCase))
                && (scalars.All(scalar => TryEuclidMathOneGreekVariantEncoded8(scalar, out _))
                    || scalars.All(scalar => TryEuclidMathOneDotlessJ(scalar, out _, out _))))
            {
                // MathType 7 stores TeX \epsilon / \varrho / \varkappa and
                // dotless-j in Euclid Math One even when MathML does not carry
                // a script mathvariant.
                return true;
            }
        }

        if (local == "mo")
        {
            if (string.Equals(variant, "script", StringComparison.OrdinalIgnoreCase)
                && scalars.All(scalar => TryEuclidMathOneOperatorEncoded8(scalar, out _)))
            {
                return true;
            }
            if (string.Equals(variant, "double-struck", StringComparison.OrdinalIgnoreCase)
                && scalars.All(scalar => TryEuclidMathTwoOperatorEncoded8(scalar, out _)))
            {
                return true;
            }
        }

        return false;
    }

    private static int? FindPrefixFontStyleIndex(
        byte[] mtef, int fontDefinition, byte expectedCharacterStyle)
    {
        var styleIndex = 0;
        foreach (var record in ReadPrefixLayout(mtef).Records.Where(r => r.Type == RecordFontStyleDef))
        {
            var position = record.Start + 1;
            styleIndex++;
            if (ReadUnsigned(mtef, ref position) == fontDefinition && mtef[position] == expectedCharacterStyle)
                return styleIndex;
        }
        return null;
    }

    private static int? FindPrefixFontStyleIndex(
        byte[] mtef, string fontName, byte expectedCharacterStyle)
    {
        var fontDefinitions = new List<string>();
        var styleIndex = 0;
        foreach (var record in ReadPrefixLayout(mtef).Records)
        {
            var position = record.Start + 1;
            if (record.Type == RecordFontDef)
            {
                _ = ReadUnsigned(mtef, ref position);
                fontDefinitions.Add(ReadPrefixNullTerminatedString(mtef, ref position));
            }
            else if (record.Type == RecordFontStyleDef)
            {
                var fontDefinitionIndex = ReadUnsigned(mtef, ref position);
                var actualCharacterStyle = mtef[position];
                styleIndex++;
                if (actualCharacterStyle == expectedCharacterStyle
                    && fontDefinitionIndex > 0 && fontDefinitionIndex <= fontDefinitions.Count
                    && string.Equals(fontDefinitions[fontDefinitionIndex - 1], fontName, StringComparison.OrdinalIgnoreCase))
                    return styleIndex;
            }
        }
        return null;
    }

    private static string ReadPrefixNullTerminatedString(
        byte[] data,
        ref int position)
    {
        var start = position;
        while (position < data.Length && data[position] != 0) position++;
        if (position >= data.Length)
            throw new EndOfStreamException("MTEF prefix string is not null terminated.");
        var value = System.Text.Encoding.Default.GetString(
            data,
            start,
            position - start);
        position++;
        return value;
    }

    private static (int EncodingDefinitions, int FontDefinitions, int FontStyleDefinitions)
        CountAvailableDefinitions(byte[] mtef)
    {
        var records = ReadPrefixLayout(mtef).Records
            .Concat(ReadRootLeadingFormattingRecords(mtef))
            .ToArray();
        return (records.Count(item => item.Type == RecordEncodingDef),
            records.Count(item => item.Type == RecordFontDef),
            records.Count(item => item.Type == RecordFontStyleDef));
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

    private static void EmitNode(XNode node, MtefWriter output)
    {
        if (node is XElement element
            && MathMlColorContract.Explicit(element) is { } color)
        {
            using (output.EnterExplicitColor(color))
                EmitNodeCore(node, output);
            return;
        }
        EmitNodeCore(node, output);
    }

    private static void EmitNodeCore(XNode node, MtefWriter output)
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
                if (element.Attribute("data-mtef-explicit-typeface") is not null)
                    EmitIdentifier(element, output);
                else EmitText(element.Value, TypefaceNumber, output);
                break;
            case "mo":
            {
                var operatorValue = element.Value.Trim();
                if (TryResolveNativeIntegral(element, out var integral))
                    EmitIntegralTemplate(integral, main: null, lower: null, upper: null, output);
                else if (TryResolveNativeBigOperator(element, out var standaloneBigOperator))
                    EmitBigOperatorTemplate(standaloneBigOperator, main: null, lower: null, upper: null, output);
                else if (IsNamedLimitOperator(element)
                    || operatorValue.Length > 1 && operatorValue.All(char.IsLetter))
                    // MathJax emits alphabetic operators such as TeX `\bmod` as
                    // one upright <mo> token. Writing each ASCII letter through
                    // the Symbol typeface made the MTEF reader recover three
                    // unrelated operator tokens (`m`, `o`, `d`), tripping the
                    // strict standalone semantic gate. MathType's native
                    // function run is the matching MTEF representation for an
                    // upright alphabetic operator and round-trips as one run.
                {
                    if (MathMlTokenStyleContract.ResolveEffectiveVariant(element) is "bold" or "italic" or "bold-italic")
                        EmitIdentifier(element, output);
                    else
                        EmitFunctionRun(operatorValue, output);
                }
                else
                    EmitOperator(element, output);
                break;
            }
            case "mtext":
                EmitText(element.Value, TypefaceText, output);
                break;
            case "mspace":
                EmitMathSpace(element, output);
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

    private static bool IsEmptyScriptBase(XElement element)
    {
        // Presentation-only wrappers cannot turn a zero-content base into an
        // atom. The same predicate governs explicit TeX prescript adoption and
        // native trailing-script isolation, so mstyle/mpadded/semantics nesting
        // never changes which neighboring object owns the script.
        if (element.Name.LocalName is not ("mrow" or "mi" or "mn" or "mtext"
            or "mstyle" or "mpadded" or "semantics" or "math"))
            return false;
        if (element.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
            return false;
        return element.Elements().Where(child => child.Name.LocalName is not ("annotation" or "annotation-xml"))
            .All(IsEmptyScriptBase);
    }

    private static void EmitScriptBase(XElement element, MtefWriter output)
    {
        // MTEF trailing-script templates consume the preceding OBJECT, not a
        // MathML token. Adjacent CHAR records of the same native run can be
        // coalesced by MathType (for example mtext "%" followed by an upright
        // identifier "e" becomes one Text object "%e"). If the base is emitted
        // as a bare CHAR, the following script then binds to that merged object
        // and silently captures the preceding sibling. A LINE is MTEF's native
        // grouping object, so every explicit MathML script base is materialized
        // as one LINE. This preserves the token/object boundary independent of
        // typeface, styling, or the preceding sibling instead of special-casing
        // percent, text, numbers, or any particular formula.
        EmitLine(element, output);
    }

    private static void EmitContainerChildren(
        XElement element,
        MtefWriter output,
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
            if (local == "mover" && IsOverDecoration(element, children[1]))
                return false;
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
        MtefWriter output)
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

    private static void EmitFraction(XElement element, MtefWriter output)
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
        output.RestoreContentSize();
    }

    private static void EmitSquareRoot(XElement element, MtefWriter output)
    {
        output.AddRange(new byte[] { RecordTemplate, 0, TemplateRoot, 0, 0 });
        EmitLineContents(SignificantChildren(element), output);
        using var scriptSize = output.EnterScriptSize();
        output.AddRange(new byte[] { RecordLine, LineNull });
        output.Add(RecordEnd);
    }

    private static void EmitNthRoot(XElement element, MtefWriter output)
    {
        var children = element.Elements().ToArray();
        if (children.Length < 2)
            throw new InvalidDataException("MathML mroot requires radicand and index.");
        output.AddRange(new byte[] { RecordTemplate, 0, TemplateRoot, 1, 0 });
        EmitLine(children[0], output);
        using var scriptSize = output.EnterScriptSize();
        EmitLine(children[1], output);
        output.Add(RecordEnd);
    }

    private static void EmitScript(XElement element, byte selector, MtefWriter output)
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
        // A native leading trailing-script has no base object. Give it its own
        // LINE so it cannot consume a preceding sibling. The TMPL is first in
        // that LINE, just as in genuine MathType storage; a preceding empty LINE
        // would supply a different baseline, and LINE_NULL is not a base at all.
        var emptyBase = IsEmptyScriptBase(children[0]);
        if (emptyBase) output.AddRange(new byte[] { RecordLine, 0 });
        else EmitScriptBase(children[0], output);
        output.AddRange(new byte[] { RecordTemplate, 0, selector, 0, 0 });
        using (output.EnterScriptSize())
        {
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
        }
        if (emptyBase) output.Add(RecordEnd);
    }

    private static void EmitOver(XElement element, MtefWriter output)
    {
        var children = element.Elements().ToArray();
        if (children.Length < 2)
        {
            EmitContainerChildren(element, output, inheritedMathVariant: null);
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
        if (!IsOverDecoration(element, children[1]))
        {
            // An upper argument is not an accent, even if its glyph resembles
            // one. The common script writer owns both ordinary superscripts
            // and big-operator limits (and their native size/object scopes).
            EmitScript(element, TemplateSup, output);
            return;
        }
        var over = NormalizeAccentMark(children[1].Value.Trim());
        var embellishmentBody = UnwrapSingleTokenAccentBody(children[0]);
        if (embellishmentBody is not null
            && !MathMlAccentGeometry.IsStretchy(children[1])
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
                EmitGlyphAccentTemplate(TemplateHat, children[0], children[1], output);
                return;
            case "~":
            case "˜":
                EmitGlyphAccentTemplate(TemplateTilde, children[0], children[1], output);
                return;
            case "⌢":
            case "⏜":
                EmitGlyphAccentTemplate(TemplateArc, children[0], children[1], output);
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
        // A centered accent on a group is a HatBox with a selected glyph, not
        // a trailing superscript. This includes dots, checks, breves and rings,
        // and works for a compound native atom (for example a contour integral)
        // without replacing that atom with a font-dependent fixed CHAR.
        if (ResolveOverAccent(element, children[1])
            && EnumerateBmpScalars(children[1].Value.Trim()).Count() == 1)
        {
            EmitGlyphAccentTemplate(TemplateHat, children[0], children[1], output);
            return;
        }
        EmitScriptBase(children[0], output);
        EmitTrailingScript(children[1], isSubscript: false, output);
    }

    private static void EmitUnder(XElement element, MtefWriter output)
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
        // Use the same object-boundary writer as msub. In particular, a null
        // basis must keep its entire trailing script inside one isolated LINE;
        // otherwise a subsequent read/rewrite can reinterpret it as a left
        // script on the following sibling.
        EmitScript(element, TemplateSub, output);
    }

    private static void EmitUnderOver(XElement element, MtefWriter output)
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
        EmitScript(element, TemplateSubSup, output);
    }

    private readonly struct NativeIntegral
    {
        public NativeIntegral(int variationKind, int integralCount,
            int symbolTypeface = TypefaceSymbol, int adornmentTypeface = TypefaceMtExtra)
        {
            VariationKind = variationKind;
            IntegralCount = integralCount;
            SymbolTypeface = symbolTypeface;
            AdornmentTypeface = adornmentTypeface;
        }

        public int VariationKind { get; }
        public int IntegralCount { get; }
        public int SymbolTypeface { get; }
        public int AdornmentTypeface { get; }
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
        if (integral.VariationKind != 0
            && int.TryParse((string?)element.Attribute("data-mtef-template-typeface"), out var styled))
            integral = new NativeIntegral(integral.VariationKind, integral.IntegralCount, styled,
                int.TryParse((string?)element.Attribute("data-mtef-adornment-typeface"), out var adornment)
                    ? adornment : TypefaceMtExtra);
        return integral.VariationKind != 0;
    }

    private static void EmitIntegralTemplate(
        NativeIntegral integral,
        XElement? main,
        XElement? lower,
        XElement? upper,
        MtefWriter output)
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
        // Limits step down from the enclosing size; a nested integral uses
        // SUB2 limits and SUBSYM, and returns to its parent's content size.
        var symbolSize = output.SymbolSize;
        using var limitsSize = output.EnterScriptSize();
        if (lower is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(lower, output);
        if (upper is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(upper, output);
        output.Add(symbolSize);
        if (integral.VariationKind == 0x05)
        {
            // Exact bytes emitted by genuine MathType 7 for the closed-loop
            // adornment: fnMTEXTRA, MTCode U+EE11, legacy position 0xD1.
            EmitScalar(
                0xEE11,
                integral.AdornmentTypeface,
                output,
                includeEncoded8: true,
                encoded8Override: 0xD1);
        }
        for (var index = 0; index < integral.IntegralCount; index++)
        {
            EmitScalar(
                0x222B,
                integral.SymbolTypeface,
                output,
                includeEncoded8: true,
                encoded8Override: 0xF2);
        }
        output.Add(RecordEnd);
        // Disposing limitsSize restores the enclosing size after the template,
        // not necessarily FULL (this integral itself may be in a script).
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
        if (bigOperator.Selector != 0
            && int.TryParse((string?)element.Attribute("data-mtef-template-typeface"), out var styled))
            bigOperator = new NativeBigOperator(bigOperator.Selector, bigOperator.MtCode, styled, bigOperator.Encoded8);
        return bigOperator.Selector != 0;
    }

    private static void EmitBigOperatorTemplate(
        NativeBigOperator bigOperator,
        XElement? main,
        XElement? lower,
        XElement? upper,
        MtefWriter output)
    {
        // MathType 7 native BigOp storage observed from genuine Equation.DSMT4:
        // variation 0x70 and line slots main/lower/upper followed by SYM + CHAR.
        EmitTemplateHeader(bigOperator.Selector, 0x70, output);
        if (main is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(main, output);
        var symbolSize = output.SymbolSize;
        using var limitsSize = output.EnterScriptSize();
        if (lower is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(lower, output);
        if (upper is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(upper, output);
        output.Add(symbolSize);
        EmitScalar(
            bigOperator.MtCode,
            bigOperator.Typeface,
            output,
            includeEncoded8: true,
            encoded8Override: bigOperator.Encoded8);
        output.Add(RecordEnd);
        // limitsSize restores the parent content size, including nested SUB.
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
        MtefWriter output)
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
        });
        using var scriptSize = output.EnterScriptSize();
        if (lower is null)
            output.AddRange(new byte[] { RecordLine, LineNull });
        else
            EmitLine(lower, output);
        if (upper is null)
            output.AddRange(new byte[] { RecordLine, LineNull });
        else
            EmitLine(upper, output);
        output.Add(RecordEnd);
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
        MtefWriter output)
    {
        output.AddRange(new byte[]
        {
            RecordTemplate,
            0,
            isSubscript ? TemplateSub : TemplateSup,
            0,
            0,
        });
        using var scriptSize = output.EnterScriptSize();
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
    }

    private static void EmitSingleSlotTemplate(
        byte selector,
        int variation,
        XElement body,
        MtefWriter output)
    {
        EmitTemplateHeader(selector, variation, output);
        EmitLine(body, output);
        output.Add(RecordEnd);
        output.RestoreContentSize();
    }

    private static void EmitGlyphAccentTemplate(
        byte selector, XElement body, XElement sourceMark, MtefWriter output)
    {
        // HatBoxClass owns a main LINE and a trailing glyph CHAR. A one-slot
        // template is incomplete native data even when our reader can infer a
        // plausible mark from its selector. The selected CHAR is authoritative:
        // tmHAT can carry a fixed dot/check/etc. as well as a stretching hat.
        EmitTemplateHeader(selector, 0, output);
        EmitLine(body, output);
        var mark = NormalizeAccentMark(sourceMark.Value.Trim());
        var expanding = MathMlAccentGeometry.IsStretchy(sourceMark);
        var combining = mark switch { "^" => 0x0302, "~" => 0x0303, "⌢" or "⏜" => 0x2322, _ => 0 };
        if (expanding && combining != 0)
            EmitScalar(combining, TypefaceFence, output);
        else
            EmitOperator(sourceMark, output);
        output.Add(RecordEnd);
        output.RestoreContentSize();
    }

    private static void EmitVectorTemplate(
        int variation,
        int combiningArrow,
        XElement body,
        MtefWriter output)
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
        output.RestoreContentSize();
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
        MtefWriter output)
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
        using (output.EnterScriptSize())
        {
            EmitColor(0, output);
            if (annotation is null)
                output.AddRange(new byte[] { RecordLine, LineNull });
            else
                EmitMathTypeColoredLine(annotation, output, resetColorAfter: false);
        }
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
        MtefWriter output)
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

        // Decoration changes only the CHAR's EMBELL list. First emit the bare
        // token through its canonical character path, including every explicit
        // font, native MTCode, encoded8 position and function-start option.
        // Re-resolving just a typeface here loses special operator glyphs as well
        // as styled identifiers; keeping their bytes makes the invariant apply
        // to all existing and future character mappings, not an accent whitelist.
        var character = output.CreateScratch();
        // Use the actual ordinary-object writer, not the fixed-glyph shortcut.
        // A large/composite operator owns a template, possibly multiple styled
        // glyphs. It must receive a group accent around that same native object.
        EmitNode(body, character);
        if (character.Count < 5 || character[0] != RecordChar)
            return false;
        var options = character[1];
        // Token emitters produce one unnudged CHAR. Do not decorate a template,
        // a run, or a pre-decorated object by guessing where its bytes end.
        if ((options & (0x08 | CharHasEmbellishment)) != 0) return false;
        var end = 2 + (character[2] == 0xFF ? 3 : 1);
        if ((options & 0x20) == 0) end += 2;
        if ((options & CharEncoded8) != 0) end++;
        if ((options & 0x10) != 0) end += 2;
        if (end != character.Count) return false;
        character[1] |= CharHasEmbellishment;
        output.AddRange(character);
        output.AddRange(new byte[] { RecordEmbellishment, 0, code, RecordEnd });
        return true;
    }

    private static bool MathTypeEmissionRequiresColorDefinition(
        XElement preparedMath,
        bool topLevelTableIsAlignedPile)
    {
        var topLevelElements = SignificantChildren(preparedMath)
            .OfType<XElement>()
            .Where(element => element.Name.LocalName is not ("annotation" or "annotation-xml"))
            .ToArray();
        XElement? alignedTopLevelTable =
            topLevelTableIsAlignedPile
            && topLevelElements.Length == 1
            && topLevelElements[0].Name.LocalName == "mtable"
                ? topLevelElements[0]
                : null;

        // EmitMatrix() is the only mtable path that writes COLOR 1. A top-level
        // table emitted through TryEmitAlignedPile() must be excluded; any nested
        // mtable still goes through EmitMatrix().
        if (preparedMath.DescendantsAndSelf()
            .Any(element => element.Name.LocalName == "mtable"
                && !ReferenceEquals(element, alignedTopLevelTable)))
            return true;

        // Native horizontal brace/bracket templates also use colored LINE slots.
        foreach (var element in preparedMath.Descendants())
        {
            if (element.Name.LocalName == "mover")
            {
                var children = element.Elements().ToArray();
                if (children.Length >= 2
                    && (IsHorizontalBraceMarker(children[1].Value.Trim(), top: true)
                        || TryGetAnnotatedHorizontalBraceBody(
                            children[0],
                            top: true,
                            out _)))
                    return true;
            }
            else if (element.Name.LocalName == "munder")
            {
                var children = element.Elements().ToArray();
                if (children.Length >= 2
                    && (IsHorizontalBraceMarker(children[1].Value.Trim(), top: false)
                        || TryGetAnnotatedHorizontalBraceBody(
                            children[0],
                            top: false,
                            out _)))
                    return true;
            }
        }
        return false;
    }

    private static bool TryEmitAlignedPile(
        XElement element,
        byte[] sourceMtef,
        MtefWriter output)
    {
        var rows = element.Elements()
            .Where(row => row.Name.LocalName is "mtr" or "mlabeledtr")
            .ToArray();
        if (rows.Length == 0) return false;

        var cells = rows.Select(row => row.Elements()
                .Where(cell => cell.Name.LocalName == "mtd")
                .ToArray())
            .ToArray();
        var columnCount = cells.Max(row => row.Length);
        if (columnCount < 2 || (columnCount & 1) != 0) return false;

        var columnAlignment = ((string?)element.Attribute("columnalign") ?? string.Empty)
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (columnAlignment.Length < columnCount) return false;
        for (var column = 0; column < columnCount; column++)
        {
            var expected = (column & 1) == 0 ? "right" : "left";
            if (!string.Equals(
                    columnAlignment[column],
                    expected,
                    StringComparison.OrdinalIgnoreCase))
                return false;
        }

        var pairCount = columnCount / 2;
        var rulerText = ((string?)element.Attribute(VisualTexMtefRulerStopsAttribute) ?? string.Empty)
            .Trim();
        var rulerStops = rulerText.Length == 0
            ? Array.Empty<int>()
            : rulerText.Split(',')
                .Select(value => int.TryParse(
                    value.Trim(),
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out var parsed)
                    ? parsed
                    : -1)
                .ToArray();
        var hasExactRuler = rulerStops.Length == pairCount
            && rulerStops.All(stop => stop > 0 && stop <= ushort.MaxValue)
            && rulerStops.Zip(rulerStops.Skip(1), (left, right) => right > left).All(valid => valid);

        // MathType's true multi-point alignment has two distinct fnMARKER
        // primitives. U+0009 starts a tab group; MTCode 0xEF00 is MathType's
        // non-printing alignment symbol inside that group. With an explicit
        // RULER, each group is positioned so its alignment symbol lands exactly
        // on the supplied stop. This is the native equivalent of one TeX
        // right/left `&` pair. A plain sequence of Ctrl+Tab markers is not enough:
        // MathType's default stops are 0.5in apart, so a wide left operand can
        // skip a stop and visibly misalign one row.
        output.AddRange(new byte[]
        {
            RecordPile,
            hasExactRuler ? (byte)0x02 : (byte)0,
            1, // horizontal alignment: left; tab groups own their reference points
            1, // vertical alignment: center line
        });
        if (hasExactRuler)
        {
            // LP_RULER is an inline field of LINE/PILE. MathType's native MTEF
            // readers consume n_stops immediately after the alignment bytes; the
            // enclosing option bit already identifies the payload as a RULER.
            // Emitting an additional standalone record tag (7) here shifts every
            // following byte and MathPage renders the equation blank.
            output.Add((byte)rulerStops.Length);
            foreach (var stop in rulerStops)
            {
                output.Add(0); // left tab; an in-group alignment symbol overrides it
                output.Add((byte)(stop & 0xff));
                output.Add((byte)((stop >> 8) & 0xff));
            }
        }

        // PILE is the root structure. Keep only root formatting state here;
        // missing font/encoding definitions are emitted globally before the
        // initial SIZE record by BuildRootStructure/AssembleRewrittenMtef.
        CopyRootLeadingFormattingRecords(sourceMtef, output);
        foreach (var row in cells)
        {
            output.AddRange(new byte[] { RecordLine, 0 });
            if (hasExactRuler)
            {
                for (var pair = 0; pair < pairCount; pair++)
                {
                    // Start the group at the next explicit ruler stop, then place
                    // the TeX pair boundary at MathType's own alignment symbol.
                    EmitScalar('\t', TypefaceMarker, output, includeEncoded8: false);
                    var leftColumn = pair * 2;
                    var rightColumn = leftColumn + 1;
                    if (leftColumn < row.Length)
                        EmitContainerChildren(row[leftColumn], output, inheritedMathVariant: null);
                    EmitScalar(
                        MathTypeAlignmentMarkerMtCode,
                        TypefaceMarker,
                        output,
                        includeEncoded8: false);
                    if (rightColumn < row.Length)
                        EmitContainerChildren(row[rightColumn], output, inheritedMathVariant: null);
                }
            }
            else
            {
                // Compatibility fallback for old MathML/third-party calls that
                // lack VisualTeX's SVG-derived ruler metadata. It preserves all
                // tab groups semantically, but production Office MathType exports
                // are expected to carry an exact ruler and use the branch above.
                for (var column = 0; column < columnCount; column++)
                {
                    if (column > 0)
                        EmitScalar('\t', TypefaceMarker, output, includeEncoded8: false);
                    if (column < row.Length)
                        EmitContainerChildren(row[column], output, inheritedMathVariant: null);
                }
            }
            output.Add(RecordEnd);
        }
        output.Add(RecordEnd);
        return true;
    }

    private static void EmitMatrix(XElement element, MtefWriter output)
    {
        var rows = element.Elements()
            .Where(row => row.Name.LocalName is "mtr" or "mlabeledtr")
            .ToArray();
        if (rows.Length == 0)
        {
            output.AddRange(new byte[] { RecordMatrix, 0, 1, 2, 1, 1, 1, 0, 0 });
            output.AddRange(new byte[] { RecordLine, LineNull, RecordEnd });
            output.RestoreContentSize();
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

    private static void EmitEnclose(XElement element, MtefWriter output)
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

    private static void EmitMultiScripts(XElement element, MtefWriter output)
    {
        var children = element.Elements().ToArray();
        if (children.Length == 0) return;

        var isolateEmpty = IsEmptyScriptBase(children[0]);
        if (isolateEmpty) output.AddRange(new byte[] { RecordLine, 0 });
        var prescriptsIndex = Array.FindIndex(
            children,
            child => child.Name.LocalName == "mprescripts");
        if (prescriptsIndex <= 0) prescriptsIndex = children.Length;

        // MTEF v5 tvSU_PRECEDES=1 makes a script template consume the next
        // object, rather than the preceding object. Emit prescript pairs from
        // outermost to innermost so the pair nearest the base is immediately
        // before it. MathML stores those pairs nearest-first.
        var prescriptCount = children.Length - prescriptsIndex - 1;
        var lastPrescriptPair = prescriptsIndex + 1
            + Math.Max(0, (prescriptCount - 1) / 2) * 2;
        for (var preIndex = lastPrescriptPair;
             prescriptCount > 0 && preIndex > prescriptsIndex;
             preIndex -= 2)
        {
            EmitScriptPair(
                children[preIndex],
                preIndex + 1 < children.Length ? children[preIndex + 1] : null,
                precedes: true,
                output);
        }

        EmitScriptBase(children[0], output);
        var index = 1;
        while (index < prescriptsIndex)
        {
            var sub = children[index];
            var sup = index + 1 < prescriptsIndex ? children[index + 1] : null;
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
                output.AddRange(new byte[] { RecordTemplate, 0, TemplateSubSup, 0, 0 });
                using (output.EnterScriptSize())
                {
                    EmitLine(sub, output);
                    EmitLine(sup, output);
                    output.Add(RecordEnd);
                }
            }
            index += 2;
        }
        if (isolateEmpty) output.Add(RecordEnd);
    }

    private static void EmitScriptPair(
        XElement? sub,
        XElement? sup,
        bool precedes,
        MtefWriter output)
    {
        var hasSub = sub is not null && sub.Name.LocalName != "none";
        var hasSup = sup is not null && sup.Name.LocalName != "none";
        if (!hasSub && !hasSup) return;

        var selector = hasSub && hasSup ? TemplateSubSup : hasSub ? TemplateSub : TemplateSup;
        output.AddRange(new byte[]
        {
            RecordTemplate,
            0,
            selector,
            precedes ? (byte)1 : (byte)0,
            0,
        });
        using var scriptSize = output.EnterScriptSize();
        if (hasSub) EmitLine(sub!, output);
        else output.AddRange(new byte[] { RecordLine, LineNull });
        if (hasSup) EmitLine(sup!, output);
        else output.AddRange(new byte[] { RecordLine, LineNull });
        output.Add(RecordEnd);
    }

    private static void EmitFenced(XElement element, MtefWriter output)
    {
        var open = NormalizeFence((string?)element.Attribute("open") ?? "(");
        var close = NormalizeFence((string?)element.Attribute("close") ?? ")");
        if ((open, close) is ("⟨", "|") or ("|", "⟩"))
        {
            // tmDIRAC owns two LINE slots and one shared vertical fence.
            // Missing halves still occupy a null LINE; ordinary CHARs lose
            // both scalable geometry and the structural bra/ket boundary.
            var left = open == "⟨";
            EmitTemplateHeader(TemplateDirac, left ? 1 : 2, output);
            EmitColor(0, output);
            if (left) EmitLineContents(SignificantChildren(element), output);
            else output.AddRange(new byte[] { RecordLine, LineNull });
            if (!left) EmitLineContents(SignificantChildren(element), output);
            else output.AddRange(new byte[] { RecordLine, LineNull });
            if (left) EmitFenceCharacter("⟨", opening: true, output);
            EmitFenceCharacter("|", opening: true, output);
            if (!left) EmitFenceCharacter("⟩", opening: false, output);
            output.Add(RecordEnd);
            return;
        }
        var selector = SelectFenceTemplate(open, close);
        if (selector is null)
        {
            if (!string.IsNullOrEmpty(open)) EmitOperator(open, output);
            foreach (var child in SignificantChildren(element)) EmitNode(child, output);
            if (!string.IsNullOrEmpty(close)) EmitOperator(close, output);
            return;
        }
        // tmINTERVAL uses two independent two-bit endpoint selectors, unlike
        // the left/right presence bits of the other ParBox templates (MTEF v5).
        var variation = selector == TemplateInterval
            ? IntervalEndpoint(open) | (IntervalEndpoint(close) << 4)
            : (string.IsNullOrEmpty(open) ? 0 : 1) | (string.IsNullOrEmpty(close) ? 0 : 2);
        EmitTemplateHeader(selector.Value, variation, output);
        EmitColor(0, output);
        EmitLineContents(SignificantChildren(element), output);
        if (!string.IsNullOrEmpty(open)) EmitFenceCharacter(open, opening: true, output);
        if (!string.IsNullOrEmpty(close)) EmitFenceCharacter(close, opening: false, output);
        output.Add(RecordEnd);
    }

    private static void EmitColor(int index, MtefWriter output) =>
        output.EmitNativeColor(index);

    private static void EmitMathTypeColoredLine(
        XElement element,
        MtefWriter output,
        bool resetColorAfter)
    {
        output.AddRange(new byte[] { RecordLine, 0 });
        EmitColor(1, output);
        EmitNode(element, output);
        output.Add(RecordEnd);
        if (resetColorAfter) EmitColor(0, output);
    }

    private static void EmitFenceCharacter(
        string value,
        bool opening,
        List<byte> output)
    {
        foreach (var scalar in EnumerateBmpScalars(value))
        {
            // MathType's fnFENCE typeface is MTCode-encoded rather than a
            // generic Unicode font. These values are captured from genuine
            // MathType 7 Equation Native streams for the same expandable
            // delimiters. Some delimiters (| and ||) use the same Unicode text
            // on both sides but distinct left/right MTCode glyphs, so the side
            // of the fence is part of the mapping.
            var mtCode = scalar switch
            {
                '(' => 0x0028,
                ')' => 0x0029,
                '[' => 0x005B,
                ']' => 0x005D,
                '{' => 0x007B,
                '}' => 0x007D,
                0x27E8 => 0x2329, // \langle
                0x27E9 => 0x232A, // \rangle
                '|' => opening ? 0xEC07 : 0xEC08,
                0x2016 => opening ? 0xEC09 : 0xEC0A, // \| / norm bars
                0x230A => 0xF8F0, // \lfloor
                0x230B => 0xF8FB, // \rfloor
                0x2308 => 0xF8EE, // \lceil
                0x2309 => 0xF8F9, // \rceil
                _ => scalar,
            };
            EmitScalar(mtCode, TypefaceFence, output, includeEncoded8: false);
        }
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
        if ((open, close) is ("⟨", "|") or ("|", "⟩")) return TemplateDirac;
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
        if (IntervalEndpoint(open) >= 0 && IntervalEndpoint(close) >= 0)
            return TemplateInterval;
        return null;
    }

    private static int IntervalEndpoint(string value) => value switch
    {
        "(" => 0, ")" => 1, "[" => 2, "]" => 3, _ => -1,
    };

    private static void EmitLine(XElement element, MtefWriter output)
    {
        output.AddRange(new byte[] { RecordLine, 0 });
        EmitNode(element, output);
        output.Add(RecordEnd);
    }

    private static void EmitLineContents(IEnumerable<XNode> nodes, MtefWriter output)
    {
        output.AddRange(new byte[] { RecordLine, 0 });
        foreach (var node in nodes) EmitNode(node, output);
        output.Add(RecordEnd);
    }

    private static void EmitIdentifier(
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
            case 0x2112: normalizedScalar = 'L'; mathVariant = "script"; return true;
            case 0x2115: normalizedScalar = 'N'; mathVariant = "double-struck"; return true;
            case 0x2119: normalizedScalar = 'P'; mathVariant = "double-struck"; return true;
            case 0x211A: normalizedScalar = 'Q'; mathVariant = "double-struck"; return true;
            case 0x211B: normalizedScalar = 'R'; mathVariant = "script"; return true;
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

    private static bool TryEuclidMathOneGreekVariantEncoded8(
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

    private static bool TryEuclidMathOneDotlessJ(
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

    private static bool TryEuclidMathOneOperatorEncoded8(
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

    private static bool TryEuclidMathTwoOperatorEncoded8(
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

    private static bool IsFarEastTextScalar(int scalar) =>
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
    private static bool NeedsFarEastCodePage(int scalar) =>
        scalar >= 0x80 && scalar <= char.MaxValue && !IsWinAllBasicCodePageScalar(scalar);

    private static bool IsFarEastCodePageOperator(int scalar)
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

    private static string NormalizeOperatorToken(string value) => value switch
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

    private static void EmitOperator(XElement element, List<byte> output)
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

    private static void EmitOperator(string value, List<byte> output)
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
            EmitScalar(
                NormalizeMathTypeMtCode(scalar),
                IsFarEastTextScalar(scalar) ? TypefaceFarEastText
                    : IsMathematicalSymbolScalar(scalar) ? TypefaceFunction : TypefaceText,
                output);
        }
    }

    private static bool IsVerifiedMtExtraOperator(int scalar, int encoded)
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

    private static void EmitText(string value, int typeface, List<byte> output)
    {
        foreach (var scalar in EnumerateBmpScalars(value))
        {
            if (IsWhiteSpaceScalar(scalar)) EmitScalar(scalar, TypefaceSpace, output);
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

    private static void DecomposeMathematicalAlphanumerics(XElement math)
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

    // The semantic MTCode and the font's encoded position are independent.
    // Keep presence and width from a native CHAR instead of coercing every
    // font position to encoded8 or inventing a missing MTCode on rewrite.
    private static bool TryEmitPreservedCharacter(
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

    private static void EmitScalar(
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
        private readonly List<string?> _colorDefinitions = new();
        private readonly Dictionary<int, byte> _builtInCharacterStyles = new();
        private readonly Dictionary<int, int> _builtInFontDefinitions = new();
        private int _position;
        private int _structureDepth;
        private int _currentColorIndex;

        private StructureScope EnterStructure()
        {
            if (_structureDepth >= FormulaXml.MaximumDepth)
                throw new InvalidDataException($"MathType MTEF exceeds the native nesting limit of {FormulaXml.MaximumDepth} levels at offset {_position}.");
            _structureDepth++;
            return new StructureScope(this);
        }

        private readonly struct StructureScope : IDisposable
        {
            private readonly MtefStructureReader _reader;
            internal StructureScope(MtefStructureReader reader) => _reader = reader;
            public void Dispose() => _reader._structureDepth--;
        }

        internal MtefStructureReader(byte[] data, int position)
        {
            _data = data;
            ScanPrefixDefinitions(position);
            _position = position;
        }

        private void ScanPrefixDefinitions(int rootPosition)
        {
            var prefix = ReadPrefixLayout(_data);
            foreach (var binding in ReadNativeCharacterStyleBindings(_data).Where(item => item.Key > 0))
            {
                _builtInFontDefinitions[binding.Key] = binding.Value.FontDefinition;
                _builtInCharacterStyles[binding.Key] = binding.Value.CharacterStyle;
            }
            if (rootPosition != prefix.Root)
                throw new InvalidDataException("MTEF structural reader received a non-root offset.");
            foreach (var record in prefix.Records)
            {
                if (record.Type == RecordEqnPrefs) continue;
                if (record.Type == RecordColorDef)
                {
                    _position = record.Start;
                    ReadColorDefinitionRecord();
                    if (_position != record.End)
                        throw new InvalidDataException($"MTEF color definition reader diverged at offset {record.Start}.");
                    continue;
                }
                if (record.Type == RecordColor)
                {
                    _position = record.Start + 1;
                    SetCurrentColorIndex(ReadUnsigned(_data, ref _position));
                    if (_position != record.End)
                        throw new InvalidDataException($"MTEF color reader diverged at offset {record.Start}.");
                    continue;
                }
                if (record.Type is not RecordEncodingDef and not RecordFontDef and not RecordFontStyleDef)
                    continue;
                _position = record.Start;
                ReadDefinitionRecord();
                if (_position != record.End)
                    throw new InvalidDataException($"MTEF definition reader diverged at offset {record.Start}.");
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
            XNode[] nodes;
            if (Peek() == RecordLine)
                nodes = ReadLine().Nodes().ToArray();
            else if (Peek() == RecordPile)
                nodes = new XNode[] { ReadPile() };
            else if (Peek() == RecordMatrix)
                nodes = new XNode[] { ReadMatrix() };
            else
                throw new InvalidDataException(
                    $"Unsupported MathType root record {Peek()} at offset {_position}.");
            if (!HasOnlyEquationEndRemaining())
                throw new InvalidDataException($"MathType MTEF root does not end at the equation boundary at offset {_position}.");
            return nodes;
        }

        internal bool HasOnlyEquationEndRemaining()
        {
            SkipFormattingRecords();
            return _position == _data.Length - 1
                && Peek() == RecordEnd;
        }

        private XElement ReadLine()
        {
            using var structure = EnterStructure();
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
            using var structure = EnterStructure();
            Expect(RecordPile);
            var options = ReadByte();
            SkipNudge(options);
            Skip(2); // horizontal and vertical alignment
            var rulerStops = (options & 0x02) != 0
                ? ReadRulerOffsets()
                : Array.Empty<int>();
            var table = new XElement("mtable");
            var hasAlignmentTabs = false;
            while (true)
            {
                SkipFormattingRecords();
                if (Peek() == RecordEnd)
                {
                    _position++;
                    break;
                }
                var line = ReadLine();
                var lineNodes = line.Nodes().ToArray();
                if (lineNodes.OfType<XElement>().Any(IsMtefAlignmentPoint)
                    && TryBuildAlignmentSymbolRow(lineNodes, out var alignedRow))
                {
                    hasAlignmentTabs = true;
                    table.Add(alignedRow);
                    continue;
                }

                var row = new XElement("mtr");
                var cell = new XElement("mtd");
                foreach (var node in lineNodes)
                {
                    if (node is XElement marker && IsMtefAlignmentTab(marker))
                    {
                        hasAlignmentTabs = true;
                        marker.Remove();
                        row.Add(cell);
                        cell = new XElement("mtd");
                        continue;
                    }
                    node.Remove();
                    cell.Add(node);
                }
                row.Add(cell);
                table.Add(row);
            }

            if (hasAlignmentTabs)
            {
                // MathType's native multi-point aligned equations are PILE rows
                // separated by fnMARKER/U+0009 tab characters. Reconstruct the
                // alternating right/left MathJax table shape so strict MTEF
                // round-trip validation and MathML->LaTeX recovery preserve every
                // `&`, including empty cells introduced by `&&`.
                table.SetAttributeValue("data-mtef-tabs", "true");
                var columnCount = table.Elements()
                    .Select(row => row.Elements().Count(cellNode => cellNode.Name.LocalName == "mtd"))
                    .DefaultIfEmpty(0)
                    .Max();
                if (columnCount > 0)
                {
                    table.SetAttributeValue(
                        "columnalign",
                        string.Join(" ", Enumerable.Range(0, columnCount)
                            .Select(column => (column & 1) == 0 ? "right" : "left")));
                    if (rulerStops.Length == columnCount / 2)
                    {
                        table.SetAttributeValue(
                            VisualTexMtefRulerStopsAttribute,
                            string.Join(",", rulerStops.Select(stop =>
                                stop.ToString(CultureInfo.InvariantCulture))));
                    }
                }
            }
            else
            {
                // Preserve the MTEF container kind. A two-row/one-column PILE
                // inside ordinary parentheses is MathType's native binomial
                // representation, while a MATRIX of the same shape is not.
                table.SetAttributeValue("data-mtef-pile", "true");
            }
            return table;
        }

        private static bool IsMtefAlignmentTab(XElement element) =>
            element.Name.LocalName == "mspace"
            && string.Equals(
                (string?)element.Attribute("data-mtef-tab"),
                "true",
                StringComparison.OrdinalIgnoreCase);

        private static bool IsMtefAlignmentPoint(XElement element) =>
            element.Name.LocalName == "mspace"
            && string.Equals(
                (string?)element.Attribute("data-mtef-align"),
                "true",
                StringComparison.OrdinalIgnoreCase);

        private static bool TryBuildAlignmentSymbolRow(
            XNode[] lineNodes,
            out XElement row)
        {
            row = new XElement("mtr");
            var groups = new List<List<XNode>>();
            var current = new List<XNode>();
            foreach (var node in lineNodes)
            {
                if (node is XElement marker && IsMtefAlignmentTab(marker))
                {
                    // A VisualTeX/MathType aligned PILE starts each tab group with
                    // a TAB. Do not materialize the leading empty group as a cell.
                    if (current.Count > 0) groups.Add(current);
                    current = new List<XNode>();
                    continue;
                }
                current.Add(node);
            }
            if (current.Count > 0) groups.Add(current);
            if (groups.Count == 0) return false;

            foreach (var group in groups)
            {
                var alignmentIndexes = group
                    .Select((node, index) => new { Node = node, Index = index })
                    .Where(item => item.Node is XElement element && IsMtefAlignmentPoint(element))
                    .Select(item => item.Index)
                    .ToArray();
                if (alignmentIndexes.Length != 1)
                {
                    row = new XElement("mtr");
                    return false;
                }

                var alignmentIndex = alignmentIndexes[0];
                row.Add(new XElement(
                    "mtd",
                    group.Take(alignmentIndex).Select(CloneNode)));
                row.Add(new XElement(
                    "mtd",
                    group.Skip(alignmentIndex + 1).Select(CloneNode)));
            }
            return row.Elements().Any();
        }

        private XElement ReadMatrix()
        {
            using var structure = EnterStructure();
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
            var pendingPrescripts = new List<ParsedTemplate>();

            void AddObject(XNode node, bool mergeAdjacentCharacters = false)
            {
                if (pendingPrescripts.Count != 0)
                {
                    node = BuildPrecedingScriptNode(node, pendingPrescripts);
                    pendingPrescripts.Clear();
                    mergeAdjacentCharacters = false;
                }
                if (mergeAdjacentCharacters && node is XElement character)
                    AppendNode(row, character);
                else
                    row.Add(node);
            }

            while (true)
            {
                SkipFormattingRecords();
                var record = Peek();
                if (record == RecordEnd)
                {
                    if (pendingPrescripts.Count != 0)
                        throw new InvalidDataException(
                            $"MathType preceding-script template at offset {_position} has no following base object.");
                    _position++;
                    RestoreExplicitNumberRuns(row);
                    return;
                }
                if (record == RecordChar)
                {
                    AddObject(ReadCharacter(), mergeAdjacentCharacters: true);
                    continue;
                }
                if (record == RecordTemplate)
                {
                    var templateColor = CurrentMathColor;
                    var template = ReadTemplate();
                    template.MathColor = templateColor;
                    if (template.Node is not null)
                        ApplyMathColor(template.Node, templateColor);
                    if (template.ScriptKind != 0)
                    {
                        if (template.ScriptPrecedes)
                        {
                            pendingPrescripts.Add(template);
                            continue;
                        }
                        if (pendingPrescripts.Count != 0)
                            throw new InvalidDataException(
                                $"MathType trailing-script template at offset {_position} appears before the base of a preceding script.");
                        // MTEF script templates store only their sub/sup slots;
                        // a native LINE may begin with one (including inside
                        // another script). Its missing base is an empty group,
                        // not a malformed stream or the following object.
                        var baseNode = row.Nodes().LastOrDefault();
                        baseNode?.Remove();
                        row.Add(BuildScriptNode(baseNode ?? new XElement("mrow"), template));
                    }
                    else if (template.Node is not null)
                    {
                        AddObject(template.Node);
                    }
                    continue;
                }
                if (record == RecordLine)
                {
                    AddObject(ReadLine());
                    continue;
                }
                if (record == RecordPile)
                {
                    AddObject(ReadPile());
                    continue;
                }
                if (record == RecordMatrix)
                {
                    AddObject(ReadMatrix());
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

        private XElement ReadCharacter() => ReadCharacter(out _);

        private XElement ReadCharacter(out int nativeTypeface)
        {
            Expect(RecordChar);
            var options = ReadByte();
            SkipNudge(options);
            var typeface = ReadSigned();
            nativeTypeface = typeface;
            int scalar = -1;
            var hasMtCode = (options & CharNoMtCode) == 0;
            if (hasMtCode)
            {
                Require(_data, _position, 2);
                scalar = _data[_position] | (_data[_position + 1] << 8);
                _position += 2;
            }
            int encoded = -1;
            if ((options & CharEncoded8) != 0) encoded = ReadByte();
            if ((options & CharEncoded16) != 0)
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
            byte explicitCharacterStyle = 0;
            var explicitEuclidOperator = false;
            NativeGlyphAuthority? explicitGlyph = null;
            if (typeface < 0)
            {
                explicitMathVariant = ResolveExplicitMathVariant(
                    -typeface,
                    scalar,
                    encoded,
                    hasMtCode,
                    out var explicitText,
                    out explicitCharacterStyle,
                    out explicitEuclidOperator,
                    out explicitGlyph);
                if (!string.IsNullOrEmpty(explicitText)) text = explicitText;
                if (explicitEuclidOperator)
                {
                    // For explicit Euclid operator glyphs the encoded8 byte is
                    // only a font position. Normally MTCode remains the semantic
                    // Unicode character; MathType's preceq/succeq pair is the
                    // exception, using private MTCode E938/E939 internally.
                    text = scalar switch
                    {
                        0xE938 => "⪯",
                        0xE939 => "⪰",
                        _ => text,
                    }; // Font-position decoding and explicit style both survive.
                }
                else if (scalar == 0xED02 && encoded == 0xF8)
                {
                    // Genuine MathType 7 dotless-j: Euclid Math One PUA MTCode
                    // 0xED02 / position 0xF8. Recover Unicode U+0237 so the
                    // public MathML/LaTeX layer sees \\jmath, not script PUA text.
                    text = "ȷ";
                    explicitMathVariant = null;
                }
            }

            if (typeface >= 0
                && _builtInFontDefinitions.TryGetValue(typeface, out var builtInFont)
                && _builtInCharacterStyles.TryGetValue(typeface, out var builtInStyle)
                && builtInFont > 0 && builtInFont <= _fontDefinitions.Count)
            {
                var definition = _fontDefinitions[builtInFont - 1];
                var encoding = definition.EncodingIndex > 0 && definition.EncodingIndex <= _encodingDefinitions.Count
                    ? _encodingDefinitions[definition.EncodingIndex - 1] : null;
                explicitGlyph = ResolveNativeGlyphAuthority(
                    definition.Name, encoding, scalar, encoded, hasMtCode);
                if (explicitGlyph is not null)
                {
                    text = ((char)explicitGlyph.Scalar).ToString();
                    explicitCharacterStyle = builtInStyle;
                    explicitMathVariant = NativeGlyphCharacterVariant(builtInStyle, explicitGlyph);
                }
            }

            XElement token;
            if (typeface >= 0 && explicitGlyph is not null)
            {
                token = new XElement(explicitGlyph.TokenClass, text);
                if (explicitMathVariant is not null)
                    token.SetAttributeValue("mathvariant", explicitMathVariant);
                if (explicitGlyph.Alternate)
                    token.SetAttributeValue("data-mjx-alternate", "1");
                SetPreservedCharacterAttributes(token, typeface, options, scalar, encoded);
                if (explicitCharacterStyle != 0)
                    token.SetAttributeValue("data-mtef-character-style", explicitCharacterStyle);
            }
            else if (typeface < 0)
            {
                if (explicitEuclidOperator || explicitGlyph is not null)
                {
                    // A native glyph authority defines token role and alternate
                    // identity independently of the signed font-style index.
                    token = new XElement(explicitGlyph?.TokenClass ?? "mo", text);
                    if (explicitGlyph?.Alternate == true)
                        token.SetAttributeValue("data-mjx-alternate", "1");
                    if (!string.IsNullOrWhiteSpace(explicitMathVariant))
                        token.SetAttributeValue("mathvariant", explicitMathVariant);
                    SetPreservedCharacterAttributes(token, typeface, options, scalar, encoded);
                    if (explicitCharacterStyle != 0)
                        token.SetAttributeValue("data-mtef-character-style", explicitCharacterStyle);
                }
                else
                {
                    // Native explicit font styles do not have a Number face;
                    // a numeric CHAR keeps its digit identity while the font
                    // definition supplies typography. Adjacent equal-style
                    // numeric CHARs are rejoined by AppendNode as usual.
                    token = new XElement(text.All(char.IsDigit) ? "mn" : "mi", text);
                    if (!string.IsNullOrWhiteSpace(explicitMathVariant))
                        token.SetAttributeValue("mathvariant", explicitMathVariant);
                    SetPreservedCharacterAttributes(token, typeface, options, scalar, encoded);
                    if (explicitCharacterStyle != 0)
                        token.SetAttributeValue(
                            "data-mtef-character-style",
                            explicitCharacterStyle);
                }
            }
            else if (typeface == TypefaceVector)
                token = new XElement(text.All(char.IsDigit) ? "mn" : "mi", new XAttribute("mathvariant", "bold"), text);
            else if (typeface == TypefaceText)
            {
                token = new XElement("mtext", text);
                if (_builtInFontDefinitions.TryGetValue(typeface, out var textFont)
                    && textFont > 0 && textFont <= _fontDefinitions.Count
                    && _builtInCharacterStyles.TryGetValue(typeface, out var textBits))
                {
                    var textVariant = NativeTextFontVariant(_fontDefinitions[textFont - 1].Name, textBits);
                    if (textVariant != "normal") token.SetAttributeValue("mathvariant", textVariant);
                }
            }
            else if (typeface == TypefaceFarEastText
                && scalar is not (0x220B or 0x25EF or 0x2661))
                token = new XElement("mtext", text);
            else if (typeface != TypefaceSymbol
                && TryNormalizeLetterlikeScalar(
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
            else if (typeface == TypefaceMarker && scalar == '\t')
                token = new XElement("mspace", new XAttribute("data-mtef-tab", "true"));
            else if (typeface == TypefaceMarker && scalar == MathTypeAlignmentMarkerMtCode)
                token = new XElement("mspace", new XAttribute("data-mtef-align", "true"));
            else if (typeface == TypefaceFarEastText && scalar == 0x220B)
                token = new XElement("mo", "∋");
            else if (typeface == TypefaceFarEastText && scalar == 0x25EF)
                token = new XElement("mo", "◯");
            else if (typeface == TypefaceFarEastText && scalar == 0x2661)
                token = new XElement("mo", "♡");
            else if (typeface == TypefaceMtExtra && scalar == 0xFFFD && encoded == 0x6E)
            {
                // MathType 7 uses this MT Extra position for the diamondsuit glyph.
                // Some unsupported TeX imports may also collapse to the same
                // persisted glyph; once that happens the original command is no
                // longer recoverable, so preserve the actual stored glyph.
                token = new XElement("mo", "♢");
            }
            else if (typeface == TypefaceMtExtra && scalar == 0x2016)
                token = new XElement("mo", text);
            else if (typeface == TypefaceMtExtra && scalar == 0x2210 && encoded == 0x43)
            {
                // MathType reuses U+2210 internally for ordinary TeX \\amalg.
                // The dedicated COPRODUCT template is what represents \\coprod;
                // a plain MT Extra CHAR at position 0x43 is amalg instead.
                token = new XElement("mo", "⨿");
            }
            else if (typeface == TypefaceMtExtra && IsVerifiedMtExtraOperator(scalar, encoded))
                token = new XElement("mo", text);
            else if (typeface == TypefaceSpace)
            {
                var space = hasMtCode
                    ? MathSpaceContract.FindNativeSpace(scalar)
                    : null;
                token = new XElement("mspace", new XAttribute("width", space?.MathMlWidth ?? "0.2em"));
                if (space is not null)
                    SetPreservedCharacterAttributes(token, typeface, options, scalar, encoded);
            }
            else if (typeface == TypefaceSymbol
                && scalar is 0x2135 or 0x2118 or 0x211C or 0x2111)
            {
                // MathType stores \aleph, \wp, \Re and \Im in Symbol typeface,
                // but exports them as identifier tokens rather than operators.
                // Preserve that token class without misclassifying ℜ/ℑ as
                // Fraktur solely from their Unicode letterlike code points.
                token = new XElement("mi", text);
                // Fixed letterlike glyphs still inherit the actual Symbol style
                // from equation preferences. A plain glyph must not become an
                // implicit-italic single mi merely because its tag is mi.
                if (_builtInCharacterStyles.TryGetValue(TypefaceSymbol, out var fixedStyle))
                    token.SetAttributeValue("mathvariant", fixedStyle switch
                    { 1 => "bold", 2 => "italic", 3 => "bold-italic", _ => "normal" });
            }
            else if (typeface == TypefaceSymbol)
                token = new XElement("mo", text);
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
            else
                token = new XElement("mi", text);

            if (typeface > 0 && token.Name.LocalName == "mi"
                && (token.Attribute("mathvariant") is null
                    || TryNativeCharacterStyle((string?)token.Attribute("mathvariant"), out _))
                && _builtInCharacterStyles.TryGetValue(typeface, out var actualIdentifierStyle)
                && _builtInFontDefinitions.TryGetValue(typeface, out var actualIdentifierFont)
                && actualIdentifierFont > 0 && actualIdentifierFont <= _fontDefinitions.Count)
            {
                // A native logical face is an adapter binding, not a MathML
                // default. Materialize its actual style and keep the font-scoped
                // CHAR tuple so unchanged native read/write remains byte exact.
                token.SetAttributeValue("mathvariant", NativeCharacterVariant(actualIdentifierStyle));
                SetPreservedCharacterAttributes(token, typeface, options, scalar, encoded);
                token.SetAttributeValue("data-mtef-character-style", actualIdentifierStyle);
            }

            // A CHAR contains a fixed glyph; actual scalable delimiters are
            // decoded by the TMPL path as mfenced. Preserve that distinction for
            // nested bra/ket characters inside a scalable absolute-value pair.
            if (token.Name.LocalName == "mo"
                && NormalizeFence(text) is "(" or ")" or "[" or "]" or "{" or "}"
                    or "⟨" or "⟩" or "⌈" or "⌉" or "⌊" or "⌋" or "|" or "‖")
            {
                token.SetAttributeValue("fence", "false");
                token.SetAttributeValue("stretchy", "false");
            }

            // EMBELL is not restricted to over-accents: native MathType stores
            // primes here, including on characters nested inside another script.
            // Apply accents to the character, then attach all right/left primes
            // at that character's script level. Never silently discard a record.
            var rightPrimes = new StringBuilder();
            var leftPrimes = new StringBuilder();
            foreach (var embellishment in embellishments)
            {
                switch (embellishment)
                {
                    case EmbellPrime: rightPrimes.Append('′'); continue;
                    case EmbellDoublePrime: rightPrimes.Append('″'); continue;
                    case EmbellTriplePrime: rightPrimes.Append('‴'); continue;
                    case EmbellBackPrime: leftPrimes.Append('‵'); continue;
                }
                var mark = EmbellishmentMark(embellishment);
                if (mark.Length == 0)
                    throw new InvalidDataException(
                        $"Unsupported MathType character embellishment {embellishment}; conversion stopped to avoid losing a symbol.");
                token = new XElement(
                    "mover",
                    new XAttribute("accent", "true"),
                    token,
                    new XElement("mo", new XAttribute("stretchy", "false"), mark));
            }
            if (leftPrimes.Length > 0)
            {
                token = new XElement("mmultiscripts", token,
                    new XElement("none"),
                    rightPrimes.Length > 0 ? new XElement("mo", rightPrimes.ToString()) : new XElement("none"),
                    new XElement("mprescripts"), new XElement("none"),
                    new XElement("mo", leftPrimes.ToString()));
            }
            else if (rightPrimes.Length > 0)
                token = new XElement("msup", token, new XElement("mo", rightPrimes.ToString()));
            ApplyMathColor(token, CurrentMathColor);
            return token;
        }

        private static void SetPreservedCharacterAttributes(
            XElement token, int typeface, byte options, int scalar, int encoded)
        {
            token.SetAttributeValue("data-mtef-explicit-typeface", typeface);
            if ((options & CharNoMtCode) != 0)
                token.SetAttributeValue("data-mtef-no-mtcode", "true");
            else
                token.SetAttributeValue("data-mtef-mtcode", scalar);
            if ((options & CharEncoded16) != 0)
                token.SetAttributeValue("data-mtef-encoded16", encoded);
            else if ((options & CharEncoded8) != 0)
                token.SetAttributeValue("data-mtef-encoded8", encoded);
            if ((options & CharFunctionStart) != 0)
                token.SetAttributeValue("data-mtef-function-start", "true");
        }

        private string? ResolveExplicitMathVariant(
            int styleIndex,
            int scalar,
            int encoded,
            bool hasMtCode,
            out string text,
            out byte characterStyle,
            out bool euclidOperator,
            out NativeGlyphAuthority? nativeGlyph)
        {
            text = encoded is >= 0x21 and <= 0x7E
                ? ((char)encoded).ToString()
                : ((char)scalar).ToString();
            string? fontName = null;
            string? encodingName = null;
            var fontDefinitionIndex = 0;
            characterStyle = 0;
            euclidOperator = false;
            nativeGlyph = null;
            if (styleIndex > 0 && styleIndex <= _fontStyleDefinitions.Count)
            {
                var style = _fontStyleDefinitions[styleIndex - 1];
                characterStyle = style.CharacterStyle;
                if (style.FontDefinitionIndex > 0
                    && style.FontDefinitionIndex <= _fontDefinitions.Count)
                {
                    var definition = _fontDefinitions[style.FontDefinitionIndex - 1];
                    fontDefinitionIndex = style.FontDefinitionIndex;
                    fontName = definition.Name;
                    if (definition.EncodingIndex > 0 && definition.EncodingIndex <= _encodingDefinitions.Count)
                        encodingName = _encodingDefinitions[definition.EncodingIndex - 1];
                }
            }

            nativeGlyph = ResolveNativeGlyphAuthority(fontName, encodingName, scalar, encoded, hasMtCode);
            if (nativeGlyph is not null)
            {
                text = ((char)nativeGlyph.Scalar).ToString();
                euclidOperator = nativeGlyph.TokenClass == "mo";
                return NativeGlyphCharacterVariant(characterStyle, nativeGlyph);
            }

            var font = fontName ?? string.Empty;
            euclidOperator =
                font.IndexOf("Euclid Math One", StringComparison.OrdinalIgnoreCase) >= 0
                && TryEuclidMathOneOperatorEncoded8(scalar, out _)
                || font.IndexOf("Euclid Math Two", StringComparison.OrdinalIgnoreCase) >= 0
                && TryEuclidMathTwoOperatorEncoded8(scalar, out _);
            var characterVariant=characterStyle switch
            { 1=>"bold", 2=>"italic", 3=>"bold-italic", _=>(string?)null };
            if (euclidOperator)
            {
                text=scalar switch { 0xE938=>"⪯", 0xE939=>"⪰", _=>((char)scalar).ToString() };
                return characterVariant;
            }
            if (fontDefinitionIndex > 0
                && _builtInFontDefinitions.TryGetValue(TypefaceFarEastText, out var farEastFont)
                && fontDefinitionIndex == farEastFont
                && string.Equals(encodingName, "WinAllCodePages", StringComparison.OrdinalIgnoreCase)
                && !IsFarEastTextScalar(scalar)
                && (IsMathematicalSymbolScalar(scalar) || NeedsFarEastCodePage(scalar)))
            {
                // VisualTeX's SimSun binding for operators such as U+25B3.
                text = ((char)scalar).ToString();
                euclidOperator = true;
                return characterVariant;
            }
            if (string.Equals(font, "MT Extra", StringComparison.OrdinalIgnoreCase))
            {
                text=scalar==0xFFFD && encoded==0x6E ? "♢"
                    : scalar==0x2210 && encoded==0x43 ? "⨿" : ((char)scalar).ToString();
                euclidOperator=IsMathematicalSymbolToken(text) || !text.All(char.IsLetter);
                return characterVariant;
            }
            if (string.Equals(font, "Symbol", StringComparison.OrdinalIgnoreCase))
            {
                // In an explicit Symbol FONT_STYLE_DEF, encoded8 is the legacy
                // Symbol-font position (for example epsilon is byte 0x65, 'e'),
                // while MTCode remains the semantic Unicode Greek scalar. Using
                // encoded8 as text silently changed a genuine bold epsilon into
                // a bold Latin e. Keep the scalar; CharacterStyle below still
                // determines bold/italic presentation.
                text = ((char)scalar).ToString();
                euclidOperator=IsMathematicalSymbolToken(text) || !text.All(char.IsLetter);
                return characterVariant;
            }
            if (font.IndexOf("Euclid Math One", StringComparison.OrdinalIgnoreCase) >= 0
                && TryEuclidMathOneGreekVariantEncoded8(scalar, out var expectedGreekEncoded8)
                && encoded == expectedGreekEncoded8)
            {
                // MathType uses Euclid Math One for several Greek variants
                // (\epsilon, \varrho, \varkappa). They remain ordinary Greek
                // identifiers, not \mathcal/script characters.
                text = ((char)scalar).ToString();
                return null;
            }
            if (font.IndexOf("Euclid Math One", StringComparison.OrdinalIgnoreCase) >= 0)
                return characterStyle is 1 or 3 ? "bold-script" : "script";
            if (font.IndexOf("Euclid Math Two", StringComparison.OrdinalIgnoreCase) >= 0)
                return "double-struck";
            if (font.IndexOf("Fraktur", StringComparison.OrdinalIgnoreCase) >= 0)
                return characterStyle is 1 or 3 ? "bold-fraktur" : "fraktur";
            if (string.Equals(font, "Courier New", StringComparison.OrdinalIgnoreCase)
                || string.Equals(font, "Arial", StringComparison.OrdinalIgnoreCase))
                return NativeTextFontVariant(font, characterStyle);

            // A lexical letterlike alias is not permission to discard explicit
            // character style. Nonzero FONT_STYLE_DEF bits remain authoritative.
            if (characterStyle == 0
                && TryNormalizeLetterlikeScalar(scalar, out var letterlikeText, out var letterlikeVariant))
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

        private static void RestoreExplicitNumberRuns(XElement row)
        {
            // An explicit FONT_STYLE_DEF replaces the logical Number face.
            // Recover decimal runs from adjacent same-font CHARs, never across
            // a template, a space, or another character style. Punctuation by
            // itself remains literal; the surrounding numeric grammar is needed.
            var children = row.Elements().ToArray();
            for (var start = 0; start < children.Length; start++)
            {
                var first = children[start];
                var face = (int?)first.Attribute("data-mtef-explicit-typeface");
                if (face is null or >= 0 || !IsNumericPart(first)) continue;
                var end = start + 1;
                while (end < children.Length && IsNumericPart(children[end])
                       && (int?)children[end].Attribute("data-mtef-explicit-typeface") == face
                       && (string?)children[end].Attribute("mathvariant") == (string?)first.Attribute("mathvariant")
                       && (string?)children[end].Attribute("mathcolor") == (string?)first.Attribute("mathcolor"))
                    end++;
                var run = new StringBuilder();
                for (var index = start; index < end; index++) run.Append(children[index].Value);
                var value = run.ToString();
                if (value.Any(char.IsDigit) && value.Count(c => c is '.' or '\u066B') == 1)
                {
                    first.Name = first.Name.Namespace + "mn";
                    first.Value = value;
                    for (var index = start + 1; index < end; index++) children[index].Remove();
                }
                start = end - 1;
            }

            static bool IsNumericPart(XElement token) => !token.HasElements
                && token.Name.LocalName is "mn" or "mi"
                && token.Value.Length > 0 && token.Value.All(c => char.IsDigit(c) || c is '.' or '\u066B');
        }

        private void ReadColorDefinitionRecord()
        {
            Expect(RecordColorDef);
            var options = ReadByte();
            var cmyk = (options & 0x01) != 0;
            var components = new int[cmyk ? 4 : 3];
            for (var index = 0; index < components.Length; index++)
            {
                Require(_data, _position, 2);
                components[index] = _data[_position] | (_data[_position + 1] << 8);
                _position += 2;
            }
            if ((options & 0x04) != 0)
                _ = ReadNullTerminatedString(_data, ref _position);
            if (components.Any(component => component is < 0 or > 1000))
            {
                _colorDefinitions.Add(null);
                return;
            }
            static int ToByte(int component) => Math.Min(255, Math.Max(0, (int)Math.Round(
                component * 255d / 1000d,
                MidpointRounding.AwayFromZero)));
            if (!cmyk)
            {
                _colorDefinitions.Add(
                    $"#{ToByte(components[0]):X2}{ToByte(components[1]):X2}{ToByte(components[2]):X2}");
                return;
            }
            var cyan = components[0] / 1000d;
            var magenta = components[1] / 1000d;
            var yellow = components[2] / 1000d;
            var black = components[3] / 1000d;
            _colorDefinitions.Add(
                $"#{Math.Min(255, Math.Max(0, (int)Math.Round(255d * (1d - cyan) * (1d - black)))):X2}"
                + $"{Math.Min(255, Math.Max(0, (int)Math.Round(255d * (1d - magenta) * (1d - black)))):X2}"
                + $"{Math.Min(255, Math.Max(0, (int)Math.Round(255d * (1d - yellow) * (1d - black)))):X2}");
        }

        private void SetCurrentColorIndex(int index)
        {
            if (index < 0 || index > _colorDefinitions.Count)
                throw new InvalidDataException(
                    $"MathType MTEF COLOR index {index} exceeds {_colorDefinitions.Count} definitions.");
            _currentColorIndex = index;
        }

        private string? CurrentMathColor =>
            _currentColorIndex <= 0 || _currentColorIndex > _colorDefinitions.Count
                ? null
                : string.Equals(_colorDefinitions[_currentColorIndex - 1], "#000000",
                    StringComparison.OrdinalIgnoreCase)
                    ? null
                    : _colorDefinitions[_currentColorIndex - 1];

        private static void ApplyMathColor(XElement element, string? color)
        {
            if (color is not null) element.SetAttributeValue("mathcolor", color);
        }

        private static void AppendNode(XElement row, XElement node)
        {
            var previous = row.Elements().LastOrDefault();
            if (node.Name.LocalName == "mtext" && !node.HasElements)
            {
                // MathType stores spaces inside an ordinary text run as its
                // Space typeface, so ReadCharacter() necessarily materializes
                // them first as <mspace width="0.2em">. Rejoin only a run of
                // those exact spaces when it is bounded by mtext on both sides.
                // This preserves `\text{Double word product}` as textual content
                // without treating general mathematical spacing as text.
                var trailingTextSpaces = row.Elements()
                    .Reverse()
                    .TakeWhile(IsMtefTextSpace)
                    .ToArray();
                if (trailingTextSpaces.Length > 0)
                {
                    var precedingText = trailingTextSpaces[trailingTextSpaces.Length - 1].PreviousNode as XElement;
                    if (precedingText is not null
                        && precedingText.Name.LocalName == "mtext"
                        && !precedingText.HasElements
                         && string.Equals((string?)precedingText.Attribute("mathvariant") ?? "normal",
                             (string?)node.Attribute("mathvariant") ?? "normal", StringComparison.OrdinalIgnoreCase)
                         && string.Equals((string?)precedingText.Attribute("mathcolor") ?? string.Empty,
                             (string?)node.Attribute("mathcolor") ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                    {
                        precedingText.Value += new string(' ', trailingTextSpaces.Length) + node.Value;
                        foreach (var space in trailingTextSpaces) space.Remove();
                        return;
                    }
                }
            }

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
                && string.Equals(
                    (string?)previous.Attribute("mathcolor") ?? string.Empty,
                    (string?)node.Attribute("mathcolor") ?? string.Empty,
                    StringComparison.OrdinalIgnoreCase)
                && !previous.HasElements
                && !node.HasElements)
            {
                previous.Value += node.Value;
                return;
            }
            row.Add(node);
        }

        private static bool IsMtefTextSpace(XElement element) =>
            element.Name.LocalName == "mspace"
            && string.Equals(
                (string?)element.Attribute("width"),
                "0.2em",
                StringComparison.OrdinalIgnoreCase)
            && !element.HasElements;

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
            internal bool ScriptPrecedes { get; set; }
            internal string? MathColor { get; set; }
            internal XElement? Node { get; set; }
            internal XElement? Subscript { get; set; }
            internal XElement? Superscript { get; set; }
        }

        private ParsedTemplate ReadTemplate()
        {
            using var structure = EnterStructure();
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
                        ScriptPrecedes = (variation & 1) != 0,
                        Subscript = subscript,
                        Superscript = superscript,
                    };
                }
                case TemplateDirac:
                {
                    var left = ReadNextLineSlot();
                    var right = ReadNextLineSlot();
                    var hasLeft = (variation & 1) != 0;
                    var hasRight = (variation & 2) != 0;
                    var fenceCount = 1 + (hasLeft ? 1 : 0) + (hasRight ? 1 : 0);
                    for (var index = 0; index < fenceCount; index++)
                    {
                        SkipFormattingRecords();
                        _ = ReadCharacter();
                    }
                    ConsumeTemplateEnd();
                    var body = new XElement("mrow");
                    if (hasLeft) body.Add(left.Nodes());
                    if (hasLeft && hasRight) body.Add(new XElement("mo", "|"));
                    if (hasRight) body.Add(right.Nodes());
                    return new ParsedTemplate
                    {
                        Node = new XElement("mfenced",
                            new XAttribute("open", hasLeft ? "⟨" : "|"),
                            new XAttribute("close", hasRight ? "⟩" : "|"),
                            new XAttribute("separators", ""), CollapseRow(body)),
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
                case TemplateInterval:
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
                    XElement? selectedAccent = null;
                    int? selectedAccentTypeface = null;
                    if (selector is TemplateHat or TemplateTilde or TemplateArc)
                    {
                        SkipFormattingRecords();
                        if (Peek() == RecordChar)
                        {
                            selectedAccent = ReadCharacter(out var typeface);
                            selectedAccentTypeface = typeface;
                        }
                    }
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
                    else if (selector is TemplateTilde or TemplateHat or TemplateArc)
                    {
                        // Preserve the real selected HatBox glyph; the selector
                        // names the layout class, not necessarily '^'/'~'. Keep
                        // reading old one-slot exports, but never generate them.
                        var mark = selectedAccent is null
                            ? selector == TemplateTilde ? "~" : selector == TemplateArc ? "⌢" : "^"
                            : NormalizeAccentMark(selectedAccent.Value);
                        var accent = new XElement("mo", mark);
                        if (selectedAccentTypeface == TypefaceFence)
                            accent.SetAttributeValue("stretchy", "true");
                        else if (selectedAccentTypeface is >= TypefaceText and <= TypefaceFarEastText
                            || (selectedAccentTypeface < 0 && selectedAccent?.Attribute("data-mtef-explicit-typeface") is null))
                            accent.SetAttributeValue("stretchy", "false");
                        var variant = (string?)selectedAccent?.Attribute("mathvariant");
                        if (!string.IsNullOrWhiteSpace(variant)) accent.SetAttributeValue("mathvariant", variant);
                        node = new XElement("mover", new XAttribute("accent", "true"), body, accent);
                    }
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
            // The selector supplies semantic operator identity; its serialized
            // glyph still supplies typography. Do not discard FONT_STYLE_DEF
            // merely because a dedicated selector already names the operator.
            var operatorVariant=(string?)serializedOperator?.Attribute("mathvariant");
            if(!string.IsNullOrWhiteSpace(operatorVariant))
                operatorNode.SetAttributeValue("mathvariant",operatorVariant);
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
            XNode WithColor(XNode node)
            {
                if (node is XElement element)
                    ApplyMathColor(element, template.MathColor);
                return node;
            }
            var sub = template.Subscript is null
                ? new XElement("mrow")
                : CollapseRow(template.Subscript);
            var sup = template.Superscript is null
                ? new XElement("mrow")
                : CollapseRow(template.Superscript);
            if (template.ScriptPrecedes)
                return WithColor(new XElement("mmultiscripts", baseNode, new XElement("mprescripts"),
                    template.ScriptKind == TemplateSup ? new XElement("none") : sub,
                    template.ScriptKind == TemplateSub ? new XElement("none") : sup));
            if (baseNode is XElement multiscripts
                && multiscripts.Name.LocalName == "mmultiscripts"
                && multiscripts.Elements().FirstOrDefault(element =>
                    element.Name.LocalName == "mprescripts") is XElement marker)
            {
                marker.AddBeforeSelf(
                    template.ScriptKind == TemplateSup ? new XElement("none") : sub,
                    template.ScriptKind == TemplateSub ? new XElement("none") : sup);
                return WithColor(multiscripts);
            }
            if (template.ScriptKind == TemplateSub)
                return WithColor(new XElement("msub", baseNode, sub));
            if (template.ScriptKind == TemplateSup)
                return WithColor(new XElement("msup", baseNode, sup));
            return WithColor(new XElement("msubsup", baseNode, sub, sup));
        }

        private static XNode BuildPrecedingScriptNode(
            XNode baseNode,
            IReadOnlyList<ParsedTemplate> templates)
        {
            var multiscripts = new XElement("mmultiscripts", baseNode, new XElement("mprescripts"));
            for (var index = templates.Count - 1; index >= 0; index--)
            {
                var template = templates[index];
                var sub = template.Subscript is null
                    ? new XElement("mrow")
                    : CollapseRow(template.Subscript);
                var sup = template.Superscript is null
                    ? new XElement("mrow")
                    : CollapseRow(template.Superscript);
                multiscripts.Add(
                    template.ScriptKind == TemplateSup ? new XElement("none") : sub,
                    template.ScriptKind == TemplateSub ? new XElement("none") : sup);
            }
            return multiscripts;
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
                    SetCurrentColorIndex(ReadUnsigned(_data, ref _position));
                    continue;
                }
                if (record == RecordColorDef)
                {
                    ReadColorDefinitionRecord();
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

        private int[] ReadRulerOffsets()
        {
            // Native MathType writes the inline ruler payload directly after a
            // LINE/PILE with LP_RULER set. Accept an explicit 7 tag as a tolerant
            // compatibility form, but do not require it.
            if (Peek() == RecordRuler) _position++;
            var count = ReadByte();
            var offsets = new int[count];
            for (var index = 0; index < count; index++)
            {
                _ = ReadByte(); // tab type: alignment symbol can override it
                var low = ReadByte();
                var high = ReadByte();
                offsets[index] = low | (high << 8);
            }
            return offsets;
        }

        private void SkipRuler() => _ = ReadRulerOffsets();

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
            if (selector == TemplateInterval)
            {
                const string endpoints = "()[]";
                return (endpoints[variation & 3].ToString(), endpoints[(variation >> 4) & 3].ToString());
            }
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
