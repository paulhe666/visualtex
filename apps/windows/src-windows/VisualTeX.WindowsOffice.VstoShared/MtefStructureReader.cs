using System.Globalization;
using System.Text;
using System.Xml.Linq;
using static VisualTeX.WordVsto.MathTypeMtefCodec;
using static VisualTeX.WordVsto.MathTypeMtefCharacters;
using static VisualTeX.WordVsto.MathTypeMtefGlyphs;
using static VisualTeX.WordVsto.MathTypeMtefPrefix;
using static VisualTeX.WordVsto.MathTypeMtefSignature;
using static VisualTeX.WordVsto.MathTypeMtefTemplates;

namespace VisualTeX.WordVsto;

internal sealed class MtefStructureReader
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
    // Size records are stream state: MathType writes one only where the
    // size changes, and a CHAR in symbol size is a large operator.
    private byte _streamSize = RecordFull;
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
        Skip(1); // vertical alignment
        var horizontalJustification = ReadByte();
        Skip(1); // vertical cell justification
        var rowCount = ReadByte();
        var columnCount = ReadByte();
        if (rowCount == 0 || columnCount == 0)
            throw new InvalidDataException("MathType MATRIX has zero rows or columns.");
        Skip(PartitionByteCount(rowCount + 1));
        Skip(PartitionByteCount(columnCount + 1));
        var table = new XElement("mtable");
        // Left- or right-justified cells (MathType's \begin{cases} is left).
        if (horizontalJustification is 0 or 2)
            table.SetAttributeValue("columnalign", string.Join(" ",
                Enumerable.Repeat(horizontalJustification == 0 ? "left" : "right", columnCount)));
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
                // Styled ASCII punctuation (MathType keeps it in Times) is
                // an operator, as in the unstyled Function face. A period
                // stays with the decimal reconstruction below.
                token = new XElement(text.All(char.IsDigit) ? "mn"
                    : text.Length == 1 && text[0] < 0x80 && text[0] != '.'
                        && (char.IsPunctuation(text[0]) || char.IsSymbol(text[0])) ? "mo"
                    : "mi", text);
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
            && scalar is not (0x220B or 0x2223 or 0x25EF or 0x2661)
            && !IsIntegralSignScalar(scalar))
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
        else if (typeface == TypefaceFarEastText && scalar == 0x2223)
            token = new XElement("mo", "∣");
        else if (typeface == TypefaceFarEastText && scalar == 0x25EF)
            token = new XElement("mo", "◯");
        else if (typeface == TypefaceFarEastText && scalar == 0x2661)
            token = new XElement("mo", "♡");
        else if (typeface == TypefaceFarEastText && IsIntegralSignScalar(scalar))
            // MathType draws the double, triple and contour integral signs
            // from its far-east face.
            token = new XElement("mo", text);
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
            // MT Extra's plain hbar is the slanted \hbar glyph: no style.
            if (scalar == 0x210F && typeface == TypefaceMtExtra && actualIdentifierStyle == 0)
                token.Attribute("mathvariant")?.Remove();
            else
                token.SetAttributeValue("mathvariant", NativeCharacterVariant(actualIdentifierStyle));
            SetPreservedCharacterAttributes(token, typeface, options, scalar, encoded);
            token.SetAttributeValue("data-mtef-character-style", actualIdentifierStyle);
        }

        // MathType stores a lone large operator as its character in symbol
        // size; union, intersection and coproduct then use the glyph of the
        // binary operator (Symbol ∪/∩, MT Extra's ⨿ position).
        if (token.Name.LocalName == "mo" && _streamSize is RecordSym or RecordSubSym)
        {
            var large = token.Value switch { "∪" => "⋃", "∩" => "⋂", "⨿" => "∐", _ => null };
            if (large is not null) token.Value = large;
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
            var strike = EmbellishmentStrike(embellishment);
            if (strike.Length > 0)
            {
                token = new XElement("menclose", new XAttribute("notation", strike), token);
                continue;
            }
            if (embellishment is >= EmbellUnderDot and <= EmbellUnderQuadrupleDot)
            {
                // Dots below the character: one to four centred dots.
                var dots = new XElement("mrow");
                for (var dot = EmbellUnderDot; dot <= embellishment; dot++)
                    dots.Add(new XElement("mo", "⋅"));
                token = new XElement("munder", new XAttribute("accentunder", "true"), token, dots);
                continue;
            }
            var mark = EmbellishmentMark(embellishment);
            if (mark.Length == 0)
                throw new InvalidDataException(
                    $"Unsupported MathType character embellishment {embellishment}; conversion stopped to avoid losing a symbol.");
            token = embellishment == EmbellUnderbar
                ? new XElement(
                    "munder",
                    new XAttribute("accentunder", "true"),
                    token,
                    new XElement("mo", new XAttribute("accent", "true"), mark))
                : embellishment >= EmbellUnderTilde
                ? new XElement(
                    "munder",
                    new XAttribute("accentunder", "true"),
                    token,
                    new XElement("mo", new XAttribute("stretchy", "false"), mark))
                : new XElement(
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
        EmbellTripleDot => "⃛",
        EmbellBreve => "˘",
        EmbellQuadrupleDot => "⃜",
        EmbellUnderbar => "―",
        EmbellRightHarpoon or EmbellUnderRightHarpoon => "⇀",
        EmbellLeftHarpoon or EmbellUnderLeftHarpoon => "↼",
        EmbellFrown or EmbellUnderFrown => "⌢",
        EmbellUnderSmile => "⌣",
        EmbellUnderTilde => "∼",
        EmbellUnderRightArrow => "→",
        EmbellUnderLeftArrow => "←",
        EmbellUnderBothArrow => "↔",
        _ => string.Empty,
    };

    // MathType's slash, cross and mid-bar embellishments strike through
    // the character, the same marks as its strike-through template.
    private static string EmbellishmentStrike(byte embellishment) => embellishment switch
    {
        EmbellNot or EmbellUpBar => "updiagonalstrike",
        EmbellDownBar => "downdiagonalstrike",
        EmbellCrossBars => "updiagonalstrike downdiagonalstrike",
        EmbellMidbar => "horizontalstrike",
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
                var fraction = new XElement(
                    "mfrac",
                    CollapseRow(numerator),
                    CollapseRow(denominator));
                // MathType's small fraction (variation 1) is TeX's \tfrac.
                return new ParsedTemplate
                {
                    Node = (variation & 0x01) != 0
                        ? new XElement("mstyle", new XAttribute("displaystyle", "false"),
                            new XAttribute("scriptlevel", "0"), fraction)
                        : fraction,
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
            case TemplateDoubleBracket:
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
        else if (selector is 21 or 22 && Peek() == RecordLine)
            serializedOperator = ReadIntegralOperatorLine();
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
        // Limits above and below the sign (0x40, or the summation-style
        // integral operator) stay there in every style.
        XNode decorated = ((variation & 0x40) != 0 || selector == 22)
                && (lower.HasElements || upper.HasElements)
            ? DecorateWithLimitsAbove(operatorNode, lower, upper)
            : DecorateWithLimits(operatorNode, lower, upper);
        var row = new XElement("mrow", decorated);
        // The template's operand slot is one group, as it is in MathType.
        // As loose tokens after the operator it lost that ownership, and
        // the OMML projection took none of it as the n-ary body.
        var operand = main.Nodes().ToArray();
        if (main.HasElements)
        {
            foreach (var node in operand) node.Remove();
            row.Add(new XElement("mrow", operand));
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
        if (!lower.HasElements && !upper.HasElements)
            return new ParsedTemplate { Node = new XElement("mrow", baseNode) };
        // MathType's \check, \acute, \grave, \breve and \mathring, and its \dot
        // and \ddot on a group: the accent in the upper slot. The base stays
        // a group, which is what keeps the accent in this template.
        var accentMark = !lower.HasElements
            && upper.Elements().All(token => token.Name.LocalName is "mo" or "mi")
            ? upper.Value switch
            {
                "ˇ" or "´" or "`" or "˘" or "˚" => upper.Value,
                "." => "˙",
                ".." => "¨",
                _ => null,
            }
            : null;
        if (accentMark is not null)
        {
            var group = new XElement("mrow", new XAttribute("data-mjx-texclass", "ORD"));
            group.Add(baseNode is XElement baseRow && baseRow.Name.LocalName == "mrow"
                ? baseRow.Nodes().ToArray()
                : new object[] { baseNode });
            return new ParsedTemplate
            {
                Node = new XElement("mover", new XAttribute("accent", "true"), group,
                    new XElement("mo", new XAttribute("stretchy", "false"), accentMark)),
            };
        }
        // The limits are centred under and over the base in every style.
        return new ParsedTemplate
        {
            Node = lower.HasElements && upper.HasElements
                ? new XElement("munderover", new XAttribute("accent", "false"),
                    new XAttribute("accentunder", "false"), baseNode, CollapseRow(lower), CollapseRow(upper))
                : lower.HasElements
                    ? new XElement("munder", new XAttribute("accentunder", "false"), baseNode, CollapseRow(lower))
                    : new XElement("mover", new XAttribute("accent", "false"), baseNode, CollapseRow(upper)),
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

    // The integral-operator templates (tmINTOP, tmSUMOP) keep the sign in
    // their own LINE slot: one far-east sign (∬, ∮) as MathType writes
    // it, or MathType's template glyphs (n ∫, after the loop adornment for
    // a contour integral).
    private XElement ReadIntegralOperatorLine()
    {
        var line = ReadLine();
        var text = line.Value;
        const char loopAdornment = (char)0xEE11; // MT Extra's contour-integral loop
        var contour = text.IndexOf(loopAdornment) >= 0;
        var signs = text.Replace(loopAdornment.ToString(), string.Empty);
        var sign = signs switch
        {
            "∫" => contour ? "∮" : "∫",
            "∫∫" => contour ? "∯" : "∬",
            "∫∫∫" => contour ? "∰" : "∭",
            // These two selectors do not name their operator: a sum, a
            // product or any other sign the user chose is in this line.
            _ when signs.Length == 1
                   && signs[0] is not (>= (char)0xE000 and <= (char)0xF8FF) and not (char)0xFFFD
                   && !char.IsWhiteSpace(signs[0]) => signs,
            _ => "∫",
        };
        var signNode = new XElement("mo", sign);
        var variant = line.Descendants()
            .Select(e => (string?)e.Attribute("mathvariant"))
            .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
        if (variant is not null) signNode.SetAttributeValue("mathvariant", variant);
        return signNode;
    }

    private static bool IsIntegralSignScalar(int scalar) => scalar is >= 0x222B and <= 0x2233;

    private static XNode DecorateWithLimitsAbove(
        XElement operatorNode,
        XElement lower,
        XElement upper)
    {
        // Limits, not accents, whatever the limit's glyph.
        operatorNode.SetAttributeValue("movablelimits", "false");
        if (lower.HasElements && upper.HasElements)
            return new XElement("munderover", new XAttribute("accent", "false"),
                new XAttribute("accentunder", "false"), operatorNode, CollapseRow(lower), CollapseRow(upper));
        return lower.HasElements
            ? new XElement("munder", new XAttribute("accentunder", "false"), operatorNode, CollapseRow(lower))
            : new XElement("mover", new XAttribute("accent", "false"), operatorNode, CollapseRow(upper));
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
        // Count in the low two bits, loop kind above them (MTEF tvINT_*):
        // 4 a loop, 8 a clockwise loop, 12 a counter-clockwise one.
        0x06 => "∯",
        0x07 => "∰",
        0x08 or 0x09 => "∲",
        0x0C or 0x0D => "∳",
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
                _streamSize = record;
                _position++;
                continue;
            }
            if (record == RecordSize)
            {
                _streamSize = RecordSize;
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
            {
                // MathType's MathML import colours each embellishment: a
                // COLOR record (stream state, like any formatting record)
                // precedes its EMBELL inside the list.
                var before = _position;
                SkipFormattingRecords();
                if (_position != before) continue;
                throw new InvalidDataException(
                    $"Unexpected MTEF embellishment record {record} at offset {_position}.");
            }
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
            TemplateDoubleBracket => ("⟦", "⟧"),
            _ => (string.Empty, string.Empty),
        };
        return (presentLeft ? pair.Item1 : string.Empty, presentRight ? pair.Item2 : string.Empty);
    }
}
