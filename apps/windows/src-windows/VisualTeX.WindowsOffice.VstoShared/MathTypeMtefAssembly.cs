using System.Globalization;
using System.Xml.Linq;
using static VisualTeX.WordVsto.MathTypeMtefCodec;
using static VisualTeX.WordVsto.MathTypeMtefCanonicalRoot;
using static VisualTeX.WordVsto.MathTypeMtefCharacters;
using static VisualTeX.WordVsto.MathTypeMtefFences;
using static VisualTeX.WordVsto.MathTypeMtefGlyphs;
using static VisualTeX.WordVsto.MathTypeMtefPrefix;
using static VisualTeX.WordVsto.MathTypeMtefTemplates;

namespace VisualTeX.WordVsto;

/// <summary>
/// Writing a new MTEF stream: MathML prepared for MathType, the color plan, the emission buffer, the root structure
/// and the stream assembled around the kept prefix.
/// </summary>
internal static class MathTypeMtefAssembly
{
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
    internal sealed class MtefWriter : List<byte>
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

    internal static byte[] BuildRootStructure(
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
            // A single-line source keeps its own ENCODING/FONT/FONT_STYLE/COLOR
            // definitions at the start of its root LINE (VisualTeX writes them
            // there). They are global and already counted by the index
            // allocation below, but a root PILE cannot carry them: move them,
            // in their order, into the prefix ahead of the new definitions, or
            // every later style index points one definition too far (an upper
            // Greek Symbol glyph then reads back as its Latin letter).
            foreach (var record in rootFormattingRecords.Where(record => record.Type
                         is RecordEncodingDef or RecordFontDef or RecordFontStyleDef or RecordColorDef))
                for (var index = record.Start; index < record.End; index++)
                    alignedPrefixDefinitions.Add(sourceMtef[index]);
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
                return CanonicalizeRootStructure(alignedRoot.ToArray());
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
        return CanonicalizeRootStructure(output.ToArray());
    }

    internal static byte[] AssembleRewrittenMtef(
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

    internal static void CopyRootLeadingFormattingRecords(
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

    internal static void CopyRootLeadingFormattingRecords(
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

    internal static IReadOnlyList<(byte Type, int Start, int End)>
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

    internal static void SkipMtefNudge(
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
}
