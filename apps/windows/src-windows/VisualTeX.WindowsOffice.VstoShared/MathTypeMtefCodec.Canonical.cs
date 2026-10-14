namespace VisualTeX.WordVsto;

// MathType stores the same equation the same way every time. These rules were
// read off MathType 7.9's own MTEF (its TeX and MathML import, the paths its
// Word commands use) for ~2000 formulas, and are applied to every root
// VisualTeX writes so that MathType's renderer and editor see the objects
// MathType itself would have stored:
//
// - Size records are stream state, written lazily: a SIZE/FULL/SUB/SYM record
//   appears only where the next slot or object needs a size different from the
//   last size record in the stream (never at the end of a line, never twice).
// - A script, limit or accent attaches to the preceding object, so MathType
//   never wraps a single base object in its own LINE. A wrapper changes where
//   MathType places the script (by 1-3 pt) and is not something its editor
//   produces. The same holds for a base that is one native run of characters
//   (a function name, a number). The wrapper is kept only where MathType would
//   otherwise merge the base with a neighbouring character into one run (the
//   script would then belong to the whole run, which is a different formula).
//
// Both rules keep every object's effective size and every object's place in
// the tree, so meaning and the reader's MathML are unchanged.
internal static partial class MathTypeMtefCodec
{
    private sealed class MtefNode
    {
        internal MtefNode(byte type, byte[] head, List<MtefNode>? children)
        {
            Type = type;
            Head = head;
            Children = children;
        }

        internal byte Type { get; }
        // Record bytes without children and without the closing END.
        internal byte[] Head { get; }
        // Object list of a LINE/TMPL/PILE/MATRIX (null for leaf records).
        internal List<MtefNode>? Children { get; set; }
        internal byte[]? EffectiveSize { get; set; }

        internal bool IsObject => Type is RecordLine or RecordChar or RecordTemplate or RecordPile or RecordMatrix;
        internal bool IsSize => Type == RecordSize || Type >= RecordFull && Type <= RecordSubSym;
    }

    internal static byte[] CanonicalizeRootStructure(byte[] root)
    {
        var position = 0;
        var nodes = ParseMtefObjectList(root, ref position);
        if (position != root.Length)
            throw new InvalidDataException($"Unexpected bytes after the MathType root structure at {position}.");
        FlattenSingleObjectLines(nodes, insideLine: false);
        WriteSizesLazily(nodes);
        var output = new List<byte>(root.Length);
        WriteMtefNodes(nodes, output);
        output.Add(RecordEnd);
        return output.ToArray();
    }

    private static void FlattenSingleObjectLines(List<MtefNode> nodes, bool insideLine)
    {
        foreach (var node in nodes)
        {
            if (node.Children is not null)
                FlattenSingleObjectLines(node.Children, node.Type == RecordLine);
        }
        if (!insideLine) return;
        for (var index = 0; index < nodes.Count; index++)
        {
            var wrapper = nodes[index];
            if (wrapper.Type != RecordLine || wrapper.Children is null
                || wrapper.Head.Length != 2 || wrapper.Head[1] != 0)
                continue;
            var objects = wrapper.Children.Where(child => child.IsObject).ToArray();
            if (objects.Length == 0
                || objects.Length > 1 && !IsSingleNativeRun(objects)
                || wrapper.Children.Any(child => !child.IsObject && !child.IsSize))
                continue;
            // A LINE holding only a script is that script's empty base; without
            // it the script would attach to whatever precedes the LINE, and at
            // the start of a line it would read as a prescript of what follows
            // (MathType's own form of _{92}^{238}U). Alone in its line it needs
            // no wrapper, and MathType writes none.
            if (IsTrailingScriptTemplate(objects[0])
                && nodes.Where((sibling, at) => at != index && sibling.IsObject).Any())
                continue;
            if (WouldJoinPrecedingRun(nodes, index, objects[0])
                || WouldJoinFollowingRun(nodes, index, objects[objects.Length - 1]))
                continue;
            // A preceding script belongs to the next object only, not to a run.
            if (objects.Length > 1 && FollowsPrecedingScript(nodes, index))
                continue;
            nodes.RemoveAt(index);
            nodes.InsertRange(index, wrapper.Children);
            index += wrapper.Children.Count - 1;
        }
    }

    private static bool IsTrailingScriptTemplate(MtefNode node)
    {
        if (node.Type != RecordTemplate) return false;
        var position = 2;
        if ((node.Head[1] & 0x08) != 0)
            position += node.Head[position] == 128 && node.Head[position + 1] == 128 ? 6 : 2;
        return node.Head[position] is TemplateSub or TemplateSup or TemplateSubSup;
    }

    // MathType merges adjacent characters of one text-like typeface (Text,
    // Function, Number, the far-east face, or one explicit font style) into a
    // single object; Variable, Greek and Symbol characters stay separate.
    private static bool WouldJoinPrecedingRun(List<MtefNode> siblings, int index, MtefNode inner)
    {
        if (inner.Type != RecordChar) return false;
        MtefNode? previous = null;
        for (var cursor = index - 1; cursor >= 0; cursor--)
        {
            if (!siblings[cursor].IsObject) continue;
            previous = siblings[cursor];
            break;
        }
        if (previous is null || previous.Type != RecordChar) return false;
        var typeface = ReadCharTypeface(inner.Head);
        return typeface == ReadCharTypeface(previous.Head)
            && typeface is not (TypefaceVariable or TypefaceLowerGreek or TypefaceUpperGreek or TypefaceSymbol);
    }

    private static bool FollowsPrecedingScript(List<MtefNode> siblings, int index)
    {
        var previous = siblings.Take(index).LastOrDefault(sibling => sibling.IsObject);
        if (previous is null || !IsTrailingScriptTemplate(previous)) return false;
        var position = 2;
        if ((previous.Head[1] & 0x08) != 0)
            position += previous.Head[position] == 128 && previous.Head[position + 1] == 128 ? 6 : 2;
        return (previous.Head[position + 1] & 0x01) != 0; // variation: precedes its base
    }

    private static bool WouldJoinFollowingRun(List<MtefNode> siblings, int index, MtefNode inner)
    {
        if (inner.Type != RecordChar) return false;
        var next = siblings.Skip(index + 1).FirstOrDefault(sibling => sibling.IsObject);
        if (next is null || next.Type != RecordChar) return false;
        var typeface = ReadCharTypeface(inner.Head);
        return typeface == ReadCharTypeface(next.Head)
            && typeface is not (TypefaceVariable or TypefaceLowerGreek or TypefaceUpperGreek or TypefaceSymbol);
    }

    // Several CHARs MathType keeps as one run, which takes a script as one
    // object: a function name (Function face, from its funcstart character) or
    // a number (Number face).
    private static bool IsSingleNativeRun(MtefNode[] objects)
    {
        if (objects.Any(node => node.Type != RecordChar || (node.Head[1] & CharHasEmbellishment) != 0))
            return false;
        var typeface = ReadCharTypeface(objects[0].Head);
        if (objects.Any(node => ReadCharTypeface(node.Head) != typeface)) return false;
        if (typeface == TypefaceNumber) return true;
        return typeface == TypefaceFunction
            && (objects[0].Head[1] & CharFunctionStart) != 0
            && objects.Skip(1).All(node => (node.Head[1] & CharFunctionStart) == 0);
    }

    private static int ReadCharTypeface(byte[] head)
    {
        var position = 2; // record type + options
        if ((head[1] & 0x08) != 0)
            position += head[position] == 128 && head[position + 1] == 128 ? 6 : 2;
        var first = head[position];
        return first != 0xFF ? first - 128 : (head[position + 1] | head[position + 2] << 8) - 32768;
    }

    private static void WriteSizesLazily(List<MtefNode> nodes)
    {
        // The writer emits sizes eagerly (on entering and leaving a scope). Replay
        // that stream to learn each object's size, then keep only the records
        // MathType writes: before a slot/object whose size differs from the last
        // size record in the stream.
        var full = new[] { RecordFull };
        var state = full;
        void Annotate(List<MtefNode> list)
        {
            foreach (var node in list)
            {
                if (node.IsSize) { state = node.Head; continue; }
                if (node.IsObject) node.EffectiveSize = state;
                if (node.Children is not null) Annotate(node.Children);
            }
        }
        Annotate(nodes);

        var stream = full;
        List<MtefNode> Rebuild(List<MtefNode> list)
        {
            var rebuilt = new List<MtefNode>(list.Count);
            foreach (var node in list)
            {
                if (node.IsSize) continue;
                if (node.IsObject && !SameBytes(node.EffectiveSize!, stream))
                {
                    rebuilt.Add(new MtefNode(node.EffectiveSize![0], node.EffectiveSize, children: null));
                    stream = node.EffectiveSize;
                }
                rebuilt.Add(node);
                if (node.Children is not null) node.Children = Rebuild(node.Children);
            }
            return rebuilt;
        }
        var result = Rebuild(nodes);
        nodes.Clear();
        nodes.AddRange(result);
    }

    private static bool SameBytes(byte[] left, byte[] right)
    {
        if (left.Length != right.Length) return false;
        for (var index = 0; index < left.Length; index++)
            if (left[index] != right[index]) return false;
        return true;
    }

    private static byte[] SliceBytes(byte[] data, int start, int end)
    {
        Require(data, start, end - start);
        var slice = new byte[end - start];
        Buffer.BlockCopy(data, start, slice, 0, slice.Length);
        return slice;
    }

    private static void WriteMtefNodes(List<MtefNode> nodes, List<byte> output)
    {
        foreach (var node in nodes)
        {
            output.AddRange(node.Head);
            if (node.Children is null) continue;
            WriteMtefNodes(node.Children, output);
            output.Add(RecordEnd);
        }
    }

    // Objects and formatting records up to (and consuming) the END of the list.
    private static List<MtefNode> ParseMtefObjectList(byte[] data, ref int position)
    {
        var nodes = new List<MtefNode>();
        while (true)
        {
            Require(data, position, 1);
            if (data[position] == RecordEnd)
            {
                position++;
                return nodes;
            }
            nodes.Add(ParseMtefNode(data, ref position));
        }
    }

    private static MtefNode ParseMtefNode(byte[] data, ref int position)
    {
        var start = position;
        var type = data[position++];
        List<MtefNode>? children = null;
        switch (type)
        {
            case RecordLine:
            {
                var options = data[position++];
                SkipMtefNudge(data, ref position, options);
                if ((options & 0x04) != 0) position += 2;
                if ((options & 0x02) != 0) SkipMtefRuler(data, ref position);
                var head = SliceBytes(data, start, position);
                if ((options & LineNull) == 0) children = ParseMtefObjectList(data, ref position);
                return new MtefNode(type, head, children);
            }
            case RecordChar:
            {
                var options = data[position++];
                SkipMtefNudge(data, ref position, options);
                _ = ReadUnsigned(data, ref position); // typeface
                if ((options & CharNoMtCode) == 0) position += 2;
                if ((options & CharEncoded8) != 0) position++;
                if ((options & CharEncoded16) != 0) position += 2;
                if ((options & CharHasEmbellishment) != 0)
                {
                    while (data[position] == RecordEmbellishment)
                    {
                        position++;
                        var embellishmentOptions = data[position++];
                        SkipMtefNudge(data, ref position, embellishmentOptions);
                        position++;
                    }
                    if (data[position++] != RecordEnd)
                        throw new InvalidDataException($"Unterminated MTEF embellishment list before {position}.");
                }
                break;
            }
            case RecordTemplate:
            {
                var options = data[position++];
                SkipMtefNudge(data, ref position, options);
                position++; // selector
                if ((data[position++] & 0x80) != 0) position++; // variation
                position++; // template-specific options
                var head = SliceBytes(data, start, position);
                children = ParseMtefObjectList(data, ref position);
                return new MtefNode(type, head, children);
            }
            case RecordPile:
            {
                var options = data[position++];
                SkipMtefNudge(data, ref position, options);
                position += 2;
                if ((options & 0x02) != 0) SkipMtefRuler(data, ref position);
                var head = SliceBytes(data, start, position);
                children = ParseMtefObjectList(data, ref position);
                return new MtefNode(type, head, children);
            }
            case RecordMatrix:
            {
                var options = data[position++];
                SkipMtefNudge(data, ref position, options);
                position += 3;
                var rows = data[position++];
                var columns = data[position++];
                position += PartitionByteCount(rows + 1) + PartitionByteCount(columns + 1);
                var head = SliceBytes(data, start, position);
                children = ParseMtefObjectList(data, ref position);
                return new MtefNode(type, head, children);
            }
            case RecordEmbellishment:
            {
                var options = data[position++];
                SkipMtefNudge(data, ref position, options);
                position++;
                break;
            }
            case RecordRuler:
                position--;
                SkipMtefRuler(data, ref position);
                break;
            case RecordFontStyleDef:
                _ = ReadUnsigned(data, ref position);
                position++;
                break;
            case RecordSize:
            {
                var form = data[position++];
                position += form == 101 ? 2 : form == 100 ? 3 : 1;
                break;
            }
            case >= RecordFull and <= RecordSubSym:
                break;
            case RecordColor:
                _ = ReadUnsigned(data, ref position);
                break;
            case RecordColorDef:
                position = SkipColorDefinition(data, start);
                break;
            case RecordFontDef:
                _ = ReadUnsigned(data, ref position);
                position = SkipNullTerminated(data, position);
                break;
            case RecordEncodingDef:
                position = SkipNullTerminated(data, position);
                break;
            case >= 100:
                position = SkipFutureRecord(data, start);
                break;
            default:
                throw new InvalidDataException($"Unexpected MTEF record {type} at offset {start}.");
        }
        Require(data, start, position - start);
        return new MtefNode(type, SliceBytes(data, start, position), children);
    }

    private static void SkipMtefRuler(byte[] data, ref int position)
    {
        // MathType writes the ruler payload directly after LINE/PILE options; an
        // explicit RULER tag is accepted as well.
        if (data[position] == RecordRuler) position++;
        var stops = data[position++];
        position += stops * 3;
    }
}
