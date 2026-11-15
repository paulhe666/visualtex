using static VisualTeX.WordVsto.MathTypeMtefCodec;

namespace VisualTeX.WordVsto;

/// <summary>
/// The MTEF prefix in front of the root equation: its record layout, where the root structure starts, and skipping
/// its records.
/// </summary>
internal static class MathTypeMtefPrefix
{
    internal static int FindRootStructureOffset(byte[] mtef)
        => ReadPrefixLayout(mtef).Root;

    internal sealed class MtefPrefixLayout
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
    internal static MtefPrefixLayout ReadPrefixLayout(byte[] mtef)
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

    internal static int SkipDimensionArray(byte[] data, int position)
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

    internal static int SkipColorDefinition(byte[] data, int position)
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

    internal static int SkipFutureRecord(byte[] data, int position)
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

    internal static int SkipInitialSizeRecord(byte[] data, int position)
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

    internal static int SkipNullTerminated(byte[] data, int position)
    {
        while (position < data.Length && data[position] != 0) position++;
        if (position >= data.Length)
            throw new EndOfStreamException("MTEF null-terminated field is truncated.");
        return position + 1;
    }
}
