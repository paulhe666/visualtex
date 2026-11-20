using System.Security.Cryptography;
using System.Text;

namespace VisualTeX.WordVsto;

/// <summary>
/// A digest of an OLE compound file's content: the root CLSID and every storage
/// and stream (name, type, stream body) in name order. The container layout
/// (sector order, free space, timestamps) is not content: WPS writes the same
/// embedded object with different bytes after an undo. Pure managed reader of
/// the [MS-CFB] structures; null for bytes that are not a readable compound file.
/// </summary>
internal static class OleCompoundFileDigest
{
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FreeSector = 0xFFFFFFFF;
    private const int MaxEntries = 4096;

    internal static string? TryCompute(byte[] file)
    {
        try { return Compute(file); }
        catch (Exception error) when (error is InvalidDataException or ArgumentException or IndexOutOfRangeException or OverflowException)
        {
            return null;
        }
    }

    private static string? Compute(byte[] file)
    {
        if (file.Length < 512) return null;
        if (BitConverter.ToUInt64(file, 0) != 0xE11AB1A1E011CFD0UL) return null;
        var sectorShift = BitConverter.ToUInt16(file, 0x1E);
        var miniShift = BitConverter.ToUInt16(file, 0x20);
        if (sectorShift is not (9 or 12) || miniShift != 6) return null;
        var sectorSize = 1 << sectorShift;
        var fatSectorCount = BitConverter.ToUInt32(file, 0x2C);
        var firstDirectory = BitConverter.ToUInt32(file, 0x30);
        var miniCutoff = BitConverter.ToUInt32(file, 0x38);
        var firstMiniFat = BitConverter.ToUInt32(file, 0x3C);
        var firstDifat = BitConverter.ToUInt32(file, 0x44);

        byte[] Sector(uint index)
        {
            var offset = checked(((long)index + 1) * sectorSize);
            if (offset + sectorSize > file.Length) throw new InvalidDataException("Compound file sector out of range.");
            var data = new byte[sectorSize];
            Buffer.BlockCopy(file, (int)offset, data, 0, sectorSize);
            return data;
        }

        // FAT sector list: 109 entries in the header, then the DIFAT chain.
        var fatSectors = new List<uint>();
        for (var i = 0; i < 109 && fatSectors.Count < fatSectorCount; i++)
            fatSectors.Add(BitConverter.ToUInt32(file, 0x4C + i * 4));
        var difat = firstDifat;
        var guard = 0;
        while (fatSectors.Count < fatSectorCount && difat != EndOfChain && difat != FreeSector)
        {
            if (++guard > 1 << 16) throw new InvalidDataException("Compound file DIFAT loops.");
            var data = Sector(difat);
            var per = sectorSize / 4 - 1;
            for (var i = 0; i < per && fatSectors.Count < fatSectorCount; i++)
                fatSectors.Add(BitConverter.ToUInt32(data, i * 4));
            difat = BitConverter.ToUInt32(data, per * 4);
        }
        var fat = new List<uint>();
        foreach (var sector in fatSectors)
        {
            var data = Sector(sector);
            for (var i = 0; i < sectorSize / 4; i++) fat.Add(BitConverter.ToUInt32(data, i * 4));
        }

        byte[] Chain(uint start, IReadOnlyList<uint> table, Func<uint, byte[]> read, int unit, long length)
        {
            using var output = new MemoryStream();
            var current = start;
            var steps = 0;
            while (current != EndOfChain && current != FreeSector && (length < 0 || output.Length < length))
            {
                if (current >= table.Count || ++steps > table.Count) throw new InvalidDataException("Compound file chain is broken.");
                var data = read(current);
                output.Write(data, 0, unit);
                current = table[(int)current];
            }
            var bytes = output.ToArray();
            if (length >= 0)
            {
                if (bytes.Length < length) throw new InvalidDataException("Compound file stream is truncated.");
                Array.Resize(ref bytes, (int)length);
            }
            return bytes;
        }

        var directory = Chain(firstDirectory, fat, Sector, sectorSize, -1);
        var count = Math.Min(directory.Length / 128, MaxEntries);
        if (count == 0) return null;
        (string Name, int Type, uint Left, uint Right, uint Child, byte[] Clsid, uint Start, long Size) Entry(uint index)
        {
            if (index >= count) throw new InvalidDataException("Compound file directory index out of range.");
            var o = (int)index * 128;
            var nameLength = Math.Min((int)BitConverter.ToUInt16(directory, o + 64), 64);
            var name = nameLength >= 2 ? Encoding.Unicode.GetString(directory, o, nameLength - 2) : string.Empty;
            var clsid = new byte[16];
            Buffer.BlockCopy(directory, o + 80, clsid, 0, 16);
            var size = sectorShift == 9 ? BitConverter.ToUInt32(directory, o + 120) : (long)BitConverter.ToUInt64(directory, o + 120);
            return (name, directory[o + 66], BitConverter.ToUInt32(directory, o + 68), BitConverter.ToUInt32(directory, o + 72),
                BitConverter.ToUInt32(directory, o + 76), clsid, BitConverter.ToUInt32(directory, o + 116), size);
        }

        var root = Entry(0);
        if (root.Type != 5) return null;
        var miniStream = root.Start == EndOfChain ? Array.Empty<byte>() : Chain(root.Start, fat, Sector, sectorSize, root.Size);
        var miniFatBytes = firstMiniFat == EndOfChain ? Array.Empty<byte>() : Chain(firstMiniFat, fat, Sector, sectorSize, -1);
        var miniFat = new List<uint>();
        for (var i = 0; i + 4 <= miniFatBytes.Length; i += 4) miniFat.Add(BitConverter.ToUInt32(miniFatBytes, i));
        byte[] MiniSector(uint index)
        {
            var offset = checked((long)index * 64);
            if (offset + 64 > miniStream.Length) throw new InvalidDataException("Compound file mini sector out of range.");
            var data = new byte[64];
            Buffer.BlockCopy(miniStream, (int)offset, data, 0, 64);
            return data;
        }

        using var digest = new MemoryStream();
        digest.Write(root.Clsid, 0, 16);
        var visited = new HashSet<uint>();
        void Children(uint child, List<uint> into)
        {
            if (child == FreeSector) return;
            if (!visited.Add(child)) throw new InvalidDataException("Compound file directory loops.");
            var entry = Entry(child);
            Children(entry.Left, into);
            into.Add(child);
            Children(entry.Right, into);
        }
        void Append(uint storage, string prefix, int depth)
        {
            if (depth > 16) throw new InvalidDataException("Compound file storage is nested too deeply.");
            var children = new List<uint>();
            Children(Entry(storage).Child, children);
            foreach (var index in children.OrderBy(i => Entry(i).Name, StringComparer.Ordinal))
            {
                var entry = Entry(index);
                var path = prefix + "/" + entry.Name;
                var header = Encoding.UTF8.GetBytes(path + "|" + entry.Type + "|");
                digest.Write(header, 0, header.Length);
                if (entry.Type == 2)
                {
                    var data = entry.Size == 0 ? Array.Empty<byte>()
                        : entry.Size < miniCutoff ? Chain(entry.Start, miniFat, MiniSector, 64, entry.Size)
                        : Chain(entry.Start, fat, Sector, sectorSize, entry.Size);
                    var length = BitConverter.GetBytes((long)data.Length);
                    digest.Write(length, 0, length.Length);
                    digest.Write(data, 0, data.Length);
                }
                else if (entry.Type == 1)
                {
                    digest.Write(entry.Clsid, 0, 16);
                    Append(index, path, depth + 1);
                }
            }
        }
        Append(0, string.Empty, 0);
        using var sha = SHA256.Create();
        return Convert.ToBase64String(sha.ComputeHash(digest.ToArray()));
    }
}
