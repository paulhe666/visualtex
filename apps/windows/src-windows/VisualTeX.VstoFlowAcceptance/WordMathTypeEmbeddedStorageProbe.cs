using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeEmbeddedStorageProbe(string artifactRoot)
    {
        var sourceOverride = Environment.GetEnvironmentVariable(
            "VISUALTEX_MATHTYPE_STORAGE_PROBE_SOURCE");
        var source = !string.IsNullOrWhiteSpace(sourceOverride)
            ? Path.GetFullPath(sourceOverride!)
            : Path.GetFullPath(Path.Combine(
                artifactRoot,
                "..",
                "mathtype-ole-to-visualtex",
                "VisualTeX-MathType7-Source-OLE.docx"));
        if (!File.Exists(source))
        {
            source = Path.GetFullPath(Path.Combine(
                AppContext.BaseDirectory,
                "..", "..", "..", "..", "..", "..",
                "artifacts", "mathtype-ole-to-visualtex",
                "VisualTeX-MathType7-Source-OLE.docx"));
        }
        if (!File.Exists(source))
            throw new FileNotFoundException("MathType source OLE fixture is missing.", source);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        object? clipboardComObject = null;
        IEnumFORMATETC? enumerator = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(source, ReadOnly: true, Visible: false);
            if (document.InlineShapes.Count != 1)
                throw new InvalidDataException("MathType source fixture must contain exactly one inline shape.");
            shape = document.InlineShapes[1];
            if (!VisualTeX.WordVsto.MathTypeOleInterop.IsMathTypeOle(shape))
                throw new InvalidDataException("Fixture is no longer recognized as MathType OLE.");

            shape.Range.Copy();
            Thread.Sleep(500);

            var hr = OleGetClipboard(out var clipboardPointer);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);
            if (clipboardPointer == IntPtr.Zero)
                throw new InvalidOperationException("OleGetClipboard returned no IDataObject.");
            try
            {
                clipboardComObject = Marshal.GetObjectForIUnknown(clipboardPointer);
            }
            finally { Marshal.Release(clipboardPointer); }
            if (clipboardComObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                throw new InvalidOperationException("OLE clipboard object does not implement IDataObject.");

            var embeddedObjectId = RegisterClipboardFormatW("Embedded Object");
            if (embeddedObjectId == 0 || embeddedObjectId > ushort.MaxValue)
                throw new InvalidOperationException("Could not resolve CFSTR_EMBEDDEDOBJECT.");

            enumerator = dataObject.EnumFormatEtc(DATADIR.DATADIR_GET);
            var formats = new FORMATETC[1];
            var fetched = new int[1];
            var foundEmbeddedObject = false;
            while (enumerator.Next(1, formats, fetched) == 0 && fetched[0] == 1)
            {
                var format = formats[0];
                var id = unchecked((ushort)format.cfFormat);
                var name = ClipboardFormatName(id);
                Console.WriteLine(
                    $"  clipboard format id={id} name='{name}' tymed={format.tymed} aspect={format.dwAspect} lindex={format.lindex}");
                if (id == embeddedObjectId)
                {
                    foundEmbeddedObject = true;
                    Console.WriteLine($"  Embedded Object advertised tymed={format.tymed}.");
                    if ((format.tymed & (TYMED.TYMED_ISTREAM | TYMED.TYMED_HGLOBAL)) == 0)
                        throw new InvalidDataException(
                            $"Word Embedded Object is not exposed as a serializable stream/global-memory payload; actual={format.tymed}.");
                }
            }
            if (!foundEmbeddedObject)
                throw new InvalidDataException("Word OLE clipboard did not advertise Embedded Object.");

            var request = new FORMATETC
            {
                cfFormat = unchecked((short)embeddedObjectId),
                dwAspect = DVASPECT.DVASPECT_CONTENT,
                lindex = -1,
                ptd = IntPtr.Zero,
                tymed = TYMED.TYMED_ISTREAM,
            };
            dataObject.GetData(ref request, out var medium);
            object? streamObject = null;
            try
            {
                if (medium.tymed != TYMED.TYMED_ISTREAM || medium.unionmember == IntPtr.Zero)
                    throw new InvalidDataException(
                        $"Embedded Object GetData returned {medium.tymed}, ptr=0x{medium.unionmember.ToInt64():X}.");
                streamObject = Marshal.GetObjectForIUnknown(medium.unionmember);
                if (streamObject is not IStream stream)
                    throw new InvalidDataException("Embedded Object TYMED_ISTREAM is not an IStream.");
                var compoundBytes = ReadAllComStream(stream);
                var header = compoundBytes.Take(32).ToArray();
                Console.WriteLine(
                    "  Embedded Object first bytes=" + BitConverter.ToString(header));
                var compound = compoundBytes.Length >= 8
                    && compoundBytes[0] == 0xD0 && compoundBytes[1] == 0xCF
                    && compoundBytes[2] == 0x11 && compoundBytes[3] == 0xE0
                    && compoundBytes[4] == 0xA1 && compoundBytes[5] == 0xB1
                    && compoundBytes[6] == 0x1A && compoundBytes[7] == 0xE1;
                if (!compound)
                    throw new InvalidDataException(
                        "Word Embedded Object stream is not a serialized Compound File Binary object.");
                var compoundPath = Path.Combine(artifactRoot, "mathtype-embedded-object.cfb");
                File.WriteAllBytes(compoundPath, compoundBytes);
                InspectMathTypeCompoundFile(compoundPath);
                Console.WriteLine(
                    $"Embedded Object GetData returned serialized Compound File IStream ptr=0x{medium.unionmember.ToInt64():X} without activating MathType.");
            }
            finally
            {
                Release(streamObject);
                ReleaseStgMedium(ref medium);
            }

            ProbeMathTypeEnhancedMetafile(dataObject, artifactRoot);
            ProbeMathTypeWindowsMetafile(dataObject, artifactRoot);

            Console.WriteLine(
                "MathType embedded-storage probe passed: Word exposes Equation.DSMT4 as a serialized Compound File through CFSTR_EMBEDDEDOBJECT/TYMED_ISTREAM, so VisualTeX can clone/rewrite/reinsert the OLE storage without a MathType server.");
        }
        finally
        {
            Release(enumerator);
            Release(clipboardComObject);
            Release(shape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MetafilePictProbe
    {
        public int MappingMode;
        public int XExt;
        public int YExt;
        public IntPtr Metafile;
    }

    private static void ProbeMathTypeWindowsMetafile(
        System.Runtime.InteropServices.ComTypes.IDataObject dataObject,
        string artifactRoot)
    {
        var request = new FORMATETC
        {
            cfFormat = 3, // CF_METAFILEPICT
            dwAspect = DVASPECT.DVASPECT_CONTENT,
            lindex = -1,
            ptd = IntPtr.Zero,
            tymed = TYMED.TYMED_MFPICT,
        };
        dataObject.GetData(ref request, out var medium);
        try
        {
            if (medium.tymed != TYMED.TYMED_MFPICT || medium.unionmember == IntPtr.Zero)
                throw new InvalidDataException(
                    $"MathType preview did not expose CF_METAFILEPICT; actual={medium.tymed}.");
            var locked = GlobalLock(medium.unionmember);
            if (locked == IntPtr.Zero)
                throw new InvalidOperationException("Could not lock MathType METAFILEPICT.");
            MetafilePictProbe pict;
            try { pict = Marshal.PtrToStructure<MetafilePictProbe>(locked); }
            finally { GlobalUnlock(medium.unionmember); }
            if (pict.Metafile == IntPtr.Zero)
                throw new InvalidDataException("MathType METAFILEPICT contains no HMETAFILE.");
            var byteCount = GetMetaFileBitsEx(pict.Metafile, 0, null);
            if (byteCount == 0 || byteCount > 32 * 1024 * 1024)
                throw new InvalidDataException($"Unexpected MathType WMF size: {byteCount}.");
            var bytes = new byte[byteCount];
            if (GetMetaFileBitsEx(pict.Metafile, byteCount, bytes) != byteCount)
                throw new InvalidDataException("GetMetaFileBitsEx returned an incomplete MathType WMF.");
            File.WriteAllBytes(Path.Combine(artifactRoot, "mathtype-preview.wmf"), bytes);
            var ascii = System.Text.Encoding.ASCII.GetString(bytes);
            var unicode = System.Text.Encoding.Unicode.GetString(bytes);
            var signals = new[] { "MathML", "<math", "mml:math", "MathType EF", "DSMT7" };
            foreach (var signal in signals)
            {
                var found = ascii.IndexOf(signal, StringComparison.OrdinalIgnoreCase) >= 0
                    || unicode.IndexOf(signal, StringComparison.OrdinalIgnoreCase) >= 0;
                Console.WriteLine($"  WMF embedded signal '{signal}'={found}.");
            }
        }
        finally { ReleaseStgMedium(ref medium); }
    }

    private static void ProbeMathTypeEnhancedMetafile(
        System.Runtime.InteropServices.ComTypes.IDataObject dataObject,
        string artifactRoot)
    {
        var request = new FORMATETC
        {
            cfFormat = 14, // CF_ENHMETAFILE
            dwAspect = DVASPECT.DVASPECT_CONTENT,
            lindex = -1,
            ptd = IntPtr.Zero,
            tymed = TYMED.TYMED_ENHMF,
        };
        dataObject.GetData(ref request, out var medium);
        try
        {
            if (medium.tymed != TYMED.TYMED_ENHMF || medium.unionmember == IntPtr.Zero)
                throw new InvalidDataException(
                    $"MathType preview did not expose CF_ENHMETAFILE; actual={medium.tymed}.");
            var byteCount = GetEnhMetaFileBits(medium.unionmember, 0, null);
            if (byteCount == 0 || byteCount > 32 * 1024 * 1024)
                throw new InvalidDataException($"Unexpected MathType EMF size: {byteCount}.");
            var bytes = new byte[byteCount];
            if (GetEnhMetaFileBits(medium.unionmember, byteCount, bytes) != byteCount)
                throw new InvalidDataException("GetEnhMetaFileBits returned an incomplete MathType EMF.");
            File.WriteAllBytes(Path.Combine(artifactRoot, "mathtype-preview.emf"), bytes);

            var ascii = System.Text.Encoding.ASCII.GetString(bytes);
            var unicode = System.Text.Encoding.Unicode.GetString(bytes);
            var signals = new[] { "MathML", "<math", "mml:math", "MathType EF", "DSMT7" };
            foreach (var signal in signals)
            {
                var found = ascii.IndexOf(signal, StringComparison.OrdinalIgnoreCase) >= 0
                    || unicode.IndexOf(signal, StringComparison.OrdinalIgnoreCase) >= 0;
                Console.WriteLine($"  EMF embedded signal '{signal}'={found}.");
            }
        }
        finally { ReleaseStgMedium(ref medium); }
    }

    private static byte[] ReadAllComStream(IStream stream)
    {
        stream.Stat(out var stat, 1);
        if (stat.cbSize < 0 || stat.cbSize > 64 * 1024 * 1024)
            throw new InvalidDataException($"Unexpected OLE stream length: {stat.cbSize}.");
        var bytes = new byte[(int)stat.cbSize];
        stream.Seek(0, 0, IntPtr.Zero);
        var readPtr = Marshal.AllocHGlobal(sizeof(int));
        try
        {
            stream.Read(bytes, bytes.Length, readPtr);
            var read = Marshal.ReadInt32(readPtr);
            if (read != bytes.Length)
                throw new EndOfStreamException($"Expected {bytes.Length} OLE bytes, read {read}.");
            return bytes;
        }
        finally { Marshal.FreeHGlobal(readPtr); }
    }

    private static void InspectMathTypeCompoundFile(string path)
    {
        const int StgmReadShareExclusive = 0x10;
        var hr = StgOpenStorage(
            path,
            IntPtr.Zero,
            StgmReadShareExclusive,
            IntPtr.Zero,
            0,
            out var storage);
        if (hr < 0) Marshal.ThrowExceptionForHR(hr);
        IEnumSTATSTGProbe? enumerator = null;
        try
        {
            storage.Stat(out var rootStat, 1);
            Console.WriteLine($"  CFB root CLSID={rootStat.clsid:B}.");
            storage.EnumElements(0, IntPtr.Zero, 0, out enumerator);
            var entries = new System.Runtime.InteropServices.ComTypes.STATSTG[1];
            try
            {
                var foundEquationNative = false;
                while (true)
                {
                    var next = enumerator.Next(1, entries, out var fetched);
                    if (next != 0 || fetched != 1) break;
                    var entry = entries[0];
                    Console.WriteLine(
                        $"  CFB entry name='{EscapeStorageName(entry.pwcsName)}' type={entry.type} size={entry.cbSize} clsid={entry.clsid:B}");
                    if (!string.Equals(entry.pwcsName, "Equation Native", StringComparison.Ordinal))
                        continue;
                    foundEquationNative = true;
                    storage.OpenStream(
                        entry.pwcsName,
                        IntPtr.Zero,
                        StgmReadShareExclusive,
                        0,
                        out var nativeStream);
                    try
                    {
                        var bytes = ReadAllComStream(nativeStream);
                        File.WriteAllBytes(
                            Path.Combine(Path.GetDirectoryName(path)!, "equation-native.bin"),
                            bytes);
                        Console.WriteLine(
                            $"  Equation Native length={bytes.Length}; first={BitConverter.ToString(bytes.Take(Math.Min(64, bytes.Length)).ToArray())}");
                        if (bytes.Length < 34)
                            throw new InvalidDataException("Equation Native is too short for EQNOLEFILEHDR + MTEF v5.");
                        var cbHdr = BitConverter.ToUInt16(bytes, 0);
                        var objectBytes = BitConverter.ToUInt32(bytes, 8);
                        var mtefVersion = bytes[cbHdr];
                        Console.WriteLine(
                            $"  EQNOLEFILEHDR cbHdr={cbHdr} cbObject={objectBytes}; MTEF version={mtefVersion}.");
                        if (cbHdr != 28 || mtefVersion != 5)
                            throw new InvalidDataException(
                                $"Unexpected MathType native layout: cbHdr={cbHdr}, MTEF={mtefVersion}.");
                    }
                    finally { Release(nativeStream); }
                }
                if (!foundEquationNative)
                    throw new InvalidDataException("MathType CFB has no Equation Native stream.");
            }
            finally { }
        }
        finally
        {
            Release(enumerator);
            Release(storage);
        }
    }

    private static string EscapeStorageName(string? name)
    {
        if (name is null) return string.Empty;
        var builder = new System.Text.StringBuilder();
        foreach (var character in name)
        {
            if (char.IsControl(character)) builder.Append($"\\x{(int)character:X2}");
            else builder.Append(character);
        }
        return builder.ToString();
    }

    private static string ClipboardFormatName(uint format)
    {
        if (format < 0xC000) return format.ToString();
        var builder = new System.Text.StringBuilder(256);
        var length = GetClipboardFormatNameW(format, builder, builder.Capacity);
        return length > 0 ? builder.ToString() : format.ToString();
    }

    [ComImport]
    [Guid("0000000B-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IStorageProbe
    {
        void CreateStream(string name, int mode, int reserved1, int reserved2, out IStream stream);
        void OpenStream(string name, IntPtr reserved1, int mode, int reserved2, out IStream stream);
        void CreateStorage(string name, int mode, int reserved1, int reserved2, out IStorageProbe storage);
        void OpenStorage(string name, IntPtr priority, int mode, IntPtr exclude, int reserved, out IStorageProbe storage);
        void CopyTo(int ciidExclude, IntPtr rgiidExclude, IntPtr snbExclude, IStorageProbe destination);
        void MoveElementTo(string name, IStorageProbe destination, string newName, int flags);
        void Commit(int flags);
        void Revert();
        void EnumElements(int reserved1, IntPtr reserved2, int reserved3, out IEnumSTATSTGProbe enumerator);
        void DestroyElement(string name);
        void RenameElement(string oldName, string newName);
        void SetElementTimes(string name, IntPtr creation, IntPtr access, IntPtr modification);
        void SetClass(ref Guid clsid);
        void SetStateBits(int stateBits, int mask);
        void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG stat, int flags);
    }

    [ComImport]
    [Guid("0000000D-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IEnumSTATSTGProbe
    {
        [PreserveSig]
        int Next(
            uint count,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
            System.Runtime.InteropServices.ComTypes.STATSTG[] entries,
            out uint fetched);
        [PreserveSig] int Skip(uint count);
        void Reset();
        void Clone(out IEnumSTATSTGProbe clone);
    }

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int StgOpenStorage(
        string name,
        IntPtr priority,
        int mode,
        IntPtr exclude,
        int reserved,
        out IStorageProbe storage);

    [DllImport("ole32.dll")]
    private static extern int OleGetClipboard(out IntPtr dataObject);

    [DllImport("ole32.dll")]
    private static extern void ReleaseStgMedium(ref STGMEDIUM medium);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern uint RegisterClipboardFormatW(string format);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GlobalLock(IntPtr memory);

    [DllImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GlobalUnlock(IntPtr memory);

    [DllImport("gdi32.dll")]
    private static extern uint GetMetaFileBitsEx(
        IntPtr metafile,
        uint bufferSize,
        [Out] byte[]? data);

    [DllImport("gdi32.dll")]
    private static extern uint GetEnhMetaFileBits(
        IntPtr enhancedMetafile,
        uint bufferSize,
        [Out] byte[]? data);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClipboardFormatNameW(
        uint format,
        System.Text.StringBuilder buffer,
        int maxCount);
}
