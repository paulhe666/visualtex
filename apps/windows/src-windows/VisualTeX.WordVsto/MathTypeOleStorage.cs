using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

/// <summary>
/// Reads and rewrites the serialized Compound File Binary payload of an existing
/// MathType OLE object without activating or requiring a MathType COM server.
/// </summary>
internal static class MathTypeOleStorage
{
    internal static readonly Guid MathTypeEquationClsid =
        new("0002CE03-0000-0000-C000-000000000046");

    private const int StgmRead = 0x00000000;
    private const int StgmWrite = 0x00000001;
    private const int StgmReadWrite = 0x00000002;
    private const int StgmShareExclusive = 0x00000010;
    private const int StatFlagNoName = 1;
    private const int AdvfNoData = 1;

    internal sealed class RewriteResult
    {
        public byte[] CompoundFile { get; set; } = Array.Empty<byte>();
        public byte[] EquationNative { get; set; } = Array.Empty<byte>();
        public byte[] Mtef { get; set; } = Array.Empty<byte>();
        public int StructureOffset { get; set; }
    }

    internal sealed class ClipboardTransaction : IDisposable
    {
        private readonly System.Runtime.InteropServices.ComTypes.IDataObject? _previousClipboard;
        private readonly object? _previousClipboardObject;
        private readonly System.Runtime.InteropServices.ComTypes.IDataObject _sourceClipboard;
        private readonly object _sourceClipboardObject;
        private MathTypeOleClipboardProxy? _replacementProxy;
        private readonly bool _oleInitializedHere;
        private bool _disposed;

        internal ClipboardTransaction(InlineShape shape)
        {
            if (shape is null) throw new ArgumentNullException(nameof(shape));
            var oleResult = OleInitialize(IntPtr.Zero);
            if (oleResult == 0 || oleResult == 1)
                _oleInitializedHere = true;
            else if (oleResult < 0)
                Marshal.ThrowExceptionForHR(oleResult);
            _previousClipboard = TryGetOleClipboard(out _previousClipboardObject);

            var range = shape.Range;
            try { range.Copy(); }
            finally { Release(range); }

            _sourceClipboard = GetOleClipboard(out var sourceClipboardObject);
            _sourceClipboardObject = sourceClipboardObject;
            CompoundFile = ReadEmbeddedObject(_sourceClipboard);
            ValidateCompoundFile(CompoundFile);
        }

        internal byte[] CompoundFile { get; }

        internal void SetReplacementClipboard(
            byte[] compoundFile,
            string? emfPath,
            bool preferEmbedSource = false,
            bool standaloneExternal = false)
        {
            if (_disposed) throw new ObjectDisposedException(nameof(ClipboardTransaction));
            ValidateCompoundFile(compoundFile);
            _replacementProxy = new MathTypeOleClipboardProxy(
                _sourceClipboard,
                compoundFile,
                emfPath,
                preferEmbedSource,
                standaloneExternal);
            if (standaloneExternal)
            {
                // Word owns the IDataObject created by Range.Copy and can otherwise
                // reuse its private DOCX/VML preview path even after OleSetClipboard.
                // Releasing OLE clipboard ownership first makes the replacement look
                // like a genuine external OLE source while the transaction still
                // restores the user's previous clipboard when it completes.
                var clearResult = OleSetClipboard(null);
                if (clearResult < 0) Marshal.ThrowExceptionForHR(clearResult);
            }
            var result = OleSetClipboard(_replacementProxy);
            if (result < 0) Marshal.ThrowExceptionForHR(result);
        }

        internal int ReplacementStorageWriteCount =>
            _replacementProxy?.StorageWriteCount ?? 0;

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try
            {
                var result = OleSetClipboard(_previousClipboard);
                if (result < 0) Marshal.ThrowExceptionForHR(result);
            }
            catch
            {
                // Formula insertion outcome is more important than clipboard
                // restoration failure; source data remains valid until here.
            }
            _replacementProxy = null;
            Release(_sourceClipboardObject);
            Release(_previousClipboardObject);
            if (_oleInitializedHere) OleUninitialize();
        }
    }

    internal static ClipboardTransaction BeginClipboardTransaction(InlineShape shape) =>
        new(shape);

    internal static byte[] CaptureCompoundFile(InlineShape shape)
    {
        if (shape is null) throw new ArgumentNullException(nameof(shape));
        try
        {
            // Preferred path: Word's Flat OPC snapshot already contains the full
            // embedded OLE Compound File. Reading it here avoids clipboard ownership,
            // OLE activation and any dependency on a registered MathType server.
            return MathTypeWordOpenXml.Read(shape).CompoundFile;
        }
        catch
        {
            // Compatibility fallback for unusual legacy Word containers whose
            // Range.WordOpenXML does not expose the OLE package parts.
            using var transaction = BeginClipboardTransaction(shape);
            return transaction.CompoundFile;
        }
    }

    internal static string ReadMathMl(InlineShape shape) =>
        ReadMathMl(CaptureCompoundFile(shape));

    internal static bool LooksLikeMathTypeCompoundFile(byte[] compoundFile)
    {
        if (!HasCompoundFileSignature(compoundFile)) return false;
        var path = MaterializeTemporaryCompoundFile(compoundFile, "identify");
        try
        {
            var storage = OpenStorage(path, StgmRead | StgmShareExclusive);
            try
            {
                storage.Stat(out var stat, StatFlagNoName);
                if (stat.clsid != MathTypeEquationClsid) return false;
                try
                {
                    var native = ReadStream(storage, "Equation Native");
                    return native.Length > 34
                        && BitConverter.ToUInt16(native, 0) == 28
                        && native[28] == 5;
                }
                catch (COMException)
                {
                    return false;
                }
            }
            finally { Release(storage); }
        }
        finally { TryDelete(path); }
    }

    internal static string ReadMathMl(byte[] compoundFile)
    {
        var equationNative = ReadEquationNative(compoundFile);
        return MathTypeMtefCodec.ReadEquationNativeMathMl(equationNative);
    }

    internal static byte[] ReadEquationNative(byte[] compoundFile)
    {
        ValidateCompoundFile(compoundFile);
        var path = MaterializeTemporaryCompoundFile(compoundFile, "read-native");
        try
        {
            var storage = OpenStorage(path, StgmRead | StgmShareExclusive);
            try
            {
                ValidateMathTypeStorage(storage);
                return ReadStream(storage, "Equation Native");
            }
            finally { Release(storage); }
        }
        finally { TryDelete(path); }
    }

    internal static void CopyCompoundFileToStorage(
        byte[] compoundFile,
        IntPtr destinationStoragePointer)
    {
        ValidateCompoundFile(compoundFile);
        if (destinationStoragePointer == IntPtr.Zero)
            throw new ArgumentException(
                "Destination IStorage pointer is null.",
                nameof(destinationStoragePointer));

        var path = MaterializeTemporaryCompoundFile(compoundFile, "copy-to-storage");
        object? destinationObject = null;
        IStorageNative? sourceStorage = null;
        try
        {
            sourceStorage = OpenStorage(path, StgmRead | StgmShareExclusive);
            ValidateMathTypeStorage(sourceStorage);

            destinationObject = Marshal.GetObjectForIUnknown(destinationStoragePointer);
            if (destinationObject is not IStorageNative destinationStorage)
                throw new InvalidDataException(
                    "Word's GetDataHere destination does not implement IStorage.");

            sourceStorage.CopyTo(
                0,
                IntPtr.Zero,
                IntPtr.Zero,
                destinationStorage);
            var clsid = MathTypeEquationClsid;
            destinationStorage.SetClass(ref clsid);
            destinationStorage.Commit(0);
        }
        finally
        {
            Release(sourceStorage);
            if (destinationObject is not null && Marshal.IsComObject(destinationObject))
            {
                try { Marshal.ReleaseComObject(destinationObject); }
                catch { }
            }
            TryDelete(path);
        }
    }

    internal static RewriteResult RewriteMathTypeCompoundFile(
        byte[] compoundFile,
        string mathMl,
        bool inline)
    {
        ValidateCompoundFile(compoundFile);
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException("MathType storage rewrite requires MathML.");

        var path = MaterializeTemporaryCompoundFile(compoundFile, "rewrite");
        try
        {
            MathTypeMtefCodec.RewriteResult rewritten;
            var storage = OpenStorage(path, StgmReadWrite | StgmShareExclusive);
            try
            {
                ValidateMathTypeStorage(storage);
                var sourceNative = ReadStream(storage, "Equation Native");
                rewritten = MathTypeMtefCodec.RewriteEquationNative(
                    sourceNative,
                    mathMl,
                    inline);
                WriteStream(storage, "Equation Native", rewritten.EquationNative);
                storage.Commit(0);
            }
            finally { Release(storage); }

            return new RewriteResult
            {
                CompoundFile = File.ReadAllBytes(path),
                EquationNative = rewritten.EquationNative,
                Mtef = rewritten.Mtef,
                StructureOffset = rewritten.StructureOffset,
            };
        }
        finally { TryDelete(path); }
    }

    internal static Guid ReadCompoundFileRootClsid(byte[] compoundFile)
    {
        ValidateCompoundFile(compoundFile);
        var path = MaterializeTemporaryCompoundFile(compoundFile, "read-root-clsid");
        try
        {
            var storage = OpenStorage(path, StgmRead | StgmShareExclusive);
            try
            {
                storage.Stat(out var stat, StatFlagNoName);
                return stat.clsid;
            }
            finally { Release(storage); }
        }
        finally { TryDelete(path); }
    }

    internal static byte[] RewriteCompoundFileRootClsid(
        byte[] compoundFile,
        Guid rootClsid)
    {
        ValidateCompoundFile(compoundFile);
        var path = MaterializeTemporaryCompoundFile(compoundFile, "rewrite-root-clsid");
        try
        {
            var storage = OpenStorage(path, StgmReadWrite | StgmShareExclusive);
            try
            {
                storage.SetClass(ref rootClsid);
                storage.Commit(0);
            }
            finally { Release(storage); }
            return File.ReadAllBytes(path);
        }
        finally { TryDelete(path); }
    }

    internal static IReadOnlyList<string> ListCompoundFileEntries(byte[] compoundFile)
    {
        ValidateCompoundFile(compoundFile);
        var path = MaterializeTemporaryCompoundFile(compoundFile, "list-entries");
        try
        {
            var storage = OpenStorage(path, StgmRead | StgmShareExclusive);
            IEnumSTATSTGNative? enumerator = null;
            try
            {
                storage.EnumElements(0, IntPtr.Zero, 0, out enumerator);
                var entries = new List<string>();
                var buffer = new System.Runtime.InteropServices.ComTypes.STATSTG[1];
                var fetched = new int[1];
                while (enumerator.Next(1, buffer, fetched) == 0 && fetched[0] == 1)
                {
                    entries.Add(buffer[0].pwcsName ?? string.Empty);
                }
                return entries;
            }
            finally
            {
                Release(enumerator);
                Release(storage);
            }
        }
        finally { TryDelete(path); }
    }

    internal static byte[] ReadEnhancedMetafilePresentationCache(byte[] compoundFile)
    {
        ValidateCompoundFile(compoundFile);
        var path = MaterializeTemporaryCompoundFile(compoundFile, "read-presentation-cache");
        object? cacheObject = null;
        try
        {
            var storage = OpenStorage(path, StgmRead | StgmShareExclusive);
            try
            {
                var cacheClass = Guid.Empty;
                var cacheIid = new Guid("0000010E-0000-0000-C000-000000000046"); // IDataObject
                var result = CreateDataCache(
                    IntPtr.Zero,
                    ref cacheClass,
                    ref cacheIid,
                    out var cachePointer);
                if (result < 0) Marshal.ThrowExceptionForHR(result);
                if (cachePointer == IntPtr.Zero)
                    throw new InvalidOperationException("Windows OLE returned no data-cache IDataObject.");
                try { cacheObject = Marshal.GetObjectForIUnknown(cachePointer); }
                finally { Marshal.Release(cachePointer); }
                if (cacheObject is not IPersistStorageNative persist
                    || cacheObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                    throw new InvalidOperationException(
                        "Windows OLE data cache does not expose IPersistStorage/IDataObject.");
                result = persist.Load(storage);
                if (result < 0) Marshal.ThrowExceptionForHR(result);

                var format = new FORMATETC
                {
                    cfFormat = 14,
                    ptd = IntPtr.Zero,
                    dwAspect = DVASPECT.DVASPECT_CONTENT,
                    lindex = -1,
                    tymed = TYMED.TYMED_ENHMF,
                };
                dataObject.GetData(ref format, out var medium);
                try
                {
                    if (medium.tymed != TYMED.TYMED_ENHMF || medium.unionmember == IntPtr.Zero)
                        throw new InvalidDataException(
                            $"OLE presentation cache returned unexpected medium {medium.tymed}.");
                    var length = GetEnhMetaFileBits(medium.unionmember, 0, null);
                    if (length == 0 || length > 32 * 1024 * 1024)
                        throw new InvalidDataException(
                            $"Unexpected cached enhanced-metafile size {length}.");
                    var bytes = new byte[length];
                    if (GetEnhMetaFileBits(medium.unionmember, length, bytes) != length)
                        throw new InvalidDataException(
                            "OLE presentation cache returned an incomplete enhanced metafile.");
                    return bytes;
                }
                finally { ReleaseStgMedium(ref medium); }
            }
            finally { Release(storage); }
        }
        finally
        {
            Release(cacheObject);
            TryDelete(path);
        }
    }

    internal static byte[] AddEnhancedMetafilePresentationCache(
        byte[] compoundFile,
        string emfPath)
    {
        ValidateCompoundFile(compoundFile);
        if (string.IsNullOrWhiteSpace(emfPath) || !File.Exists(emfPath))
            throw new FileNotFoundException(
                "MathType presentation cache requires an existing EMF preview.",
                emfPath);

        var path = MaterializeTemporaryCompoundFile(compoundFile, "presentation-cache");
        object? cacheObject = null;
        try
        {
            var storage = OpenStorage(path, StgmReadWrite | StgmShareExclusive);
            try
            {
                ValidateMathTypeStorage(storage);
                var cacheClass = Guid.Empty;
                var cacheIid = new Guid("0000011E-0000-0000-C000-000000000046");
                var result = CreateDataCache(
                    IntPtr.Zero,
                    ref cacheClass,
                    ref cacheIid,
                    out var cachePointer);
                if (result < 0) Marshal.ThrowExceptionForHR(result);
                if (cachePointer == IntPtr.Zero)
                    throw new InvalidOperationException("Windows OLE returned no presentation cache object.");
                try { cacheObject = Marshal.GetObjectForIUnknown(cachePointer); }
                finally { Marshal.Release(cachePointer); }

                if (cacheObject is not IPersistStorageNative persist
                    || cacheObject is not IOleCacheNative cache)
                    throw new InvalidOperationException(
                        "Windows OLE data cache does not expose IPersistStorage/IOleCache.");

                result = persist.Load(storage);
                if (result < 0) Marshal.ThrowExceptionForHR(result);

                var format = new FORMATETC
                {
                    cfFormat = 14, // CF_ENHMETAFILE
                    ptd = IntPtr.Zero,
                    dwAspect = DVASPECT.DVASPECT_CONTENT,
                    lindex = -1,
                    tymed = TYMED.TYMED_ENHMF,
                };
                result = cache.Cache(ref format, AdvfNoData, out _);
                if (result < 0) Marshal.ThrowExceptionForHR(result);

                var metafile = GetEnhMetaFileW(emfPath);
                if (metafile == IntPtr.Zero)
                    throw new InvalidDataException(
                        $"Windows could not load the VisualTeX EMF preview '{emfPath}'.");
                try
                {
                    var medium = new STGMEDIUM
                    {
                        tymed = TYMED.TYMED_ENHMF,
                        unionmember = metafile,
                        pUnkForRelease = null,
                    };
                    result = cache.SetData(ref format, ref medium, false);
                    if (result < 0) Marshal.ThrowExceptionForHR(result);
                }
                finally { DeleteEnhMetaFile(metafile); }

                result = persist.Save(storage, true);
                if (result < 0) Marshal.ThrowExceptionForHR(result);
                result = persist.SaveCompleted(storage);
                if (result < 0) Marshal.ThrowExceptionForHR(result);
                storage.Commit(0);
                result = persist.HandsOffStorage();
                if (result < 0) Marshal.ThrowExceptionForHR(result);
                Release(cacheObject);
                cacheObject = null;
            }
            finally { Release(storage); }

            var cached = File.ReadAllBytes(path);
            // The presentation cache must never alter MathType's semantic stream.
            _ = ReadMathMl(cached);
            return cached;
        }
        finally
        {
            Release(cacheObject);
            TryDelete(path);
        }
    }

    private static void ValidateCompoundFile(byte[] compoundFile)
    {
        if (!HasCompoundFileSignature(compoundFile))
            throw new InvalidDataException(
                "The embedded MathType object is not a Compound File Binary payload.");
    }

    private static bool HasCompoundFileSignature(byte[] data) =>
        data is { Length: >= 8 }
        && data[0] == 0xD0 && data[1] == 0xCF
        && data[2] == 0x11 && data[3] == 0xE0
        && data[4] == 0xA1 && data[5] == 0xB1
        && data[6] == 0x1A && data[7] == 0xE1;

    private static void ValidateMathTypeStorage(IStorageNative storage)
    {
        storage.Stat(out var stat, StatFlagNoName);
        if (stat.clsid != MathTypeEquationClsid)
            throw new InvalidDataException(
                $"Embedded OLE storage CLSID {stat.clsid:B} is not MathType Equation.DSMT4.");
        var native = ReadStream(storage, "Equation Native");
        if (native.Length < 34 || BitConverter.ToUInt16(native, 0) != 28)
            throw new InvalidDataException("MathType Equation Native stream has an unsupported header.");
        if (native[28] != 5)
            throw new InvalidDataException(
                $"VisualTeX currently supports direct MathType preservation for MTEF v5 only, actual={native[28]}.");
    }

    private static IStorageNative OpenStorage(string path, int mode)
    {
        var result = StgOpenStorage(
            path,
            IntPtr.Zero,
            mode,
            IntPtr.Zero,
            0,
            out var storage);
        if (result < 0) Marshal.ThrowExceptionForHR(result);
        return storage;
    }

    private static byte[] ReadEmbeddedObject(
        System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
    {
        var formatId = RegisterClipboardFormatW("Embedded Object");
        if (formatId == 0 || formatId > ushort.MaxValue)
            throw new InvalidOperationException("Could not register CFSTR_EMBEDDEDOBJECT.");
        var request = new FORMATETC
        {
            cfFormat = unchecked((short)formatId),
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
                    $"Word did not expose the embedded OLE object as TYMED_ISTREAM; actual={medium.tymed}.");
            streamObject = Marshal.GetObjectForIUnknown(medium.unionmember);
            if (streamObject is not IStream stream)
                throw new InvalidDataException("CFSTR_EMBEDDEDOBJECT is not an IStream.");
            return ReadComStream(stream);
        }
        finally
        {
            Release(streamObject);
            ReleaseStgMedium(ref medium);
        }
    }

    private static byte[] ReadComStream(IStream stream)
    {
        stream.Stat(out var stat, StatFlagNoName);
        if (stat.cbSize < 0 || stat.cbSize > 64L * 1024 * 1024)
            throw new InvalidDataException($"Unexpected embedded OLE stream length: {stat.cbSize}.");
        var bytes = new byte[(int)stat.cbSize];
        stream.Seek(0, 0, IntPtr.Zero);
        var readPointer = Marshal.AllocHGlobal(sizeof(int));
        try
        {
            stream.Read(bytes, bytes.Length, readPointer);
            var read = Marshal.ReadInt32(readPointer);
            if (read != bytes.Length)
                throw new EndOfStreamException(
                    $"Expected {bytes.Length} embedded OLE bytes, read {read}.");
            return bytes;
        }
        finally { Marshal.FreeHGlobal(readPointer); }
    }

    private static System.Runtime.InteropServices.ComTypes.IDataObject GetOleClipboard(
        out object clipboardObject)
    {
        const int ClipbrdECantOpen = unchecked((int)0x800401D0);
        var retryDelaysMs = new[] { 0, 15, 35, 70, 120, 200 };
        var lastResult = 0;

        foreach (var delayMs in retryDelaysMs)
        {
            if (delayMs > 0)
                System.Threading.Thread.Sleep(delayMs);

            var result = OleGetClipboard(out var pointer);
            lastResult = result;
            if (result == ClipbrdECantOpen || (result >= 0 && pointer == IntPtr.Zero))
                continue;
            if (result < 0)
                Marshal.ThrowExceptionForHR(result);

            try
            {
                clipboardObject = Marshal.GetObjectForIUnknown(pointer);
            }
            finally { Marshal.Release(pointer); }
            return clipboardObject as System.Runtime.InteropServices.ComTypes.IDataObject
                ?? throw new InvalidOperationException("Windows OLE clipboard object does not implement IDataObject.");
        }

        if (lastResult < 0)
            Marshal.ThrowExceptionForHR(lastResult);
        throw new InvalidOperationException("Windows OLE clipboard returned no IDataObject after bounded retries.");
    }

    private static System.Runtime.InteropServices.ComTypes.IDataObject? TryGetOleClipboard(
        out object? clipboardObject)
    {
        clipboardObject = null;
        var result = OleGetClipboard(out var pointer);
        if (result < 0 || pointer == IntPtr.Zero) return null;
        try { clipboardObject = Marshal.GetObjectForIUnknown(pointer); }
        finally { Marshal.Release(pointer); }
        return clipboardObject as System.Runtime.InteropServices.ComTypes.IDataObject;
    }

    private static byte[] ReadStream(IStorageNative storage, string name)
    {
        storage.OpenStream(
            name,
            IntPtr.Zero,
            StgmRead | StgmShareExclusive,
            0,
            out var stream);
        try
        {
            stream.Stat(out var stat, StatFlagNoName);
            if (stat.cbSize < 0 || stat.cbSize > 64L * 1024 * 1024)
                throw new InvalidDataException(
                    $"Unexpected MathType storage stream '{name}' length: {stat.cbSize}.");
            var bytes = new byte[(int)stat.cbSize];
            stream.Seek(0, 0, IntPtr.Zero);
            var readPointer = Marshal.AllocHGlobal(sizeof(int));
            try
            {
                stream.Read(bytes, bytes.Length, readPointer);
                var read = Marshal.ReadInt32(readPointer);
                if (read != bytes.Length)
                    throw new EndOfStreamException(
                        $"MathType stream '{name}' expected {bytes.Length} bytes, read {read}.");
            }
            finally { Marshal.FreeHGlobal(readPointer); }
            return bytes;
        }
        finally { Release(stream); }
    }

    private static void WriteStream(
        IStorageNative storage,
        string name,
        byte[] data)
    {
        storage.OpenStream(
            name,
            IntPtr.Zero,
            StgmReadWrite | StgmShareExclusive,
            0,
            out var stream);
        try
        {
            stream.SetSize(data.LongLength);
            stream.Seek(0, 0, IntPtr.Zero);
            var writtenPointer = Marshal.AllocHGlobal(sizeof(int));
            try
            {
                stream.Write(data, data.Length, writtenPointer);
                var written = Marshal.ReadInt32(writtenPointer);
                if (written != data.Length)
                    throw new IOException(
                        $"MathType stream '{name}' expected to write {data.Length} bytes, wrote {written}.");
            }
            finally { Marshal.FreeHGlobal(writtenPointer); }
            stream.Commit(0);
        }
        finally { Release(stream); }
    }

    private static string MaterializeTemporaryCompoundFile(
        byte[] compoundFile,
        string purpose)
    {
        var root = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(root);
        var path = Path.Combine(
            root,
            $"mathtype-{purpose}-{Guid.NewGuid():N}.ole");
        File.WriteAllBytes(path, compoundFile);
        return path;
    }

    private static void TryDelete(string path)
    {
        try { File.Delete(path); }
        catch { }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            try { Marshal.FinalReleaseComObject(value); }
            catch { }
        }
    }

    [ComImport]
    [Guid("0000000B-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IStorageNative
    {
        void CreateStream(string name, int mode, int reserved1, int reserved2, out IStream stream);
        void OpenStream(string name, IntPtr reserved1, int mode, int reserved2, out IStream stream);
        void CreateStorage(string name, int mode, int reserved1, int reserved2, out IStorageNative storage);
        void OpenStorage(string name, IntPtr priority, int mode, IntPtr exclude, int reserved, out IStorageNative storage);
        void CopyTo(int ciidExclude, IntPtr rgiidExclude, IntPtr snbExclude, IStorageNative destination);
        void MoveElementTo(string name, IStorageNative destination, string newName, int flags);
        void Commit(int flags);
        void Revert();
        void EnumElements(int reserved1, IntPtr reserved2, int reserved3, out IEnumSTATSTGNative enumerator);
        void DestroyElement(string name);
        void RenameElement(string oldName, string newName);
        void SetElementTimes(string name, IntPtr creation, IntPtr access, IntPtr modification);
        void SetClass(ref Guid clsid);
        void SetStateBits(int stateBits, int mask);
        void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG stat, int flags);
    }

    [ComImport]
    [Guid("0000010A-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPersistStorageNative
    {
        [PreserveSig] int GetClassID(out Guid classId);
        [PreserveSig] int IsDirty();
        [PreserveSig] int InitNew(IStorageNative storage);
        [PreserveSig] int Load(IStorageNative storage);
        [PreserveSig] int Save(IStorageNative storage, [MarshalAs(UnmanagedType.Bool)] bool sameAsLoad);
        [PreserveSig] int SaveCompleted(IStorageNative storage);
        [PreserveSig] int HandsOffStorage();
    }

    [ComImport]
    [Guid("0000011E-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleCacheNative
    {
        [PreserveSig] int Cache(ref FORMATETC format, int advf, out int connection);
        [PreserveSig] int Uncache(int connection);
        [PreserveSig] int EnumCache(out IntPtr enumerator);
        [PreserveSig] int InitCache(
            [MarshalAs(UnmanagedType.Interface)]
            System.Runtime.InteropServices.ComTypes.IDataObject dataObject);
        [PreserveSig] int SetData(
            ref FORMATETC format,
            ref STGMEDIUM medium,
            [MarshalAs(UnmanagedType.Bool)] bool release);
    }

    [ComImport]
    [Guid("0000000D-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IEnumSTATSTGNative
    {
        [PreserveSig]
        int Next(
            int celt,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
            System.Runtime.InteropServices.ComTypes.STATSTG[] entries,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeConst = 1)] int[] fetched);
        [PreserveSig] int Skip(int celt);
        [PreserveSig] int Reset();
        void Clone(out IEnumSTATSTGNative clone);
    }

    [DllImport("ole32.dll")]
    private static extern int CreateDataCache(
        IntPtr outer,
        ref Guid classId,
        ref Guid interfaceId,
        out IntPtr cacheObject);

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr GetEnhMetaFileW(string fileName);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeleteEnhMetaFile(IntPtr metafile);

    [DllImport("gdi32.dll")]
    private static extern uint GetEnhMetaFileBits(
        IntPtr metafile,
        uint bufferSize,
        [Out] byte[]? data);

    [DllImport("ole32.dll")]
    private static extern void ReleaseStgMedium(ref STGMEDIUM medium);

    [DllImport("ole32.dll")]
    private static extern int OleInitialize(IntPtr reserved);

    [DllImport("ole32.dll")]
    private static extern void OleUninitialize();

    [DllImport("ole32.dll")]
    private static extern int OleGetClipboard(out IntPtr dataObject);

    [DllImport("ole32.dll")]
    private static extern int OleSetClipboard(
        [MarshalAs(UnmanagedType.Interface)]
        System.Runtime.InteropServices.ComTypes.IDataObject? dataObject);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern uint RegisterClipboardFormatW(string format);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int StgOpenStorage(
        string name,
        IntPtr priority,
        int mode,
        IntPtr exclude,
        int reserved,
        out IStorageNative storage);
}
