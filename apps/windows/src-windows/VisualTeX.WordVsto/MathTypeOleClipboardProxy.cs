using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace VisualTeX.WordVsto;

/// <summary>
/// OLE clipboard proxy used to let Word embed a rewritten MathType CFB without
/// activating a MathType server. Descriptor/native presentation formats are
/// delegated to Word's original copied OLE IDataObject, while Embedded Object
/// and (optionally) CF_ENHMETAFILE are supplied by VisualTeX.
/// </summary>
internal sealed class MathTypeOleClipboardProxy : System.Runtime.InteropServices.ComTypes.IDataObject
{
    private const int S_OK = 0;
    private const int DV_E_FORMATETC = unchecked((int)0x80040064);
    private const int E_NOTIMPL = unchecked((int)0x80004001);
    private const uint GmemMoveable = 0x0002;
    private const short CfMetafilePict = 3;
    private const short CfEnhMetafile = 14;
    private const int MmAnisotropic = 8;
    private const uint GmemZeroInit = 0x0040;

    private readonly System.Runtime.InteropServices.ComTypes.IDataObject? _source;
    private readonly byte[] _compoundFile;
    private readonly string? _emfPath;
    private readonly short _embeddedObjectFormat;
    private readonly short _embedSourceFormat;
    private readonly short _nativeFormat;
    private readonly short _objectDescriptorFormat;
    private readonly bool _preferEmbedSource;
    private readonly bool _standaloneExternal;
    private readonly byte[]? _objectDescriptorBytes;

    internal int StorageWriteCount { get; private set; }

    internal MathTypeOleClipboardProxy(
        System.Runtime.InteropServices.ComTypes.IDataObject source,
        byte[] compoundFile,
        string? emfPath,
        bool preferEmbedSource = false,
        bool standaloneExternal = false)
        : this(
            source ?? throw new ArgumentNullException(nameof(source)),
            compoundFile,
            emfPath,
            preferEmbedSource,
            standaloneExternal,
            objectDescriptorBytes: null)
    {
    }

    internal MathTypeOleClipboardProxy(
        byte[] compoundFile,
        string? emfPath,
        byte[] objectDescriptorBytes)
        : this(
            source: null,
            compoundFile,
            emfPath,
            preferEmbedSource: false,
            standaloneExternal: true,
            objectDescriptorBytes)
    {
    }

    private MathTypeOleClipboardProxy(
        System.Runtime.InteropServices.ComTypes.IDataObject? source,
        byte[] compoundFile,
        string? emfPath,
        bool preferEmbedSource,
        bool standaloneExternal,
        byte[]? objectDescriptorBytes)
    {
        if (!standaloneExternal && source is null)
            throw new ArgumentNullException(nameof(source));
        _source = source;
        _compoundFile = compoundFile ?? throw new ArgumentNullException(nameof(compoundFile));
        _emfPath = string.IsNullOrWhiteSpace(emfPath) ? null : emfPath;
        _embeddedObjectFormat = RegisterOleFormat("Embedded Object");
        _embedSourceFormat = RegisterOleFormat("Embed Source");
        _nativeFormat = RegisterOleFormat("Native");
        _objectDescriptorFormat = RegisterOleFormat("Object Descriptor");
        _preferEmbedSource = preferEmbedSource;
        _standaloneExternal = standaloneExternal;
        _objectDescriptorBytes = standaloneExternal
            ? objectDescriptorBytes
                ?? (source is not null
                    ? TryReadHGlobalFormat(source, _objectDescriptorFormat)
                    : null)
            : null;
        if (standaloneExternal && (_objectDescriptorBytes is null || _objectDescriptorBytes.Length == 0))
            throw new InvalidDataException(
                "Standalone MathType OLE data requires an Object Descriptor.");
    }

    internal static byte[] CreateStandaloneObjectDescriptor(
        float widthPt,
        float heightPt)
    {
        const int fixedDescriptorBytes = 52;
        const uint contentAspect = 1;
        var fullUserType = Encoding.Unicode.GetBytes("MathType 7.0 Equation\0");
        var sourceOfCopy = Encoding.Unicode.GetBytes("VisualTeX\0");
        var fullUserTypeOffset = fixedDescriptorBytes;
        var sourceOfCopyOffset = fixedDescriptorBytes + fullUserType.Length;
        var totalBytes = sourceOfCopyOffset + sourceOfCopy.Length;

        using var stream = new MemoryStream(totalBytes);
        using var writer = new BinaryWriter(stream, Encoding.Unicode, leaveOpen: true);
        writer.Write((uint)totalBytes);
        writer.Write(MathTypeOleStorage.MathTypeEquationClsid.ToByteArray());
        writer.Write(contentAspect);
        writer.Write(PointsToHimetric(widthPt));
        writer.Write(PointsToHimetric(heightPt));
        writer.Write(0);
        writer.Write(0);
        writer.Write(0u);
        writer.Write((uint)fullUserTypeOffset);
        writer.Write((uint)sourceOfCopyOffset);
        if (stream.Position != fixedDescriptorBytes)
            throw new InvalidOperationException(
                $"VisualTeX built an invalid OBJECTDESCRIPTOR header size: {stream.Position}.");
        writer.Write(fullUserType);
        writer.Write(sourceOfCopy);
        writer.Flush();
        return stream.ToArray();
    }

    private static int PointsToHimetric(float points)
    {
        var safePoints = Math.Max(1d / 72d, points);
        return Math.Max(1, checked((int)Math.Round(safePoints * 2540d / 72d)));
    }

    public void GetData(ref FORMATETC format, out STGMEDIUM medium)
    {
        Trace($"GetData cf={unchecked((ushort)format.cfFormat)} tymed={format.tymed}");
        if (_preferEmbedSource && format.cfFormat == _embeddedObjectFormat)
        {
            Marshal.ThrowExceptionForHR(DV_E_FORMATETC);
        }
        if (IsObjectPayloadFormat(format.cfFormat))
        {
            medium = CreateEmbeddedObjectMedium(format.tymed);
            return;
        }
        if (_standaloneExternal && format.cfFormat == _objectDescriptorFormat)
        {
            medium = CreateHGlobalMedium(_objectDescriptorBytes!);
            return;
        }
        if (format.cfFormat == CfEnhMetafile
            && _emfPath is not null
            && ShouldExposeEnhancedMetafile())
        {
            medium = CreateEnhancedMetafileMedium();
            return;
        }
        if (format.cfFormat == CfMetafilePict
            && _emfPath is not null
            && ShouldExposeMetafilePicture())
        {
            medium = CreateMetafilePictureMedium();
            return;
        }
        if (_standaloneExternal)
            Marshal.ThrowExceptionForHR(DV_E_FORMATETC);
        _source!.GetData(ref format, out medium);
    }

    public void GetDataHere(ref FORMATETC format, ref STGMEDIUM medium)
    {
        Trace($"GetDataHere cf={unchecked((ushort)format.cfFormat)} request={format.tymed} medium={medium.tymed}");
        if (_preferEmbedSource && format.cfFormat == _embeddedObjectFormat)
            Marshal.ThrowExceptionForHR(DV_E_FORMATETC);
        if (IsObjectPayloadFormat(format.cfFormat)
            && medium.tymed == TYMED.TYMED_ISTORAGE
            && medium.unionmember != IntPtr.Zero)
        {
            MathTypeOleStorage.CopyCompoundFileToStorage(
                _compoundFile,
                medium.unionmember);
            StorageWriteCount++;
            Trace("GetDataHere wrote rewritten MathType CFB into destination IStorage");
            return;
        }
        if (_standaloneExternal)
            Marshal.ThrowExceptionForHR(DV_E_FORMATETC);
        _source!.GetDataHere(ref format, ref medium);
    }

    [PreserveSig]
    public int QueryGetData(ref FORMATETC format)
    {
        Trace($"QueryGetData cf={unchecked((ushort)format.cfFormat)} tymed={format.tymed}");
        if (_preferEmbedSource && format.cfFormat == _embeddedObjectFormat)
            return DV_E_FORMATETC;
        if (IsObjectPayloadFormat(format.cfFormat))
        {
            return (format.tymed & (TYMED.TYMED_ISTREAM | TYMED.TYMED_HGLOBAL | TYMED.TYMED_ISTORAGE)) != 0
                ? S_OK
                : DV_E_FORMATETC;
        }
        if (_standaloneExternal && format.cfFormat == _objectDescriptorFormat)
            return (format.tymed & TYMED.TYMED_HGLOBAL) != 0 ? S_OK : DV_E_FORMATETC;
        if (format.cfFormat == CfEnhMetafile
            && _emfPath is not null
            && ShouldExposeEnhancedMetafile())
            return (format.tymed & TYMED.TYMED_ENHMF) != 0 ? S_OK : DV_E_FORMATETC;
        if (format.cfFormat == CfMetafilePict
            && _emfPath is not null
            && ShouldExposeMetafilePicture())
            return (format.tymed & TYMED.TYMED_MFPICT) != 0 ? S_OK : DV_E_FORMATETC;
        return _standaloneExternal
            ? DV_E_FORMATETC
            : _source!.QueryGetData(ref format);
    }

    [PreserveSig]
    public int GetCanonicalFormatEtc(ref FORMATETC formatIn, out FORMATETC formatOut)
    {
        if (_standaloneExternal)
        {
            formatOut = formatIn;
            formatOut.ptd = IntPtr.Zero;
            return E_NOTIMPL;
        }
        return _source!.GetCanonicalFormatEtc(ref formatIn, out formatOut);
    }

    public void SetData(ref FORMATETC formatIn, ref STGMEDIUM medium, bool release)
    {
        if (_standaloneExternal) Marshal.ThrowExceptionForHR(E_NOTIMPL);
        _source!.SetData(ref formatIn, ref medium, release);
    }

    public IEnumFORMATETC EnumFormatEtc(DATADIR direction)
    {
        if (_standaloneExternal)
        {
            if (direction != DATADIR.DATADIR_GET)
                return new FormatEtcEnumerator(Array.Empty<FORMATETC>());
            var standaloneFormats = new List<FORMATETC>();
            EnsureFormat(
                standaloneFormats,
                _preferEmbedSource ? _embedSourceFormat : _embeddedObjectFormat,
                TYMED.TYMED_ISTREAM | TYMED.TYMED_HGLOBAL | TYMED.TYMED_ISTORAGE);
            EnsureFormat(standaloneFormats, _objectDescriptorFormat, TYMED.TYMED_HGLOBAL);
            if (_emfPath is not null)
            {
                if (ShouldExposeEnhancedMetafile())
                    EnsureFormat(standaloneFormats, CfEnhMetafile, TYMED.TYMED_ENHMF);
                if (ShouldExposeMetafilePicture())
                    EnsureFormat(standaloneFormats, CfMetafilePict, TYMED.TYMED_MFPICT);
            }
            Trace(
                "EnumFormatEtc standalone formats="
                + string.Join(",", standaloneFormats.Select(item => unchecked((ushort)item.cfFormat))));
            return new FormatEtcEnumerator(standaloneFormats);
        }

        if (!_preferEmbedSource || direction != DATADIR.DATADIR_GET)
            return _source!.EnumFormatEtc(direction);

        var formats = ReadFormats(_source!.EnumFormatEtc(direction));
        formats.RemoveAll(format => format.cfFormat == _embeddedObjectFormat);
        EnsureFormat(
            formats,
            _embedSourceFormat,
            TYMED.TYMED_ISTREAM | TYMED.TYMED_HGLOBAL | TYMED.TYMED_ISTORAGE);
        if (_emfPath is not null)
        {
            formats.RemoveAll(format =>
                format.cfFormat == CfEnhMetafile || format.cfFormat == CfMetafilePict);
            EnsureFormat(formats, CfEnhMetafile, TYMED.TYMED_ENHMF);
            EnsureFormat(formats, CfMetafilePict, TYMED.TYMED_MFPICT);
        }
        Trace(
            "EnumFormatEtc exposes Embed Source="
            + formats.Any(format => format.cfFormat == _embedSourceFormat)
            + ", Embedded Object="
            + formats.Any(format => format.cfFormat == _embeddedObjectFormat)
            + ", ENHMF="
            + formats.Any(format => format.cfFormat == CfEnhMetafile));
        return new FormatEtcEnumerator(formats);
    }

    [PreserveSig]
    public int DAdvise(ref FORMATETC pFormatetc, ADVF advf, IAdviseSink adviseSink, out int connection)
    {
        try
        {
            if (_standaloneExternal)
            {
                connection = 0;
                return E_NOTIMPL;
            }
            return _source!.DAdvise(ref pFormatetc, advf, adviseSink, out connection);
        }
        catch (COMException error)
        {
            connection = 0;
            return error.ErrorCode;
        }
        catch
        {
            connection = 0;
            return E_NOTIMPL;
        }
    }

    public void DUnadvise(int connection)
    {
        if (_standaloneExternal) Marshal.ThrowExceptionForHR(E_NOTIMPL);
        _source!.DUnadvise(connection);
    }

    [PreserveSig]
    public int EnumDAdvise(out IEnumSTATDATA? enumAdvise)
    {
        try
        {
            if (_standaloneExternal)
            {
                enumAdvise = null;
                return E_NOTIMPL;
            }
            return _source!.EnumDAdvise(out enumAdvise);
        }
        catch (COMException error)
        {
            enumAdvise = null;
            return error.ErrorCode;
        }
        catch
        {
            enumAdvise = null;
            return E_NOTIMPL;
        }
    }

    private static bool ShouldExposeEnhancedMetafile()
    {
        var mode = Environment.GetEnvironmentVariable(
            "VISUALTEX_ACCEPTANCE_MATHTYPE_PREVIEW_FORMAT");
        return !string.Equals(mode, "wmf-only", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ShouldExposeMetafilePicture()
    {
        var mode = Environment.GetEnvironmentVariable(
            "VISUALTEX_ACCEPTANCE_MATHTYPE_PREVIEW_FORMAT");
        return !string.Equals(mode, "emf-only", StringComparison.OrdinalIgnoreCase);
    }

    private static void Trace(string message)
    {
        if (string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            Console.WriteLine("    [MathType clipboard proxy] " + message);
    }

    private bool IsObjectPayloadFormat(short format) =>
        format == _embeddedObjectFormat
        || format == _embedSourceFormat
        || format == _nativeFormat;

    private static List<FORMATETC> ReadFormats(IEnumFORMATETC enumerator)
    {
        var formats = new List<FORMATETC>();
        try
        {
            var buffer = new FORMATETC[1];
            var fetched = new int[1];
            while (enumerator.Next(1, buffer, fetched) == S_OK && fetched[0] == 1)
            {
                var format = buffer[0];
                // The target-device pointer is caller-owned COM memory and is not
                // needed for VisualTeX's synthesized object/presentation formats.
                format.ptd = IntPtr.Zero;
                formats.Add(format);
            }
        }
        finally
        {
            if (Marshal.IsComObject(enumerator))
            {
                try { Marshal.ReleaseComObject(enumerator); }
                catch { }
            }
        }
        return formats;
    }

    private static void EnsureFormat(
        List<FORMATETC> formats,
        short clipboardFormat,
        TYMED tymed)
    {
        if (formats.Any(format => format.cfFormat == clipboardFormat)) return;
        formats.Add(new FORMATETC
        {
            cfFormat = clipboardFormat,
            ptd = IntPtr.Zero,
            dwAspect = DVASPECT.DVASPECT_CONTENT,
            lindex = -1,
            tymed = tymed,
        });
    }

    private sealed class FormatEtcEnumerator : IEnumFORMATETC
    {
        private readonly IReadOnlyList<FORMATETC> _formats;
        private int _index;

        internal FormatEtcEnumerator(IReadOnlyList<FORMATETC> formats, int index = 0)
        {
            _formats = formats;
            _index = index;
        }

        public int Next(int celt, FORMATETC[] rgelt, int[] pceltFetched)
        {
            if (celt <= 0 || rgelt is null || rgelt.Length < celt)
                return 1;
            var fetched = 0;
            while (fetched < celt && _index < _formats.Count)
            {
                rgelt[fetched] = _formats[_index];
                fetched++;
                _index++;
            }
            if (pceltFetched is not null && pceltFetched.Length > 0)
                pceltFetched[0] = fetched;
            return fetched == celt ? S_OK : 1;
        }

        public int Skip(int celt)
        {
            if (celt <= 0) return S_OK;
            var remaining = _formats.Count - _index;
            var skipped = Math.Min(celt, Math.Max(0, remaining));
            _index += skipped;
            return skipped == celt ? S_OK : 1;
        }

        public int Reset()
        {
            _index = 0;
            return S_OK;
        }

        public void Clone(out IEnumFORMATETC newEnum) =>
            newEnum = new FormatEtcEnumerator(_formats, _index);
    }

    private static byte[]? TryReadHGlobalFormat(
        System.Runtime.InteropServices.ComTypes.IDataObject source,
        short clipboardFormat)
    {
        var request = new FORMATETC
        {
            cfFormat = clipboardFormat,
            ptd = IntPtr.Zero,
            dwAspect = DVASPECT.DVASPECT_CONTENT,
            lindex = -1,
            tymed = TYMED.TYMED_HGLOBAL,
        };
        try
        {
            source.GetData(ref request, out var medium);
            try
            {
                if (medium.tymed != TYMED.TYMED_HGLOBAL || medium.unionmember == IntPtr.Zero)
                    return null;
                var size = GlobalSize(medium.unionmember).ToUInt64();
                if (size == 0 || size > 4 * 1024 * 1024)
                    return null;
                var pointer = GlobalLock(medium.unionmember);
                if (pointer == IntPtr.Zero) return null;
                try
                {
                    var bytes = new byte[(int)size];
                    Marshal.Copy(pointer, bytes, 0, bytes.Length);
                    return bytes;
                }
                finally { GlobalUnlock(medium.unionmember); }
            }
            finally { ReleaseStgMedium(ref medium); }
        }
        catch
        {
            return null;
        }
    }

    private static STGMEDIUM CreateHGlobalMedium(byte[] bytes)
    {
        var memory = GlobalAlloc(GmemMoveable, new UIntPtr((uint)bytes.Length));
        if (memory == IntPtr.Zero)
            throw new OutOfMemoryException("GlobalAlloc failed for OLE clipboard data.");
        var target = GlobalLock(memory);
        if (target == IntPtr.Zero)
        {
            GlobalFree(memory);
            throw new OutOfMemoryException("GlobalLock failed for OLE clipboard data.");
        }
        try { Marshal.Copy(bytes, 0, target, bytes.Length); }
        finally { GlobalUnlock(memory); }
        return new STGMEDIUM
        {
            tymed = TYMED.TYMED_HGLOBAL,
            unionmember = memory,
            pUnkForRelease = null,
        };
    }

    private static short RegisterOleFormat(string name)
    {
        var format = RegisterClipboardFormatW(name);
        if (format == 0 || format > ushort.MaxValue)
            throw new InvalidOperationException($"Could not register OLE clipboard format '{name}'.");
        return unchecked((short)format);
    }

    private STGMEDIUM CreateEmbeddedObjectMedium(TYMED requested)
    {
        if ((requested & TYMED.TYMED_ISTREAM) != 0)
        {
            var stream = SHCreateMemStream(_compoundFile, (uint)_compoundFile.Length);
            if (stream == IntPtr.Zero)
                throw new OutOfMemoryException("SHCreateMemStream failed for MathType CFB.");
            return new STGMEDIUM
            {
                tymed = TYMED.TYMED_ISTREAM,
                unionmember = stream,
                pUnkForRelease = null,
            };
        }
        if ((requested & TYMED.TYMED_HGLOBAL) != 0)
            return CreateHGlobalMedium(_compoundFile);
        Marshal.ThrowExceptionForHR(DV_E_FORMATETC);
        return default;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RectLong
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct SizeLong
    {
        public int Cx;
        public int Cy;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct EnhancedMetafileHeader
    {
        public uint Type;
        public uint Size;
        public RectLong Bounds;
        public RectLong Frame;
        public uint Signature;
        public uint Version;
        public uint Bytes;
        public uint Records;
        public ushort Handles;
        public ushort Reserved;
        public uint DescriptionCharacters;
        public uint DescriptionOffset;
        public uint PaletteEntries;
        public SizeLong Device;
        public SizeLong Millimeters;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MetafilePicture
    {
        public int MappingMode;
        public int XExt;
        public int YExt;
        public IntPtr Metafile;
    }

    private STGMEDIUM CreateMetafilePictureMedium()
    {
        if (_emfPath is null || !File.Exists(_emfPath))
            Marshal.ThrowExceptionForHR(DV_E_FORMATETC);

        var enhanced = GetEnhMetaFileW(_emfPath!);
        if (enhanced == IntPtr.Zero)
            throw new InvalidDataException(
                $"VisualTeX could not open the MathType replacement EMF: {_emfPath}");

        IntPtr recordingDc = IntPtr.Zero;
        IntPtr metafile = IntPtr.Zero;
        IntPtr metafilePicture = IntPtr.Zero;
        try
        {
            var header = new EnhancedMetafileHeader();
            var headerSize = (uint)Marshal.SizeOf<EnhancedMetafileHeader>();
            if (GetEnhMetaFileHeader(enhanced, headerSize, ref header) == 0)
                throw new InvalidDataException("VisualTeX could not read the replacement EMF header.");
            var frameWidth = Math.Max(1, header.Frame.Right - header.Frame.Left);
            var frameHeight = Math.Max(1, header.Frame.Bottom - header.Frame.Top);
            if (header.Device.Cx <= 0 || header.Device.Cy <= 0
                || header.Millimeters.Cx <= 0 || header.Millimeters.Cy <= 0)
                throw new InvalidDataException("VisualTeX replacement EMF has invalid device metrics.");

            // rclFrame is in 0.01 mm. Reconstruct the exact pixel canvas used by
            // the source recording from szlDevice/szlMillimeters rather than
            // allowing GetWinMetaFileBits to infer a printer-dependent transform.
            var canvasWidth = Math.Max(
                1,
                (int)Math.Round(
                    frameWidth / 100d * header.Device.Cx / header.Millimeters.Cx));
            var canvasHeight = Math.Max(
                1,
                (int)Math.Round(
                    frameHeight / 100d * header.Device.Cy / header.Millimeters.Cy));

            recordingDc = CreateMetaFileW(null);
            if (recordingDc == IntPtr.Zero)
                throw new InvalidOperationException("VisualTeX could not create an in-memory WMF DC.");
            SetMapMode(recordingDc, MmAnisotropic);
            if (!SetWindowExtEx(recordingDc, canvasWidth, canvasHeight, IntPtr.Zero)
                || !SetViewportExtEx(recordingDc, canvasWidth, canvasHeight, IntPtr.Zero))
                throw new InvalidOperationException("VisualTeX could not configure the WMF logical canvas.");
            var destination = new RectLong
            {
                Left = 0,
                Top = 0,
                Right = canvasWidth,
                Bottom = canvasHeight,
            };
            if (!PlayEnhMetaFile(recordingDc, enhanced, ref destination))
                throw new InvalidDataException("VisualTeX could not replay the EMF into the WMF canvas.");
            metafile = CloseMetaFile(recordingDc);
            recordingDc = IntPtr.Zero;
            if (metafile == IntPtr.Zero)
                throw new InvalidDataException("VisualTeX could not finalize the replacement WMF.");

            metafilePicture = GlobalAlloc(
                GmemMoveable | GmemZeroInit,
                new UIntPtr((uint)Marshal.SizeOf<MetafilePicture>()));
            if (metafilePicture == IntPtr.Zero)
                throw new OutOfMemoryException("GlobalAlloc failed for MathType METAFILEPICT.");
            var locked = GlobalLock(metafilePicture);
            if (locked == IntPtr.Zero)
                throw new OutOfMemoryException("GlobalLock failed for MathType METAFILEPICT.");
            try
            {
                var picture = new MetafilePicture
                {
                    MappingMode = MmAnisotropic,
                    XExt = frameWidth,
                    YExt = frameHeight,
                    Metafile = metafile,
                };
                Marshal.StructureToPtr(picture, locked, false);
            }
            finally { GlobalUnlock(metafilePicture); }

            metafile = IntPtr.Zero;
            var result = new STGMEDIUM
            {
                tymed = TYMED.TYMED_MFPICT,
                unionmember = metafilePicture,
                pUnkForRelease = null,
            };
            metafilePicture = IntPtr.Zero;
            return result;
        }
        finally
        {
            if (recordingDc != IntPtr.Zero)
            {
                var abandoned = CloseMetaFile(recordingDc);
                if (abandoned != IntPtr.Zero) DeleteMetaFile(abandoned);
            }
            DeleteEnhMetaFile(enhanced);
            if (metafile != IntPtr.Zero) DeleteMetaFile(metafile);
            if (metafilePicture != IntPtr.Zero) GlobalFree(metafilePicture);
        }
    }

    private STGMEDIUM CreateEnhancedMetafileMedium()
    {
        if (_emfPath is null || !File.Exists(_emfPath))
            Marshal.ThrowExceptionForHR(DV_E_FORMATETC);
        var metafile = GetEnhMetaFileW(_emfPath!);
        if (metafile == IntPtr.Zero)
            throw new InvalidDataException(
                $"VisualTeX could not open the MathType replacement EMF: {_emfPath}");
        return new STGMEDIUM
        {
            tymed = TYMED.TYMED_ENHMF,
            unionmember = metafile,
            pUnkForRelease = null,
        };
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern uint RegisterClipboardFormatW(string format);

    [DllImport("shlwapi.dll")]
    private static extern IntPtr SHCreateMemStream(byte[] initial, uint length);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GlobalAlloc(uint flags, UIntPtr bytes);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GlobalLock(IntPtr memory);

    [DllImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GlobalUnlock(IntPtr memory);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GlobalFree(IntPtr memory);

    [DllImport("kernel32.dll")]
    private static extern UIntPtr GlobalSize(IntPtr memory);

    [DllImport("ole32.dll")]
    private static extern void ReleaseStgMedium(ref STGMEDIUM medium);

    [DllImport("user32.dll")]
    private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr dpiContext);

    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr window);

    [DllImport("user32.dll")]
    private static extern int ReleaseDC(IntPtr window, IntPtr deviceContext);

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr GetEnhMetaFileW(string fileName);

    [DllImport("gdi32.dll")]
    private static extern uint GetEnhMetaFileHeader(
        IntPtr enhancedMetafile,
        uint bufferSize,
        ref EnhancedMetafileHeader header);

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateMetaFileW(string? fileName);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CloseMetaFile(IntPtr deviceContext);

    [DllImport("gdi32.dll")]
    private static extern int SetMapMode(IntPtr deviceContext, int mode);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetWindowExtEx(
        IntPtr deviceContext,
        int x,
        int y,
        IntPtr previousSize);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetViewportExtEx(
        IntPtr deviceContext,
        int x,
        int y,
        IntPtr previousSize);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PlayEnhMetaFile(
        IntPtr deviceContext,
        IntPtr enhancedMetafile,
        ref RectLong destination);

    [DllImport("gdi32.dll")]
    private static extern uint GetWinMetaFileBits(
        IntPtr enhancedMetafile,
        uint bufferSize,
        [Out] byte[]? data,
        int mapMode,
        IntPtr referenceDeviceContext);

    [DllImport("gdi32.dll")]
    private static extern IntPtr SetMetaFileBitsEx(uint byteCount, byte[] data);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeleteEnhMetaFile(IntPtr enhancedMetafile);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeleteMetaFile(IntPtr metafile);
}
