using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WordVsto;

[ComVisible(true)]
[Guid("00000118-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IMathTypeOleClientSiteNative
{
    [PreserveSig] int SaveObject();
    [PreserveSig] int GetMoniker(uint assign, uint whichMoniker, out IntPtr moniker);
    [PreserveSig] int GetContainer(out IntPtr container);
    [PreserveSig] int ShowObject();
    [PreserveSig] int OnShowWindow([MarshalAs(UnmanagedType.Bool)] bool show);
    [PreserveSig] int RequestNewObjectLayout();
}

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
public sealed class MathTypeOleClientSiteNative : IMathTypeOleClientSiteNative
{
    public int SaveObject() => 0;

    public int GetMoniker(uint assign, uint whichMoniker, out IntPtr moniker)
    {
        moniker = IntPtr.Zero;
        return unchecked((int)0x80004001); // E_NOTIMPL
    }

    public int GetContainer(out IntPtr container)
    {
        container = IntPtr.Zero;
        return unchecked((int)0x80004002); // E_NOINTERFACE
    }

    public int ShowObject() => 0;
    public int OnShowWindow(bool show) => 0;
    public int RequestNewObjectLayout() => unchecked((int)0x80004001); // E_NOTIMPL
}

internal static class MathTypeOleInterop
{
    private const uint GmemMoveable = 0x0002;
    private const uint GmemZeroInit = 0x0040;
    private const uint OleCloseNoSave = 1;
    private const int DvAspectContent = 1;
    private const int MaxMathMlBytes = 16 * 1024 * 1024;
    private static readonly string[] MathMlFormats =
    {
        "MathML",
        "MathML Presentation",
        "application/mathml+xml",
    };

    internal sealed class Capabilities
    {
        public string ProgId { get; set; } = string.Empty;
        public Guid ResolvedClsid { get; set; }
        public string? ServerPath { get; set; }
        public int RunForConversionVerb { get; set; } = 2;
        public bool RegisteredMathMlGetSet { get; set; }
    }

    internal static bool IsMathTypeOle(InlineShape shape)
    {
        OLEFormat? format = null;
        try
        {
            if (shape.Type is not WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                and not WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                return false;

            // Fast path when Word can expose the stored ProgID directly.
            try
            {
                format = shape.OLEFormat;
                var progId = format.ProgID;
                if (string.Equals(progId, "Equation.DSMT4", StringComparison.OrdinalIgnoreCase))
                    return true;
                if (TryResolveCapabilities(progId, out _))
                    return true;
            }
            catch
            {
                // A missing MathType registration can make OLEFormat access fail.
                // Fall through to the serialized WordOpenXML/CFB identity below.
            }
            finally
            {
                Release(format);
                format = null;
            }

            try
            {
                var fragment = MathTypeWordOpenXml.Read(shape);
                return string.Equals(fragment.ProgId, "Equation.DSMT4", StringComparison.OrdinalIgnoreCase)
                    || fragment.ProgId.StartsWith("Equation.", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }
        catch
        {
            return false;
        }
        finally { Release(format); }
    }

    internal static bool TryResolveCapabilities(string? progId, out Capabilities capabilities)
    {
        capabilities = new Capabilities();
        if (string.IsNullOrWhiteSpace(progId)
            || !(string.Equals(progId, "Equation", StringComparison.OrdinalIgnoreCase)
                || progId.StartsWith("Equation.", StringComparison.OrdinalIgnoreCase)))
            return false;

        if (CLSIDFromProgID(progId, out var current) != 0)
            return false;

        var visited = new HashSet<Guid>();
        for (var depth = 0; depth < 16 && visited.Add(current); depth++)
        {
            var hr = OleGetAutoConvert(ref current, out var next);
            if (hr != 0 || next == Guid.Empty || next == current)
                break;
            current = next;
        }

        var clsidKey = "CLSID\\{" + current.ToString("D") + "}";
        using var classKey = Registry.ClassesRoot.OpenSubKey(clsidKey);
        if (classKey is null) return false;
        var className = classKey.GetValue(null) as string;
        var serverPath = ReadDefaultString(classKey, "LocalServer32")
            ?? ReadDefaultString(classKey, "LocalServer");
        var serverFileName = string.IsNullOrWhiteSpace(serverPath)
            ? null
            : Path.GetFileName(serverPath.Trim().Trim('"'));
        var isMathType = !string.IsNullOrWhiteSpace(className)
                && className.IndexOf("MathType", StringComparison.OrdinalIgnoreCase) >= 0
            || string.Equals(serverFileName, "MathType.exe", StringComparison.OrdinalIgnoreCase);
        if (!isMathType) return false;

        var supportsMathMl = false;
        using (var formats = classKey.OpenSubKey("DataFormats\\GetSet"))
        {
            if (formats is not null)
            {
                foreach (var name in formats.GetSubKeyNames())
                {
                    using var item = formats.OpenSubKey(name);
                    var value = item?.GetValue(null) as string;
                    if (value is null) continue;
                    if (MathMlFormats.Any(format =>
                            value.StartsWith(format + ",", StringComparison.OrdinalIgnoreCase)))
                    {
                        supportsMathMl = true;
                        break;
                    }
                }
            }
        }

        var conversionVerb = 2;
        using (var verbs = classKey.OpenSubKey("Verb"))
        {
            if (verbs is not null)
            {
                foreach (var name in verbs.GetSubKeyNames())
                {
                    using var item = verbs.OpenSubKey(name);
                    var value = item?.GetValue(null) as string;
                    if (value is null
                        || value.IndexOf("RunForConversion", StringComparison.OrdinalIgnoreCase) < 0
                        || !int.TryParse(name, out var parsed))
                        continue;
                    conversionVerb = parsed;
                    break;
                }
            }
        }

        capabilities = new Capabilities
        {
            ProgId = progId,
            ResolvedClsid = current,
            ServerPath = serverPath,
            RunForConversionVerb = conversionVerb,
            RegisteredMathMlGetSet = supportsMathMl,
        };
        return true;
    }

    internal static FormulaMetadata ReadMetadata(
        Microsoft.Office.Interop.Word.Application application,
        InlineShape shape)
    {
        string mathMl;
        try
        {
            // Preferred path: Word already owns the complete Equation.DSMT4 CFB.
            // Parse Equation Native/MTEF directly so VisualTeX works even when
            // MathType is not installed on this machine.
            mathMl = MathTypeOleStorage.ReadMathMl(shape);
        }
        catch (Exception directError)
        {
            // Compatibility fallback only for an MTEF structure the direct parser
            // does not support yet and only when a MathType server is actually
            // installed. Never make MathType a requirement for supported formulas.
            OLEFormat? fallbackFormat = null;
            try
            {
                fallbackFormat = shape.OLEFormat;
                if (!TryResolveCapabilities(fallbackFormat.ProgID, out _))
                    throw new InvalidDataException(
                        "VisualTeX could not parse this MathType OLE formula directly, and no MathType server is installed for compatibility fallback.",
                        directError);
            }
            finally { Release(fallbackFormat); }
            mathMl = ReadMathMl(shape);
        }
        var latex = MathMlToLatexConverter.Convert(mathMl).Trim();
        if (string.IsNullOrWhiteSpace(latex))
            throw new InvalidDataException("MathType OLE returned MathML that VisualTeX could not convert to LaTeX.");

        Range? range = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        var displayMode = "inline";
        var fontSizePt = FormulaFontSize.DefaultPt;
        try
        {
            range = shape.Range;
            displayMode = InferDisplayMode(range);
            font = range.Font;
            if (font.Size > 0 && font.Size <= 200)
                fontSizePt = FormulaFontSize.Normalize(font.Size);
        }
        catch { }
        finally
        {
            Release(font);
            Release(range);
        }

        var now = DateTimeOffset.UtcNow.ToString("O");
        var metadata = new FormulaMetadata
        {
            FormulaId = Guid.NewGuid().ToString("D"),
            Title = "MathType Formula",
            Latex = latex,
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
            },
            CodeFormat = "raw",
            DisplayMode = displayMode,
            Numbered = false,
            FontSizePt = fontSizePt,
            RenderFontSizePt = fontSizePt,
            CreatedWithVersion = "1.2.5",
            UpdatedWithVersion = "1.2.5",
            CreatedAt = now,
            UpdatedAt = now,
        };
        metadata.Validate();
        return metadata;
    }

    internal static string ReadMathMl(InlineShape shape, bool activateForConversion = true)
    {
        OLEFormat? format = null;
        object? runningObject = null;
        try
        {
            format = shape.OLEFormat;
            if (!TryResolveCapabilities(format.ProgID, out var capabilities))
                throw new InvalidOperationException("The selected OLE object is not backed by an installed MathType server.");
            if (!capabilities.RegisteredMathMlGetSet)
                throw new InvalidOperationException("The installed MathType OLE server does not register a MathML Get/Set format.");

            runningObject = activateForConversion
                ? GetRunningMathTypeObject(format, capabilities)
                : TryGetObject(format);
            if (runningObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                throw new InvalidOperationException("MathType OLE does not expose IDataObject.");
            return ReadMathMlFromDataObject(dataObject);
        }
        finally
        {
            CloseOleObject(runningObject);
            Release(runningObject);
            Release(format);
        }
    }

    internal static string ProbeStandaloneSetFormats(string progId = "Equation.DSMT4")
    {
        var probes = new (string Name, string Payload)[]
        {
            ("MathML", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mn>1</mn></math>"),
            ("MathML Presentation", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mn>1</mn></math>"),
            ("application/mathml+xml", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mn>1</mn></math>"),
            ("TeX Input Language", "x+1"),
        };
        var results = new List<string>();
        foreach (var probe in probes)
        {
            object? serverObject = null;
            try
            {
                serverObject = CreateStandaloneMathTypeObject(progId);
                if (serverObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                {
                    results.Add($"{probe.Name}=no-idataobject");
                    continue;
                }
                var formatId = RegisterClipboardFormat(probe.Name);
                if (formatId == 0 || formatId > ushort.MaxValue)
                {
                    results.Add($"{probe.Name}=register-failed");
                    continue;
                }
                var request = CreateFormatEtc(unchecked((short)formatId));
                var global = Marshal.StringToHGlobalAnsi(probe.Payload);
                try
                {
                    var medium = new STGMEDIUM
                    {
                        tymed = TYMED.TYMED_HGLOBAL,
                        unionmember = global,
                        pUnkForRelease = null,
                    };
                    try
                    {
                        dataObject.SetData(ref request, ref medium, false);
                        results.Add($"{probe.Name}=S_OK");
                    }
                    catch (COMException error)
                    {
                        results.Add($"{probe.Name}=0x{error.HResult:X8}");
                    }
                }
                finally { Marshal.FreeHGlobal(global); }
            }
            finally
            {
                CloseOleObject(serverObject);
                Release(serverObject);
            }
        }

        // CF_TEXT is not a registered custom format; probe it separately because
        // MathType advertises it in DATADIR_SET even on a fresh equation object.
        object? textServer = null;
        try
        {
            textServer = CreateStandaloneMathTypeObject(progId);
            if (textServer is System.Runtime.InteropServices.ComTypes.IDataObject textDataObject)
            {
                var request = CreateFormatEtc(1); // CF_TEXT
                var global = Marshal.StringToHGlobalAnsi("x+1");
                try
                {
                    var medium = new STGMEDIUM
                    {
                        tymed = TYMED.TYMED_HGLOBAL,
                        unionmember = global,
                        pUnkForRelease = null,
                    };
                    try
                    {
                        textDataObject.SetData(ref request, ref medium, false);
                        results.Add("CF_TEXT=S_OK");
                    }
                    catch (COMException error)
                    {
                        results.Add($"CF_TEXT=0x{error.HResult:X8}");
                    }
                }
                finally { Marshal.FreeHGlobal(global); }
            }
        }
        finally
        {
            CloseOleObject(textServer);
            Release(textServer);
        }
        return string.Join(" | ", results);
    }

    internal static string DescribeStandaloneServerFormats(string progId = "Equation.DSMT4")
    {
        object? serverObject = null;
        try
        {
            serverObject = CreateStandaloneMathTypeObject(progId);
            if (serverObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                return "no-idataobject";
            var parts = new List<string>();
            AppendFormats(parts, dataObject, DATADIR.DATADIR_GET, "GET");
            AppendFormats(parts, dataObject, DATADIR.DATADIR_SET, "SET");
            foreach (var name in MathMlFormats)
            {
                var id = RegisterClipboardFormat(name);
                if (id == 0 || id > ushort.MaxValue) continue;
                var request = CreateFormatEtc(unchecked((short)id));
                parts.Add($"QUERY {name}({id})=0x{dataObject.QueryGetData(ref request):X8}");
            }
            return string.Join(" | ", parts);
        }
        finally
        {
            CloseOleObject(serverObject);
            Release(serverObject);
        }
    }

    internal static string RoundTripStandaloneMathMl(
        string mathMl,
        string progId = "Equation.DSMT4")
    {
        object? serverObject = null;
        IntPtr lockBytes = IntPtr.Zero;
        IntPtr storage = IntPtr.Zero;
        try
        {
            serverObject = CreateStandaloneMathTypeObject(progId);
            InitializeStandaloneOleStorage(serverObject, out lockBytes, out storage);
            if (serverObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                throw new InvalidOperationException("Standalone MathType OLE object does not expose IDataObject.");
            WriteMathMlToDataObject(dataObject, mathMl);
            return ReadMathMlFromDataObject(dataObject);
        }
        finally
        {
            CloseOleObject(serverObject);
            ReleaseRawComPointer(ref storage);
            ReleaseRawComPointer(ref lockBytes);
            Release(serverObject);
        }
    }

    internal static string DescribeOleCreateClientFormats(
        string progId = "Equation.DSMT4")
    {
        object? oleObject = null;
        IntPtr lockBytes = IntPtr.Zero;
        IntPtr storage = IntPtr.Zero;
        try
        {
            oleObject = CreateContainedMathTypeOleObject(
                progId,
                out lockBytes,
                out storage,
                out _);
            if (oleObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                return "no-idataobject";
            var parts = new List<string>();
            AppendFormats(parts, dataObject, DATADIR.DATADIR_GET, "GET");
            AppendFormats(parts, dataObject, DATADIR.DATADIR_SET, "SET");
            foreach (var name in MathMlFormats)
            {
                var id = RegisterClipboardFormat(name);
                if (id == 0 || id > ushort.MaxValue) continue;
                var request = CreateFormatEtc(unchecked((short)id));
                parts.Add($"QUERY {name}({id})=0x{dataObject.QueryGetData(ref request):X8}");
            }
            return string.Join(" | ", parts);
        }
        finally
        {
            CloseOleObject(oleObject);
            ReleaseRawComPointer(ref storage);
            ReleaseRawComPointer(ref lockBytes);
            Release(oleObject);
        }
    }

    internal static string DescribeOleCreateAttachedDataFormats(
        string progId = "Equation.DSMT4")
    {
        object? oleObject = null;
        object? attachedDataObject = null;
        IntPtr lockBytes = IntPtr.Zero;
        IntPtr storage = IntPtr.Zero;
        try
        {
            oleObject = CreateContainedMathTypeOleObject(
                progId,
                out lockBytes,
                out storage,
                out _);
            attachedDataObject = GetOleClipboardDataObject(oleObject);
            if (attachedDataObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                return "no-idataobject";
            var parts = new List<string>();
            AppendFormats(parts, dataObject, DATADIR.DATADIR_GET, "GET");
            AppendFormats(parts, dataObject, DATADIR.DATADIR_SET, "SET");
            foreach (var name in MathMlFormats)
            {
                var id = RegisterClipboardFormat(name);
                if (id == 0 || id > ushort.MaxValue) continue;
                var request = CreateFormatEtc(unchecked((short)id));
                parts.Add($"QUERY {name}({id})=0x{dataObject.QueryGetData(ref request):X8}");
            }
            return string.Join(" | ", parts);
        }
        finally
        {
            Release(attachedDataObject);
            CloseOleObject(oleObject);
            ReleaseRawComPointer(ref storage);
            ReleaseRawComPointer(ref lockBytes);
            Release(oleObject);
        }
    }

    internal static string RoundTripOleCreateMathMl(
        string mathMl,
        string progId = "Equation.DSMT4")
    {
        object? oleObject = null;
        object? attachedDataObject = null;
        IntPtr lockBytes = IntPtr.Zero;
        IntPtr storage = IntPtr.Zero;
        try
        {
            oleObject = CreateContainedMathTypeOleObject(
                progId,
                out lockBytes,
                out storage,
                out var clientSite);
            if (oleObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                throw new InvalidOperationException(
                    "OleCreate MathType object does not expose IDataObject.");
            WriteMathMlToDataObject(dataObject, mathMl);
            ConvertContainedMathTypeOleObject(oleObject, clientSite);
            SaveContainedMathTypeOleObject(oleObject, storage);
            return ReadMathMlFromDataObject(dataObject);
        }
        finally
        {
            Release(attachedDataObject);
            CloseOleObject(oleObject);
            ReleaseRawComPointer(ref storage);
            ReleaseRawComPointer(ref lockBytes);
            Release(oleObject);
        }
    }

    internal static string DescribeInitializedStandaloneServerFormats(
        string progId = "Equation.DSMT4")
    {
        object? serverObject = null;
        IntPtr lockBytes = IntPtr.Zero;
        IntPtr storage = IntPtr.Zero;
        try
        {
            serverObject = CreateStandaloneMathTypeObject(progId);
            InitializeStandaloneOleStorage(serverObject, out lockBytes, out storage);
            if (serverObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                return "no-idataobject";
            var parts = new List<string>();
            AppendFormats(parts, dataObject, DATADIR.DATADIR_GET, "GET");
            AppendFormats(parts, dataObject, DATADIR.DATADIR_SET, "SET");
            foreach (var name in MathMlFormats)
            {
                var id = RegisterClipboardFormat(name);
                if (id == 0 || id > ushort.MaxValue) continue;
                var request = CreateFormatEtc(unchecked((short)id));
                parts.Add($"QUERY {name}({id})=0x{dataObject.QueryGetData(ref request):X8}");
            }
            return string.Join(" | ", parts);
        }
        finally
        {
            CloseOleObject(serverObject);
            ReleaseRawComPointer(ref storage);
            ReleaseRawComPointer(ref lockBytes);
            Release(serverObject);
        }
    }

    internal static string DescribeDataFormats(InlineShape shape, bool activateForConversion = true)
    {
        OLEFormat? format = null;
        object? runningObject = null;
        try
        {
            format = shape.OLEFormat;
            if (!TryResolveCapabilities(format.ProgID, out var capabilities))
                return "not-mathtype";
            runningObject = activateForConversion
                ? GetRunningMathTypeObject(format, capabilities)
                : TryGetObject(format);
            if (runningObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                return "no-idataobject";
            var parts = new List<string>();
            AppendFormats(parts, dataObject, DATADIR.DATADIR_GET, "GET");
            AppendFormats(parts, dataObject, DATADIR.DATADIR_SET, "SET");
            foreach (var name in MathMlFormats)
            {
                var id = RegisterClipboardFormat(name);
                if (id == 0 || id > ushort.MaxValue) continue;
                var request = CreateFormatEtc(unchecked((short)id));
                parts.Add($"QUERY {name}({id})=0x{dataObject.QueryGetData(ref request):X8}");
            }
            return string.Join(" | ", parts);
        }
        finally
        {
            CloseOleObject(runningObject);
            Release(runningObject);
            Release(format);
        }
    }

    internal static string ProbeExistingOleMathMlSetVariants(
        InlineShape shape,
        string mathMl)
    {
        OLEFormat? format = null;
        object? runningObject = null;
        try
        {
            format = shape.OLEFormat;
            if (!TryResolveCapabilities(format.ProgID, out var capabilities))
                throw new InvalidOperationException(
                    "The selected OLE object is not backed by an installed MathType server.");
            runningObject = GetRunningMathTypeObject(format, capabilities);
            if (runningObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                throw new InvalidOperationException("MathType OLE does not expose IDataObject.");

            var results = new List<string>();
            foreach (var formatName in MathMlFormats)
            {
                var clipboardId = RegisterClipboardFormat(formatName);
                if (clipboardId == 0 || clipboardId > ushort.MaxValue) continue;
                var request = CreateFormatEtc(unchecked((short)clipboardId));
                results.Add($"{formatName}:Query=0x{dataObject.QueryGetData(ref request):X8}");
                results.Add(formatName + ":Marshal/keep="
                    + ProbeSetDataWithMarshalHGlobal(dataObject, request, mathMl, false));
                results.Add(formatName + ":GlobalMoveable/keep="
                    + ProbeSetDataWithMoveableHGlobal(dataObject, request, mathMl, false));
                results.Add(formatName + ":GlobalMoveable/release="
                    + ProbeSetDataWithMoveableHGlobal(dataObject, request, mathMl, true));
            }
            return string.Join(" | ", results);
        }
        finally
        {
            CloseOleObject(runningObject);
            Release(runningObject);
            Release(format);
        }
    }

    private static string ProbeSetDataWithMarshalHGlobal(
        System.Runtime.InteropServices.ComTypes.IDataObject dataObject,
        FORMATETC request,
        string text,
        bool release)
    {
        var memory = Marshal.StringToHGlobalAnsi(text);
        if (memory == IntPtr.Zero) return "alloc-failed";
        try
        {
            var medium = new STGMEDIUM
            {
                tymed = TYMED.TYMED_HGLOBAL,
                unionmember = memory,
                pUnkForRelease = null,
            };
            try
            {
                dataObject.SetData(ref request, ref medium, release);
                if (release) memory = IntPtr.Zero;
                return "S_OK";
            }
            catch (COMException error)
            {
                return $"0x{error.HResult:X8}";
            }
        }
        finally
        {
            if (memory != IntPtr.Zero) Marshal.FreeHGlobal(memory);
        }
    }

    private static string ProbeSetDataWithMoveableHGlobal(
        System.Runtime.InteropServices.ComTypes.IDataObject dataObject,
        FORMATETC request,
        string text,
        bool release)
    {
        var bytes = Encoding.ASCII.GetBytes(text + "\0");
        const uint GmemMoveable = 0x0002;
        const uint GmemZeroInit = 0x0040;
        var memory = GlobalAlloc(GmemMoveable | GmemZeroInit, new UIntPtr((uint)bytes.Length));
        if (memory == IntPtr.Zero) return "alloc-failed";
        var locked = GlobalLock(memory);
        if (locked == IntPtr.Zero)
        {
            GlobalFree(memory);
            return "lock-failed";
        }
        try { Marshal.Copy(bytes, 0, locked, bytes.Length); }
        finally { GlobalUnlock(memory); }

        var transferred = false;
        try
        {
            var medium = new STGMEDIUM
            {
                tymed = TYMED.TYMED_HGLOBAL,
                unionmember = memory,
                pUnkForRelease = null,
            };
            try
            {
                dataObject.SetData(ref request, ref medium, release);
                transferred = release;
                return "S_OK";
            }
            catch (COMException error)
            {
                return $"0x{error.HResult:X8}";
            }
        }
        finally
        {
            // If SetData succeeds with fRelease=true the server owns the medium.
            // On a failed HRESULT keep caller ownership and release it here.
            if (!transferred) GlobalFree(memory);
        }
    }

    internal static void WriteTextForExistingOle(
        InlineShape shape,
        string text,
        bool activateForConversion = true)
    {
        if (string.IsNullOrWhiteSpace(text))
            throw new InvalidDataException("VisualTeX did not provide formula text for MathType OLE.");

        OLEFormat? format = null;
        object? runningObject = null;
        IntPtr global = IntPtr.Zero;
        try
        {
            format = shape.OLEFormat;
            if (!TryResolveCapabilities(format.ProgID, out var capabilities))
                throw new InvalidOperationException(
                    "The selected OLE object is not backed by an installed MathType server.");
            runningObject = activateForConversion
                ? GetRunningMathTypeObject(format, capabilities)
                : TryGetObject(format);
            if (runningObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                throw new InvalidOperationException("MathType OLE does not expose IDataObject.");

            var request = CreateFormatEtc(1); // CF_TEXT, advertised by MathType 7 for DATADIR_SET.
            global = Marshal.StringToHGlobalAnsi(text + "\0");
            if (global == IntPtr.Zero)
                throw new OutOfMemoryException("Unable to allocate MathType text transfer memory.");
            var medium = new STGMEDIUM
            {
                tymed = TYMED.TYMED_HGLOBAL,
                unionmember = global,
                pUnkForRelease = null,
            };
            dataObject.SetData(ref request, ref medium, false);
        }
        finally
        {
            if (global != IntPtr.Zero) Marshal.FreeHGlobal(global);
            CloseOleObject(runningObject);
            Release(runningObject);
            Release(format);
        }
    }

    internal static void WriteMathMl(InlineShape shape, string mathMl, bool activateForConversion = true)
    {
        if (string.IsNullOrWhiteSpace(mathMl)
            || !mathMl.TrimStart().StartsWith("<math", StringComparison.Ordinal))
            throw new InvalidDataException("VisualTeX did not provide valid MathML for MathType OLE.");

        OLEFormat? format = null;
        object? runningObject = null;
        IntPtr global = IntPtr.Zero;
        var ownershipTransferred = false;
        try
        {
            format = shape.OLEFormat;
            if (!TryResolveCapabilities(format.ProgID, out var capabilities))
                throw new InvalidOperationException("The selected OLE object is not backed by an installed MathType server.");
            runningObject = activateForConversion
                ? GetRunningMathTypeObject(format, capabilities)
                : TryGetObject(format);
            if (runningObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                throw new InvalidOperationException("MathType OLE does not expose IDataObject.");
            WriteMathMlToDataObject(dataObject, mathMl);
            ownershipTransferred = true;
        }
        finally
        {
            if (!ownershipTransferred && global != IntPtr.Zero)
                GlobalFree(global);
            CloseOleObject(runningObject);
            Release(runningObject);
            Release(format);
        }
    }

    private static object GetRunningMathTypeObject(OLEFormat format, Capabilities capabilities)
    {
        if (TryDescribeVisibleStandaloneMathType(out var activeWindow))
        {
            throw new InvalidOperationException(
                "MathType currently has an interactive equation window open. "
                + "Finish or close that MathType editing window before VisualTeX accesses this OLE equation. "
                + $"Active MathType window: {activeWindow}");
        }

        // MathType's documented IDataObject contract requires the OLE server to
        // be started with its custom RunForConversion verb before Object is
        // queried. Word can expose a dormant OLE proxy through format.Object
        // before that transition; it may answer QueryGetData but reject SetData
        // with DV_E_FORMATETC. Always enter MathType's conversion mode first.
        object verb = capabilities.RunForConversionVerb;
        Exception? activationError = null;
        try { format.DoVerb(ref verb); }
        catch (Exception error) when (error is COMException or InvalidCastException)
        {
            activationError = error;
        }

        object? running = null;
        var delayMilliseconds = 20;
        for (var attempt = 0; attempt < 12; attempt++)
        {
            running = TryGetObject(format);
            if (running is not null) return running;
            if (attempt == 11) break;
            Thread.Sleep(delayMilliseconds);
            delayMilliseconds = Math.Min(80, delayMilliseconds + 10);
        }

        throw new COMException(
            "MathType OLE server started but Word did not expose IDataObject."
            + (activationError is null ? string.Empty : $" {activationError.Message}"));
    }

    private static bool TryDescribeVisibleStandaloneMathType(out string description)
    {
        description = string.Empty;
        try
        {
            foreach (var process in Process.GetProcessesByName("MathType"))
            {
                try
                {
                    process.Refresh();
                    if (process.MainWindowHandle == IntPtr.Zero || process.HasExited)
                        continue;
                    var title = process.MainWindowTitle?.Trim();
                    description = string.IsNullOrWhiteSpace(title)
                        ? $"PID {process.Id}"
                        : $"{title} (PID {process.Id})";
                    return true;
                }
                finally { process.Dispose(); }
            }
        }
        catch { }
        return false;
    }

    private static object? TryGetObject(OLEFormat format)
    {
        try { return format.Object; }
        catch (Exception error) when (error is COMException or InvalidCastException)
        {
            return null;
        }
    }

    private static object CreateContainedMathTypeOleObject(
        string progId,
        out IntPtr lockBytes,
        out IntPtr storage,
        out MathTypeOleClientSiteNative clientSite)
    {
        if (!TryResolveCapabilities(progId, out var capabilities))
            throw new InvalidOperationException(
                $"The installed OLE class '{progId}' is not recognized as MathType.");
        lockBytes = IntPtr.Zero;
        storage = IntPtr.Zero;
        clientSite = new MathTypeOleClientSiteNative();
        const uint stgmReadWrite = 0x00000002;
        const uint stgmShareExclusive = 0x00000010;
        const uint stgmCreate = 0x00001000;
        var hr = CreateILockBytesOnHGlobal(IntPtr.Zero, true, out lockBytes);
        if (hr < 0) Marshal.ThrowExceptionForHR(hr);
        hr = StgCreateDocfileOnILockBytes(
            lockBytes,
            stgmReadWrite | stgmShareExclusive | stgmCreate,
            0,
            out storage);
        if (hr < 0) Marshal.ThrowExceptionForHR(hr);

        var iidOleObject = new Guid("00000112-0000-0000-C000-000000000046");
        var clsid = capabilities.ResolvedClsid;
        IntPtr oleObjectPointer = IntPtr.Zero;
        IntPtr clientSitePointer = IntPtr.Zero;
        try
        {
            clientSitePointer = Marshal.GetComInterfaceForObject(
                clientSite,
                typeof(IMathTypeOleClientSiteNative));
            hr = OleCreate(
                ref clsid,
                ref iidOleObject,
                1, // OLERENDER_DRAW: matches COleClientItem::CreateNewItem default used by MathType's MFC sample.
                IntPtr.Zero,
                clientSitePointer,
                storage,
                out oleObjectPointer);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);
            if (oleObjectPointer == IntPtr.Zero)
                throw new COMException("OleCreate returned no MathType IOleObject pointer.");
            var oleObject = Marshal.GetObjectForIUnknown(oleObjectPointer);
            hr = OleRun(oleObject);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);
            return oleObject;
        }
        finally
        {
            if (clientSitePointer != IntPtr.Zero)
                Marshal.Release(clientSitePointer);
            if (oleObjectPointer != IntPtr.Zero)
                Marshal.Release(oleObjectPointer);
        }
    }

    private static object GetOleClipboardDataObject(object oleObject)
    {
        if (oleObject is not IOleObjectNative native)
            throw new InvalidOperationException(
                "MathType OleCreate object does not expose IOleObject.");
        IntPtr dataObjectPointer = IntPtr.Zero;
        try
        {
            var hr = native.GetClipboardData(0, out dataObjectPointer);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);
            if (dataObjectPointer == IntPtr.Zero)
                throw new COMException(
                    "MathType IOleObject.GetClipboardData returned no data object.");
            return Marshal.GetObjectForIUnknown(dataObjectPointer);
        }
        finally
        {
            if (dataObjectPointer != IntPtr.Zero)
                Marshal.Release(dataObjectPointer);
        }
    }

    private static void ConvertContainedMathTypeOleObject(
        object oleObject,
        MathTypeOleClientSiteNative clientSite)
    {
        if (oleObject is not IOleObjectNative native)
            throw new InvalidOperationException("MathType OleCreate object does not expose IOleObject.");
        var rectangle = new OleRect();
        var clientSitePointer = Marshal.GetComInterfaceForObject(
            clientSite,
            typeof(IMathTypeOleClientSiteNative));
        try
        {
            var hr = native.DoVerb(
                2, // MathType kConvert / RunForConversion
                IntPtr.Zero,
                clientSitePointer,
                0,
                IntPtr.Zero,
                ref rectangle);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);
        }
        finally
        {
            Marshal.Release(clientSitePointer);
        }
    }

    private static void SaveContainedMathTypeOleObject(object oleObject, IntPtr storage)
    {
        if (oleObject is not IPersistStorageFull persistStorage)
            throw new InvalidOperationException("MathType OleCreate object does not expose IPersistStorage.");
        var hr = persistStorage.Save(storage, true);
        if (hr < 0) Marshal.ThrowExceptionForHR(hr);
        hr = persistStorage.SaveCompleted(storage);
        if (hr < 0) Marshal.ThrowExceptionForHR(hr);
    }

    private static void InitializeStandaloneOleStorage(
        object serverObject,
        out IntPtr lockBytes,
        out IntPtr storage)
    {
        lockBytes = IntPtr.Zero;
        storage = IntPtr.Zero;
        const uint stgmReadWrite = 0x00000002;
        const uint stgmShareExclusive = 0x00000010;
        const uint stgmCreate = 0x00001000;
        var hr = CreateILockBytesOnHGlobal(IntPtr.Zero, true, out lockBytes);
        if (hr < 0)
            Marshal.ThrowExceptionForHR(hr);
        hr = StgCreateDocfileOnILockBytes(
            lockBytes,
            stgmReadWrite | stgmShareExclusive | stgmCreate,
            0,
            out storage);
        if (hr < 0)
            Marshal.ThrowExceptionForHR(hr);
        if (serverObject is not IPersistStorageInit persistStorage)
            throw new InvalidOperationException(
                "Standalone MathType OLE object does not implement IPersistStorage.");
        hr = persistStorage.InitNew(storage);
        if (hr < 0)
            Marshal.ThrowExceptionForHR(hr);
        hr = OleRun(serverObject);
        if (hr < 0)
            Marshal.ThrowExceptionForHR(hr);
    }

    private static void ReleaseRawComPointer(ref IntPtr pointer)
    {
        if (pointer == IntPtr.Zero) return;
        try { Marshal.Release(pointer); } catch { }
        pointer = IntPtr.Zero;
    }

    private static object CreateStandaloneMathTypeObject(string progId)
    {
        if (!TryResolveCapabilities(progId, out var capabilities))
            throw new InvalidOperationException(
                $"The installed OLE class '{progId}' is not recognized as MathType.");
        var type = Type.GetTypeFromCLSID(capabilities.ResolvedClsid, throwOnError: true)
            ?? throw new COMException("Unable to resolve the installed MathType OLE COM class.");
        return Activator.CreateInstance(type)
            ?? throw new COMException("Unable to create a standalone MathType OLE COM object.");
    }

    private static string ReadMathMlFromDataObject(
        System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
    {
        var clipboardFormat = ResolveSupportedMathMlFormat(dataObject);
        var request = CreateFormatEtc(clipboardFormat);
        dataObject.GetData(ref request, out var medium);
        try { return ReadMathMlFromMedium(medium); }
        finally { ReleaseStgMedium(ref medium); }
    }

    private static void WriteMathMlToDataObject(
        System.Runtime.InteropServices.ComTypes.IDataObject dataObject,
        string mathMl)
    {
        var request = ResolveSupportedMathMlSetFormat(dataObject);

        // Match MathType's published .NET sample as closely as possible. Its
        // SetData path explicitly uses the CLR Marshal allocator for the HGLOBAL
        // payload instead of constructing a movable GlobalAlloc block manually.
        // Keep ownership in the caller (release=false) and free it after SetData.
        var global = Marshal.StringToHGlobalAnsi(mathMl);
        if (global == IntPtr.Zero)
            throw new OutOfMemoryException("Unable to allocate MathType MathML transfer memory.");
        try
        {
            var medium = new STGMEDIUM
            {
                tymed = TYMED.TYMED_HGLOBAL,
                unionmember = global,
                pUnkForRelease = null,
            };
            dataObject.SetData(ref request, ref medium, false);
        }
        finally
        {
            Marshal.FreeHGlobal(global);
        }
    }

    private static FORMATETC ResolveSupportedMathMlSetFormat(
        System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
    {
        var setFormats = EnumerateFormats(dataObject, DATADIR.DATADIR_SET);
        foreach (var name in MathMlFormats)
        {
            var id = RegisterClipboardFormat(name);
            if (id == 0 || id > ushort.MaxValue) continue;
            var signedId = unchecked((short)id);
            var advertised = setFormats.FirstOrDefault(item =>
                item.cfFormat == signedId
                && (item.tymed & TYMED.TYMED_HGLOBAL) != 0
                && item.dwAspect == (DVASPECT)DvAspectContent);
            if (advertised.cfFormat == signedId)
            {
                advertised.ptd = IntPtr.Zero;
                advertised.lindex = -1;
                advertised.tymed = TYMED.TYMED_HGLOBAL;
                return advertised;
            }
        }

        // MathType's documented programmatic-insertion flow does not require
        // QueryGetData before SetData. A newly-created OLE client item can reject
        // QueryGetData(MathML) until conversion while still accepting SetData for
        // the registered "MathML" format. Use that exact documented format here.
        var mathMlFormat = RegisterClipboardFormat("MathML");
        if (mathMlFormat == 0 || mathMlFormat > ushort.MaxValue)
            throw new InvalidOperationException("Windows did not register MathType's MathML clipboard format.");
        return CreateFormatEtc(unchecked((short)mathMlFormat));
    }

    private static short ResolveSupportedMathMlFormat(System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
    {
        foreach (var name in MathMlFormats)
        {
            var id = RegisterClipboardFormat(name);
            if (id == 0 || id > ushort.MaxValue) continue;
            var request = CreateFormatEtc(unchecked((short)id));
            if (dataObject.QueryGetData(ref request) == 0)
                return unchecked((short)id);
        }
        throw new InvalidOperationException("This MathType OLE object does not expose a supported MathML clipboard format.");
    }

    private static List<FORMATETC> EnumerateFormats(
        System.Runtime.InteropServices.ComTypes.IDataObject dataObject,
        DATADIR direction)
    {
        var result = new List<FORMATETC>();
        IEnumFORMATETC? enumerator = null;
        try
        {
            enumerator = dataObject.EnumFormatEtc(direction);
            if (enumerator is null) return result;
            var buffer = new FORMATETC[1];
            var fetched = new int[1];
            while (enumerator.Next(1, buffer, fetched) == 0 && fetched[0] == 1)
            {
                result.Add(buffer[0]);
                buffer[0] = default;
                fetched[0] = 0;
            }
            return result;
        }
        catch (COMException)
        {
            return result;
        }
        finally
        {
            Release(enumerator);
        }
    }

    private static void AppendFormats(
        List<string> parts,
        System.Runtime.InteropServices.ComTypes.IDataObject dataObject,
        DATADIR direction,
        string label)
    {
        foreach (var item in EnumerateFormats(dataObject, direction))
        {
            var unsignedFormat = unchecked((ushort)item.cfFormat);
            var name = ClipboardFormatName(unsignedFormat);
            parts.Add($"{label} {name}({unsignedFormat}) aspect={(int)item.dwAspect} lindex={item.lindex} tymed=0x{(int)item.tymed:X}");
        }
    }

    private static string ClipboardFormatName(uint format)
    {
        if (format < 0xC000) return $"CF_{format}";
        var builder = new StringBuilder(256);
        var length = GetClipboardFormatName(format, builder, builder.Capacity);
        return length > 0 ? builder.ToString() : $"CF_{format}";
    }

    private static FORMATETC CreateFormatEtc(short clipboardFormat) => new()
    {
        cfFormat = clipboardFormat,
        dwAspect = (DVASPECT)DvAspectContent,
        lindex = -1,
        ptd = IntPtr.Zero,
        tymed = TYMED.TYMED_HGLOBAL,
    };

    private static string ReadMathMlFromMedium(STGMEDIUM medium)
    {
        if (medium.tymed != TYMED.TYMED_HGLOBAL || medium.unionmember == IntPtr.Zero)
            throw new InvalidDataException("MathType MathML was not returned as an HGLOBAL payload.");
        var sizeValue = GlobalSize(medium.unionmember).ToUInt64();
        if (sizeValue == 0 || sizeValue > MaxMathMlBytes)
            throw new InvalidDataException("MathType MathML payload size is invalid.");
        var size = checked((int)sizeValue);
        var locked = GlobalLock(medium.unionmember);
        if (locked == IntPtr.Zero)
            throw new InvalidOperationException("Unable to lock MathType MathML payload.");
        byte[] bytes;
        try
        {
            bytes = new byte[size];
            Marshal.Copy(locked, bytes, 0, size);
        }
        finally { GlobalUnlock(medium.unionmember); }

        string decoded;
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
        {
            var length = FindUtf16Terminator(bytes, 2);
            decoded = Encoding.Unicode.GetString(bytes, 2, Math.Max(0, length - 2));
        }
        else if (bytes.Length >= 2 && bytes[0] == (byte)'<' && bytes[1] == 0)
        {
            var length = FindUtf16Terminator(bytes, 0);
            decoded = Encoding.Unicode.GetString(bytes, 0, length);
        }
        else
        {
            var length = Array.IndexOf(bytes, (byte)0);
            if (length < 0) length = bytes.Length;
            try { decoded = new UTF8Encoding(false, true).GetString(bytes, 0, length); }
            catch (DecoderFallbackException)
            {
                decoded = Encoding.Default.GetString(bytes, 0, length);
            }
        }

        decoded = decoded.Trim().TrimStart('\uFEFF');
        try
        {
            // MathType commonly emits an XML declaration + translator comment and
            // a namespace-prefixed root such as <mml:math>. Never key detection
            // off the literal '<math' spelling; XML namespace prefixes are not
            // semantically significant.
            var document = XDocument.Parse(decoded, LoadOptions.None);
            var root = document.Root;
            if (root is null
                || !string.Equals(root.Name.LocalName, "math", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "MathType returned XML whose root element is not MathML <math>.");
            return root.ToString(SaveOptions.DisableFormatting);
        }
        catch (System.Xml.XmlException error)
        {
            throw new InvalidDataException(
                "MathType returned malformed MathML XML.",
                error);
        }
    }

    private static int FindUtf16Terminator(byte[] bytes, int start)
    {
        for (var index = start; index + 1 < bytes.Length; index += 2)
        {
            if (bytes[index] == 0 && bytes[index + 1] == 0)
                return index;
        }
        return bytes.Length - (bytes.Length - start) % 2;
    }

    private static string InferDisplayMode(Range range)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = range.Paragraphs;
            if (paragraphs.Count != 1) return "inline";
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            // MathType's right-numbered display equation does not expose the
            // equation number as ordinary paragraph text.  The number is owned by
            // a MACROBUTTON MTPlaceRef field (with nested SEQ MTEqn/MTSec fields),
            // and Word commonly exposes only the field-end control character
            // U+0015 in Paragraph.Range.Text.  If that field shares the paragraph
            // with the MathType OLE, this is unambiguously a display equation.
            if (ContainsMathTypeDisplayNumberField(paragraphRange)) return "block";

            var text = (paragraphRange.Text ?? string.Empty)
                .Replace("\r", string.Empty)
                .Replace("\a", string.Empty)
                .Replace("\u0001", string.Empty)
                .Replace("\uFFFC", string.Empty)
                // Word field begin/separator/end controls are structural, not
                // visible prose.  In particular U+0015 is what made genuine
                // MathType MTPlaceRef rows look non-empty and therefore inline.
                .Replace("\u0013", string.Empty)
                .Replace("\u0014", string.Empty)
                .Replace("\u0015", string.Empty)
                .Trim();
            if (text.Length == 0) return "block";

            // Some MathType/Word combinations materialize the display number as
            // visible numeric decoration instead of an empty MTPlaceRef result.
            // Keep supporting those forms without treating ordinary prose as a
            // display equation.
            return LooksLikeDisplayEquationDecoration(text) ? "block" : "inline";
        }
        catch { return "inline"; }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool ContainsMathTypeDisplayNumberField(Range paragraphRange)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            fields = paragraphRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var instruction = code.Text ?? string.Empty;
                if (instruction.IndexOf("MACROBUTTON MTPlaceRef", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static bool LooksLikeDisplayEquationDecoration(string text)
    {
        var sawNumber = false;
        foreach (var character in text)
        {
            if (char.IsWhiteSpace(character)) continue;
            if (char.IsDigit(character))
            {
                sawNumber = true;
                continue;
            }

            switch (character)
            {
                case '(':
                case ')':
                case '[':
                case ']':
                case '{':
                case '}':
                case '.':
                case ',':
                case ':':
                case ';':
                case '-':
                case '–':
                case '—':
                case '/':
                case '\\':
                    continue;
                default:
                    return false;
            }
        }
        return sawNumber;
    }

    private static string? ReadDefaultString(RegistryKey parent, string subKeyName)
    {
        using var subKey = parent.OpenSubKey(subKeyName);
        return subKey?.GetValue(null) as string;
    }

    private static void CloseOleObject(object? value)
    {
        if (value is not IOleObjectClose oleObject) return;
        try { oleObject.Close(OleCloseNoSave); }
        catch { }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct OleRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [ComImport]
    [Guid("00000112-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleObjectNative
    {
        [PreserveSig] int SetClientSite(IntPtr clientSite);
        [PreserveSig] int GetClientSite(out IntPtr clientSite);
        [PreserveSig] int SetHostNames(
            [MarshalAs(UnmanagedType.LPWStr)] string containerApp,
            [MarshalAs(UnmanagedType.LPWStr)] string containerObject);
        [PreserveSig] int Close(uint saveOption);
        [PreserveSig] int SetMoniker(uint whichMoniker, IntPtr moniker);
        [PreserveSig] int GetMoniker(uint assign, uint whichMoniker, out IntPtr moniker);
        [PreserveSig] int InitFromData(
            [MarshalAs(UnmanagedType.Interface)] object dataObject,
            [MarshalAs(UnmanagedType.Bool)] bool creation,
            uint reserved);
        [PreserveSig] int GetClipboardData(uint reserved, out IntPtr dataObject);
        [PreserveSig] int DoVerb(
            int verb,
            IntPtr message,
            IntPtr activeSite,
            int index,
            IntPtr parentWindow,
            ref OleRect positionRectangle);
    }

    [ComImport]
    [Guid("0000010A-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPersistStorageFull
    {
        [PreserveSig] int GetClassID(out Guid classId);
        [PreserveSig] int IsDirty();
        [PreserveSig] int InitNew(IntPtr storage);
        [PreserveSig] int Load(IntPtr storage);
        [PreserveSig] int Save(IntPtr storage, [MarshalAs(UnmanagedType.Bool)] bool sameAsLoad);
        [PreserveSig] int SaveCompleted(IntPtr storage);
        [PreserveSig] int HandsOffStorage();
    }

    [ComImport]
    [Guid("0000010A-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPersistStorageInit
    {
        [PreserveSig] int GetClassID(out Guid classId);
        [PreserveSig] int IsDirty();
        [PreserveSig] int InitNew(IntPtr storage);
    }

    [ComImport]
    [Guid("00000112-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleObjectClose
    {
        [PreserveSig] int SetClientSite(IntPtr clientSite);
        [PreserveSig] int GetClientSite(out IntPtr clientSite);
        [PreserveSig] int SetHostNames(
            [MarshalAs(UnmanagedType.LPWStr)] string containerApp,
            [MarshalAs(UnmanagedType.LPWStr)] string containerObject);
        [PreserveSig] int Close(uint saveOption);
    }

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string progId, out Guid clsid);

    [DllImport("ole32.dll")]
    private static extern int OleCreate(
        ref Guid classId,
        ref Guid interfaceId,
        uint renderOption,
        IntPtr format,
        IntPtr clientSite,
        IntPtr storage,
        out IntPtr oleObject);

    [DllImport("ole32.dll")]
    private static extern int CreateILockBytesOnHGlobal(
        IntPtr globalMemory,
        [MarshalAs(UnmanagedType.Bool)] bool deleteOnRelease,
        out IntPtr lockBytes);

    [DllImport("ole32.dll")]
    private static extern int StgCreateDocfileOnILockBytes(
        IntPtr lockBytes,
        uint mode,
        uint reserved,
        out IntPtr storage);

    [DllImport("ole32.dll")]
    private static extern int OleRun([MarshalAs(UnmanagedType.IUnknown)] object unknown);

    [DllImport("ole32.dll")]
    private static extern int OleGetAutoConvert(ref Guid clsidOld, out Guid clsidNew);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint RegisterClipboardFormat(string format);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int GetClipboardFormatName(
        uint format,
        StringBuilder buffer,
        int maxCount);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalAlloc(uint flags, UIntPtr bytes);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalLock(IntPtr memory);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GlobalUnlock(IntPtr memory);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern UIntPtr GlobalSize(IntPtr memory);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalFree(IntPtr memory);

    [DllImport("ole32.dll")]
    private static extern void ReleaseStgMedium(ref STGMEDIUM medium);
}
