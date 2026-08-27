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
    internal const string CanonicalProgId = "Equation.DSMT4";
    internal const string CanonicalUserType = "MathType 7.0 Equation";
    internal static readonly Guid CanonicalClsid =
        new("0002CE03-0000-0000-C000-000000000046");

    private static readonly object CapabilityCacheSync = new();
    private static readonly Dictionary<string, CapabilityCacheEntry> CapabilityCache =
        new(StringComparer.OrdinalIgnoreCase);
    private static readonly TimeSpan CapabilityCacheLifetime = TimeSpan.FromSeconds(30);
    private const int MaximumCapabilityCacheEntries = 32;

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
        public string? ClassName { get; set; }
        public string? ServerPath { get; set; }
        public int RunForConversionVerb { get; set; } = 2;
        public bool RegisteredMathMlGetSet { get; set; }
    }

    private sealed class CapabilityCacheEntry
    {
        public DateTimeOffset ExpiresAtUtc { get; set; }
        public bool Success { get; set; }
        public Capabilities Value { get; set; } = new();
    }

    internal sealed class StorageIdentity
    {
        public string ProgId { get; set; } = CanonicalProgId;
        public Guid Clsid { get; set; } = CanonicalClsid;
        public string UserType { get; set; } = CanonicalUserType;
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
                if (string.Equals(progId, CanonicalProgId, StringComparison.OrdinalIgnoreCase))
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
                // Never identify an object merely because its ProgID starts with
                // "Equation.". Microsoft Equation Editor and unrelated equation
                // servers use that prefix too. The Flat OPC reader validates the
                // embedded CFB as a MathType MTEF-v5 object, including the stored
                // user type/class identity, which also works when MathType is not
                // installed on the current computer.
                return MathTypeOleStorage.LooksLikeMathTypeCompoundFile(
                    fragment.CompoundFile);
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
        if (string.IsNullOrWhiteSpace(progId)) return false;

        // Do not assume every MathType release/registration uses an Equation.*
        // ProgID. Resolve any registered ProgID, then accept it only when the
        // resulting class name or executable is actually MathType. This remains
        // fail-closed for Microsoft Equation Editor and unrelated OLE servers.
        if (CLSIDFromProgID(progId!, out var current) != 0
            && !TryReadProgIdClsid(progId!, out current))
            return false;
        return TryResolveCapabilities(progId!, current, out capabilities);
    }

    internal static bool TryResolveCapabilities(
        InlineShape? shape,
        out Capabilities capabilities)
    {
        capabilities = new Capabilities();
        if (shape is null) return false;
        OLEFormat? format = null;
        try
        {
            if (shape.Type is not WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                and not WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                return false;

            try
            {
                format = shape.OLEFormat;
                if (TryResolveCapabilities(format.ProgID, out capabilities))
                    return true;
            }
            catch
            {
                // A damaged/missing ProgID registration can make OLEFormat fail.
                // Fall through to the serialized storage identity below.
            }
            finally
            {
                Release(format);
                format = null;
            }

            try
            {
                var fragment = MathTypeWordOpenXml.Read(shape);
                var identity = MathTypeOleStorage.ReadCompoundFileIdentity(
                    fragment.CompoundFile);
                return identity.Clsid != Guid.Empty
                    && TryResolveCapabilities(
                        CanonicalProgId,
                        identity.Clsid,
                        out capabilities);
            }
            catch
            {
                return false;
            }
        }
        catch
        {
            capabilities = new Capabilities();
            return false;
        }
        finally { Release(format); }
    }

    internal static bool TryResolveCapabilities(
        Guid clsid,
        out Capabilities capabilities)
    {
        capabilities = new Capabilities();
        return clsid != Guid.Empty
            && TryResolveCapabilities(CanonicalProgId, clsid, out capabilities);
    }

    internal static bool IsRegisteredMathTypeClass(Guid clsid) =>
        IsDirectlyRegisteredMathTypeClass(clsid);

    private static bool AllowsMathTypeAutoConvertAlias(string progId) =>
        string.Equals(progId, CanonicalProgId, StringComparison.OrdinalIgnoreCase)
        || string.Equals(progId, "DSEquations", StringComparison.OrdinalIgnoreCase)
        || string.Equals(progId, "Equation", StringComparison.OrdinalIgnoreCase)
        || progId.StartsWith("Equation.DSMT", StringComparison.OrdinalIgnoreCase);

    internal static bool IsDirectlyRegisteredMathTypeClass(Guid clsid)
    {
        if (clsid == Guid.Empty) return false;
        using var classKey = OpenClassesSubKey(
            "CLSID\\{" + clsid.ToString("D") + "}");
        if (classKey is null) return false;
        var className = classKey.GetValue(null) as string;
        var serverPath = NormalizeLocalServerPath(
            ReadDefaultString(classKey, "LocalServer32")
                ?? ReadDefaultString(classKey, "LocalServer"));
        return !string.IsNullOrWhiteSpace(className)
                && className!.IndexOf("MathType", StringComparison.OrdinalIgnoreCase) >= 0
            || string.Equals(
                Path.GetFileName(serverPath),
                "MathType.exe",
                StringComparison.OrdinalIgnoreCase);
    }

    internal static StorageIdentity ResolvePreferredStorageIdentity()
    {
        foreach (var progId in new[] { CanonicalProgId, "DSEquations", "Equation" })
        {
            if (!TryResolveCapabilities(progId, out var capabilities)) continue;
            return new StorageIdentity
            {
                ProgId = string.IsNullOrWhiteSpace(capabilities.ProgId)
                    ? CanonicalProgId
                    : capabilities.ProgId,
                Clsid = capabilities.ResolvedClsid,
                UserType = string.IsNullOrWhiteSpace(capabilities.ClassName)
                    ? CanonicalUserType
                    : capabilities.ClassName!,
            };
        }
        return new StorageIdentity();
    }

    private static bool TryResolveCapabilities(
        string progId,
        Guid initialClsid,
        out Capabilities capabilities)
    {
        capabilities = new Capabilities();
        var current = initialClsid;
        // Microsoft Equation Editor may register OleAutoConvert to MathType after
        // MathType is installed. That migration hint must not make Equation.3 (or
        // another unrelated equation server) look like an existing MathType OLE.
        // Follow auto-conversion only for MathType's own known aliases/family, or
        // when the starting class is already directly registered as MathType.
        if (!AllowsMathTypeAutoConvertAlias(progId)
            && !IsDirectlyRegisteredMathTypeClass(initialClsid))
            return false;
        var visited = new HashSet<Guid>();
        for (var depth = 0;
             depth < 16 && current != Guid.Empty && visited.Add(current);
             depth++)
        {
            var probe = current;
            var hr = OleGetAutoConvert(ref probe, out var next);
            if (hr != 0 || next == Guid.Empty || next == current)
                break;
            current = next;
        }

        var clsidKey = "CLSID\\{" + current.ToString("D") + "}";
        using var classKey = OpenClassesSubKey(clsidKey);
        if (classKey is null) return false;
        var className = classKey.GetValue(null) as string;
        var registeredProgId = ReadDefaultString(classKey, "ProgID");
        var serverPath = NormalizeLocalServerPath(
            ReadDefaultString(classKey, "LocalServer32")
                ?? ReadDefaultString(classKey, "LocalServer"));
        var serverFileName = string.IsNullOrWhiteSpace(serverPath)
            ? null
            : Path.GetFileName(serverPath);
        var isMathType = !string.IsNullOrWhiteSpace(className)
                && className!.IndexOf("MathType", StringComparison.OrdinalIgnoreCase) >= 0
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
            ProgId = string.IsNullOrWhiteSpace(registeredProgId)
                ? progId
                : registeredProgId!,
            ResolvedClsid = current,
            ClassName = className,
            ServerPath = serverPath,
            RunForConversionVerb = conversionVerb,
            RegisteredMathMlGetSet = supportsMathMl,
        };
        return true;
    }

    internal static FormulaMetadata ReadMetadata(
        Microsoft.Office.Interop.Word.Application application,
        InlineShape shape,
        string? knownMathMl = null)
    {
        string mathMl;
        try
        {
            // Batch format conversion captures the CFB once and passes the parsed
            // MathML here. Avoid reading Equation Native a second time for every
            // MathType object in a large document.
            mathMl = !string.IsNullOrWhiteSpace(knownMathMl)
                ? knownMathMl!
                : MathTypeOleStorage.ReadMathMl(shape);
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
        var numbered = false;
        var fontSizePt = FormulaFontSize.DefaultPt;
        try
        {
            range = shape.Range;
            displayMode = InferDisplayMode(range);
            numbered = ContainsMathTypeDisplayNumberFieldAtRange(range);
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
            Numbered = numbered,
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
            {
                // Some MathType 7.x builds expose MathML through IDataObject but
                // omit or vary the optional DataFormats\\GetSet registry entries.
                // Treat the registry as a discovery hint only; QueryGetData below
                // is the authoritative runtime capability probe.
                WordDoubleClickHook.TraceMessage(
                    $"mathtype-mathml-registry-hint-missing progId={format.ProgID} clsid={capabilities.ResolvedClsid}");
            }

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
            ConvertContainedMathTypeOleObject(oleObject, clientSite, progId);
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

    internal static byte[] CreateServerBackedCompoundFileFromText(
        string text,
        string progId = "Equation.DSMT4")
    {
        if (string.IsNullOrWhiteSpace(text))
            throw new InvalidDataException(
                "MathType server-backed storage creation requires formula text.");

        object? oleObject = null;
        IntPtr lockBytes = IntPtr.Zero;
        IntPtr storage = IntPtr.Zero;
        IntPtr globalMemory = IntPtr.Zero;
        IntPtr locked = IntPtr.Zero;
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

            ConvertContainedMathTypeOleObject(oleObject, clientSite, progId);
            var request = CreateFormatEtc(1); // CF_TEXT; MathType advertises this for DATADIR_SET.
            var textGlobal = Marshal.StringToHGlobalAnsi(text + "\0");
            if (textGlobal == IntPtr.Zero)
                throw new OutOfMemoryException(
                    "Unable to allocate MathType text transfer memory.");
            try
            {
                var medium = new STGMEDIUM
                {
                    tymed = TYMED.TYMED_HGLOBAL,
                    unionmember = textGlobal,
                    pUnkForRelease = null,
                };
                dataObject.SetData(ref request, ref medium, false);
            }
            finally { Marshal.FreeHGlobal(textGlobal); }
            SaveContainedMathTypeOleObject(oleObject, storage);

            var hr = GetHGlobalFromILockBytes(lockBytes, out globalMemory);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);
            if (globalMemory == IntPtr.Zero)
                throw new InvalidOperationException(
                    "MathType OleCreate storage did not expose its backing HGLOBAL.");
            var size = GlobalSize(globalMemory).ToUInt64();
            if (size == 0 || size > 64UL * 1024UL * 1024UL)
                throw new InvalidDataException(
                    $"MathType OleCreate produced an unexpected storage size of {size} bytes.");
            locked = GlobalLock(globalMemory);
            if (locked == IntPtr.Zero)
                throw new InvalidOperationException(
                    "Windows could not lock the MathType OleCreate storage HGLOBAL.");
            var bytes = new byte[checked((int)size)];
            Marshal.Copy(locked, bytes, 0, bytes.Length);
            if (!MathTypeOleStorage.LooksLikeMathTypeCompoundFile(bytes))
                throw new InvalidDataException(
                    "MathType OleCreate did not persist a valid Equation.DSMT4 Compound File.");
            _ = MathTypeOleStorage.ReadMathMl(bytes);
            return bytes;
        }
        finally
        {
            if (locked != IntPtr.Zero && globalMemory != IntPtr.Zero)
            {
                try { GlobalUnlock(globalMemory); } catch { }
            }
            CloseOleObject(oleObject);
            ReleaseRawComPointer(ref storage);
            ReleaseRawComPointer(ref lockBytes);
            Release(oleObject);
        }
    }

    internal static byte[] CreateServerBackedCompoundFile(
        string mathMl,
        string progId = "Equation.DSMT4")
    {
        if (string.IsNullOrWhiteSpace(mathMl)
            || !mathMl.TrimStart().StartsWith("<math", StringComparison.Ordinal))
            throw new InvalidDataException(
                "MathType server-backed storage creation requires valid MathML.");

        object? oleObject = null;
        IntPtr lockBytes = IntPtr.Zero;
        IntPtr storage = IntPtr.Zero;
        IntPtr globalMemory = IntPtr.Zero;
        IntPtr locked = IntPtr.Zero;
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

            // MathType's OLE contract requires the client item to enter its
            // RunForConversion verb before accepting programmatic MathML SetData.
            // A dormant OleCreate object otherwise returns DV_E_FORMATETC.
            ConvertContainedMathTypeOleObject(oleObject, clientSite, progId);
            WriteMathMlToDataObject(dataObject, mathMl);
            SaveContainedMathTypeOleObject(oleObject, storage);

            var hr = GetHGlobalFromILockBytes(lockBytes, out globalMemory);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);
            if (globalMemory == IntPtr.Zero)
                throw new InvalidOperationException(
                    "MathType OleCreate storage did not expose its backing HGLOBAL.");

            var size = GlobalSize(globalMemory).ToUInt64();
            if (size == 0 || size > 64UL * 1024UL * 1024UL)
                throw new InvalidDataException(
                    $"MathType OleCreate produced an unexpected storage size of {size} bytes.");
            locked = GlobalLock(globalMemory);
            if (locked == IntPtr.Zero)
                throw new InvalidOperationException(
                    "Windows could not lock the MathType OleCreate storage HGLOBAL.");

            var bytes = new byte[checked((int)size)];
            Marshal.Copy(locked, bytes, 0, bytes.Length);
            if (!MathTypeOleStorage.LooksLikeMathTypeCompoundFile(bytes))
                throw new InvalidDataException(
                    "MathType OleCreate did not persist a valid Equation.DSMT4 Compound File.");
            _ = MathTypeOleStorage.ReadMathMl(bytes);
            return bytes;
        }
        finally
        {
            if (locked != IntPtr.Zero && globalMemory != IntPtr.Zero)
            {
                try { GlobalUnlock(globalMemory); } catch { }
            }
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

            // MathType 7.x does not expose one stable DATADIR_SET contract across
            // point releases. Some builds advertise CF_TEXT but reject it with
            // DV_E_FORMATETC, while accepting the registered TeX input format that
            // is absent from the SET enumeration. Treat enumeration as a hint and
            // perform the actual SetData calls in semantic-preference order.
            var failures = new List<string>();
            var texFormat = RegisterClipboardFormat("TeX Input Language");
            COMException? texError = null;
            if (texFormat != 0 && texFormat <= ushort.MaxValue)
            {
                if (TrySetAnsiData(
                        dataObject,
                        CreateFormatEtc(unchecked((short)texFormat)),
                        text,
                        out texError))
                    return;
                if (texError is not null)
                    failures.Add($"TeX Input Language=0x{texError.HResult:X8}");
            }

            if (TrySetAnsiData(
                    dataObject,
                    CreateFormatEtc(1),
                    text,
                    out var textError))
                return;
            if (textError is not null)
                failures.Add($"CF_TEXT=0x{textError.HResult:X8}");

            throw new InvalidOperationException(
                "The installed MathType OLE server rejected all supported text SetData formats"
                + (failures.Count == 0 ? "." : ": " + string.Join(", ", failures)));
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

        var watch = Stopwatch.StartNew();
        var delayMilliseconds = 20;
        while (watch.Elapsed < TimeSpan.FromSeconds(5))
        {
            var running = TryGetObject(format);
            if (running is not null) return running;
            Thread.Sleep(delayMilliseconds);
            delayMilliseconds = Math.Min(120, delayMilliseconds + 15);
        }

        throw new COMException(
            "MathType OLE server started but Word did not expose IDataObject within 5 seconds."
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
            hr = OleSetContainedObject(oleObject, true);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);
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
        MathTypeOleClientSiteNative clientSite,
        string progId)
    {
        if (oleObject is not IOleObjectNative native)
            throw new InvalidOperationException("MathType OleCreate object does not expose IOleObject.");
        if (!TryResolveCapabilities(progId, out var capabilities))
            throw new InvalidOperationException(
                $"The installed OLE class '{progId}' is not recognized as MathType.");
        var rectangle = new OleRect();
        var clientSitePointer = Marshal.GetComInterfaceForObject(
            clientSite,
            typeof(IMathTypeOleClientSiteNative));
        try
        {
            var hr = native.DoVerb(
                capabilities.RunForConversionVerb,
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
        // MathType 7.x registry/EnumFormatEtc declarations vary by point release.
        // The actual SetData HRESULT is authoritative. Try each MathML spelling
        // first, then fall back to MathType's TeX/text import channels using the
        // same semantic equation rather than failing solely because one version
        // omitted a SET advertisement.
        var failures = new List<string>();
        var asciiMathMl = ToAsciiMathMlPayload(mathMl);
        foreach (var name in MathMlFormats)
        {
            var id = RegisterClipboardFormat(name);
            if (id == 0 || id > ushort.MaxValue) continue;
            if (TrySetAnsiData(
                    dataObject,
                    CreateFormatEtc(unchecked((short)id)),
                    asciiMathMl,
                    out var error))
                return;
            if (error is not null)
                failures.Add($"{name}=0x{error.HResult:X8}");
        }

        var latex = MathMlToLatexConverter.Convert(mathMl).Trim();
        if (!string.IsNullOrWhiteSpace(latex))
        {
            var texId = RegisterClipboardFormat("TeX Input Language");
            if (texId != 0 && texId <= ushort.MaxValue)
            {
                if (TrySetAnsiData(
                        dataObject,
                        CreateFormatEtc(unchecked((short)texId)),
                        latex,
                        out var texError))
                    return;
                if (texError is not null)
                    failures.Add($"TeX Input Language=0x{texError.HResult:X8}");
            }

            if (TrySetAnsiData(
                    dataObject,
                    CreateFormatEtc(1),
                    latex,
                    out var textError))
                return;
            if (textError is not null)
                failures.Add($"CF_TEXT=0x{textError.HResult:X8}");
        }

        throw new InvalidOperationException(
            "The installed MathType OLE server rejected all supported equation SetData formats"
            + (failures.Count == 0 ? "." : ": " + string.Join(", ", failures)));
    }

    private static bool TrySetAnsiData(
        System.Runtime.InteropServices.ComTypes.IDataObject dataObject,
        FORMATETC request,
        string payload,
        out COMException? error)
    {
        error = null;
        var global = Marshal.StringToHGlobalAnsi(payload + "\0");
        if (global == IntPtr.Zero)
            throw new OutOfMemoryException("Unable to allocate MathType transfer memory.");
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
                return true;
            }
            catch (COMException setError)
            {
                error = setError;
                return false;
            }
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
            return NormalizeMathMlRoot(root);
        }
        catch (System.Xml.XmlException error)
        {
            throw new InvalidDataException(
                "MathType returned malformed MathML XML.",
                error);
        }
    }

    private static string ToAsciiMathMlPayload(string mathMl)
    {
        var builder = new StringBuilder(mathMl.Length + 64);
        for (var index = 0; index < mathMl.Length; index++)
        {
            var value = mathMl[index];
            if (value <= 0x7F)
            {
                builder.Append(value);
                continue;
            }

            int codePoint;
            if (char.IsHighSurrogate(value)
                && index + 1 < mathMl.Length
                && char.IsLowSurrogate(mathMl[index + 1]))
            {
                codePoint = char.ConvertToUtf32(value, mathMl[++index]);
            }
            else
            {
                codePoint = value;
            }
            builder.Append("&#x")
                .Append(codePoint.ToString("X", System.Globalization.CultureInfo.InvariantCulture))
                .Append(';');
        }
        return builder.ToString();
    }

    private static string NormalizeMathMlRoot(XElement root)
    {
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";

        XNode CloneNode(XNode node)
        {
            if (node is XElement element)
            {
                return new XElement(
                    mathMl + element.Name.LocalName,
                    element.Attributes()
                        .Where(attribute => !attribute.IsNamespaceDeclaration)
                        .Select(attribute =>
                            attribute.Name.Namespace == XNamespace.Xml
                                ? new XAttribute(XNamespace.Xml + attribute.Name.LocalName, attribute.Value)
                                : new XAttribute(attribute.Name.LocalName, attribute.Value)),
                    element.Nodes().Select(CloneNode));
            }
            if (node is XText text) return new XText(text.Value);
            if (node is XCData cdata) return new XCData(cdata.Value);
            return new XText(string.Empty);
        }

        return ((XElement)CloneNode(root)).ToString(SaveOptions.DisableFormatting);
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
        InlineShapes? paragraphShapes = null;
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

            // An unnumbered MathType display equation is a single equation object
            // occupying its paragraph. A paragraph containing multiple inline
            // objects is an inline formula run even when there is no prose between
            // them. Otherwise zero-gap OLE+OLE is misclassified as display math
            // and the conversion safety check refuses the neighboring object.
            paragraphShapes = paragraphRange.InlineShapes;
            if (paragraphShapes.Count != 1) return "inline";

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
            Release(paragraphShapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static bool TryReadDisplayNumberPosition(
        InlineShape shape,
        out string numberPosition)
    {
        numberPosition = "right";
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            fields = paragraphRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(result);
                result = null;
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                result = field.Result;
                var fieldStart = Math.Max(paragraphRange.Start, code.Start - 1);
                var fieldEnd = Math.Min(paragraphRange.End, result.End + 1);
                if (fieldEnd <= shapeRange.Start)
                    numberPosition = "left";
                else if (fieldStart >= shapeRange.End)
                    numberPosition = "right";
                else
                    numberPosition = code.Start < shapeRange.Start ? "left" : "right";
                return true;
            }
            return false;
        }
        catch
        {
            numberPosition = "right";
            return false;
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static bool ContainsMathTypeDisplayNumberFieldAtRange(Range range)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = range.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return ContainsMathTypeDisplayNumberField(paragraphRange);
        }
        catch { return false; }
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

    internal static string? ResolveInstalledServerPath()
    {
        foreach (var progId in new[] { CanonicalProgId, "DSEquations", "Equation" })
        {
            if (TryResolveCapabilities(progId, out var capabilities)
                && !string.IsNullOrWhiteSpace(capabilities.ServerPath)
                && File.Exists(capabilities.ServerPath))
                return capabilities.ServerPath;
        }

        var appPath = ReadRegisteredApplicationPath("MathType.exe");
        if (!string.IsNullOrWhiteSpace(appPath) && File.Exists(appPath))
            return appPath;

        foreach (var root in new[]
                 {
                     Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                     Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                 })
        {
            if (string.IsNullOrWhiteSpace(root)) continue;
            foreach (var relative in new[]
                     {
                         Path.Combine("MathType", "MathType.exe"),
                         Path.Combine("WIRIS", "MathType", "MathType.exe"),
                     })
            {
                var candidate = Path.Combine(root, relative);
                if (File.Exists(candidate)) return candidate;
            }
        }
        return null;
    }

    internal static string? ResolveMathPagePath()
    {
        var overridePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_MATHTYPE_MATHPAGE_PATH");
        if (!string.IsNullOrWhiteSpace(overridePath))
        {
            var expanded = Environment.ExpandEnvironmentVariables(
                overridePath.Trim().Trim('"'));
            if (File.Exists(expanded)
                && IsMathPageBinaryCompatibleWithCurrentProcess(expanded))
                return expanded;
        }

        var architecture = Environment.Is64BitProcess ? "64" : "32";
        var candidates = new List<string>();
        var serverPath = ResolveInstalledServerPath();
        if (!string.IsNullOrWhiteSpace(serverPath))
        {
            var installRoot = Path.GetDirectoryName(serverPath);
            if (!string.IsNullOrWhiteSpace(installRoot))
                AddMathPageCandidates(candidates, installRoot!, architecture);
        }
        foreach (var root in new[]
                 {
                     Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                     Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                 })
        {
            if (!string.IsNullOrWhiteSpace(root))
            {
                AddMathPageCandidates(
                    candidates,
                    Path.Combine(root, "MathType"),
                    architecture);
                AddMathPageCandidates(
                    candidates,
                    Path.Combine(root, "WIRIS", "MathType"),
                    architecture);
            }
        }
        return candidates
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(path =>
                File.Exists(path)
                && IsMathPageBinaryCompatibleWithCurrentProcess(path));
    }

    private static bool TryReadProgIdClsid(string progId, out Guid clsid)
    {
        clsid = Guid.Empty;
        using var key = OpenClassesSubKey(progId + "\\CLSID");
        var value = key?.GetValue(null) as string;
        return Guid.TryParse(value, out clsid) && clsid != Guid.Empty;
    }

    private static RegistryKey? OpenClassesSubKey(string relativePath)
    {
        try
        {
            var merged = Registry.ClassesRoot.OpenSubKey(relativePath);
            if (merged is not null) return merged;
        }
        catch { }

        var views = Environment.Is64BitOperatingSystem
            ? new[] { RegistryView.Registry64, RegistryView.Registry32 }
            : new[] { RegistryView.Registry32 };
        foreach (var view in views)
        {
            foreach (var hive in new[]
                     {
                         RegistryHive.CurrentUser,
                         RegistryHive.LocalMachine,
                     })
            {
                try
                {
                    using var baseKey = RegistryKey.OpenBaseKey(hive, view);
                    var key = baseKey.OpenSubKey("Software\\Classes\\" + relativePath);
                    if (key is not null) return key;
                }
                catch
                {
                    // Continue across per-user/per-machine and 32/64-bit views.
                }
            }
        }
        return null;
    }

    private static string? ReadRegisteredApplicationPath(string executableName)
    {
        var views = Environment.Is64BitOperatingSystem
            ? new[] { RegistryView.Registry64, RegistryView.Registry32 }
            : new[] { RegistryView.Registry32 };
        foreach (var view in views)
        {
            foreach (var hive in new[]
                     {
                         RegistryHive.CurrentUser,
                         RegistryHive.LocalMachine,
                     })
            {
                try
                {
                    using var baseKey = RegistryKey.OpenBaseKey(hive, view);
                    using var key = baseKey.OpenSubKey(
                        "Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\"
                        + executableName);
                    var path = NormalizeLocalServerPath(key?.GetValue(null) as string);
                    if (!string.IsNullOrWhiteSpace(path)) return path;
                }
                catch { }
            }
        }
        return null;
    }

    private static void AddMathPageCandidates(
        ICollection<string> candidates,
        string installRoot,
        string architecture)
    {
        if (string.IsNullOrWhiteSpace(installRoot)) return;
        candidates.Add(Path.Combine(
            installRoot,
            "MathPage",
            architecture,
            "MathPage.wll"));
        candidates.Add(Path.Combine(installRoot, "MathPage", "MathPage.wll"));
        candidates.Add(Path.Combine(installRoot, "Office Support", "MathPage.wll"));
        candidates.Add(Path.Combine(installRoot, "MathPage.wll"));

        var mathPageRoot = Path.Combine(installRoot, "MathPage");
        try
        {
            if (!Directory.Exists(mathPageRoot)) return;
            foreach (var path in Directory.GetFiles(
                         mathPageRoot,
                         "MathPage.wll",
                         SearchOption.AllDirectories))
                candidates.Add(path);
        }
        catch
        {
            // Installation discovery is best-effort; explicit candidates remain.
        }
    }

    private static bool IsMathPageBinaryCompatibleWithCurrentProcess(string path)
    {
        try
        {
            using var stream = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            using var reader = new BinaryReader(stream);
            if (reader.ReadUInt16() != 0x5A4D) return false; // MZ
            stream.Position = 0x3C;
            var peOffset = reader.ReadInt32();
            if (peOffset < 0 || peOffset > stream.Length - 6) return false;
            stream.Position = peOffset;
            if (reader.ReadUInt32() != 0x00004550) return false; // PE\0\0
            var machine = reader.ReadUInt16();
            return RuntimeInformation.ProcessArchitecture switch
            {
                Architecture.X86 => machine == 0x014C,
                Architecture.X64 => machine == 0x8664,
                Architecture.Arm => machine == 0x01C4,
                Architecture.Arm64 => machine == 0xAA64,
                _ => false,
            };
        }
        catch
        {
            return false;
        }
    }

    private static string? NormalizeLocalServerPath(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var source = Environment.ExpandEnvironmentVariables(value!.Trim());
        string candidate;
        if (source.StartsWith("\"", StringComparison.Ordinal))
        {
            var closingQuote = source.IndexOf('"', 1);
            candidate = closingQuote > 1
                ? source.Substring(1, closingQuote - 1)
                : source.Trim('"');
        }
        else
        {
            var exeEnd = source.IndexOf(
                ".exe",
                StringComparison.OrdinalIgnoreCase);
            candidate = exeEnd >= 0
                ? source.Substring(0, exeEnd + 4)
                : source.Split(new[] { ' ', '\t' }, 2)[0];
        }
        return candidate.Trim().Trim('"');
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
    private static extern int GetHGlobalFromILockBytes(
        IntPtr lockBytes,
        out IntPtr globalMemory);

    [DllImport("ole32.dll")]
    private static extern int StgCreateDocfileOnILockBytes(
        IntPtr lockBytes,
        uint mode,
        uint reserved,
        out IntPtr storage);

    [DllImport("ole32.dll")]
    private static extern int OleRun([MarshalAs(UnmanagedType.IUnknown)] object unknown);

    [DllImport("ole32.dll")]
    private static extern int OleSetContainedObject(
        [MarshalAs(UnmanagedType.IUnknown)] object unknown,
        [MarshalAs(UnmanagedType.Bool)] bool contained);

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
