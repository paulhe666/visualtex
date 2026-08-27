using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

/// <summary>
/// Thin bridge over the same MathPage.WLL entry points used by MathType's own
/// WordCmds.dot MTSetData module. This deliberately does not synthesize an OLE
/// compound file or use the process-global clipboard: Word inserts MathType's
/// own BlankEqn.doc client item, RunForConversion activates that real server
/// object, and MTSetEqnFromLangStr writes MathML through MathType itself.
/// </summary>
internal static class MathTypeWordCommandsBridge
{
    private const short MtInitLaunchAsNeeded = 0;
    private const short MtLangMathMl = 2;
    private const int MtOk = 0;

    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    private delegate int MtInitApiDelegate(short options, short timeout);

    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    private delegate int MtTermApiDelegate();

    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    private delegate int MtSetEqnFromLangStrDelegate(
        [MarshalAs(UnmanagedType.IDispatch)] object oleObject,
        short langType,
        IntPtr langBuffer,
        int langLength);

    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    private delegate int MtGetLangStrFromEqnDelegate(
        [MarshalAs(UnmanagedType.IDispatch)] object oleObject,
        short langType,
        IntPtr langBuffer,
        ref int langLength);

    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    private delegate int MtCloseOleObjectDelegate(
        int saveOptions,
        [MarshalAs(UnmanagedType.IDispatch)] object oleObject);

    internal static string? ResolveBlankEquationPath()
    {
        var candidates = new List<string>();
        var serverPath = MathTypeOleInterop.ResolveInstalledServerPath();
        if (!string.IsNullOrWhiteSpace(serverPath))
        {
            var installRoot = Path.GetDirectoryName(serverPath);
            if (!string.IsNullOrWhiteSpace(installRoot))
                candidates.Add(Path.Combine(installRoot, "Office Support", "BlankEqn.doc"));
        }
        foreach (var root in new[]
                 {
                     Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                     Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                 })
        {
            if (!string.IsNullOrWhiteSpace(root))
                candidates.Add(Path.Combine(root, "MathType", "Office Support", "BlankEqn.doc"));
        }
        return candidates.FirstOrDefault(File.Exists);
    }

    internal static string ReadMathMl(InlineShape shape)
    {
        if (shape is null) throw new ArgumentNullException(nameof(shape));

        var mathPage = ResolveMathPagePath()
            ?? throw new FileNotFoundException("MathType MathPage.WLL was not found.");
        IntPtr module = IntPtr.Zero;
        MtTermApiDelegate? term = null;
        var initialized = false;
        OLEFormat? format = null;
        object? runningObject = null;
        try
        {
            module = LoadLibraryW(mathPage);
            if (module == IntPtr.Zero)
                throw new InvalidOperationException(
                    $"Windows could not load MathType Word API bridge '{mathPage}'.");

            var init = GetDelegate<MtInitApiDelegate>(module, "MTInitAPI")
                ?? throw new MissingMethodException("MathPage.WLL", "MTInitAPI");
            term = GetDelegate<MtTermApiDelegate>(module, "MTTermAPI")
                ?? throw new MissingMethodException("MathPage.WLL", "MTTermAPI");
            var getEquation = GetDelegate<MtGetLangStrFromEqnDelegate>(
                    module,
                    "MTGetLangStrFromEqn")
                ?? throw new MissingMethodException("MathPage.WLL", "MTGetLangStrFromEqn");
            var closeEquation = GetDelegate<MtCloseOleObjectDelegate>(
                    module,
                    "MTCloseOleObject")
                ?? throw new MissingMethodException("MathPage.WLL", "MTCloseOleObject");

            var initStatus = init(MtInitLaunchAsNeeded, 30);
            if (initStatus < 0)
                throw new InvalidOperationException(
                    $"MathType MTInitAPI failed with status {initStatus}.");
            initialized = true;

            format = shape.OLEFormat;
            if (!MathTypeOleInterop.TryResolveCapabilities(shape, out var capabilities))
                throw new InvalidDataException(
                    $"MathType server read could not resolve the OLE class '{format.ProgID}'.");

            object runForConversion = capabilities.RunForConversionVerb;
            format.DoVerb(ref runForConversion);

            runningObject = WaitForRunningMathTypeObject(
                format,
                TimeSpan.FromSeconds(5),
                out var objectError);
            if (runningObject is null)
                throw new InvalidOperationException(
                    "MathType RunForConversion did not expose OLEFormat.Object for readback within 5 seconds.",
                    objectError);

            var serverMathMl = ReadMathMlFromServer(getEquation, runningObject);
            var closeStatus = closeEquation(1, runningObject);
            if (closeStatus != MtOk)
                throw new InvalidOperationException(
                    $"MathType MTCloseOleObject failed with status {closeStatus} after readback.");
            return serverMathMl;
        }
        finally
        {
            Release(runningObject);
            Release(format);
            if (initialized && term is not null)
            {
                try { term(); } catch { }
            }
            if (module != IntPtr.Zero) FreeLibrary(module);
        }
    }

    internal static string SetMathMl(InlineShape shape, string mathMl)
    {
        if (shape is null) throw new ArgumentNullException(nameof(shape));
        if (string.IsNullOrWhiteSpace(mathMl)
            || mathMl.IndexOf("<math", StringComparison.OrdinalIgnoreCase) < 0)
            throw new InvalidDataException("MathType server write requires MathML.");

        var mathPage = ResolveMathPagePath()
            ?? throw new FileNotFoundException("MathType MathPage.WLL was not found.");
        IntPtr module = IntPtr.Zero;
        MtTermApiDelegate? term = null;
        var initialized = false;
        OLEFormat? format = null;
        object? runningObject = null;
        Range? shapeRange = null;
        Fields? fields = null;
        GCHandle mathMlHandle = default;
        try
        {
            module = LoadLibraryW(mathPage);
            if (module == IntPtr.Zero)
                throw new InvalidOperationException(
                    $"Windows could not load MathType Word API bridge '{mathPage}'.");

            var init = GetDelegate<MtInitApiDelegate>(module, "MTInitAPI")
                ?? throw new MissingMethodException("MathPage.WLL", "MTInitAPI");
            term = GetDelegate<MtTermApiDelegate>(module, "MTTermAPI")
                ?? throw new MissingMethodException("MathPage.WLL", "MTTermAPI");
            var setEquation = GetDelegate<MtSetEqnFromLangStrDelegate>(
                    module,
                    "MTSetEqnFromLangStr")
                ?? throw new MissingMethodException("MathPage.WLL", "MTSetEqnFromLangStr");
            var getEquation = GetDelegate<MtGetLangStrFromEqnDelegate>(
                    module,
                    "MTGetLangStrFromEqn")
                ?? throw new MissingMethodException("MathPage.WLL", "MTGetLangStrFromEqn");
            var closeEquation = GetDelegate<MtCloseOleObjectDelegate>(
                    module,
                    "MTCloseOleObject")
                ?? throw new MissingMethodException("MathPage.WLL", "MTCloseOleObject");

            var initStatus = init(MtInitLaunchAsNeeded, 30);
            if (initStatus < 0)
                throw new InvalidOperationException(
                    $"MathType MTInitAPI failed with status {initStatus}.");
            initialized = true;

            format = shape.OLEFormat;
            if (!MathTypeOleInterop.TryResolveCapabilities(shape, out var capabilities))
                throw new InvalidDataException(
                    $"MathType server write could not resolve the OLE class '{format.ProgID}'.");

            // Enter the installed MathType class's registered RunForConversion verb.
            // MathType 7 normally uses verb 2, but discovery is authoritative for
            // older/newer desktop builds and localized registrations.
            object runForConversion = capabilities.RunForConversionVerb;
            format.DoVerb(ref runForConversion);

            // Word/MathType may need a short COM pump before OLEFormat.Object is
            // exposed after RunForConversion. Do not activate any visible editor.
            runningObject = WaitForRunningMathTypeObject(
                format,
                TimeSpan.FromSeconds(5),
                out var objectError);
            if (runningObject is null)
                throw new InvalidOperationException(
                    "MathType RunForConversion did not expose OLEFormat.Object within 5 seconds.",
                    objectError);

            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                try
                {
                    var beforeSetMathMl = ReadMathMlFromServer(getEquation, runningObject);
                    Console.WriteLine(
                        "    [MathType server-before-set] signature='"
                        + MathTypeMtefCodec.SemanticSignature(beforeSetMathMl)
                        + "'; mathMl='" + beforeSetMathMl + "'");
                }
                catch (Exception beforeSetError)
                {
                    Console.WriteLine("    [MathType server-before-set] diagnostic failed: " + beforeSetError.Message);
                }
            }

            // Exact equivalent of MathType's MTSetData.SetMTData:
            //     Dim mmlUnicode() As Byte
            //     mmlUnicode = mmlStr
            //     MTSetEqnFromLangStr(myObj, mtlangMATHML, mmlUnicode(0), Len(mmlStr))
            var unicode = Encoding.Unicode.GetBytes(mathMl + "\0");
            mathMlHandle = GCHandle.Alloc(unicode, GCHandleType.Pinned);
            var setStatus = setEquation(
                runningObject,
                MtLangMathMl,
                mathMlHandle.AddrOfPinnedObject(),
                mathMl.Length);
            if (setStatus != MtOk)
                throw new InvalidOperationException(
                    $"MathType MTSetEqnFromLangStr failed with status {setStatus}.");

            // Read the equation back from MathType's own live server before it is
            // closed. This is the authoritative semantic check; it deliberately
            // bypasses VisualTeX's MTEF decoder so a decoder glyph-table bug cannot
            // be mistaken for a MathType write failure.
            var serverMathMl = ReadMathMlFromServer(getEquation, runningObject);

            // Exact equivalent of MathType's MTSetData.ShutdownMT for modern Word.
            shapeRange = shape.Range;
            fields = shapeRange.Fields;
            try { fields.Update(); } catch { }
            var closeStatus = closeEquation(1, runningObject);
            if (closeStatus != MtOk)
                throw new InvalidOperationException(
                    $"MathType MTCloseOleObject failed with status {closeStatus}.");
            return serverMathMl;
        }
        finally
        {
            if (mathMlHandle.IsAllocated) mathMlHandle.Free();
            Release(fields);
            Release(shapeRange);
            Release(runningObject);
            Release(format);
            if (initialized && term is not null)
            {
                try { term(); } catch { }
            }
            if (module != IntPtr.Zero) FreeLibrary(module);
        }
    }

    private static object? WaitForRunningMathTypeObject(
        OLEFormat format,
        TimeSpan timeout,
        out Exception? lastError)
    {
        lastError = null;
        var watch = System.Diagnostics.Stopwatch.StartNew();
        var delayMilliseconds = 15;
        do
        {
            try
            {
                var running = format.Object;
                if (running is not null) return running;
            }
            catch (Exception error) when (error is COMException or InvalidCastException)
            {
                lastError = error;
            }

            if (watch.Elapsed >= timeout) break;
            Thread.Sleep(delayMilliseconds);
            delayMilliseconds = Math.Min(120, delayMilliseconds + 15);
        }
        while (watch.Elapsed < timeout);
        return null;
    }

    private static string ReadMathMlFromServer(
        MtGetLangStrFromEqnDelegate getEquation,
        object runningObject)
    {
        var length = 0;
        var status = getEquation(runningObject, MtLangMathMl, IntPtr.Zero, ref length);
        if (status != MtOk || length <= 0 || length > 4 * 1024 * 1024)
            throw new InvalidOperationException(
                $"MathType MTGetLangStrFromEqn length query failed with status {status}, length {length}.");

        // MathType's installed VBA declaration for MTGetLangStrFromEqn uses
        // `ByVal langStr As String`, which VBA marshals as an ANSI mutable buffer.
        // This differs deliberately from MTSetEqnFromLangStr where MathType's own
        // MTSetData module passes a UTF-16LE Byte() pointer via `As Any`.
        var bufferBytes = checked(length + 2);
        var buffer = Marshal.AllocHGlobal(bufferBytes);
        try
        {
            for (var offset = 0; offset < bufferBytes; offset++)
                Marshal.WriteByte(buffer, offset, 0);
            var capacity = length;
            status = getEquation(runningObject, MtLangMathMl, buffer, ref capacity);
            if (status != MtOk)
                throw new InvalidOperationException(
                    $"MathType MTGetLangStrFromEqn failed with status {status}.");
            var safeLength = Math.Max(0, Math.Min(capacity, length));
            var result = Marshal.PtrToStringAnsi(buffer, safeLength) ?? string.Empty;
            var nullIndex = result.IndexOf('\0');
            if (nullIndex >= 0) result = result.Substring(0, nullIndex);
            var trimmed = result.Trim();
            XElement? math = null;
            try
            {
                var document = XDocument.Parse(trimmed, LoadOptions.PreserveWhitespace);
                math = document.Root?.DescendantsAndSelf()
                    .FirstOrDefault(element =>
                        string.Equals(element.Name.LocalName, "math", StringComparison.OrdinalIgnoreCase));
                if (math is null)
                    throw new InvalidDataException("MathType XML contains no MathML <math> element.");
            }
            catch (Exception error) when (error is System.Xml.XmlException or InvalidDataException)
            {
                throw new InvalidDataException(
                    $"MathType MTGetLangStrFromEqn did not return parseable MathML. length={length}; returned={capacity}; prefix='{trimmed.Substring(0, Math.Min(trimmed.Length, 120))}'.",
                    error);
            }
            return NormalizeMathMlElement(math);
        }
        finally { Marshal.FreeHGlobal(buffer); }
    }

    private static string NormalizeMathMlElement(XElement math)
    {
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";

        static XNode CloneNode(XNode node, XNamespace targetNamespace)
        {
            if (node is XElement element)
            {
                var clone = new XElement(
                    targetNamespace + element.Name.LocalName,
                    element.Attributes()
                        .Where(attribute => !attribute.IsNamespaceDeclaration)
                        .Select(attribute =>
                            attribute.Name.Namespace == XNamespace.Xml
                                ? new XAttribute(XNamespace.Xml + attribute.Name.LocalName, attribute.Value)
                                : new XAttribute(attribute.Name.LocalName, attribute.Value)),
                    element.Nodes().Select(child => CloneNode(child, targetNamespace)));
                return clone;
            }
            if (node is XText text) return new XText(text.Value);
            if (node is XCData cdata) return new XCData(cdata.Value);
            return new XText(string.Empty);
        }

        var normalized = (XElement)CloneNode(math, mathMl);
        return normalized.ToString(SaveOptions.DisableFormatting);
    }

    private static string? ResolveMathPagePath() =>
        MathTypeOleInterop.ResolveMathPagePath();

    private static T? GetDelegate<T>(IntPtr module, string name) where T : class
    {
        var address = GetProcAddress(module, name);
        if (address == IntPtr.Zero) return null;
        return Marshal.GetDelegateForFunctionPointer(address, typeof(T)) as T;
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.FinalReleaseComObject(value); } catch { }
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr LoadLibraryW(string path);

    [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
    private static extern IntPtr GetProcAddress(IntPtr module, string name);

    [DllImport("kernel32.dll")]
    private static extern bool FreeLibrary(IntPtr module);
}
