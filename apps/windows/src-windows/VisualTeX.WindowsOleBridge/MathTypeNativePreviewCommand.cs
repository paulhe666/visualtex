using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Win32;

namespace VisualTeX.WindowsOleBridge;

/// <summary>
/// Isolated MathType MathPage renderer used by the Word VSTO add-in. MathPage.wll
/// is legacy native code and can terminate a host process with AccessViolation for
/// malformed/unsupported input. Running it in this already-shipped sidecar keeps
/// WINWORD.EXE outside that failure boundary while still producing MathType's own
/// WMF presentation, dimensions and baseline.
/// </summary>
internal static class MathTypeNativePreviewCommand
{
    private const short MtInitLaunchAsNeeded = 0;
    private const short MtXfmLocal = -3;
    private const short MtXfmFile = -4;
    private const short MtXfmMtef = 4;
    private const short MtXfmPict = 6;
    private const int MtOk = 0;
    private const int MtError = -9999;

    [StructLayout(LayoutKind.Sequential, Pack = 2)]
    private struct RectNative
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 2)]
    private struct DimsNative
    {
        public short Baseline;
        public RectNative Bounds;
    }

    [UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet = CharSet.Ansi)]
    private delegate int MtInitApiDelegate(short options, short timeout);

    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    private delegate int MtTermApiDelegate();

    [UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet = CharSet.Ansi)]
    private delegate int MtXFormEqnDelegate(
        short src,
        short srcFormat,
        byte[] srcData,
        int srcLength,
        short dst,
        short dstFormat,
        IntPtr dstData,
        int dstLength,
        [MarshalAs(UnmanagedType.LPStr)] string dstPath,
        ref DimsNative dims);

    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    private delegate int MtGetLastDimensionDelegate(short index);

    private sealed class Request
    {
        public string ResultPath { get; set; } = string.Empty;
        public string MathTypeServerPath { get; set; } = string.Empty;
        public List<RequestItem> Items { get; set; } = new();
    }

    private sealed class RequestItem
    {
        public string Id { get; set; } = string.Empty;
        public string MtefPath { get; set; } = string.Empty;
        public string WmfPath { get; set; } = string.Empty;
    }

    private sealed class Response
    {
        public string Error { get; set; } = string.Empty;
        public List<ResponseItem> Items { get; set; } = new();
    }

    private sealed class ResponseItem
    {
        public string Id { get; set; } = string.Empty;
        public bool Success { get; set; }
        public string Error { get; set; } = string.Empty;
        public string WmfPath { get; set; } = string.Empty;
        public float WidthPt { get; set; }
        public float HeightPt { get; set; }
        public int WordPosition { get; set; }
    }

    internal static int Run(string manifestPath)
    {
        if (string.IsNullOrWhiteSpace(manifestPath) || !File.Exists(manifestPath))
            return 2;

        Request? request = null;
        try
        {
            request = JsonSerializer.Deserialize<Request>(File.ReadAllText(manifestPath));
            if (request is null
                || string.IsNullOrWhiteSpace(request.ResultPath)
                || request.Items.Count == 0)
                return 2;
        }
        catch
        {
            return 2;
        }

        var response = new Response();
        IntPtr module = IntPtr.Zero;
        MtTermApiDelegate? term = null;
        var initialized = false;
        var privateDllDirectoryApplied = false;
        PrivateRuntimeScope? privateRuntimeScope = null;
        IntPtr overriddenLocalMachine = IntPtr.Zero;
        var localMachineOverridden = false;
        try
        {
            var mathPage = ResolveMathPagePath(request.MathTypeServerPath);
            if (mathPage is null)
            {
                response.Error = "MathType MathPage.wll is not installed.";
                return WriteResponse(request.ResultPath, response, 3);
            }

            if (IsPrivateMathPagePath(mathPage))
            {
                var runtimeDirectory = Path.GetDirectoryName(mathPage);
                if (!string.IsNullOrWhiteSpace(runtimeDirectory))
                    privateDllDirectoryApplied = SetDllDirectoryW(runtimeDirectory);
            }

            module = LoadLibraryW(mathPage);
            if (module == IntPtr.Zero)
            {
                response.Error = "MathType MathPage.wll could not be loaded.";
                return WriteResponse(request.ResultPath, response, 3);
            }

            var virtualLocalMachineRoot = Environment.GetEnvironmentVariable(
                "VISUALTEX_MATHTYPE_VIRTUAL_HKLM_ROOT");
            if (string.IsNullOrWhiteSpace(virtualLocalMachineRoot)
                && IsPrivateMathPagePath(mathPage))
            {
                if (!PrivateRuntimeScope.TryCreate(
                        mathPage,
                        out privateRuntimeScope,
                        out var privateRuntimeError))
                {
                    response.Error = privateRuntimeError;
                    return WriteResponse(request.ResultPath, response, 3);
                }
            }
            else if (!string.IsNullOrWhiteSpace(virtualLocalMachineRoot))
            {
                var openStatus = RegOpenKeyExW(
                    HkeyCurrentUser,
                    virtualLocalMachineRoot.Trim().Trim('\\'),
                    0,
                    KeyRead,
                    out overriddenLocalMachine);
                if (openStatus != 0 || overriddenLocalMachine == IntPtr.Zero)
                {
                    response.Error = $"Private MathType virtual HKLM root could not be opened; status {openStatus}.";
                    return WriteResponse(request.ResultPath, response, 3);
                }
                var overrideStatus = RegOverridePredefKey(HkeyLocalMachine, overriddenLocalMachine);
                if (overrideStatus != 0)
                {
                    response.Error = $"Private MathType virtual HKLM override failed; status {overrideStatus}.";
                    return WriteResponse(request.ResultPath, response, 3);
                }
                localMachineOverridden = true;
            }

            var init = GetDelegate<MtInitApiDelegate>(module, "MTInitAPI");
            term = GetDelegate<MtTermApiDelegate>(module, "MTTermAPI");
            var transform = GetDelegate<MtXFormEqnDelegate>(module, "MTXFormEqn");
            var getDimension = GetDelegate<MtGetLastDimensionDelegate>(module, "MTGetLastDimension");
            if (init is null || term is null || transform is null || getDimension is null)
            {
                response.Error = "MathType MathPage API exports are incomplete.";
                return WriteResponse(request.ResultPath, response, 3);
            }

            var initStatus = init(MtInitLaunchAsNeeded, 8);
            initialized = initStatus >= 0;
            if (!initialized)
            {
                response.Error = $"MathType MathPage API initialization failed with status {initStatus}.";
                return WriteResponse(request.ResultPath, response, 3);
            }

            foreach (var item in request.Items)
            {
                var result = new ResponseItem
                {
                    Id = item.Id,
                    WmfPath = item.WmfPath,
                };
                response.Items.Add(result);
                try
                {
                    if (string.IsNullOrWhiteSpace(item.Id)
                        || string.IsNullOrWhiteSpace(item.MtefPath)
                        || !File.Exists(item.MtefPath)
                        || string.IsNullOrWhiteSpace(item.WmfPath))
                    {
                        result.Error = "Invalid native-preview request item.";
                        continue;
                    }
                    var mtef = File.ReadAllBytes(item.MtefPath);
                    if (mtef.Length == 0)
                    {
                        result.Error = "MTEF input is empty.";
                        continue;
                    }
                    var outputDirectory = Path.GetDirectoryName(item.WmfPath);
                    if (!string.IsNullOrWhiteSpace(outputDirectory))
                        Directory.CreateDirectory(outputDirectory);
                    try { File.Delete(item.WmfPath); } catch { }

                    var dims = new DimsNative();
                    var status = transform(
                        MtXfmLocal,
                        MtXfmMtef,
                        mtef,
                        mtef.Length,
                        MtXfmFile,
                        MtXfmPict,
                        IntPtr.Zero,
                        0,
                        item.WmfPath,
                        ref dims);
                    if (status != MtOk
                        || !File.Exists(item.WmfPath)
                        || new FileInfo(item.WmfPath).Length <= 22)
                    {
                        result.Error = $"MathType native renderer returned status {status}.";
                        continue;
                    }

                    var width32 = getDimension(1);
                    var height32 = getDimension(2);
                    var baseline32 = getDimension(3);
                    if (width32 <= 0 || height32 <= 0 || baseline32 < 0
                        || width32 == MtError || height32 == MtError || baseline32 == MtError)
                    {
                        result.Error = "MathType native renderer returned invalid dimensions.";
                        continue;
                    }

                    result.WidthPt = width32 / 32f;
                    result.HeightPt = height32 / 32f;
                    result.WordPosition = -(int)Math.Round(
                        baseline32 / 32d,
                        MidpointRounding.AwayFromZero);
                    result.Success = true;
                }
                catch (Exception error)
                {
                    result.Error = error.Message;
                }

                // MathPage is legacy native code and can terminate this sidecar
                // with an AccessViolation on a later transform. Persist progress
                // after every item so the parent can keep already-rendered WMFs
                // and retry only the unfinished tail in a fresh process.
                TryCheckpointResponse(request.ResultPath, response);
            }
            return WriteResponse(request.ResultPath, response, 0);
        }
        catch (Exception error)
        {
            response.Error = error.Message;
            return WriteResponse(request.ResultPath, response, 4);
        }
        finally
        {
            if (initialized && term is not null)
            {
                try { term(); } catch { }
            }
            if (module != IntPtr.Zero) FreeLibrary(module);
            if (localMachineOverridden)
            {
                try { RegOverridePredefKey(HkeyLocalMachine, IntPtr.Zero); } catch { }
            }
            if (overriddenLocalMachine != IntPtr.Zero) RegCloseKey(overriddenLocalMachine);
            privateRuntimeScope?.Dispose();
            if (privateDllDirectoryApplied) SetDllDirectoryW(null);
        }
    }

    private const string PrivateMathTypeClsid =
        "{0002CE03-0000-0000-C000-000000000046}";
    private const string PrivateRuntimeMutexName =
        @"Local\VisualTeX.PrivateMathTypeRuntimeBootstrap.v1";
    private const string PrivateRuntimeRootEnvironment =
        "VISUALTEX_MATHTYPE_RUNTIME_ROOT";

    private sealed class PrivateLocalServerSnapshot
    {
        private readonly bool _parentExisted;
        private readonly bool _keyExisted;
        private readonly bool _valueExisted;
        private readonly object? _value;
        private readonly RegistryValueKind _valueKind;
        private readonly string _temporaryValue;

        private static string ParentPath =>
            $@"Software\Classes\CLSID\{PrivateMathTypeClsid}";
        private static string KeyPath => ParentPath + @"\LocalServer32";

        private PrivateLocalServerSnapshot(
            bool parentExisted,
            bool keyExisted,
            bool valueExisted,
            object? value,
            RegistryValueKind valueKind,
            string temporaryValue)
        {
            _parentExisted = parentExisted;
            _keyExisted = keyExisted;
            _valueExisted = valueExisted;
            _value = value;
            _valueKind = valueKind;
            _temporaryValue = temporaryValue;
        }

        internal static PrivateLocalServerSnapshot CaptureAndInstall(
            string serverPath)
        {
            using var baseKey = RegistryKey.OpenBaseKey(
                RegistryHive.CurrentUser,
                RegistryView.Registry32);
            using var parent = baseKey.OpenSubKey(ParentPath, writable: false);
            var parentExisted = parent is not null;
            using var existing = baseKey.OpenSubKey(KeyPath, writable: false);
            var keyExisted = existing is not null;
            var valueExisted = existing?.GetValueNames().Any(name => name.Length == 0)
                == true;
            var value = valueExisted
                ? existing!.GetValue(
                    string.Empty,
                    null,
                    RegistryValueOptions.DoNotExpandEnvironmentNames)
                : null;
            var valueKind = valueExisted
                ? existing!.GetValueKind(string.Empty)
                : RegistryValueKind.String;

            using var key = baseKey.CreateSubKey(KeyPath, writable: true)
                ?? throw new InvalidOperationException(
                    "VisualTeX could not create the temporary MathType LocalServer32 registration.");
            key.SetValue(string.Empty, serverPath, RegistryValueKind.String);
            return new PrivateLocalServerSnapshot(
                parentExisted,
                keyExisted,
                valueExisted,
                value,
                valueKind,
                serverPath);
        }

        internal void Restore()
        {
            try
            {
                using var baseKey = RegistryKey.OpenBaseKey(
                    RegistryHive.CurrentUser,
                    RegistryView.Registry32);
                using (var key = baseKey.OpenSubKey(KeyPath, writable: true))
                {
                    if (key is null) return;
                    var current = key.GetValue(
                        string.Empty,
                        null,
                        RegistryValueOptions.DoNotExpandEnvironmentNames) as string;
                    if (!string.Equals(
                            current,
                            _temporaryValue,
                            StringComparison.OrdinalIgnoreCase))
                        return;

                    if (_valueExisted)
                        key.SetValue(string.Empty, _value!, _valueKind);
                    else
                        key.DeleteValue(string.Empty, throwOnMissingValue: false);
                }

                if (!_keyExisted)
                {
                    using var key = baseKey.OpenSubKey(KeyPath, writable: false);
                    if (key is not null
                        && key.GetSubKeyNames().Length == 0
                        && key.GetValueNames().Length == 0)
                    {
                        key.Dispose();
                        try { baseKey.DeleteSubKey(KeyPath, throwOnMissingSubKey: false); }
                        catch { }
                    }
                }

                if (!_parentExisted)
                {
                    using var parent = baseKey.OpenSubKey(ParentPath, writable: false);
                    if (parent is not null
                        && parent.GetSubKeyNames().Length == 0
                        && parent.GetValueNames().Length == 0)
                    {
                        parent.Dispose();
                        try { baseKey.DeleteSubKey(ParentPath, throwOnMissingSubKey: false); }
                        catch { }
                    }
                }
            }
            catch
            {
                // Restoration is fail-closed. Never overwrite a value that changed
                // while the short bootstrap window was active.
            }
        }
    }

    private sealed class PrivateRuntimeScope : IDisposable
    {
        private readonly Mutex _mutex;
        private readonly string _virtualRootPath;
        private readonly PrivateLocalServerSnapshot _localServer;
        private IntPtr _virtualRootHandle;
        private bool _ownsMutex;
        private bool _localMachineOverridden;
        private bool _disposed;

        private PrivateRuntimeScope(
            Mutex mutex,
            bool ownsMutex,
            string virtualRootPath,
            PrivateLocalServerSnapshot localServer,
            IntPtr virtualRootHandle,
            bool localMachineOverridden)
        {
            _mutex = mutex;
            _ownsMutex = ownsMutex;
            _virtualRootPath = virtualRootPath;
            _localServer = localServer;
            _virtualRootHandle = virtualRootHandle;
            _localMachineOverridden = localMachineOverridden;
        }

        internal static bool TryCreate(
            string mathPagePath,
            out PrivateRuntimeScope? scope,
            out string error)
        {
            scope = null;
            error = string.Empty;
            if (!TryResolvePrivateRuntimeRoot(
                    mathPagePath,
                    out var runtimeRoot,
                    out error))
                return false;

            Mutex? mutex = null;
            PrivateLocalServerSnapshot? localServer = null;
            string? virtualRootPath = null;
            IntPtr virtualRootHandle = IntPtr.Zero;
            var localMachineOverridden = false;
            var ownsMutex = false;
            try
            {
                mutex = new Mutex(initiallyOwned: false, PrivateRuntimeMutexName);
                try
                {
                    ownsMutex = mutex.WaitOne(TimeSpan.FromSeconds(10));
                }
                catch (AbandonedMutexException)
                {
                    ownsMutex = true;
                }
                if (!ownsMutex)
                {
                    error = "VisualTeX private MathType runtime bootstrap is busy.";
                    mutex.Dispose();
                    return false;
                }

                var serverPath = Path.Combine(runtimeRoot, "MathType.exe");
                localServer = PrivateLocalServerSnapshot.CaptureAndInstall(serverPath);
                virtualRootPath = CreatePrivateVirtualHklm(runtimeRoot);
                var openStatus = RegOpenKeyExW(
                    HkeyCurrentUser,
                    virtualRootPath,
                    0,
                    KeyRead,
                    out virtualRootHandle);
                if (openStatus != 0 || virtualRootHandle == IntPtr.Zero)
                    throw new InvalidOperationException(
                        $"VisualTeX private MathType virtual HKLM root could not be opened; status {openStatus}.");

                var overrideStatus = RegOverridePredefKey(
                    HkeyLocalMachine,
                    virtualRootHandle);
                if (overrideStatus != 0)
                    throw new InvalidOperationException(
                        $"VisualTeX private MathType virtual HKLM override failed; status {overrideStatus}.");
                localMachineOverridden = true;

                scope = new PrivateRuntimeScope(
                    mutex,
                    ownsMutex,
                    virtualRootPath,
                    localServer,
                    virtualRootHandle,
                    localMachineOverridden);
                return true;
            }
            catch (Exception exception)
            {
                error = exception.Message;
                if (localMachineOverridden)
                {
                    try { RegOverridePredefKey(HkeyLocalMachine, IntPtr.Zero); }
                    catch { }
                }
                if (virtualRootHandle != IntPtr.Zero)
                    RegCloseKey(virtualRootHandle);
                if (!string.IsNullOrWhiteSpace(virtualRootPath))
                    DeletePrivateVirtualHklm(virtualRootPath!);
                localServer?.Restore();
                if (ownsMutex && mutex is not null)
                {
                    try { mutex.ReleaseMutex(); } catch { }
                }
                mutex?.Dispose();
                return false;
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            if (_localMachineOverridden)
            {
                try { RegOverridePredefKey(HkeyLocalMachine, IntPtr.Zero); }
                catch { }
                _localMachineOverridden = false;
            }
            if (_virtualRootHandle != IntPtr.Zero)
            {
                RegCloseKey(_virtualRootHandle);
                _virtualRootHandle = IntPtr.Zero;
            }
            DeletePrivateVirtualHklm(_virtualRootPath);
            _localServer.Restore();
            if (_ownsMutex)
            {
                try { _mutex.ReleaseMutex(); } catch { }
                _ownsMutex = false;
            }
            _mutex.Dispose();
        }
    }

    private static bool TryResolvePrivateRuntimeRoot(
        string mathPagePath,
        out string runtimeRoot,
        out string error)
    {
        runtimeRoot = string.Empty;
        error = string.Empty;
        var candidates = new List<string>();
        var overrideRoot = Environment.GetEnvironmentVariable(
            PrivateRuntimeRootEnvironment);
        if (!string.IsNullOrWhiteSpace(overrideRoot))
        {
            try
            {
                candidates.Add(Path.GetFullPath(
                    Environment.ExpandEnvironmentVariables(
                        overrideRoot.Trim().Trim('"'))));
            }
            catch { }
        }

        try
        {
            var directory = new DirectoryInfo(
                Path.GetDirectoryName(Path.GetFullPath(mathPagePath))!);
            if ((string.Equals(directory.Name, "64", StringComparison.OrdinalIgnoreCase)
                 || string.Equals(directory.Name, "32", StringComparison.OrdinalIgnoreCase))
                && directory.Parent is not null
                && string.Equals(
                    directory.Parent.Name,
                    "MathPage",
                    StringComparison.OrdinalIgnoreCase)
                && directory.Parent.Parent is not null)
            {
                candidates.Add(directory.Parent.Parent.FullName);
            }
            else if (string.Equals(
                         directory.Name,
                         "mathtype-runtime",
                         StringComparison.OrdinalIgnoreCase))
            {
                candidates.Add(directory.FullName);
            }
        }
        catch { }

        var bundledRuntimeRoot = Path.Combine(
            AppContext.BaseDirectory,
            "mathtype-runtime");
        if (IsPathWithinRoot(mathPagePath, bundledRuntimeRoot))
            candidates.Insert(0, bundledRuntimeRoot);

        foreach (var candidate in candidates
                     .Where(value => !string.IsNullOrWhiteSpace(value))
                     .Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (ValidatePrivateRuntime(candidate, mathPagePath, out var validationError))
            {
                runtimeRoot = Path.GetFullPath(candidate);
                return true;
            }
            error = validationError;
        }
        if (string.IsNullOrWhiteSpace(error))
            error = "VisualTeX private MathType runtime root could not be resolved.";
        return false;
    }

    private static bool ValidatePrivateRuntime(
        string runtimeRoot,
        string mathPagePath,
        out string error)
    {
        error = string.Empty;
        string root;
        try { root = Path.GetFullPath(runtimeRoot); }
        catch
        {
            error = "VisualTeX private MathType runtime path is invalid.";
            return false;
        }

        var requiredFiles = new[]
        {
            "MathType.exe",
            "MT7.dsc",
            Path.Combine("System", "MathTypeLib.exe"),
            Path.Combine("System", "MT6.dll"),
            Path.Combine("System", "64", "MT6.dll"),
            Path.Combine("System", "rt", "lib", "rt.jar"),
            Path.Combine("System", "rt", "lib", "charsets.jar"),
            Path.Combine("System", "rt", "lib", "jce.jar"),
            Path.Combine("System", "rt", "lib", "jsse.jar"),
            Path.Combine("System", "rt", "lib", "security", "java.security"),
        };
        foreach (var relative in requiredFiles)
        {
            var path = Path.Combine(root, relative);
            if (!File.Exists(path))
            {
                error = $"VisualTeX private MathType runtime is incomplete: missing {relative}.";
                return false;
            }
        }
        foreach (var relative in new[]
                 {
                     Path.Combine("System", "rt", "bin"),
                     Path.Combine("System", "rt", "lib", "ext"),
                     Path.Combine("System", "rt", "lib", "security"),
                 })
        {
            if (!Directory.Exists(Path.Combine(root, relative)))
            {
                error = $"VisualTeX private MathType runtime is incomplete: missing {relative}.";
                return false;
            }
        }
        if (!IsPathWithinRoot(mathPagePath, root))
        {
            error = "MathPage.wll is outside the resolved VisualTeX private MathType runtime.";
            return false;
        }
        return true;
    }

    private static string CreatePrivateVirtualHklm(string runtimeRoot)
    {
        var rootPath = $@"Software\VisualTeX\PrivateMathTypeRuntime\VirtualHKLM\{Environment.ProcessId}-{Guid.NewGuid():N}";
        using var baseKey = RegistryKey.OpenBaseKey(
            RegistryHive.CurrentUser,
            RegistryView.Registry64);
        using (var directories = baseKey.CreateSubKey(
                   rootPath + @"\SOFTWARE\Design Science\DSMT7\Directories",
                   writable: true)
               ?? throw new InvalidOperationException(
                   "VisualTeX could not create the private MathType directory registry view."))
        {
            directories.SetValue("ProgDir", runtimeRoot, RegistryValueKind.String);
            directories.SetValue(
                "AppSystemDir",
                Path.Combine(runtimeRoot, "System"),
                RegistryValueKind.String);
            directories.SetValue(
                "AppSystemDir32",
                Path.Combine(runtimeRoot, "System", "32"),
                RegistryValueKind.String);
            directories.SetValue(
                "AppSystemDir64",
                Path.Combine(runtimeRoot, "System", "64"),
                RegistryValueKind.String);
            directories.SetValue(
                "LangDir",
                Path.Combine(runtimeRoot, "Language"),
                RegistryValueKind.String);
            directories.SetValue(
                "TranslatorDir",
                Path.Combine(runtimeRoot, "Translators"),
                RegistryValueKind.String);
            directories.SetValue(
                "PrefsDir",
                Path.Combine(runtimeRoot, "Preferences"),
                RegistryValueKind.String);
            directories.SetValue(
                "MathPageDir",
                Path.Combine(runtimeRoot, "MathPage"),
                RegistryValueKind.String);
        }

        using (var currentVersion = baseKey.CreateSubKey(
                   rootPath + @"\SOFTWARE\Microsoft\Windows\CurrentVersion",
                   writable: true)
               ?? throw new InvalidOperationException(
                   "VisualTeX could not create the private MathType Windows directory registry view."))
        {
            SetRegistryStringIfPresent(
                currentVersion,
                "ProgramFilesDir",
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles));
            SetRegistryStringIfPresent(
                currentVersion,
                "CommonFilesDir",
                Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles));
            SetRegistryStringIfPresent(
                currentVersion,
                "ProgramFilesDir (x86)",
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86));
            SetRegistryStringIfPresent(
                currentVersion,
                "CommonFilesDir (x86)",
                Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86));
        }
        return rootPath;
    }

    private static void SetRegistryStringIfPresent(
        RegistryKey key,
        string name,
        string value)
    {
        if (!string.IsNullOrWhiteSpace(value))
            key.SetValue(name, value, RegistryValueKind.String);
    }

    private static void DeletePrivateVirtualHklm(string rootPath)
    {
        if (string.IsNullOrWhiteSpace(rootPath)) return;
        try
        {
            using var baseKey = RegistryKey.OpenBaseKey(
                RegistryHive.CurrentUser,
                RegistryView.Registry64);
            baseKey.DeleteSubKeyTree(rootPath, throwOnMissingSubKey: false);
            DeleteRegistryKeyIfEmpty(
                baseKey,
                @"Software\VisualTeX\PrivateMathTypeRuntime\VirtualHKLM");
            DeleteRegistryKeyIfEmpty(
                baseKey,
                @"Software\VisualTeX\PrivateMathTypeRuntime");
        }
        catch { }
    }

    private static void DeleteRegistryKeyIfEmpty(
        RegistryKey baseKey,
        string path)
    {
        using var key = baseKey.OpenSubKey(path, writable: false);
        if (key is null
            || key.GetSubKeyNames().Length != 0
            || key.GetValueNames().Length != 0)
            return;
        key.Dispose();
        try { baseKey.DeleteSubKey(path, throwOnMissingSubKey: false); }
        catch { }
    }

    private static bool IsPathWithinRoot(string path, string root)
    {
        try
        {
            var fullPath = Path.GetFullPath(path);
            var fullRoot = Path.GetFullPath(root)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            return fullPath.StartsWith(fullRoot, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static void TryCheckpointResponse(string path, Response response)
    {
        try
        {
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
            File.WriteAllText(path, JsonSerializer.Serialize(response));
        }
        catch { }
    }

    private static int WriteResponse(string path, Response response, int exitCode)
    {
        try
        {
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
            File.WriteAllText(path, JsonSerializer.Serialize(response));
        }
        catch
        {
            return exitCode == 0 ? 5 : exitCode;
        }
        return exitCode;
    }

    private static bool IsPrivateMathPagePath(string mathPagePath)
    {
        var overridePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_MATHTYPE_MATHPAGE_PATH");
        if (!string.IsNullOrWhiteSpace(overridePath))
        {
            var expanded = Environment.ExpandEnvironmentVariables(
                overridePath.Trim().Trim('"'));
            if (string.Equals(
                    Path.GetFullPath(expanded),
                    Path.GetFullPath(mathPagePath),
                    StringComparison.OrdinalIgnoreCase))
                return true;
        }

        var bundledRuntimeRoot = Path.Combine(
            AppContext.BaseDirectory,
            "mathtype-runtime");
        return IsPathWithinRoot(mathPagePath, bundledRuntimeRoot);
    }

    private static string? ResolveMathPagePath(string? mathTypeServerPath)
    {
        var overridePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_MATHTYPE_MATHPAGE_PATH");
        if (!string.IsNullOrWhiteSpace(overridePath))
        {
            var expanded = Environment.ExpandEnvironmentVariables(
                overridePath.Trim().Trim('"'));
            if (File.Exists(expanded)
                && IsNativeLibraryCompatibleWithCurrentProcess(expanded))
                return expanded;
        }

        // The preview command runs in the packaged x64 sidecar, independently of
        // Word's bitness. Resolve MathPage for this process rather than assuming the
        // Office/VSTO architecture. This keeps x86 Word compatible with an x64
        // renderer while rejecting a wrong-bitness override before LoadLibrary.
        var architecture = Environment.Is64BitProcess ? "64" : "32";
        var candidates = new List<string>();
        // VisualTeX can carry a private MathPage runtime next to the packaged
        // preview sidecar. Keep it out of Word's STARTUP path: the WLL is loaded
        // only inside this isolated helper process and never registered as a
        // Word add-in. It is the fallback when no compatible MathType 7 is installed.
        var bundledRuntimeRoot = Path.Combine(
            AppContext.BaseDirectory,
            "mathtype-runtime");
        foreach (var bundledMathPage in new[]
                 {
                     Path.Combine(
                         bundledRuntimeRoot,
                         "MathPage",
                         architecture,
                         "MathPage.wll"),
                     Path.Combine(bundledRuntimeRoot, "MathPage", "MathPage.wll"),
                     Path.Combine(bundledRuntimeRoot, "MathPage.wll"),
                 })
        {
            if (File.Exists(bundledMathPage)
                && IsNativeLibraryCompatibleWithCurrentProcess(bundledMathPage))
                return bundledMathPage;
        }
        if (!string.IsNullOrWhiteSpace(mathTypeServerPath))
        {
            var expandedServer = Environment.ExpandEnvironmentVariables(
                mathTypeServerPath.Trim().Trim('"'));
            var installRoot = Path.GetDirectoryName(expandedServer);
            if (!string.IsNullOrWhiteSpace(installRoot))
                AddMathPageCandidates(candidates, installRoot!, architecture);
        }
        foreach (var root in new[]
                 {
                     Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                     Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                 })
        {
            if (string.IsNullOrWhiteSpace(root)) continue;
            AddMathPageCandidates(
                candidates,
                Path.Combine(root, "MathType"),
                architecture);
            AddMathPageCandidates(
                candidates,
                Path.Combine(root, "WIRIS", "MathType"),
                architecture);
        }
        return candidates
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(path =>
                File.Exists(path)
                && IsNativeLibraryCompatibleWithCurrentProcess(path));
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

        // MathType point releases have used more than one MathPage subdirectory.
        // Search only the bounded MathPage subtree, never an entire Program Files
        // hierarchy, and validate PE architecture before loading a candidate.
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
            // Explicit candidates above remain usable when enumeration is denied.
        }
    }

    private static bool IsNativeLibraryCompatibleWithCurrentProcess(string path)
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
            return System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture switch
            {
                System.Runtime.InteropServices.Architecture.X86 => machine == 0x014C,
                System.Runtime.InteropServices.Architecture.X64 => machine == 0x8664,
                System.Runtime.InteropServices.Architecture.Arm => machine == 0x01C4,
                System.Runtime.InteropServices.Architecture.Arm64 => machine == 0xAA64,
                _ => false,
            };
        }
        catch
        {
            return false;
        }
    }

    private static T? GetDelegate<T>(IntPtr module, string name) where T : class
    {
        var address = GetProcAddress(module, name);
        if (address == IntPtr.Zero) return null;
        return Marshal.GetDelegateForFunctionPointer(address, typeof(T)) as T;
    }

    private static readonly IntPtr HkeyCurrentUser = new(unchecked((int)0x80000001));
    private static readonly IntPtr HkeyLocalMachine = new(unchecked((int)0x80000002));
    private const int KeyRead = 0x20019;

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
    private static extern int RegOpenKeyExW(
        IntPtr hKey,
        string lpSubKey,
        int ulOptions,
        int samDesired,
        out IntPtr phkResult);

    [DllImport("advapi32.dll")]
    private static extern int RegOverridePredefKey(IntPtr hKey, IntPtr hNewHKey);

    [DllImport("advapi32.dll")]
    private static extern int RegCloseKey(IntPtr hKey);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr LoadLibraryW(string path);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetDllDirectoryW(string? path);

    [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
    private static extern IntPtr GetProcAddress(IntPtr module, string name);

    [DllImport("kernel32.dll")]
    private static extern bool FreeLibrary(IntPtr module);
}
