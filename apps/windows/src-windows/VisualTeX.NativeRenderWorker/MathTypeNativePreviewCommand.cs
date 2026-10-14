using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Win32;

namespace VisualTeX.NativeRenderWorker;

/// <summary>
/// Native rendering implementation hosted only by the dedicated service worker.
/// The worker's session, identity and owned job isolate legacy MathPage failures
/// from Office and the interactive MathType application.
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
    // Worker exit status for a session that cannot resolve MathType's fonts.
    // NativeWorkerBackend maps it to NativeRenderProtocol.FontsUnavailableError.
    internal const int FontsUnavailableExitCode = 6;

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

    internal static int Run(string manifestPath, string requiredPrivateRuntimeRoot)
    {
        using var session = new NativeSession(requiredPrivateRuntimeRoot);
        return session.Render(manifestPath);
    }

    // One STA session belongs to one service-owned worker. Keeping only the job
    // alive is insufficient: MTTermAPI releases MathPage's last native lease.
    internal sealed class NativeSession : IDisposable
    {
        private readonly string requiredPrivateRuntimeRoot;
        private IntPtr module = IntPtr.Zero;
        private MtTermApiDelegate? term;
        private MtXFormEqnDelegate? transform;
        private MtGetLastDimensionDelegate? getDimension;
        private bool initialized;
        private bool privateDllDirectoryApplied;
        private PrivateRuntimeScope? privateRuntimeScope;

        internal NativeSession(string runtimeRoot) => requiredPrivateRuntimeRoot = runtimeRoot;

        internal int Render(string manifestPath)
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
        try
        {
            if (!initialized)
            {
            var mathPage = Path.Combine(Path.GetFullPath(requiredPrivateRuntimeRoot), "MathPage", "64", "MathPage.wll");
            if (mathPage is null || !File.Exists(mathPage)
                || !IsNativeLibraryCompatibleWithCurrentProcess(mathPage))
            {
                response.Error = "MathType MathPage.wll is not installed.";
                return WriteResponse(request.ResultPath, response, 3);
            }

            // MathType must find its own fonts before it starts. Otherwise it
            // writes placeholders into pictures that Office then saves.
            if (!MathTypeFontGuard.Ensure(Path.GetFullPath(requiredPrivateRuntimeRoot), out var fontError))
            {
                response.Error = fontError;
                return WriteResponse(request.ResultPath, response, FontsUnavailableExitCode);
            }

            // Discovery is confined to this exact bundled dependency tree.
            {
                var runtimeDirectory = Path.GetDirectoryName(mathPage);
                if (!string.IsNullOrWhiteSpace(runtimeDirectory))
                    privateDllDirectoryApplied = SetDllDirectoryW(runtimeDirectory);
                if (!PrivateRuntimeScope.TryCreate(
                        requiredPrivateRuntimeRoot,
                        mathPage,
                        out privateRuntimeScope,
                        out var privateRuntimeError))
                {
                    response.Error = privateRuntimeError;
                    return WriteResponse(request.ResultPath, response, 3);
                }
            }

            module = LoadLibraryW(mathPage);
            if (module == IntPtr.Zero)
            {
                response.Error = "MathType MathPage.wll could not be loaded.";
                return WriteResponse(request.ResultPath, response, 3);
            }

            var init = GetDelegate<MtInitApiDelegate>(module, "MTInitAPI");
            term = GetDelegate<MtTermApiDelegate>(module, "MTTermAPI");
            transform = GetDelegate<MtXFormEqnDelegate>(module, "MTXFormEqn");
            getDimension = GetDelegate<MtGetLastDimensionDelegate>(module, "MTGetLastDimension");
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
                    var status = transform!(
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

                    var width32 = getDimension!(1);
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
        }

        public void Dispose()
        {
            if (initialized && term is not null)
            {
                try { term(); } catch { }
                initialized = false;
            }
            if (module != IntPtr.Zero) { FreeLibrary(module); module = IntPtr.Zero; }
            privateRuntimeScope?.Dispose();
            privateRuntimeScope = null;
            if (privateDllDirectoryApplied) { SetDllDirectoryW(null); privateDllDirectoryApplied = false; }
        }
    }

    private const string PrivateMathTypeClsid =
        "{0002CE03-0000-0000-C000-000000000046}";
    private const string PrivateRuntimeMutexName =
        @"Local\VisualTeX.PrivateMathTypeRuntimeBootstrap.v1";

    private sealed class PrivateRuntimeScope : IDisposable
    {
        private readonly Mutex _mutex;
        private readonly string _virtualSessionPath;
        private IntPtr _virtualLocalMachineHandle;
        private IntPtr _virtualClassesRootHandle;
        private bool _ownsMutex;
        private bool _localMachineOverridden;
        private bool _classesRootOverridden;
        private bool _disposed;

        private PrivateRuntimeScope(
            Mutex mutex,
            bool ownsMutex,
            string virtualSessionPath,
            IntPtr virtualLocalMachineHandle,
            IntPtr virtualClassesRootHandle,
            bool localMachineOverridden,
            bool classesRootOverridden)
        {
            _mutex = mutex;
            _ownsMutex = ownsMutex;
            _virtualSessionPath = virtualSessionPath;
            _virtualLocalMachineHandle = virtualLocalMachineHandle;
            _virtualClassesRootHandle = virtualClassesRootHandle;
            _localMachineOverridden = localMachineOverridden;
            _classesRootOverridden = classesRootOverridden;
        }

        internal static bool TryCreate(
            string runtimeRoot,
            string mathPagePath,
            out PrivateRuntimeScope? scope,
            out string error)
        {
            scope = null;
            error = string.Empty;
            if (!ValidatePrivateRuntime(runtimeRoot, mathPagePath, out error))
                return false;

            Mutex? mutex = null;
            string? virtualSessionPath = null;
            IntPtr virtualLocalMachineHandle = IntPtr.Zero;
            IntPtr virtualClassesRootHandle = IntPtr.Zero;
            var localMachineOverridden = false;
            var classesRootOverridden = false;
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
                DeleteAbandonedPrivateVirtualRegistry();
                virtualSessionPath = CreatePrivateVirtualRegistry(
                    runtimeRoot,
                    serverPath,
                    out var virtualLocalMachinePath,
                    out var virtualClassesRootPath);
                var openStatus = RegOpenKeyExW(
                    HkeyCurrentUser,
                    virtualLocalMachinePath,
                    0,
                    KeyRead,
                    out virtualLocalMachineHandle);
                if (openStatus != 0 || virtualLocalMachineHandle == IntPtr.Zero)
                    throw new InvalidOperationException(
                        $"VisualTeX private MathType virtual HKLM root could not be opened; status {openStatus}.");

                var overrideStatus = RegOverridePredefKey(
                    HkeyLocalMachine,
                    virtualLocalMachineHandle);
                if (overrideStatus != 0)
                    throw new InvalidOperationException(
                        $"VisualTeX private MathType virtual HKLM override failed; status {overrideStatus}.");
                localMachineOverridden = true;

                openStatus = RegOpenKeyExW(
                    HkeyCurrentUser,
                    virtualClassesRootPath,
                    0,
                    KeyRead,
                    out virtualClassesRootHandle);
                if (openStatus != 0 || virtualClassesRootHandle == IntPtr.Zero)
                    throw new InvalidOperationException(
                        $"VisualTeX private MathType virtual HKCR root could not be opened; status {openStatus}.");

                overrideStatus = RegOverridePredefKey(
                    HkeyClassesRoot,
                    virtualClassesRootHandle);
                if (overrideStatus != 0)
                    throw new InvalidOperationException(
                        $"VisualTeX private MathType virtual HKCR override failed; status {overrideStatus}.");
                classesRootOverridden = true;

                // MTInitAPI owns launch-as-needed. Starting MathType explicitly
                // for every sidecar also launches it when its RPC server already
                // exists; MathType forwards that second launch to its editor and
                // opens an Untitled window. Keep discovery process-local, and
                // let MathPage connect to or start the renderer exactly once.

                scope = new PrivateRuntimeScope(
                    mutex,
                    ownsMutex,
                    virtualSessionPath,
                    virtualLocalMachineHandle,
                    virtualClassesRootHandle,
                    localMachineOverridden,
                    classesRootOverridden);
                return true;
            }
            catch (Exception exception)
            {
                error = exception.Message;
                if (classesRootOverridden)
                {
                    try { RegOverridePredefKey(HkeyClassesRoot, IntPtr.Zero); }
                    catch { }
                }
                if (localMachineOverridden)
                {
                    try { RegOverridePredefKey(HkeyLocalMachine, IntPtr.Zero); }
                    catch { }
                }
                if (virtualClassesRootHandle != IntPtr.Zero)
                    RegCloseKey(virtualClassesRootHandle);
                if (virtualLocalMachineHandle != IntPtr.Zero)
                    RegCloseKey(virtualLocalMachineHandle);
                if (!string.IsNullOrWhiteSpace(virtualSessionPath))
                    DeletePrivateVirtualRegistry(virtualSessionPath!);
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
            if (_classesRootOverridden)
            {
                try { RegOverridePredefKey(HkeyClassesRoot, IntPtr.Zero); }
                catch { }
                _classesRootOverridden = false;
            }
            if (_localMachineOverridden)
            {
                try { RegOverridePredefKey(HkeyLocalMachine, IntPtr.Zero); }
                catch { }
                _localMachineOverridden = false;
            }
            if (_virtualClassesRootHandle != IntPtr.Zero)
            {
                RegCloseKey(_virtualClassesRootHandle);
                _virtualClassesRootHandle = IntPtr.Zero;
            }
            if (_virtualLocalMachineHandle != IntPtr.Zero)
            {
                RegCloseKey(_virtualLocalMachineHandle);
                _virtualLocalMachineHandle = IntPtr.Zero;
            }
            DeletePrivateVirtualRegistry(_virtualSessionPath);
            if (_ownsMutex)
            {
                try { _mutex.ReleaseMutex(); } catch { }
                _ownsMutex = false;
            }
            _mutex.Dispose();
        }
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
            Path.Combine("Language", "MT7chs.dll"),
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
        // RegOverridePredefKey is process-local; the native MathType child does
        // not inherit it. Do not start a reduced child that would fall back to
        // English and overwrite the standalone application's saved locale.
        using var user = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32);
        using var configuration = user.OpenSubKey(@"Software\Design Science\DSMT7\Config");
        var appLanguage = configuration?.GetValue("AppLang") as string;
        if (!string.IsNullOrWhiteSpace(appLanguage))
        {
            var parts = appLanguage.Split(',');
            var language = parts.Length == 2 ? parts[1].Trim() : string.Empty;
            if (language.Length != 3 || !language.All(character =>
                    character is >= 'A' and <= 'Z' or >= 'a' and <= 'z'))
            {
                error = "MathType's saved language is invalid; the private renderer will not change it.";
                return false;
            }
            if (!string.Equals(language, "ENU", StringComparison.OrdinalIgnoreCase)
                && !File.Exists(Path.Combine(root, "Language", $"MT7{language}.dll")))
            {
                error = $"The private MathType renderer does not include the saved {language} language. "
                    + "Use a complete MathType installation; the saved language was not changed.";
                return false;
            }
        }
        return true;
    }

    private static string CreatePrivateVirtualRegistry(
        string runtimeRoot,
        string serverPath,
        out string virtualLocalMachinePath,
        out string virtualClassesRootPath)
    {
        var sessionPath = $@"Software\VisualTeX\PrivateMathTypeRuntime\Sessions\{Environment.ProcessId}-{Guid.NewGuid():N}";
        virtualLocalMachinePath = sessionPath + @"\HKLM";
        virtualClassesRootPath = sessionPath + @"\HKCR";
        using var baseKey = RegistryKey.OpenBaseKey(
            RegistryHive.CurrentUser,
            RegistryView.Registry64);
        using (var directories = baseKey.CreateSubKey(
                   virtualLocalMachinePath + @"\SOFTWARE\Design Science\DSMT7\Directories",
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
                   virtualLocalMachinePath + @"\SOFTWARE\Microsoft\Windows\CurrentVersion",
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

        using (var classKey = baseKey.CreateSubKey(
                   virtualClassesRootPath + $@"\CLSID\{PrivateMathTypeClsid}",
                   writable: true)
               ?? throw new InvalidOperationException(
                   "VisualTeX could not create the private MathType class view."))
        {
            classKey.SetValue(string.Empty, "MathType 7 Equation", RegistryValueKind.String);
        }
        using (var localServer = baseKey.CreateSubKey(
                   virtualClassesRootPath + $@"\CLSID\{PrivateMathTypeClsid}\LocalServer32",
                   writable: true)
               ?? throw new InvalidOperationException(
                   "VisualTeX could not create the private MathType server view."))
        {
            localServer.SetValue(string.Empty, serverPath, RegistryValueKind.String);
        }
        foreach (var progId in new[] { "Equation.DSMT4", "DSEquations" })
        {
            using var clsid = baseKey.CreateSubKey(
                virtualClassesRootPath + $@"\{progId}\CLSID",
                writable: true);
            clsid?.SetValue(string.Empty, PrivateMathTypeClsid, RegistryValueKind.String);
        }

        return sessionPath;
    }

    private static void SetRegistryStringIfPresent(
        RegistryKey key,
        string name,
        string value)
    {
        if (!string.IsNullOrWhiteSpace(value))
            key.SetValue(name, value, RegistryValueKind.String);
    }

    private static void DeleteAbandonedPrivateVirtualRegistry()
    {
        try
        {
            using var baseKey = RegistryKey.OpenBaseKey(
                RegistryHive.CurrentUser,
                RegistryView.Registry64);
            baseKey.DeleteSubKeyTree(
                @"Software\VisualTeX\PrivateMathTypeRuntime\Sessions",
                throwOnMissingSubKey: false);
            // Releases before this fix used a separate VirtualHKLM tree. It was
            // never part of MathType's public registration and is safe to prune.
            baseKey.DeleteSubKeyTree(
                @"Software\VisualTeX\PrivateMathTypeRuntime\VirtualHKLM",
                throwOnMissingSubKey: false);
            DeleteRegistryKeyIfEmpty(
                baseKey,
                @"Software\VisualTeX\PrivateMathTypeRuntime");
        }
        catch { }
    }

    private static void DeletePrivateVirtualRegistry(string sessionPath)
    {
        if (string.IsNullOrWhiteSpace(sessionPath)) return;
        try
        {
            using var baseKey = RegistryKey.OpenBaseKey(
                RegistryHive.CurrentUser,
                RegistryView.Registry64);
            baseKey.DeleteSubKeyTree(sessionPath, throwOnMissingSubKey: false);
            DeleteRegistryKeyIfEmpty(
                baseKey,
                @"Software\VisualTeX\PrivateMathTypeRuntime\Sessions");
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

    private static readonly IntPtr HkeyClassesRoot = new(unchecked((int)0x80000000));
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
