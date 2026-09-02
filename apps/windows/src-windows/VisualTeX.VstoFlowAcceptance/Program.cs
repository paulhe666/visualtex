using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Win32;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using WinForms = System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const float ExportWidth = 160f;
    private const float ExportHeight = 32f;
    private const float ExportBaseline = 24f;
    private const uint MouseLeftDown = 0x0002;
    private const uint MouseLeftUp = 0x0004;
    private const double WordPerformanceLimitMilliseconds = 500.0;

    [ComImport]
    [Guid("00000016-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(
            int callType,
            IntPtr callerTask,
            int tickCount,
            IntPtr interfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(
            IntPtr calleeTask,
            int tickCount,
            int rejectType);

        [PreserveSig]
        int MessagePending(
            IntPtr calleeTask,
            int tickCount,
            int pendingType);
    }

    private sealed class OfficeComMessageFilter : IOleMessageFilter, IDisposable
    {
        private IOleMessageFilter? _previous;
        private bool _disposed;

        private OfficeComMessageFilter()
        {
        }

        internal static OfficeComMessageFilter Register()
        {
            var filter = new OfficeComMessageFilter();
            var result = CoRegisterMessageFilter(filter, out filter._previous);
            if (result < 0)
                Marshal.ThrowExceptionForHR(result);
            return filter;
        }

        public int HandleInComingCall(
            int callType,
            IntPtr callerTask,
            int tickCount,
            IntPtr interfaceInfo) => 0;

        public int RetryRejectedCall(
            IntPtr calleeTask,
            int tickCount,
            int rejectType)
        {
            const int ServerCallRejected = 1;
            const int ServerCallRetryLater = 2;
            if ((rejectType == ServerCallRejected || rejectType == ServerCallRetryLater)
                && tickCount < 30_000)
                return 100;
            return -1;
        }

        public int MessagePending(
            IntPtr calleeTask,
            int tickCount,
            int pendingType) => 2;

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try { CoRegisterMessageFilter(_previous, out _); }
            catch { }
            _previous = null;
        }
    }

    private sealed class UserOfficeAddInAutoLoadSuppression : IDisposable
    {
        private readonly string _keyPath;
        private readonly bool _keyExisted;
        private readonly bool _loadBehaviorExisted;
        private readonly object? _loadBehavior;
        private readonly RegistryValueKind _loadBehaviorKind;
        private bool _disposed;

        private UserOfficeAddInAutoLoadSuppression(
            string keyPath,
            bool keyExisted,
            bool loadBehaviorExisted,
            object? loadBehavior,
            RegistryValueKind loadBehaviorKind)
        {
            _keyPath = keyPath;
            _keyExisted = keyExisted;
            _loadBehaviorExisted = loadBehaviorExisted;
            _loadBehavior = loadBehavior;
            _loadBehaviorKind = loadBehaviorKind;
        }

        internal static UserOfficeAddInAutoLoadSuppression Create(
            string applicationName,
            string progId)
        {
            var keyPath = $@"Software\Microsoft\Office\{applicationName}\Addins\{progId}";
            using var existingKey = Registry.CurrentUser.OpenSubKey(keyPath);
            var keyExisted = existingKey is not null;
            using var key = Registry.CurrentUser.CreateSubKey(keyPath, writable: true)
                ?? throw new InvalidOperationException(
                    $"The per-user {applicationName} add-in acceptance override could not be created.");
            var valueNames = key.GetValueNames();
            var loadBehaviorExisted = valueNames.Any(name =>
                string.Equals(name, "LoadBehavior", StringComparison.OrdinalIgnoreCase));
            var loadBehavior = loadBehaviorExisted
                ? key.GetValue("LoadBehavior", null, RegistryValueOptions.DoNotExpandEnvironmentNames)
                : null;
            var loadBehaviorKind = loadBehaviorExisted
                ? key.GetValueKind("LoadBehavior")
                : RegistryValueKind.DWord;
            if (!keyExisted)
            {
                key.SetValue("FriendlyName", "VisualTeX", RegistryValueKind.String);
                key.SetValue(
                    "Description",
                    "VisualTeX VSTO flow acceptance auto-load suppression",
                    RegistryValueKind.String);
            }
            key.SetValue("LoadBehavior", 0, RegistryValueKind.DWord);
            return new UserOfficeAddInAutoLoadSuppression(
                keyPath,
                keyExisted,
                loadBehaviorExisted,
                loadBehavior,
                loadBehaviorKind);
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            if (!_keyExisted)
            {
                try { Registry.CurrentUser.DeleteSubKeyTree(_keyPath, throwOnMissingSubKey: false); }
                catch { }
                return;
            }

            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(_keyPath, writable: true);
                if (key is null) return;
                if (_loadBehaviorExisted && _loadBehavior is not null)
                    key.SetValue("LoadBehavior", _loadBehavior, _loadBehaviorKind);
                else
                    key.DeleteValue("LoadBehavior", throwOnMissingValue: false);
            }
            catch { }
        }
    }

    private sealed class PerformanceFormulaEntry
    {
        public int Index { get; set; }
        public string FormulaId { get; set; } = string.Empty;
        public string ObjectMode { get; set; } = string.Empty;
        public string DisplayMode { get; set; } = string.Empty;
        public string Latex { get; set; } = string.Empty;
    }

    private sealed class PerformanceTimingEntry
    {
        public string Phase { get; set; } = string.Empty;
        public int Index { get; set; }
        public string FormulaId { get; set; } = string.Empty;
        public string ObjectMode { get; set; } = string.Empty;
        public string DisplayMode { get; set; } = string.Empty;
        public double OpenMilliseconds { get; set; }
        public double ApplyMilliseconds { get; set; }
    }

    private sealed class PerformanceSummary
    {
        public string Phase { get; set; } = string.Empty;
        public int Count { get; set; }
        public double OpenP50Milliseconds { get; set; }
        public double OpenP95Milliseconds { get; set; }
        public double OpenMaximumMilliseconds { get; set; }
        public double ApplyP50Milliseconds { get; set; }
        public double ApplyP95Milliseconds { get; set; }
        public double ApplyMaximumMilliseconds { get; set; }
    }

    private delegate bool EnumWindowsCallback(IntPtr windowHandle, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsCallback callback, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr windowHandle);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(IntPtr windowHandle);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(
        IntPtr windowHandle,
        StringBuilder text,
        int maximumCount);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(
        IntPtr windowHandle,
        out uint processId);

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(
        IOleMessageFilter? newFilter,
        out IOleMessageFilter? oldFilter);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr windowHandle);

    [DllImport("user32.dll")]
    private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiContext);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GetConsoleWindow();

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr windowHandle, int command);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr windowHandle, out NativeRect rectangle);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(
        IntPtr windowHandle,
        IntPtr insertAfter,
        int x,
        int y,
        int width,
        int height,
        uint flags);

    [DllImport("user32.dll")]
    private static extern void mouse_event(
        uint flags,
        uint dx,
        uint dy,
        uint data,
        UIntPtr extraInfo);

    private static readonly string SessionRoot = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "com.visualtex.studio",
        "office",
        "sessions");

    private static bool AttachActiveWord => string.Equals(
        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE_ATTACH_WORD"),
        "1",
        StringComparison.OrdinalIgnoreCase);

    private static Word.Application CreateWordApplication(bool visible)
    {
        if (AttachActiveWord)
        {
            var active = Marshal.GetActiveObject("Word.Application") as Word.Application
                ?? throw new InvalidOperationException("No active Word instance is available.");
            active.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            if (visible) active.Visible = true;
            return active;
        }

        // Register the pinned font before WINWORD.EXE is created. Office Math and
        // Word's PDF exporter snapshot available math fonts during process startup;
        // registering the same file only after Word starts can leave equations on
        // Cambria fallback even though the OMML run names say Latin Modern Math.
        Console.WriteLine("Acceptance Word startup: verifying Latin Modern Math.");
        WordOfficeMathFontLoader.EnsureLoaded();
        Console.WriteLine(
            $"Acceptance Word startup: font ready path='{WordOfficeMathFontLoader.LoadedPath}', sessionRegistration={WordOfficeMathFontLoader.SessionRegistrationUsed}.");
        Console.WriteLine("Acceptance Word startup: creating WINWORD COM application.");
        var showInteractiveWordAlerts = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_WORD_ALERTS"),
            "1",
            StringComparison.Ordinal);
        var created = new Word.Application
        {
            Visible = visible,
            DisplayAlerts = showInteractiveWordAlerts
                ? Word.WdAlertLevel.wdAlertsAll
                : Word.WdAlertLevel.wdAlertsNone,
        };
        if (showInteractiveWordAlerts)
            Console.WriteLine("Acceptance Word startup: interactive Word alerts ENABLED.");
        Console.WriteLine("Acceptance Word startup: WINWORD COM application created.");
        try
        {
            var hwnd = Convert.ToInt32(((dynamic)created).Hwnd);
            _ = GetWindowThreadProcessId(new IntPtr(hwnd), out var processId);
            Console.WriteLine(
                $"Acceptance Word application: PID={processId}, Hwnd={hwnd}, fontPath='{WordOfficeMathFontLoader.LoadedPath}', sessionRegistration={WordOfficeMathFontLoader.SessionRegistrationUsed}.");
        }
        catch { }
        return created;
    }

    private static void QuitWordApplicationIfOwned(Word.Application? application)
    {
        if (application is null || AttachActiveWord) return;
        application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
        WordOfficeMathFontLoader.UnloadSessionRegistration();
    }

    private static string ResolveAcceptanceArtifactRoot(
        string? artifactArgument,
        string mode)
    {
        var configuredRoot = Environment.GetEnvironmentVariable(
            "VISUALTEX_ACCEPTANCE_ARTIFACT_ROOT");
        var baseRoot = string.IsNullOrWhiteSpace(configuredRoot)
            ? Path.Combine(
                Path.GetTempPath(),
                "VisualTeX",
                "acceptance",
                "vsto-flow")
            : Path.GetFullPath(Environment.ExpandEnvironmentVariables(
                configuredRoot!.Trim().Trim('"')));

        if (string.IsNullOrWhiteSpace(artifactArgument))
        {
            var safeMode = string.Concat((mode ?? "all").Select(character =>
                char.IsLetterOrDigit(character) || character is '-' or '_'
                    ? character
                    : '-'));
            return Path.Combine(
                baseRoot,
                $"{safeMode}-{DateTime.Now:yyyyMMdd-HHmmss-fff}-{Guid.NewGuid():N}");
        }

        var expanded = Environment.ExpandEnvironmentVariables(
            artifactArgument!.Trim().Trim('"'));
        if (Path.IsPathRooted(expanded))
            return Path.GetFullPath(expanded);

        // Relative acceptance paths used to be resolved against the checkout and
        // created hundreds of artifacts/* and Tempvisualtex-* directories beside
        // source files. Keep the caller's grouping, but anchor it under one OS-temp
        // root. Reject traversal rather than allowing a relative argument to escape.
        var segments = expanded
            .Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar)
            .Split(new[] { Path.DirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries)
            .Where(segment => !string.Equals(segment, ".", StringComparison.Ordinal))
            .ToArray();
        if (segments.Any(segment => string.Equals(segment, "..", StringComparison.Ordinal)))
            throw new InvalidDataException(
                "A relative acceptance artifact path cannot contain '..'. Use an absolute path for an explicit external destination.");
        var relative = segments.Length == 0
            ? $"{mode}-{DateTime.Now:yyyyMMdd-HHmmss-fff}"
            : Path.Combine(segments);
        return Path.GetFullPath(Path.Combine(baseRoot, relative));
    }

    [STAThread]
    private static int Main(string[] args)
    {
        try { SetProcessDpiAwarenessContext(new IntPtr(-4)); } catch { }
        var instanceMutexName = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_ACCEPTANCE_MUTEX_NAME");
        if (string.IsNullOrWhiteSpace(instanceMutexName))
            instanceMutexName = @"Local\VisualTeX.VstoFlowAcceptance";
        using var instanceMutex = new Mutex(
            initiallyOwned: true,
            name: instanceMutexName,
            createdNew: out var createdNew);
        if (!createdNew)
        {
            Console.Error.WriteLine("Another VisualTeX VSTO acceptance instance is already running.");
            return 4;
        }

        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", "1");
        var mode = args
            .FirstOrDefault(argument => argument.StartsWith("--mode=", StringComparison.OrdinalIgnoreCase))
            ?.Substring("--mode=".Length)
            ?? "all";
        var artifactArgument = args.FirstOrDefault(argument => !argument.StartsWith("--", StringComparison.Ordinal));
        var artifactRoot = ResolveAcceptanceArtifactRoot(artifactArgument, mode);
        Directory.CreateDirectory(artifactRoot);

        using var log = new StreamWriter(
            Path.Combine(artifactRoot, "acceptance.log"),
            append: false,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false))
        {
            AutoFlush = true,
        };
        var originalOut = Console.Out;
        var originalError = Console.Error;
        Console.SetOut(new TeeTextWriter(originalOut, log));
        Console.SetError(new TeeTextWriter(originalError, log));
        Console.WriteLine($"Acceptance mode: {mode}");

        using var officeMessageFilter = OfficeComMessageFilter.Register();
        var exerciseInstalledWordAddIn = string.Equals(
                mode,
                "word-installed-unsaved-first-omml",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-mathtype-left-right-stability",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-mathtype-right-left-live",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-installed-format-conversion",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-live-format-conversion-fixture",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-installed-visualtex-number-toggle-close",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-installed-omml-mathtype-format-conversion",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-installed-mathtype-native-regression",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-installed-mathtype-left-ui-e2e",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-installed-mathtype-reedit-regression",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-installed-mathtype-50-edit-stress",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-installed-format-50-batch-stress",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-installed-vt-omml-50-direct-stress",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-user-100-mathtype-conversion",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-user-100-mathtype-reverse",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-user-100-omml-to-mathtype",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-active-user-100-failure-inspect",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-active-mathtype-omml-source-audit",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                mode,
                "word-active-mathtype-omml-copy-diagnostic",
                StringComparison.OrdinalIgnoreCase);
        using var installedWordAutoLoadSuppression = exerciseInstalledWordAddIn
            ? null
            : UserOfficeAddInAutoLoadSuppression.Create("Word", "VisualTeX.WordVsto");
        var exerciseInstalledPowerPointAddIn = string.Equals(
            mode,
            "powerpoint-installed-ole-presentation",
            StringComparison.OrdinalIgnoreCase);
        using var installedPowerPointAutoLoadSuppression = exerciseInstalledPowerPointAddIn
            ? null
            : UserOfficeAddInAutoLoadSuppression.Create("PowerPoint", "VisualTeX.PowerPointVsto");
        Console.WriteLine(exerciseInstalledWordAddIn
            ? "Installed Word add-in remains enabled for this MathType real-environment stability acceptance."
            : exerciseInstalledPowerPointAddIn
                ? "Installed Word add-in auto-load is suppressed; installed PowerPoint add-in remains enabled for this acceptance."
                : "Installed Word and PowerPoint add-in auto-load is suppressed for this isolated acceptance process.");

        // Pure COM/Word performance acceptance does not depend on the desktop
        // companion. Running it before VisualTeXSessionClient also prevents an
        // unrelated companion-health or assembly-binding failure from polluting
        // the timing data we are trying to measure.
        if (string.Equals(
                mode,
                "word-sparse-numbered-omml-performance",
                StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                RunWordSparseNumberedOmmlPerformanceAcceptance(artifactRoot);
                Console.WriteLine("VisualTeX real VSTO formula flow acceptance passed.");
                Console.WriteLine($"Artifacts: {artifactRoot}");
                return 0;
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error);
                Console.Error.WriteLine($"Acceptance artifacts retained: {artifactRoot}");
                return 1;
            }
        }

        try
        {
            using var client = new VisualTeXSessionClient();
            client.EnsureHealthyAsync(CancellationToken.None).GetAwaiter().GetResult();
            if (string.Equals(mode, "health", StringComparison.OrdinalIgnoreCase))
            {
                // Run twice on the same client so the second request exercises
                // HttpClient TLS connection pooling, where the certificate
                // callback may not run again for an already-authenticated socket.
                client.EnsureHealthyAsync(CancellationToken.None).GetAwaiter().GetResult();
                Console.WriteLine("VisualTeX .NET Framework companion health acceptance passed twice on one pooled client.");
            }
            else if (string.Equals(mode, "office-session-mathtype-number-position", StringComparison.OrdinalIgnoreCase))
            {
                RunOfficeSessionMathTypeNumberPositionAcceptance(client);
            }
            else if (string.Equals(mode, "word-native-crossref-probe", StringComparison.OrdinalIgnoreCase))
            {
                ProbeNativeEquationCrossReference();
            }
            else if (string.Equals(mode, "word-crossref", StringComparison.OrdinalIgnoreCase))
            {
                RunWordNativeCrossReference(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-formula-fonts", StringComparison.OrdinalIgnoreCase))
            {
                RunWordFormulaFontAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-inline-fraction-line-grid", StringComparison.OrdinalIgnoreCase))
            {
                RunWordInlineFractionLineGridAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "office-ole-formula-fonts", StringComparison.OrdinalIgnoreCase))
            {
                RunOleFormulaFontAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-create", StringComparison.OrdinalIgnoreCase))
            {
                RunWord(client, artifactRoot, initialOnly: true);
            }
            else if (string.Equals(mode, "powerpoint-context-safety", StringComparison.OrdinalIgnoreCase))
            {
                RunPowerPointContextSafetyAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "powerpoint-dense-zorder", StringComparison.OrdinalIgnoreCase))
            {
                RunPowerPointDenseZOrderAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "powerpoint-ole-svg-delete", StringComparison.OrdinalIgnoreCase))
            {
                RunPowerPointOleSvgDeleteAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "powerpoint-font-size", StringComparison.OrdinalIgnoreCase))
            {
                RunPowerPointFontSizeAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "powerpoint-copy-conversion", StringComparison.OrdinalIgnoreCase))
            {
                RunPowerPointCopyConversionAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "powerpoint-ole-svg-geometry", StringComparison.OrdinalIgnoreCase))
            {
                RunPowerPointOleSvgGeometryAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "powerpoint-current-geometry-probe", StringComparison.OrdinalIgnoreCase))
            {
                RunPowerPointCurrentGeometryProbe();
            }
            else if (string.Equals(mode, "powerpoint-current-ole-cache-probe", StringComparison.OrdinalIgnoreCase))
            {
                RunPowerPointCurrentOleCacheProbe();
            }
            else if (string.Equals(mode, "powerpoint-installed-ole-presentation", StringComparison.OrdinalIgnoreCase))
            {
                RunPowerPointInstalledOlePresentationAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "powerpoint-direct-omml-stress", StringComparison.OrdinalIgnoreCase))
            {
                RunPowerPointDirectOmmlStressAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "powerpoint-picture-edit", StringComparison.OrdinalIgnoreCase))
            {
                RunPowerPoint(client, artifactRoot, stopAfterPictureEdit: true);
            }
            else if (string.Equals(mode, "powerpoint", StringComparison.OrdinalIgnoreCase))
            {
                RunPowerPoint(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-native-omml-first-double-click", StringComparison.OrdinalIgnoreCase))
            {
                RunWordNativeOmmlFirstDoubleClickAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-ole-interop", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeOleInteropAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "mathtype-standalone-com", StringComparison.OrdinalIgnoreCase))
            {
                RunMathTypeStandaloneComAcceptance();
            }
            else if (string.Equals(mode, "mathtype-set-format-probe", StringComparison.OrdinalIgnoreCase))
            {
                RunMathTypeSetFormatProbe();
            }
            else if (string.Equals(mode, "mathtype-ui-mathml-clipboard", StringComparison.OrdinalIgnoreCase))
            {
                RunMathTypeUiMathMlClipboardAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-ole-copy-formats", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeOleCopyFormatsAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-embedded-storage-probe", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeEmbeddedStorageProbe(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-offline-storage", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeOfflineStorageAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "mathtype-adjacent-batch-isolation", StringComparison.OrdinalIgnoreCase))
            {
                RunMathTypeAdjacentBatchIsolationAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-adjacent-format-conversion", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeAdjacentFormatConversionAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-caption-frame-repair", StringComparison.OrdinalIgnoreCase))
            {
                RunWordCaptionFrameRepairAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-visualtex-mathtype-adjacent-frame", StringComparison.OrdinalIgnoreCase))
            {
                RunWordVisualTeXMathTypeAdjacentFrameAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-active-user-100-failure-inspect", StringComparison.OrdinalIgnoreCase))
            {
                RunActiveUserHundredFailureInspection();
            }
            else if (string.Equals(mode, "word-active-mathtype-omml-source-audit", StringComparison.OrdinalIgnoreCase))
            {
                RunActiveMathTypeOmmlSourceAudit();
            }
            else if (string.Equals(mode, "word-active-mathtype-omml-copy-diagnostic", StringComparison.OrdinalIgnoreCase))
            {
                RunActiveMathTypeOmmlCopyDiagnostic(artifactRoot);
            }
            else if (string.Equals(mode, "word-active-mathtype-source-double-click", StringComparison.OrdinalIgnoreCase))
            {
                RunActiveMathTypeSourceDoubleClickProbe();
            }
            else if (string.Equals(mode, "word-user-100-mathtype-source-audit", StringComparison.OrdinalIgnoreCase))
            {
                RunUserHundredMathTypeSourceAudit();
            }
            else if (string.Equals(mode, "word-user-omml-mathtype-source-audit", StringComparison.OrdinalIgnoreCase))
            {
                RunUserOmmlMathTypeSourceAudit();
            }
            else if (string.Equals(mode, "word-user-100-omml-to-mathtype", StringComparison.OrdinalIgnoreCase))
            {
                RunUserOmmlMathTypeConversionAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-user-100-mathtype-reverse", StringComparison.OrdinalIgnoreCase))
            {
                RunUserHundredMathTypeReverseAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-user-100-mathtype-conversion", StringComparison.OrdinalIgnoreCase))
            {
                RunWordUserHundredMathTypeConversionAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-user-100-mathtype-preview-scan", StringComparison.OrdinalIgnoreCase))
            {
                RunUserHundredMathTypePreviewScan(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-standalone-codec", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeStandaloneCodecAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-insert-scaling", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeInsertScalingAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-insertion-complexity", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeInsertionComplexityAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-preview-fallback", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypePreviewFallbackAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "mathtype-mtef-root-compat", StringComparison.OrdinalIgnoreCase))
            {
                RunMathTypeMtefRootCompatibilityAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "mathtype-ole-storage-robustness", StringComparison.OrdinalIgnoreCase))
            {
                RunMathTypeOleStorageRobustnessAcceptance();
            }
            else if (string.Equals(mode, "mathtype-capability-resolution", StringComparison.OrdinalIgnoreCase))
            {
                RunMathTypeCapabilityResolutionAcceptance();
            }
            else if (string.Equals(mode, "word-mathtype-addole-from-cfb", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeAddOleFromCfbAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-direct-paste-symbols", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeDirectPasteSymbolsAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "mathtype-native-preview-complex", StringComparison.OrdinalIgnoreCase))
            {
                RunMathTypeNativePreviewComplexAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "mathtype-native-preview-single-performance", StringComparison.OrdinalIgnoreCase))
            {
                RunMathTypeNativePreviewSinglePerformanceAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "mathtype-native-preview-shared-lifecycle", StringComparison.OrdinalIgnoreCase))
            {
                RunMathTypeNativePreviewSharedLifecycleAcceptance();
            }
            else if (string.Equals(mode, "word-mathtype-openxml-clone", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeOpenXmlCloneAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "mathtype-window-inventory", StringComparison.OrdinalIgnoreCase))
            {
                RunMathTypeWindowInventory();
            }
            else if (string.Equals(mode, "mathtype-clean-acceptance-windows", StringComparison.OrdinalIgnoreCase))
            {
                RunMathTypeAcceptanceWindowCleanup();
            }
            else if (string.Equals(mode, "word-mathtype-ole-to-visualtex", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeOleToVisualTeXAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-ole-product-roundtrip", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeOleProductRoundTripAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-display-layout", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeDisplayLayoutAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-ole-create", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeOleCreateAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-ole-create-structure", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeOleCreateStructureAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-installed-format-conversion", StringComparison.OrdinalIgnoreCase))
            {
                RunWordInstalledFormatConversionAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-doc4-visualtex-source-fixture", StringComparison.OrdinalIgnoreCase))
            {
                RunWordDoc4VisualTeXSourceFixture(artifactRoot);
            }
            else if (string.Equals(mode, "word-live-format-conversion-fixture", StringComparison.OrdinalIgnoreCase))
            {
                RunWordLiveFormatConversionFixtureCapture(artifactRoot);
            }
            else if (string.Equals(mode, "word-format-conversion-rollback-residual-fixture", StringComparison.OrdinalIgnoreCase))
            {
                RunWordFormatConversionRollbackResidualFixtureAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-live-mathtype-dump", StringComparison.OrdinalIgnoreCase))
            {
                RunWordLiveMathTypeDump(artifactRoot);
            }
            else if (string.Equals(mode, "word-live-backup-unsaved", StringComparison.OrdinalIgnoreCase))
            {
                RunWordLiveUnsavedBackup(artifactRoot);
            }
            else if (string.Equals(mode, "word-simple-format-conversion-numbering", StringComparison.OrdinalIgnoreCase))
            {
                RunWordSimpleFormatConversionNumberingAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-single-numbered-omml-to-visualtex", StringComparison.OrdinalIgnoreCase))
            {
                RunWordSingleNumberedOmmlToVisualTeXAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-visualtex-numbered-roundtrip", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlVisualTeXNumberedRoundTripAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-mathtype-format-conversion", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlMathTypeFormatConversionAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-omml-consecutive-numbered", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeOmmlConsecutiveNumberedAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-mathtype-single-performance", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlMathTypeSinglePerformanceAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-omml-selection-view", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeOmmlSelectionViewAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-installed-omml-mathtype-format-conversion", StringComparison.OrdinalIgnoreCase))
            {
                RunWordInstalledOmmlMathTypeFormatConversionAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-installed-mathtype-native-regression", StringComparison.OrdinalIgnoreCase))
            {
                RunWordInstalledMathTypeNativeRegressionAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-installed-mathtype-left-ui-e2e", StringComparison.OrdinalIgnoreCase))
            {
                RunWordInstalledMathTypeLeftUiE2eAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-installed-mathtype-reedit-regression", StringComparison.OrdinalIgnoreCase))
            {
                RunWordInstalledMathTypeReeditRegressionAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-installed-mathtype-50-edit-stress", StringComparison.OrdinalIgnoreCase))
            {
                RunWordInstalledMathTypeFiftyEditStressAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-installed-format-50-batch-stress", StringComparison.OrdinalIgnoreCase))
            {
                RunWordInstalledFormatFiftyBatchStressAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-installed-vt-omml-50-direct-stress", StringComparison.OrdinalIgnoreCase))
            {
                RunWordInstalledVisualTeXOmmlFiftyDirectStressAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-right-left-live", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeRightThenLeftLiveAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-left-right-stability", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeLeftThenRightStabilityAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-native-number-reference", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeNativeNumberReferenceAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-number-format", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeNumberFormatAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "win32-thread-execution-probe", StringComparison.OrdinalIgnoreCase))
            {
                RunWin32ThreadExecutionProbe();
            }
            else if (string.Equals(mode, "ole-clipboard-flush-probe", StringComparison.OrdinalIgnoreCase))
            {
                RunOleClipboardFlushProbe();
            }
            else if (string.Equals(mode, "word-mathtype-native-editor", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeNativeEditorAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-direct-set-real-ole", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeDirectSetRealOleAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-ole-real-double-click", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeOleRealDoubleClickAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-double-click-event-probe", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlDoubleClickEventProbe(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-double-click-fixtures", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlDoubleClickFixtures(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-office2021-omml-boundary", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOffice2021OmmlBoundaryAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-ole-real-double-click", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOleRealDoubleClick(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-double-click-hit-test-existing", StringComparison.OrdinalIgnoreCase))
            {
                RunExistingWordDoubleClickHitTest(client);
            }
            else if (string.Equals(mode, "word-font-size", StringComparison.OrdinalIgnoreCase))
            {
                RunWordFontSizeAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-font-size-persistence", StringComparison.OrdinalIgnoreCase))
            {
                RunWordFontSizePersistenceAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-native-omml-source", StringComparison.OrdinalIgnoreCase))
            {
                RunWordNativeOmmlSourceProbe(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-ole-picture-roundtrip", StringComparison.OrdinalIgnoreCase))
            {
                RunWord(
                    client,
                    artifactRoot,
                    stopAfterOlePictureRoundTrip: true,
                    skipDoubleClickForOlePictureRoundTrip: true);
            }
            else if (string.Equals(mode, "word-unchanged", StringComparison.OrdinalIgnoreCase))
            {
                RunWord(client, artifactRoot, stopAfterUnchanged: true);
            }
            else if (string.Equals(mode, "word-100-performance", StringComparison.OrdinalIgnoreCase))
            {
                RunWordHundredFormulaPerformance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-latex-redraw-omml-only", StringComparison.OrdinalIgnoreCase))
            {
                RunWordLatexRedrawOmmlOnly(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-latex-redraw-mathtype", StringComparison.OrdinalIgnoreCase))
            {
                RunWordLatexRedrawMathTypeOnly(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-latex-redraw-distinct-formulas", StringComparison.OrdinalIgnoreCase))
            {
                RunWordLatexRedrawDistinctFormulas(artifactRoot);
            }
            else if (string.Equals(mode, "word-latex-redraw-source-context", StringComparison.OrdinalIgnoreCase))
            {
                RunWordLatexRedrawSourceContextStress();
            }
            else if (string.Equals(mode, "word-omml-boundary-digit-direct", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlBoundaryDigitDirect(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-anchor-recovery", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlAnchorRecovery(artifactRoot);
            }
            else if (string.Equals(mode, "word-inline-ole-visual-baseline", StringComparison.OrdinalIgnoreCase))
            {
                RunWordInlineOleVisualBaseline(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-inline-ole-visual-baseline-existing", StringComparison.OrdinalIgnoreCase))
            {
                RunExistingWordInlineOleVisualBaseline(artifactRoot);
            }
            else if (string.Equals(mode, "word-select-existing-inline-ole", StringComparison.OrdinalIgnoreCase))
            {
                SelectExistingWordInlineOle();
            }
            else if (string.Equals(mode, "word-inline-ole-typing-baseline", StringComparison.OrdinalIgnoreCase))
            {
                RunWordInlineOleTypingBaseline(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-inline-ole-initial-typing-baseline", StringComparison.OrdinalIgnoreCase))
            {
                RunWordInlineOleInitialTypingBaseline(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-inline-ole-font-style-existing", StringComparison.OrdinalIgnoreCase))
            {
                RunExistingWordInlineOleFontStyle(artifactRoot);
            }
            else if (string.Equals(mode, "word-latex-redraw", StringComparison.OrdinalIgnoreCase))
            {
                RunWordLatexRedraw(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-formula-to-latex", StringComparison.OrdinalIgnoreCase))
            {
                RunWordFormulaToLatex(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-formula-font", StringComparison.OrdinalIgnoreCase))
            {
                RunWordFormulaFontAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-tab-numbering", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlTabNumberingAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-numbered-omml-true-display-complex", StringComparison.OrdinalIgnoreCase))
            {
                RunWordNumberedOmmlTrueDisplayComplexAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-1x3-number-lifecycle", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlTableNumberLifecycleAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-1x3-number-stress", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlTableNumberStressAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-mathtype-reedit-regression", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeToOmmlReeditRegressionAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-unsaved-first-vsto-regression", StringComparison.OrdinalIgnoreCase))
            {
                RunWordUnsavedFirstNumberedOmmlVstoRegressionAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-installed-unsaved-first-omml", StringComparison.OrdinalIgnoreCase))
            {
                RunWordInstalledUnsavedFirstNumberedOmmlAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-empty-line-insertion-regression", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlEmptyLineInsertionRegressionAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-number-end-enter-regression", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlNumberEndEnterRegressionAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-deferred-finalization-cost", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlDeferredFinalizationCostAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-selection-change-idle", StringComparison.OrdinalIgnoreCase))
            {
                RunWordSelectionChangeIdleAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-converted-visualtex-numbering-scaffold", StringComparison.OrdinalIgnoreCase))
            {
                RunWordConvertedVisualTeXNumberingScaffoldAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mathtype-to-visualtex-numbered-core", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMathTypeToVisualTeXNumberedCoreAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-1x3-native-edit", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlTableNativeEditAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-1x3-complex", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlTableComplexAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-1x3-navigable-reference", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlTableNavigableReferenceAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-hash-to-1x3-migration", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlHashToTableMigrationAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-bad-eqarr-migration", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlDisplayMigrationAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-legacy-shape-migration", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlLegacyShapeMigrationAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mixed-legacy-numbering-migration", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMixedLegacyNumberingMigrationAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-numbered-omml-font-size", StringComparison.OrdinalIgnoreCase))
            {
                RunWordNumberedOmmlFontSizeAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-numbered-omml-empty-row", StringComparison.OrdinalIgnoreCase))
            {
                RunWordNumberedOmmlEmptyRowAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-bulk-import-dialog-options", StringComparison.OrdinalIgnoreCase))
            {
                RunWordBulkImportDialogOptionAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-bulk-import-mathtype", StringComparison.OrdinalIgnoreCase))
            {
                RunWordBulkImportMathTypeAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-bulk-import-latex-spacing", StringComparison.OrdinalIgnoreCase))
            {
                RunWordBulkImportLatexSpacing(artifactRoot);
            }
            else if (string.Equals(mode, "word-bulk-import-multiline", StringComparison.OrdinalIgnoreCase))
            {
                var objectMode = args
                    .FirstOrDefault(argument => argument.StartsWith(
                        "--object-mode=",
                        StringComparison.OrdinalIgnoreCase))
                    ?.Substring("--object-mode=".Length)
                    ?? "omml";
                RunWordBulkImportMultiline(client, artifactRoot, objectMode);
            }
            else if (string.Equals(mode, "office-package-command-probe", StringComparison.OrdinalIgnoreCase))
            {
                RunOfficePackageCommandProbe(client);
            }
            else if (string.Equals(mode, "office-ole-viewbox-probe", StringComparison.OrdinalIgnoreCase))
            {
                RunOfficeOleViewBoxProbe(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-bulk-import-ole-viewbox", StringComparison.OrdinalIgnoreCase))
            {
                RunWordBulkImportOleViewBox(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-editor-native-close", StringComparison.OrdinalIgnoreCase))
            {
                RunWordEditorNativeClose(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-numbered-omml-tab-scale", StringComparison.OrdinalIgnoreCase))
            {
                RunWordNumberedOmmlTabScaleAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-tab-format-roundtrip", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlTabFormatConversionAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-numbering-migration", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlNumberingMigrationAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-visualtex-omml-tab-numbering", StringComparison.OrdinalIgnoreCase))
            {
                RunWordVisualTeXOmmlTabAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-native-seq-lifecycle", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlNativeSequenceLifecycleAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-mixed-visualtex-sequence", StringComparison.OrdinalIgnoreCase))
            {
                RunWordMixedVisualTeXSequenceAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-native-number-toggle", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlNativeNumberToggleAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-sparse-numbered-omml-performance", StringComparison.OrdinalIgnoreCase))
            {
                RunWordSparseNumberedOmmlPerformanceAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-native-alias-recovery", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlNativeAliasRecoveryAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-native-f9", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlNativeF9Acceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-native-hash-production", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlNativeHashProductionAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-hash-number-production", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlHashNumberProductionAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-native-hash-complex-font", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlNativeHashComplexFontAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-native-hash-copy-paste", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlNativeHashCopyPasteAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-native-hash-scale", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlNativeHashScaleAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-true-display-tab-prototype", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlTrueDisplayTabPrototypeAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-hash-number-prototype", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlHashNumberPrototypeAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-true-display-shape-prototype", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlTrueDisplayShapePrototypeAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-visualtex-number-parenthesis", StringComparison.OrdinalIgnoreCase))
            {
                RunWordVisualTeXNumberParenthesisAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-visualtex-number-toggle", StringComparison.OrdinalIgnoreCase))
            {
                RunWordVisualTeXNumberToggleAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-installed-visualtex-number-toggle-close", StringComparison.OrdinalIgnoreCase))
            {
                RunWordInstalledVisualTeXNumberToggleCloseAcceptance(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-display-spacing", StringComparison.OrdinalIgnoreCase))
            {
                RunWordDisplaySpacing(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-equation-number-format", StringComparison.OrdinalIgnoreCase))
            {
                RunWordEquationNumberFormat(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-office2019-sequential-numbered-insertion", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOffice2019SequentialNumberedInsertion(artifactRoot);
            }
            else if (string.Equals(mode, "word-ole-numbering-migration", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOleNumberingMigration(artifactRoot);
            }
            else if (string.Equals(mode, "word-ole-copy-edit", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOleCopyEditAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-copy-paste-reedit", StringComparison.OrdinalIgnoreCase))
            {
                RunWordCopyPasteReeditAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-omml-copy-paste-reedit", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlCopyPasteReeditAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-numbered-formula-performance", StringComparison.OrdinalIgnoreCase))
            {
                RunWordNumberedFormulaPerformanceAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-numbered-middle-artifact-dump", StringComparison.OrdinalIgnoreCase))
            {
                RunWordNumberedMiddleArtifactDump(artifactRoot);
            }
            else if (string.Equals(mode, "word-numbered-omml-performance", StringComparison.OrdinalIgnoreCase))
            {
                RunWordNumberedOmmlPerformanceAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-numbered-structural-performance", StringComparison.OrdinalIgnoreCase))
            {
                RunWordNumberedStructuralPerformanceAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-number-format-performance", StringComparison.OrdinalIgnoreCase))
            {
                RunWordNumberFormatPerformanceAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-numbered-existing-performance", StringComparison.OrdinalIgnoreCase))
            {
                RunWordExistingNumberedPerformanceAcceptance(artifactRoot);
            }
            else if (string.Equals(mode, "word-bulk-import-performance", StringComparison.OrdinalIgnoreCase))
            {
                var objectMode = args
                    .FirstOrDefault(argument => argument.StartsWith(
                        "--object-mode=",
                        StringComparison.OrdinalIgnoreCase))
                    ?.Substring("--object-mode=".Length)
                    ?? "omml";
                RunWordBulkImportPerformance(client, artifactRoot, objectMode);
            }
            else if (string.Equals(mode, "targeted-powerpoint-46", StringComparison.OrdinalIgnoreCase))
            {
                RunTargetedPowerPoint46(client, artifactRoot);
            }
            else if (string.Equals(mode, "targeted-2364", StringComparison.OrdinalIgnoreCase))
            {
                RunTargetedWord2364(client, artifactRoot);
                RunTargetedPowerPoint46(client, artifactRoot);
            }
            else
            {
                RunWord(client, artifactRoot);
                RunPowerPoint(client, artifactRoot);
            }
            Console.WriteLine("VisualTeX real VSTO formula flow acceptance passed.");
            Console.WriteLine($"Artifacts: {artifactRoot}");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            Console.Error.WriteLine($"Acceptance artifacts retained: {artifactRoot}");
            return 1;
        }
    }

    private static void ProbeNativeEquationCrossReference()
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.Fields? fields = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            range = document.Range(0, 0);
            range.InsertCaption(
                Label: Word.WdCaptionLabelID.wdCaptionEquation,
                Title: " included",
                Position: Word.WdCaptionPosition.wdCaptionPositionBelow,
                ExcludeLabel: false);
            Release(range);
            range = null;

            var documentEnd = document.Content.End - 1;
            object secondStart = documentEnd;
            object secondEnd = documentEnd;
            range = document.Range(ref secondStart, ref secondEnd);
            range.InsertParagraphAfter();
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            range.InsertCaption(
                Label: Word.WdCaptionLabelID.wdCaptionEquation,
                Title: " excluded",
                Position: Word.WdCaptionPosition.wdCaptionPositionBelow,
                ExcludeLabel: true);
            Release(range);
            range = null;

            var label = application.CaptionLabels[Word.WdCaptionLabelID.wdCaptionEquation];
            var labelName = label.Name;
            Release(label);
            documentEnd = document.Content.End - 1;
            object thirdStart = documentEnd;
            object thirdEnd = documentEnd;
            range = document.Range(ref thirdStart, ref thirdEnd);
            range.InsertParagraphAfter();
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            range.Text = "(";
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            var manualField = document.Fields.Add(
                range,
                Word.WdFieldType.wdFieldEmpty,
                $"SEQ {labelName} \\* ARABIC",
                true);
            Release(manualField);
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            range.InsertAfter(") manual");
            Release(range);
            range = null;

            documentEnd = document.Content.End - 1;
            object fourthStart = documentEnd;
            object fourthEnd = documentEnd;
            range = document.Range(ref fourthStart, ref fourthEnd);
            range.InsertParagraphAfter();
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            range.InsertCaption(
                Label: Word.WdCaptionLabelID.wdCaptionEquation,
                Title: string.Empty,
                Position: Word.WdCaptionPosition.wdCaptionPositionBelow,
                ExcludeLabel: true);
            Word.Font? helperFont = null;
            Word.ParagraphFormat? helperParagraph = null;
            try
            {
                helperFont = range.Font;
                helperFont.Hidden = 0;
                helperFont.Size = 1f;
                helperFont.Color = Word.WdColor.wdColorWhite;
                helperParagraph = range.ParagraphFormat;
                helperParagraph.SpaceBefore = 0f;
                helperParagraph.SpaceAfter = 0f;
                helperParagraph.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                helperParagraph.LineSpacing = 1f;
            }
            finally
            {
                Release(helperParagraph);
                Release(helperFont);
            }
            Release(range);
            range = null;

            fields = document.Fields;
            Console.WriteLine($"Native equation caption field count: {fields.Count}");
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? field = null;
                Word.Range? code = null;
                Word.Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    result = field.Result;
                    Console.WriteLine(
                        $"Caption field {index}: code=[{code.Text?.Trim()}] result=[{result.Text?.Trim()}]");
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }

            var items = document.GetCrossReferenceItems(Word.WdCaptionLabelID.wdCaptionEquation);
            if (items is not Array array)
                throw new InvalidDataException("Word did not return an equation cross-reference array.");
            Console.WriteLine($"Native equation cross-reference item count: {array.Length}");
            for (var index = array.GetLowerBound(0); index <= array.GetUpperBound(0); index++)
                Console.WriteLine($"Native equation item {index}: [{array.GetValue(index)}]");
            if (array.Length < 4)
                throw new InvalidDataException("Word hidden native equation caption was not listed for cross-reference.");

            var referenceStart = document.Content.End - 1;
            object referenceStartObject = referenceStart;
            object referenceEndObject = referenceStart;
            range = document.Range(ref referenceStartObject, ref referenceEndObject);
            range.InsertParagraphAfter();
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            Word.Font? visibleReferenceFont = null;
            try
            {
                visibleReferenceFont = range.Font;
                visibleReferenceFont.Hidden = 0;
            }
            finally { Release(visibleReferenceFont); }
            range.InsertCrossReference(
                ReferenceType: Word.WdCaptionLabelID.wdCaptionEquation,
                ReferenceKind: Word.WdReferenceKind.wdEntireCaption,
                ReferenceItem: 4,
                InsertAsHyperlink: true,
                IncludePosition: false);
            fields = document.Fields;
            Console.WriteLine($"Field count after native reference insertion: {fields.Count}");
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? nativeField = null;
                Word.Range? nativeCode = null;
                Word.Range? nativeResult = null;
                Word.Font? nativeResultFont = null;
                try
                {
                    nativeField = fields[index];
                    nativeCode = nativeField.Code;
                    nativeResult = nativeField.Result;
                    nativeResultFont = nativeResult.Font;
                    Console.WriteLine(
                        $"Post-reference field {index}: type={nativeField.Type} " +
                        $"code=[{nativeCode.Text?.Trim()}] result=[{nativeResult.Text?.Trim()}] " +
                        $"hidden={nativeResultFont.Hidden}");
                }
                finally
                {
                    Release(nativeResultFont);
                    Release(nativeResult);
                    Release(nativeCode);
                    Release(nativeField);
                }
            }
        }
        finally
        {
            Release(fields);
            Release(range);
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

    private static void RunWordNativeCrossReference(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? firstShape = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            Console.WriteLine("[Native cross-reference 1/7] Starting Word...");
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);

            Console.WriteLine("[Native cross-reference 2/7] Creating equation (1)...");
            var existing = SnapshotSessionIds();
            addIn.OnInsertDisplay(new object());
            var firstSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var firstSession = client.GetSessionAsync(firstSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(client, firstSession, "block", "nativeOle", "a=b", numbered: true);
            var final = WaitForTerminal(client, firstSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "First numbered formula did not complete.");
            client.CloseEditorAsync(firstSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            var firstFormulaId = final.FormulaId
                ?? throw new InvalidDataException("First numbered formula has no formulaId.");
            WaitForWordInlineShapeCount(document, 1, TimeSpan.FromSeconds(15));

            Console.WriteLine("[Native cross-reference 3/7] Creating equation (2)...");
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            existing = SnapshotSessionIds();
            addIn.OnInsertDisplay(new object());
            var secondSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var secondSession = client.GetSessionAsync(secondSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(client, secondSession, "block", "nativeOle", "E=mc^2", numbered: true);
            final = WaitForTerminal(client, secondSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "Second numbered formula did not complete.");
            client.CloseEditorAsync(secondSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            var secondFormulaId = final.FormulaId
                ?? throw new InvalidDataException("Second numbered formula has no formulaId.");
            WaitForWordInlineShapeCount(document, 2, TimeSpan.FromSeconds(15));

            Console.WriteLine("[Native cross-reference 4/7] Checking Word's built-in Equation list...");
            var nativeItems = document.GetCrossReferenceItems(Word.WdCaptionLabelID.wdCaptionEquation) as Array;
            if (nativeItems is null || nativeItems.Length != 2)
                throw new InvalidDataException(
                    $"Word native Equation list should contain two VisualTeX formulas, actual count: {nativeItems?.Length ?? 0}.");
            var firstItem = Convert.ToString(nativeItems.GetValue(nativeItems.GetLowerBound(0)))?.Trim();
            var secondItem = Convert.ToString(nativeItems.GetValue(nativeItems.GetLowerBound(0) + 1))?.Trim();
            if (string.IsNullOrWhiteSpace(firstItem) || string.IsNullOrWhiteSpace(secondItem))
                throw new InvalidDataException("Word native Equation list returned an empty VisualTeX number.");
            if (string.Equals(firstItem, secondItem, StringComparison.Ordinal))
                throw new InvalidDataException("Two numbered VisualTeX formulas rendered the same native equation number.");

            Console.WriteLine($"[Native cross-reference 5/7] Inserting a native REF to equation ({secondItem})...");
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            application.Selection.TypeText("See ");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REFERENCE_TARGET_INDEX", "1");
            addIn.OnInsertEquationReference(new object());
            var nativeReferenceCode = WaitForWordNativeReferenceResult(
                document,
                expectedResult: secondItem!,
                expectedCode: null,
                TimeSpan.FromSeconds(15));
            if (!DocumentTextContains(document, $"({secondItem})"))
                throw new InvalidDataException($"Native Word reference did not render as ({secondItem}).");

            Console.WriteLine("[Native cross-reference 6/7] Deleting equation (1) and updating fields...");
            firstShape = document.InlineShapes[1];
            firstShape.Range.Select();
            addIn.OnDeleteSelected(new object());
            Release(firstShape);
            firstShape = null;
            WaitForWordInlineShapeCount(document, 1, TimeSpan.FromSeconds(15));
            addIn.OnUpdateEquationNumbers(new object());
            nativeItems = document.GetCrossReferenceItems(Word.WdCaptionLabelID.wdCaptionEquation) as Array;
            if (nativeItems is null || nativeItems.Length != 1)
                throw new InvalidDataException(
                    $"Word native Equation list should contain one item after deletion, actual count: {nativeItems?.Length ?? 0}.");
            var remainingItem = Convert.ToString(nativeItems.GetValue(nativeItems.GetLowerBound(0)))?.Trim();
            if (string.IsNullOrWhiteSpace(remainingItem))
                throw new InvalidDataException("Remaining Word native Equation item is empty after deletion.");
            WaitForWordNativeReferenceResult(
                document,
                expectedResult: remainingItem!,
                expectedCode: nativeReferenceCode,
                TimeSpan.FromSeconds(15));
            if (!DocumentTextContains(document, $"({remainingItem})"))
                throw new InvalidDataException($"Native Word reference did not update to ({remainingItem}).");
            if (WordBookmarkExists(document, $"VTEq_{Guid.Parse(firstFormulaId):N}"))
                throw new InvalidDataException("Deleted formula retained its visible number bookmark.");
            if (!WordBookmarkExists(document, $"VTEq_{Guid.Parse(secondFormulaId):N}"))
                throw new InvalidDataException("Remaining formula lost its visible number bookmark.");

            var path = Path.Combine(artifactRoot, "VisualTeX-Word-Native-CrossReference.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine($"[Native cross-reference 7/7] Saved {path}; native Equation list and REF update passed.");
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REFERENCE_TARGET_INDEX", null);
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            Release(firstShape);
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

    private static void RunExistingWordDoubleClickHitTest(
        VisualTeXSessionClient client)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Window? window = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? ommlRange = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        Word.Range? oleRange = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = Marshal.GetActiveObject("Word.Application") as Word.Application
                ?? throw new InvalidOperationException("No active Word instance is available.");
            document = application.ActiveDocument;
            window = application.ActiveWindow;
            maths = document.OMaths;
            shapes = document.InlineShapes;
            if (maths.Count < 1 || shapes.Count < 1)
                throw new InvalidDataException(
                    "The active Word document must contain at least one OMML and one inline formula.");

            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);
            var service = new VisualTeX.WordVsto.WordFormulaService(application);
            var callback = typeof(VisualTeX.WordVsto.ThisAddIn).GetMethod(
                "OnNativeWordDoubleClick",
                System.Reflection.BindingFlags.Instance
                    | System.Reflection.BindingFlags.NonPublic)
                ?? throw new MissingMethodException(
                    "The coordinate-aware Word double-click callback is missing.");

            math = maths[1];
            ommlRange = math.Range;
            window.GetPoint(
                out var ommlLeft,
                out var ommlTop,
                out var ommlWidth,
                out var ommlHeight,
                ommlRange);
            if (ommlWidth <= 0 || ommlHeight <= 0)
                throw new InvalidDataException("Word did not return the OMML screen rectangle.");
            var ommlX = ommlLeft + ommlWidth / 2;
            var ommlY = ommlTop + ommlHeight / 2;
            var ommlSelection = service.ReadVisualTeXOmmlAtScreenPoint(ommlX, ommlY)
                ?? throw new InvalidDataException(
                    "The OMML center was not resolved as a VisualTeX formula.");
            AssertEqual(
                FormulaOleContract.WordOmmlMode,
                ommlSelection.ObjectMode,
                "The OMML center resolved to the wrong object mode.");

            shape = shapes[1];
            oleRange = shape.Range;
            oleRange.Select();
            WinForms.Application.DoEvents();
            Thread.Sleep(200);
            var oleSelection = service.ReadSelection();
            AssertEqual(
                FormulaOleContract.NativeOleMode,
                oleSelection.ObjectMode,
                "The selected inline formula is not a native VisualTeX OLE object.");
            window.GetPoint(
                out var oleLeft,
                out var oleTop,
                out var oleWidth,
                out var oleHeight,
                oleRange);
            if (oleWidth <= 0 || oleHeight <= 0)
                throw new InvalidDataException("Word did not return the OLE screen rectangle.");
            var oleX = oleLeft + oleWidth / 2;
            var oleY = oleTop + oleHeight / 2;
            AssertTrue(
                service.IsFormulaAtScreenPoint(oleSelection, oleX, oleY),
                "The OLE center did not hit its real screen rectangle.");

            if (!GetWindowRect(new IntPtr(window.Hwnd), out var wordRectangle))
                throw new InvalidDataException("Word did not return its window rectangle.");
            var candidates = new[]
            {
                (X: wordRectangle.Left + 70, Y: ommlY),
                (X: wordRectangle.Right - 120, Y: ommlY),
                (X: ommlX, Y: wordRectangle.Bottom - 90),
                (X: oleX, Y: wordRectangle.Bottom - 90),
            };
            var blankFound = false;
            var blankX = 0;
            var blankY = 0;
            foreach (var candidate in candidates)
            {
                if (service.ReadVisualTeXOmmlAtScreenPoint(candidate.X, candidate.Y) is not null)
                    continue;
                if (service.IsFormulaAtScreenPoint(ommlSelection, candidate.X, candidate.Y)
                    || service.IsFormulaAtScreenPoint(oleSelection, candidate.X, candidate.Y))
                    continue;
                blankX = candidate.X;
                blankY = candidate.Y;
                blankFound = true;
                break;
            }
            if (!blankFound)
                throw new InvalidDataException(
                    "No unambiguous blank Word screen point was available for the hit-test acceptance.");

            void AssertBlankCallbackDoesNotOpen(
                Word.Range selectedFormulaRange,
                string description)
            {
                selectedFormulaRange.Select();
                WinForms.Application.DoEvents();
                Thread.Sleep(200);
                var existing = SnapshotSessionIds();
                callback.Invoke(addIn, new object[] { false, blankX, blankY });
                var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(2);
                while (DateTime.UtcNow < deadline)
                {
                    WinForms.Application.DoEvents();
                    Thread.Sleep(50);
                }
                var created = SnapshotSessionIds()
                    .Except(existing, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                if (created.Length == 0) return;
                foreach (var sessionId in created)
                {
                    try
                    {
                        client.CloseEditorAsync(sessionId, CancellationToken.None)
                            .GetAwaiter().GetResult();
                    }
                    catch { }
                }
                throw new InvalidDataException(
                    description + " incorrectly created an editor Session at blank coordinates: "
                    + string.Join(", ", created));
            }

            AssertBlankCallbackDoesNotOpen(
                ommlRange,
                "A stale OMML selection");
            AssertBlankCallbackDoesNotOpen(
                oleRange,
                "A stale OLE selection");

            void AssertFormulaCallbackOpens(
                Word.Range formulaRange,
                bool interceptedNativeOle,
                int screenX,
                int screenY,
                string expectedObjectMode,
                string description)
            {
                formulaRange.Select();
                WinForms.Application.DoEvents();
                Thread.Sleep(200);
                var existing = SnapshotSessionIds();
                callback.Invoke(
                    addIn,
                    new object[]
                    {
                        interceptedNativeOle,
                        screenX,
                        screenY,
                    });
                var sessionId = WaitForNewSession(
                    existing,
                    "word",
                    TimeSpan.FromSeconds(30));
                var session = WaitForUnchangedEditorReady(
                    client,
                    sessionId,
                    TimeSpan.FromSeconds(15));
                AssertEqual(
                    "edit",
                    session.Mode,
                    description + " did not create an edit Session.");
                AssertEqual(
                    expectedObjectMode,
                    session.ObjectMode,
                    description + " created the wrong object mode.");
                client.CloseEditorAsync(sessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
                var terminal = WaitForTerminal(
                    client,
                    sessionId,
                    TimeSpan.FromSeconds(30));
                AssertEqual(
                    "completed",
                    terminal.Status,
                    terminal.Error
                        ?? description + " unchanged editor did not close cleanly.");
                client.CloseEditorAsync(sessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            }

            AssertFormulaCallbackOpens(
                ommlRange,
                interceptedNativeOle: false,
                ommlX,
                ommlY,
                FormulaOleContract.WordOmmlMode,
                "The real OMML center callback");
            AssertFormulaCallbackOpens(
                oleRange,
                interceptedNativeOle: true,
                oleX,
                oleY,
                FormulaOleContract.NativeOleMode,
                "The real OLE center callback");

            Console.WriteLine(
                "Existing Word double-click hit-test passed: OMML/OLE centers opened "
                + "the correct editor, blank point "
                + $"{blankX},{blankY} missed both, and stale formula selections "
                + "created no editor Session.");
        }
        finally
        {
            if (addIn is not null)
            {
                try
                {
                    addIn.OnDisconnection(
                        ext_DisconnectMode.ext_dm_UserClosed,
                        ref custom);
                }
                catch { }
            }
            Release(oleRange);
            Release(shape);
            Release(shapes);
            Release(ommlRange);
            Release(math);
            Release(maths);
            Release(window);
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordOleRealDoubleClick(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.Range? range = null;
        Word.Window? window = null;
        COMAddIns? installedAddIns = null;
        COMAddIn? installedAddIn = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        var consoleWindow = GetConsoleWindow();
        string? converterSessionId = null;
        var hookTracePath = Path.Combine(artifactRoot, "word-ole-hook-trace.log");
        var previousOfficePreferencesPath = Environment.GetEnvironmentVariable(
            "VISUALTEX_OFFICE_PREFERENCES_PATH");
        var isolatedPreferencesPath = Path.Combine(
            artifactRoot,
            "office-preferences-visualtex-ole-mathtype-disabled.json");
        File.WriteAllText(
            isolatedPreferencesPath,
            "{\"powerpointDefaultFontSizePt\":20.0,\"mathtypeDoubleClickEditEnabled\":false}");
        Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", hookTracePath);
        Environment.SetEnvironmentVariable(
            "VISUALTEX_OFFICE_PREFERENCES_PATH",
            isolatedPreferencesPath);
        try
        {
            Console.WriteLine("[Word real OLE 1/8] Starting visible Word...");
            application = CreateWordApplication(visible: true);

            // This acceptance must exercise the current source assembly, not the
            // formally installed previous build. Disconnect the registered add-in
            // only for this Word process and reconnect it before shutdown.
            installedAddIns = application.COMAddIns;
            try
            {
                object addInIndex = "VisualTeX.WordVsto";
                installedAddIn = installedAddIns.Item(ref addInIndex);
                if (installedAddIn.Connect)
                    installedAddIn.Connect = false;
            }
            catch
            {
                Release(installedAddIn);
                installedAddIn = null;
            }

            document = application.Documents.Add();
            application.Selection.TypeText("Real OLE contour-integral double-click:");
            application.Selection.TypeParagraph();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);

            const string contourLatex = "\\oiint_{\\Sigma}F\\,\\mathrm{d}S11";
            Console.WriteLine("[Word real OLE 2/8] Rendering \\oiint through the production Office export path...");
            var sourceExportFixturePath = Path.Combine(artifactRoot, "oiint-source-export.json");
            OfficeExportDocument contourExport;
            if (File.Exists(sourceExportFixturePath))
            {
                contourExport = JsonSerializer.Deserialize<OfficeExportDocument>(
                        File.ReadAllText(sourceExportFixturePath, Encoding.UTF8),
                        new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                    ?? throw new InvalidDataException(
                        $"Source \\oiint export fixture is invalid: {sourceExportFixturePath}");
                contourExport.PngBase64 ??= CreatePngDataUrl(
                    contourLatex,
                    contourExport.Width,
                    contourExport.Height);
                Console.WriteLine(
                    $"  Using current-source export fixture: {sourceExportFixturePath}");
            }
            else
            {
                var converterLine = new FormulaLine
                {
                    Id = Guid.NewGuid().ToString("D"),
                    Latex = contourLatex,
                };
                var converterSession = client.CreateSessionAsync(
                        new CreateVstoSessionRequest
                        {
                            Mode = "create",
                            Host = "word",
                            Title = "OLE contour-integral acceptance renderer",
                            Lines = new List<FormulaLine> { converterLine },
                            ActiveLineId = converterLine.Id,
                            CodeFormat = "latex",
                            DisplayMode = "block",
                            ObjectMode = FormulaOleContract.NativeOleMode,
                            Numbered = false,
                            FontSizePt = 14d,
                            AutoCommitOnClose = false,
                        },
                        CancellationToken.None)
                    .GetAwaiter().GetResult();
                converterSessionId = converterSession.Id;
                client.OpenConverterAsync(converterSessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
                converterSession = client.WaitForCommitAsync(
                        converterSessionId,
                        TimeSpan.FromMinutes(3),
                        CancellationToken.None)
                    .GetAwaiter().GetResult();
                if (string.Equals(converterSession.Status, "failed", StringComparison.Ordinal))
                    throw new InvalidDataException(
                        converterSession.Error ?? "The production Office converter failed for \\oiint.");
                contourExport = converterSession.ExportResult
                    ?? throw new InvalidDataException(
                        "The production Office converter returned no \\oiint export.");
            }
            var contourSvg = contourExport.Svg;
            if (string.IsNullOrWhiteSpace(contourSvg)
                && !string.IsNullOrWhiteSpace(contourExport.SvgBase64))
            {
                var comma = contourExport.SvgBase64.IndexOf(',');
                if (comma >= 0)
                    contourSvg = Encoding.UTF8.GetString(
                        Convert.FromBase64String(contourExport.SvgBase64.Substring(comma + 1)));
            }
            if (string.IsNullOrWhiteSpace(contourSvg))
                throw new InvalidDataException("The production Office converter returned no readable \\oiint SVG.");
            AssertTrue(
                contourSvg!.IndexOf(
                    "data-visualtex-integral=\"oiint\"",
                    StringComparison.Ordinal) >= 0,
                "OLE \\oiint did not use the shared VisualTeX contour-integral vector glyph.");
            AssertTrue(
                contourSvg.IndexOf(">∯</text>", StringComparison.Ordinal) < 0,
                "OLE \\oiint fell back to the old small system-font character.");
            AssertTrue(
                contourExport.Height > 30f,
                $"OLE \\oiint did not retain display-integral height; exportHeight={contourExport.Height:0.###}.");
            AssertTrue(
                (contourExport.MathMl ?? string.Empty).IndexOf("\\oiint", StringComparison.Ordinal) < 0,
                "OLE \\oiint leaked into MathML as unresolved command text.");

            Console.WriteLine("[Word real OLE 3/8] Inserting the production vector as native OLE...");
            var existing = SnapshotSessionIds();
            addIn.OnInsertDisplay(new object());
            var createSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var createSession = client.GetSessionAsync(createSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            var createLineId = createSession.Lines.First().Id;
            client.PatchAsync(
                    createSessionId,
                    new Dictionary<string, object>
                    {
                        ["lines"] = new[]
                        {
                            new Dictionary<string, object>
                            {
                                ["id"] = createLineId,
                                ["latex"] = contourLatex,
                            },
                        },
                        ["activeLineId"] = createLineId,
                        ["codeFormat"] = "latex",
                        ["displayMode"] = "block",
                        ["objectMode"] = FormulaOleContract.NativeOleMode,
                        ["numbered"] = false,
                        ["fontSizePt"] = 14d,
                        ["exportWidth"] = contourExport.Width,
                        ["exportHeight"] = contourExport.Height,
                        ["exportResult"] = contourExport,
                        ["dirty"] = true,
                        ["status"] = "committing",
                    },
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            var created = WaitForTerminal(client, createSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", created.Status,
                created.Error ?? "Real OLE \\oiint fixture creation did not complete.");
            client.CloseEditorAsync(createSessionId, CancellationToken.None).GetAwaiter().GetResult();
            if (!string.IsNullOrWhiteSpace(converterSessionId))
            {
                client.CompleteAsync(converterSessionId!, CancellationToken.None)
                    .GetAwaiter().GetResult();
                converterSessionId = null;
            }
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            AssertEqual(1, document.InlineShapes.Count,
                "Real OLE \\oiint fixture should contain exactly one inline shape.");
            shape = document.InlineShapes[1];
            AssertEqual(FormulaOleContract.ProgId, shape.OLEFormat.ProgID,
                "Real OLE \\oiint fixture has the wrong ProgID.");
            AssertTrue(
                shape.Height > 20f,
                $"Word received a small OLE \\oiint preview; shapeHeight={shape.Height:0.###} pt.");

            var path = Path.Combine(artifactRoot, "VisualTeX-Word-Real-OLE-OIINT-DoubleClick.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            Release(shape);
            shape = null;
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: true);

            Console.WriteLine("[Word real OLE 4/8] Verifying save/reopen metadata and edit source...");
            AssertEqual(1, document.InlineShapes.Count,
                "Save/reopen changed the OLE \\oiint shape inventory.");
            shape = document.InlineShapes[1];
            shape.Range.Select();
            var reopenedSource = new VisualTeX.WordVsto.WordFormulaService(application).ReadSelection();
            AssertEqual(FormulaOleContract.NativeOleMode, reopenedSource.ObjectMode,
                "Save/reopen changed the OLE \\oiint object mode.");
            var reopenedLatex = reopenedSource.Metadata?.Latex ?? string.Empty;
            AssertTrue(
                reopenedLatex.IndexOf("\\oiint", StringComparison.Ordinal) >= 0,
                $"Save/reopen lost the canonical \\oiint command; actual='{reopenedLatex}'.");
            AssertTrue(
                reopenedLatex.IndexOf('∯') < 0,
                "Save/reopen serialized OLE \\oiint as raw Unicode instead of canonical LaTeX.");

            Console.WriteLine("[Word real OLE 5/8] Resolving the reopened formula screen rectangle...");
            AssertEqual(FormulaOleContract.ProgId, shape.OLEFormat.ProgID,
                "Reopened OLE \\oiint fixture has the wrong ProgID.");
            range = shape.Range;
            range.Select();
            application.ActiveWindow.Activate();
            WinForms.Application.DoEvents();
            Thread.Sleep(500);
            var addInType = typeof(VisualTeX.WordVsto.ThisAddIn);
            var activeField = addInType.GetField(
                "_nativeOleTargetActive",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            var leftField = addInType.GetField(
                "_nativeOleTargetLeft",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            var topField = addInType.GetField(
                "_nativeOleTargetTop",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            var rightField = addInType.GetField(
                "_nativeOleTargetRight",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            var bottomField = addInType.GetField(
                "_nativeOleTargetBottom",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            Console.WriteLine(
                $"  OLE hook cache active={activeField?.GetValue(addIn)}; "
                + $"rect={leftField?.GetValue(addIn)},{topField?.GetValue(addIn)},"
                + $"{rightField?.GetValue(addIn)},{bottomField?.GetValue(addIn)}.");
            window = application.ActiveWindow;
            window.GetPoint(
                out var left,
                out var top,
                out var width,
                out var height,
                range);
            if (width <= 0 || height <= 0)
                throw new InvalidDataException("Word did not return a visible OLE formula rectangle.");

            Console.WriteLine("[Word real OLE 6/9] Rejecting a blank-area double-click after selecting the formula...");
            if (consoleWindow != IntPtr.Zero) ShowWindow(consoleWindow, 0);
            var wordWindowHandle = new IntPtr(window.Hwnd);
            const uint noMoveNoSizeShow = 0x0001 | 0x0002 | 0x0040;
            SetWindowPos(wordWindowHandle, new IntPtr(-1), 0, 0, 0, 0, noMoveNoSizeShow);
            var foregroundSet = SetForegroundWindow(wordWindowHandle);
            if (!GetWindowRect(wordWindowHandle, out var wordWindowRectangle))
                throw new InvalidDataException("Word did not return its window rectangle.");
            var titleX = wordWindowRectangle.Left
                + Math.Max(40, (wordWindowRectangle.Right - wordWindowRectangle.Left) / 2);
            var titleY = wordWindowRectangle.Top + 18;
            SetCursorPos(titleX, titleY);
            mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
            WinForms.Application.DoEvents();
            Thread.Sleep(600);
            Console.WriteLine($"  Word foreground request accepted={foregroundSet}.");

            var x = left + width / 2;
            var y = top + height / 2;
            SetCursorPos(x, y);
            mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
            WinForms.Application.DoEvents();
            Thread.Sleep(250);

            var leftBlankX = wordWindowRectangle.Left + 80;
            var rightBlankX = wordWindowRectangle.Right - 120;
            var blankX = Math.Abs(leftBlankX - x) >= Math.Abs(rightBlankX - x)
                ? leftBlankX
                : rightBlankX;
            var blankY = y;
            if (VisualTeX.WindowsOffice.VstoShared.WordDoubleClickRouting
                    .ScreenPointHitsFormulaRectangle(
                    blankX,
                    blankY,
                    left,
                    top,
                    width,
                    height))
                throw new InvalidDataException("The OLE blank-area probe still overlaps the formula rectangle.");

            var blankSessionsBefore = SnapshotSessionIds();
            SetCursorPos(blankX, blankY);
            Thread.Sleep(120);
            for (var click = 0; click < 2; click++)
            {
                mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(90);
            }
            var blankDeadline = DateTime.UtcNow + TimeSpan.FromSeconds(2);
            while (DateTime.UtcNow < blankDeadline)
            {
                WinForms.Application.DoEvents();
                Thread.Sleep(50);
            }
            var unexpectedBlankSessions = SnapshotSessionIds()
                .Except(blankSessionsBefore, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (unexpectedBlankSessions.Length > 0)
                throw new InvalidDataException(
                    "Double-clicking blank Word space after selecting an OLE formula "
                    + "incorrectly created an editor Session: "
                    + string.Join(", ", unexpectedBlankSessions));
            Console.WriteLine(
                $"  Blank-area double-click rejected at {blankX},{blankY}; no Session created.");

            Console.WriteLine("[Word real OLE 7/9] Sending a real formula double-click...");
            existing = SnapshotSessionIds();
            SetCursorPos(x, y);
            Thread.Sleep(150);
            for (var click = 0; click < 2; click++)
            {
                mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(90);
            }
            var editSessionId = WaitForNewSession(
                existing,
                "word",
                TimeSpan.FromSeconds(30));
            var editSession = client.GetSessionAsync(editSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("edit", editSession.Mode,
                "Real OLE double-click did not create an edit Session.");
            AssertEqual(FormulaOleContract.NativeOleMode, editSession.ObjectMode,
                "Real OLE double-click created the wrong object mode.");
            var editLatex = string.Join("\n", editSession.Lines.Select(line => line.Latex));
            AssertTrue(
                editLatex.IndexOf("\\oiint", StringComparison.Ordinal) >= 0,
                $"Real OLE double-click reopened unresolved/noncanonical source: '{editLatex}'.");
            AssertTrue(
                editLatex.IndexOf('∯') < 0,
                "Real OLE double-click returned raw Unicode ∯ instead of \\oiint.");
            Console.WriteLine(
                $"  Real mouse OLE Session={editSessionId}; rectangle={left},{top},{width},{height}.");

            Console.WriteLine("[Word real OLE 8/9] Closing the unchanged editor...");
            var ready = WaitForUnchangedEditorReady(
                client,
                editSessionId,
                TimeSpan.FromSeconds(10));
            AssertEqual(false, ready.Dirty,
                "Real OLE double-click editor became dirty before input.");
            client.CloseEditorAsync(editSessionId, CancellationToken.None).GetAwaiter().GetResult();
            var closed = WaitForTerminal(client, editSessionId, TimeSpan.FromSeconds(30));
            AssertEqual("completed", closed.Status,
                closed.Error ?? "Real OLE double-click editor did not close cleanly.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            document.Save();
            Console.WriteLine(
                $"[Word real OLE 9/9] Saved {path}; production \\oiint vector, "
                + "blank-space rejection, save/reopen metadata, and real mouse edit checks passed.");
        }
        finally
        {
            if (consoleWindow != IntPtr.Zero) ShowWindow(consoleWindow, 5);
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", null);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_OFFICE_PREFERENCES_PATH",
                previousOfficePreferencesPath);
            if (!string.IsNullOrWhiteSpace(converterSessionId))
            {
                try
                {
                    client.CompleteAsync(converterSessionId!, CancellationToken.None)
                        .GetAwaiter().GetResult();
                }
                catch { }
            }
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            if (installedAddIn is not null)
            {
                try { installedAddIn.Connect = true; } catch { }
            }
            Release(installedAddIn);
            Release(installedAddIns);
            Release(window);
            Release(range);
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

    private static void RunWordOmmlDoubleClickEventProbe(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        var path = Path.Combine(artifactRoot, "VisualTeX-Word-OMML-DoubleClick-Fixtures.docx");
        if (!File.Exists(path))
            throw new FileNotFoundException("The OMML double-click fixture document is missing.", path);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? range = null;
        Word.Selection? selection = null;
        COMAddIns? installedAddIns = null;
        COMAddIn? installedAddIn = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = CreateWordApplication(visible: true);
            installedAddIns = application.COMAddIns;
            try
            {
                object addInIndex = "VisualTeX.WordVsto";
                installedAddIn = installedAddIns.Item(ref addInIndex);
                if (installedAddIn.Connect) installedAddIn.Connect = false;
            }
            catch
            {
                Release(installedAddIn);
                installedAddIn = null;
            }

            document = application.Documents.Open(path, ReadOnly: false, Visible: true);
            document.Activate();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);
            maths = document.OMaths;
            math = maths[1];
            range = math.Range;
            range.Select();
            selection = application.Selection;

            var sessionsBefore = SnapshotSessionIds();
            var handler = typeof(VisualTeX.WordVsto.ThisAddIn).GetMethod(
                "OnWindowBeforeDoubleClick",
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.NonPublic)
                ?? throw new MissingMethodException("Word OMML double-click callback is missing.");
            var arguments = new object[] { selection, false };
            handler.Invoke(addIn, arguments);
            var sessionId = WaitForNewSession(
                sessionsBefore,
                "word",
                TimeSpan.FromSeconds(30));
            var session = WaitForUnchangedEditorReady(
                client,
                sessionId,
                TimeSpan.FromSeconds(15));
            AssertEqual("edit", session.Mode,
                "Direct OMML double-click callback did not create an edit Session.");
            AssertEqual(FormulaOleContract.WordOmmlMode, session.ObjectMode,
                "Direct OMML double-click callback changed the object mode.");
            if (!string.Equals(session.Lines.FirstOrDefault()?.Latex, "x+y", StringComparison.Ordinal))
                throw new InvalidDataException(
                    $"Direct OMML double-click callback opened the wrong source: '{session.Lines.FirstOrDefault()?.Latex}'.");
            client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            Console.WriteLine(
                "Word OMML direct double-click callback probe passed; handler logic opened the correct OMML formula.");
        }
        finally
        {
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            if (installedAddIn is not null)
            {
                try { installedAddIn.Connect = true; } catch { }
            }
            Release(selection);
            Release(range);
            Release(math);
            Release(maths);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            Release(installedAddIn);
            Release(installedAddIns);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordOmmlDoubleClickFixtures(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? oleShape = null;
        COMAddIns? installedAddIns = null;
        COMAddIn? installedAddIn = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        var consoleWindow = GetConsoleWindow();
        try
        {
            Console.WriteLine("[OMML fixtures 1/8] Starting Word...");
            application = CreateWordApplication(visible: false);
            installedAddIns = application.COMAddIns;
            try
            {
                object addInIndex = "VisualTeX.WordVsto";
                installedAddIn = installedAddIns.Item(ref addInIndex);
                if (installedAddIn.Connect)
                    installedAddIn.Connect = false;
            }
            catch
            {
                Release(installedAddIn);
                installedAddIn = null;
            }
            document = application.Documents.Add();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);

            void MoveToNewLabeledParagraph(string label, bool first = false)
            {
                application.Selection.EndKey(Word.WdUnits.wdStory);
                if (!first) application.Selection.TypeParagraph();
                application.Selection.TypeText(label);
            }

            void CreateOmml(
                Action<object> command,
                string displayMode,
                string latex,
                string mathMl,
                bool numbered)
            {
                var existing = SnapshotSessionIds();
                command(new object());
                var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
                var session = client.GetSessionAsync(sessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
                AssertEqual(FormulaOleContract.WordOmmlMode, session.ObjectMode,
                    "OMML fixture command did not request Word OMML.");
                Commit(
                    client,
                    session,
                    displayMode,
                    FormulaOleContract.WordOmmlMode,
                    latex,
                    numbered: numbered,
                    mathMl: mathMl);
                var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
                AssertEqual("completed", final.Status,
                    final.Error ?? $"OMML fixture '{latex}' did not complete.");
                client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
                WaitForAddInIdle(addIn!, TimeSpan.FromSeconds(10));
                Console.WriteLine($"  OMML inventory after '{latex}': {document.OMaths.Count}.");
            }

            void ConvertSelected(
                Action<object> command,
                string displayMode,
                string objectMode,
                string latex,
                bool numbered,
                string? mathMl = null)
            {
                var existing = SnapshotSessionIds();
                command(new object());
                var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
                var session = client.GetSessionAsync(sessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
                AssertEqual(objectMode, session.ObjectMode,
                    $"Conversion to {objectMode} requested the wrong object mode.");
                Commit(
                    client,
                    session,
                    displayMode,
                    objectMode,
                    latex,
                    numbered: numbered,
                    dirty: false,
                    mathMl: mathMl);
                var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
                AssertEqual("completed", final.Status,
                    final.Error ?? $"Conversion to {objectMode} did not complete.");
                client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
                WaitForAddInIdle(addIn!, TimeSpan.FromSeconds(10));
            }

            void AssertNumberedTableCellNormalized(int tableIndex, string stage)
            {
                Word.Table? table = null;
                Word.Cell? centerCell = null;
                Word.Cell? numberCell = null;
                Word.Range? centerRange = null;
                Word.Range? numberRange = null;
                Word.Bookmarks? numberBookmarks = null;
                Word.Bookmark? labelBookmark = null;
                Word.Range? labelRange = null;
                Word.Font? labelFont = null;
                Word.Paragraphs? centerParagraphs = null;
                Word.Paragraphs? numberParagraphs = null;
                try
                {
                    table = document.Tables[tableIndex];
                    centerCell = table.Cell(1, 2);
                    numberCell = table.Cell(1, 3);
                    centerRange = centerCell.Range;
                    numberRange = numberCell.Range;
                    centerParagraphs = centerRange.Paragraphs;
                    numberParagraphs = numberRange.Paragraphs;
                    AssertEqual(1, centerParagraphs.Count,
                        $"{stage}: formula cell contains extra paragraphs.");
                    AssertEqual(1, numberParagraphs.Count,
                        $"{stage}: number cell contains extra paragraphs.");
                    AssertEqual(
                        Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter,
                        centerCell.VerticalAlignment,
                        $"{stage}: formula cell is not vertically centered.");
                    AssertEqual(
                        Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter,
                        numberCell.VerticalAlignment,
                        $"{stage}: number cell is not vertically centered.");
                    AssertEqual(
                        Word.WdParagraphAlignment.wdAlignParagraphCenter,
                        centerRange.ParagraphFormat.Alignment,
                        $"{stage}: formula paragraph is not horizontally centered.");
                    AssertEqual(
                        Word.WdParagraphAlignment.wdAlignParagraphRight,
                        numberRange.ParagraphFormat.Alignment,
                        $"{stage}: number paragraph is not right aligned.");
                    AssertNear(0f, numberRange.Font.Position, 0.1f,
                        $"{stage}: number has a manual baseline offset.");

                    numberBookmarks = numberRange.Bookmarks;
                    for (var bookmarkIndex = 1;
                         bookmarkIndex <= numberBookmarks.Count;
                         bookmarkIndex++)
                    {
                        Word.Bookmark? candidate = null;
                        try
                        {
                            candidate = numberBookmarks[bookmarkIndex];
                            var name = candidate.Name ?? string.Empty;
                            if (!name.StartsWith("VTEq_", StringComparison.Ordinal)
                                || name.StartsWith("VTEqCap_", StringComparison.Ordinal)
                                || name.StartsWith("VTEqNum_", StringComparison.Ordinal))
                                continue;
                            labelBookmark = candidate;
                            candidate = null;
                            break;
                        }
                        finally { Release(candidate); }
                    }
                    if (labelBookmark is null)
                        throw new InvalidDataException(
                            $"{stage}: visible equation-number bookmark is missing.");
                    labelRange = labelBookmark.Range;
                    labelFont = labelRange.Font;
                    var labelText = labelRange.Text ?? string.Empty;
                    AssertTrue(
                        labelText.StartsWith("(", StringComparison.Ordinal)
                        && labelText.EndsWith(")", StringComparison.Ordinal),
                        $"{stage}: equation-number bookmark does not contain both parentheses.");
                    AssertEqual("Cambria Math", labelFont.Name,
                        $"{stage}: parentheses and number do not share the compatibility font.");
                    AssertNear(0f, labelFont.Position, 0.1f,
                        $"{stage}: parentheses and number have different baselines.");
                    AssertEqual(0, labelFont.Hidden,
                        $"{stage}: visible number inherited hidden caption formatting.");
                    AssertEqual(0, labelFont.Bold,
                        $"{stage}: visible number inherited bold formatting.");
                    AssertEqual(0, labelFont.Italic,
                        $"{stage}: visible number inherited italic formatting.");

                    var cellXml = XDocument.Parse(centerRange.WordOpenXML);
                    XNamespace word =
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    AssertEqual(0, cellXml.Descendants(word + "br").Count(),
                        $"{stage}: formula cell retained hidden manual line breaks.");
                }
                finally
                {
                    Release(labelFont);
                    Release(labelRange);
                    Release(labelBookmark);
                    Release(numberBookmarks);
                    Release(numberParagraphs);
                    Release(centerParagraphs);
                    Release(numberRange);
                    Release(centerRange);
                    Release(numberCell);
                    Release(centerCell);
                    Release(table);
                }
            }

            Console.WriteLine("[OMML fixtures 2/7] Creating inline OMML...");
            MoveToNewLabeledParagraph("1. Inline OMML: ", first: true);
            CreateOmml(
                addIn.OnInsertInlineOmml,
                "inline",
                "x+y",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mi>y</mi></math>",
                numbered: false);

            // Regression: OMML -> OLE used to leave the caret carrying the
            // formula's raised baseline. Exercise the actual ribbon conversion,
            // type after it, then return the fixture to OMML for later checks.
            document.OMaths[1].Range.Select();
            ConvertSelected(
                addIn.OnConvertSelected,
                "inline",
                FormulaOleContract.NativeOleMode,
                "x+y",
                numbered: false);
            AssertNear(0f, application.Selection.Font.Position, 0.1f,
                "Caret after inline OMML-to-OLE conversion inherited a baseline offset.");
            var inlineSuffixStart = application.Selection.Start;
            application.Selection.TypeText(" baseline-ok");
            object inlineSuffixRangeStart = inlineSuffixStart;
            object inlineSuffixRangeEnd = application.Selection.Start;
            Word.Range? inlineSuffixRange = document.Range(
                ref inlineSuffixRangeStart,
                ref inlineSuffixRangeEnd);
            try
            {
                AssertNear(0f, inlineSuffixRange.Font.Position, 0.1f,
                    "Text after inline OMML-to-OLE conversion inherited a baseline offset.");
            }
            finally { Release(inlineSuffixRange); }
            document.InlineShapes[1].Range.Select();
            ConvertSelected(
                addIn.OnConvertSelectedToOmml,
                "inline",
                FormulaOleContract.WordOmmlMode,
                "x+y",
                numbered: false,
                mathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mi>y</mi></math>");
            AssertEqual(1, document.OMaths.Count,
                "Inline OMML-to-OLE-to-OMML round-trip lost or duplicated the equation.");

            Console.WriteLine("[OMML fixtures 3/8] Reopening inline OMML and appending digit 1...");
            document.OMaths[1].Range.Select();
            var boundaryExisting = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            var boundaryEditSessionId = WaitForNewSession(
                boundaryExisting,
                "word",
                TimeSpan.FromSeconds(30));
            var boundaryEditSession = WaitForUnchangedEditorReady(
                client,
                boundaryEditSessionId,
                TimeSpan.FromSeconds(15));
            var reopenedLatex = string.Join(
                "\n",
                boundaryEditSession.Lines.Select(line => line.Latex));
            AssertTrue(
                reopenedLatex.IndexOf('\u200B') < 0
                && reopenedLatex.IndexOf('\u200C') < 0
                && reopenedLatex.IndexOf('\u2060') < 0,
                "Inline OMML edit session imported a VisualTeX typing anchor into LaTeX: "
                + reopenedLatex);
            AssertEqual("x+y", reopenedLatex,
                "Inline OMML edit session duplicated or changed the original formula.");
            Commit(
                client,
                boundaryEditSession,
                "inline",
                FormulaOleContract.WordOmmlMode,
                "x+y1",
                numbered: false,
                mathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                    + "<mi>x</mi><mo>+</mo><mi>y</mi>"
                    + "<mrow data-mjx-texclass=\"ORD\"><mo>&#x200C;</mo></mrow>"
                    + "<mn>1</mn></math>");
            var boundaryFinal = WaitForTerminal(
                client,
                boundaryEditSessionId,
                TimeSpan.FromSeconds(45));
            AssertEqual("completed", boundaryFinal.Status,
                boundaryFinal.Error ?? "Appending digit 1 to inline OMML did not complete.");
            client.CloseEditorAsync(boundaryEditSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            WaitForAddInIdle(addIn!, TimeSpan.FromSeconds(10));
            AssertEqual(1, document.OMaths.Count,
                "Appending digit 1 duplicated the inline OMML equation.");
            Word.OMath? updatedInlineMath = null;
            Word.Range? updatedInlineRange = null;
            try
            {
                updatedInlineMath = document.OMaths[1];
                updatedInlineRange = updatedInlineMath.Range;
                var updatedText = updatedInlineRange.Text ?? string.Empty;
                AssertTrue(updatedText.IndexOf('1') >= 0,
                    "The appended digit 1 was not written to Word OMML.");
                AssertTrue(
                    updatedText.IndexOf('\u200B') < 0
                    && updatedText.IndexOf('\u200C') < 0
                    && updatedText.IndexOf('\u2060') < 0,
                    "The updated Word OMML still contains a VisualTeX typing anchor.");
                AssertTrue(
                    updatedInlineRange.WordOpenXML.IndexOf("200C", StringComparison.OrdinalIgnoreCase) < 0,
                    "The updated Word OMML XML still contains U+200C.");
            }
            finally
            {
                Release(updatedInlineRange);
                Release(updatedInlineMath);
            }

            Console.WriteLine("[OMML fixtures 4/8] Creating unnumbered display OMML...");
            MoveToNewLabeledParagraph("2. Display OMML (unnumbered):");
            CreateOmml(
                addIn.OnInsertDisplayOmml,
                "block",
                "\\frac{a+b}{c+d}",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mfrac><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mrow><mi>c</mi><mo>+</mo><mi>d</mi></mrow></mfrac></math>",
                numbered: false);

            Console.WriteLine("[OMML fixtures 4/7] Creating numbered display OMML...");
            MoveToNewLabeledParagraph("3. Display OMML (numbered):");
            CreateOmml(
                addIn.OnInsertDisplayOmml,
                "block",
                "\\sum_{n=1}^{\\infty}\\frac{1}{n^2}",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><munderover><mo>∑</mo><mrow><mi>n</mi><mo>=</mo><mn>1</mn></mrow><mi>∞</mi></munderover><mfrac><mn>1</mn><msup><mi>n</mi><mn>2</mn></msup></mfrac></mrow></math>",
                numbered: true);

            var naryCases = new (string Latex, string MathMl)[]
            {
                ("\\sum_b^z c", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><munderover><mo>∑</mo><mi>b</mi><mi>z</mi></munderover><mi>c</mi></math>"),
                ("\\sum_{b}^{z} c", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><munderover><mo>∑</mo><mrow><mi>b</mi></mrow><mrow><mi>z</mi></mrow></munderover><mi>c</mi></math>"),
                ("\\oint_l^u x\\,dy", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msubsup><mo>∮</mo><mi>l</mi><mi>u</mi></msubsup><mi>x</mi><mstyle><mspace width=\"0.167em\"/></mstyle><mi>d</mi><mi>y</mi></math>"),
                ("\\oint_l x\\,dy", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msub><mo>∮</mo><mi>l</mi></msub><mi>x</mi><mstyle><mspace width=\"0.167em\"/></mstyle><mi>d</mi><mi>y</mi></math>"),
                ("\\int_0^1 x^2\\,dx", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup><msup><mi>x</mi><mn>2</mn></msup><mstyle><mspace width=\"0.167em\"/></mstyle><mi>d</mi><mi>x</mi></math>"),
                ("\\prod_{i=1}^{n} a_i", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><munderover><mo>∏</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mrow><mi>n</mi></mrow></munderover><msub><mi>a</mi><mi>i</mi></msub></math>"),
            };
            foreach (var (latex, mathMl) in naryCases)
            {
                MoveToNewLabeledParagraph($"N-ary regression: {latex}");
                CreateOmml(
                    addIn.OnInsertDisplayOmml,
                    "block",
                    latex,
                    mathMl,
                    numbered: false);
            }

            Console.WriteLine("[OMML fixtures 5/7] Creating OLE then converting it to OMML...");
            MoveToNewLabeledParagraph("4. Numbered display OLE converted to OMML: ");
            var existing = SnapshotSessionIds();
            addIn.OnInsertDisplay(new object());
            var oleSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var oleSession = client.GetSessionAsync(oleSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(
                client,
                oleSession,
                "block",
                FormulaOleContract.NativeOleMode,
                "\\int_0^1 x^2\\,dx",
                numbered: true);
            var finalOle = WaitForTerminal(client, oleSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", finalOle.Status,
                finalOle.Error ?? "OLE fixture creation did not complete.");
            client.CloseEditorAsync(oleSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            oleShape = document.InlineShapes[document.InlineShapes.Count];
            AssertNumberedTableCellNormalized(2, "new numbered OLE");
            oleShape.Range.Select();
            existing = SnapshotSessionIds();
            addIn.OnConvertSelectedToOmml(new object());
            var conversionSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var conversionSession = client.GetSessionAsync(conversionSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual(FormulaOleContract.WordOmmlMode, conversionSession.ObjectMode,
                "OLE to OMML fixture conversion requested the wrong object mode.");
            Commit(
                client,
                conversionSession,
                "block",
                FormulaOleContract.WordOmmlMode,
                "\\int_0^1 x^2\\,dx",
                numbered: true,
                dirty: false,
                mathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup><msup><mi>x</mi><mn>2</mn></msup><mi>d</mi><mi>x</mi></math>");
            var converted = WaitForTerminal(client, conversionSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", converted.Status,
                converted.Error ?? "OLE to OMML fixture conversion did not complete.");
            client.CloseEditorAsync(conversionSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            Console.WriteLine($"  OMML inventory after numbered OLE conversion: {document.OMaths.Count}.");
            AssertNumberedTableCellNormalized(2, "initial OLE-to-OMML conversion");
            Release(oleShape);
            oleShape = null;

            // Regression: a numbered OMML -> OLE -> OMML round-trip used to
            // create a nested table and several empty paragraphs. Inserting the
            // next display formula could then remove the previous equation.
            for (var round = 1; round <= 3; round++)
            {
                document.OMaths[document.OMaths.Count].Range.Select();
                ConvertSelected(
                    addIn.OnConvertSelected,
                    "block",
                    FormulaOleContract.NativeOleMode,
                    "\\int_0^1 x^2\\,dx",
                    numbered: true);
                AssertEqual(2, document.Tables.Count,
                    $"Round {round} OMML-to-OLE created an extra equation table.");
                AssertEqual(9, document.OMaths.Count,
                    $"Round {round} OMML-to-OLE did not replace exactly one OMML equation.");
                AssertNumberedTableCellNormalized(
                    2,
                    $"round {round} OMML-to-OLE");

                document.InlineShapes[document.InlineShapes.Count].Range.Select();
                ConvertSelected(
                    addIn.OnConvertSelectedToOmml,
                    "block",
                    FormulaOleContract.WordOmmlMode,
                    "\\int_0^1 x^2\\,dx",
                    numbered: true,
                    mathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msubsup><mo>\u222b</mo><mn>0</mn><mn>1</mn></msubsup><msup><mi>x</mi><mn>2</mn></msup><mi>d</mi><mi>x</mi></math>");
                AssertEqual(2, document.Tables.Count,
                    $"Round {round} OLE-to-OMML created a nested/extra equation table.");
                AssertEqual(10, document.OMaths.Count,
                    $"Round {round} lost or duplicated the original display equation.");
                AssertNumberedTableCellNormalized(
                    2,
                    $"round {round} OLE-to-OMML");
            }

            MoveToNewLabeledParagraph("5. Display OMML inserted after round-trip:");
            CreateOmml(
                addIn.OnInsertDisplayOmml,
                "block",
                "a+b=c",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>a</mi><mo>+</mo><mi>b</mi><mo>=</mo><mi>c</mi></mrow></math>",
                numbered: true);
            AssertEqual(11, document.OMaths.Count,
                "Inserting the next display equation removed the round-tripped equation.");

            Console.WriteLine("[OMML fixtures 6/8] Validating formula inventory...");
            const int expectedEquationCount = 11;
            if (document.OMaths.Count != expectedEquationCount)
                throw new InvalidDataException(
                    $"OMML double-click fixture document should contain {expectedEquationCount} equations, actual: {document.OMaths.Count}.");
            XNamespace mathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            XNamespace wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            for (var equationIndex = 1; equationIndex <= document.OMaths.Count; equationIndex++)
            {
                Word.OMath? equation = null;
                Word.Range? equationRange = null;
                try
                {
                    equation = document.OMaths[equationIndex];
                    equationRange = equation.Range;
                    var xml = XDocument.Parse(equationRange.WordOpenXML);
                    if (xml.Descendants(wordNamespace + "color")
                            .Any(element => string.Equals(
                                (string?)element.Attribute(wordNamespace + "val"),
                                "FFFFFF",
                                StringComparison.OrdinalIgnoreCase))
                        || xml.Descendants(wordNamespace + "sz")
                            .Any(element => (string?)element.Attribute(wordNamespace + "val") == "2"))
                        throw new InvalidDataException(
                            $"Equation {equationIndex} inherited the hidden-caption white 1pt style.");
                    foreach (var nary in xml.Descendants(mathNamespace + "nary"))
                    {
                        var operand = nary.Element(mathNamespace + "e");
                        if (operand is null || !operand.Elements().Any())
                            throw new InvalidDataException(
                                $"Equation {equationIndex} contains an empty n-ary operand.");
                    }
                    if (equationIndex > 1 && equation.Type != Word.WdOMathType.wdOMathDisplay)
                        throw new InvalidDataException(
                            $"Display equation {equationIndex} degraded to inline OMath.");
                }
                finally
                {
                    Release(equationRange);
                    Release(equation);
                }
            }

            if (document.Tables.Count != 3 || document.Fields.Count != 6)
                throw new InvalidDataException(
                    $"Expected three numbered equation tables with SEQ+REF fields; "
                    + $"tables={document.Tables.Count}, fields={document.Fields.Count}.");
            for (var tableIndex = 1; tableIndex <= 3; tableIndex++)
            {
                Word.Table? table = null;
                Word.Cell? numberCell = null;
                Word.Range? numberCellRange = null;
                Word.Paragraphs? numberCellParagraphs = null;
                try
                {
                    table = document.Tables[tableIndex];
                    numberCell = table.Cell(1, 3);
                    numberCellRange = numberCell.Range;
                    numberCellParagraphs = numberCellRange.Paragraphs;
                    AssertEqual(
                        1,
                        numberCellParagraphs.Count,
                        $"Numbered equation table {tableIndex} contains an extra empty paragraph that shifts the number down.");
                    AssertNear(0f, numberCellRange.Font.Position, 0.1f,
                        $"Numbered equation table {tableIndex} applies a manual baseline shift instead of cell centering.");
                    var visibleNumber = (numberCellRange.Text ?? string.Empty)
                        .Trim('\r', '\a', ' ');
                    AssertEqual(
                        $"({tableIndex})",
                        visibleNumber,
                        $"Numbered equation table {tableIndex} has no visible right-aligned number.");
                }
                finally
                {
                    Release(numberCellParagraphs);
                    Release(numberCellRange);
                    Release(numberCell);
                    Release(table);
                }
            }

            // Word can finish auto-loading the installed VSTO add-in after the
            // initial COMAddIns.Connect check. This acceptance also hosts the
            // current build in-process, so a late installed instance would add a
            // second WindowBeforeDoubleClick handler and create a false duplicate
            // Session. Disconnect it again immediately before the real mouse
            // phase, after Word startup has fully settled.
            if (installedAddIn is not null)
            {
                if (installedAddIn.Connect)
                    installedAddIn.Connect = false;
                Thread.Sleep(250);
                if (installedAddIn.Connect)
                    throw new InvalidDataException(
                        "The installed Word add-in remained connected during the isolated OMML double-click acceptance.");
            }

            var path = Path.Combine(artifactRoot, "VisualTeX-Word-OMML-DoubleClick-Fixtures.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            Console.WriteLine("[OMML fixtures 7/8] Reopening and exercising real VisualTeX double-click editing...");
            document = application.Documents.Open(path, ReadOnly: false, Visible: true);
            application.Visible = true;
            for (var tableIndex = 1; tableIndex <= 3; tableIndex++)
                AssertNumberedTableCellNormalized(
                    tableIndex,
                    $"reopened numbered table {tableIndex}");
            Word.Window? focusWindow = null;
            try
            {
                focusWindow = application.ActiveWindow;
                focusWindow.Activate();
                if (consoleWindow != IntPtr.Zero) ShowWindow(consoleWindow, 0);
                var wordWindowHandle = new IntPtr(focusWindow.Hwnd);
                const uint noMoveNoSizeShow = 0x0001 | 0x0002 | 0x0040;
                SetWindowPos(wordWindowHandle, new IntPtr(-1), 0, 0, 0, 0, noMoveNoSizeShow);
                SetForegroundWindow(wordWindowHandle);
                if (GetWindowRect(wordWindowHandle, out var wordWindowRectangle))
                {
                    SetCursorPos(
                        wordWindowRectangle.Left
                            + Math.Max(40, (wordWindowRectangle.Right - wordWindowRectangle.Left) / 2),
                        wordWindowRectangle.Top + 18);
                    mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                    mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                }
                WinForms.Application.DoEvents();
                Thread.Sleep(600);
            }
            finally
            {
                Release(focusWindow);
            }

            var openedFormulaIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var equationIndex = 1; equationIndex <= expectedEquationCount; equationIndex++)
            {
                Word.OMaths? maths = null;
                Word.OMath? math = null;
                Word.Range? equationRange = null;
                Word.Window? window = null;
                try
                {
                    maths = document.OMaths;
                    math = maths[equationIndex];
                    equationRange = math.Range;
                    var beforeXml = equationRange.WordOpenXML;
                    var beforeMathXml = XDocument.Parse(beforeXml);
                    var beforeMathText = string.Concat(
                        beforeMathXml.Descendants(mathNamespace + "t").Select(text => text.Value));
                    var beforeFractionCount = beforeMathXml.Descendants(mathNamespace + "f").Count();
                    var beforeNaryCount = beforeMathXml.Descendants(mathNamespace + "nary").Count();
                    var beforeMathCount = document.OMaths.Count;
                    var beforeParagraphCount = document.Paragraphs.Count;
                    var beforeStart = equationRange.Start;
                    var beforeAlignment = equationRange.ParagraphFormat.Alignment;
                    var beforeSpaceBefore = equationRange.ParagraphFormat.SpaceBefore;
                    var beforeSpaceAfter = equationRange.ParagraphFormat.SpaceAfter;
                    var beforeLineSpacingRule = equationRange.ParagraphFormat.LineSpacingRule;
                    equationRange.Select();
                    WinForms.Application.DoEvents();
                    Thread.Sleep(180);
                    window = application.ActiveWindow;
                    var wordWindowHandle = new IntPtr(window.Hwnd);
                    SetForegroundWindow(wordWindowHandle);
                    window.GetPoint(
                        out var left,
                        out var top,
                        out var width,
                        out var height,
                        equationRange);
                    if (width <= 0 || height <= 0)
                        throw new InvalidDataException(
                            $"Word did not return a visible rectangle for OMML equation {equationIndex}.");

                    var sessionsBefore = SnapshotSessionIds();
                    SetCursorPos(left + width / 2, top + height / 2);
                    Thread.Sleep(120);
                    for (var click = 0; click < 2; click++)
                    {
                        mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                        mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                        Thread.Sleep(90);
                    }

                    var editSessionId = WaitForNewSession(
                        sessionsBefore,
                        "word",
                        TimeSpan.FromSeconds(30));
                    var editSession = WaitForUnchangedEditorReady(
                        client,
                        editSessionId,
                        TimeSpan.FromSeconds(15));
                    AssertEqual("edit", editSession.Mode,
                        $"OMML equation {equationIndex} double-click did not create an edit Session.");
                    AssertEqual(FormulaOleContract.WordOmmlMode, editSession.ObjectMode,
                        $"OMML equation {equationIndex} double-click changed the object mode.");
                    if (string.IsNullOrWhiteSpace(editSession.FormulaId)
                        || !openedFormulaIds.Add(editSession.FormulaId))
                        throw new InvalidDataException(
                            $"OMML equation {equationIndex} double-click did not resolve one unique persistent formulaId.");

                    if (equationIndex == 2)
                    {
                        // The click target is intentionally inside the display
                        // OMath. This reproduces Word's clipped OMath.Range bug:
                        // replacement must use the bookmark-resolved full range,
                        // otherwise the new formula is inserted into the old one.
                        Commit(
                            client,
                            editSession,
                            "block",
                            FormulaOleContract.WordOmmlMode,
                            "\\frac{a+b}{c+d}+1",
                            numbered: false,
                            mathMl:
                                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
                                + "<mfrac><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow>"
                                + "<mrow><mi>c</mi><mo>+</mo><mi>d</mi></mrow></mfrac>"
                                + "<mo>+</mo><mn>1</mn></math>");
                    }
                    else
                    {
                        client.CloseEditorAsync(editSessionId, CancellationToken.None)
                            .GetAwaiter().GetResult();
                    }

                    var terminal = WaitForTerminal(
                        client,
                        editSessionId,
                        TimeSpan.FromSeconds(45));
                    AssertEqual("completed", terminal.Status,
                        terminal.Error
                            ?? $"OMML equation {equationIndex} double-click Session did not complete.");
                    client.CloseEditorAsync(editSessionId, CancellationToken.None)
                        .GetAwaiter().GetResult();
                    WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

                    Release(equationRange);
                    equationRange = null;
                    Release(math);
                    math = null;
                    Release(maths);
                    maths = document.OMaths;
                    math = maths[equationIndex];
                    equationRange = math.Range;

                    AssertEqual(beforeMathCount, document.OMaths.Count,
                        $"OMML equation {equationIndex} double-click edit lost or duplicated an equation.");
                    AssertEqual(beforeParagraphCount, document.Paragraphs.Count,
                        $"OMML equation {equationIndex} double-click edit inserted or removed a paragraph.");
                    AssertEqual(beforeStart, equationRange.Start,
                        $"OMML equation {equationIndex} was not replaced in place.");
                    AssertEqual(beforeAlignment, equationRange.ParagraphFormat.Alignment,
                        $"OMML equation {equationIndex} changed paragraph alignment.");
                    AssertNear(beforeSpaceBefore, equationRange.ParagraphFormat.SpaceBefore, 0.1f,
                        $"OMML equation {equationIndex} changed paragraph space-before.");
                    AssertNear(beforeSpaceAfter, equationRange.ParagraphFormat.SpaceAfter, 0.1f,
                        $"OMML equation {equationIndex} changed paragraph space-after.");
                    AssertEqual(beforeLineSpacingRule, equationRange.ParagraphFormat.LineSpacingRule,
                        $"OMML equation {equationIndex} changed line-spacing rules.");

                    if (equationIndex == 2)
                    {
                        var editedXml = XDocument.Parse(equationRange.WordOpenXML);
                        AssertEqual(1, editedXml.Descendants(mathNamespace + "f").Count(),
                            "Partial-range OMML edit left the old fraction behind the replacement.");
                        AssertEqual(1, editedXml.Descendants(mathNamespace + "oMath").Count(),
                            "Partial-range OMML edit nested or duplicated the native equation.");
                        var editedMathText = string.Concat(
                            editedXml.Descendants(mathNamespace + "t").Select(text => text.Value));
                        if (editedMathText.IndexOf('1') < 0)
                            throw new InvalidDataException(
                                "Partial-range OMML edit did not persist the added digit. "
                                + $"ActualMathText='{editedMathText}'.");
                    }
                    else
                    {
                        // Word may rewrite volatile run properties or rsid data
                        // merely by entering/leaving an equation. Compare the
                        // mathematical content and structure rather than the
                        // byte-for-byte WordOpenXML wrapper.
                        var unchangedXml = XDocument.Parse(equationRange.WordOpenXML);
                        AssertEqual(
                            beforeMathText,
                            string.Concat(unchangedXml
                                .Descendants(mathNamespace + "t")
                                .Select(text => text.Value)),
                            $"Unchanged OMML equation {equationIndex} changed visible math text.");
                        AssertEqual(beforeFractionCount,
                            unchangedXml.Descendants(mathNamespace + "f").Count(),
                            $"Unchanged OMML equation {equationIndex} changed fraction structure.");
                        AssertEqual(beforeNaryCount,
                            unchangedXml.Descendants(mathNamespace + "nary").Count(),
                            $"Unchanged OMML equation {equationIndex} changed n-ary structure.");
                    }

                    Console.WriteLine(
                        equationIndex == 2
                            ? "  OMML equation 2: VisualTeX opened and full-range in-place edit passed."
                            : $"  OMML equation {equationIndex}: VisualTeX opened and unchanged close passed.");
                }
                finally
                {
                    Release(window);
                    Release(equationRange);
                    Release(math);
                    Release(maths);
                }
            }

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: true);
            AssertEqual(expectedEquationCount, document.OMaths.Count,
                "Save/reopen changed the OMML equation inventory after VisualTeX editing.");
            var reopenedEditXml = XDocument.Parse(document.OMaths[2].Range.WordOpenXML);
            AssertEqual(1, reopenedEditXml.Descendants(mathNamespace + "f").Count(),
                "Save/reopen restored a duplicated fraction after in-place OMML editing.");
            var reopenedEditMathText = string.Concat(
                reopenedEditXml.Descendants(mathNamespace + "t").Select(text => text.Value));
            if (reopenedEditMathText.IndexOf('1') < 0)
                throw new InvalidDataException(
                    "Save/reopen lost the digit added through the OMML VisualTeX editor. "
                    + $"ActualMathText='{reopenedEditMathText}'.");
            Console.WriteLine($"[OMML fixtures 8/8] Saved {path}; VisualTeX double-click and in-place edit checks passed.");
        }
        finally
        {
            if (consoleWindow != IntPtr.Zero) ShowWindow(consoleWindow, 5);
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            if (installedAddIn is not null)
            {
                try { installedAddIn.Connect = true; } catch { }
            }
            Release(installedAddIn);
            Release(installedAddIns);
            Release(oleShape);
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

    private static void RunTargetedWord2364(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? range = null;
        Word.Selection? eventSelection = null;
        Word.InlineShape? shape = null;
        COMAddIns? installedAddIns = null;
        COMAddIn? installedAddIn = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            Console.WriteLine("[Targeted Word 1/6] Starting Word from the clean baseline...");
            application = CreateWordApplication(visible: false);
            installedAddIns = application.COMAddIns;
            try
            {
                object addInIndex = "VisualTeX.WordVsto";
                installedAddIn = installedAddIns.Item(ref addInIndex);
                if (installedAddIn.Connect) installedAddIn.Connect = false;
            }
            catch
            {
                Release(installedAddIn);
                installedAddIn = null;
            }
            document = application.Documents.Add();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);

            Console.WriteLine("[Targeted Word 2/6] Inserting a bare display integral and validating native OMML growth...");
            var existing = SnapshotSessionIds();
            addIn.OnInsertDisplayOmml(new object());
            var bareSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var bareSession = client.GetSessionAsync(bareSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            const string bareIntegralLatex = "\\int f(x)\\,dx";
            const string bareIntegralMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
                + "<mo>∫</mo><mrow><mi>f</mi><mo>(</mo><mi>x</mi><mo>)</mo>"
                + "<mspace width=\"0.167em\"/><mi>d</mi><mi>x</mi></mrow></math>";
            Commit(
                client,
                bareSession,
                "block",
                FormulaOleContract.WordOmmlMode,
                bareIntegralLatex,
                numbered: false,
                mathMl: bareIntegralMathMl);
            var final = WaitForTerminal(client, bareSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "Bare display integral did not complete.");
            client.CloseEditorAsync(bareSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            var bareFormulaId = final.FormulaId
                ?? throw new InvalidDataException("Bare display integral has no formulaId.");

            maths = document.OMaths;
            AssertEqual(1, maths.Count, "Bare display integral did not create exactly one OMath.");
            math = maths[1];
            AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
                "Bare integral degraded to inline OMML.");
            range = math.Range;
            var bareXml = XDocument.Parse(range.WordOpenXML);
            XNamespace mathNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var nary = bareXml.Descendants(mathNamespace + "nary").SingleOrDefault()
                ?? throw new InvalidDataException("Bare display integral did not become native m:nary.");
            var naryProperties = nary.Element(mathNamespace + "naryPr");
            AssertEqual(
                "1",
                naryProperties?.Element(mathNamespace + "grow")
                    ?.Attribute(mathNamespace + "val")?.Value ?? string.Empty,
                "Bare display integral is not configured to grow.");
            AssertEqual(
                "1",
                naryProperties?.Element(mathNamespace + "subHide")
                    ?.Attribute(mathNamespace + "val")?.Value ?? string.Empty,
                "Bare display integral exposes an empty lower-limit placeholder.");
            AssertEqual(
                "1",
                naryProperties?.Element(mathNamespace + "supHide")
                    ?.Attribute(mathNamespace + "val")?.Value ?? string.Empty,
                "Bare display integral exposes an empty upper-limit placeholder.");

            Console.WriteLine("[Targeted Word 3/6] Verifying VisualTeX OMML double-click editing...");
            range.Select();
            existing = SnapshotSessionIds();
            eventSelection = application.Selection;
            var doubleClickHandler = typeof(VisualTeX.WordVsto.ThisAddIn).GetMethod(
                "OnWindowBeforeDoubleClick",
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.NonPublic)
                ?? throw new MissingMethodException("Word double-click handler is missing.");
            var doubleClickArguments = new object[] { eventSelection, false };
            doubleClickHandler.Invoke(addIn, doubleClickArguments);
            AssertEqual(true, (bool)doubleClickArguments[1],
                "VisualTeX OMML double-click was not intercepted.");
            Release(eventSelection);
            eventSelection = null;
            var editSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var editSession = WaitForUnchangedEditorReady(
                client,
                editSessionId,
                TimeSpan.FromSeconds(10));
            AssertEqual("edit", editSession.Mode,
                "OMML double-click did not create an edit Session.");
            AssertEqual(FormulaOleContract.WordOmmlMode, editSession.ObjectMode,
                "OMML double-click changed the object mode.");
            AssertEqual(bareFormulaId, editSession.FormulaId,
                "OMML double-click opened the wrong formula.");
            client.CloseEditorAsync(editSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            final = WaitForTerminal(client, editSessionId, TimeSpan.FromSeconds(30));
            AssertEqual("completed", final.Status,
                final.Error ?? "Unchanged OMML double-click Session did not complete.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            Console.WriteLine("[Targeted Word 4/6] Verifying direct OMML↔OLE conversion without an editor window...");
            Release(range);
            range = null;
            Release(math);
            math = null;
            Release(maths);
            maths = document.OMaths;
            math = maths[1];
            range = math.Range;
            range.Select();
            existing = SnapshotSessionIds();
            final = WaitForDirectConversion(
                client,
                existing,
                "word",
                FormulaOleContract.NativeOleMode,
                () => addIn.OnConvertSelected(new object()),
                TimeSpan.FromSeconds(45),
                out var ommlToOleElapsed);
            AssertEqual("completed", final.Status,
                final.Error ?? "Direct Word OMML-to-OLE conversion did not complete.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            AssertEqual(1, document.InlineShapes.Count,
                "Direct Word OMML-to-OLE conversion did not create one OLE object.");
            shape = document.InlineShapes[1];
            AssertEqual(FormulaOleContract.ProgId, shape.OLEFormat.ProgID,
                "Direct Word conversion created the wrong OLE class.");

            shape.Range.Select();
            existing = SnapshotSessionIds();
            final = WaitForDirectConversion(
                client,
                existing,
                "word",
                FormulaOleContract.WordOmmlMode,
                () => addIn.OnConvertSelectedToOmml(new object()),
                TimeSpan.FromSeconds(45),
                out var oleToOmmlElapsed);
            AssertEqual("completed", final.Status,
                final.Error ?? "Direct Word OLE-to-OMML conversion did not complete.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            AssertEqual(1, document.OMaths.Count,
                "Direct Word OLE-to-OMML conversion lost or duplicated the equation.");
            Console.WriteLine(
                $"  Direct conversions completed in {ommlToOleElapsed.TotalSeconds:F2}s and "
                + $"{oleToOmmlElapsed.TotalSeconds:F2}s without a visible VisualTeX window.");

            var focusedPath = Path.Combine(artifactRoot, "VisualTeX-Targeted-Word-2364.docx");
            document.SaveAs2(focusedPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Add();

            Console.WriteLine("[Targeted Word 5/6] Measuring numbered-formula insertion with existing formulas...");
            TimeSpan finalInsertionElapsed = TimeSpan.Zero;
            for (var index = 1; index <= 6; index++)
            {
                if (index > 1) application.Selection.TypeParagraph();
                existing = SnapshotSessionIds();
                addIn.OnInsertDisplay(new object());
                var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
                var session = client.GetSessionAsync(sessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
                var stopwatch = Stopwatch.StartNew();
                Commit(
                    client,
                    session,
                    "block",
                    FormulaOleContract.NativeOleMode,
                    $"x_{{{index}}}={index}",
                    numbered: true);
                final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
                stopwatch.Stop();
                AssertEqual("completed", final.Status,
                    final.Error ?? $"Numbered formula {index} did not complete.");
                client.CloseEditorAsync(sessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
                if (index == 6) finalInsertionElapsed = stopwatch.Elapsed;
            }
            if (finalInsertionElapsed > TimeSpan.FromSeconds(6))
                throw new InvalidDataException(
                    $"The sixth numbered formula took {finalInsertionElapsed.TotalSeconds:F2}s "
                    + "after commit; insertion still performs excessive full-document work.");
            var referenceItems = document.GetCrossReferenceItems(
                Word.WdCaptionLabelID.wdCaptionEquation) as Array;
            AssertEqual(6, referenceItems?.Length ?? 0,
                "Optimized insertion changed the native Equation reference inventory.");
            Console.WriteLine($"  Field inventory after insertion: {document.Fields.Count} fields.");
            var embedCount = 0;
            var refResults = new List<string>();
            var seqResults = new List<string>();
            for (var fieldIndex = 1; fieldIndex <= document.Fields.Count; fieldIndex++)
            {
                Word.Field? inventoryField = null;
                Word.Range? inventoryCode = null;
                Word.Range? inventoryResult = null;
                try
                {
                    inventoryField = document.Fields[fieldIndex];
                    inventoryCode = inventoryField.Code;
                    inventoryResult = inventoryField.Result;
                    var codeText = (inventoryCode.Text ?? string.Empty).Trim();
                    var resultText = (inventoryResult.Text ?? string.Empty).Trim();
                    if (codeText.StartsWith(
                            "EMBED VisualTeX.Formula.1",
                            StringComparison.OrdinalIgnoreCase))
                        embedCount++;
                    else if (codeText.StartsWith(
                                 "REF VTEqNum_",
                                 StringComparison.OrdinalIgnoreCase))
                        refResults.Add(resultText);
                    else if (codeText.IndexOf("SEQ ", StringComparison.OrdinalIgnoreCase) >= 0)
                        seqResults.Add(resultText);
                    Console.WriteLine(
                        $"    [{fieldIndex}] pos={inventoryResult.Start}, "
                        + $"code='{codeText}', result='{resultText}'");
                }
                finally
                {
                    Release(inventoryResult);
                    Release(inventoryCode);
                    Release(inventoryField);
                }
            }
            AssertEqual(6, embedCount,
                "Numbered OLE insertion changed the embedded-object field inventory.");
            AssertEqual("1,2,3,4,5,6", string.Join(",", seqResults),
                "Native SEQ fields were not numbered in document order.");
            AssertEqual("1,2,3,4,5,6", string.Join(",", refResults),
                "Visible REF numbers were stale immediately after insertion.");

            Console.WriteLine("[Targeted Word 6/6] Running the explicit full numbering refresh...");
            addIn.OnUpdateEquationNumbers(new object());
            Thread.Sleep(100);
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(15));
            AssertEqual(18, document.Fields.Count,
                "Explicit numbering refresh changed the EMBED/SEQ/REF field inventory.");
            var performancePath = Path.Combine(
                artifactRoot,
                "VisualTeX-Targeted-Word-Numbering-Performance.docx");
            document.SaveAs2(performancePath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                $"  Sixth numbered insertion commit completed in "
                + $"{finalInsertionElapsed.TotalSeconds:F2}s; numbering and references remained intact.");
        }
        finally
        {
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            if (installedAddIn is not null)
            {
                try { installedAddIn.Connect = true; } catch { }
            }
            Release(installedAddIn);
            Release(installedAddIns);
            Release(shape);
            Release(eventSelection);
            Release(range);
            Release(math);
            Release(maths);
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

    private static void RunWordFontSizeAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        Word.InlineShape? oleShape = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Selection? selection = null;
        COMAddIns? installedAddIns = null;
        COMAddIn? installedAddIn = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = CreateWordApplication(visible: true);
            installedAddIns = application.COMAddIns;
            try
            {
                object addInIndex = "VisualTeX.WordVsto";
                installedAddIn = installedAddIns.Item(ref addInIndex);
                if (installedAddIn.Connect) installedAddIn.Connect = false;
            }
            catch
            {
                Release(installedAddIn);
                installedAddIn = null;
            }

            document = application.Documents.Add();
            document.Activate();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);

            Console.WriteLine("[Word font size 1/6] Creating an inline OLE formula...");
            var existing = SnapshotSessionIds();
            addIn.OnInsertInline(new object());
            var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var session = client.GetSessionAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            Commit(client, session, "inline", FormulaOleContract.NativeOleMode, "x+y");
            var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status, final.Error ?? "Word OLE font-size fixture failed.");
            client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            oleShape = document.InlineShapes[1];
            var oleWidth12 = oleShape.Width;
            var oleHeight12 = oleShape.Height;
            var oleRatio12 = oleWidth12 / oleHeight12;
            var olePosition12 = oleShape.Range.Font.Position;
            oleShape.Range.Select();
            AssertTrue(addIn.GetFormulaFontSizeEnabled(null!),
                "Word font-size control was disabled for inline OLE.");

            Console.WriteLine("[Word font size 2/6] Setting OLE to 36 pt, then decreasing and increasing...");
            addIn.OnFormulaFontSizeChanged(null!, "36");
            Release(oleShape);
            oleShape = document.InlineShapes[1];
            var oleWidth36 = oleShape.Width;
            var oleHeight36 = oleShape.Height;
            var oleRatio36 = oleWidth36 / oleHeight36;
            AssertTrue(oleWidth36 > oleWidth12 && oleHeight36 > oleHeight12,
                $"Word OLE did not enlarge at 36 pt: {oleWidth12:F2}x{oleHeight12:F2} -> {oleWidth36:F2}x{oleHeight36:F2}.");
            AssertNear(oleRatio12, oleRatio36, 0.08f,
                "Word OLE font-size change distorted aspect ratio.");
            oleShape.Range.Select();
            var oleSelection36 = new VisualTeX.WordVsto.WordFormulaService(application).ReadSelection();
            var oleMetadata36 = oleSelection36.Metadata
                ?? throw new InvalidDataException("Word OLE 36 pt metadata could not be read.");
            var expectedOlePosition36 = WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                oleHeight36,
                (float)(oleMetadata36.RenderHeightPx ?? 0),
                oleMetadata36.Baseline.HasValue ? (float?)oleMetadata36.Baseline.Value : null,
                existingFontPosition: olePosition12,
                sourceSemanticFontSizePoints: FormulaFontSize.DefaultPt,
                targetSemanticFontSizePoints: 36);
            AssertNear(expectedOlePosition36, oleShape.Range.Font.Position, 0.1f,
                "Word OLE baseline did not follow its stored render geometry.");
            Console.WriteLine(
                $"  OLE baseline position: {olePosition12:F0} pt at {FormulaFontSize.DefaultPt:F0} pt -> "
                + $"{oleShape.Range.Font.Position:F0} pt at 36 pt (expected {expectedOlePosition36}; "
                + $"renderHeight={oleMetadata36.RenderHeightPx?.ToString("F2", System.Globalization.CultureInfo.InvariantCulture) ?? "null"}, "
                + $"baseline={oleMetadata36.Baseline?.ToString("F2", System.Globalization.CultureInfo.InvariantCulture) ?? "null"}).");
            AssertEqual(FormulaFontSize.FormatDisplay(36), addIn.GetFormulaFontSizeText(null!),
                "Word OLE font-size control did not report 36 pt.");
            addIn.OnDecreaseFormulaFontSize(new object());
            Release(oleShape);
            oleShape = document.InlineShapes[1];
            oleShape.Range.Select();
            var oleTextAfterDecrease = addIn.GetFormulaFontSizeText(null!);
            var oleWidthAfterDecrease = oleShape.Width;
            var oleHeightAfterDecrease = oleShape.Height;
            Console.WriteLine(
                $"  OLE after decrease: control='{oleTextAfterDecrease}', "
                + $"size={oleWidthAfterDecrease:F2}x{oleHeightAfterDecrease:F2}.");
            if (!(oleWidthAfterDecrease < oleWidth36 && oleHeightAfterDecrease < oleHeight36))
                Console.WriteLine(
                    $"  [DISCREPANCY] Word 2021 OLE decrease did not reduce both dimensions: "
                    + $"{oleWidth36:F2}x{oleHeight36:F2} -> "
                    + $"{oleWidthAfterDecrease:F2}x{oleHeightAfterDecrease:F2}; "
                    + $"control='{oleTextAfterDecrease}'.");

            addIn.OnIncreaseFormulaFontSize(new object());
            Release(oleShape);
            oleShape = document.InlineShapes[1];
            oleShape.Range.Select();
            var oleTextAfterIncrease = addIn.GetFormulaFontSizeText(null!);
            Console.WriteLine(
                $"  OLE after increase: control='{oleTextAfterIncrease}', "
                + $"size={oleShape.Width:F2}x{oleShape.Height:F2}.");
            if (!string.Equals(
                    oleTextAfterIncrease,
                    FormulaFontSize.FormatDisplay(36),
                    StringComparison.Ordinal)
                || Math.Abs(oleShape.Width - oleWidth36) > 1.0f
                || Math.Abs(oleShape.Height - oleHeight36) > 0.6f)
                Console.WriteLine(
                    $"  [DISCREPANCY] Word 2021 OLE decrease/increase did not restore 36 pt: "
                    + $"control='{oleTextAfterIncrease}', size={oleShape.Width:F2}x{oleShape.Height:F2}.");

            addIn.OnFormulaFontSizeChanged(null!, "36");
            Release(oleShape);
            oleShape = document.InlineShapes[1];
            oleShape.Range.Select();
            AssertNear(oleWidth36, oleShape.Width, 1.0f,
                "Word OLE direct 36 pt restore did not recover width.");
            AssertNear(oleHeight36, oleShape.Height, 0.6f,
                "Word OLE direct 36 pt restore did not recover height.");

            Console.WriteLine("[Word font size 3/6] Creating an inline OMML formula...");
            selection = application.Selection;
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeText(" text ");
            existing = SnapshotSessionIds();
            addIn.OnInsertInlineOmml(new object());
            sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            session = client.GetSessionAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            const string mathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>a</mi><mo>+</mo><mi>b</mi></math>";
            Commit(
                client,
                session,
                "inline",
                FormulaOleContract.WordOmmlMode,
                "a+b",
                mathMl: mathMl);
            final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status, final.Error ?? "Word OMML font-size fixture failed.");
            client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            maths = document.OMaths;
            AssertEqual(1, maths.Count, "Word OMML font-size fixture did not create one OMath.");
            math = maths[1];
            mathRange = math.Range;
            mathRange.Select();
            AssertTrue(addIn.GetFormulaFontSizeEnabled(null!),
                "Word font-size control was disabled for inline OMML.");

            Console.WriteLine("[Word font size 4/6] Setting OMML to 36 pt, then decreasing and increasing...");
            addIn.OnFormulaFontSizeChanged(null!, "36");
            Release(mathRange);
            Release(math);
            Release(maths);
            maths = document.OMaths;
            math = maths[1];
            mathRange = math.Range;
            AssertNear(36f, mathRange.Font.Size, 0.6f,
                "Word OMML native range did not become 36 pt.");
            AssertNear(0f, mathRange.Font.Position, 0.1f,
                "Word OMML range received a manual baseline offset.");
            mathRange.Select();
            AssertEqual(FormulaFontSize.FormatDisplay(36), addIn.GetFormulaFontSizeText(null!),
                "Word OMML font-size control did not report 36 pt.");
            addIn.OnDecreaseFormulaFontSize(new object());
            Release(mathRange);
            Release(math);
            Release(maths);
            maths = document.OMaths;
            math = maths[1];
            mathRange = math.Range;
            AssertTrue(mathRange.Font.Size < 36f,
                "Word OMML decrease did not reduce the native font size.");
            mathRange.Select();
            addIn.OnIncreaseFormulaFontSize(new object());
            Release(mathRange);
            Release(math);
            Release(maths);
            maths = document.OMaths;
            math = maths[1];
            mathRange = math.Range;
            AssertNear(36f, mathRange.Font.Size, 0.6f,
                "Word OMML decrease/increase did not restore 36 pt.");

            Console.WriteLine("[Word font size 5/6] Saving and reopening both object types...");
            var path = Path.Combine(artifactRoot, "VisualTeX-Word-Font-Size.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            reopened = application.Documents.Open(path, ReadOnly: false, Visible: true);
            AssertEqual(1, reopened.InlineShapes.Count,
                "Reopened Word font-size document lost the OLE formula.");
            AssertEqual(1, reopened.OMaths.Count,
                "Reopened Word font-size document lost the OMML formula.");
            Release(oleShape);
            oleShape = reopened.InlineShapes[1];
            AssertNear(oleWidth36, oleShape.Width, 1.0f,
                "Word OLE width changed after save/reopen.");
            AssertNear(oleHeight36, oleShape.Height, 0.6f,
                "Word OLE height changed after save/reopen.");
            AssertNear(expectedOlePosition36, oleShape.Range.Font.Position, 0.1f,
                "Word OLE baseline position changed after save/reopen.");
            oleShape.Range.Select();
            AssertEqual(FormulaFontSize.FormatDisplay(36), addIn.GetFormulaFontSizeText(null!),
                "Word OLE semantic size was lost after save/reopen.");
            Release(mathRange);
            Release(math);
            Release(maths);
            maths = reopened.OMaths;
            math = maths[1];
            mathRange = math.Range;
            AssertNear(36f, mathRange.Font.Size, 0.6f,
                "Word OMML native size changed after save/reopen.");
            mathRange.Select();
            AssertEqual(FormulaFontSize.FormatDisplay(36), addIn.GetFormulaFontSizeText(null!),
                "Word OMML semantic size was lost after save/reopen.");

            Console.WriteLine(
                $"[Word font size 6/6] OLE {oleWidth12:F1}x{oleHeight12:F1} -> "
                + $"{oleWidth36:F1}x{oleHeight36:F1}; OMML 36 pt; presets and save/reopen passed. Artifact: {path}");
        }
        finally
        {
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            if (installedAddIn is not null)
            {
                try { installedAddIn.Connect = true; } catch { }
            }
            Release(selection);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(oleShape);
            if (reopened is not null)
            {
                try { reopened.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(reopened);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            Release(installedAddIn);
            Release(installedAddIns);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordNativeOmmlSourceProbe(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        COMAddIns? installedAddIns = null;
        COMAddIn? installedAddIn = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = CreateWordApplication(visible: false);
            installedAddIns = application.COMAddIns;
            try
            {
                object addInIndex = "VisualTeX.WordVsto";
                installedAddIn = installedAddIns.Item(ref addInIndex);
                if (installedAddIn.Connect) installedAddIn.Connect = false;
            }
            catch
            {
                Release(installedAddIn);
                installedAddIn = null;
            }

            document = application.Documents.Add();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);
            const string initialMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
                + "<msup><mi>x</mi><mn>2</mn></msup><mo>+</mo>"
                + "<msup><mi>y</mi><mn>2</mn></msup></math>";

            Console.WriteLine("[Native OMML source 1/4] Creating a VisualTeX OMML formula...");
            var existing = SnapshotSessionIds();
            addIn.OnInsertDisplayOmml(new object());
            var createSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var createSession = client.GetSessionAsync(createSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(
                client,
                createSession,
                "block",
                FormulaOleContract.WordOmmlMode,
                "x^2+y^2",
                numbered: false,
                mathMl: initialMathMl);
            var final = WaitForTerminal(client, createSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "Native OMML source fixture did not complete.");
            client.CloseEditorAsync(createSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            Console.WriteLine("[Native OMML source 2/4] Appending +z^3 with Word's native equation object...");
            AppendToLastWordOmmlAndSelect(document, "+z^3");

            Console.WriteLine("[Native OMML source 3/4] Reopening through the Ribbon edit command...");
            existing = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            var editSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var editSession = WaitForUnchangedEditorReady(
                client,
                editSessionId,
                TimeSpan.FromSeconds(15));
            AssertEqual(FormulaOleContract.WordOmmlMode, editSession.ObjectMode,
                "Native-edited OMML opened with the wrong object mode.");
            var imported = string.Join("\n", editSession.Lines.Select(line => line.Latex));
            if (imported.IndexOf("z", StringComparison.Ordinal) < 0
                || (imported.IndexOf("z^3", StringComparison.Ordinal) < 0
                    && imported.IndexOf("z^{3}", StringComparison.Ordinal) < 0))
                throw new InvalidDataException(
                    "The editor used stale metadata instead of Word's latest OMML. "
                    + $"Imported source: {imported}");
            client.CloseEditorAsync(editSessionId, CancellationToken.None).GetAwaiter().GetResult();
            final = WaitForTerminal(client, editSessionId, TimeSpan.FromSeconds(30));
            AssertEqual("completed", final.Status,
                final.Error ?? "Native-edited OMML unchanged edit did not complete.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            var path = Path.Combine(artifactRoot, "VisualTeX-Word-Native-OMML-Latest-Source.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                $"[Native OMML source 4/4] Saved {path}; Ribbon edit imported Word's latest +z^3 source.");
        }
        finally
        {
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            if (installedAddIn is not null)
            {
                try { installedAddIn.Connect = true; } catch { }
            }
            Release(installedAddIn);
            Release(installedAddIns);
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

    private static void RunWord(
        VisualTeXSessionClient client,
        string artifactRoot,
        bool initialOnly = false,
        bool stopAfterUnchanged = false,
        bool stopAfterOlePictureRoundTrip = false,
        bool skipDoubleClickForOlePictureRoundTrip = false)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.OLEFormat? wordOleFormat = null;
        Word.Range? typedRange = null;
        Word.Selection? eventSelection = null;
        Word.InlineShape? numberedShape = null;
        Process? testOleServerProcess = null;
        object? oleServerKeepAlive = null;
        COMAddIns? installedAddIns = null;
        COMAddIn? installedAddIn = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            var testOleServerPath = Environment.GetEnvironmentVariable("VISUALTEX_TEST_OLE_SERVER_PATH");
            if (!string.IsNullOrWhiteSpace(testOleServerPath))
            {
                testOleServerProcess = Process.Start(new ProcessStartInfo
                {
                    FileName = testOleServerPath,
                    Arguments = "-Embedding",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                }) ?? throw new InvalidOperationException("Failed to start the test VisualTeX OLE server.");
                Thread.Sleep(500);
                if (testOleServerProcess.HasExited)
                    throw new InvalidOperationException("The test VisualTeX OLE server exited before Word COM activation.");
                var oleServerType = Type.GetTypeFromProgID(FormulaOleContract.ProgId, throwOnError: true)
                    ?? throw new InvalidOperationException("VisualTeX OLE server class is not registered.");
                oleServerKeepAlive = Activator.CreateInstance(oleServerType)
                    ?? throw new InvalidOperationException("VisualTeX OLE server keep-alive could not be created.");
                Console.WriteLine(
                    $"[Word probe] Test OLE server pinned: pid={testOleServerProcess.Id}, path={testOleServerPath}");
            }

            Console.WriteLine("[Word 1/12] Starting Word and creating an inline Session...");
            application = CreateWordApplication(visible: false);
            installedAddIns = application.COMAddIns;
            try
            {
                object addInIndex = "VisualTeX.WordVsto";
                installedAddIn = installedAddIns.Item(ref addInIndex);
                if (installedAddIn.Connect)
                    installedAddIn.Connect = false;
            }
            catch
            {
                Release(installedAddIn);
                installedAddIn = null;
            }
            document = application.Documents.Add();
            application.Selection.TypeText("Before ");
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);
            var existing = SnapshotSessionIds();
            addIn.OnInsertInline(new object());
            var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var session = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("nativeOle", session.ObjectMode, "Word must request native OLE.");
            AssertEqual("inline", session.DisplayMode, "Word inline Session mode is wrong.");

            Console.WriteLine("[Word 2/12] Committing the initial vector formula...");
            Commit(client, session, "inline", "nativeOle", "x+y=z");
            var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status, final.Error ?? "Word Session did not complete.");
            client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();

            Console.WriteLine("[Word 3/12] Checking OLE size, aspect ratio and formula baseline...");
            AssertEqual(1, document.InlineShapes.Count, "Word should contain one inline formula.");
            shape = document.InlineShapes[1];
            AssertNear(120f, shape.Width, 0.5f, "Word formula width is incorrect.");
            AssertNear(24f, shape.Height, 0.5f, "Word formula height is incorrect.");
            AssertNear(5f, shape.Width / shape.Height, 0.05f,
                "Word formula aspect ratio is distorted.");
            var expectedPosition = WordInlineAlignment.CalculateFontPosition(
                shape.Height,
                ExportHeight,
                ExportBaseline);
            AssertNear(expectedPosition, shape.Range.Font.Position, 0.1f,
                "Word inline formula baseline is incorrect.");
            AssertEqual(FormulaOleContract.ProgId, shape.OLEFormat.ProgID,
                "Word inserted the wrong OLE class.");
            if (initialOnly)
            {
                var probePath = Path.Combine(artifactRoot, "VisualTeX-Word-Create-Probe.docx");
                document.SaveAs2(probePath, Word.WdSaveFormat.wdFormatXMLDocument);
                Console.WriteLine($"[Word probe] Saved {probePath}; initial OLE creation passed.");
                return;
            }

            Console.WriteLine("[Word 4/12] Verifying that typing after the formula returns to the text baseline...");
            AssertNear(0f, application.Selection.Font.Position, 0.1f,
                "Word caret inherited the formula baseline offset.");
            var textStart = application.Selection.Start;
            application.Selection.TypeText(" after");
            object rangeStart = textStart;
            object rangeEnd = application.Selection.Start;
            typedRange = document.Range(ref rangeStart, ref rangeEnd);
            AssertNear(0f, typedRange.Font.Position, 0.1f,
                "Text typed after the formula inherited the formula baseline offset.");
            Release(typedRange);
            typedRange = null;

            Console.WriteLine("[Word 4b/12] Re-clicking the formula tail and typing on the body-text baseline...");
            var formulaTail = shape.Range.End;
            application.Selection.SetRange(formulaTail, formulaTail);
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(100);
            System.Windows.Forms.Application.DoEvents();
            AssertNear(0f, application.Selection.Font.Position, 0.1f,
                "Word caret at the re-clicked formula tail kept the OLE baseline offset.");
            var retypedStart = application.Selection.Start;
            application.Selection.TypeText(" reclick");
            object retypedRangeStart = retypedStart;
            object retypedRangeEnd = application.Selection.Start;
            typedRange = document.Range(ref retypedRangeStart, ref retypedRangeEnd);
            AssertNear(0f, typedRange.Font.Position, 0.1f,
                "Text typed after re-clicking the formula tail inherited the OLE baseline offset.");
            Release(typedRange);
            typedRange = null;

            Console.WriteLine("[Word 5/12] Closing an unchanged edit and waiting for the add-in to unlock...");
            shape.Range.Select();
            existing = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            var unchangedSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var unchangedSession = WaitForUnchangedEditorReady(
                client,
                unchangedSessionId,
                TimeSpan.FromSeconds(10));
            AssertEqual(false, unchangedSession.Dirty,
                "Word unchanged edit Session became dirty before closing.");
            client.CloseEditorAsync(unchangedSessionId, CancellationToken.None).GetAwaiter().GetResult();
            final = WaitForTerminal(client, unchangedSessionId, TimeSpan.FromSeconds(30));
            AssertEqual("completed", final.Status,
                final.Error ?? "Word unchanged edit did not complete after closing the window.");
            AssertEqual(false, final.Dirty,
                "Word unchanged edit was incorrectly marked dirty.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            Console.WriteLine(skipDoubleClickForOlePictureRoundTrip
                ? "[Word 6/12] Reopening through Edit Selected for the OLE/picture compatibility probe..."
                : "[Word 6/12] Reopening immediately through the double-click interception...");
            shape.Range.Select();
            existing = SnapshotSessionIds();
            if (skipDoubleClickForOlePictureRoundTrip)
            {
                addIn.OnEditSelected(new object());
            }
            else
            {
                eventSelection = application.Selection;
                var handler = typeof(VisualTeX.WordVsto.ThisAddIn).GetMethod(
                    "OnWindowBeforeDoubleClick",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                    ?? throw new MissingMethodException("Word double-click handler is missing.");
                var handlerArguments = new object[] { eventSelection, false };
                handler.Invoke(addIn, handlerArguments);
                AssertEqual(true, (bool)handlerArguments[1],
                    "Word did not suppress built-in OLE activation on double-click.");
                Release(eventSelection);
                eventSelection = null;
            }
            var editSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var editSession = client.GetSessionAsync(editSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("edit", editSession.Mode, "Word double-click did not create an edit Session.");
            AssertEqual("nativeOle", editSession.ObjectMode,
                "Word edit Session changed the object mode.");
            if (stopAfterUnchanged)
            {
                editSession = WaitForUnchangedEditorReady(
                    client,
                    editSessionId,
                    TimeSpan.FromSeconds(10));
                AssertEqual(false, editSession.Dirty,
                    "Word reopened edit Session was already dirty.");
                client.CloseEditorAsync(editSessionId, CancellationToken.None).GetAwaiter().GetResult();
                final = WaitForTerminal(client, editSessionId, TimeSpan.FromSeconds(30));
                AssertEqual("completed", final.Status,
                    final.Error ?? "Word reopened unchanged edit did not complete.");
                AssertEqual(false, final.Dirty,
                    "Word reopened unchanged edit was incorrectly marked dirty.");
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
                var probePath = Path.Combine(artifactRoot, "VisualTeX-Word-Unchanged-Probe.docx");
                document.SaveAs2(probePath, Word.WdSaveFormat.wdFormatXMLDocument);
                Console.WriteLine($"[Word unchanged probe] Saved {probePath}; close and immediate reopen passed.");
                return;
            }

            Console.WriteLine("[Word 7/12] Editing to a wider formula and checking natural resize...");
            Commit(
                client,
                editSession,
                "inline",
                "nativeOle",
                "x+y+z+a+b+c=d",
                renderWidth: 320f,
                renderHeight: 32f,
                baseline: 24f);
            final = WaitForTerminal(client, editSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status, final.Error ?? "Word edit Session did not complete.");
            client.CloseEditorAsync(editSessionId, CancellationToken.None).GetAwaiter().GetResult();
            Release(shape);
            shape = document.InlineShapes[1];
            Console.WriteLine(
                $"  Word edited OLE geometry: {shape.Width:F3} x {shape.Height:F3} pt, "
                + $"ratio={shape.Width / shape.Height:F4}");
            AssertNear(240f, shape.Width, 1.5f,
                "Word edited formula retained the old width.");
            AssertNear(24f, shape.Height, 0.5f,
                "Word edited formula height is incorrect.");
            AssertNear(10f, shape.Width / shape.Height, 0.08f,
                "Word edited formula is compressed into the old aspect ratio.");
            AssertNear(0f, application.Selection.Font.Position, 0.1f,
                "Word caret baseline was not restored after editing.");

            Console.WriteLine("[Word 8/12] Exporting the selected OLE formula as an editable picture...");
            shape.Range.Select();
            Release(shape);
            shape = null;
            _ = new VisualTeX.WordVsto.WordFormulaService(application).ExportSelectedOleAsPicture();
            shape = WaitForWordShapeMode(document, nativeOle: false, TimeSpan.FromSeconds(15));
            AssertEqual(Word.WdInlineShapeType.wdInlineShapePicture, shape.Type,
                "Word OLE export did not create a normal picture.");

            Console.WriteLine("[Word 9/12] Converting the unchanged picture back to native OLE...");
            shape.Range.Select();
            existing = SnapshotSessionIds();
            addIn.OnConvertSelected(new object());
            var convertSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var convertSession = client.GetSessionAsync(convertSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("edit", convertSession.Mode,
                "Word convert command did not create an edit Session.");
            AssertEqual("nativeOle", convertSession.ObjectMode,
                "Word convert command did not request native OLE.");
            Commit(
                client,
                convertSession,
                "inline",
                "nativeOle",
                "x+y+z+a+b+c=d",
                renderWidth: 320f,
                renderHeight: 32f,
                baseline: 24f,
                dirty: false);
            final = WaitForTerminal(client, convertSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "Word OLE conversion did not complete.");
            client.CloseEditorAsync(convertSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            Release(shape);
            shape = WaitForWordShapeMode(document, nativeOle: true, TimeSpan.FromSeconds(15));
            wordOleFormat = shape.OLEFormat;
            AssertEqual(FormulaOleContract.ProgId, wordOleFormat.ProgID,
                "Word converted to the wrong OLE class.");

            Console.WriteLine("[Word 10/12] Exercising the converted OLE show verb...");
            wordOleFormat.DoVerb(-1);
            Release(wordOleFormat);
            wordOleFormat = null;

            Console.WriteLine("[Word 11/17] Rechecking size after picture-to-OLE conversion...");
            AssertNear(240f, shape.Width, 1.5f,
                "Word picture-to-OLE conversion changed the formula width.");
            AssertNear(24f, shape.Height, 0.5f,
                "Word picture-to-OLE conversion changed the formula height.");
            if (stopAfterOlePictureRoundTrip)
            {
                var roundTripPath = Path.Combine(
                    artifactRoot,
                    "VisualTeX-Word-OLE-Picture-RoundTrip.docx");
                document.SaveAs2(roundTripPath, Word.WdSaveFormat.wdFormatXMLDocument);
                Console.WriteLine(
                    $"[Word OLE/picture round-trip] Saved {roundTripPath}; edit resize, "
                    + "picture export, picture-to-OLE conversion, OLE verb, dimensions, and baseline passed.");
                return;
            }
            Release(shape);
            shape = null;

            Console.WriteLine("[Word 12/17] Creating the first numbered display formula...");
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            existing = SnapshotSessionIds();
            addIn.OnInsertDisplay(new object());
            var firstNumberSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var firstNumberSession = client.GetSessionAsync(firstNumberSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(
                client,
                firstNumberSession,
                "block",
                "nativeOle",
                "a=b",
                numbered: true);
            final = WaitForTerminal(client, firstNumberSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "First numbered Word formula did not complete.");
            client.CloseEditorAsync(firstNumberSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            var firstNumberFormulaId = final.FormulaId
                ?? throw new InvalidDataException("First numbered formula has no formulaId.");
            WaitForWordInlineShapeCount(document, 2, TimeSpan.FromSeconds(15));

            Console.WriteLine("[Word 13/17] Creating the second numbered display formula...");
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            existing = SnapshotSessionIds();
            addIn.OnInsertDisplay(new object());
            var secondNumberSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var secondNumberSession = client.GetSessionAsync(secondNumberSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(
                client,
                secondNumberSession,
                "block",
                "nativeOle",
                "E=mc^2",
                numbered: true);
            final = WaitForTerminal(client, secondNumberSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "Second numbered Word formula did not complete.");
            client.CloseEditorAsync(secondNumberSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            var secondNumberFormulaId = final.FormulaId
                ?? throw new InvalidDataException("Second numbered formula has no formulaId.");
            WaitForWordInlineShapeCount(document, 3, TimeSpan.FromSeconds(15));

            Console.WriteLine("[Word 14/17] Inserting a live reference to equation (2)...");
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            application.Selection.TypeText("See ");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REFERENCE_TARGET_INDEX", "1");
            var nativeItems = document.GetCrossReferenceItems(Word.WdCaptionLabelID.wdCaptionEquation) as Array;
            if (nativeItems is null || nativeItems.Length != 2)
                throw new InvalidDataException(
                    $"Word native Equation list should contain two VisualTeX formulas, actual count: {nativeItems?.Length ?? 0}.");
            addIn.OnInsertEquationReference(new object());
            var nativeReferenceCode = WaitForWordNativeReferenceResult(
                document,
                expectedResult: "2",
                expectedCode: null,
                TimeSpan.FromSeconds(15));
            if (!DocumentTextContains(document, "(2)"))
                throw new InvalidDataException("Word native reference did not include parenthesized equation number (2).");

            Console.WriteLine("[Word 15/17] Deleting equation (1) through the VisualTeX command...");
            numberedShape = document.InlineShapes[2];
            numberedShape.Range.Select();
            addIn.OnDeleteSelected(new object());
            Release(numberedShape);
            numberedShape = null;
            WaitForWordInlineShapeCount(document, 2, TimeSpan.FromSeconds(15));

            Console.WriteLine("[Word 16/17] Updating equation numbers and the REF field...");
            addIn.OnUpdateEquationNumbers(new object());
            WaitForWordNativeReferenceResult(
                document,
                expectedResult: "1",
                expectedCode: nativeReferenceCode,
                TimeSpan.FromSeconds(15));
            if (!DocumentTextContains(document, "(1)"))
                throw new InvalidDataException("Word native reference did not update to equation number (1).");
            if (WordBookmarkExists(document, $"VTEq_{Guid.Parse(firstNumberFormulaId):N}"))
                throw new InvalidDataException("Deleted equation retained its VisualTeX number bookmark.");
            if (!WordBookmarkExists(document, $"VTEq_{Guid.Parse(secondNumberFormulaId):N}"))
                throw new InvalidDataException("Referenced equation lost its persistent VisualTeX number bookmark.");
            nativeItems = document.GetCrossReferenceItems(Word.WdCaptionLabelID.wdCaptionEquation) as Array;
            if (nativeItems is null || nativeItems.Length != 1)
                throw new InvalidDataException(
                    $"Word native Equation list should contain one formula after deletion, actual count: {nativeItems?.Length ?? 0}.");

            Console.WriteLine("[Word 17/21] Creating a real Word OMML formula through the VisualTeX editor...");
            const string ommlInitialMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
                + "<msup><mi>x</mi><mn>2</mn></msup><mo>+</mo>"
                + "<msup><mi>y</mi><mn>2</mn></msup></math>";
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            existing = SnapshotSessionIds();
            addIn.OnInsertDisplayOmml(new object());
            var ommlCreateSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var ommlCreateSession = client.GetSessionAsync(ommlCreateSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("wordOmml", ommlCreateSession.ObjectMode,
                "Word OMML insert command did not create a wordOmml Session.");
            Commit(
                client,
                ommlCreateSession,
                "block",
                "wordOmml",
                "x^2+y^2",
                numbered: false,
                mathMl: ommlInitialMathMl);
            final = WaitForTerminal(client, ommlCreateSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "Word OMML creation did not complete.");
            client.CloseEditorAsync(ommlCreateSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            Console.WriteLine("[Word 18/21] Adding +z^3 through Word's native equation object...");
            AppendToLastWordOmmlAndSelect(document, "+z^3");

            Console.WriteLine("[Word 19/21] Opening the selected native-edited OMML through VisualTeX...");
            existing = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            var ommlEditSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var ommlEditSession = WaitForUnchangedEditorReady(
                client,
                ommlEditSessionId,
                TimeSpan.FromSeconds(10));
            AssertEqual("edit", ommlEditSession.Mode,
                "Native-edited OMML did not open an edit Session.");
            AssertEqual("wordOmml", ommlEditSession.ObjectMode,
                "Native-edited OMML opened with the wrong object mode.");
            var importedNativeLatex = string.Join("\n", ommlEditSession.Lines.Select(line => line.Latex));
            if (importedNativeLatex.IndexOf("z", StringComparison.Ordinal) < 0
                || importedNativeLatex.IndexOf("^{3}", StringComparison.Ordinal) < 0)
                throw new InvalidDataException(
                    "VisualTeX editor Session did not include the +z^3 inserted by Word's native equation editor. "
                    + $"Imported source: {importedNativeLatex}");

            Console.WriteLine("[Word 20/21] Closing the unchanged OMML editor after verifying imported source...");
            client.CloseEditorAsync(ommlEditSessionId, CancellationToken.None).GetAwaiter().GetResult();
            final = WaitForTerminal(client, ommlEditSessionId, TimeSpan.FromSeconds(30));
            AssertEqual("completed", final.Status,
                final.Error ?? "Native-edited OMML unchanged edit did not complete.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            var path = Path.Combine(artifactRoot, "VisualTeX-Word-Flow.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                $"[Word 21/21] Saved {path}; conversion, foreground editor, live cross-reference, and native OMML source-import checks passed.");
        }
        finally
        {
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            if (installedAddIn is not null)
            {
                try { installedAddIn.Connect = true; } catch { }
            }
            Release(installedAddIn);
            Release(installedAddIns);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REFERENCE_TARGET_INDEX", null);
            Release(numberedShape);
            Release(eventSelection);
            Release(typedRange);
            Release(wordOleFormat);
            Release(shape);
            Release(oleServerKeepAlive);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            if (testOleServerProcess is not null)
            {
                try
                {
                    if (!testOleServerProcess.HasExited)
                        testOleServerProcess.Kill();
                }
                catch { }
                testOleServerProcess.Dispose();
            }
            ForceComCleanup();
        }
    }

    private static void RunTargetedPowerPoint46(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        PowerPoint.OLEFormat? oleFormat = null;
        Process? testOleServerProcess = null;
        object? oleServerKeepAlive = null;
        COMAddIns? installedAddIns = null;
        COMAddIn? installedAddIn = null;
        VisualTeX.PowerPointVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            var testOleServerPath = Environment.GetEnvironmentVariable("VISUALTEX_TEST_OLE_SERVER_PATH");
            if (!string.IsNullOrWhiteSpace(testOleServerPath))
            {
                testOleServerProcess = Process.Start(new ProcessStartInfo
                {
                    FileName = testOleServerPath,
                    Arguments = "-Embedding",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                }) ?? throw new InvalidOperationException("Failed to start the test VisualTeX OLE server.");
                Thread.Sleep(500);
                if (testOleServerProcess.HasExited)
                    throw new InvalidOperationException("The test VisualTeX OLE server exited before PowerPoint COM activation.");
                var oleServerType = Type.GetTypeFromProgID(FormulaOleContract.ProgId, throwOnError: true)
                    ?? throw new InvalidOperationException("VisualTeX OLE server class is not registered.");
                oleServerKeepAlive = Activator.CreateInstance(oleServerType)
                    ?? throw new InvalidOperationException("VisualTeX OLE server keep-alive could not be created.");
                Console.WriteLine(
                    $"[Targeted PowerPoint] Test OLE server pinned: pid={testOleServerProcess.Id}, path={testOleServerPath}");
            }

            Console.WriteLine("[Targeted PowerPoint 1/4] Starting PowerPoint and creating a picture formula...");
            application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            installedAddIns = application.COMAddIns;
            try
            {
                object addInIndex = "VisualTeX.PowerPointVsto";
                installedAddIn = installedAddIns.Item(ref addInIndex);
                if (installedAddIn.Connect) installedAddIn.Connect = false;
            }
            catch
            {
                Release(installedAddIn);
                installedAddIn = null;
            }
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            addIn = new VisualTeX.PowerPointVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);

            var existing = SnapshotSessionIds();
            addIn.OnNewFormula(new object());
            var createSessionId = WaitForNewSession(
                existing,
                "powerpoint",
                TimeSpan.FromSeconds(30));
            var createSession = client.GetSessionAsync(createSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(
                client,
                createSession,
                "block",
                FormulaOleContract.CrossPlatformPictureMode,
                "x+y=z");
            var final = WaitForTerminal(client, createSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "PowerPoint picture formula did not complete.");
            client.CloseEditorAsync(createSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            var formulaId = final.FormulaId
                ?? throw new InvalidDataException("PowerPoint picture formula has no formulaId.");
            AssertEqual(1, slide.Shapes.Count,
                "PowerPoint targeted fixture did not create one formula shape.");
            shape = slide.Shapes[1];
            AssertTrue(IsPowerPointEditablePictureShape(shape),
                $"PowerPoint targeted fixture did not start as an editable picture/graphic. Actual type: {(int)shape.Type}.");
            var pictureLeft = shape.Left;
            var pictureTop = shape.Top;
            var pictureWidth = shape.Width;
            var pictureHeight = shape.Height;
            var pictureRatio = pictureWidth / pictureHeight;
            var sourcePicturePng = Path.Combine(artifactRoot, "targeted-current-vsto-source.png");
            shape.Export(sourcePicturePng, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
            var sourceInk = AnalyzeDarkPixels(sourcePicturePng);
            AssertTrue(sourceInk.Count > 0, "Current VSTO source picture export contains no formula pixels.");

            Console.WriteLine("[Targeted PowerPoint 2/4] Converting picture to OLE without opening the editor...");
            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            final = WaitForDirectConversion(
                client,
                existing,
                "powerpoint",
                FormulaOleContract.NativeOleMode,
                () => addIn.OnConvertSelected(new object()),
                TimeSpan.FromSeconds(45),
                out var conversionElapsed);
            AssertEqual("completed", final.Status,
                final.Error ?? "Direct PowerPoint picture-to-OLE conversion did not complete.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            var conversionDeadline = DateTime.UtcNow.AddSeconds(15);
            while (DateTime.UtcNow < conversionDeadline)
            {
                WinForms.Application.DoEvents();
                Thread.Sleep(100);
                Release(shape);
                shape = slide.Shapes[1];
                if (shape.Type == MsoShapeType.msoEmbeddedOLEObject) break;
            }
            AssertEqual(MsoShapeType.msoEmbeddedOLEObject, shape.Type,
                "Direct PowerPoint conversion did not create an embedded OLE object.");
            oleFormat = shape.OLEFormat;
            AssertEqual(FormulaOleContract.ProgId, oleFormat.ProgID,
                "Direct PowerPoint conversion created the wrong OLE class.");
            AssertNear(pictureLeft, shape.Left, 0.2f,
                "Current VSTO picture-to-OLE conversion moved horizontally.");
            AssertNear(pictureTop, shape.Top, 0.2f,
                "Current VSTO picture-to-OLE conversion moved vertically.");
            AssertNear(pictureWidth, shape.Width, 0.2f,
                "Current VSTO picture-to-OLE conversion changed width.");
            AssertNear(pictureHeight, shape.Height, 0.2f,
                "Current VSTO picture-to-OLE conversion changed height.");
            AssertNear(pictureRatio, shape.Width / shape.Height, 0.01f,
                "Current VSTO picture-to-OLE conversion changed aspect ratio.");
            var currentVstoOlePng = Path.Combine(artifactRoot, "targeted-current-vsto-ole.png");
            shape.Export(currentVstoOlePng, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
            var oleInk = AnalyzeDarkPixels(currentVstoOlePng);
            AssertTrue(oleInk.Count > 0, "Current VSTO OLE export contains no formula pixels.");
            var sourceInkRatio = sourceInk.Width / (double)Math.Max(1, sourceInk.Height);
            var oleInkRatio = oleInk.Width / (double)Math.Max(1, oleInk.Height);
            if (Math.Abs(sourceInkRatio - oleInkRatio) / sourceInkRatio > 0.03)
                throw new InvalidOperationException(
                    $"Current VSTO picture-to-OLE visually distorted the formula. " +
                    $"Source ink ratio={sourceInkRatio:F4}, OLE ink ratio={oleInkRatio:F4}.");
            Console.WriteLine(
                $"  Direct PowerPoint conversion completed in "
                + $"{conversionElapsed.TotalSeconds:F2}s without a visible editor; "
                + $"geometry stayed {shape.Left:F3},{shape.Top:F3} {shape.Width:F3}x{shape.Height:F3} pt.");

            Console.WriteLine("[Targeted PowerPoint 3/4] Verifying the OLE verb does not reveal the VisualTeX main window...");
            var baselineWindows = VisibleVisualTeXWindowTitles();
            oleFormat.DoVerb(0);
            WinForms.Application.DoEvents();
            Thread.Sleep(1000);
            AssertNoNewVisibleVisualTeXWindows(
                baselineWindows,
                "PowerPoint OLE default verb");

            Console.WriteLine("[Targeted PowerPoint 4/4] Verifying OLE double-click opens only the formula editor Session...");
            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            var doubleClickHandler = typeof(VisualTeX.PowerPointVsto.ThisAddIn).GetMethod(
                "OnNativeDoubleClick",
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.NonPublic)
                ?? throw new MissingMethodException("PowerPoint double-click callback is missing.");
            doubleClickHandler.Invoke(addIn, null);
            var editSessionId = WaitForNewSession(
                existing,
                "powerpoint",
                TimeSpan.FromSeconds(30));
            var editSession = WaitForUnchangedEditorReady(
                client,
                editSessionId,
                TimeSpan.FromSeconds(10));
            AssertEqual("edit", editSession.Mode,
                "PowerPoint OLE double-click did not create an edit Session.");
            AssertEqual(FormulaOleContract.NativeOleMode, editSession.ObjectMode,
                "PowerPoint OLE double-click changed the object mode.");
            AssertEqual(formulaId, editSession.FormulaId,
                "PowerPoint OLE double-click opened the wrong formula.");
            client.CloseEditorAsync(editSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            final = WaitForTerminal(client, editSessionId, TimeSpan.FromSeconds(30));
            AssertEqual("completed", final.Status,
                final.Error ?? "PowerPoint OLE double-click Session did not complete.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            var path = Path.Combine(artifactRoot, "VisualTeX-Targeted-PowerPoint-46.pptx");
            presentation.SaveAs(
                path,
                PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                MsoTriState.msoFalse);
            Console.WriteLine($"  Saved {path}; direct conversion and OLE double-click routing passed.");
        }
        finally
        {
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            if (installedAddIn is not null)
            {
                try { installedAddIn.Connect = true; } catch { }
            }
            Release(installedAddIn);
            Release(installedAddIns);
            Release(oleFormat);
            Release(shape);
            Release(oleServerKeepAlive);
            Release(slide);
            if (presentation is not null)
            {
                try { presentation.Close(); } catch { }
            }
            Release(presentation);
            if (application is not null)
            {
                try { application.Quit(); } catch { }
            }
            Release(application);
            if (testOleServerProcess is not null)
            {
                try
                {
                    if (!testOleServerProcess.HasExited)
                        testOleServerProcess.Kill();
                }
                catch { }
                testOleServerProcess.Dispose();
            }
            ForceComCleanup();
        }
    }

    private static void RunPowerPointContextSafetyAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? firstPresentation = null;
        PowerPoint.Presentation? secondPresentation = null;
        PowerPoint.Presentation? readOnlyPresentation = null;
        PowerPoint.Slide? firstSlide = null;
        PowerPoint.Slide? secondSlide = null;
        PowerPoint.Slide? otherSlide = null;
        PowerPoint.SlideShowWindow? slideShowWindow = null;
        VisualTeX.PowerPointVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            firstPresentation = application.Presentations.Add(MsoTriState.msoTrue);
            firstSlide = firstPresentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            secondSlide = firstPresentation.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutBlank);
            firstPresentation.Windows[1].Activate();
            application.ActiveWindow.View.GotoSlide(1);
            addIn = new VisualTeX.PowerPointVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);

            Console.WriteLine("[PowerPoint context 1/5] Starting on slide 1, switching to slide 2 before commit...");
            var existing = SnapshotSessionIds();
            addIn.OnNewFormula(new object());
            var sessionId = WaitForNewSession(existing, "powerpoint", TimeSpan.FromSeconds(30));
            var session = client.GetSessionAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            application.ActiveWindow.View.GotoSlide(2);
            Commit(client, session, "block", FormulaOleContract.CrossPlatformPictureMode, "s_1");
            var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "PowerPoint source-slide insertion failed.");
            client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            var sourceSlideCount = firstSlide.Shapes.Count;
            var activeSlideCount = secondSlide.Shapes.Count;
            Console.WriteLine(
                $"  Source slide shapes={sourceSlideCount}; active slide shapes={activeSlideCount}.");
            if (sourceSlideCount != 1 || activeSlideCount != 0)
            {
                var sourceNames = Enumerable.Range(1, sourceSlideCount)
                    .Select(index => firstSlide.Shapes[index].Name)
                    .ToArray();
                var activeNames = Enumerable.Range(1, activeSlideCount)
                    .Select(index => secondSlide.Shapes[index].Name)
                    .ToArray();
                Console.WriteLine(
                    $"  [DISCREPANCY] Source-slide routing failed. "
                    + $"source=[{string.Join(", ", sourceNames)}], active=[{string.Join(", ", activeNames)}].");
            }
            while (firstSlide.Shapes.Count > 0) firstSlide.Shapes[1].Delete();
            while (secondSlide.Shapes.Count > 0) secondSlide.Shapes[1].Delete();

            Console.WriteLine("[PowerPoint context 2/5] Switching presentation before commit and requiring rejection...");
            firstPresentation.Windows[1].Activate();
            application.ActiveWindow.View.GotoSlide(1);
            existing = SnapshotSessionIds();
            addIn.OnNewFormula(new object());
            sessionId = WaitForNewSession(existing, "powerpoint", TimeSpan.FromSeconds(30));
            session = client.GetSessionAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            secondPresentation = application.Presentations.Add(MsoTriState.msoTrue);
            otherSlide = secondPresentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            secondPresentation.Windows[1].Activate();
            Commit(client, session, "block", FormulaOleContract.CrossPlatformPictureMode, "must_not_insert");
            final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
            AssertEqual("failed", final.Status,
                "PowerPoint accepted a formula after the active presentation changed.");
            AssertEqual(0, firstSlide.Shapes.Count,
                "Rejected cross-presentation commit changed the source presentation.");
            AssertEqual(0, otherSlide.Shapes.Count,
                "Rejected cross-presentation commit wrote into the wrong presentation.");
            client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            Console.WriteLine("[PowerPoint context 3/5] Inserting in presentation 2, then Undo and Redo...");
            secondPresentation.Windows[1].Activate();
            application.ActiveWindow.View.GotoSlide(1);
            existing = SnapshotSessionIds();
            addIn.OnNewFormula(new object());
            sessionId = WaitForNewSession(existing, "powerpoint", TimeSpan.FromSeconds(30));
            session = client.GetSessionAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            Commit(client, session, "block", FormulaOleContract.CrossPlatformPictureMode, "u+v=w");
            final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "PowerPoint undo/redo fixture insertion failed.");
            client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            AssertEqual(1, otherSlide.Shapes.Count,
                "PowerPoint undo/redo fixture was not inserted.");
            application.CommandBars.ExecuteMso("Undo");
            Thread.Sleep(300);
            AssertEqual(0, otherSlide.Shapes.Count,
                "PowerPoint Undo did not remove the latest formula.");
            application.CommandBars.ExecuteMso("Redo");
            Thread.Sleep(300);
            AssertEqual(1, otherSlide.Shapes.Count,
                "PowerPoint Redo did not restore the latest formula.");

            Console.WriteLine("[PowerPoint context 4/5] Rejecting formula creation during slide show...");
            slideShowWindow = secondPresentation.SlideShowSettings.Run();
            Thread.Sleep(500);
            existing = SnapshotSessionIds();
            addIn.OnNewFormula(new object());
            Thread.Sleep(800);
            var slideShowSessions = SnapshotSessionIds().Except(existing, StringComparer.Ordinal).ToArray();
            AssertEqual(0, slideShowSessions.Length,
                "PowerPoint created an editor Session during slide show.");
            AssertTrue(!string.IsNullOrWhiteSpace(addIn.DiagnosticLastError),
                "PowerPoint slide-show rejection did not report an error.");
            slideShowWindow.View.Exit();
            Release(slideShowWindow);
            slideShowWindow = null;

            Console.WriteLine("[PowerPoint context 5/5] Reopening read-only and requiring insertion rejection...");
            var path = Path.Combine(artifactRoot, "VisualTeX-PowerPoint-Context-Safety.pptx");
            secondPresentation.SaveAs(
                path,
                PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                MsoTriState.msoFalse);
            secondPresentation.Close();
            Release(otherSlide);
            otherSlide = null;
            Release(secondPresentation);
            secondPresentation = null;
            readOnlyPresentation = application.Presentations.Open(
                path,
                ReadOnly: MsoTriState.msoTrue,
                Untitled: MsoTriState.msoFalse,
                WithWindow: MsoTriState.msoTrue);
            otherSlide = readOnlyPresentation.Slides[1];
            var readOnlyBefore = otherSlide.Shapes.Count;
            existing = SnapshotSessionIds();
            addIn.OnNewFormula(new object());
            Thread.Sleep(800);
            var readOnlySessionIds = SnapshotSessionIds().Except(existing, StringComparer.Ordinal).ToArray();
            if (readOnlySessionIds.Length == 0)
            {
                AssertTrue(!string.IsNullOrWhiteSpace(addIn.DiagnosticLastError),
                    "PowerPoint read-only rejection did not report an error.");
            }
            else
            {
                AssertEqual(1, readOnlySessionIds.Length,
                    "PowerPoint read-only insertion created multiple Sessions.");
                session = client.GetSessionAsync(readOnlySessionIds[0], CancellationToken.None)
                    .GetAwaiter().GetResult();
                Commit(client, session, "block", FormulaOleContract.CrossPlatformPictureMode, "blocked");
                final = WaitForTerminal(client, readOnlySessionIds[0], TimeSpan.FromSeconds(45));
                AssertEqual("failed", final.Status,
                    "PowerPoint read-only insertion unexpectedly completed.");
                client.CloseEditorAsync(readOnlySessionIds[0], CancellationToken.None).GetAwaiter().GetResult();
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            }
            AssertEqual(readOnlyBefore, otherSlide.Shapes.Count,
                "PowerPoint wrote a formula into the read-only presentation.");
            Console.WriteLine(
                $"PowerPoint slide/presentation context, Undo/Redo, slide-show, and read-only safety passed. Artifact: {path}");
        }
        finally
        {
            if (slideShowWindow is not null)
            {
                try { slideShowWindow.View.Exit(); } catch { }
            }
            Release(slideShowWindow);
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            Release(otherSlide);
            Release(secondSlide);
            Release(firstSlide);
            if (readOnlyPresentation is not null)
            {
                try { readOnlyPresentation.Close(); } catch { }
            }
            Release(readOnlyPresentation);
            if (secondPresentation is not null)
            {
                try { secondPresentation.Close(); } catch { }
            }
            Release(secondPresentation);
            if (firstPresentation is not null)
            {
                try { firstPresentation.Close(); } catch { }
            }
            Release(firstPresentation);
            if (application is not null)
            {
                try { application.Quit(); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunPowerPointOleSvgDeleteAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Presentation? reopened = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        VisualTeX.PowerPointVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            Console.WriteLine("[PowerPoint OLE/SVG/delete 1/6] Creating an editable SVG/graphic formula...");
            application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            addIn = new VisualTeX.PowerPointVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);

            var existing = SnapshotSessionIds();
            addIn.OnNewFormula(new object());
            var sessionId = WaitForNewSession(existing, "powerpoint", TimeSpan.FromSeconds(30));
            var session = client.GetSessionAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            Commit(
                client,
                session,
                "block",
                FormulaOleContract.CrossPlatformPictureMode,
                "\\frac{a+b}{c+d}=1",
                renderWidth: 320f,
                renderHeight: 80f,
                baseline: 60f);
            var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status, final.Error ?? "PowerPoint OLE/SVG fixture did not complete.");
            client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            shape = slide.Shapes[1];
            var originalWidth = shape.Width;
            var originalHeight = shape.Height;
            var originalMetadata = DecodePowerPointMetadata(shape)
                ?? throw new InvalidDataException("Initial PowerPoint SVG metadata could not be decoded.");
            var formulaId = originalMetadata.FormulaId;

            Console.WriteLine("[PowerPoint OLE/SVG/delete 2/6] Converting SVG/graphic to native OLE...");
            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            final = WaitForDirectConversion(
                client,
                existing,
                "powerpoint",
                FormulaOleContract.NativeOleMode,
                () => addIn.OnConvertSelected(new object()),
                TimeSpan.FromSeconds(45),
                out _);
            AssertEqual("completed", final.Status, final.Error ?? "PowerPoint SVG-to-OLE conversion failed.");
            Release(shape);
            shape = slide.Shapes[1];
            AssertEqual(MsoShapeType.msoEmbeddedOLEObject, shape.Type,
                "PowerPoint SVG-to-OLE did not create an embedded OLE object.");
            AssertNear(originalWidth, shape.Width, 1.0f,
                "PowerPoint SVG-to-OLE changed formula width.");
            AssertNear(originalHeight, shape.Height, 0.8f,
                "PowerPoint SVG-to-OLE changed formula height.");

            Console.WriteLine("[PowerPoint OLE/SVG/delete 3/6] Exporting OLE back to editable SVG picture...");
            application.ActiveWindow.Activate();
            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            final = WaitForDirectConversion(
                client,
                existing,
                "powerpoint",
                FormulaOleContract.CrossPlatformPictureMode,
                () => addIn.OnExportSelectedAsPicture(new object()),
                TimeSpan.FromSeconds(45),
                out _);
            AssertEqual("completed", final.Status, final.Error ?? "PowerPoint OLE-to-SVG conversion failed.");
            Release(shape);
            shape = slide.Shapes[1];
            AssertTrue(IsPowerPointEditablePictureShape(shape),
                $"PowerPoint OLE-to-SVG returned the wrong shape type: {(int)shape.Type}.");
            AssertNear(originalWidth, shape.Width, 1.0f,
                "PowerPoint OLE-to-SVG changed formula width.");
            AssertNear(originalHeight, shape.Height, 0.8f,
                "PowerPoint OLE-to-SVG changed formula height.");
            AssertTrue(shape.Width > 0 && shape.Height > 0,
                "PowerPoint OLE-to-SVG produced a zero-size formula.");
            var svgMetadata = DecodePowerPointMetadata(shape)
                ?? throw new InvalidDataException("OLE-to-SVG metadata could not be decoded.");
            AssertEqual(formulaId, svgMetadata.FormulaId,
                "OLE-to-SVG changed the formula ID.");
            AssertTrue(svgMetadata.Lines.Any(line => line.Latex.IndexOf("frac", StringComparison.Ordinal) >= 0),
                "OLE-to-SVG lost the LaTeX source.");

            Console.WriteLine("[PowerPoint OLE/SVG/delete 4/6] Exporting the slide and checking visible formula pixels...");
            var pngPath = Path.Combine(artifactRoot, "VisualTeX-PowerPoint-OLE-SVG.png");
            slide.Export(pngPath, "PNG", 960, 540);
            var pixels = AnalyzeDarkPixels(pngPath);
            AssertTrue(pixels.Count >= 40,
                $"PowerPoint OLE-to-SVG slide export is blank or nearly blank ({pixels.Count} dark pixels).");

            Console.WriteLine("[PowerPoint OLE/SVG/delete 5/6] Saving, reopening, and editing the SVG formula...");
            var path = Path.Combine(artifactRoot, "VisualTeX-PowerPoint-OLE-SVG-Delete.pptx");
            presentation.SaveAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoFalse);
            presentation.Close();
            Release(shape);
            shape = null;
            Release(slide);
            slide = null;
            Release(presentation);
            presentation = null;
            reopened = application.Presentations.Open(path, WithWindow: MsoTriState.msoTrue);
            slide = reopened.Slides[1];
            shape = slide.Shapes[1];
            AssertTrue(IsPowerPointEditablePictureShape(shape),
                "Reopened OLE-to-SVG formula is no longer an editable picture/graphic.");
            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            var editSessionId = WaitForNewSession(existing, "powerpoint", TimeSpan.FromSeconds(30));
            var editSession = WaitForUnchangedEditorReady(client, editSessionId, TimeSpan.FromSeconds(12));
            AssertEqual(formulaId, editSession.FormulaId,
                "Reopened SVG edit targeted the wrong formula ID.");
            AssertTrue(editSession.Lines.Any(line => line.Latex.IndexOf("frac", StringComparison.Ordinal) >= 0),
                "Reopened SVG edit lost the LaTeX source.");
            client.CloseEditorAsync(editSessionId, CancellationToken.None).GetAwaiter().GetResult();
            _ = WaitForTerminal(client, editSessionId, TimeSpan.FromSeconds(30));
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            Console.WriteLine("[PowerPoint OLE/SVG/delete 6/6] Deleting the selected formula and saving the empty slide...");
            shape.Select(MsoTriState.msoTrue);
            addIn.OnDeleteSelected(new object());
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            AssertEqual(0, slide.Shapes.Count,
                "PowerPoint delete left the formula shape on the slide.");
            reopened.Save();
            reopened.Close();
            Release(reopened);
            reopened = application.Presentations.Open(path, WithWindow: MsoTriState.msoTrue);
            slide = reopened.Slides[1];
            AssertEqual(0, slide.Shapes.Count,
                "PowerPoint deleted formula reappeared after save/reopen.");
            Console.WriteLine(
                $"PowerPoint OLE-to-SVG, visible export, metadata/source, save/reopen, and delete cleanup passed. Artifact: {path}");
        }
        finally
        {
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            Release(shape);
            Release(slide);
            if (reopened is not null)
            {
                try { reopened.Close(); } catch { }
            }
            Release(reopened);
            if (presentation is not null)
            {
                try { presentation.Close(); } catch { }
            }
            Release(presentation);
            if (application is not null)
            {
                try { application.Quit(); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunPowerPointFontSizeAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Presentation? reopened = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        VisualTeX.PowerPointVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            Console.WriteLine("[PowerPoint font size 1/6] Creating an editable SVG/graphic formula...");
            application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            addIn = new VisualTeX.PowerPointVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);

            var existing = SnapshotSessionIds();
            addIn.OnNewFormula(new object());
            var sessionId = WaitForNewSession(existing, "powerpoint", TimeSpan.FromSeconds(30));
            var session = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(client, session, "block", FormulaOleContract.CrossPlatformPictureMode, "x+y=z");
            var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status, final.Error ?? "PowerPoint font-size fixture did not complete.");
            client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            shape = slide.Shapes[1];
            AssertTrue(IsPowerPointEditablePictureShape(shape),
                $"Font-size fixture is not an editable picture/graphic. Actual type: {(int)shape.Type}.");
            shape.Select(MsoTriState.msoTrue);
            var originalWidth = shape.Width;
            var originalHeight = shape.Height;
            var originalCenterX = shape.Left + shape.Width / 2f;
            var originalCenterY = shape.Top + shape.Height / 2f;

            Console.WriteLine("[PowerPoint font size 2/6] Setting the picture formula to 36 pt...");
            addIn.OnFormulaFontSizeChanged(null!, "36");
            Release(shape);
            shape = slide.Shapes[1];
            var width36 = shape.Width;
            var height36 = shape.Height;
            AssertTrue(width36 > originalWidth && height36 > originalHeight,
                $"36 pt did not enlarge the picture formula: {originalWidth:F2}x{originalHeight:F2} -> {width36:F2}x{height36:F2}.");
            AssertNear(originalWidth / originalHeight, width36 / height36, 0.05f,
                "Picture font-size change distorted the formula aspect ratio.");
            AssertNear(originalCenterX, shape.Left + shape.Width / 2f, 0.5f,
                "Picture font-size change moved the formula horizontally.");
            AssertNear(originalCenterY, shape.Top + shape.Height / 2f, 0.5f,
                "Picture font-size change moved the formula vertically.");
            var metadata36 = FormulaMetadataCodec.Decode(shape.AlternativeText)
                ?? throw new InvalidDataException("36 pt picture metadata could not be decoded.");
            AssertNear(36f, (float)(metadata36.FontSizePt ?? 0), 0.1f,
                "Picture metadata did not store 36 pt.");

            Console.WriteLine("[PowerPoint font size 3/6] Exercising decrease and increase presets...");
            shape.Select(MsoTriState.msoTrue);
            addIn.OnDecreaseFormulaFontSize(null!);
            Release(shape);
            shape = slide.Shapes[1];
            var decreasedMetadata = FormulaMetadataCodec.Decode(shape.AlternativeText)
                ?? throw new InvalidDataException("Decreased picture metadata could not be decoded.");
            AssertTrue((decreasedMetadata.FontSizePt ?? 0) < 36 && shape.Width < width36,
                "PowerPoint decrease did not reduce semantic and physical formula size.");
            shape.Select(MsoTriState.msoTrue);
            addIn.OnIncreaseFormulaFontSize(null!);
            Release(shape);
            shape = slide.Shapes[1];
            var restoredMetadata = FormulaMetadataCodec.Decode(shape.AlternativeText)
                ?? throw new InvalidDataException("Restored picture metadata could not be decoded.");
            AssertNear(36f, (float)(restoredMetadata.FontSizePt ?? 0), 0.1f,
                "Decrease/increase did not restore 36 pt.");
            AssertNear(width36, shape.Width, 0.8f,
                "Decrease/increase did not restore picture width.");
            AssertNear(height36, shape.Height, 0.5f,
                "Decrease/increase did not restore picture height.");

            Console.WriteLine("[PowerPoint font size 4/6] Converting the 36 pt picture to OLE...");
            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            final = WaitForDirectConversion(
                client,
                existing,
                "powerpoint",
                FormulaOleContract.NativeOleMode,
                () => addIn.OnConvertSelected(new object()),
                TimeSpan.FromSeconds(45),
                out _);
            AssertEqual("completed", final.Status, final.Error ?? "36 pt picture-to-OLE conversion failed.");
            Release(shape);
            shape = slide.Shapes[1];
            AssertEqual(MsoShapeType.msoEmbeddedOLEObject, shape.Type,
                "36 pt conversion did not create an embedded OLE object.");
            AssertNear(width36, shape.Width, 1.0f,
                "Picture-to-OLE conversion changed 36 pt width.");
            AssertNear(height36, shape.Height, 0.6f,
                "Picture-to-OLE conversion changed 36 pt height.");

            Console.WriteLine("[PowerPoint font size 5/6] Setting the OLE formula to 48 pt...");
            application.ActiveWindow.Activate();
            application.ActiveWindow.View.GotoSlide(1);
            shape.Select(MsoTriState.msoTrue);
            WinForms.Application.DoEvents();
            Thread.Sleep(200);
            var enabledBefore48 = addIn.GetFormulaFontSizeEnabled(null!);
            var textBefore48 = addIn.GetFormulaFontSizeText(null!);
            var metadataBefore48 = DecodePowerPointMetadata(shape);
            Console.WriteLine(
                $"  OLE before 48 pt: {shape.Width:F2}x{shape.Height:F2}, enabled={enabledBefore48}, "
                + $"control='{textBefore48}', metadata={metadataBefore48?.FontSizePt?.ToString("F1", System.Globalization.CultureInfo.InvariantCulture) ?? "null"}.");
            addIn.OnFormulaFontSizeChanged(null!, "48");
            WinForms.Application.DoEvents();
            Thread.Sleep(200);
            Release(shape);
            shape = slide.Shapes[1];
            var width48 = shape.Width;
            var height48 = shape.Height;
            var metadata48 = DecodePowerPointMetadata(shape);
            shape.Select(MsoTriState.msoTrue);
            var textAfter48 = addIn.GetFormulaFontSizeText(null!);
            Console.WriteLine(
                $"  OLE after 48 pt: {width48:F2}x{height48:F2}, control='{textAfter48}', "
                + $"outerMetadata={metadata48?.FontSizePt?.ToString("F1", System.Globalization.CultureInfo.InvariantCulture) ?? "not exposed"}, "
                + $"diagnostic='{addIn.DiagnosticLastError}'.");
            AssertTrue(height48 > height36,
                $"48 pt did not increase the OLE formula height: {height36:F2} -> {height48:F2}; "
                + $"enabled={enabledBefore48}, controlBefore='{textBefore48}', controlAfter='{textAfter48}', diagnostic='{addIn.DiagnosticLastError}'.");
            AssertEqual("48", textAfter48,
                "PowerPoint OLE font-size control did not report 48 pt after resizing.");
            var ratio36 = width36 / height36;
            var ratio48 = width48 / height48;
            AssertNear(ratio36, ratio48, 0.1f,
                $"PowerPoint OLE font-size change distorted aspect ratio; "
                + $"width {width36:F2} -> {width48:F2}, height {height36:F2} -> {height48:F2}.");

            Console.WriteLine("[PowerPoint font size 6/6] Saving, reopening, and rechecking OLE size metadata...");
            var path = Path.Combine(artifactRoot, "VisualTeX-PowerPoint-Font-Size.pptx");
            presentation.SaveAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoFalse);
            presentation.Close();
            Release(shape);
            shape = null;
            Release(slide);
            slide = null;
            Release(presentation);
            presentation = null;
            reopened = application.Presentations.Open(path, WithWindow: MsoTriState.msoTrue);
            slide = reopened.Slides[1];
            shape = slide.Shapes[1];
            AssertEqual(MsoShapeType.msoEmbeddedOLEObject, shape.Type,
                "Reopened 48 pt formula is no longer OLE.");
            AssertNear(width48, shape.Width, 1.0f,
                "48 pt OLE width changed after save/reopen.");
            AssertNear(height48, shape.Height, 0.6f,
                "48 pt OLE height changed after save/reopen.");
            shape.Select(MsoTriState.msoTrue);
            AssertTrue(addIn.GetFormulaFontSizeEnabled(null!),
                "PowerPoint font-size controls were disabled for reopened OLE.");
            AssertEqual("48", addIn.GetFormulaFontSizeText(null!),
                "48 pt semantic size was lost after save/reopen.");
            Console.WriteLine(
                $"PowerPoint font-size acceptance passed: picture {originalWidth:F1}x{originalHeight:F1} -> "
                + $"36 pt {width36:F1}x{height36:F1}; OLE 48 pt {width48:F1}x{height48:F1}; save/reopen stable.");
        }
        finally
        {
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            Release(shape);
            Release(slide);
            if (reopened is not null)
            {
                try { reopened.Close(); } catch { }
            }
            Release(reopened);
            if (presentation is not null)
            {
                try { presentation.Close(); } catch { }
            }
            Release(presentation);
            if (application is not null)
            {
                try { application.Quit(); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunPowerPoint(
        VisualTeXSessionClient client,
        string artifactRoot,
        bool stopAfterPictureEdit = false)
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        PowerPoint.OLEFormat? oleFormat = null;
        VisualTeX.PowerPointVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            Console.WriteLine("[PowerPoint 1/10] Starting PowerPoint and creating a formula Session...");
            application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            addIn = new VisualTeX.PowerPointVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);
            var existing = SnapshotSessionIds();
            addIn.OnNewFormula(new object());
            var sessionId = WaitForNewSession(existing, "powerpoint", TimeSpan.FromSeconds(30));
            var session = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("crossPlatformPicture", session.ObjectMode,
                "PowerPoint must use the stable editable picture path by default.");

            Console.WriteLine("[PowerPoint 2/10] Committing the initial picture formula...");
            Commit(client, session, "block", "crossPlatformPicture", "x+y=z");
            var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status, final.Error ?? "PowerPoint Session did not complete.");
            client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();

            Console.WriteLine("[PowerPoint 3/10] Checking the inserted editable picture...");
            AssertEqual(1, slide.Shapes.Count, "PowerPoint should contain one formula shape.");
            shape = slide.Shapes[1];
            AssertTrue(IsPowerPointEditablePictureShape(shape),
                $"PowerPoint formula must be an editable picture/graphic, not an OLE placeholder. Actual type: {(int)shape.Type}.");
            AssertNear(120f, shape.Width, 0.5f, "PowerPoint formula width is incorrect.");
            AssertNear(24f, shape.Height, 0.5f, "PowerPoint formula height is incorrect.");
            AssertNear(5f, shape.Width / shape.Height, 0.05f,
                "PowerPoint formula aspect ratio is distorted.");
            ReportPowerPointMetadata("initial", shape);

            Console.WriteLine("[PowerPoint create reset] Creating again while the previous formula remains selected...");
            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            addIn.OnNewFormula(new object());
            var blankCreateSessionId = WaitForNewSession(
                existing,
                "powerpoint",
                TimeSpan.FromSeconds(30));
            var blankCreateSession = client.GetSessionAsync(
                    blankCreateSessionId,
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("create", blankCreateSession.Mode,
                "PowerPoint New Formula reused the previous formula as an edit Session.");
            AssertEqual(1, blankCreateSession.Lines.Count,
                "PowerPoint New Formula should start with exactly one empty line.");
            AssertEqual(string.Empty, blankCreateSession.Lines[0].Latex,
                "PowerPoint New Formula inherited LaTeX from the selected previous formula.");
            if (string.Equals(blankCreateSession.FormulaId, final.FormulaId, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "PowerPoint New Formula reused the selected formulaId.");
            Commit(client, blankCreateSession, "block", "crossPlatformPicture", "u=v");
            var blankCreateFinal = WaitForTerminal(
                client,
                blankCreateSessionId,
                TimeSpan.FromSeconds(45));
            AssertEqual("completed", blankCreateFinal.Status,
                blankCreateFinal.Error ?? "Second blank PowerPoint create Session did not complete.");
            client.CloseEditorAsync(blankCreateSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            AssertEqual(2, slide.Shapes.Count,
                "PowerPoint second create did not insert an independent formula.");
            slide.Shapes[2].Delete();
            Release(shape);
            shape = slide.Shapes[1];

            Console.WriteLine("[PowerPoint 4/10] Closing an unchanged edit and waiting for the add-in to unlock...");
            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            var unchangedSessionId = WaitForNewSession(existing, "powerpoint", TimeSpan.FromSeconds(30));
            var unchangedSession = WaitForUnchangedEditorReady(
                client,
                unchangedSessionId,
                TimeSpan.FromSeconds(10));
            AssertEqual(false, unchangedSession.Dirty,
                "PowerPoint unchanged edit Session became dirty before closing.");
            client.CloseEditorAsync(unchangedSessionId, CancellationToken.None).GetAwaiter().GetResult();
            final = WaitForTerminal(client, unchangedSessionId, TimeSpan.FromSeconds(30));
            AssertEqual("completed", final.Status,
                final.Error ?? "PowerPoint unchanged edit did not complete after closing the window.");
            AssertEqual(false, final.Dirty,
                "PowerPoint unchanged edit was incorrectly marked dirty.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            Console.WriteLine("[PowerPoint 5/10] Reopening immediately and editing through the Ribbon callback...");
            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            var editSessionId = WaitForNewSession(existing, "powerpoint", TimeSpan.FromSeconds(30));
            var editSession = client.GetSessionAsync(editSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("edit", editSession.Mode,
                "PowerPoint edit button did not create an edit Session.");
            AssertEqual("crossPlatformPicture", editSession.ObjectMode,
                "PowerPoint picture edit changed the object mode.");
            Commit(
                client,
                editSession,
                "block",
                "crossPlatformPicture",
                "x+y+z+a+b+c=d",
                renderWidth: 320f,
                renderHeight: 32f,
                baseline: 24f);
            final = WaitForTerminal(client, editSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "PowerPoint edit Session did not complete.");
            client.CloseEditorAsync(editSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            Release(shape);
            shape = slide.Shapes[1];
            AssertTrue(IsPowerPointEditablePictureShape(shape),
                $"PowerPoint edit unexpectedly changed the editable picture/graphic object type. Actual type: {(int)shape.Type}.");
            AssertNear(240f, shape.Width, 0.8f,
                "PowerPoint edited formula retained the old width.");
            AssertNear(24f, shape.Height, 0.5f,
                "PowerPoint edited formula height is incorrect.");
            AssertNear(10f, shape.Width / shape.Height, 0.08f,
                "PowerPoint edited formula is compressed into the old aspect ratio.");
            ReportPowerPointMetadata("edited", shape);

            Console.WriteLine("[PowerPoint 6/10] Exercising the double-click edit callback...");
            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            var doubleClickHandler = typeof(VisualTeX.PowerPointVsto.ThisAddIn).GetMethod(
                "OnNativeDoubleClick",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                ?? throw new MissingMethodException("PowerPoint double-click callback is missing.");
            doubleClickHandler.Invoke(addIn, null);
            var doubleClickSessionId = WaitForNewSession(
                existing,
                "powerpoint",
                TimeSpan.FromSeconds(30));
            var doubleClickSession = client.GetSessionAsync(
                    doubleClickSessionId,
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("edit", doubleClickSession.Mode,
                "PowerPoint double-click did not create an edit Session.");
            Commit(
                client,
                doubleClickSession,
                "block",
                "crossPlatformPicture",
                "x+y+z+a+b+c=d",
                renderWidth: 320f,
                renderHeight: 32f,
                baseline: 24f,
                dirty: false);
            final = WaitForTerminal(client, doubleClickSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "PowerPoint double-click Session did not complete.");
            client.CloseEditorAsync(doubleClickSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            if (stopAfterPictureEdit)
            {
                var picturePath = Path.Combine(
                    artifactRoot,
                    "VisualTeX-PowerPoint-Picture-Edit.pptx");
                presentation.SaveAs(
                    picturePath,
                    PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                    MsoTriState.msoFalse);
                Console.WriteLine(
                    $"[PowerPoint picture/edit] Saved {picturePath}; independent create, unchanged close, "
                    + "Ribbon edit resize, aspect ratio, metadata, and double-click callback passed.");
                return;
            }

            Console.WriteLine("[PowerPoint 7/10] Converting the selected formula to native OLE...");
            application.ActiveWindow.Activate();
            application.ActiveWindow.View.GotoSlide(1);
            shape.Select(MsoTriState.msoTrue);
            WinForms.Application.DoEvents();
            Thread.Sleep(150);
            existing = SnapshotSessionIds();
            addIn.OnConvertSelected(new object());
            var convertSessionId = WaitForNewSession(
                existing,
                "powerpoint",
                TimeSpan.FromSeconds(30),
                () => string.IsNullOrWhiteSpace(addIn.DiagnosticLastError)
                    ? null
                    : Convert.ToBase64String(
                        Encoding.UTF8.GetBytes(addIn.DiagnosticLastError)));
            var convertSession = client.GetSessionAsync(convertSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("nativeOle", convertSession.ObjectMode,
                "PowerPoint convert command did not request native OLE.");
            Commit(
                client,
                convertSession,
                "block",
                "nativeOle",
                "x+y+z+a+b+c=d",
                renderWidth: 320f,
                renderHeight: 32f,
                baseline: 24f,
                dirty: false);
            final = WaitForTerminal(client, convertSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "PowerPoint OLE conversion did not complete.");
            client.CloseEditorAsync(convertSessionId, CancellationToken.None).GetAwaiter().GetResult();
            Release(shape);
            shape = slide.Shapes[1];
            AssertEqual(MsoShapeType.msoEmbeddedOLEObject, shape.Type,
                "PowerPoint convert command did not create an embedded OLE object.");
            oleFormat = shape.OLEFormat;
            AssertEqual(FormulaOleContract.ProgId, oleFormat.ProgID,
                "PowerPoint converted to the wrong OLE class.");
            AssertNear(240f, shape.Width, 0.8f,
                "PowerPoint OLE conversion changed the formula width.");
            AssertNear(24f, shape.Height, 0.5f,
                "PowerPoint OLE conversion changed the formula height.");
            Console.WriteLine("[PowerPoint 8/10] Exercising the converted OLE show verb...");
            oleFormat.DoVerb(0);
            Release(oleFormat);
            oleFormat = null;

            Console.WriteLine("[PowerPoint 9/10] Exporting the final slide and checking visible content...");
            var presentationPath = Path.Combine(artifactRoot, "VisualTeX-PowerPoint-Flow.pptx");
            presentation.SaveAs(
                presentationPath,
                PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                MsoTriState.msoFalse);
            var slidePng = Path.Combine(artifactRoot, "VisualTeX-PowerPoint-Flow.png");
            slide.Export(slidePng, "PNG", 960, 540);
            var preview = AnalyzeDarkPixels(slidePng);
            if (preview.Count < 40)
                throw new InvalidDataException(
                    $"PowerPoint export is blank or nearly blank ({preview.Count} dark pixels).");
            if (preview.Width > 140 || preview.Height > 45)
                throw new InvalidDataException(
                    $"PowerPoint OLE export still resembles the placeholder cache " +
                    $"({preview.Width}x{preview.Height} dark-pixel bounds).");
            Console.WriteLine(
                $"[PowerPoint 10/10] Saved {presentationPath}; preview has {preview.Count} dark pixels " +
                $"inside {preview.Width}x{preview.Height} bounds.");
        }
        finally
        {
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            Release(oleFormat);
            Release(shape);
            Release(slide);
            if (presentation is not null)
            {
                try { presentation.Close(); } catch { }
            }
            Release(presentation);
            if (application is not null)
            {
                try { application.Quit(); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void AppendToLastWordOmmlAndSelect(
        Word.Document document,
        string suffix)
    {
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? insertion = null;
        Word.Range? selectedRange = null;
        try
        {
            maths = document.OMaths;
            if (maths.Count == 0)
                throw new InvalidDataException("Word document contains no OMML equation to edit.");
            math = maths[maths.Count];
            math.Linearize();
            insertion = math.Range.Duplicate;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            insertion.InsertBefore(suffix);
            math.BuildUp();
            selectedRange = math.Range;
            selectedRange.Select();
        }
        finally
        {
            Release(selectedRange);
            Release(insertion);
            Release(math);
            Release(maths);
        }
    }

    private static void Commit(
        VisualTeXSessionClient client,
        OfficeSessionDocument session,
        string displayMode,
        string objectMode,
        string latex,
        float renderWidth = ExportWidth,
        float renderHeight = ExportHeight,
        float baseline = ExportBaseline,
        bool dirty = true,
        bool numbered = false,
        string? mathMl = null)
    {
        var lineId = session.Lines.First().Id;
        var svg = CreateSvg(renderWidth, renderHeight);
        var exportResult = new Dictionary<string, object?>
        {
            ["svg"] = svg,
            ["svgBase64"] = "data:image/svg+xml;base64," +
                Convert.ToBase64String(Encoding.UTF8.GetBytes(svg)),
            ["pngBase64"] = CreatePngDataUrl(latex, renderWidth, renderHeight),
            ["width"] = renderWidth,
            ["height"] = renderHeight,
            ["baseline"] = baseline,
        };
        if (!string.IsNullOrWhiteSpace(mathMl)) exportResult["mathMl"] = mathMl;
        var patch = new Dictionary<string, object>
        {
            ["lines"] = new[]
            {
                new Dictionary<string, object>
                {
                    ["id"] = lineId,
                    ["latex"] = latex,
                },
            },
            ["activeLineId"] = lineId,
            ["codeFormat"] = "latex",
            ["displayMode"] = displayMode,
            ["objectMode"] = objectMode,
            ["numbered"] = numbered,
            ["exportWidth"] = renderWidth,
            ["exportHeight"] = renderHeight,
            ["exportResult"] = exportResult,
            ["dirty"] = dirty,
            ["status"] = "committing",
        };
        client.PatchAsync(session.Id, patch, CancellationToken.None).GetAwaiter().GetResult();
    }

    private static void WaitForWordInlineShapeCount(
        Word.Document document,
        int expectedCount,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        var lastCount = -1;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
            Word.InlineShapes? shapes = null;
            try
            {
                shapes = document.InlineShapes;
                lastCount = shapes.Count;
                if (lastCount == expectedCount) return;
            }
            finally { Release(shapes); }
        }
        throw new TimeoutException(
            $"Expected {expectedCount} Word inline shapes, last count was {lastCount}.");
    }

    private static string WaitForWordNativeReferenceResult(
        Word.Document document,
        string expectedResult,
        string? expectedCode,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        var lastResult = string.Empty;
        var lastCode = string.Empty;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
            Word.Fields? fields = null;
            try
            {
                fields = document.Fields;
                for (var index = 1; index <= fields.Count; index++)
                {
                    Word.Field? field = null;
                    Word.Range? code = null;
                    Word.Range? result = null;
                    try
                    {
                        field = fields[index];
                        if (field.Type != Word.WdFieldType.wdFieldRef) continue;
                        code = field.Code;
                        var codeText = (code.Text ?? string.Empty).Trim();
                        if (!string.IsNullOrWhiteSpace(expectedCode)
                            && !string.Equals(
                                NormalizeFieldCode(codeText),
                                NormalizeFieldCode(expectedCode),
                                StringComparison.OrdinalIgnoreCase))
                            continue;
                        field.Update();
                        result = field.Result;
                        lastCode = codeText;
                        lastResult = (result.Text ?? string.Empty).Trim();
                        if (string.Equals(lastResult, expectedResult, StringComparison.Ordinal))
                            return codeText;
                    }
                    finally
                    {
                        Release(result);
                        Release(code);
                        Release(field);
                    }
                }
            }
            finally { Release(fields); }
        }
        throw new TimeoutException(
            $"Word native REF field did not become {expectedResult}; " +
            $"last result was [{lastResult}], last code was [{lastCode}].");
    }

    private static string NormalizeFieldCode(string value) =>
        string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private static bool DocumentTextContains(Word.Document document, string value)
    {
        Word.Range? content = null;
        try
        {
            content = document.Content;
            return (content.Text ?? string.Empty).IndexOf(value, StringComparison.Ordinal) >= 0;
        }
        finally { Release(content); }
    }

    private static bool WordBookmarkExists(Word.Document document, string name)
    {
        Word.Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            return bookmarks.Exists(name);
        }
        finally { Release(bookmarks); }
    }

    private static Word.InlineShape WaitForWordShapeMode(
        Word.Document document,
        bool nativeOle,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
            Word.InlineShapes? shapes = null;
            Word.InlineShape? candidate = null;
            Word.OLEFormat? format = null;
            try
            {
                shapes = document.InlineShapes;
                if (shapes.Count != 1) continue;
                candidate = shapes[1];
                var isNativeOle = false;
                if (candidate.Type is Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                    or Word.WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                {
                    try
                    {
                        format = candidate.OLEFormat;
                        isNativeOle = string.Equals(
                            format.ProgID,
                            FormulaOleContract.ProgId,
                            StringComparison.OrdinalIgnoreCase);
                    }
                    catch
                    {
                        isNativeOle = false;
                    }
                }
                if (isNativeOle == nativeOle)
                {
                    var result = candidate;
                    candidate = null;
                    return result;
                }
            }
            finally
            {
                Release(format);
                Release(candidate);
                Release(shapes);
            }
        }
        throw new TimeoutException(nativeOle
            ? "Word formula did not become a VisualTeX native OLE object."
            : "Word formula did not become a cross-platform picture.");
    }

    private static OfficeSessionDocument WaitForUnchangedEditorReady(
        VisualTeXSessionClient client,
        string sessionId,
        TimeSpan timeout)
    {
        var startedAt = DateTime.UtcNow;
        var deadline = startedAt + timeout;
        OfficeSessionDocument? session = null;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
            session = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            if (session.Dirty)
                throw new InvalidOperationException(
                    $"Session {sessionId} became dirty before the user changed the formula.");
            if (session.Status is "completed" or "failed" or "cancelled")
                throw new InvalidOperationException(
                    $"Session {sessionId} reached {session.Status} before the unchanged editor was closed: {session.Error}");
            // A pristine editor is allowed to remain in `created`; no autosave is
            // necessary until the user changes something. Wait long enough for
            // the WebView and its close handlers to mount before closing it.
            if (DateTime.UtcNow - startedAt >= TimeSpan.FromSeconds(3)
                && session.Status is "created" or "editing")
                return session;
        }
        throw new TimeoutException(
            $"Session {sessionId} was not ready for unchanged close; " +
            $"last status was {session?.Status ?? "unknown"}.");
    }

    private static OfficeSessionDocument WaitForDirectConversion(
        VisualTeXSessionClient client,
        HashSet<string> existingSessionIds,
        string expectedHost,
        string expectedObjectMode,
        Action command,
        TimeSpan timeout,
        out TimeSpan elapsed,
        Func<string?>? errorProvider = null)
    {
        var baselineWindows = VisibleVisualTeXWindowTitles();
        var stopwatch = Stopwatch.StartNew();
        command();
        var sessionId = WaitForNewSession(
            existingSessionIds,
            expectedHost,
            TimeSpan.FromSeconds(30),
            errorProvider);
        var deadline = DateTime.UtcNow + timeout;
        OfficeSessionDocument? session = null;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
            AssertNoNewVisibleVisualTeXWindows(
                baselineWindows,
                $"{expectedHost} direct format conversion");
            session = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual(
                expectedObjectMode,
                session.ObjectMode,
                $"{expectedHost} direct conversion requested the wrong object mode.");
            if (session.Status is "completed" or "failed" or "cancelled")
            {
                stopwatch.Stop();
                elapsed = stopwatch.Elapsed;
                return session;
            }
        }
        stopwatch.Stop();
        elapsed = stopwatch.Elapsed;
        throw new TimeoutException(
            $"Direct {expectedHost} conversion Session {sessionId} did not finish; "
            + $"last status was {session?.Status ?? "unknown"}.");
    }

    private static HashSet<string> VisibleVisualTeXWindowTitles()
    {
        var titles = new HashSet<string>(StringComparer.Ordinal);
        EnumWindows((windowHandle, _) =>
        {
            if (!IsWindowVisible(windowHandle)) return true;
            GetWindowThreadProcessId(windowHandle, out var processId);
            if (processId == 0) return true;
            try
            {
                using var process = Process.GetProcessById((int)processId);
                if (!string.Equals(
                        process.ProcessName,
                        "visualtex",
                        StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch
            {
                return true;
            }

            var length = GetWindowTextLength(windowHandle);
            if (length <= 0) return true;
            var text = new StringBuilder(length + 1);
            GetWindowText(windowHandle, text, text.Capacity);
            var title = text.ToString().Trim();
            if (title.Length > 0) titles.Add(title);
            return true;
        }, IntPtr.Zero);
        return titles;
    }

    private static void AssertNoNewVisibleVisualTeXWindows(
        HashSet<string> baseline,
        string stage)
    {
        var added = VisibleVisualTeXWindowTitles()
            .Except(baseline, StringComparer.Ordinal)
            .ToArray();
        if (added.Length > 0)
            throw new InvalidDataException(
                $"{stage} unexpectedly opened a visible VisualTeX window: "
                + string.Join(" | ", added));
    }

    private static OfficeSessionDocument WaitForTerminal(
        VisualTeXSessionClient client,
        string sessionId,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        OfficeSessionDocument? session = null;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
            session = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            if (session.Status is "completed" or "failed" or "cancelled") return session;
        }
        throw new TimeoutException(
            $"Session {sessionId} did not finish; last status was {session?.Status ?? "unknown"}.");
    }

    private static void WaitForAddInIdle(object addIn, TimeSpan timeout)
    {
        var field = addIn.GetType().GetField(
            "_operationGate",
            System.Reflection.BindingFlags.Instance |
            System.Reflection.BindingFlags.NonPublic)
            ?? throw new MissingFieldException("Office add-in operation gate is missing.");
        var gate = field.GetValue(addIn) as SemaphoreSlim
            ?? throw new InvalidOperationException("Office add-in operation gate is unavailable.");
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            if (gate.CurrentCount == 1) return;
            Thread.Sleep(25);
        }
        throw new TimeoutException("Office add-in did not return to the idle state.");
    }

    private static HashSet<string> SnapshotSessionIds()
    {
        Directory.CreateDirectory(SessionRoot);
        return Directory.EnumerateDirectories(SessionRoot)
            .Select(Path.GetFileName)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Cast<string>()
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    private static string WaitForNewSession(
        HashSet<string> existing,
        string expectedHost,
        TimeSpan timeout,
        Func<string?>? errorProvider = null)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            var diagnosticError = errorProvider?.Invoke();
            if (!string.IsNullOrWhiteSpace(diagnosticError))
                throw new InvalidOperationException(diagnosticError);
            Thread.Sleep(100);
            foreach (var directory in Directory.EnumerateDirectories(SessionRoot)
                         .OrderByDescending(Directory.GetLastWriteTimeUtc))
            {
                var id = Path.GetFileName(directory);
                if (string.IsNullOrWhiteSpace(id) || existing.Contains(id)) continue;
                var sessionPath = Path.Combine(directory, "session.json");
                if (!File.Exists(sessionPath)) continue;
                var json = File.ReadAllText(sessionPath);
                if (json.IndexOf($"\"host\": \"{expectedHost}\"", StringComparison.OrdinalIgnoreCase) >= 0
                    || json.IndexOf($"\"host\":\"{expectedHost}\"", StringComparison.OrdinalIgnoreCase) >= 0)
                    return id;
            }
        }
        throw new TimeoutException($"No new {expectedHost} Office Session appeared.");
    }

    private static string CreateSvg(float width, float height) =>
        $"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 {width:F0} {height:F0}\">" +
        "<path fill=\"#111111\" d=\"" +
        "M4 5 L10 5 L18 14 L26 5 L32 5 L21 17 L33 29 L27 29 L18 20 L9 29 L3 29 L15 17 Z " +
        "M48 14 L60 14 L60 2 L66 2 L66 14 L78 14 L78 20 L66 20 L66 32 L60 32 L60 20 L48 20 Z " +
        "M94 5 L100 5 L108 14 L116 5 L122 5 L111 17 L123 29 L117 29 L108 20 L99 29 L93 29 L105 17 Z" +
        "\"/></svg>";

    private static string CreatePngDataUrl(string latex, float width, float height)
    {
        var pixelWidth = Math.Max(32, (int)Math.Ceiling(width * 2));
        var pixelHeight = Math.Max(24, (int)Math.Ceiling(height * 2));
        using var bitmap = new Bitmap(pixelWidth, pixelHeight, PixelFormat.Format32bppArgb);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.Transparent);
        using var font = new Font("Cambria Math", 28f, FontStyle.Regular, GraphicsUnit.Pixel);
        using var brush = new SolidBrush(Color.Black);
        graphics.DrawString(latex, font, brush, new PointF(2, 4));
        using var stream = new MemoryStream();
        bitmap.Save(stream, ImageFormat.Png);
        return "data:image/png;base64," + Convert.ToBase64String(stream.ToArray());
    }

    private static bool IsPowerPointEditablePictureShape(PowerPoint.Shape shape)
    {
        const int MsoGraphic = 28;
        return shape.Type == MsoShapeType.msoPicture
            || (int)shape.Type == MsoGraphic;
    }

    private static FormulaMetadata? DecodePowerPointMetadata(PowerPoint.Shape shape)
    {
        var decoded = FormulaMetadataCodec.Decode(shape.AlternativeText ?? string.Empty);
        if (decoded is not null) return decoded;
        PowerPoint.Tags? tags = null;
        try
        {
            tags = shape.Tags;
            string value = string.Empty;
            try { value = tags["VisualTeXMetadata"] ?? string.Empty; } catch { }
            return FormulaMetadataCodec.Decode(value);
        }
        finally { Release(tags); }
    }

    private static void ReportPowerPointMetadata(string stage, PowerPoint.Shape shape)
    {
        var alternativeText = shape.AlternativeText ?? string.Empty;
        var decodedAlternative = FormulaMetadataCodec.Decode(alternativeText);
        string tagValue = string.Empty;
        PowerPoint.Tags? tags = null;
        try
        {
            tags = shape.Tags;
            try { tagValue = tags["VisualTeXMetadata"] ?? string.Empty; } catch { }
        }
        finally { Release(tags); }
        Console.WriteLine(
            $"  {stage} metadata: alt={alternativeText.Length} " +
            $"altDecoded={decodedAlternative is not null} tag={tagValue.Length} " +
            $"tagDecoded={FormulaMetadataCodec.Decode(tagValue) is not null}");
    }

    private static (int Count, int Width, int Height) AnalyzeDarkPixels(string path)
    {
        using var bitmap = new Bitmap(path);
        var count = 0;
        var minimumX = bitmap.Width;
        var minimumY = bitmap.Height;
        var maximumX = -1;
        var maximumY = -1;
        for (var y = 0; y < bitmap.Height; y++)
        for (var x = 0; x < bitmap.Width; x++)
        {
            var pixel = bitmap.GetPixel(x, y);
            if (pixel.A < 16) continue;
            if (pixel.R >= 120 && pixel.G >= 120 && pixel.B >= 120) continue;
            count++;
            minimumX = Math.Min(minimumX, x);
            minimumY = Math.Min(minimumY, y);
            maximumX = Math.Max(maximumX, x);
            maximumY = Math.Max(maximumY, y);
        }
        return maximumX < minimumX || maximumY < minimumY
            ? (0, 0, 0)
            : (count, maximumX - minimumX + 1, maximumY - minimumY + 1);
    }

    private static void AssertTrue(bool condition, string message)
    {
        if (!condition) throw new InvalidDataException(message);
    }

    private static void AssertNear(float expected, float actual, float tolerance, string message)
    {
        if (Math.Abs(expected - actual) > tolerance)
            throw new InvalidOperationException(
                $"{message} Expected {expected:F3}, actual {actual:F3}.");
    }

    private static void AssertEqual<T>(T expected, T actual, string message)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new InvalidOperationException(
                $"{message} Expected {expected}, actual {actual}.");
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            try { Marshal.ReleaseComObject(value); } catch { }
        }
    }

    private static void ForceComCleanup()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
        Thread.Sleep(500);
    }

    private sealed class TeeTextWriter : TextWriter
    {
        private readonly TextWriter _primary;
        private readonly TextWriter _secondary;

        public TeeTextWriter(TextWriter primary, TextWriter secondary)
        {
            _primary = primary;
            _secondary = secondary;
        }

        public override Encoding Encoding => _primary.Encoding;

        public override void Write(char value)
        {
            _primary.Write(value);
            _secondary.Write(value);
        }

        public override void Write(string? value)
        {
            _primary.Write(value);
            _secondary.Write(value);
        }

        public override void WriteLine(string? value)
        {
            _primary.WriteLine(value);
            _secondary.WriteLine(value);
        }

        public override void Flush()
        {
            _primary.Flush();
            _secondary.Flush();
        }
    }
}
