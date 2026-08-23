using System.Runtime.InteropServices;
using System.Text.Json;

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
        try
        {
            var mathPage = ResolveMathPagePath(request.MathTypeServerPath);
            if (mathPage is null)
            {
                response.Error = "MathType MathPage.wll is not installed.";
                return WriteResponse(request.ResultPath, response, 3);
            }

            module = LoadLibraryW(mathPage);
            if (module == IntPtr.Zero)
            {
                response.Error = "MathType MathPage.wll could not be loaded.";
                return WriteResponse(request.ResultPath, response, 3);
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

            initialized = init(MtInitLaunchAsNeeded, 8) >= 0;
            if (!initialized)
            {
                response.Error = "MathType MathPage API initialization failed.";
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

    private static string? ResolveMathPagePath(string? mathTypeServerPath)
    {
        var overridePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_MATHTYPE_MATHPAGE_PATH");
        if (!string.IsNullOrWhiteSpace(overridePath))
        {
            var expanded = Environment.ExpandEnvironmentVariables(
                overridePath.Trim().Trim('"'));
            if (File.Exists(expanded)) return expanded;
        }

        var architecture = Environment.Is64BitProcess ? "64" : "32";
        var candidates = new List<string>();
        if (!string.IsNullOrWhiteSpace(mathTypeServerPath))
        {
            var expandedServer = Environment.ExpandEnvironmentVariables(
                mathTypeServerPath.Trim().Trim('"'));
            var installRoot = Path.GetDirectoryName(expandedServer);
            if (!string.IsNullOrWhiteSpace(installRoot))
                candidates.Add(Path.Combine(
                    installRoot,
                    "MathPage",
                    architecture,
                    "MathPage.wll"));
        }
        foreach (var root in new[]
                 {
                     Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                     Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                 })
        {
            if (!string.IsNullOrWhiteSpace(root))
                candidates.Add(Path.Combine(
                    root,
                    "MathType",
                    "MathPage",
                    architecture,
                    "MathPage.wll"));
        }
        return candidates.FirstOrDefault(File.Exists);
    }

    private static T? GetDelegate<T>(IntPtr module, string name) where T : class
    {
        var address = GetProcAddress(module, name);
        if (address == IntPtr.Zero) return null;
        return Marshal.GetDelegateForFunctionPointer(address, typeof(T)) as T;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr LoadLibraryW(string path);

    [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
    private static extern IntPtr GetProcAddress(IntPtr module, string name);

    [DllImport("kernel32.dll")]
    private static extern bool FreeLibrary(IntPtr module);
}
