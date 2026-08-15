using System.Runtime.InteropServices;

namespace VisualTeX.WordVsto;

internal static class MathTypeNativePreviewRenderer
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

    internal sealed class Result : IDisposable
    {
        public string WmfPath { get; set; } = string.Empty;
        public float WidthPt { get; set; }
        public float HeightPt { get; set; }
        public int WordPosition { get; set; }

        public void Dispose()
        {
            if (string.IsNullOrWhiteSpace(WmfPath)) return;
            try { File.Delete(WmfPath); } catch { }
        }
    }

    internal static bool TryRender(byte[] mtef, string outputDirectory, out Result result)
    {
        result = new Result();
        if (mtef is null || mtef.Length == 0) return false;
        var mathPage = ResolveMathPagePath();
        if (mathPage is null) return false;

        Directory.CreateDirectory(outputDirectory);
        var outputPath = Path.Combine(outputDirectory, $"mathtype-native-{Guid.NewGuid():N}.wmf");
        IntPtr module = IntPtr.Zero;
        MtTermApiDelegate? term = null;
        var initialized = false;
        try
        {
            module = LoadLibraryW(mathPage);
            if (module == IntPtr.Zero) return false;
            var init = GetDelegate<MtInitApiDelegate>(module, "MTInitAPI");
            term = GetDelegate<MtTermApiDelegate>(module, "MTTermAPI");
            var transform = GetDelegate<MtXFormEqnDelegate>(module, "MTXFormEqn");
            var getDimension = GetDelegate<MtGetLastDimensionDelegate>(module, "MTGetLastDimension");
            if (init is null || term is null || transform is null || getDimension is null) return false;

            var visibleBefore = CountVisibleMathTypeWindows();
            var initResult = init(MtInitLaunchAsNeeded, 8);
            initialized = initResult >= 0;
            if (!initialized) return false;

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
                outputPath,
                ref dims);
            if (status != MtOk || !File.Exists(outputPath) || new FileInfo(outputPath).Length <= 22)
                return false;

            var width32 = getDimension(1);
            var height32 = getDimension(2);
            var baseline32 = getDimension(3);
            if (width32 <= 0 || height32 <= 0 || baseline32 < 0
                || width32 == MtError || height32 == MtError || baseline32 == MtError)
                return false;

            if (CountVisibleMathTypeWindows() > visibleBefore)
                return false;

            result = new Result
            {
                WmfPath = outputPath,
                WidthPt = width32 / 32f,
                HeightPt = height32 / 32f,
                WordPosition = -(int)Math.Round(baseline32 / 32d, MidpointRounding.AwayFromZero),
            };
            outputPath = string.Empty;
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            if (initialized && term is not null)
            {
                try { term(); } catch { }
            }
            if (module != IntPtr.Zero) FreeLibrary(module);
            if (!string.IsNullOrWhiteSpace(outputPath))
            {
                try { File.Delete(outputPath); } catch { }
            }
        }
    }

    private static string? ResolveMathPagePath()
    {
        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        if (string.IsNullOrWhiteSpace(programFilesX86)) return null;
        var architecture = Environment.Is64BitProcess ? "64" : "32";
        var path = Path.Combine(programFilesX86, "MathType", "MathPage", architecture, "MathPage.wll");
        return File.Exists(path) ? path : null;
    }

    private static T? GetDelegate<T>(IntPtr module, string name) where T : class
    {
        var address = GetProcAddress(module, name);
        if (address == IntPtr.Zero) return null;
        return Marshal.GetDelegateForFunctionPointer(address, typeof(T)) as T;
    }

    private static int CountVisibleMathTypeWindows()
    {
        var count = 0;
        try
        {
            foreach (var process in System.Diagnostics.Process.GetProcessesByName("MathType"))
            {
                try
                {
                    process.Refresh();
                    if (!process.HasExited && process.MainWindowHandle != IntPtr.Zero) count++;
                }
                catch { }
                finally { process.Dispose(); }
            }
        }
        catch { }
        return count;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr LoadLibraryW(string path);

    [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
    private static extern IntPtr GetProcAddress(IntPtr module, string name);

    [DllImport("kernel32.dll")]
    private static extern bool FreeLibrary(IntPtr module);
}
