using System.Collections.Concurrent;
using System.Diagnostics;
using VisualTeX.WordVsto;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // Covers the preview caller boundary only. No native editor activation and
    // no unrelated formatting/numbering corpus is needed for a launch regression.
    private static void RunWordMathTypeHeadlessAcceptance(string artifactRoot)
    {
        var baseline = SnapshotInteractiveMathTypeIdentities();
        var visibleWindows = new ConcurrentDictionary<int, string>();
        using var stop = new CancellationTokenSource();
        var observer = Task.Run(() =>
        {
            while (!stop.IsCancellationRequested)
            {
                foreach (var process in Process.GetProcessesByName("MathType"))
                {
                    using (process)
                    {
                        try
                        {
                            if (!baseline.ContainsKey(process.Id) && process.MainWindowHandle != IntPtr.Zero)
                                visibleWindows[process.Id] = process.MainWindowTitle;
                        }
                        catch (InvalidOperationException) { }
                    }
                }
                stop.Token.WaitHandle.WaitOne(25);
            }
        });
        var previousAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        try
        {
            var svgPath = Path.Combine(artifactRoot, "insert.svg");
            File.WriteAllText(svgPath,
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"96\" viewBox=\"0 0 240 96\"><text x=\"4\" y=\"64\" font-size=\"48\">x+1</text></svg>");
            var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 96);
            RunWordMathTypeInlineCreateAcceptance(artifactRoot, emfPath);

            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", "1");
            using var host = new WordPerformanceHost(documentPath: null);
            AssertTrue(MathTypeNativePreviewRenderer.IsAvailable, "The private render service is not running.");
            const string token = "VT_HEADLESS_CONVERT";
            host.Document.Content.Text = "before " + token + " after\r";
            var range = InsertPureNativeOmml(host.Document, token, "(x+1)/(y-2)");
            try { range.Select(); }
            finally { Release(range); }
            host.AddIn.OnConvertOmmlToMathTypeSelection(new object());
            WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(15));
            AssertEqual(0, host.Document.OMaths.Count, "OMML conversion did not finish.");
            AssertEqual(1, CountMathTypeOleShapes(host.Document), "Conversion did not create exactly one MathType OLE.");
            AssertEveryMathTypeProgId(host.Document);
            host.Save(Path.Combine(artifactRoot, "headless-conversion.docx"));
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", previousAcceptance);
            stop.Cancel();
            observer.GetAwaiter().GetResult();
        }
        AssertEqual(0, visibleWindows.Count,
            "MathType native editor appeared: " + string.Join(", ", visibleWindows.Values));
        var after = SnapshotInteractiveMathTypeIdentities();
        AssertTrue(baseline.Count == after.Count && baseline.All(pair =>
                after.TryGetValue(pair.Key, out var created) && created == pair.Value),
            "Word insertion/conversion changed interactive MathType process identities.");
        Console.WriteLine("PASS Word insert/save/reopen and OMML conversion; no native MathType editor window.");
    }
}
