using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// The PowerPoint add-in reaches the shared Word converter by reflection
// (PowerPointNativeOmmlKernel.ResolveTransform and ReadNativeMathMl), so the
// compiler cannot catch a signature change that breaks native PowerPoint
// equations. The test host cannot load Office's interop types, so this compares
// the sources: the exact parameter lists the kernel asks for must still be
// declared.
public sealed class WordOmmlConverterReflectionContractTests
{
    private static string SourcesRoot([CallerFilePath] string path = "") =>
        Path.GetFullPath(Path.Combine(Path.GetDirectoryName(path)!, ".."));

    [Fact]
    public void PowerPointKernelFindsTheSharedWriterAndReader()
    {
        var kernel = File.ReadAllText(Path.Combine(
            SourcesRoot(), "VisualTeX.PowerPointVsto", "PowerPointNativeOmmlKernel.cs"));
        var converter = File.ReadAllText(Path.Combine(
            SourcesRoot(), "VisualTeX.WindowsOffice.VstoShared", "WordOmmlConverter.cs"));

        Assert.Matches(
            new Regex(
                @"""TransformMathMlToOmml"",[^;]*?types: new\[\]\s*\{\s*typeof\(string\)\s*\}",
                RegexOptions.Singleline),
            kernel);
        Assert.Matches(
            new Regex(@"internal static string TransformMathMlToOmml\(string mathMl\)"),
            converter);
        Assert.Matches(
            new Regex(
                @"""TransformOmmlToMathMl"",[^;]*?types: new\[\]\s*\{\s*typeof\(string\),\s*typeof\(bool\)\s*\}",
                RegexOptions.Singleline),
            kernel);
        Assert.Matches(
            new Regex(@"internal static string TransformOmmlToMathMl\(string wordOpenXml, bool display\)"),
            converter);
    }
}
