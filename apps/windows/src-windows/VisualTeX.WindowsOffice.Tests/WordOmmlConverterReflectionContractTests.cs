using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// The PowerPoint add-in reaches the shared Word converter by reflection
// (PowerPointNativeOmmlKernel.ResolveConverterMethods), so the compiler cannot
// catch a signature change that breaks native PowerPoint equations. The test
// host cannot load Office's interop types, so this compares the sources: the
// exact parameter list the kernel asks for must still be declared.
public sealed class WordOmmlConverterReflectionContractTests
{
    private static string SourcesRoot([CallerFilePath] string path = "") =>
        Path.GetFullPath(Path.Combine(Path.GetDirectoryName(path)!, ".."));

    [Fact]
    public void PowerPointKernelFindsTheDirectOmmlReplacement()
    {
        var kernel = File.ReadAllText(Path.Combine(
            SourcesRoot(), "VisualTeX.PowerPointVsto", "PowerPointNativeOmmlKernel.cs"));
        var converter = File.ReadAllText(Path.Combine(
            SourcesRoot(), "VisualTeX.WindowsOffice.VstoShared", "WordOmmlConverter.cs"));

        Assert.Matches(
            new Regex(
                @"""ReplaceWithPreparedOmmlDirect"",[^;]*?types: new\[\]\s*\{\s*typeof\(Word\.Document\),\s*typeof\(Word\.Range\),\s*typeof\(string\),\s*typeof\(bool\),\s*typeof\(string\),\s*typeof\(bool\),\s*\}",
                RegexOptions.Singleline),
            kernel);
        Assert.Matches(
            new Regex(
                @"internal static Range ReplaceWithPreparedOmmlDirect\(\s*Document targetDocument,\s*Range targetRange,\s*string omml,\s*bool display,\s*string\? mathFontName = null,\s*bool allowWordNativeInlinePromotion = false\)",
                RegexOptions.Singleline),
            converter);
    }
}
