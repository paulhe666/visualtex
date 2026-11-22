using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// Word answers Range.Information(wdWithInTable) from its page layout: asked once per
// formula of a batch (LaTeX redraw, bulk import), right after each write, it lays
// the document out again up to the probe. That made a 200-formula Word redraw take
// minutes (a few seconds in WPS). The per-formula table checks of the shared XML
// import and of the paragraph snapshot use the table structure in Word and keep
// Information for WPS only. The test host cannot load Office's interop types, so
// this checks the sources.
public sealed class WordLayoutFreeTableCheckTests
{
    private static string SourcesRoot([CallerFilePath] string path = "") =>
        Path.GetFullPath(Path.Combine(Path.GetDirectoryName(path)!, ".."));

    private static readonly Regex InformationCall = new(
        @"get_Information\(\s*WdInformation\.wdWithInTable\s*\)");

    private static readonly Regex MethodSignature = new(
        @"static\s+[\w<>?\[\]]+\s+(?<name>\w+)\s*\(");

    // Methods that only WPS reaches (WpsXmlInsertion.InsertInWps and its helpers).
    private static readonly string[] WpsOnlyMethods = { "EmptyParagraphAt" };

    [Theory]
    [InlineData("VisualTeX.WordVsto", "WpsXmlInsertion.cs")]
    [InlineData("VisualTeX.WordVsto", "WordParagraphSnapshot.cs")]
    public void PerFormulaTableChecksAskWordForTheTableStructure(string project, string file)
    {
        var source = File.ReadAllText(Path.Combine(SourcesRoot(), project, file));
        var offenders = new List<string>();
        foreach (Match call in InformationCall.Matches(source))
        {
            var before = source.Substring(0, call.Index);
            var method = MethodSignature.Matches(before).Cast<Match>().LastOrDefault()?.Groups["name"].Value ?? "?";
            var expressionStart = Math.Max(before.LastIndexOf(';'), before.LastIndexOf('{'));
            var expression = before.Substring(expressionStart + 1);
            var wpsBranch = Regex.IsMatch(expression, @"IsWps\s*\?\s*\w+\.$");
            if (!wpsBranch && !WpsOnlyMethods.Contains(method))
                offenders.Add($"{file}: {method}");
        }
        Assert.Empty(offenders);
        // and Word gets the structural check (Range.Tables) instead
        Assert.Matches(new Regex(@"IsWhollyInsideTable\(|\.Tables;"), source);
    }
}
