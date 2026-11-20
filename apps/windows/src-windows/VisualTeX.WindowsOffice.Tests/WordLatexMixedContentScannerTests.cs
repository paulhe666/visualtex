using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WindowsOffice.Tests;

// LaTeX in ordinary Word prose: prices, stray dollar signs and paragraph
// breaks must never shift which text the redraw/import treats as a formula.
public sealed class WordLatexMixedContentScannerTests
{
    private static string[] Redraw(string source) =>
        WordBulkImportParser.FindFormulaSpans(source).Select(span => span.Latex).ToArray();

    private static string[] Import(string source) =>
        WordBulkImportParser.Parse(source, WordBulkSourceFormat.Markdown, WordBulkFormulaObjectMode.Omml)
            .Blocks.SelectMany(block => block.Runs).Where(run => run.IsFormula).Select(run => run.Latex).ToArray();

    public static TheoryData<string, string[]> MixedContent => new()
    {
        // prices are not formulas, and do not pair with the real formula after them
        { "价格 $5，折后 $3。公式 $x^2$。", new[] { "x^2" } },
        { "US$100 和 $x+y$ 都有", new[] { "x+y" } },
        { "区间 $5-$10，变量 $a$", new[] { "a" } },
        { "It costs $5 and $10, while $n$ is free.", new[] { "n" } },
        { "Total: $ 5 and $ 10", Array.Empty<string>() },
        // an inline formula never spans a paragraph: the stray $ stays text
        { "开头 $x\r下一段 y$ 结尾 $z$", new[] { "z" } },
        { "\\(a\rb\\) 与 \\(c\\)", new[] { "c" } },
        // spaced but mathematical bodies are formulas
        { "设 $ x^2 $ 与 $ a $，以及 $x $", new[] { "x^2", "a", "x" } },
        { "$5$ 个苹果，$n$个人", new[] { "5", "n" } },
        // escapes and adjacency keep working
        { "\\$ 5 和 $a$", new[] { "a" } },
        { "相邻 $a$$b$。", new[] { "a", "b" } },
        // multi-line display formulas are fine; a blank line ends a display formula
        { "$$\ra=b\r$$", new[] { "a=b" } },
        { "未闭合 $$ a\r\r下一段 $$b$$", new[] { "b" } },
    };

    [Theory]
    [MemberData(nameof(MixedContent))]
    public void RedrawFindsOnlyTheFormulas(string source, string[] expected) =>
        Assert.Equal(expected, Redraw(source));

    public static TheoryData<string, string[]> ImportLines => new()
    {
        { "价格 $5，折后 $3。公式 $x^2$。", new[] { "x^2" } },
        { "It costs $5 and $10, while $n$ is free.", new[] { "n" } },
        { "设 $ x^2 $ 与 $ a $", new[] { "x^2", "a" } },
        { "相邻 $a$$b$。", new[] { "a", "b" } },
    };

    [Theory]
    [MemberData(nameof(ImportLines))]
    public void ImportFindsOnlyTheFormulas(string source, string[] expected) =>
        Assert.Equal(expected, Import(source));

    [Fact]
    public void RedrawSpanCoordinatesSurviveSkippedPrices()
    {
        const string source = "It costs $5 and $10, while $n$ is free.";
        var span = Assert.Single(WordBulkImportParser.FindFormulaSpans(source));
        Assert.Equal("$n$", source.Substring(span.Start, span.Length));
    }
}
