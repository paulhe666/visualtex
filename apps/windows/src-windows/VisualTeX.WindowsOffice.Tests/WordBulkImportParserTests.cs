using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordBulkImportParserTests
{
    public static IEnumerable<object[]> HeadingHierarchyCases()
    {
        using var fixture = System.Text.Json.JsonDocument.Parse(File.ReadAllText(
            Path.Combine(AppContext.BaseDirectory, "document-import-heading-hierarchy.json")));
        foreach (var item in fixture.RootElement.EnumerateArray())
            yield return new object[] { item.GetProperty("source").GetString()!,
                string.Join(",", item.GetProperty("levels").EnumerateArray().Select(level => level.GetInt32())) };
    }

    [Theory]
    [MemberData(nameof(HeadingHierarchyCases))]
    public void LatexHeadingHierarchyPreservesArticleAndBookStructure(string source, string expected)
    {
        var document = WordBulkImportParser.Parse(source, WordBulkSourceFormat.Latex, WordBulkFormulaObjectMode.Omml);
        Assert.Equal(expected, string.Join(",", document.Blocks
            .Where(block => block.Kind == WordBulkBlockKind.Heading).Select(block => block.Level)));
    }

    [Fact]
    public void MarkdownProducesNativeTextAndIndependentInlineAndDisplayFormulas()
    {
        const string source = """
            # 示例标题

            这是 **粗体**、*斜体* 与行内公式 $E=mc^2$。

            $$
            \int_0^1 x^2\,\mathrm{d}x=\frac13
            $$

            - 第一项含公式 $a+b$
            - 第二项
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(5, document.Blocks.Count);
        Assert.Equal(3, document.FormulaCount);
        Assert.Equal(2, document.InlineFormulaCount);
        Assert.Equal(1, document.DisplayFormulaCount);
        Assert.Equal(WordBulkBlockKind.Heading, document.Blocks[0].Kind);
        Assert.Contains(document.Blocks[1].Runs, run => !run.IsFormula && run.Bold && run.Text == "粗体");
        Assert.Contains(document.Blocks[1].Runs, run => !run.IsFormula && run.Italic && run.Text == "斜体");
        Assert.Contains(document.Blocks[1].Runs, run => run.IsFormula && run.Latex == "E=mc^2");
        Assert.Equal(WordBulkBlockKind.DisplayFormula, document.Blocks[2].Kind);
        Assert.Equal("\\int_0^1 x^2\\,\\mathrm{d}x=\\frac13", document.Blocks[2].Runs.Single().Latex);
        Assert.All(document.Blocks.Skip(3), block => Assert.Equal(WordBulkBlockKind.Bullet, block.Kind));
        Assert.Empty(document.Warnings);
    }

    [Fact]
    public void LatexDocumentStripsPreambleAndConvertsSectionsListsAndMath()
    {
        const string source = """
            \documentclass{article}
            \usepackage{amsmath}
            \begin{document}
            \section{理论背景}
            普通文字与 \textbf{重点}，以及 \(\alpha+\beta\)。
            \begin{equation}
            F=ma
            \end{equation}
            \begin{enumerate}
            \item 第一项 $x_1$
            \item 第二项
            \end{enumerate}
            \end{document}
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Auto,
            WordBulkFormulaObjectMode.Ole);

        Assert.Equal(WordBulkSourceFormat.Latex, document.SourceFormat);
        Assert.Equal(WordBulkFormulaObjectMode.Ole, document.FormulaObjectMode);
        Assert.Equal(3, document.FormulaCount);
        Assert.Equal(2, document.InlineFormulaCount);
        Assert.Equal(1, document.DisplayFormulaCount);
        Assert.Equal(WordBulkBlockKind.Heading, document.Blocks[0].Kind);
        Assert.Equal(1, document.Blocks[0].Level);
        Assert.Contains(document.Blocks[1].Runs, run => run.Bold && run.Text == "重点");
        Assert.Contains(document.Blocks[1].Runs, run => run.IsFormula && run.Latex == "\\alpha+\\beta");
        Assert.Equal("F=ma", document.Blocks[2].Runs.Single().Latex);
        Assert.All(document.Blocks.Skip(3), block => Assert.Equal(WordBulkBlockKind.Numbered, block.Kind));
        Assert.DoesNotContain(document.Blocks.SelectMany(block => block.Runs), run => run.Text.Contains("documentclass"));
    }

    [Fact]
    public void ProofOnlyFragmentAutoDetectsLatexAndPreservesHeadingAndBody()
    {
        const string source = """
            \begin{proof}[谱展开]
            对任意 $N$，部分和收敛。
            \end{proof}
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Auto,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(WordBulkSourceFormat.Latex, document.SourceFormat);
        Assert.Equal(2, document.Blocks.Count);
        var heading = Assert.Single(document.Blocks[0].Runs);
        Assert.True(heading.Bold);
        Assert.False(heading.Italic);
        Assert.Equal("证明（谱展开）：", heading.Text);
        var body = Assert.Single(document.Blocks[1].Runs, run =>
            !run.IsFormula && run.Text.Contains("对任意", StringComparison.Ordinal));
        Assert.False(body.Bold);
        Assert.False(body.Italic);
        Assert.Contains(document.Blocks[1].Runs, run => run.IsFormula && run.Latex == "N");
        Assert.DoesNotContain(document.Blocks.SelectMany(block => block.Runs), run => run.Text.Contains("begin{proof}", StringComparison.Ordinal));
    }

    [Fact]
    public void EquationStarPhysicalSourceLinesRemainOneFormulaExpression()
    {
        const string source = """
            \begin{equation*}
            \lim_{N\to +\infty}\bigl\|\,|\psi\rangle-|s_N\rangle\,\bigr\|
            =
            \lim_{N\to +\infty}
            \left\|\,|\psi\rangle-\sum_{i=1}^{N} c_i |u_i\rangle\,\right\|
            =0.
            \end{equation*}
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Omml);

        var formula = Assert.Single(document.Blocks).Runs.Single();
        Assert.True(formula.IsFormula);
        Assert.Equal("block", formula.DisplayMode);
        Assert.DoesNotContain('\n', formula.Latex);
        Assert.Equal(
            "\\lim_{N\\to +\\infty}\\bigl\\|\\,|\\psi\\rangle-|s_N\\rangle\\,\\bigr\\| = \\lim_{N\\to +\\infty} \\left\\|\\,|\\psi\\rangle-\\sum_{i=1}^{N} c_i |u_i\\rangle\\,\\right\\| =0.",
            formula.Latex);
    }

    [Fact]
    public void AlignPhysicalSourceLinesKeepEveryExplicitMathRow()
    {
        const string source = """
            \begin{align*}
            a &= b+c \\
            d &= e-f \\
            g &= h
            \end{align*}
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Ole);

        var formula = Assert.Single(document.Blocks).Runs.Single();
        Assert.DoesNotContain('\n', formula.Latex);
        Assert.StartsWith("\\begin{aligned}", formula.Latex, StringComparison.Ordinal);
        Assert.EndsWith("\\end{aligned}", formula.Latex, StringComparison.Ordinal);
        Assert.Contains("a &= b+c \\\\ d &= e-f \\\\ g &= h", formula.Latex, StringComparison.Ordinal);
    }

    [Fact]
    public void LatexCjkInlineMathWhitespaceFollowsTexSemantics()
    {
        const string source =
            "中文 $x=1$ 中文；English $y=2$ words；"
            + "中文\\ $z=3$\\ 中文；中文~$q=4$~中文；"
            + "中文\\quad$r=5$\\qquad中文";

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Omml);

        var serialized = string.Concat(
            document.Blocks.SelectMany(block => block.Runs).Select(run =>
                run.IsFormula ? $"<{run.Latex}>" : run.Text));
        Assert.Equal(
            "中文<x=1>中文；English <y=2> words；"
            + "中文\u00A0<z=3>\u00A0中文；中文\u00A0<q=4>\u00A0中文；"
            + "中文\u00A0<r=5>\u00A0\u00A0中文",
            serialized);
    }

    [Fact]
    public void LatexVerbatimDocumentExamplesDoNotTruncateLargeFragment()
    {
        var source = new System.Text.StringBuilder();
        source.AppendLine("\\section*{综合 LaTeX 正文示例}");
        source.AppendLine();
        source.AppendLine("这是一段不含 \\verb|\\documentclass|、导言区和宏包加载命令的正文代码，可直接放入已有文档的 \\verb|\\begin{document}| 与 \\verb|\\end{document}| 之间。");
        source.AppendLine("\\textbf{\\Large 常见公式与排版语法综合练习}");
        source.AppendLine();
        for (var number = 1; number <= 20; number++)
        {
            source.AppendLine("\\section{" + number + ". 综合公式}");
            source.AppendLine(
                "正文 $x_{" + number + "}=a_{" + number + "}+b_{" + number + "}$、"
                + "$y_{" + number + "}=c_{" + number + "}-d_{" + number + "}$ 与 "
                + "$z_{" + number + "}=e_{" + number + "}f_{" + number + "}$。");
            source.AppendLine("\\[");
            source.AppendLine("A_{" + number + "}=\\frac{" + number + "}{" + (number + 1) + "}");
            source.AppendLine("\\]");
            source.AppendLine("\\[");
            source.AppendLine("B_{" + number + "}=\\sum_{k=1}^{" + number + "}k");
            source.AppendLine("\\]");
        }
        source.AppendLine("\\section{21. 综合案例与表达规范}");
        source.AppendLine("最后一节没有额外公式，但必须完整保留。");
        source.AppendLine("\\section*{排版检查清单}");
        source.AppendLine("正文结束。");

        var document = WordBulkImportParser.Parse(
            source.ToString(),
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(23, document.Blocks.Count(block => block.Kind == WordBulkBlockKind.Heading));
        Assert.Equal(100, document.FormulaCount);
        Assert.Equal(60, document.InlineFormulaCount);
        Assert.Equal(40, document.DisplayFormulaCount);
        Assert.Contains(
            document.Blocks.SelectMany(block => block.Runs),
            run => !run.IsFormula && run.Code && run.Text == "\\documentclass");
        Assert.Contains(
            document.Blocks.SelectMany(block => block.Runs),
            run => !run.IsFormula && run.Code && run.Text == "\\begin{document}");
        Assert.Contains(
            document.Blocks.SelectMany(block => block.Runs),
            run => !run.IsFormula && run.Code && run.Text == "\\end{document}");
        Assert.Contains(
            document.Blocks.SelectMany(block => block.Runs),
            run => !run.IsFormula && run.Bold && run.Text.Contains("常见公式与排版语法综合练习", StringComparison.Ordinal));
        Assert.DoesNotContain(
            document.Blocks.SelectMany(block => block.Runs),
            run => !run.IsFormula && run.Text.Contains("\\Large", StringComparison.Ordinal));
        Assert.Contains(
            document.Blocks.SelectMany(block => block.Runs),
            run => !run.IsFormula && run.Text.Contains("最后一节没有额外公式", StringComparison.Ordinal));
        Assert.Contains(
            document.Blocks.SelectMany(block => block.Runs),
            run => !run.IsFormula && run.Text.Contains("正文结束", StringComparison.Ordinal));
        Assert.Empty(document.Warnings);
    }

    [Fact]
    public void LatexDocumentBoundariesIgnoreCommentsAndLiteralRegions()
    {
        const string source = """
            % \begin{document}
            \documentclass{article}
            \begin{verbatim}
            \begin{document}
            \end{document}
            \end{verbatim}
            \begin{document}
            正文前 \verb|\end{document}| 正文后 $x=1$。
            \end{document}
            尾部不应导入。
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(1, document.FormulaCount);
        Assert.Contains(
            document.Blocks.SelectMany(block => block.Runs),
            run => !run.IsFormula && run.Text.Contains("正文前", StringComparison.Ordinal));
        Assert.Contains(
            document.Blocks.SelectMany(block => block.Runs),
            run => !run.IsFormula && run.Code && run.Text == "\\end{document}");
        Assert.DoesNotContain(
            document.Blocks.SelectMany(block => block.Runs),
            run => !run.IsFormula && run.Text.Contains("尾部不应导入", StringComparison.Ordinal));
        Assert.Empty(document.Warnings);
    }

    [Fact]
    public void LatexAmsDisplayEnvironmentsRemainRenderableAndEditable()
    {
        const string source = """
            \begin{document}
            \begin{align}
            a+b &= c \\
            d &= e+f
            \end{align}
            \begin{gather*}
            x=1 \\
            y=2
            \end{gather*}
            \begin{multline}
            p+q+r \\
            = s+t
            \end{multline}
            \end{document}
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(3, document.DisplayFormulaCount);
        var formulas = document.Blocks
            .Where(block => block.Kind == WordBulkBlockKind.DisplayFormula)
            .Select(block => block.Runs.Single().Latex.Replace("\r", string.Empty))
            .ToArray();
        Assert.Equal(
            "\\begin{aligned}a+b &= c \\\\ d &= e+f\\end{aligned}",
            formulas[0]);
        Assert.Equal(
            "\\begin{gathered}x=1 \\\\ y=2\\end{gathered}",
            formulas[1]);
        Assert.Equal(
            "\\begin{gathered}p+q+r \\\\ = s+t\\end{gathered}",
            formulas[2]);
    }

    [Fact]
    public void MarkdownAndLatexPreserveExtendedIntegralCommandsVerbatim()
    {
        const string markdown = """
            闭合曲面积分 $\oiint_{\Sigma} a\,\mathrm{d}S$。

            $$
            \oiiint_V f\,\mathrm{d}V+\intclockwise_C g\,\mathrm{d}s
            $$
            """;
        const string latex = """
            \documentclass{article}
            \begin{document}
            行内 \(\varointclockwise_C f\,\mathrm{d}s\)。
            \[
            \ointctrclockwise_C g\,\mathrm{d}s+\intctrclockwise_C h\,\mathrm{d}s
            \]
            \end{document}
            """;

        var markdownDocument = WordBulkImportParser.Parse(
            markdown,
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Ole);
        var latexDocument = WordBulkImportParser.Parse(
            latex,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Omml);

        var markdownFormulas = markdownDocument.Blocks
            .SelectMany(block => block.Runs)
            .Where(run => run.IsFormula)
            .Select(run => run.Latex)
            .ToArray();
        var latexFormulas = latexDocument.Blocks
            .SelectMany(block => block.Runs)
            .Where(run => run.IsFormula)
            .Select(run => run.Latex)
            .ToArray();

        Assert.Equal(2, markdownFormulas.Length);
        Assert.Contains("\\oiint_{\\Sigma}", markdownFormulas[0], StringComparison.Ordinal);
        Assert.Contains("\\oiiint_V", markdownFormulas[1], StringComparison.Ordinal);
        Assert.Contains("\\intclockwise_C", markdownFormulas[1], StringComparison.Ordinal);
        Assert.Equal(2, latexFormulas.Length);
        Assert.Contains("\\varointclockwise_C", latexFormulas[0], StringComparison.Ordinal);
        Assert.Contains("\\ointctrclockwise_C", latexFormulas[1], StringComparison.Ordinal);
        Assert.Contains("\\intctrclockwise_C", latexFormulas[1], StringComparison.Ordinal);
        Assert.Empty(markdownDocument.Warnings);
        Assert.Empty(latexDocument.Warnings);
    }

    [Fact]
    public void EscapedDollarRemainsNativeText()
    {
        var document = WordBulkImportParser.Parse(
            "价格写作 \\$5，而公式是 $x+1$。",
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(1, document.FormulaCount);
        var runs = document.Blocks.Single().Runs;
        Assert.Contains(runs, run => !run.IsFormula && run.Text.Contains("$5"));
        Assert.Contains(runs, run => run.IsFormula && run.Latex == "x+1");
    }

    [Fact]
    public void UnclosedDisplayFormulaIsImportedWithWarning()
    {
        var document = WordBulkImportParser.Parse(
            "正文\n\n$$\na+b",
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(1, document.DisplayFormulaCount);
        Assert.Single(document.Warnings);
        Assert.Contains("缺少结束标记", document.Warnings[0]);
    }

    [Fact]
    public void LatexNestedListsQuotesAndLiteralCodeRemainStructured()
    {
        const string source = """
            \documentclass{book}
            \begin{document}
            \chapter{结构测试}
            \begin{quote}
            第一行包含 \textbf{重点}。
            第二行包含行内公式 $x+y$。
            \end{quote}
            \begin{itemize}
            \item 外层项目
            \begin{enumerate}
            \item 内层编号项目 $n=1$
            \end{enumerate}
            \end{itemize}
            \begin{verbatim}
            value_1 = 20 % 这里不是注释
            \end{verbatim}
            \end{document}
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Auto,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(WordBulkSourceFormat.Latex, document.SourceFormat);
        Assert.Equal(WordBulkBlockKind.Heading, document.Blocks[0].Kind);
        Assert.Equal(1, document.Blocks[0].Level);
        var quote = Assert.Single(document.Blocks, block => block.Kind == WordBulkBlockKind.Quote);
        Assert.Contains(quote.Runs, run => !run.IsFormula && run.Bold && run.Text == "重点");
        Assert.Contains(quote.Runs, run => run.IsFormula && run.Latex == "x+y");
        var lists = document.Blocks
            .Where(block => block.Kind is WordBulkBlockKind.Bullet or WordBulkBlockKind.Numbered)
            .ToArray();
        Assert.Equal(2, lists.Length);
        Assert.Equal(WordBulkBlockKind.Bullet, lists[0].Kind);
        Assert.Equal(0, lists[0].Level);
        Assert.Equal(WordBulkBlockKind.Numbered, lists[1].Kind);
        Assert.Equal(1, lists[1].Level);
        var code = Assert.Single(document.Blocks, block => block.Kind == WordBulkBlockKind.Code);
        Assert.True(code.Runs.Single().Code);
        Assert.Contains("value_1 = 20 % 这里不是注释", code.Runs.Single().Text);
        Assert.Empty(document.Warnings);
    }

    [Fact]
    public void MarkdownMultilineQuoteAndIndentedListsPreserveLevels()
    {
        const string source = """
            > 第一行引用
            > 第二行含公式 $q=1$

            - 一级项目
              - 二级项目
            1. 一级编号
                1. 三级编号
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Ole);

        var quote = Assert.Single(document.Blocks, block => block.Kind == WordBulkBlockKind.Quote);
        Assert.Contains(quote.Runs, run => !run.IsFormula && run.Text.Contains("第一行引用 第二行含公式"));
        Assert.Contains(quote.Runs, run => run.IsFormula && run.Latex == "q=1");
        var lists = document.Blocks.Skip(1).ToArray();
        Assert.Equal(
            new[] { 0, 1, 0, 2 },
            lists.Select(block => block.Level).ToArray());
        Assert.Equal(
            new[]
            {
                WordBulkBlockKind.Bullet,
                WordBulkBlockKind.Bullet,
                WordBulkBlockKind.Numbered,
                WordBulkBlockKind.Numbered,
            },
            lists.Select(block => block.Kind).ToArray());
        Assert.Empty(document.Warnings);
    }

    [Fact]
    public void EmbeddedAndAdjacentDisplayDelimitersAreSplitFromNativeParagraphs()
    {
        const string source = """
            对于高频相应，也可以对时间常数做一个近似计算： $\tau_H=\sum_{i=1}^{n}\tau_i$

            对于共射极放大电路： \[ R_1=\left(R_{si}\parallel R_b+r_{bb'}\right)\parallel r_{be}\]\[\tau_1=R_1C_{be}\]\[\tau_2=R_2C_{bc}=\left[\left(1+g_mR_L'\right)R_1+R_L'\right]C_{bc}\]
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Auto,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(WordBulkSourceFormat.Latex, document.SourceFormat);
        Assert.Equal(4, document.FormulaCount);
        Assert.Equal(1, document.InlineFormulaCount);
        Assert.Equal(3, document.DisplayFormulaCount);
        Assert.Equal(5, document.Blocks.Count);
        Assert.Equal(WordBulkBlockKind.Paragraph, document.Blocks[0].Kind);
        Assert.Equal(WordBulkBlockKind.Paragraph, document.Blocks[1].Kind);
        Assert.Equal("对于共射极放大电路：", string.Concat(
            document.Blocks[1].Runs.Where(run => !run.IsFormula).Select(run => run.Text)).Trim());
        var display = document.Blocks
            .Where(block => block.Kind == WordBulkBlockKind.DisplayFormula)
            .Select(block => block.Runs.Single().Latex)
            .ToArray();
        Assert.Equal("R_1=\\left(R_{si}\\parallel R_b+r_{bb'}\\right)\\parallel r_{be}", display[0]);
        Assert.Equal("\\tau_1=R_1C_{be}", display[1]);
        Assert.Equal("\\tau_2=R_2C_{bc}=\\left[\\left(1+g_mR_L'\\right)R_1+R_L'\\right]C_{bc}", display[2]);
        Assert.DoesNotContain(document.Blocks.SelectMany(block => block.Runs), run =>
            !run.IsFormula && (run.Text.Contains("\\[") || run.Text.Contains("\\]")));
        Assert.Empty(document.Warnings);
    }

    [Fact]
    public void EmbeddedDisplayFormulaCanSpanLinesAndKeepTrailingText()
    {
        const string source = """
            前文 \[
            \begin{aligned}
            a+b&=c\\
            d&=e
            \end{aligned}
            \] 后文仍然是正文。
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Ole);

        Assert.Equal(1, document.DisplayFormulaCount);
        Assert.Equal(3, document.Blocks.Count);
        Assert.Equal("前文", document.Blocks[0].Runs.Single().Text.Trim());
        Assert.Contains("\\begin{aligned}", document.Blocks[1].Runs.Single().Latex);
        Assert.Equal("后文仍然是正文。", document.Blocks[2].Runs.Single().Text.Trim());
        Assert.Empty(document.Warnings);
    }

    [Fact]
    public void DisplaySourceFormattingDoesNotCreateInvalidAlignedRows()
    {
        const string source = """
            对于高频相应，也可以对时间常数做一个近似计算：$$\tau _H=\sum_{i=1}^n{\tau _i}$$
            $$f_{_{\mathrm{H}}}=\frac{1}{2\pi \tau _{_{\mathrm{H}}}}$$
            对于共射级放大电路：\[
            R_1
            =
            \left(R_{si}\parallel R_b+r_{bb'}\right)
            \parallel r_{b'e}
            \]\[
            \tau_1
            =
            R_1 C_{b'e}
            \]\[
            \tau_2
            =
            R_2 C_{b'c}
            =
            \left[
            \left(1+g_mR_L'\right)R_1
            +
            R_L'
            \right]C_{b'c}
            \]
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Auto,
            WordBulkFormulaObjectMode.Omml);

        Assert.Equal(WordBulkSourceFormat.Latex, document.SourceFormat);
        Assert.Equal(5, document.FormulaCount);
        Assert.Equal(0, document.InlineFormulaCount);
        Assert.Equal(5, document.DisplayFormulaCount);
        var formulas = document.Blocks
            .Where(block => block.Kind == WordBulkBlockKind.DisplayFormula)
            .Select(block => block.Runs.Single().Latex)
            .ToArray();
        Assert.All(formulas, formula =>
        {
            Assert.DoesNotContain('\r', formula);
            Assert.DoesNotContain('\n', formula);
        });
        Assert.Equal(
            "\\tau_2 = R_2 C_{b'c} = \\left[ \\left(1+g_mR_L'\\right)R_1 + R_L' \\right]C_{b'c}",
            formulas[4]);
        Assert.DoesNotContain(document.Blocks.SelectMany(block => block.Runs), run =>
            !run.IsFormula && (run.Text.Contains("$$") || run.Text.Contains("\\[") || run.Text.Contains("\\]")));
        Assert.Empty(document.Warnings);
    }

    [Theory]
    [InlineData("\\[x+y\\tag{4.8.4}\\]", "4.8.4")]
    [InlineData("\\[x+y\\tag 4.8.4\\]", "4.8.4")]
    public void SeparatesTrailingEquationTagsFromEditableLatex(
        string source,
        string expectedTag)
    {
        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Omml);

        var formula = Assert.Single(
            document.Blocks.SelectMany(block => block.Runs).Where(run => run.IsFormula));
        Assert.Equal("x+y", formula.Latex);
        Assert.Equal(expectedTag, formula.EquationTag);
        Assert.DoesNotContain("\\tag", formula.Latex, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownStrongMarkersAndNestedEmphasisRemainStructured()
    {
        const string source = """
            **asterisk bold**

            __underscore bold__

            ***bold italic***

            ___underscore bold italic___

            **outer with *inner italic***
            """;

        var document = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Omml);

        var runs = document.Blocks.SelectMany(block => block.Runs).ToArray();
        Assert.Contains(runs, run => !run.IsFormula && run.Bold && !run.Italic && run.Text == "asterisk bold");
        Assert.Contains(runs, run => !run.IsFormula && run.Bold && !run.Italic && run.Text == "underscore bold");
        Assert.Contains(runs, run => !run.IsFormula && run.Bold && run.Italic && run.Text == "bold italic");
        Assert.Contains(runs, run => !run.IsFormula && run.Bold && run.Italic && run.Text == "underscore bold italic");
        Assert.Contains(runs, run => !run.IsFormula && run.Bold && !run.Italic && run.Text.Contains("outer with"));
        Assert.Contains(runs, run => !run.IsFormula && run.Bold && run.Italic && run.Text == "inner italic");
        Assert.Empty(document.Warnings);
    }

    [Fact]
    public void SerializedPreviewDocumentIsUsedVerbatimByWordImport()
    {
        const string serialized = """
            {
              "format": "markdown",
              "blocks": [
                {
                  "kind": "heading",
                  "level": 2,
                  "runs": [{ "kind": "text", "text": "Title", "bold": true }]
                },
                {
                  "kind": "paragraph",
                  "level": 0,
                  "runs": [
                    { "kind": "text", "text": "deleted", "strike": true },
                    { "kind": "text", "text": "underlined", "underline": true },
                    { "kind": "formula", "latex": "x+y", "display": false }
                  ]
                },
                {
                  "kind": "display",
                  "level": 0,
                  "runs": [
                    { "kind": "formula", "latex": "E=mc^2", "display": true, "equationTag": "4.2" }
                  ]
                },
                {
                  "kind": "code",
                  "level": 0,
                  "runs": [{ "kind": "text", "text": "const x = 1;", "code": true }]
                }
              ],
              "warnings": ["fallback warning"]
            }
            """;

        var document = WordBulkImportParser.ParseSerialized(
            serialized,
            WordBulkFormulaObjectMode.Ole);

        Assert.Equal(WordBulkSourceFormat.Markdown, document.SourceFormat);
        Assert.Equal(WordBulkFormulaObjectMode.Ole, document.FormulaObjectMode);
        Assert.Equal(4, document.Blocks.Count);
        Assert.Equal(2, document.FormulaCount);
        Assert.Equal(1, document.InlineFormulaCount);
        Assert.Equal(1, document.DisplayFormulaCount);
        Assert.Equal(WordBulkBlockKind.Heading, document.Blocks[0].Kind);
        Assert.Equal(2, document.Blocks[0].Level);
        Assert.True(document.Blocks[1].Runs[0].Strike);
        Assert.True(document.Blocks[1].Runs[1].Underline);
        Assert.Equal("x+y", document.Blocks[1].Runs[2].Latex);
        Assert.Equal("E=mc^2", document.Blocks[2].Runs.Single().Latex);
        Assert.Equal("4.2", document.Blocks[2].Runs.Single().EquationTag);
        Assert.Equal(WordBulkBlockKind.Code, document.Blocks[3].Kind);
        Assert.True(document.Blocks[3].Runs.Single().Code);
        Assert.Equal("fallback warning", Assert.Single(document.Warnings));
    }

    [Fact]
    public void SerializedPreviewDocumentRejectsInvalidOrEmptyPayloads()
    {
        Assert.Throws<InvalidDataException>(() => WordBulkImportParser.ParseSerialized(
            "not-json",
            WordBulkFormulaObjectMode.Omml));
        Assert.Throws<InvalidDataException>(() => WordBulkImportParser.ParseSerialized(
            "{\"format\":\"markdown\",\"blocks\":[]}",
            WordBulkFormulaObjectMode.Omml));
    }

    [Fact]
    public void FindsExactLatexSpansForInPlaceWordRedraw()
    {
        const string source =
            "紫外线 $UVI>2$，相邻公式 $a$$b$。\r"
            + "\\(x+y\\)\r$$E=mc^2$$\r\\[k^2\\]\r"
            + "\\begin{align}a&=b\\\\c&=d\\end{align}";

        var spans = WordBulkImportParser.FindFormulaSpans(source);

        Assert.Equal(7, spans.Count);
        Assert.Equal(
            new[] { "UVI>2", "a", "b", "x+y", "E=mc^2", "k^2", "\\begin{aligned}a&=b\\\\c&=d\\end{aligned}" },
            spans.Select(span => span.Latex).ToArray());
        Assert.Equal(
            new[] { "inline", "inline", "inline", "inline", "block", "block", "block" },
            spans.Select(span => span.DisplayMode).ToArray());
        Assert.Equal("$UVI>2$", source.Substring(spans[0].Start, spans[0].Length));
        Assert.Equal("$$E=mc^2$$", source.Substring(spans[4].Start, spans[4].Length));
        Assert.Equal("\\[k^2\\]", source.Substring(spans[5].Start, spans[5].Length));
    }

    [Fact]
    public void RedrawScannerIgnoresEscapedOrUnclosedDelimiters()
    {
        const string source = @"价格为 \$5，保留 $unclosed，只有 \(x\) 可转换。";

        var spans = WordBulkImportParser.FindFormulaSpans(source);

        var span = Assert.Single(spans);
        Assert.Equal("x", span.Latex);
        Assert.Equal("\\(x\\)", source.Substring(span.Start, span.Length));
    }

    [Fact]
    public void RedrawScannerTreatsWordTableCellEndsAsHardFormulaBoundaries()
    {
        const string source =
            "左格 $unclosed\r\a"
            + "右格正文 和 $b$\r\a"
            + "第三格 $$c+d$$\r\a"
            + "第四格 \\begin{align}x&=1\r\a"
            + "第五格 y&=2\\end{align}\r\a"
            + "第六格 \\(z\\)";

        var spans = WordBulkImportParser.FindFormulaSpans(source);

        Assert.Equal(3, spans.Count);
        Assert.Equal(new[] { "b", "c+d", "z" }, spans.Select(span => span.Latex).ToArray());
        Assert.Equal(new[] { "inline", "block", "inline" }, spans.Select(span => span.DisplayMode).ToArray());
        Assert.Equal("$b$", source.Substring(spans[0].Start, spans[0].Length));
        Assert.Equal("$$c+d$$", source.Substring(spans[1].Start, spans[1].Length));
        Assert.Equal("\\(z\\)", source.Substring(spans[2].Start, spans[2].Length));
        Assert.All(spans, span =>
            Assert.DoesNotContain('\a', source.Substring(span.Start, span.Length)));
    }

    [Fact]
    public void RejectsEmptyOrExcessivelyLargeInput()
    {
        Assert.Throws<InvalidDataException>(() => WordBulkImportParser.Parse(
            "   ",
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Omml));
        Assert.Throws<InvalidDataException>(() => WordBulkImportParser.Parse(
            new string('a', 5_000_001),
            WordBulkSourceFormat.Markdown,
            WordBulkFormulaObjectMode.Omml));
    }
}
