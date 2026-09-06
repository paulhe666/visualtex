import assert from "node:assert/strict";
import { webcrypto } from "node:crypto";
import { readFileSync } from "node:fs";
import {
  parseDocumentImport,
  type DocumentImportBlock,
  type DocumentImportRun,
  type DocumentSourceFormat,
  type ParsedDocumentImport,
} from "../src/office/documentImport/documentImportParser.ts";

if (!globalThis.crypto) {
  Object.defineProperty(globalThis, "crypto", { value: webcrypto });
}

type CoverageCase = {
  name: string;
  format: DocumentSourceFormat;
  source: string;
  check: (parsed: ParsedDocumentImport) => boolean;
};

function runs(parsed: ParsedDocumentImport): DocumentImportRun[] {
  return parsed.blocks.flatMap((block) => block.runs);
}

function text(parsed: ParsedDocumentImport): string {
  return runs(parsed)
    .filter((run): run is Extract<DocumentImportRun, { kind: "text" }> => run.kind === "text")
    .map((run) => run.text)
    .join("");
}

function formulas(parsed: ParsedDocumentImport) {
  return runs(parsed).filter(
    (run): run is Extract<DocumentImportRun, { kind: "formula" }> => run.kind === "formula",
  );
}

function blocksOf(parsed: ParsedDocumentImport, kind: DocumentImportBlock["kind"]) {
  return parsed.blocks.filter((block) => block.kind === kind);
}

const markdownBaseline: CoverageCase[] = [
  {
    name: "ATX headings # through ######",
    format: "markdown",
    source: "# H1\n\n## H2\n\n### H3\n\n#### H4\n\n##### H5\n\n###### H6",
    check: (parsed) =>
      blocksOf(parsed, "heading").length === 6 &&
      blocksOf(parsed, "heading").every((block, index) => block.level === index + 1),
  },
  {
    name: "paragraph line wrapping",
    format: "markdown",
    source: "第一行\n第二行\n\n新段落",
    check: (parsed) =>
      blocksOf(parsed, "paragraph").length === 2 && text(parsed).includes("第一行 第二行"),
  },
  {
    name: "bold italic nested emphasis and inline code",
    format: "markdown",
    source: "**粗体**、*斜体*、_斜体_、**粗体含 *斜体***、`code()`",
    check: (parsed) => {
      const all = runs(parsed);
      return (
        all.some((run) => run.kind === "text" && run.bold && run.text === "粗体") &&
        all.filter((run) => run.kind === "text" && run.italic).length >= 3 &&
        all.some((run) => run.kind === "text" && run.bold && run.italic) &&
        all.some((run) => run.kind === "text" && run.code && run.text === "code()")
      );
    },
  },
  {
    name: "escaped Markdown punctuation",
    format: "markdown",
    source: String.raw`\*literal\* \_name\_ \#hash \{brace\} \!bang \$5`,
    check: (parsed) =>
      parsed.formulaCount === 0 && text(parsed).includes("*literal* _name_ #hash {brace} !bang $5"),
  },
  {
    name: "inline math dollar and parenthesis delimiters",
    format: "markdown",
    source: String.raw`行内 $E=mc^2$ 与 \(a+b=c\)。价格 \$5。`,
    check: (parsed) =>
      formulas(parsed).filter((run) => !run.display).map((run) => run.latex).join("|") ===
        "E=mc^2|a+b=c" && text(parsed).includes("$5"),
  },
  {
    name: "display math dollar delimiters same-line and multiline",
    format: "markdown",
    source: String.raw`前文 $$x=1$$ 后文

$$
\int_0^1 x\,dx
$$`,
    check: (parsed) =>
      formulas(parsed).filter((run) => run.display).length === 2 &&
      formulas(parsed).some((run) => run.display && run.latex === "x=1") &&
      formulas(parsed).some((run) => run.display && run.latex.includes("\\int_0^1")),
  },
  {
    name: "display math bracket delimiters",
    format: "markdown",
    source: String.raw`正文 \[\frac{a}{b}\] 结尾`,
    check: (parsed) =>
      formulas(parsed).some((run) => run.display && run.latex === String.raw`\frac{a}{b}`),
  },
  {
    name: "multiline blockquote",
    format: "markdown",
    source: "> 第一行\n> 第二行 $q=1$",
    check: (parsed) =>
      blocksOf(parsed, "quote").length === 1 &&
      formulas(parsed).some((run) => run.latex === "q=1"),
  },
  {
    name: "unordered list markers and nested indentation",
    format: "markdown",
    source: "- dash\n+ plus\n* star\n  - nested",
    check: (parsed) => {
      const list = blocksOf(parsed, "bullet");
      return list.length === 4 && list.map((block) => block.level).join(",") === "0,0,0,1";
    },
  },
  {
    name: "ordered list dot and parenthesis markers",
    format: "markdown",
    source: "1. first\n2) second\n    1. nested",
    check: (parsed) => {
      const list = blocksOf(parsed, "numbered");
      return list.length === 3 && list.map((block) => block.level).join(",") === "0,0,2";
    },
  },
  {
    name: "backtick fenced code with info string",
    format: "markdown",
    source: "```ts\nconst x = `$notMath`;\n```",
    check: (parsed) => {
      const code = blocksOf(parsed, "code");
      return code.length === 1 && text(parsed).includes("const x = `$notMath`;");
    },
  },
  {
    name: "unclosed fenced code warning",
    format: "markdown",
    source: "```\nconst x = 1;",
    check: (parsed) => blocksOf(parsed, "code").length === 1 && parsed.warnings.length === 1,
  },
];

const latexVerbatimBoundaryFixture = [
  String.raw`\section*{综合 LaTeX 正文示例}`,
  "",
  String.raw`这是一段不含 \verb|\documentclass|、导言区和宏包加载命令的正文代码，可直接放入已有文档的 \verb|\begin{document}| 与 \verb|\end{document}| 之间。`,
  "",
  String.raw`\begin{center}`,
  String.raw`\textbf{\Large 常见公式与排版语法综合练习}`,
  String.raw`\end{center}`,
  "",
  String.raw`\begin{quote}`,
  String.raw`\emph{排版的目标不是堆叠命令，而是让结构、符号和论证关系更加清楚。}`,
  String.raw`\end{quote}`,
  "",
  String.raw`\begin{itemize}`,
  String.raw`\item \textbf{行内公式：}适合嵌入普通句子。`,
  String.raw`\item \textbf{行间公式：}适合核心结论。`,
  String.raw`\item \textbf{高亮语法：}包括粗体和斜体。`,
  String.raw`\end{itemize}`,
  "",
  String.raw`\begin{enumerate}`,
  String.raw`\item 先定义符号与假设；`,
  String.raw`\item 再给出推导过程；`,
  String.raw`\item 最后解释结果。`,
  String.raw`\end{enumerate}`,
  "",
  String.raw`\begin{description}`,
  String.raw`\item[定义] 明确对象是什么。`,
  String.raw`\item[定理] 陈述在什么条件下成立。`,
  String.raw`\item[证明] 给出逻辑过程。`,
  String.raw`\end{description}`,
  "",
  ...Array.from({ length: 20 }, (_, index) => {
    const number = index + 1;
    return [
      String.raw`\section{${number}. 综合公式}`,
      `正文 $x_${number}=a_${number}+b_${number}$、$y_${number}=c_${number}-d_${number}$ 与 $z_${number}=e_${number}f_${number}$。`,
      String.raw`\[A_{${number}}=\frac{${number}}{${number}+1}\]`,
      String.raw`\[B_{${number}}=\sum_{k=1}^{${number}}k\]`,
    ].join("\n");
  }),
  String.raw`\section{21. 综合案例与表达规范}`,
  "最后一节没有额外公式，但必须完整保留。",
  String.raw`\section*{排版检查清单}`,
  "正文结束。",
].join("\n");

const exactLatexUserFixture = readFileSync(
  new URL("./fixtures/latex_100_formulas_4000_chinese.tex", import.meta.url),
  "utf8",
);

const latexBaseline: CoverageCase[] = [
  {
    name: "full document preamble stripping and comments",
    format: "auto",
    source: String.raw`\documentclass{article}
\usepackage{amsmath}
\begin{document}
正文 % remove this
保留 \% 百分号
\end{document}`,
    check: (parsed) =>
      parsed.format === "latex" &&
      !text(parsed).includes("documentclass") &&
      !text(parsed).includes("remove this") &&
      text(parsed).includes("保留 % 百分号"),
  },
  {
    name: "verbatim document command examples do not truncate a LaTeX fragment",
    format: "latex",
    source: latexVerbatimBoundaryFixture,
    check: (parsed) => {
      const allFormulas = formulas(parsed);
      const codeRuns = runs(parsed).filter(
        (run): run is Extract<DocumentImportRun, { kind: "text" }> =>
          run.kind === "text" && Boolean(run.code),
      );
      return (
        blocksOf(parsed, "heading").length === 23 &&
        allFormulas.length === 100 &&
        allFormulas.filter((run) => !run.display).length === 60 &&
        allFormulas.filter((run) => run.display).length === 40 &&
        codeRuns.map((run) => run.text).join("|") ===
          String.raw`\documentclass|\begin{document}|\end{document}` &&
        text(parsed).includes("最后一节没有额外公式，但必须完整保留。") &&
        text(parsed).includes("常见公式与排版语法综合练习") &&
        !text(parsed).includes("\\Large") &&
        text(parsed).includes("正文结束。")
      );
    },
  },
  {
    name: "uploaded 8743-character 100-formula LaTeX fixture parses completely",
    format: "latex",
    source: exactLatexUserFixture,
    check: (parsed) => {
      const allFormulas = formulas(parsed);
      return (
        exactLatexUserFixture.replace(/\r\n/g, "\n").length === 8743 &&
        blocksOf(parsed, "heading").length === 23 &&
        allFormulas.length === 100 &&
        allFormulas.filter((run) => !run.display).length === 60 &&
        allFormulas.filter((run) => run.display).length === 40 &&
        blocksOf(parsed, "bullet").length === 6 &&
        blocksOf(parsed, "numbered").length === 8 &&
        text(parsed).includes("这是一段不含 \\documentclass、导言区和宏包加载命令的正文代码") &&
        text(parsed).includes("正文结束：共包含 100 个数学公式示例") &&
        !text(parsed).includes("\\Large")
      );
    },
  },
  {
    name: "LaTeX CJK inline math spacing follows TeX semantics",
    format: "latex",
    source: String.raw`中文 $x=1$ 中文；English $y=2$ words；中文\ $z=3$\ 中文；中文~$q=4$~中文；中文\quad$r=5$\qquad中文`,
    check: (parsed) =>
      runs(parsed)
        .map((run) => (run.kind === "formula" ? `<${run.latex}>` : run.text))
        .join("") ===
      `中文<x=1>中文；English <y=2> words；中文\u00a0<z=3>\u00a0中文；中文\u00a0<q=4>\u00a0中文；中文\u00a0<r=5>\u00a0\u00a0中文`,
  },
  {
    name: "real document boundaries ignore comments and inline literal examples",
    format: "latex",
    source: String.raw`% \begin{document}
\documentclass{article}
\begin{verbatim}
\begin{document}
\end{document}
\end{verbatim}
\begin{document}
正文前 \verb|\end{document}| 正文后 $x=1$。
\end{document}
尾部不应导入。`,
    check: (parsed) =>
      text(parsed).includes("正文前 \\end{document} 正文后") &&
      !text(parsed).includes("尾部不应导入") &&
      formulas(parsed).length === 1,
  },
  {
    name: "all common sectioning commands",
    format: "latex",
    source: String.raw`\part{P}
\chapter{C}
\section{S}
\subsection{SS}
\subsubsection{SSS}
\paragraph{P4}
\subparagraph{P5}`,
    check: (parsed) => {
      const headings = blocksOf(parsed, "heading");
      return headings.length === 7 && headings.map((block) => block.level).join(",") === "1,2,3,4,5,6,7";
    },
  },
  {
    name: "nested textbf textit emph and texttt",
    format: "latex",
    source: String.raw`\textbf{粗体和 \textit{粗斜体}} \emph{强调} \texttt{code_1}`,
    check: (parsed) => {
      const all = runs(parsed);
      return (
        all.some((run) => run.kind === "text" && run.bold && run.text.includes("粗体和")) &&
        all.some((run) => run.kind === "text" && run.bold && run.italic && run.text === "粗斜体") &&
        all.some((run) => run.kind === "text" && run.italic && run.text.includes("强调")) &&
        all.some((run) => run.kind === "text" && run.code && run.text === "code_1")
      );
    },
  },
  {
    name: "escaped LaTeX text characters",
    format: "latex",
    source: String.raw`A\% B\_ C\& D\# E\$ F\{ G\} H\textbackslash{}I~J`,
    check: (parsed) => text(parsed).includes("A% B_ C& D# E$ F{ G} H\\I\u00a0J"),
  },
  {
    name: "inline math dollar and parenthesis delimiters",
    format: "latex",
    source: String.raw`公式 $x_1$ 与 \(\alpha+\beta\)。`,
    check: (parsed) =>
      formulas(parsed).filter((run) => !run.display).map((run) => run.latex).join("|") ===
      String.raw`x_1|\alpha+\beta`,
  },
  {
    name: "display dollar and bracket delimiters including adjacency",
    format: "latex",
    source: String.raw`前 $$a=1$$ 中 \[b=2\]\[c=3\] 后`,
    check: (parsed) =>
      formulas(parsed).filter((run) => run.display).map((run) => run.latex).join("|") === "a=1|b=2|c=3",
  },
  {
    name: "equation and equation* environments",
    format: "latex",
    source: String.raw`\begin{equation}a=1\end{equation}
\begin{equation*}b=2\end{equation*}`,
    check: (parsed) => formulas(parsed).filter((run) => run.display).length === 2,
  },
  {
    name: "equation* formatting newlines are TeX whitespace",
    format: "latex",
    source: String.raw`\begin{equation*}
\lim_{N\to +\infty}\bigl\|\,|\psi\rangle-|s_N\rangle\,\bigr\|
=
\lim_{N\to +\infty}
\left\|\,|\psi\rangle-\sum_{i=1}^{N}c_i|u_i\rangle\,\right\|
=0.
\end{equation*}`,
    check: (parsed) => {
      const formula = formulas(parsed).find((run) => run.display);
      return Boolean(
        formula &&
        !formula.latex.includes("\n") &&
        formula.latex.includes(String.raw`\lim_{N\to +\infty}`) &&
        formula.latex.endsWith("=0."),
      );
    },
  },
  {
    name: "align source newlines preserve explicit math rows",
    format: "latex",
    source: String.raw`\begin{align*}
a &= b+c \\
d &= e-f \\
g &= h
\end{align*}`,
    check: (parsed) => {
      const formula = formulas(parsed).find((run) => run.display);
      return Boolean(
        formula &&
        !formula.latex.includes("\n") &&
        formula.latex.startsWith(String.raw`\begin{aligned}`) &&
        formula.latex.includes(String.raw`a &= b+c \\ d &= e-f \\ g &= h`) &&
        formula.latex.endsWith(String.raw`\end{aligned}`),
      );
    },
  },
  {
    name: "align gather multline and displaymath environments",
    format: "latex",
    source: String.raw`\begin{align}a&=1\\b&=2\end{align}
\begin{gather*}c=3\\d=4\end{gather*}
\begin{multline}e+f\\=g\end{multline}
\begin{displaymath}h=5\end{displaymath}`,
    check: (parsed) => {
      const display = formulas(parsed).filter((run) => run.display);
      return (
        display.length === 4 &&
        display[0].latex.includes("begin{aligned}") &&
        display[1].latex.includes("begin{gathered}") &&
        display[2].latex.includes("begin{gathered}")
      );
    },
  },
  {
    name: "equation tags separated from editable formula",
    format: "latex",
    source: String.raw`\[x+y\tag{4.8.4}\]`,
    check: (parsed) => {
      const formula = formulas(parsed)[0];
      return formula?.latex === "x+y" && formula.equationTag === "4.8.4";
    },
  },
  {
    name: "nested itemize enumerate and optional item label",
    format: "latex",
    source: String.raw`\begin{itemize}
\item outer
\begin{enumerate}
\item[1)] inner $n=1$
\end{enumerate}
\end{itemize}`,
    check: (parsed) => {
      const list = parsed.blocks.filter((block) => block.kind === "bullet" || block.kind === "numbered");
      return list.length === 2 && list[0].kind === "bullet" && list[0].level === 0 && list[1].kind === "numbered" && list[1].level === 1;
    },
  },
  {
    name: "quote and quotation environments",
    format: "latex",
    source: String.raw`\begin{quote}
Q1 $x$
\end{quote}
\begin{quotation}
Q2
\end{quotation}`,
    check: (parsed) => blocksOf(parsed, "quote").length === 2,
  },
  {
    name: "verbatim and lstlisting literal code",
    format: "latex",
    source: String.raw`\begin{verbatim}
a_1 = 20 % literal
\end{verbatim}
\begin{lstlisting}[language=Python]
print("$x$")
\end{lstlisting}`,
    check: (parsed) => {
      const code = blocksOf(parsed, "code");
      return code.length === 2 && text(parsed).includes("a_1 = 20 % literal") && text(parsed).includes('print("$x$")');
    },
  },
  {
    name: "formula internals matrices cases arrays and operators pass through",
    format: "latex",
    source: String.raw`\[\begin{cases}x^2,&x>0\\0,&x\le0\end{cases}\]
\[\begin{pmatrix}a&b\\c&d\end{pmatrix}+\sum_{i=1}^n i+\oiint_\Sigma f\,dS\]`,
    check: (parsed) => {
      const display = formulas(parsed).filter((run) => run.display);
      return display.length === 2 && display[0].latex.includes("begin{cases}") && display[1].latex.includes("begin{pmatrix}") && display[1].latex.includes("\\oiint");
    },
  },
  {
    name: "unclosed structures produce warnings rather than crashes",
    format: "latex",
    source: String.raw`\begin{itemize}
\item item
\begin{quote}
text
\[
x+y`,
    check: (parsed) => parsed.warnings.length >= 2 && formulas(parsed).some((run) => run.display),
  },
];

const markdownExtended: CoverageCase[] = [
  {
    name: "escaped square brackets alongside the \\[...\\] math extension",
    format: "markdown",
    source: String.raw`\[literal bracket\]`,
    check: (parsed) => parsed.formulaCount === 0 && text(parsed) === "[literal bracket]",
  },
  {
    name: "Setext headings",
    format: "markdown",
    source: "Heading\n=======",
    check: (parsed) => blocksOf(parsed, "heading").length === 1,
  },
  {
    name: "double-underscore strong emphasis",
    format: "markdown",
    source: "__bold__",
    check: (parsed) => runs(parsed).some((run) => run.kind === "text" && run.bold && run.text === "bold"),
  },
  {
    name: "triple-marker bold italic",
    format: "markdown",
    source: "***both***",
    check: (parsed) => runs(parsed).some((run) => run.kind === "text" && run.bold && run.italic && run.text === "both"),
  },
  {
    name: "GFM strikethrough",
    format: "markdown",
    source: "~~deleted~~",
    check: (parsed) =>
      runs(parsed).some((run) => run.kind === "text" && run.strike && run.text === "deleted"),
  },
  {
    name: "inline links",
    format: "markdown",
    source: "[OpenAI](https://openai.com)",
    check: (parsed) => text(parsed).includes("OpenAI") && text(parsed).includes("https://openai.com"),
  },
  {
    name: "reference links",
    format: "markdown",
    source: "[OpenAI][oa]\n\n[oa]: https://openai.com",
    check: (parsed) =>
      text(parsed).includes("OpenAI") &&
      !text(parsed).includes("[oa]") &&
      text(parsed).includes("https://openai.com"),
  },
  {
    name: "images with alt text",
    format: "markdown",
    source: "![diagram](diagram.png)",
    check: (parsed) => text(parsed).includes("diagram") && text(parsed).includes("diagram.png"),
  },
  {
    name: "GFM tables",
    format: "markdown",
    source: "| A | B |\n|---|---|\n| 1 | 2 |",
    check: (parsed) =>
      text(parsed).includes("A") &&
      text(parsed).includes("B") &&
      text(parsed).includes("1") &&
      text(parsed).includes("2") &&
      !text(parsed).includes("---"),
  },
  {
    name: "horizontal rules",
    format: "markdown",
    source: "before\n\n---\n\nafter",
    check: (parsed) => text(parsed).includes("────────────────────"),
  },
  {
    name: "task list semantics",
    format: "markdown",
    source: "- [x] done\n- [ ] todo",
    check: (parsed) => {
      const items = blocksOf(parsed, "bullet");
      return items.length === 2 && text(parsed).includes("☒ done") && text(parsed).includes("☐ todo");
    },
  },
  {
    name: "tilde fenced code",
    format: "markdown",
    source: "~~~js\nconst x = 1;\n~~~",
    check: (parsed) => blocksOf(parsed, "code").length === 1,
  },
  {
    name: "indented code blocks",
    format: "markdown",
    source: "    const x = 1;",
    check: (parsed) => blocksOf(parsed, "code").length === 1,
  },
  {
    name: "multi-backtick inline code",
    format: "markdown",
    source: "``code with ` tick``",
    check: (parsed) => runs(parsed).some((run) => run.kind === "text" && run.code && run.text === "code with ` tick"),
  },
  {
    name: "hard line breaks",
    format: "markdown",
    source: "line one  \nline two",
    check: (parsed) => text(parsed).includes("line one\nline two"),
  },
  {
    name: "raw HTML blocks",
    format: "markdown",
    source: "<details><summary>Title</summary>Body</details>",
    check: (parsed) => text(parsed) === "TitleBody",
  },
  {
    name: "footnotes",
    format: "markdown",
    source: "Text[^1]\n\n[^1]: note",
    check: (parsed) =>
      text(parsed).includes("Text〔注 1〕") &&
      text(parsed).includes("注 1：") &&
      text(parsed).includes("note"),
  },
  {
    name: "shortcut reference links",
    format: "markdown",
    source: "[OpenAI]\n\n[OpenAI]: https://openai.com",
    check: (parsed) =>
      text(parsed).includes("OpenAI") && text(parsed).includes("https://openai.com"),
  },
  {
    name: "HTML entities",
    format: "markdown",
    source: "A &amp; B, &#x03B1; &lt; x &gt; &quot;q&quot;",
    check: (parsed) => text(parsed).includes('A & B, α < x > "q"'),
  },
  {
    name: "HTML comments are hidden",
    format: "markdown",
    source: "before <!-- hidden --> after",
    check: (parsed) => text(parsed).includes("before") && text(parsed).includes("after") && !text(parsed).includes("hidden"),
  },
  {
    name: "YAML front matter is preserved as metadata code",
    format: "markdown",
    source: "---\ntitle: Demo\nauthor: Test\n---\n# Body",
    check: (parsed) =>
      blocksOf(parsed, "code").length === 1 &&
      blocksOf(parsed, "heading").length === 1 &&
      text(parsed).includes("title: Demo") &&
      text(parsed).includes("Body"),
  },
];

const latexExtended: CoverageCase[] = [
  {
    name: "description lists",
    format: "latex",
    source: String.raw`\begin{description}\item[Term] Definition\end{description}`,
    check: (parsed) =>
      blocksOf(parsed, "bullet").length === 1 &&
      text(parsed).includes("Term") &&
      text(parsed).includes("Definition"),
  },
  {
    name: "flalign environments",
    format: "latex",
    source: String.raw`\begin{flalign}a&=b&&\end{flalign}`,
    check: (parsed) => formulas(parsed).some((run) => run.display),
  },
  {
    name: "alignat environments",
    format: "latex",
    source: String.raw`\begin{alignat}{2}a&=b&\quad c&=d\end{alignat}`,
    check: (parsed) => formulas(parsed).some((run) => run.display),
  },
  {
    name: "eqnarray environments",
    format: "latex",
    source: String.raw`\begin{eqnarray}a&=&b\end{eqnarray}`,
    check: (parsed) => formulas(parsed).some((run) => run.display),
  },
  {
    name: "math environments",
    format: "latex",
    source: String.raw`text \begin{math}x+y\end{math}`,
    check: (parsed) => formulas(parsed).some((run) => !run.display && run.latex === "x+y"),
  },
  {
    name: "center flushleft and flushright environments",
    format: "latex",
    source: String.raw`\begin{center}Centered\end{center}`,
    check: (parsed) => text(parsed) === "Centered",
  },
  {
    name: "tabular and table structures",
    format: "latex",
    source: String.raw`\begin{tabular}{cc}A&B\\1&2\end{tabular}`,
    check: (parsed) =>
      parsed.blocks.length === 2 &&
      text(parsed).includes("A ｜ B") &&
      text(parsed).includes("1 ｜ 2"),
  },
  {
    name: "figure and includegraphics",
    format: "latex",
    source: String.raw`\begin{figure}\includegraphics{plot.png}\caption{Plot}\end{figure}`,
    check: (parsed) =>
      text(parsed).includes("图片：plot.png") &&
      text(parsed).includes("图注：") &&
      text(parsed).includes("Plot"),
  },
  {
    name: "title author date and maketitle",
    format: "latex",
    source: String.raw`\title{Title}\author{Author}\date{}\begin{document}\maketitle\end{document}`,
    check: (parsed) => text(parsed).includes("Title") && text(parsed).includes("Author"),
  },
  {
    name: "href and url commands",
    format: "latex",
    source: String.raw`\href{https://example.com}{Example} \url{https://example.com}`,
    check: (parsed) =>
      text(parsed).includes("Example") &&
      text(parsed).split("https://example.com").length === 3,
  },
  {
    name: "footnotes",
    format: "latex",
    source: String.raw`Text\footnote{Note}`,
    check: (parsed) => text(parsed) === "Text（注：Note）",
  },
  {
    name: "citations and references",
    format: "latex",
    source: String.raw`See \cite{key}, Eq.~\ref{eq:a}.`,
    check: (parsed) => !text(parsed).includes("\\cite") && !text(parsed).includes("\\ref"),
  },
  {
    name: "proof-only fragments are auto-detected as LaTeX",
    format: "auto",
    source: String.raw`\begin{proof}[Spectral argument]
For every $N$, the partial sum converges.
\qedhere
\end{proof}`,
    check: (parsed) =>
      parsed.format === "latex" &&
      blocksOf(parsed, "heading").some(
        (block) => block.level === 4 && block.runs.some((run) => run.kind === "text" && run.text.includes("证明（Spectral argument）")),
      ) &&
      blocksOf(parsed, "quote").length === 0 &&
      text(parsed).includes("partial sum converges") &&
      text(parsed).includes("□") &&
      formulas(parsed).some((run) => run.latex === "N"),
  },
  {
    name: "starred theorem-like environments",
    format: "latex",
    source: String.raw`\begin{lemma*}A starred lemma.\end{lemma*}`,
    check: (parsed) =>
      blocksOf(parsed, "heading").some(
        (block) => block.level === 4 && block.runs.some((run) => run.kind === "text" && run.text === "引理"),
      ) &&
      blocksOf(parsed, "quote").length === 1 &&
      !text(parsed).includes("引理 1") &&
      text(parsed).includes("A starred lemma."),
  },
  {
    name: "custom newtheorem environments",
    format: "auto",
    source: String.raw`\newtheorem{spectralclaim}[theorem]{Spectral Claim}
\begin{document}
\begin{spectralclaim}[Finite truncation]
The approximation error tends to zero.
\end{spectralclaim}
\end{document}`,
    check: (parsed) =>
      parsed.format === "latex" &&
      blocksOf(parsed, "heading").some(
        (block) => block.level === 4 && block.runs.some((run) => run.kind === "text" && run.text.includes("Spectral Claim 1（Finite truncation）")),
      ) &&
      blocksOf(parsed, "quote").length === 1 &&
      text(parsed).includes("approximation error tends to zero") &&
      !text(parsed).includes("newtheorem"),
  },
  {
    name: "theorem-like environments",
    format: "latex",
    source: String.raw`\begin{theorem}Every finite subgroup is...\end{theorem}`,
    check: (parsed) =>
      blocksOf(parsed, "heading").some(
        (block) => block.level === 4 && block.runs.some((run) => run.kind === "text" && run.text === "定理 1"),
      ) &&
      blocksOf(parsed, "quote").length === 1 &&
      text(parsed).includes("Every finite subgroup is..."),
  },
  {
    name: "newtheorem shared counters",
    format: "latex",
    source: String.raw`\newtheorem{thm}{定理}
\newtheorem{lem}[thm]{引理}
\begin{thm}First.\end{thm}
\begin{lem}Second.\end{lem}`,
    check: (parsed) => {
      const headingText = blocksOf(parsed, "heading")
        .flatMap((block) => block.runs)
        .filter((run): run is Extract<DocumentImportRun, { kind: "text" }> => run.kind === "text")
        .map((run) => run.text);
      return headingText.includes("定理 1") && headingText.includes("引理 2");
    },
  },
  {
    name: "newtheorem star is unnumbered",
    format: "latex",
    source: String.raw`\newtheorem*{specialremark}{特别说明}
\begin{specialremark}[边界情况]No number.\end{specialremark}`,
    check: (parsed) => {
      const headingText = blocksOf(parsed, "heading")
        .flatMap((block) => block.runs)
        .filter((run): run is Extract<DocumentImportRun, { kind: "text" }> => run.kind === "text")
        .map((run) => run.text)
        .join("|");
      return headingText.includes("特别说明（边界情况）") && !headingText.includes("特别说明 1");
    },
  },
  {
    name: "nested proof overrides theorem quote body style",
    format: "latex",
    source: String.raw`\begin{theorem}
Theorem body.
\begin{proof}
Proof body $x=1$.\qed
\end{proof}
Theorem tail.
\end{theorem}`,
    check: (parsed) =>
      blocksOf(parsed, "heading").length === 2 &&
      blocksOf(parsed, "quote").some((block) => block.runs.some((run) => run.kind === "text" && run.text.includes("Theorem body"))) &&
      blocksOf(parsed, "paragraph").some((block) => block.runs.some((run) => run.kind === "text" && run.text.includes("Proof body"))) &&
      blocksOf(parsed, "quote").some((block) => block.runs.some((run) => run.kind === "text" && run.text.includes("Theorem tail"))) &&
      formulas(parsed).some((run) => run.latex === "x=1"),
  },
  {
    name: "custom macro expansion",
    format: "latex",
    source: String.raw`\newcommand{\R}{\mathbb{R}}\begin{document}$x\in\R$\end{document}`,
    check: (parsed) => formulas(parsed).some((run) => run.latex.includes("\\mathbb{R}")),
  },
  {
    name: "four argument nested custom macros",
    format: "latex",
    source: String.raw`\newcommand{\four}[4]{\frac{#1_{#2}}{#3^{#4}}}
\begin{document}$\four{a+b}{i}{c+d}{2}$\end{document}`,
    check: (parsed) =>
      formulas(parsed).some((run) => run.latex === String.raw`\frac{a+b_{i}}{c+d^{2}}`),
  },
  {
    name: "custom macros with default first argument",
    format: "latex",
    source: String.raw`\newcommand{\decorate}[2][*]{#1#2#1}
$\decorate{x}+\decorate[!]{y}$`,
    check: (parsed) => formulas(parsed).some((run) => run.latex === "*x*+!y!"),
  },
  {
    name: "def macros with parameters",
    format: "latex",
    source: String.raw`\def\tensor#1#2#3{#1_{#2}^{#3}} $\tensor{T}{ij}{k}$`,
    check: (parsed) => formulas(parsed).some((run) => run.latex === "T_{ij}^{k}"),
  },
  {
    name: "DeclareMathOperator definitions",
    format: "latex",
    source: String.raw`\DeclareMathOperator{\Spec}{Spec} $\Spec(A)$`,
    check: (parsed) => formulas(parsed).some((run) =>
      run.latex.includes(String.raw`\operatorname{Spec}(A)`),
    ),
  },
  {
    name: "DeclarePairedDelimiter definitions",
    format: "latex",
    source: String.raw`\DeclarePairedDelimiter{\ceil}{\lceil}{\rceil} $\ceil*{x+y}$`,
    check: (parsed) => formulas(parsed).some((run) =>
      run.latex.includes(String.raw`\left\lceilx+y\right\rceil`),
    ),
  },
  {
    name: "preamble commands in fragments are hidden",
    format: "auto",
    source: String.raw`\documentclass{article}
\usepackage[amsmath]{mathtools}
\geometry{margin=2cm}
\begin{proof}Visible body.\end{proof}`,
    check: (parsed) =>
      parsed.format === "latex" &&
      text(parsed).includes("Visible body.") &&
      !text(parsed).includes("documentclass") &&
      !text(parsed).includes("usepackage") &&
      !text(parsed).includes("margin=2cm"),
  },
  {
    name: "formula labels and no-number controls are removed",
    format: "latex",
    source: String.raw`\begin{align}
a&=b\label{eq:a}\notag\\
c&=d\nonumber
\end{align}`,
    check: (parsed) => {
      const formula = formulas(parsed)[0]?.latex ?? "";
      return formula.includes("a&=b") && formula.includes("c&=d") &&
        !formula.includes(String.raw`\label`) &&
        !formula.includes(String.raw`\notag`) &&
        !formula.includes(String.raw`\nonumber`);
    },
  },
  {
    name: "verb and verb star commands",
    format: "latex",
    source: String.raw`\verb|a_b%$| and \verb*+x y+`,
    check: (parsed) => {
      const codeRuns = runs(parsed).filter((run) => run.kind === "text" && run.code);
      return codeRuns.length === 2 && codeRuns[0].text === "a_b%$" && codeRuns[1].text === "x y";
    },
  },
  {
    name: "underline and strike formatting commands",
    format: "latex",
    source: String.raw`\underline{under} \uline{again} \sout{gone}`,
    check: (parsed) =>
      runs(parsed).some((run) => run.kind === "text" && run.underline && run.text === "under") &&
      runs(parsed).some((run) => run.kind === "text" && run.underline && run.text === "again") &&
      runs(parsed).some((run) => run.kind === "text" && run.strike && run.text === "gone"),
  },
  {
    name: "color and transform boxes preserve visible arguments",
    format: "latex",
    source: String.raw`\textcolor{red}{Red} \colorbox{yellow}{Box} \fcolorbox{black}{white}{Frame} \scalebox{2}{Scale} \resizebox{3cm}{!}{Resize} \rotatebox{90}{Rotate}`,
    check: (parsed) =>
      ["Red", "Box", "Frame", "Scale", "Resize", "Rotate"].every((value) => text(parsed).includes(value)) &&
      !text(parsed).includes("yellow") &&
      !text(parsed).includes("3cm"),
  },
  {
    name: "box commands with optional arguments preserve content",
    format: "latex",
    source: String.raw`\makebox[3cm][c]{Make} \framebox[2cm]{Frame} \parbox{4cm}{Paragraph} \raisebox{1ex}{Raised}`,
    check: (parsed) =>
      ["Make", "Frame", "Paragraph", "Raised"].every((value) => text(parsed).includes(value)) &&
      !text(parsed).includes("3cm"),
  },
  {
    name: "optional section and list arguments",
    format: "latex",
    source: String.raw`\section[Short]{Long title}
\begin{enumerate}[label=(\alph*)]
\item First
\end{enumerate}`,
    check: (parsed) =>
      blocksOf(parsed, "heading").some((block) => block.runs.some((run) => run.kind === "text" && run.text === "Long title")) &&
      blocksOf(parsed, "numbered").length === 1 &&
      text(parsed).includes("First"),
  },
  {
    name: "minipage wrappers preserve body",
    format: "latex",
    source: String.raw`\begin{minipage}[t]{0.45\textwidth}Left $x$\end{minipage}`,
    check: (parsed) => text(parsed).includes("Left") && formulas(parsed).some((run) => run.latex === "x"),
  },
  {
    name: "thebibliography and bibitem entries",
    format: "latex",
    source: String.raw`\begin{thebibliography}{9}
\bibitem[Knuth84]{knuth} Donald Knuth, TeXbook.
\end{thebibliography}`,
    check: (parsed) =>
      blocksOf(parsed, "heading").length === 1 &&
      blocksOf(parsed, "bullet").length === 1 &&
      text(parsed).includes("Knuth84") &&
      text(parsed).includes("TeXbook"),
  },
  {
    name: "external input and bibliography resources degrade visibly",
    format: "latex",
    source: String.raw`\input{chapter1} \include{appendix} \bibliography{refs} \addbibresource{more.bib}`,
    check: (parsed) =>
      text(parsed).includes("外部 LaTeX 文件：chapter1") &&
      text(parsed).includes("外部 LaTeX 文件：appendix") &&
      text(parsed).includes("参考文献数据：refs") &&
      text(parsed).includes("参考文献数据：more.bib"),
  },
  {
    name: "two argument custom macros",
    format: "latex",
    source: String.raw`\newcommand{\pair}[2]{#1 and #2}\begin{document}\pair{left}{right}\end{document}`,
    check: (parsed) => text(parsed).includes("left and right"),
  },
  {
    name: "unknown environments preserve visible body with warning",
    format: "latex",
    source: String.raw`\begin{custombox}Visible $x+1$\end{custombox}`,
    check: (parsed) =>
      text(parsed).includes("Visible") &&
      formulas(parsed).some((run) => run.latex === "x+1") &&
      parsed.warnings.some((warning) => warning.includes("custombox")),
  },
];

function evaluate(cases: CoverageCase[]) {
  return cases.map((testCase) => {
    try {
      const parsed = parseDocumentImport(testCase.source, testCase.format);
      return {
        name: testCase.name,
        supported: Boolean(testCase.check(parsed)),
        blocks: parsed.blocks.length,
        formulas: parsed.formulaCount,
        warnings: parsed.warnings,
        text: text(parsed),
        blockKinds: parsed.blocks.map((block) => `${block.kind}:${block.level}`),
      };
    } catch (error) {
      return {
        name: testCase.name,
        supported: false,
        error: error instanceof Error ? error.message : String(error),
      };
    }
  });
}

function summarize(label: string, results: ReturnType<typeof evaluate>) {
  const supported = results.filter((result) => result.supported).length;
  return {
    label,
    supported,
    total: results.length,
    unsupported: results.filter((result) => !result.supported).map((result) => result.name),
  };
}

const results = {
  markdownBaseline: evaluate(markdownBaseline),
  latexBaseline: evaluate(latexBaseline),
  markdownExtended: evaluate(markdownExtended),
  latexExtended: evaluate(latexExtended),
};

// The installed import window serializes this parser's blocks to Word. Keep
// its heading contract identical to the VSTO source-parser fallback.
const headingCases: Array<{ source: string; levels: number[] }> = JSON.parse(
  readFileSync(new URL("../test-fixtures/document-import-heading-hierarchy.json", import.meta.url), "utf8"),
);
for (const fixture of headingCases) {
  assert.deepEqual(blocksOf(parseDocumentImport(fixture.source, "latex"), "heading")
    .map((block) => block.level), fixture.levels, fixture.source);
}
console.log(`Shared Word/frontend heading cases: ${headingCases.length} passed`);

const summary = [
  summarize("Markdown baseline", results.markdownBaseline),
  summarize("LaTeX baseline", results.latexBaseline),
  summarize("Markdown extended/common dialect features", results.markdownExtended),
  summarize("LaTeX extended/document features", results.latexExtended),
];

const showDetails = process.argv.includes("--details");
console.log(JSON.stringify(showDetails ? { summary, results } : { summary }, null, 2));

assert.equal(
  results.markdownBaseline.filter((result) => result.supported).length,
  markdownBaseline.length,
  "Markdown baseline coverage regressed",
);
assert.equal(
  results.latexBaseline.filter((result) => result.supported).length,
  latexBaseline.length,
  "LaTeX baseline coverage regressed",
);
assert.equal(
  results.markdownExtended.filter((result) => result.supported).length,
  markdownExtended.length,
  "Markdown extended coverage regressed",
);
assert.equal(
  results.latexExtended.filter((result) => result.supported).length,
  latexExtended.length,
  "LaTeX extended coverage regressed",
);
