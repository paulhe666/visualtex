import assert from "node:assert/strict";
import { DOMParser } from "@xmldom/xmldom";
import { parseLatexMarkdownDocument } from "../src/office/documentImport/documentImportParser.ts";
import { normalizeFormulaEditorDocument } from "../src/office/shared/formulaEditorDocument.ts";
import { renderOfficeFormulaArtifacts } from "../src/office/shared/formulaRenderArtifacts.ts";

globalThis.DOMParser ??= DOMParser;
const domProbe = new DOMParser().parseFromString("<root/>", "application/xml");
const documentPrototype = Object.getPrototypeOf(domProbe);
const elementPrototype = Object.getPrototypeOf(domProbe.documentElement);
if (typeof documentPrototype.querySelector !== "function") {
  documentPrototype.querySelector = function querySelector(name) {
    return this.getElementsByTagName(name)?.item(0) ?? null;
  };
}
if (!("children" in elementPrototype)) {
  Object.defineProperty(elementPrototype, "children", {
    configurable: true,
    get() {
      return Array.from(this.childNodes ?? []).filter((node) => node.nodeType === 1);
    },
  });
}

const allLineEndings = ["lf", "crlf", "cr"];

const fixtures = [
  {
    name: "latex-inline-dollar",
    kind: "latex",
    source: "正文 $a^2+b^2=c^2$ 结尾。",
    expected: [{ mode: "inline", contains: "a^2+b^2" }],
  },
  {
    name: "latex-inline-paren",
    kind: "latex",
    source: String.raw`正文 \(E=mc^2\) 结尾。`,
    expected: [{ mode: "inline", contains: "E=mc^2" }],
  },
  {
    name: "latex-inline-math-environment",
    kind: "latex",
    source: String.raw`正文 \begin{math}\alpha+\beta\end{math} 结尾。`,
    expected: [{ mode: "inline", contains: "\\alpha+\\beta" }],
  },
  {
    name: "latex-display-dollar-single-line",
    kind: "latex",
    source: "前文\n$$E=mc^2$$\n后文",
    lineEndings: allLineEndings,
    expected: [{ mode: "block", numbered: false, contains: "E=mc^2" }],
  },
  {
    name: "latex-display-dollar-multiline",
    kind: "latex",
    source: "$$\n\\frac{1}{2}mv^2\n$$",
    lineEndings: allLineEndings,
    expected: [{ mode: "block", contains: "\\frac{1}{2}" }],
  },
  {
    name: "latex-display-bracket",
    kind: "latex",
    source: String.raw`\[
\int_0^1 x^2\,dx
\]`,
    lineEndings: allLineEndings,
    expected: [{ mode: "block", contains: "\\int_0^1" }],
  },
  {
    name: "latex-equation",
    kind: "latex",
    source: String.raw`\begin{equation}
F=ma
\end{equation}`,
    lineEndings: allLineEndings,
    expected: [{ mode: "block", numbered: true, contains: "\\begin{equation}" }],
  },
  {
    name: "latex-equation-star",
    kind: "latex",
    source: String.raw`\begin{equation*}
E=mc^2
\end{equation*}`,
    expected: [{ mode: "block", numbered: false, contains: "equation*" }],
  },
  {
    name: "latex-align",
    kind: "latex",
    source: String.raw`\begin{align}
a&=b+c\\
d&=e-f
\end{align}`,
    lineEndings: allLineEndings,
    expected: [{ mode: "block", numbered: true, contains: "\\begin{align}" }],
  },
  {
    name: "latex-align-star",
    kind: "latex",
    source: String.raw`\begin{align*}
x&=y\\
y&=z
\end{align*}`,
    expected: [{ mode: "block", numbered: false, contains: "align*" }],
  },
  {
    name: "latex-alignat",
    kind: "latex",
    source: String.raw`\begin{alignat}{2}
a&=b&\qquad c&=d\\
e&=f& g&=h
\end{alignat}`,
    expected: [{ mode: "block", numbered: true, contains: "alignat" }],
  },
  {
    name: "latex-alignat-star",
    kind: "latex",
    source: String.raw`\begin{alignat*}{2}
a&=b& c&=d\\
e&=f& g&=h
\end{alignat*}`,
    expected: [{ mode: "block", numbered: false, contains: "alignat*" }],
  },
  {
    name: "latex-flalign",
    kind: "latex",
    source: String.raw`\begin{flalign}
a&=b+c&&\\
d&=e-f&&
\end{flalign}`,
    expected: [{ mode: "block", numbered: true, contains: "flalign" }],
  },
  {
    name: "latex-flalign-star",
    kind: "latex",
    source: String.raw`\begin{flalign*}
a&=b&&\\
c&=d&&
\end{flalign*}`,
    expected: [{ mode: "block", numbered: false, contains: "flalign*" }],
  },
  {
    name: "latex-eqnarray",
    kind: "latex",
    source: String.raw`\begin{eqnarray}
a&=&b+c\\
d&=&e-f
\end{eqnarray}`,
    expected: [{ mode: "block", numbered: true, contains: "eqnarray" }],
  },
  {
    name: "latex-gather",
    kind: "latex",
    source: String.raw`\begin{gather}
a=b\\
c=d
\end{gather}`,
    expected: [{ mode: "block", numbered: true, contains: "gather" }],
  },
  {
    name: "latex-gather-star",
    kind: "latex",
    source: String.raw`\begin{gather*}
a=b\\
c=d
\end{gather*}`,
    expected: [{ mode: "block", numbered: false, contains: "gather*" }],
  },
  {
    name: "latex-multline",
    kind: "latex",
    source: String.raw`\begin{multline}
a+b+c+d+e\\
=f+g+h
\end{multline}`,
    expected: [{ mode: "block", numbered: true, contains: "multline" }],
  },
  {
    name: "latex-multline-star",
    kind: "latex",
    source: String.raw`\begin{multline*}
a+b+c\\
=d+e
\end{multline*}`,
    expected: [{ mode: "block", numbered: false, contains: "multline*" }],
  },
  {
    name: "latex-equation-split",
    kind: "latex",
    source: String.raw`\begin{equation}
\begin{split}
a&=b+c\\
 &=d+e
\end{split}
\end{equation}`,
    lineEndings: allLineEndings,
    expected: [{ mode: "block", numbered: true, contains: "\\begin{split}" }],
  },
  {
    name: "latex-matrix",
    kind: "latex",
    source: String.raw`\[
A=\begin{matrix}a&b\\c&d\end{matrix}
\]`,
    expected: [{ mode: "block", contains: "\\begin{matrix}" }],
  },
  {
    name: "latex-pmatrix",
    kind: "latex",
    source: String.raw`\[A=\begin{pmatrix}1&2\\3&4\end{pmatrix}\]`,
    expected: [{ mode: "block", contains: "pmatrix" }],
  },
  {
    name: "latex-bmatrix",
    kind: "latex",
    source: String.raw`\[A=\begin{bmatrix}1&0\\0&1\end{bmatrix}\]`,
    expected: [{ mode: "block", contains: "bmatrix" }],
  },
  {
    name: "latex-vmatrix-determinant",
    kind: "latex",
    source: String.raw`\[\det A=\begin{vmatrix}a&b\\c&d\end{vmatrix}\]`,
    expected: [{ mode: "block", contains: "vmatrix" }],
  },
  {
    name: "latex-cases",
    kind: "latex",
    source: String.raw`\[f(x)=\begin{cases}x^2,&x\ge0\\-x,&x<0\end{cases}\]`,
    expected: [{ mode: "block", contains: "cases" }],
  },
  {
    name: "latex-array",
    kind: "latex",
    source: String.raw`\[\begin{array}{c|cc}&A&B\\x&1&2\end{array}\]`,
    expected: [{ mode: "block", contains: "\\begin{array}" }],
  },
  {
    name: "latex-alignedat-inside-brackets",
    kind: "latex",
    source: String.raw`\[\begin{alignedat}{2}a&=b&\quad c&=d\\e&=f&g&=h\end{alignedat}\]`,
    expected: [{ mode: "block", contains: "alignedat" }],
  },
  {
    name: "latex-nested-fractions-roots",
    kind: "latex",
    source: String.raw`\[\frac{1}{1+\frac{x}{\sqrt{1+x^2}}}\]`,
    expected: [{ mode: "block", contains: "\\sqrt" }],
  },
  {
    name: "latex-accents-decorators",
    kind: "latex",
    source: String.raw`\[\hat{x}+\bar{y}+\vec{v}+\overset{!}{=}+\underbrace{a+b}_{c}\]`,
    expected: [{ mode: "block", contains: "\\underbrace" }],
  },
  {
    name: "latex-operators-integrals",
    kind: "latex",
    source: String.raw`\[\sum_{n=1}^{\infty}\int_{-\infty}^{\infty}e^{-x^2}\,dx\]`,
    expected: [{ mode: "block", contains: "\\sum" }],
  },
  {
    name: "latex-label-and-tag-removal",
    kind: "latex",
    source: String.raw`\begin{equation}
E=mc^2\label{eq:e}\tag{A}
\end{equation}`,
    expected: [{ mode: "block", numbered: true, excludes: ["\\label", "\\tag"] }],
  },
  {
    name: "latex-comments-and-escaped-percent",
    kind: "latex",
    source: String.raw`正文 50\% 与 $x=1$。 % remove this
\[
y=2 % remove formula comment
\]`,
    lineEndings: allLineEndings,
    expected: [
      { mode: "inline", contains: "x=1" },
      { mode: "block", contains: "y=2" },
    ],
  },
  {
    name: "latex-escaped-symbols",
    kind: "latex",
    source: String.raw`价格 \$5，比例 50\%，变量 $a\_1+b\&c$。`,
    expected: [{ mode: "inline", contains: "a\\_1" }],
  },
  {
    name: "latex-complete-document",
    kind: "auto",
    source: String.raw`\documentclass[12pt]{article}
\usepackage{amsmath}
\title{测试文档}
\author{VisualTeX}
\begin{document}
\maketitle
\section{公式}
正文 $x=1$。
\begin{equation*}
y=2
\end{equation*}
\end{document}`,
    lineEndings: allLineEndings,
    expected: [
      { mode: "inline", contains: "x=1" },
      { mode: "block", numbered: false, contains: "equation*" },
    ],
    textIncludes: ["测试文档", "VisualTeX", "公式"],
  },
  {
    name: "latex-itemize-inline-formulas",
    kind: "auto",
    source: String.raw`\begin{itemize}
\item 第一项 $a=1$
\item 第二项 \(b=2\)
\end{itemize}`,
    expected: [
      { mode: "inline", contains: "a=1" },
      { mode: "inline", contains: "b=2" },
    ],
  },
  {
    name: "latex-nested-lists",
    kind: "latex",
    source: String.raw`\begin{enumerate}
\item 外层
\begin{itemize}
\item 内层 $x=1$
\end{itemize}
\item 末项
\end{enumerate}`,
    expected: [{ mode: "inline", contains: "x=1" }],
  },
  {
    name: "latex-quote-and-center",
    kind: "latex",
    source: String.raw`\begin{quote}
引用包含 $q=mv$。
\end{quote}
\begin{center}
\[E=mc^2\]
\end{center}`,
    expected: [
      { mode: "inline", contains: "q=mv" },
      { mode: "block", contains: "E=mc^2" },
    ],
  },
  {
    name: "latex-multiple-inline-one-paragraph",
    kind: "latex",
    source: "先 $a=1$，再 $b=2$，最后 \\(c=3\\)。",
    expected: [
      { mode: "inline", contains: "a=1" },
      { mode: "inline", contains: "b=2" },
      { mode: "inline", contains: "c=3" },
    ],
  },
  {
    name: "latex-flexible-environment-spacing",
    kind: "latex",
    source: String.raw`\begin {equation}
E=mc^2
\end {equation}`,
    expected: [{ mode: "block", numbered: true, contains: "E=mc^2" }],
  },
  {
    name: "latex-verbatim-protected",
    kind: "latex",
    source: String.raw`\begin{verbatim}
$not_math$ and \[not_math\]
\end{verbatim}
正文 $x=1$。`,
    expected: [{ mode: "inline", contains: "x=1" }],
    forbiddenFormulaText: ["not_math"],
  },
  {
    name: "latex-verb-protected",
    kind: "latex",
    source: String.raw`命令 \verb|$not_math$| 后面是 $x=1$。`,
    expected: [{ mode: "inline", contains: "x=1" }],
    forbiddenFormulaText: ["not_math"],
  },
  {
    name: "latex-minted-protected",
    kind: "latex",
    source: String.raw`\begin{minted}{tex}
$x$ and \[y\]
\end{minted}
\[z=3\]`,
    expected: [{ mode: "block", contains: "z=3" }],
    forbiddenFormulaText: ["$x$", "y"],
  },
  {
    name: "auto-markdown-with-matrix-environment",
    kind: "auto",
    source: String.raw`**粗体 Markdown** 与矩阵：

\[
A=\begin{matrix}1&2\\3&4\end{matrix}
\]`,
    expected: [{ mode: "block", contains: "\\begin{matrix}" }],
    textIncludes: ["粗体 Markdown"],
    textExcludes: ["**"],
  },
  {
    name: "markdown-heading-inline",
    kind: "markdown",
    source: "# 标题\n\n正文含 $x=1$。",
    lineEndings: allLineEndings,
    expected: [{ mode: "inline", contains: "x=1" }],
    textIncludes: ["标题"],
  },
  {
    name: "markdown-display-dollar",
    kind: "markdown",
    source: "前文\n\n$$\nE=mc^2\n$$\n\n后文",
    lineEndings: allLineEndings,
    expected: [{ mode: "block", contains: "E=mc^2" }],
  },
  {
    name: "markdown-display-bracket",
    kind: "markdown",
    source: String.raw`说明

\[
\frac{a}{b}
\]`,
    expected: [{ mode: "block", contains: "\\frac{a}{b}" }],
  },
  {
    name: "markdown-bullet-list",
    kind: "markdown",
    source: "- 第一项 $a=1$\n- 第二项 $b=2$",
    expected: [
      { mode: "inline", contains: "a=1" },
      { mode: "inline", contains: "b=2" },
    ],
  },
  {
    name: "markdown-nested-list",
    kind: "markdown",
    source: "- 外层\n  - 内层 $x=1$\n    - 更深 $y=2$",
    expected: [
      { mode: "inline", contains: "x=1" },
      { mode: "inline", contains: "y=2" },
    ],
  },
  {
    name: "markdown-numbered-list",
    kind: "markdown",
    source: "1. 第一项 $a=1$\n2) 第二项 $b=2$",
    expected: [
      { mode: "inline", contains: "a=1" },
      { mode: "inline", contains: "b=2" },
    ],
  },
  {
    name: "markdown-quote",
    kind: "markdown",
    source: "> 引用中的公式 $q=mv$\n> 第二行",
    expected: [{ mode: "inline", contains: "q=mv" }],
  },
  {
    name: "markdown-fenced-code-protected",
    kind: "markdown",
    source: "```tex\n$not_math$\n\\[not_math\\]\n```\n\n正文 $x=1$。",
    expected: [{ mode: "inline", contains: "x=1" }],
    forbiddenFormulaText: ["not_math"],
  },
  {
    name: "markdown-tilde-fence-protected",
    kind: "markdown",
    source: "~~~latex\n$$not_math$$\n~~~\n\n正文 $x=1$。",
    expected: [{ mode: "inline", contains: "x=1" }],
    forbiddenFormulaText: ["not_math"],
  },
  {
    name: "markdown-inline-code-protected",
    kind: "markdown",
    source: "代码 `$not_math$`，真实公式 $x=1$。",
    expected: [{ mode: "inline", contains: "x=1" }],
    forbiddenFormulaText: ["not_math"],
  },
  {
    name: "markdown-links-emphasis",
    kind: "markdown",
    source: "**粗体**、*斜体*、[链接](https://example.com) 与 $x=1$。",
    expected: [{ mode: "inline", contains: "x=1" }],
    textIncludes: ["粗体", "斜体", "链接"],
  },
  {
    name: "markdown-multiple-inline",
    kind: "markdown",
    source: "同一段 $a=1$、$b=2$ 和 \\(c=3\\)。",
    expected: [
      { mode: "inline", contains: "a=1" },
      { mode: "inline", contains: "b=2" },
      { mode: "inline", contains: "c=3" },
    ],
  },
  {
    name: "markdown-escaped-dollar",
    kind: "markdown",
    source: "价格是 \\$5，不是公式；真正公式 $x=1$。",
    expected: [{ mode: "inline", contains: "x=1" }],
  },
  {
    name: "markdown-table",
    kind: "markdown",
    source: "| 量 | 值 |\n|---|---|\n| 能量 | $E=mc^2$ |",
    expected: [{ mode: "inline", contains: "E=mc^2" }],
  },
  {
    name: "markdown-mixed-delimiters",
    kind: "markdown",
    source: String.raw`行内 $a=1$ 与 \(b=2\)。

$$c=3$$

\[d=4\]`,
    expected: [
      { mode: "inline", contains: "a=1" },
      { mode: "inline", contains: "b=2" },
      { mode: "block", contains: "c=3" },
      { mode: "block", contains: "d=4" },
    ],
  },
  {
    name: "markdown-multiline-display-body",
    kind: "markdown",
    source: "$$\n\\begin{aligned}\na&=b\\\\\nc&=d\n\\end{aligned}\n$$",
    expected: [{ mode: "block", contains: "\\begin{aligned}" }],
  },
  {
    name: "markdown-indented-code-protected",
    kind: "markdown",
    source: "    $not_math$\n\n正文 $x=1$。",
    expected: [{ mode: "inline", contains: "x=1" }],
    forbiddenFormulaText: ["not_math"],
  },
  {
    name: "markdown-crlf-paragraphs",
    kind: "markdown",
    source: "第一段 $a=1$。\n\n第二段。\n\n$$b=2$$",
    lineEndings: allLineEndings,
    expected: [
      { mode: "inline", contains: "a=1" },
      { mode: "block", contains: "b=2" },
    ],
  },
  {
    name: "latex-preamble-and-custom-command-fallback",
    kind: "auto",
    source: String.raw`\documentclass[12pt]{article}
\usepackage{amsmath,physics}
\newcommand{\vect}[1]{\mathbf{#1}}
\begin{document}
正文中的标准公式
\[
E=mc^2
\]
\end{document}`,
    lineEndings: allLineEndings,
    expected: [{ mode: "block", contains: "E=mc^2" }],
    textIncludes: [
      String.raw`\documentclass[12pt]{article}`,
      String.raw`\usepackage{amsmath,physics}`,
      String.raw`\newcommand{\vect}[1]{\mathbf{#1}}`,
    ],
    codeTextIncludes: [String.raw`\usepackage{amsmath,physics}`],
  },
  {
    name: "latex-multiline-newcommand-protects-inner-math",
    kind: "latex",
    source: String.raw`\newcommand{\custombox}[1]{
  原样内容 $not_a_document_formula$：#1
}
正文 $x=1$。`,
    expected: [{ mode: "inline", contains: "x=1" }],
    forbiddenFormulaText: ["not_a_document_formula"],
    textIncludes: [
      String.raw`\newcommand{\custombox}[1]{`,
      String.raw`原样内容 $not_a_document_formula$：#1`,
    ],
  },
  {
    name: "latex-custom-environment-preserved-literally",
    kind: "auto",
    source: String.raw`\begin{mytheorem}[自定义标题]
这里的 $not_parsed$ 和 \begin{equation}z=9\end{equation} 都属于自定义内容。
\end{mytheorem}
正文 \(y=2\)。`,
    expected: [{ mode: "inline", contains: "y=2" }],
    forbiddenFormulaText: ["not_parsed", "z=9"],
    textIncludes: [
      String.raw`\begin{mytheorem}[自定义标题]`,
      String.raw`\end{mytheorem}`,
      String.raw`$not_parsed$`,
    ],
    codeTextIncludes: [String.raw`\begin{mytheorem}[自定义标题]`],
  },
  {
    name: "latex-newenvironment-definition-preserved",
    kind: "latex",
    source: String.raw`\newenvironment{warningbox}
  {\begin{quote}\bfseries}
  {\end{quote}}
正文 $a=b$。`,
    expected: [{ mode: "inline", contains: "a=b" }],
    textIncludes: [
      String.raw`\newenvironment{warningbox}`,
      String.raw`{\begin{quote}\bfseries}`,
      String.raw`{\end{quote}}`,
    ],
  },
  {
    name: "latex-makeatletter-block-preserved",
    kind: "latex",
    source: String.raw`\makeatletter
\def\custom@name#1{value-$not_math$-#1}
\makeatother
\[q=3\]`,
    expected: [{ mode: "block", contains: "q=3" }],
    forbiddenFormulaText: ["not_math"],
    textIncludes: [
      String.raw`\makeatletter`,
      String.raw`\def\custom@name#1{value-$not_math$-#1}`,
      String.raw`\makeatother`,
    ],
  },
  {
    name: "latex-expl3-block-preserved",
    kind: "latex",
    source: String.raw`\ExplSyntaxOn
\cs_new:Npn \my_macro:n #1 { raw_$not_math$_#1 }
\ExplSyntaxOff
正文 $r=4$。`,
    expected: [{ mode: "inline", contains: "r=4" }],
    forbiddenFormulaText: ["not_math"],
    textIncludes: [
      String.raw`\ExplSyntaxOn`,
      String.raw`\cs_new:Npn \my_macro:n #1 { raw_$not_math$_#1 }`,
      String.raw`\ExplSyntaxOff`,
    ],
  },
  {
    name: "latex-inline-unknown-command-preserved",
    kind: "latex",
    source: String.raw`正文保留 \myterm[opt]{alpha_beta}，公式为 $k=5$。`,
    expected: [{ mode: "inline", contains: "k=5" }],
    textIncludes: [String.raw`\myterm[opt]{alpha_beta}`],
  },
  {
    name: "latex-package-options-and-declarations-preserved",
    kind: "latex",
    source: String.raw`\RequirePackage{xcolor}
\PassOptionsToPackage{unicode}{hyperref}
\DeclareMathOperator*{\argmax}{arg\,max}
\definecolor{brand}{RGB}{10,20,30}
正文 $s=6$。`,
    expected: [{ mode: "inline", contains: "s=6" }],
    textIncludes: [
      String.raw`\RequirePackage{xcolor}`,
      String.raw`\PassOptionsToPackage{unicode}{hyperref}`,
      String.raw`\DeclareMathOperator*{\argmax}{arg\,max}`,
      String.raw`\definecolor{brand}{RGB}{10,20,30}`,
    ],
  },
  {
    name: "latex-unknown-figure-environment-preserved",
    kind: "latex",
    source: String.raw`\begin{figure}[htbp]
\centering
\includegraphics{custom.pdf}
\caption{图中写有 $not_formula$}
\end{figure}
正文 $t=7$。`,
    expected: [{ mode: "inline", contains: "t=7" }],
    forbiddenFormulaText: ["not_formula"],
    textIncludes: [
      String.raw`\begin{figure}[htbp]`,
      String.raw`\includegraphics{custom.pdf}`,
      String.raw`\caption{图中写有 $not_formula$}`,
      String.raw`\end{figure}`,
    ],
  },
  {
    name: "latex-custom-content-mixed-with-known-structure",
    kind: "auto",
    source: String.raw`\usepackage{custompkg}
\section{可识别标题}
\customnote{这一行原样保留}
\begin{itemize}
\item 列表公式 $u=8$
\end{itemize}`,
    expected: [{ mode: "inline", contains: "u=8" }],
    textIncludes: [
      String.raw`\usepackage{custompkg}`,
      String.raw`\customnote{这一行原样保留}`,
      "可识别标题",
    ],
  },
  {
    name: "latex-built-in-theorem-structured",
    kind: "auto",
    source: String.raw`\begin{theorem}[勾股定理]
设直角三角形满足 $a^2+b^2=c^2$。
\[
A=\frac{1}{2}ab
\]
\end{theorem}`,
    lineEndings: allLineEndings,
    expected: [
      { mode: "inline", contains: "a^2+b^2=c^2" },
      { mode: "block", contains: "\\frac{1}{2}" },
    ],
    headingTextIncludes: ["定理 1（勾股定理）"],
    styledTextIncludes: [
      { style: "quote", text: "设直角三角形满足" },
    ],
    formulaStyleExpectations: [
      { contains: "a^2+b^2=c^2", style: "quote" },
    ],
  },
  {
    name: "latex-proof-structured-and-qed",
    kind: "auto",
    source: String.raw`\begin{proof}[充分性]
由 $x>0$ 可立即得到结论。\qedhere
\end{proof}`,
    expected: [{ mode: "inline", contains: "x>0" }],
    headingTextIncludes: ["证明（充分性）"],
    styledTextIncludes: [
      { style: "normal", text: "由" },
      { style: "normal", text: "□" },
    ],
    formulaStyleExpectations: [{ contains: "x>0", style: "normal" }],
  },
  {
    name: "latex-dynamic-newtheorem-structured",
    kind: "auto",
    source: String.raw`\newtheorem{thm}{自定义定理}
\begin{thm}[谱定理]
正文含有行内公式 $Av=\lambda v$。
\end{thm}`,
    expected: [{ mode: "inline", contains: "Av=\\lambda v" }],
    headingTextIncludes: ["自定义定理 1（谱定理）"],
    styledTextIncludes: [{ style: "quote", text: "正文含有行内公式" }],
    codeTextIncludes: [String.raw`\newtheorem{thm}{自定义定理}`],
  },
  {
    name: "latex-newtheorem-shared-counter",
    kind: "latex",
    source: String.raw`\newtheorem{thm}{定理}
\newtheorem{lem}[thm]{引理}
\begin{thm}第一个结论。\end{thm}
\begin{lem}第二个结论。\end{lem}`,
    expected: [],
    headingTextIncludes: ["定理 1", "引理 2"],
    styledTextIncludes: [
      { style: "quote", text: "第一个结论" },
      { style: "quote", text: "第二个结论" },
    ],
    codeTextIncludes: [String.raw`\newtheorem{lem}[thm]{引理}`],
  },
  {
    name: "latex-newtheorem-star-unnumbered",
    kind: "latex",
    source: String.raw`\newtheorem*{specialremark}{特别说明}
\begin{specialremark}[边界情况]
这一段不应自动编号。
\end{specialremark}`,
    expected: [],
    headingTextIncludes: ["特别说明（边界情况）"],
    headingTextExcludes: ["特别说明 1"],
    styledTextIncludes: [
      { style: "quote", text: "这一段不应自动编号" },
    ],
  },
  {
    name: "latex-common-theorem-family",
    kind: "auto",
    source: String.raw`\begin{lemma}引理正文。\end{lemma}
\begin{proposition}命题正文。\end{proposition}
\begin{corollary}推论正文。\end{corollary}
\begin{definition}定义正文。\end{definition}
\begin{axiom}公理正文。\end{axiom}
\begin{conjecture}猜想正文。\end{conjecture}
\begin{claim}断言正文。\end{claim}
\begin{example}例子正文。\end{example}
\begin{exercise}练习正文。\end{exercise}
\begin{remark}备注正文。\end{remark}
\begin{notation}记号正文。\end{notation}
\begin{solution}解答正文。\end{solution}`,
    expected: [],
    headingTextIncludes: [
      "引理 1",
      "命题 1",
      "推论 1",
      "定义 1",
      "公理 1",
      "猜想 1",
      "断言 1",
      "例 1",
      "练习 1",
      "注",
      "记号",
      "解答",
    ],
    styledTextIncludes: [
      { style: "quote", text: "引理正文" },
      { style: "quote", text: "定义正文" },
      { style: "normal", text: "解答正文" },
    ],
  },
  {
    name: "latex-theorem-list-and-formula",
    kind: "latex",
    source: String.raw`\begin{theorem}
满足以下条件：
\begin{itemize}
\item 第一项 $x=1$
\item 第二项
\end{itemize}
\end{theorem}`,
    expected: [{ mode: "inline", contains: "x=1" }],
    headingTextIncludes: ["定理 1"],
    styledTextIncludes: [
      { style: "quote", text: "满足以下条件" },
      { style: "quote", text: "第一项", listKind: "bullet" },
      { style: "quote", text: "第二项", listKind: "bullet" },
    ],
    formulaStyleExpectations: [
      { contains: "x=1", style: "quote", listKind: "bullet" },
    ],
  },
  {
    name: "latex-theorem-title-with-inline-formula",
    kind: "latex",
    source: String.raw`\begin{theorem}[关于 $f(x)$ 的结论]
正文为 $f(0)=0$。
\end{theorem}`,
    expected: [
      { mode: "inline", contains: "f(x)" },
      { mode: "inline", contains: "f(0)=0" },
    ],
    headingTextIncludes: ["定理 1（关于", "的结论）"],
    formulaStyleExpectations: [
      { contains: "f(x)", style: "heading4" },
      { contains: "f(0)=0", style: "quote" },
    ],
  },
  {
    name: "latex-nested-proof-inside-theorem",
    kind: "latex",
    source: String.raw`\begin{theorem}
定理的第一段。
\begin{proof}
证明中的公式 $y=2$。\qed
\end{proof}
定理的最后一段。
\end{theorem}`,
    expected: [{ mode: "inline", contains: "y=2" }],
    headingTextIncludes: ["定理 1", "证明"],
    styledTextIncludes: [
      { style: "quote", text: "定理的第一段" },
      { style: "normal", text: "证明中的公式" },
      { style: "quote", text: "定理的最后一段" },
    ],
  },
  {
    name: "latex-newtheorem-section-reset-syntax",
    kind: "auto",
    source: String.raw`\newtheorem{result}{Result}[section]
\begin{result}[Local form]
Result body with $z=3$.
\end{result}`,
    lineEndings: allLineEndings,
    expected: [{ mode: "inline", contains: "z=3" }],
    headingTextIncludes: ["Result 1（Local form）"],
    styledTextIncludes: [{ style: "quote", text: "Result body" }],
    codeTextIncludes: [String.raw`\newtheorem{result}{Result}[section]`],
  },
  {
    name: "latex-auto-percent-comment-banner",
    kind: "auto",
    source: String.raw`这就是色散媒质中频域形式的本构关系。

% ==================== % 6. 磁色散媒质中的本构关系 %
====================

如果磁响应也具有色散，则磁化强度为
\[
M(\omega)=\chi_m(\omega)H(\omega)
\]`,
    lineEndings: allLineEndings,
    expected: [{ mode: "block", contains: "\\chi_m" }],
    textIncludes: [
      "这就是色散媒质中频域形式的本构关系。",
      "如果磁响应也具有色散，则磁化强度为",
    ],
    textExcludes: ["磁色散媒质中的本构关系", "====================", "%"],
  },
  {
    name: "latex-escaped-percent-and-inline-comment",
    kind: "latex",
    source: String.raw`效率为 50\%，公式为 $x=1$。 % 这一段应被删除
下一句。`,
    lineEndings: allLineEndings,
    expected: [{ mode: "inline", contains: "x=1" }],
    textIncludes: ["效率为 50%", "下一句。"],
    textExcludes: ["这一段应被删除"],
  },
  {
    name: "latex-cjk-inline-formula-boundary-spacing",
    kind: "latex",
    source: String.raw`对于良导体低频近似，若 \(\varepsilon_r(\omega)\) 的本征极化部分可以忽略，则成立。`,
    expected: [{ mode: "inline", contains: "\\varepsilon_r" }],
    blockSequence: [
      { kind: "text", value: "对于良导体低频近似，若" },
      { kind: "formula", value: String.raw`\varepsilon_r(\omega)` },
      { kind: "text", value: "的本征极化部分可以忽略，则成立。" },
    ],
  },
  {
    name: "markdown-cjk-inline-formula-boundary-spacing",
    kind: "markdown",
    source: String.raw`对于良导体低频近似，若 $\varepsilon_r(\omega)$ 的本征极化部分可以忽略，则成立。`,
    expected: [{ mode: "inline", contains: "\\varepsilon_r" }],
    blockSequence: [
      { kind: "text", value: "对于良导体低频近似，若" },
      { kind: "formula", value: String.raw`\varepsilon_r(\omega)` },
      { kind: "text", value: "的本征极化部分可以忽略，则成立。" },
    ],
  },
  {
    name: "markdown-english-inline-spacing-preserved",
    kind: "markdown",
    source: "Let $x$ be a variable.",
    expected: [{ mode: "inline", contains: "x" }],
    blockSequence: [
      { kind: "text", value: "Let " },
      { kind: "formula", value: "x" },
      { kind: "text", value: " be a variable." },
    ],
  },
  {
    name: "latex-equation-with-inner-aligned-is-idempotent",
    kind: "latex",
    source: String.raw`\begin{equation}
\begin{aligned}
f^{*}(\mathbf{x})
&=
\frac{1}{p(\mathbf{x})}
\int t\,p(\mathbf{x},t)\,\mathrm{d}t  \\
&=
\int t\,p(t\mid\mathbf{x})\,\mathrm{d}t
=
\mathbb{E}_{t}[t\mid\mathbf{x}]
\end{aligned}
\end{equation}`,
    lineEndings: allLineEndings,
    expected: [
      {
        mode: "block",
        numbered: true,
        contains: String.raw`\begin{aligned}`,
        artifactCodeFormat: "equation",
        artifactLineCount: 1,
        canonicalIncludes: [
          String.raw`\begin{equation}`,
          String.raw`\begin{aligned}`,
          String.raw`\end{aligned}`,
          String.raw`\end{equation}`,
        ],
        minSvgWidth: 240,
        maxSvgHeight: 130,
        ommlIncludes: ["<m:eqArr>", "&amp;="],
        ommlExcludes: ["<m:m>", "<m:e></m:e>"],
        ommlCount: { fragment: "&amp;=", count: 2 },
      },
    ],
  },
  {
    name: "latex-equation-star-with-inner-alignedat-is-idempotent",
    kind: "latex",
    source: String.raw`\begin{equation*}
\begin{alignedat}{2}
a&=b &\qquad c&=d \\
e&=f &\qquad g&=h
\end{alignedat}
\end{equation*}`,
    expected: [
      {
        mode: "block",
        numbered: false,
        contains: String.raw`\begin{alignedat}`,
        artifactCodeFormat: "equation-star",
        artifactLineCount: 1,
        canonicalIncludes: [String.raw`\begin{equation*}`, String.raw`\begin{alignedat}`],
        maxSvgHeight: 100,
        ommlIncludes: ["<m:eqArr>", "&amp;="],
        ommlExcludes: ["<m:m>", "<m:e></m:e>"],
      },
    ],
  },
];

function withLineEnding(source, lineEnding) {
  const normalized = source.replace(/\r\n?/g, "\n");
  if (lineEnding === "crlf") return normalized.replace(/\n/g, "\r\n");
  if (lineEnding === "cr") return normalized.replace(/\n/g, "\r");
  return normalized;
}

let executedCases = 0;
let renderedFormulas = 0;
for (const fixture of fixtures) {
  const variants = fixture.lineEndings ?? ["lf"];
  for (const lineEnding of variants) {
    executedCases += 1;
    const label = `${fixture.name}:${lineEnding}`;
    const blocks = parseLatexMarkdownDocument(
      withLineEnding(fixture.source, lineEnding),
      fixture.kind,
      12,
    );
    const formulas = blocks.filter((block) => block.kind === "formula");
    assert.equal(
      formulas.length,
      fixture.expected.length,
      `${label} formula count`,
    );

    for (const [index, expected] of fixture.expected.entries()) {
      const formula = formulas[index];
      assert.equal(formula.displayMode, expected.mode, `${label} formula ${index + 1} mode`);
      if (typeof expected.numbered === "boolean") {
        assert.equal(
          formula.numbered,
          expected.numbered,
          `${label} formula ${index + 1} numbering`,
        );
      }
      if (expected.contains) {
        assert.ok(
          formula.latex.includes(expected.contains),
          `${label} formula ${index + 1} must contain ${expected.contains}`,
        );
      }
      for (const excluded of expected.excludes ?? []) {
        assert.ok(
          !formula.latex.includes(excluded),
          `${label} formula ${index + 1} must remove ${excluded}`,
        );
      }

      const document = normalizeFormulaEditorDocument(
        [{ id: `fixture-${executedCases}-${index}`, latex: formula.latex }],
        "raw",
      );
      let artifacts;
      try {
        artifacts = await renderOfficeFormulaArtifacts({
          lines: document.lines,
          codeFormat: document.codeFormat,
          displayMode: formula.displayMode,
          host: "word",
        });
      } catch (error) {
        throw new Error(
          `${label} formula ${index + 1} artifact generation failed: ${error instanceof Error ? error.message : String(error)}`,
          { cause: error },
        );
      }
      assert.ok(artifacts.canonicalLatex.trim(), `${label} canonical LaTeX`);
      assert.ok(artifacts.svg.svg.includes("<svg"), `${label} SVG artifact`);
      assert.match(
        artifacts.svg.svg,
        /(?:fill|stroke)=["']#000000["']/i,
        `${label} explicit Word black paint`,
      );
      assert.ok(artifacts.omml?.ommlBase64, `${label} OMML artifact`);
      assert.ok(artifacts.omml?.ommlDocxBase64, `${label} OMML DOCX artifact`);
      if (expected.artifactCodeFormat) {
        assert.equal(
          artifacts.codeFormat,
          expected.artifactCodeFormat,
          `${label} artifact code format`,
        );
      }
      if (expected.artifactLineCount) {
        assert.equal(
          artifacts.lines.length,
          expected.artifactLineCount,
          `${label} artifact logical line count`,
        );
      }
      for (const sourceFragment of expected.canonicalIncludes ?? []) {
        assert.ok(
          artifacts.canonicalLatex.includes(sourceFragment),
          `${label} canonical LaTeX must contain ${sourceFragment}`,
        );
      }
      if (expected.minSvgWidth) {
        assert.ok(
          artifacts.svg.width >= expected.minSvgWidth,
          `${label} SVG width ${artifacts.svg.width} must be at least ${expected.minSvgWidth}`,
        );
      }
      if (expected.maxSvgHeight) {
        assert.ok(
          artifacts.svg.height <= expected.maxSvgHeight,
          `${label} SVG height ${artifacts.svg.height} must be at most ${expected.maxSvgHeight}`,
        );
      }
      const omml = artifacts.omml?.omml ?? "";
      if (
        [
          "align",
          "align-star",
          "aligned",
          "equation-split",
          "equation-star-split",
        ].includes(artifacts.codeFormat)
      ) {
        assert.ok(
          omml.includes("<m:eqArr>"),
          `${label} aligned OMML must use a Word equation array`,
        );
        assert.ok(
          !omml.includes("<m:m>"),
          `${label} aligned OMML must not fall back to a Word matrix`,
        );
        assert.ok(
          !omml.includes("<m:e></m:e>"),
          `${label} aligned OMML must not contain an empty placeholder slot`,
        );
      }
      for (const fragment of expected.ommlIncludes ?? []) {
        assert.ok(
          omml.includes(fragment),
          `${label} OMML must contain ${fragment}`,
        );
      }
      for (const fragment of expected.ommlExcludes ?? []) {
        assert.ok(
          !omml.includes(fragment),
          `${label} OMML must not contain ${fragment}`,
        );
      }
      if (expected.ommlCount) {
        assert.equal(
          omml.split(expected.ommlCount.fragment).length - 1,
          expected.ommlCount.count,
          `${label} OMML ${expected.ommlCount.fragment} count`,
        );
      }
      renderedFormulas += 1;
    }

    const allFormulaText = formulas.map((formula) => formula.latex).join("\n");
    for (const forbidden of fixture.forbiddenFormulaText ?? []) {
      assert.ok(
        !allFormulaText.includes(forbidden),
        `${label} protected code must not become a formula: ${forbidden}`,
      );
    }
    const prose = blocks
      .filter((block) => block.kind === "text")
      .map((block) => block.text)
      .join("\n");
    for (const text of fixture.textIncludes ?? []) {
      assert.ok(prose.includes(text), `${label} prose must contain ${text}`);
    }
    for (const text of fixture.textExcludes ?? []) {
      assert.ok(!prose.includes(text), `${label} prose must remove ${text}`);
    }
    if (fixture.blockSequence) {
      assert.deepEqual(
        blocks.map((block) => ({
          kind: block.kind,
          value: block.kind === "text" ? block.text : block.latex,
        })),
        fixture.blockSequence,
        `${label} paragraph/formula boundary sequence`,
      );
    }
    for (const text of fixture.codeTextIncludes ?? []) {
      const literalBlock = blocks.find(
        (block) =>
          block.kind === "text" &&
          block.paragraphStyle === "code" &&
          block.text.includes(text),
      );
      assert.ok(literalBlock, `${label} literal fallback must preserve ${text}`);
    }
    const headingText = blocks
      .filter(
        (block) =>
          block.kind === "text" &&
          ["heading1", "heading2", "heading3", "heading4"].includes(
            block.paragraphStyle ?? "",
          ),
      )
      .map((block) => block.text)
      .join("\n");
    for (const text of fixture.headingTextIncludes ?? []) {
      assert.ok(
        headingText.includes(text),
        `${label} structured heading must contain ${text}`,
      );
    }
    for (const text of fixture.headingTextExcludes ?? []) {
      assert.ok(
        !headingText.includes(text),
        `${label} structured heading must not contain ${text}`,
      );
    }
    for (const expected of fixture.styledTextIncludes ?? []) {
      const matchingBlock = blocks.find(
        (block) =>
          block.kind === "text" &&
          block.paragraphStyle === expected.style &&
          block.text.includes(expected.text) &&
          (expected.listKind === undefined || block.listKind === expected.listKind),
      );
      assert.ok(
        matchingBlock,
        `${label} must contain ${expected.style} text ${expected.text}`,
      );
    }
    for (const expected of fixture.formulaStyleExpectations ?? []) {
      const matchingFormula = formulas.find((formula) =>
        formula.latex.includes(expected.contains),
      );
      assert.ok(
        matchingFormula,
        `${label} must contain formula ${expected.contains}`,
      );
      assert.equal(
        matchingFormula.paragraphStyle,
        expected.style,
        `${label} formula ${expected.contains} paragraph style`,
      );
      if (expected.listKind !== undefined) {
        assert.equal(
          matchingFormula.listKind,
          expected.listKind,
          `${label} formula ${expected.contains} list kind`,
        );
      }
    }
  }
}

assert.ok(fixtures.length >= 40, "The parser corpus must keep at least 40 syntax fixtures");
assert.ok(executedCases >= 60, "Line-ending expansion must execute at least 60 cases");
console.log(
  `Document import parser corpus passed: ${fixtures.length} syntax fixtures, ${executedCases} line-ending cases, ${renderedFormulas} rendered formulas`,
);
