import assert from "node:assert/strict";
import { documentImportErrorMessage } from "../src/office/documentImport/documentImportErrors.ts";
import {
  mergeDocumentImportBlocks,
  parseLatexMarkdownDocument,
} from "../src/office/documentImport/documentImportParser.ts";

assert.equal(
  documentImportErrorMessage({ error: { message: "SVG staging failed" } }, "fallback"),
  "SVG staging failed",
);
assert.notEqual(
  documentImportErrorMessage({ code: 4608, details: { stage: "Word" } }, "fallback"),
  "[object Object]",
);

const markdown = [
  "动量满足 $p=mv$，这是行内公式。",
  "",
  "$$E=mc^2$$",
  "",
  "结尾文字。",
].join("\n");
const markdownBlocks = parseLatexMarkdownDocument(markdown, "markdown", 11);
const markdownFormulas = markdownBlocks.filter((block) => block.kind === "formula");
assert.equal(markdownFormulas.length, 2);
assert.deepEqual(
  markdownFormulas.map((block) => [block.displayMode, block.numbered, block.fontSizePt]),
  [
    ["inline", false, 11],
    ["block", false, 11],
  ],
);
assert.equal(markdownFormulas[0].latex, "p=mv");
assert.equal(markdownFormulas[1].latex, "E=mc^2");
assert.match(
  markdownBlocks.filter((block) => block.kind === "text").map((block) => block.text).join(""),
  /动量满足.*这是行内公式。.*结尾文字。/s,
);

// A space between text and a formula at the paragraph's start or end stays.
const edgeBlocks = parseLatexMarkdownDocument("- item one $x^2$\n\n$y$ and more", "markdown", 12);
assert.deepEqual(
  edgeBlocks.map((block) => (block.kind === "text" ? block.text : `[${block.latex}]`)),
  ["item one ", "[x^2]", "[y]", " and more"],
);

const latex = String.raw`\documentclass{article}
\begin{document}
\section{测试}
正文包含 \(a^2+b^2=c^2\)。
\begin{equation}
E=mc^2\label{eq:energy}
\end{equation}
\begin{equation*}
F=ma
\end{equation*}
\begin{align}
a&=b+c\\
d&=e-f
\end{align}
\end{document}`;
const latexBlocks = parseLatexMarkdownDocument(latex, "latex", 12);
const latexFormulas = latexBlocks.filter((block) => block.kind === "formula");
assert.equal(latexFormulas.length, 4);
assert.deepEqual(
  latexFormulas.map((block) => [block.displayMode, block.numbered]),
  [
    ["inline", false],
    ["block", true],
    ["block", false],
    ["block", true],
  ],
);
assert.equal(latexFormulas[0].latex, "a^2+b^2=c^2");
assert.match(latexFormulas[1].latex, /^\\begin\{equation\}/);
assert.doesNotMatch(latexFormulas[1].latex, /\\label/);
assert.match(latexFormulas[3].latex, /^\\begin\{align\}/);

const customized = latexBlocks.map((block) =>
  block.kind === "formula"
    ? { ...block, fontSizePt: 18, numbered: block.displayMode === "block" }
    : block,
);
const reparsed = parseLatexMarkdownDocument(latex, "latex", 10);
const merged = mergeDocumentImportBlocks(customized, reparsed);
const mergedFormulas = merged.filter((block) => block.kind === "formula");
assert.ok(mergedFormulas.every((block) => block.fontSizePt === 18));
assert.equal(mergedFormulas[0].numbered, false);
assert.ok(mergedFormulas.slice(1).every((block) => block.numbered));

const semanticLatex = String.raw`\documentclass{article}
\begin{document}
\section{结构化标题}
正文中的公式 $x=1$ 与文字属于同一段。
\begin{itemize}
\item 第一项
\item 第二项包含 $y=2$
\end{itemize}
\begin{enumerate}
\item 编号一
\item 编号二
\end{enumerate}
$$E=mc^2$$
公式后的正文。
\end{document}`;
const semanticBlocks = parseLatexMarkdownDocument(semanticLatex, "latex", 12);
const heading = semanticBlocks.find(
  (block) => block.kind === "text" && block.text === "结构化标题",
);
assert.equal(heading?.paragraphStyle, "heading1");
assert.equal(heading?.paragraphStart, true);
assert.equal(heading?.paragraphEnd, true);
const bulletParagraphs = new Set(
  semanticBlocks
    .filter((block) => block.listKind === "bullet")
    .map((block) => block.paragraphId),
);
assert.equal(bulletParagraphs.size, 2);
const numberedParagraphs = new Set(
  semanticBlocks
    .filter((block) => block.listKind === "number")
    .map((block) => block.paragraphId),
);
assert.equal(numberedParagraphs.size, 2);
const inlineListFormula = semanticBlocks.find(
  (block) => block.kind === "formula" && block.latex === "y=2",
);
assert.equal(inlineListFormula?.listKind, "bullet");
assert.equal(inlineListFormula?.paragraphEnd, true);
const displayFormula = semanticBlocks.find(
  (block) => block.kind === "formula" && block.latex === "E=mc^2",
);
assert.equal(displayFormula?.displayMode, "block");
assert.equal(displayFormula?.paragraphId, undefined);
assert.ok(
  semanticBlocks
    .filter((block) => block.kind === "text")
    .every((block) => !block.text.includes("\n\n")),
  "Display-formula boundaries must not preserve blank text paragraphs",
);

const semanticMarkdown = parseLatexMarkdownDocument(
  "# 标题\n\n- 项目一\n- 项目二含 $z=3$\n\n1. 编号一\n2. 编号二",
  "markdown",
  12,
);
assert.equal(semanticMarkdown[0]?.paragraphStyle, "heading1");
assert.equal(
  semanticMarkdown.find((block) => block.kind === "formula" && block.latex === "z=3")
    ?.listKind,
  "bullet",
);
assert.equal(
  semanticMarkdown.filter((block) => block.listKind === "number" && block.paragraphStart)
    .length,
  2,
);

const pastedItemizeFragment = String.raw`多项式逼近和幂级数逼近的区别
	\begin{itemize}
		\item 多项式逼近的逼近系数和原函数原则上说没什么硬性关系，而幂级数的系数是严格依赖原函数在某些点的解析性的
		\item 多项式逼近是全局性的，而幂级数逼近是有收敛半径的
		\item 形式上的区别：幂级数的形式是$\sum_{i=1}^{\infty}{a_k\left( x-x_0 \right) ^k}$而多项式级数的形式是$\sum_{i=1}^{\infty}{a_kP_k\left( x \right)}$
	\end{itemize}
	所以从逼近效果来看，多项式级数逼近要比幂级数逼近要容易得多`;
const pastedBlocks = parseLatexMarkdownDocument(pastedItemizeFragment, "auto", 12);
const pastedBulletParagraphs = new Set(
  pastedBlocks
    .filter((block) => block.listKind === "bullet" && block.paragraphStart)
    .map((block) => block.paragraphId),
);
assert.equal(
  pastedBulletParagraphs.size,
  3,
  "A headerless itemize fragment must auto-detect as LaTeX and create three list items",
);
const pastedFormulas = pastedBlocks.filter((block) => block.kind === "formula");
assert.deepEqual(
  pastedFormulas.map((block) => block.latex),
  [
    String.raw`\sum_{i=1}^{\infty}{a_k\left( x-x_0 \right) ^k}`,
    String.raw`\sum_{i=1}^{\infty}{a_kP_k\left( x \right)}`,
  ],
);
assert.ok(pastedFormulas.every((block) => block.listKind === "bullet"));
assert.equal(pastedFormulas[0]?.paragraphEnd, false);
assert.equal(pastedFormulas[1]?.paragraphEnd, true);
assert.equal(
  pastedBlocks.at(-1)?.kind === "text" ? pastedBlocks.at(-1)?.listKind : undefined,
  "none",
);
assert.ok(
  pastedBlocks
    .filter((block) => block.kind === "text")
    .every((block) => !/[\\](?:begin|end|item)\b/.test(block.text)),
  "LaTeX list commands must never leak into imported Word prose",
);

const longPhysicsNotes = String.raw`laplace算子的表达式：\[
g_{ii}=h_i^2,\qquad \sqrt{g}=h_1h_2h_3
\]

\[
\nabla^2=\frac{1}{h_1h_2h_3}\sum_{i=1}^3\frac{\partial}{\partial x^i}\left(\frac{h_1h_2h_3}{h_i^2}\frac{\partial}{\partial x^i}\right)
\]

\[
\nabla^2=\frac{1}{h_1h_2h_3}\left[
\frac{\partial}{\partial x^1}\left(\frac{h_2h_3}{h_1}\frac{\partial}{\partial x^1}\right)
+\frac{\partial}{\partial x^2}\left(\frac{h_1h_3}{h_2}\frac{\partial}{\partial x^2}\right)
+\frac{\partial}{\partial x^3}\left(\frac{h_1h_2}{h_3}\frac{\partial}{\partial x^3}\right)
\right]
\]
在圆形边界中，本征值只可取离散值，原因可以理解为是因为在周期边界中出现自干涉，导致的量子化条件，和玻尔-索末菲的量子化条件有点像

图伦不变量？是什么

圆形边界内laplace方程的第一类边值问题的通解可以表示为：平面上的电势多极展开
规定原点处有界可以避免多解的情况

利用圆形边界的解的系数形式，可以写成：
\[
u(r,\varphi)=\frac{a^{2}-r^{2}}{2\pi}\int_{0}^{2\pi}\frac{g(\varphi')}{r^{2}+a^{2}-2ar\cos(\varphi-\varphi')}\,d\varphi'
\]

柱贝塞尔方程：\[
\frac{1}{r}\frac{d}{dr}\left(r\frac{dR}{dr}\right)+\left[k^{2}-\lambda-\frac{\mu}{r^{2}}\right]R=0.
\]
球坐标分离变量

legendre方程也是一个SL方程，而且如果$sin\theta$在边界为零，则这个微分算子是自伴的。有奇异性的SL方程？

角向的两个方程一个就是$L^2$和$L_z$，而且它们是对易的。所以他们构成共同本征基。

然后$L_z=-i\hbar\frac{\partial}{\partial \varphi}$,理解方法是我们类比动量的表达式，可以凑出来这个形式，我们也可以直接算出来：由
\begin{equation*}
L=-i\hbar \vec{r} \times \nabla
\end{equation*}
所以
\[
\hat{L}_z = x \hat{p}_y - y \hat{p}_x = - i \hbar \left( x \frac{\partial}{\partial y} - y \frac{\partial}{\partial x} \right)
\]
然后把直角坐标系变成球坐标系即可

\begin{equation*}
L=-\ i\hbar \left(\overrightarrow{e_{\varphi }}\frac{\partial }{\partial \theta } -\overrightarrow{e_{\theta }}\frac{1}{sin\theta }\frac{\partial }{\partial \varphi }\right)
\end{equation*}
坐标变换的时候要注意奇异项的问题，会多一个delta函数的解，所以要引入自然边界条件来解决这个问题
\begin{equation*}
L_{z} =-i\hbar \frac{\partial }{\partial \varphi }
\end{equation*}
角动量算子和旋转对称变换
对于一个x平移变换，我们可以写成
\begin{equation*}
\mathcal{T}( a) =e^{-\frac{i}{\hbar } aP_{x}}
\end{equation*}
一个算符造成物理事实不变的表述是：这个算符和哈密顿量对易，然后根据平移算符导致哈密顿量不变，所以推出动量算符和哈密顿量对易。

角动量算符：
\[
R_{n}(\theta)=\exp\left[-\frac{i}{\hbar}\theta L_{n}\right]
\]
一种更深层次的理解：角动量算子是旋转变换的生成元，然后由角动量算子的不对易性和乘法逆的封闭性，可以自然得到\[
[L_x,L_y]=iL_z
\]，所以实际上只要满足这些东西的算子都可以被叫做角动量算子`;
const longPhysicsBlocks = parseLatexMarkdownDocument(longPhysicsNotes, "auto", 11);
const longPhysicsFormulas = longPhysicsBlocks.filter((block) => block.kind === "formula");
assert.equal(longPhysicsFormulas.length, 16);
assert.equal(
  longPhysicsFormulas.filter((formula) => formula.displayMode === "inline").length,
  4,
);
assert.ok(longPhysicsFormulas.every((formula) => !formula.numbered));

console.log("Document import parser smoke test passed");
