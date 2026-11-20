import assert from "node:assert/strict";

import {
  findWindowsWordLatexRedrawSpans,
  type WordLatexRedrawSpan,
} from "../src/office/redraw/wordLatexRedrawParser.ts";

function assertExactSpan(source: string, span: WordLatexRedrawSpan) {
  assert.equal(source.slice(span.start, span.end), span.sourceText);
  assert.ok(span.end > span.start);
}

const mixedSource = [
  "前😀缀 $x+1$ 中间 \\$5 不是公式。",
  "显示：\\[ y^2 \\]，行内：\\(z\\)。",
  "被转义的分隔符：\\\\(not-math\\) 与 \\\\[still-not-math\\]。",
].join("\n");
const mixed = findWindowsWordLatexRedrawSpans(mixedSource);
assert.deepEqual(
  mixed.map((span) => [span.sourceText, span.latex, span.displayMode]),
  [
    ["$x+1$", "x+1", "inline"],
    ["\\[ y^2 \\]", "y^2", "block"],
    ["\\(z\\)", "z", "inline"],
  ],
);
for (const span of mixed) assertExactSpan(mixedSource, span);
assert.equal(mixed[0].start, "前😀缀 ".length);

const wordStyledQuadratic = String.raw`$𝑥=\\𝑓𝑟𝑎𝑐{−𝑏\\𝑝𝑚 \\𝑠𝑞𝑟𝑡{𝑏^{2}−4𝑎𝑐}}{2𝑎}$`;
const wordStyledQuadraticSpan =
  findWindowsWordLatexRedrawSpans(wordStyledQuadratic)[0];
assertExactSpan(wordStyledQuadratic, wordStyledQuadraticSpan);
assert.equal(
  wordStyledQuadraticSpan.sourceText,
  wordStyledQuadratic,
  "Word range verification must retain the exact styled source text",
);
assert.equal(
  wordStyledQuadraticSpan.latex,
  String.raw`x=\frac{-b\pm \sqrt{b^{2}-4ac}}{2a}`,
  "Word mathematical Unicode and duplicated command slashes must normalize back to standard LaTeX",
);

const displayWhitespace = "before $$\n  x +\n y  \n$$ after";
assert.equal(
  findWindowsWordLatexRedrawSpans(displayWhitespace)[0].latex,
  "x +\n y",
  "The shared scanner preserves newlines until TeX comments have been processed",
);

const environmentSource = [
  String.raw`\begin{equation}E=mc^2\end{equation}`,
  String.raw`\begin{align*}a&=b\\c&=d\end{align*}`,
  String.raw`\begin{gather}x\\y\end{gather}`,
  String.raw`\begin{multline*}p\\q\end{multline*}`,
  String.raw`\begin{displaymath}r+s\end{displaymath}`,
].join(" ");
assert.deepEqual(
  findWindowsWordLatexRedrawSpans(environmentSource).map((span) => [
    span.sourceText,
    span.latex,
    span.displayMode,
  ]),
  [
    [String.raw`\begin{equation}E=mc^2\end{equation}`, String.raw`\begin{equation}E=mc^2\end{equation}`, "block"],
    [
      String.raw`\begin{align*}a&=b\\c&=d\end{align*}`,
      String.raw`\begin{align*}a&=b\\c&=d\end{align*}`,
      "block",
    ],
    [
      String.raw`\begin{gather}x\\y\end{gather}`,
      String.raw`\begin{gather}x\\y\end{gather}`,
      "block",
    ],
    [
      String.raw`\begin{multline*}p\\q\end{multline*}`,
      String.raw`\begin{multline*}p\\q\end{multline*}`,
      "block",
    ],
    [String.raw`\begin{displaymath}r+s\end{displaymath}`, String.raw`\begin{displaymath}r+s\end{displaymath}`, "block"],
  ],
);

assert.deepEqual(
  findWindowsWordLatexRedrawSpans(
    String.raw`\begin{alignat}{2}x&=1&y&=2\end{alignat} \begin{math}z\end{math}`,
  ).map((span) => span.displayMode),
  ["block", "inline"],
  "Every import-supported math environment is also available to redraw",
);

assert.equal(
  findWindowsWordLatexRedrawSpans("unclosed $x+1").length,
  0,
);
assert.equal(
  findWindowsWordLatexRedrawSpans(String.raw`\begin{align}x&=1`).length,
  0,
);
assert.equal(findWindowsWordLatexRedrawSpans("$$$$").length, 0);

const replacementSource = "甲😀 $a$ 乙 $$b$$ 丙 \\(c\\) 丁";
const replacementSpans = findWindowsWordLatexRedrawSpans(replacementSource);
let replaced = replacementSource;
for (let index = replacementSpans.length - 1; index >= 0; index -= 1) {
  const span = replacementSpans[index];
  replaced =
    replaced.slice(0, span.start) +
    `<FORMULA:${span.latex}>` +
    replaced.slice(span.end);
}
assert.equal(
  replaced,
  "甲😀 <FORMULA:a> 乙 <FORMULA:b> 丙 <FORMULA:c> 丁",
);

const exactlyOneThousand = Array.from(
  { length: 1000 },
  (_value, index) => `$${index}$`,
).join(" ");
assert.equal(findWindowsWordLatexRedrawSpans(exactlyOneThousand).length, 1000);
assert.throws(
  () => findWindowsWordLatexRedrawSpans(`${exactlyOneThousand} $overflow$`),
  /at most 1000/,
);
assert.throws(
  () => findWindowsWordLatexRedrawSpans("x".repeat(5 * 1024 * 1024 + 1)),
  /5 MB/,
);

console.log("VisualTeX shared Word LaTeX redraw parser regression: PASS");
