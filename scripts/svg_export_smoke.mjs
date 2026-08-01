import assert from "node:assert/strict";
import {
  latexToMathMl,
  latexToSvg,
  svgToBase64,
} from "../src/export/runtime.ts";
import { normalizeChineseLatex } from "../src/editor/normalizeChineseLatex.ts";
import { EXTENDED_INTEGRAL_SYMBOLS } from "../src/math/extendedIntegralCompatibility.ts";

const matrixRows = Array.from({ length: 10 }, (_, row) =>
  Array.from({ length: 10 }, (_, column) => `a_{${row + 1}${column + 1}}`).join("&"),
).join("\\\\");

const cases = [
  ["fraction", String.raw`\frac{a+b}{c+d}`],
  ["root", String.raw`\sqrt[n]{x^2+y^2}`],
  ["integral", String.raw`\int_{-\infty}^{\infty} e^{-x^2}\,\mathrm{d}x`],
  ["sum", String.raw`\sum_{i=1}^{n} i^2`],
  ["matrix", String.raw`\begin{pmatrix}${matrixRows}\end{pmatrix}`],
  ["chinese", String.raw`\text{测试}+\alpha`],
  ["multiline", "a=b+c\nd=e-f\ng=h"],
  ["long", Array.from({ length: 25 }, (_, index) => `x_{${index + 1}}`).join("+")],
  ["bm-single-token", String.raw`A\bm v=\lambda\bm v`],
  ["bm-group", String.raw`\nabla\cdot\bm{F}+\boldsymbol{\alpha}`],
  ["math-fonts", String.raw`\mathbf{x}+\mathrm{d}+\operatorname{rank}(A)+\mathbb{R}+\mathcal{L}+\mathfrak{g}`],
  ["accents", String.raw`\vec{x}+\hat{x}+\bar{x}+\dot{x}+\ddot{x}`],
  ["over-under", String.raw`\overset{a}{=}+\underset{b}{=}`],
  ["substack", String.raw`\sum_{\substack{i=1\\j=2}}^n a_{ij}`],
  ["cases", String.raw`f(x)=\begin{cases}x^2,&x>0\\0,&x\le 0\end{cases}`],
  ["matrix-family", String.raw`\begin{matrix}a&b\\c&d\end{matrix}+\begin{pmatrix}a&b\\c&d\end{pmatrix}+\begin{bmatrix}a&b\\c&d\end{bmatrix}+\begin{vmatrix}a&b\\c&d\end{vmatrix}+\begin{Vmatrix}a&b\\c&d\end{Vmatrix}`],
  ["scalable-delimiters", String.raw`\left(\frac{a}{b}\right)+\left\lVert\bm v\right\rVert`],
  ["vector-calculus", String.raw`\partial_x+\nabla f+\nabla\cdot\bm F+\nabla\times\bm F+\nabla^2 f`],
];

function assertNoUnknownMathCommand(mathMl, context) {
  assert.doesNotMatch(
    mathMl,
    /<mtext\b[^>]*mathcolor="red"[^>]*>\s*\\/i,
    `${context} contains a MathJax unknown-command error`,
  );
}

const extendedIntegralCases = Object.entries(EXTENDED_INTEGRAL_SYMBOLS).map(
  ([command, character]) => [
    command,
    character,
    character.codePointAt(0).toString(16).toUpperCase(),
  ],
);

for (const [name, latex] of cases) {
  const result = await latexToSvg(latex, {
    displayMode: true,
    fontSizePt: 14,
    paddingPx: 10,
    background: name === "root" ? "white" : "transparent",
  });
  assert.match(result.svg, /^<svg\b/);
  assert.match(result.svg, /\bviewBox=/);
  assert.ok(result.width > 0, `${name} width`);
  assert.ok(result.height > 0, `${name} height`);
  assert.ok((result.baseline ?? -1) >= 0, `${name} baseline`);
  assert.ok(!/<foreignObject\b/i.test(result.svg), `${name} foreignObject`);
  assert.ok(!/<link\b|@import\b/i.test(result.svg), `${name} external CSS`);
  assert.ok(
    !/\b(?:href|xlink:href)=["'](?!#|data:)[^"']+/i.test(result.svg),
    `${name} external href`,
  );
  assert.ok(!/url\(\s*["']?https?:/i.test(result.svg), `${name} remote CSS URL`);
  if (name !== "root") {
    assert.match(
      result.svg,
      /<rect\b[^>]*fill-opacity="0\.001"/,
      `${name} transparent PowerPoint hit target`,
    );
  }
  const mathMl = latexToMathMl(latex, true);
  assert.match(mathMl, /^<math\b/);
  assert.match(mathMl, /xmlns="http:\/\/www\.w3\.org\/1998\/Math\/MathML"/);
  assertNoUnknownMathCommand(mathMl, name);
  if (name === "fraction") assert.match(mathMl, /<mfrac>/);
  if (name === "root") assert.match(mathMl, /<mroot>/);
  if (name === "matrix" || name === "matrix-family" || name === "cases" || name === "substack") {
    assert.match(mathMl, /<mtable(?:\s|>)/);
  }
  if (name.startsWith("bm-")) {
    assert.match(mathMl, /mathvariant="bold-italic"/);
    assert.doesNotMatch(mathMl, /\\bm(?:<|\s)/);
  }
  assert.equal(result.base64, svgToBase64(result.svg));
  const decoded = new TextDecoder().decode(
    Uint8Array.from(atob(result.base64), (character) => character.charCodeAt(0)),
  );
  assert.equal(decoded, result.svg, `${name} UTF-8 base64 round trip`);
}

for (const [command, character, codePoint] of extendedIntegralCases) {
  const latex = `\\${command}_{\\Sigma} a\\,\\mathrm{d}S`;
  const mathMl = latexToMathMl(latex, true);
  const svgResult = latexToSvg(latex, {
    displayMode: true,
    fontSizePt: 14,
    paddingPx: 0,
    background: "transparent",
  });
  const svg = svgResult.svg;

  assert.match(mathMl, new RegExp(`&#x${codePoint};`, "i"), `${command} MathML symbol`);
  assert.match(mathMl, /<msub>/, `${command} keeps its lower limit`);
  assertNoUnknownMathCommand(mathMl, command);
  assert.doesNotMatch(mathMl, new RegExp(`\\\\${command}(?:<|$)`), `${command} is not literal text`);
  assert.doesNotMatch(svg, /mathcolor|fill="red"|#FF0000/i, `${command} SVG has no error glyph`);
  assert.doesNotMatch(svg, new RegExp(`\\\\${command}`), `${command} SVG has no literal command`);
  assert.match(
    svg,
    /data-visualtex-integral="[A-Za-z]+"/,
    `${command} OLE SVG uses a VisualTeX vector glyph`,
  );
  assert.doesNotMatch(
    svg,
    new RegExp(`<text[^>]*>${character}</text>`),
    `${command} OLE SVG must not fall back to a small system-font character`,
  );
  assert.ok(svgResult.height > 30, `${command} display operator keeps large-integral height`);

  const normalizedUnicode = normalizeChineseLatex(`${character}_{S}F`);
  assert.match(
    normalizedUnicode,
    /^\\[A-Za-z]+\s*_\{S\}F$/,
    `${command} Unicode serialization is restored to canonical LaTeX`,
  );
  const reopenedSvg = latexToSvg(normalizedUnicode, {
    displayMode: true,
    fontSizePt: 14,
    paddingPx: 0,
    background: "transparent",
  }).svg;
  assert.match(
    reopenedSvg,
    /data-visualtex-integral="[A-Za-z]+"/,
    `${command} remains resolved after Unicode save/reopen normalization`,
  );
}

assert.throws(
  () =>
    assertNoUnknownMathCommand(
      latexToMathMl(String.raw`\definitelyUnknownVisualTeXCommand`, true),
      "unknown-command guard",
    ),
  /did not resolve LaTeX command/,
);

assert.throws(
  () =>
    latexToSvg("", {
      displayMode: true,
      fontSizePt: 12,
      paddingPx: 8,
      background: "transparent",
    }),
  /Cannot export an empty formula/,
);

console.log(
  `SVG export smoke test passed (${cases.length} formula classes, ${extendedIntegralCases.length} extended integral operators)`,
);
