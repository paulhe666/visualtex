import assert from "node:assert/strict";
import {
  latexToMathMl,
  latexToSvg,
  svgToBase64,
} from "../src/export/runtime.ts";
import { normalizeChineseLatex } from "../src/editor/normalizeChineseLatex.ts";
import { EXTENDED_INTEGRAL_SYMBOLS } from "../src/math/extendedIntegralCompatibility.ts";
import { isIncompleteLatexDraft } from "../src/math/latexCompatibility.ts";

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
  ["tagged-equation", String.raw`L^\dagger=p_2(x)\frac{d^2}{dx^2}+\bigl[2p_2(x)-p_1(x)\bigr]\frac{d}{dx}\tag{9.27}`],
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
  ["physics-package", String.raw`\qty(\frac{a}{b})+\dv{f}{x}+\pdv{g}{y}+\abs{x}+\norm{\bm v}`],
  ["siunitx-package", String.raw`\SI{3}{\meter\per\second}+\si{\kilogram}+\unit{\joule}+\qty{5}{\tesla}`],
  ["bbm-package", String.raw`\mathbbm{1}_{A}`],
  ["physics-derivative-orders", String.raw`\dv{x}+\dv[2]{f}{x}+\pdv[3]{g}{y}+\fdv{S}{\phi}`],
  ["physics-vectors-operators", String.raw`\vb{v}+\va{a}+\vu{n}+\pb{f}{g}+\order{x^2}+\Tr A+\rank A`],
  ["physics-matrix-quantities", String.raw`\mqty{a&b\\c&d}+\pmqty{1&0\\0&1}+\vmqty{x&y\\z&w}`],
  ["siunitx-options-ranges", String.raw`\SI[round-mode=places]{3.14}{\kilo\meter\per\second}+\qty[round-mode=figures]{5}{\tesla}+\qtyrange{1}{10}{\milli\second}+\ang{30}`],
];

function assertNoUnknownMathCommand(mathMl, context) {
  assert.doesNotMatch(
    mathMl,
    /<mtext\b[^>]*mathcolor="red"[^>]*>\s*\\/i,
    `${context} contains a MathJax unknown-command error`,
  );
}

const extendedIntegralCases = Object.entries(EXTENDED_INTEGRAL_SYMBOLS).map(
  ([command, replacement]) => [command, replacement],
);

function assertExtendedIntegralMathMl(mathMl, command, replacement) {
  if (replacement.startsWith("\\")) {
    assert.ok(
      (mathMl.match(/&#x222B;/gi) ?? []).length >= 2,
      `${command} composite MathML keeps both integral operators`,
    );
    assert.match(mathMl, /&#x22EF;/i, `${command} composite MathML keeps its dots`);
    return;
  }

  const expectedCounts = new Map();
  for (const character of Array.from(replacement)) {
    const codePoint = character.codePointAt(0).toString(16).toUpperCase();
    expectedCounts.set(codePoint, (expectedCounts.get(codePoint) ?? 0) + 1);
  }
  for (const [codePoint, count] of expectedCounts) {
    const actual = (mathMl.match(new RegExp(`&#x${codePoint};`, "gi")) ?? []).length;
    assert.ok(actual >= count, `${command} MathML symbol U+${codePoint}`);
  }
}

for (const [name, latex] of cases) {
  const result = await latexToSvg(latex, {
    displayMode: true,
    fontSizePt: 14,
    paddingPx: 10,
    background: name === "root" ? "white" : "transparent",
  });
  assert.match(result.svg, /^<svg\b/);
  const rootOpening = result.svg.match(/^<svg\b[^>]*>/)?.[0] ?? "";
  assert.match(rootOpening, /\bviewBox=/, `${name} root viewBox`);
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
  if (name === "tagged-equation") {
    assert.ok(result.width > 250, "tagged equation keeps its full intrinsic width");
    const nestedViewports = [...result.svg.matchAll(
      /<svg\b[^>]*\bdata-(?:table|labels)=["'][^"']+["'][^>]*>/g,
    )].map((match) => match[0]);
    assert.equal(nestedViewports.length, 2, "tagged equation table and label viewports");
    for (const viewport of nestedViewports) {
      assert.match(viewport, /\bviewBox=["'][^"']+["']/);
      assert.match(viewport, /\bwidth=["'][-+\d.eE]+["']/);
      assert.match(viewport, /\bheight=["'][-+\d.eE]+["']/);
      assert.match(viewport, /\boverflow=["']visible["']/);
    }
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

for (const [command, character] of extendedIntegralCases) {
  const latex = `\\${command}_{\\Sigma} a\\,\\mathrm{d}S`;
  const mathMl = latexToMathMl(latex, true);
  const svgResult = latexToSvg(latex, {
    displayMode: true,
    fontSizePt: 14,
    paddingPx: 0,
    background: "transparent",
  });
  const svg = svgResult.svg;

  assertExtendedIntegralMathMl(mathMl, command, character);
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

  if (!character.startsWith("\\")) {
    const normalizedUnicode = normalizeChineseLatex(`${character}_{S}F`);
    assert.match(
      normalizedUnicode,
      /^(?:\\[A-Za-z]+\s*)+_\{S\}F$/,
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
}

assert.equal(
  isIncompleteLatexDraft(String.raw`x+\placeholder{}`),
  true,
  "structural placeholder is an incomplete editor draft",
);
assert.equal(
  isIncompleteLatexDraft(
    String.raw`x+\alp`,
    new Error("MathJax did not resolve LaTeX command \\alp."),
  ),
  true,
  "a trailing partial command is an incomplete editor draft",
);
assert.equal(
  isIncompleteLatexDraft(String.raw`\frac{a}{`),
  true,
  "an unclosed group is an incomplete editor draft",
);
assert.equal(
  isIncompleteLatexDraft(String.raw`\begin{matrix}a&b`),
  true,
  "an unclosed environment is an incomplete editor draft",
);
assert.equal(
  isIncompleteLatexDraft(String.raw`x+\alpha`),
  false,
  "a complete command is not an incomplete draft",
);
assert.equal(
  isIncompleteLatexDraft(
    String.raw`x+\definitelyUnknownVisualTeXCommand+y`,
    new Error(
      "MathJax did not resolve LaTeX command \\definitelyUnknownVisualTeXCommand.",
    ),
  ),
  false,
  "a complete unknown command in the formula remains a real error",
);
assert.equal(
  isIncompleteLatexDraft(
    String.raw`x+\definitelyUnknownVisualTeXCommand+\alpha`,
    new Error(
      "MathJax did not resolve LaTeX command \\definitelyUnknownVisualTeXCommand.",
    ),
  ),
  false,
  "a valid trailing command must not hide an earlier unknown command",
);

assert.throws(
  () => latexToSvg(String.raw`x+\placeholder{}`),
  /empty VisualTeX placeholders/,
);
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
