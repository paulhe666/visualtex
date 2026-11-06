import assert from "node:assert/strict";
import {
  formatLatex,
  latexCodeFormats,
  parseLatexSource,
} from "../src/clipboard/LatexCopyService.ts";
import { VISUALTEX_ALIGNMENT_MARKER_LATEX } from "../src/editor/alignmentMarkers.ts";
import {
  normalizeChineseLatex,
  normalizeContextualUprightSymbols,
  normalizeMathLiveCanonicalUprightCommands,
  resolveVisualTexInlineShortcuts,
  visualTexAutoEscapeInlineShortcuts,
} from "../src/editor/normalizeChineseLatex.ts";
import { applyInputJudgment } from "../src/editor/inputJudgment.ts";

const formulas = [
  "a=b+c",
  "\\frac{x}{y}=z",
  "\\begin{matrix}a&b\\\\c&d\\end{matrix}=M",
];
const latex = formulas.join("\n");

for (const format of latexCodeFormats) {
  const source = formatLatex(latex, format.id);
  assert.ok(source.length > 0, `${format.id} produced empty source`);

  const parsed = parseLatexSource(source, format.id);
  assert.deepEqual(
    parsed,
    formulas,
    `${format.id} failed format/parse round trip:\n${source}`,
  );
}

const alignSource = formatLatex(latex, "align-star");
assert.match(alignSource, /\\begin\{align\*\}/);
assert.doesNotMatch(
  alignSource,
  /a&=b|\\frac\{x\}\{y\}&=z/,
  "align formatting must not invent alignment points",
);
assert.match(
  alignSource,
  /\\begin\{matrix\}a&b\\\\c&d\\end\{matrix\}=M/,
  "nested matrix separators or row breaks were changed",
);

const explicitAlignFormulas = [
  `a${VISUALTEX_ALIGNMENT_MARKER_LATEX}=b+c`,
  `\\frac{x}{y}${VISUALTEX_ALIGNMENT_MARKER_LATEX}=z`,
  `\\begin{matrix}a&b\\\\c&d\\end{matrix}${VISUALTEX_ALIGNMENT_MARKER_LATEX}=M`,
];
const explicitAlignSource = formatLatex(
  explicitAlignFormulas.join("\n"),
  "align-star",
);
assert.match(explicitAlignSource, /a&=b\+c \\\\/);
assert.match(explicitAlignSource, /\\frac\{x\}\{y\}&=z \\\\/);
assert.match(
  explicitAlignSource,
  /\\begin\{matrix\}a&b\\\\c&d\\end\{matrix\}&=M/,
  "explicit alignment point after a nested matrix was not preserved",
);
assert.deepEqual(
  parseLatexSource(explicitAlignSource, "align-star"),
  explicitAlignFormulas,
  "explicit alignment points failed format/parse round trip",
);
assert.doesNotMatch(
  formatLatex(explicitAlignFormulas.join("\n"), "raw"),
  /visualtex-align-marker|\\class/,
  "VisualTeX alignment markers must not leak into non-align source formats",
);

assert.deepEqual(
  parseLatexSource(
    "\\begin{align*}\na &= b \\\\[4pt]\nc &= d\n\\end{align*}",
    "align-star",
  ),
  [
    `a ${VISUALTEX_ALIGNMENT_MARKER_LATEX}= b`,
    `c ${VISUALTEX_ALIGNMENT_MARKER_LATEX}= d`,
  ],
  "optional row spacing or explicit alignment points were not parsed correctly",
);

assert.match(formatLatex("a=b", "equation"), /\\begin\{equation\}/);
assert.doesNotMatch(formatLatex("a=b", "equation"), /equation\*/);
assert.match(formatLatex("a=b", "equation-star"), /\\begin\{equation\*\}/);
assert.match(
  formatLatex("a=b\nc=d", "equation-split"),
  /\\begin\{equation\}[\s\S]*\\begin\{split\}[\s\S]*\\end\{split\}[\s\S]*\\end\{equation\}/,
);
assert.match(
  formatLatex("a=b\nc=d", "aligned"),
  /^\\\[[\s\S]*\\begin\{aligned\}[\s\S]*\\end\{aligned\}[\s\S]*\\\]$/,
);

const canonicalUpright = String.raw`\differentialD x+\capitalDifferentialD y+\exponentialE^{\imaginaryI x}+\imaginaryJ`;
assert.equal(
  normalizeMathLiveCanonicalUprightCommands(canonicalUpright),
  String.raw`\mathrm{d}x+\mathrm{D}y+\mathrm{e}^{\mathrm{i}x}+\mathrm{j}`,
  "MathLive upright commands must be converted to portable LaTeX",
);
assert.equal(
  normalizeMathLiveCanonicalUprightCommands(
    String.raw`d+e+i+j+distance+limit+imaginaryIndex+\mathrm{d}`,
  ),
  String.raw`d+e+i+j+distance+limit+imaginaryIndex+\mathrm{d}`,
  "ordinary variables and identifiers must not be over-normalized",
);
assert.equal(
  normalizeMathLiveCanonicalUprightCommands(String.raw`\mathrm{dx}`),
  String.raw`\mathrm{d}x`,
  "a grouped differential variable must be split so only d is upright",
);
assert.equal(
  formatLatex(canonicalUpright, "raw"),
  String.raw`\mathrm{d}x+\mathrm{D}y+\mathrm{e}^{\mathrm{i}x}+\mathrm{j}`,
  "copied LaTeX must never expose MathLive-only upright commands",
);
// Typed words are decided by the one input judgment once finished
// (src/editor/inputJudgment.ts); MathLive converts only symbol sequences as
// they are typed.
const typed = (latex, escapes = true) =>
  applyInputJudgment(latex, { upright: true, escapes, typed: true });
assert.equal(typed("dtheta"), String.raw`\mathrm{d}\theta`);
assert.equal(
  typed("dx"),
  String.raw`\mathrm{d}x`,
  "dx must keep x in the normal italic math alphabet",
);
assert.equal(typed("dy"), String.raw`\mathrm{d}y`);
assert.equal(typed("dt"), String.raw`\mathrm{d}t`);
for (const word of ["spin", "point", "hnu"]) {
  assert.equal(typed(word), word, "a name inside other letters must not convert");
}
assert.deepEqual(
  resolveVisualTexInlineShortcuts({ alpha: String.raw`\alpha` }, false),
  {},
  "with quick escapes off nothing converts as it is typed",
);
assert.equal(typed("alpha", false), "alpha");
assert.equal(
  typed("dx", false),
  String.raw`\mathrm{d}x`,
  "upright differential detection must stay active when quick escapes are off",
);
const enabledShortcuts = resolveVisualTexInlineShortcuts(
  {
    nativeOnly: String.raw`\star`,
    xx: String.raw`\times`,
    mathbb: String.raw`\mathbb{#?}`,
    mathcal: String.raw`\mathcal{#?}`,
  },
  true,
);
assert.equal(enabledShortcuts[">="], String.raw`\ge`);
for (const word of ["nativeOnly", "xx", "mathbb", "mathcal", "alpha"]) {
  assert.equal(enabledShortcuts[word], undefined, word + " must not convert as it is typed");
}
assert.equal(
  typed("xx"),
  "xx",
  "xx must remain ordinary x input instead of becoming a multiplication sign",
);
assert.equal(
  typed("mathbb"),
  "mathbb",
  "font-variant commands must require a backslash or toolbar action",
);
assert.equal(typed("mathcal"), "mathcal");
for (const word of ["pp", "ss", "dd"]) {
  assert.equal(typed(word), word, "products of variables are not shortcuts");
}
assert.equal(typed("eq"), "=");
assert.equal(typed("geq"), String.raw`\geq`);
assert.equal(typed("leq"), String.raw`\leq`);
assert.equal(typed("neq"), String.raw`\neq`);
assert.equal(typed("varphi"), String.raw`\varphi`);
assert.equal(typed("hat"), String.raw`\hat{\placeholder{}}`);
assert.equal(visualTexAutoEscapeInlineShortcuts.mathbf, undefined);
assert.equal(visualTexAutoEscapeInlineShortcuts.mathbb, undefined);
const contextualDifferentialCases = [
  [
    String.raw`dr/d\theta`,
    String.raw`\mathrm{d}r/\mathrm{d}\theta`,
    "slash derivative with Latin and Greek variables",
  ],
  [
    String.raw`\frac{dr}{d\theta}`,
    String.raw`\frac{\mathrm{d}r}{\mathrm{d}\theta}`,
    "fraction derivative with Latin and Greek variables",
  ],
  [
    String.raw`\frac{d}{dr}`,
    String.raw`\frac{\mathrm{d}}{\mathrm{d}r}`,
    "standalone derivative operator in the numerator",
  ],
  [
    String.raw`\frac{d^2y}{dx^2}`,
    String.raw`\frac{\mathrm{d}^2y}{\mathrm{d}x^2}`,
    "second derivative",
  ],
  [
    String.raw`\frac{d^2y}{\mathrm{d}x^2}`,
    String.raw`\frac{\mathrm{d}^2y}{\mathrm{d}x^2}`,
    "incremental second derivative with an already-upright denominator",
  ],
  [
    String.raw`\frac{d\mathbf{r}}{dt}`,
    String.raw`\frac{\mathrm{d}\mathbf{r}}{\mathrm{d}t}`,
    "styled vector differential",
  ],
  [
    String.raw`\int_0^1 f(x) dx`,
    String.raw`\int_0^1 f(x) \mathrm{d}x`,
    "single integral measure",
  ],
  [
    String.raw`\iint_S f dA`,
    String.raw`\iint_S f \mathrm{d}A`,
    "surface integral measure",
  ],
  [
    String.raw`\int f d\theta`,
    String.raw`\int f \mathrm{d}\theta`,
    "Greek integral measure",
  ],
  [
    String.raw`\int \mathbf{F}\cdot d\mathbf{r}`,
    String.raw`\int \mathbf{F}\cdot \mathrm{d}\mathbf{r}`,
    "vector line element",
  ],
];
for (const [source, expected, description] of contextualDifferentialCases) {
  assert.equal(
    normalizeContextualUprightSymbols(source),
    expected,
    description,
  );
  assert.equal(normalizeChineseLatex(source), expected, description);
}
assert.equal(
  normalizeChineseLatex(
    String.raw`distance+dimension+driver+dr+d\theta+dV+\frac{\partial f}{\partial x}`,
  ),
  String.raw`distance+dimension+driver+dr+d\theta+dV+\frac{\partial f}{\partial x}`,
  "ordinary identifiers, standalone d variables, and partial derivatives must not be over-normalized",
);
console.log(`LaTeX format smoke test passed (${latexCodeFormats.length} formats)`);
