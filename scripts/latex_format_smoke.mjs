import assert from "node:assert/strict";
import {
  formatLatex,
  latexCodeFormats,
  parseLatexSource,
} from "../src/clipboard/LatexCopyService.ts";
import {
  normalizeChineseLatex,
  normalizeContextualUprightSymbols,
  normalizeMathLiveCanonicalUprightCommands,
  resolveVisualTexInlineShortcuts,
  visualTexAutoEscapeInlineShortcuts,
  visualTexUprightInlineShortcuts,
} from "../src/editor/normalizeChineseLatex.ts";

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
assert.match(alignSource, /a&=b\+c \\\\/);
assert.match(alignSource, /\\frac\{x\}\{y\}&=z \\\\/);
assert.match(
  alignSource,
  /\\begin\{matrix\}a&b\\\\c&d\\end\{matrix\}&=M/,
  "matrix alignment markers or row breaks were changed",
);

assert.deepEqual(
  parseLatexSource(
    "\\begin{align*}\na &= b \\\\[4pt]\nc &= d\n\\end{align*}",
    "align-star",
  ),
  ["a = b", "c = d"],
  "optional row spacing was not parsed correctly",
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
  String.raw`\mathrm{d} x+\mathrm{D} y+\mathrm{e}^{\mathrm{i} x}+\mathrm{j}`,
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
  String.raw`\mathrm{d} x+\mathrm{D} y+\mathrm{e}^{\mathrm{i} x}+\mathrm{j}`,
  "copied LaTeX must never expose MathLive-only upright commands",
);
for (const shortcut of Object.values(visualTexUprightInlineShortcuts)) {
  assert.doesNotMatch(
    shortcut.after,
    /(?:^|\+)letter(?:\+|$)|(?:^|\+)digit(?:\+|$)/,
    "VisualTeX upright shortcuts must not trigger inside identifiers",
  );
}
assert.equal(
  visualTexUprightInlineShortcuts.dtheta.value,
  String.raw`\mathrm{d}\theta`,
);
assert.equal(
  visualTexUprightInlineShortcuts.dx.value,
  String.raw`\mathrm{d}x`,
  "dx must keep x in the normal italic math alphabet",
);
assert.equal(
  visualTexUprightInlineShortcuts.dy.value,
  String.raw`\mathrm{d}y`,
);
assert.equal(
  visualTexUprightInlineShortcuts.dt.value,
  String.raw`\mathrm{d}t`,
);
assert.equal(
  visualTexUprightInlineShortcuts.dr,
  undefined,
  "VisualTeX must not introduce new two-character shortcuts such as dr",
);
assert.deepEqual(
  resolveVisualTexInlineShortcuts({ alpha: String.raw`\alpha` }, false),
  {},
  "disabling automatic conversion must remove every inline shortcut",
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
assert.equal(enabledShortcuts.nativeOnly, String.raw`\star`);
assert.equal(
  enabledShortcuts.xx,
  undefined,
  "xx must remain ordinary x input instead of becoming a multiplication sign",
);
assert.equal(
  enabledShortcuts.mathbb,
  undefined,
  "font-variant commands must require a backslash or toolbar action",
);
assert.equal(enabledShortcuts.mathcal, undefined);
assert.equal(enabledShortcuts.pp, "+");
assert.equal(enabledShortcuts.ss, "-");
assert.equal(enabledShortcuts.mm, String.raw`\times`);
assert.equal(enabledShortcuts.dd, String.raw`\div`);
assert.equal(enabledShortcuts.eq, "=");
assert.equal(enabledShortcuts[">="], String.raw`\ge`);
assert.equal(enabledShortcuts.geq, String.raw`\geq`);
assert.equal(enabledShortcuts.leq, String.raw`\leq`);
assert.equal(enabledShortcuts.neq, String.raw`\neq`);
assert.equal(enabledShortcuts.varphi, String.raw`\varphi`);
assert.equal(enabledShortcuts.hat, String.raw`\hat{#?}`);
assert.equal(visualTexAutoEscapeInlineShortcuts.mathbf, undefined);
assert.equal(visualTexAutoEscapeInlineShortcuts.mathbb, undefined);
assert.equal(
  visualTexAutoEscapeInlineShortcuts.dx.value,
  String.raw`\mathrm{d}x`,
);
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
