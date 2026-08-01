import assert from "node:assert/strict";
import {
  formatLatex,
  latexCodeFormats,
  parseLatexSource,
  parseLatexSourceDraft,
} from "../src/clipboard/LatexCopyService.ts";

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
  assert.deepEqual(
    parseLatexSourceDraft(source, format.id),
    { valid: true, values: formulas },
    `${format.id} failed strict source parsing:\n${source}`,
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

const inlineTextFormula =
  "\\text{速度}v=\\frac{s}{t}+x^{\\text{中文}}+y_{\\text{下标}}";
const inlineTextSource = formatLatex(
  inlineTextFormula,
  "inline-text-double-dollar",
);
assert.equal(
  inlineTextSource,
  "速度$$v=\\frac{s}{t}+x^{\\text{中文}}+y_{\\text{下标}}$$",
);
assert.doesNotMatch(inlineTextSource, /^\$\$\\text\{/);
assert.match(inlineTextSource, /x\^\{\\text\{中文\}\}/);
assert.deepEqual(
  parseLatexSourceDraft(
    inlineTextSource,
    "inline-text-double-dollar",
  ),
  { valid: true, values: [inlineTextFormula] },
);

for (const [source, format, expectedError, expectedValues] of [
  ["$$\\fra$$", "display-dollar", "unknown-command", ["\\fra"]],
  [
    "$$a+b+\\frac{a}{b}+\\mathbb{H}+\\p$$",
    "display-dollar",
    "unknown-command",
    ["a+b+\\frac{a}{b}+\\mathbb{H}+\\p"],
  ],
  [
    "$$\\frac{a}$$",
    "display-dollar",
    "incomplete-command-arguments",
    ["\\frac{a}"],
  ],
  ["$$x^$$", "display-dollar", "incomplete-script", ["x^"]],
  ["$$\\frac{a}{b}", "display-dollar", "incomplete-format-wrapper", []],
  [
    "文字$$\\frac{a}{b}",
    "inline-text-double-dollar",
    "incomplete-format-wrapper",
    [],
  ],
]) {
  const result = parseLatexSourceDraft(source, format);
  assert.equal(result.valid, false, JSON.stringify({ source, format, result }));
  assert.equal(result.error, expectedError, JSON.stringify({ source, format, result }));
  assert.deepEqual(
    result.values,
    expectedValues,
    JSON.stringify({ source, format, result }),
  );
}

console.log(`LaTeX format smoke test passed (${latexCodeFormats.length} formats)`);
