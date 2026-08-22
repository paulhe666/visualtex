import assert from "node:assert/strict";
import { additionalCommands } from "../src/autocomplete/additionalCommands.ts";
import { latexToSvg } from "../src/export/runtime.ts";

const REPORTED_TIMES_FORMULA = String.raw`\theta=\arccos\left(\frac{(A-B)\cdot(C-B)}{\left|A-B\right|\left|C-B\right|}\right)`;
const CUSTOM_LETTER_FONTS = ["times", "cambria", "stix", "palatino", "helvetica"];
const BASELINE_CASES = [
  ["reported-arccos", REPORTED_TIMES_FORMULA],
  ["simple-superscript", String.raw`x^2`],
  ["subscript", String.raw`L_z+v_{n+1}`],
  ["superscript", String.raw`L^2+x^{n+1}`],
  ["fraction", String.raw`\frac{a}{b}`],
  ["fraction-mixed", String.raw`\frac{a_i+b^2}{c_j-d}`],
  ["accents", String.raw`\bar{x}_i+\hat{y}^{2}+\vec{v}_{n+1}`],
  ["integral", String.raw`\int_0^x f(t)\,\mathrm{d}t`],
  ["sum", String.raw`\sum_i a_i`],
  ["sqrt", String.raw`\sqrt{x}`],
  ["delimiters", String.raw`\left(\frac{a}{b}\right)+\left|x_i\right|`],
];
const MULTI_LETTER_OPERATORS = [
  String.raw`\sin x`,
  String.raw`\cos x`,
  String.raw`\tan x`,
  String.raw`\cot x`,
  String.raw`\sec x`,
  String.raw`\csc x`,
  String.raw`\arcsin x`,
  String.raw`\arccos x`,
  String.raw`\arctan x`,
  String.raw`\sinh x`,
  String.raw`\cosh x`,
  String.raw`\tanh x`,
  String.raw`\exp x`,
  String.raw`\log x`,
  String.raw`\ln x`,
  String.raw`\lim_{x\to0}f(x)`,
  String.raw`\max_x f(x)`,
  String.raw`\min_x f(x)`,
  String.raw`\det A`,
  String.raw`\gcd(a,b)`,
  String.raw`\operatorname{rank}(A)`,
  String.raw`\operatorname{diag}(A)`,
  String.raw`\operatorname{tr}(A)`,
  String.raw`\operatorname{sgn}(x)`,
  String.raw`\operatorname{erf}(x)`,
];

function render(latex, formulaLetterFont) {
  return latexToSvg(latex, {
    displayMode: false,
    fontSizePt: 14,
    paddingPx: 1,
    background: "transparent",
    forceExplicitBlack: true,
    formulaLetterFont,
  });
}

function collapsedCustomTextGroups(svg) {
  const collapsed = [];
  for (const group of svg.matchAll(/<g\b[^>]*>([\s\S]*?)<\/g>/g)) {
    const texts = [
      ...group[1].matchAll(
        /<text\b[^>]*data-visualtex-output-letter-font=["'][^"']+["'][^>]*transform=["']([^"']+)["'][^>]*>([^<]*)<\/text>/g,
      ),
    ];
    if (texts.length <= 1) continue;
    const misplaced = texts
      .slice(1)
      .filter((match) => match[1] === "scale(1,-1)")
      .map((match) => match[2]);
    if (misplaced.length > 0) collapsed.push(misplaced.join(""));
  }
  return collapsed;
}

const commandIds = new Set(additionalCommands.map((command) => command.id));
assert.ok(commandIds.has("arcsin"), "arcsin command candidate is missing");
assert.ok(commandIds.has("arccos"), "arccos command candidate is missing");
assert.ok(commandIds.has("arctan"), "arctan command candidate is missing");

const reportedReference = render(REPORTED_TIMES_FORMULA, "katex");
const reportedTimes = render(REPORTED_TIMES_FORMULA, "times");
assert.equal(reportedTimes.width, reportedReference.width);
assert.equal(reportedTimes.height, reportedReference.height);
assert.equal(reportedTimes.baseline, reportedReference.baseline);
assert.deepEqual(collapsedCustomTextGroups(reportedTimes.svg), []);
for (const [codePoint, translate] of [
  ["72", "500"],
  ["63", "892"],
  ["6F", "1780"],
  ["73", "2280"],
]) {
  assert.match(
    reportedTimes.svg,
    new RegExp(
      `<text\\b[^>]*data-c=["']${codePoint}["'][^>]*transform=["']translate\\(${translate},0\\) scale\\(1,-1\\)["']`,
    ),
    `reported Times \\arccos lost glyph translation ${translate}`,
  );
}

for (const [name, latex] of BASELINE_CASES) {
  const reference = render(latex, "katex");
  assert.ok(reference.baseline > 0 && reference.baseline <= reference.height, `${name} reference baseline`);
  for (const font of CUSTOM_LETTER_FONTS) {
    const result = render(latex, font);
    assert.equal(result.width, reference.width, `${name}/${font} width geometry changed`);
    assert.equal(result.height, reference.height, `${name}/${font} height geometry changed`);
    assert.equal(result.baseline, reference.baseline, `${name}/${font} mathematical baseline changed`);
    assert.deepEqual(
      collapsedCustomTextGroups(result.svg),
      [],
      `${name}/${font} collapsed custom-font glyph positions`,
    );
  }
}

for (const latex of MULTI_LETTER_OPERATORS) {
  const result = render(latex, "times");
  assert.deepEqual(
    collapsedCustomTextGroups(result.svg),
    [],
    `Times multi-letter operator collapsed: ${latex}`,
  );
}

const commandFailures = [];
const commandCollapsed = [];
for (const command of additionalCommands) {
  try {
    const result = render(command.previewLatex, "times");
    const collapsed = collapsedCustomTextGroups(result.svg);
    if (collapsed.length > 0) {
      commandCollapsed.push({ id: command.id, previewLatex: command.previewLatex, collapsed });
    }
  } catch (error) {
    commandFailures.push({
      id: command.id,
      previewLatex: command.previewLatex,
      error: error instanceof Error ? error.message : String(error),
    });
  }
}
assert.deepEqual(commandFailures, [], `Times command render failures: ${JSON.stringify(commandFailures)}`);
assert.deepEqual(commandCollapsed, [], `Times command positioning failures: ${JSON.stringify(commandCollapsed)}`);

console.log(
  `VisualTeX Word image font regression: PASS (${additionalCommands.length} Times command previews, ${BASELINE_CASES.length} baseline classes, ${MULTI_LETTER_OPERATORS.length} multi-letter operators)`,
);
