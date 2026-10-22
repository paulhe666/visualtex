import assert from "node:assert/strict";
import { additionalCommands } from "../src/autocomplete/additionalCommands.ts";
import { latexToSvg } from "../src/export/runtime.ts";
import { renderOfficeFormulaArtifacts } from "../src/office/shared/formulaRenderArtifacts.ts";

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

async function render(latex, formulaLetterFont) {
  return await latexToSvg(latex, {
    displayMode: false,
    fontSizePt: 14,
    paddingPx: 1,
    background: "transparent",
    forceExplicitBlack: true,
    formulaLetterFont,
  });
}

// Word images are drawn by the editor's layout engine: every glyph is a <use>
// of an outline whose id names its TeX font (Computer Modern for KaTeX, newtx
// for "Times New Roman"); Latin letters are never system-font <text>.
function glyphFonts(svg) {
  return new Set((svg.match(/id="vtx-([a-z0-9]+?)-\d+"/g) ?? []).map((m) => m.slice(8, m.lastIndexOf("-"))));
}

function assertEngineImage(result, label, font) {
  assert.ok(result.width > 0 && result.height > 0, `${label} has a size`);
  assert.ok(result.baseline > 0 && result.baseline <= result.height, `${label} baseline inside the image`);
  assert.ok(!/<text\b/.test(result.svg), `${label} draws no system-font text`);
  const fonts = [...glyphFonts(result.svg)];
  assert.ok(fonts.length > 0, `${label} has glyph outlines`);
  if (font === "times") {
    assert.ok(fonts.some((f) => f.startsWith("ntx")), `${label} uses the newtx (Times) fonts: ${fonts}`);
  } else {
    assert.ok(!fonts.some((f) => f.startsWith("ntx")), `${label} uses Computer Modern only: ${fonts}`);
  }
}

const inlineTextStyle = await renderOfficeFormulaArtifacts({
  lines: [{ id: "inline-style-test", latex: String.raw`\sum_{i=1}^{n}\frac{x_i}{y_i}` }],
  codeFormat: "raw",
  displayMode: "inline",
  inlineImageMathStyle: "text",
  host: "word",
  includeWordOmml: false,
  fontSizePt: 14,
});
const inlineDisplayStyle = await renderOfficeFormulaArtifacts({
  lines: [{ id: "inline-style-test", latex: String.raw`\sum_{i=1}^{n}\frac{x_i}{y_i}` }],
  codeFormat: "raw",
  displayMode: "inline",
  inlineImageMathStyle: "display",
  host: "word",
  includeWordOmml: false,
  fontSizePt: 14,
});
assert.notEqual(
  inlineDisplayStyle.svg.svg,
  inlineTextStyle.svg.svg,
  "Word inline image display style must produce different SVG layout from text style",
);
assert.ok(
  inlineDisplayStyle.svg.height > inlineTextStyle.svg.height,
  "Word inline image display style should expand large operators/fractions vertically",
);

const commandIds = new Set(additionalCommands.map((command) => command.id));
assert.ok(commandIds.has("arcsin"), "arcsin command candidate is missing");
assert.ok(commandIds.has("arccos"), "arccos command candidate is missing");
assert.ok(commandIds.has("arctan"), "arctan command candidate is missing");

const reportedReference = await render(REPORTED_TIMES_FORMULA, "katex");
const reportedTimes = await render(REPORTED_TIMES_FORMULA, "times");
assertEngineImage(reportedReference, "reported \\arccos (KaTeX)", "katex");
assertEngineImage(reportedTimes, "reported \\arccos (Times)", "times");

for (const [name, latex] of BASELINE_CASES) {
  const reference = await render(latex, "katex");
  assertEngineImage(reference, `${name}/katex`, "katex");
  const times = await render(latex, "times");
  assertEngineImage(times, `${name}/times`, "times");
  // fonts of earlier versions (Cambria, STIX, ...) are drawn as KaTeX
  for (const font of CUSTOM_LETTER_FONTS.filter((f) => f !== "times")) {
    const result = await render(latex, font);
    assert.equal(result.svg, reference.svg, `${name}/${font} is drawn as KaTeX`);
  }
}

for (const latex of MULTI_LETTER_OPERATORS) {
  assertEngineImage(await render(latex, "times"), `Times ${latex}`, "times");
}

const commandFailures = [];
for (const command of additionalCommands) {
  try {
    const result = await render(command.previewLatex, "times");
    if (/<text\b/.test(result.svg) && !/[\u3400-\u9fff]/.test(command.previewLatex)) {
      commandFailures.push({ id: command.id, previewLatex: command.previewLatex, error: "system-font text" });
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

console.log(
  `VisualTeX Word image font regression: PASS (${additionalCommands.length} Times command previews, ${BASELINE_CASES.length} baseline classes, ${MULTI_LETTER_OPERATORS.length} multi-letter operators)`,
);
