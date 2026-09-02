import assert from "node:assert/strict";
import {
  compatibilityCommandNames,
  compatibilityRawPlaceholderTemplates,
  compatibilityRequiredArgumentCounts,
  compatibilityWrapperCanonicalTargets,
  compatibilityWrapperPreviews,
} from "../src/autocomplete/compatibilityCommands.ts";
import { commandRegistry } from "../src/autocomplete/commandRegistry.ts";
import {
  formatLatex,
  parseLatexSourceDraft,
} from "../src/clipboard/LatexCopyService.ts";
import { convertVisualTexLatexToMarkup } from "../src/editor/mathLiveIntegralCompatibility.ts";
import { latexToMathMl, latexToSvg } from "../src/export/runtime.ts";

const wrapperCommands = [
  "\\bm",
  "\\mathbfit",
  "\\symbfit",
  "\\simbfit",
  "\\symbf",
  "\\symbfup",
  "\\symbb",
  "\\symcal",
  "\\symfrak",
  "\\symtt",
  "\\boldmath",
  "\\bold",
  "\\pmb",
  "\\bf",
  "\\Bbb",
  "\\frak",
  "\\abs",
  "\\norm",
  "\\dd",
  "\\bra",
  "\\ket",
  "\\expval",
  "\\vb",
  "\\va",
  "\\vu",
  "\\pmod",
  "\\pod",
];

for (const command of wrapperCommands) {
  assert.ok(
    commandRegistry.some((entry) => entry.command === command),
    `${command} is missing from the command registry`,
  );
  assert.ok(
    compatibilityWrapperPreviews.has(command) ||
      command === "\\bm" ||
      command === "\\mathbfit",
    `${command} is missing wrapper-input registration`,
  );
}

for (const command of ["\\bmod", "\\mod"]) {
  assert.ok(
    commandRegistry.some((entry) => entry.command === command),
    `${command} is missing from the command registry`,
  );
}

for (const command of [
  "\\comm",
  "\\acomm",
  "\\pb",
  "\\dv",
  "\\pdv",
  "\\braket",
  "\\ketbra",
  "\\mel",
]) {
  assert.ok(
    commandRegistry.some((entry) => entry.command === command),
    `${command} is missing from the command registry`,
  );
  assert.ok(
    compatibilityRawPlaceholderTemplates.has(command),
    `${command} is missing structural placeholder registration`,
  );
}

assert.equal(
  compatibilityWrapperCanonicalTargets.get("\\boldmath"),
  "\\mathbfit",
);
assert.equal(
  compatibilityWrapperCanonicalTargets.get("\\simbfit"),
  "\\symbfit",
);
assert.equal(
  compatibilityWrapperCanonicalTargets.get("\\Bbb"),
  "\\mathbb",
);
assert.equal(
  compatibilityWrapperCanonicalTargets.get("\\frak"),
  "\\mathfrak",
);
assert.ok(compatibilityCommandNames.has("symbfit"));
assert.ok(compatibilityCommandNames.has("boldmath"));
assert.ok(compatibilityCommandNames.has("dv"));
for (const command of ["bmod", "mod", "pmod", "pod"]) {
  assert.ok(
    compatibilityCommandNames.has(command),
    `\\${command} is missing source-validation compatibility`,
  );
}
assert.equal(compatibilityRequiredArgumentCounts.get("symbfit"), 1);
assert.equal(compatibilityRequiredArgumentCounts.get("abs"), 1);
assert.equal(compatibilityRequiredArgumentCounts.get("dv"), 2);
assert.equal(compatibilityRequiredArgumentCounts.get("mel"), 3);
assert.equal(compatibilityRequiredArgumentCounts.get("pmod"), 1);
assert.equal(compatibilityRequiredArgumentCounts.get("pod"), 1);
assert.equal(compatibilityRequiredArgumentCounts.has("bmod"), false);
assert.equal(compatibilityRequiredArgumentCounts.has("mod"), false);

for (const source of [
  String.raw`\symbfit{J}`,
  String.raw`\bm{\alpha J}`,
  String.raw`\abs{x}`,
  String.raw`\dv{f}{x}`,
  String.raw`\braket{\phi}{\psi}`,
  String.raw`(ab)\bmod\beta`,
  String.raw`a\equiv b\mod n`,
  String.raw`a\equiv b\pmod n`,
  String.raw`a\equiv b\pod n`,
]) {
  const parsed = parseLatexSourceDraft(source, "raw");
  assert.equal(parsed.valid, true, `source editor rejected ${source}: ${parsed.error}`);
}

for (const source of [
  String.raw`\bm`,
  String.raw`\mathbfit`,
  String.raw`\symbfit`,
  String.raw`\abs`,
  String.raw`\dv`,
  String.raw`\dv{f}`,
  String.raw`\mel{a}{b}`,
  String.raw`a\pmod`,
  String.raw`a\pod`,
]) {
  const parsed = parseLatexSourceDraft(source, "raw");
  assert.equal(parsed.valid, false, `source editor accepted incomplete ${source}`);
  assert.equal(
    parsed.error,
    "incomplete-command-arguments",
    `source editor returned the wrong error for ${source}`,
  );
}

for (const source of [
  String.raw`\symbfit{J}`,
  String.raw`\bm{J}`,
  String.raw`\abs{x}+\norm{x}`,
  String.raw`(ab)\bmod\beta`,
  String.raw`a\equiv b\mod n`,
  String.raw`a\equiv b\pmod n`,
  String.raw`a\equiv b\pod n`,
]) {
  const markup = convertVisualTexLatexToMarkup(source, { defaultMode: "math" });
  assert.ok(markup.length > 0, `static MathLive preview is empty for ${source}`);
  assert.doesNotMatch(markup, /ML__error/, `static MathLive preview errored for ${source}`);
}

const boldItalicMarkup = convertVisualTexLatexToMarkup(
  String.raw`\symbfit{J}`,
  { defaultMode: "math" },
);
assert.match(
  boldItalicMarkup,
  /ML__mathbfit|ML__cmr[^\"]*ML__bold[^\"]*ML__it|ML__cmr[^\"]*ML__it[^\"]*ML__bold/,
  "symbfit static preview must resolve through a MathLive bold-italic variant",
);

const exportSource = String.raw`\symbfit{J}+\bm{\alpha}+\abs{x}+\dv{f}{x}`;
assert.equal(
  formatLatex(exportSource, "raw"),
  exportSource,
  "raw source copy must preserve compatibility command spellings",
);
assert.equal(
  formatLatex(exportSource, "display-dollar"),
  `$$\n${exportSource}\n$$`,
  "formatted source copy must preserve compatibility command spellings",
);
const svg = latexToSvg(exportSource, { displayMode: false }).svg;
assert.ok(svg.length > 0);
assert.doesNotMatch(svg, /data-mml-node=["']merror["']|mathcolor=["']?red/i);
const mathMl = latexToMathMl(exportSource, false);
assert.ok(mathMl.length > 0);
assert.doesNotMatch(mathMl, /<merror\b|mathcolor=["']?red/i);

const issue15Source = String.raw`\begin{aligned}
\langle p_1,p_0\rangle &\leftarrow \operatorname{umul}(a,b)=ab
&&\text{Double word product}\\
p_0 &\leftarrow \operatorname{umullo}(a,b)=(ab)\bmod\beta
&&\text{Low word}\\
p_1 &\leftarrow \operatorname{umulhi}(a,b)=\left\lfloor\frac{ab}{\beta}\right\rfloor
&&\text{High word.}
\end{aligned}`;
const issue15Draft = parseLatexSourceDraft(issue15Source, "raw");
assert.equal(
  issue15Draft.valid,
  true,
  `Issue #15 source was rejected: ${issue15Draft.error}`,
);
assert.equal(
  formatLatex(issue15Source, "raw"),
  issue15Source,
  "Issue #15 source copy must preserve \\bmod and aligned column markers",
);
const issue15Markup = convertVisualTexLatexToMarkup(issue15Source, {
  defaultMode: "math",
});
assert.ok(issue15Markup.length > 0, "Issue #15 MathLive preview is empty");
assert.doesNotMatch(
  issue15Markup,
  /ML__error/,
  "Issue #15 MathLive preview contains an unknown-command error",
);
for (const formulaLetterFont of [
  "katex",
  "times",
  "cambria",
  "stix",
  "palatino",
  "helvetica",
] as const) {
  const issue15Svg = latexToSvg(issue15Source, {
    displayMode: true,
    formulaLetterFont,
  });
  assert.ok(issue15Svg.width > 400, `${formulaLetterFont} Issue #15 SVG is too narrow`);
  assert.ok(issue15Svg.height > 70, `${formulaLetterFont} Issue #15 SVG is too short`);
  assert.doesNotMatch(
    issue15Svg.svg,
    /data-mml-node=["']merror["']|mathcolor=["']?red/i,
    `${formulaLetterFont} Issue #15 SVG contains a MathJax error`,
  );
}
const issue15MathMl = latexToMathMl(issue15Source, true);
assert.equal((issue15MathMl.match(/<mtr\b/g) ?? []).length, 3);
assert.equal((issue15MathMl.match(/<mtd\b/g) ?? []).length, 12);
assert.match(issue15MathMl, />mod</, "Issue #15 MathML lost the modulo operator");
assert.doesNotMatch(issue15MathMl, /<merror\b|mathcolor=["']?red/i);

console.log("VisualTeX math variant/package shorthand compatibility regression passed");
