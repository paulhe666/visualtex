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
assert.equal(compatibilityRequiredArgumentCounts.get("symbfit"), 1);
assert.equal(compatibilityRequiredArgumentCounts.get("abs"), 1);
assert.equal(compatibilityRequiredArgumentCounts.get("dv"), 2);
assert.equal(compatibilityRequiredArgumentCounts.get("mel"), 3);

for (const source of [
  String.raw`\symbfit{J}`,
  String.raw`\bm{\alpha J}`,
  String.raw`\abs{x}`,
  String.raw`\dv{f}{x}`,
  String.raw`\braket{\phi}{\psi}`,
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

console.log("VisualTeX math variant/package shorthand compatibility regression passed");
