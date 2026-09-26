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
import { normalizePackageLatexCommands } from "../src/math/latexCompatibility.ts";
import { VISUALTEX_PACKAGE_COMPATIBILITY_MACROS } from "../src/math/packageMacroCompatibility.ts";

const wrapperCommands = [
  "\\bm", "\\mathbfit", "\\symbfit", "\\simbfit", "\\symbf",
  "\\symbfup", "\\symbb", "\\symcal", "\\symfrak", "\\symtt",
  "\\boldmath", "\\bold", "\\pmb", "\\bf", "\\Bbb", "\\frak",
  "\\abs", "\\norm", "\\dd", "\\bra", "\\ket", "\\expval",
  "\\vb", "\\va", "\\vu",
];
for (const command of wrapperCommands) {
  assert.ok(
    commandRegistry.some((entry) => entry.command === command),
    `${command} is missing from the command registry`,
  );
  assert.ok(
    compatibilityWrapperPreviews.has(command) || command === "\\bm" || command === "\\mathbfit",
    `${command} is missing wrapper-input registration`,
  );
}

for (const command of ["\\comm", "\\acomm", "\\pb", "\\dv", "\\pdv", "\\braket", "\\ketbra", "\\mel"]) {
  assert.ok(commandRegistry.some((entry) => entry.command === command), `${command} is missing from the command registry`);
  assert.ok(compatibilityRawPlaceholderTemplates.has(command), `${command} is missing structural placeholder registration`);
}

assert.equal(compatibilityWrapperCanonicalTargets.get("\\boldmath"), "\\mathbfit");
assert.equal(compatibilityWrapperCanonicalTargets.get("\\simbfit"), "\\symbfit");
assert.equal(compatibilityWrapperCanonicalTargets.get("\\Bbb"), "\\mathbb");
assert.equal(compatibilityWrapperCanonicalTargets.get("\\frak"), "\\mathfrak");
assert.ok(compatibilityCommandNames.has("symbfit"));
assert.ok(compatibilityCommandNames.has("boldmath"));
assert.ok(compatibilityCommandNames.has("dv"));
assert.equal(compatibilityRequiredArgumentCounts.get("symbfit"), 1);
assert.equal(compatibilityRequiredArgumentCounts.get("abs"), 1);
assert.equal(compatibilityRequiredArgumentCounts.get("dv"), 2);
assert.equal(compatibilityRequiredArgumentCounts.get("mel"), 3);
assert.equal(compatibilityRequiredArgumentCounts.get("braket"), 1);

for (const source of [
  String.raw`\symbfit{J}`,
  String.raw`\bm{\alpha J}`,
  String.raw`\abs{x}`,
  String.raw`\dv{f}{x}`,
  String.raw`\braket{\phi}{\psi}`,
  String.raw`\braket{\phi|\psi}`,
  String.raw`\braket{\phi|A|\psi}`,
]) {
  const parsed = parseLatexSourceDraft(source, "raw");
  assert.equal(parsed.valid, true, `source editor rejected ${source}: ${parsed.error}`);
}

for (const source of [
  String.raw`\bm`, String.raw`\mathbfit`, String.raw`\symbfit`, String.raw`\abs`,
  String.raw`\dv`, String.raw`\dv{f}`, String.raw`\mel{a}{b}`,
]) {
  const parsed = parseLatexSourceDraft(source, "raw");
  assert.equal(parsed.valid, false, `source editor accepted incomplete ${source}`);
  assert.equal(parsed.error, "incomplete-command-arguments", `source editor returned the wrong error for ${source}`);
}

const exportSource = String.raw`\symbfit{J}+\bm{\alpha}+\abs{x}+\dv{f}{x}`;
assert.equal(formatLatex(exportSource, "raw"), exportSource, "raw source copy must preserve compatibility command spellings");
assert.equal(formatLatex(exportSource, "display-dollar"), `$$\n${exportSource}\n$$`, "formatted source copy must preserve compatibility command spellings");
const normalizedExportSource = normalizePackageLatexCommands(exportSource);
assert.doesNotMatch(normalizedExportSource, /\\(?:symbfit|bm|abs|dv)(?![A-Za-z@])/u, "supported aliases must normalize before the core parser");
assert.match(normalizedExportSource, /\\boldsymbol/u);
assert.match(normalizedExportSource, /\\left\|/u);
assert.match(normalizedExportSource, /\\frac/u);

for (const [source, pattern] of [
  [String.raw`\braket{\phi}{\psi}`, /\\left\\langle\s*\\phi\\middle\\vert\s*\\psi\\right\\rangle/u],
  [String.raw`\braket{\phi|\psi}`, /\\left\\langle\s*\\phi\\middle\\vert\s*\\psi\\right\\rangle/u],
  [String.raw`\braket{\phi|A|\psi}`, /\\left\\langle\s*\\phi\\middle\\vert\s*A\\middle\\vert\s*\\psi\\right\\rangle/u],
] as const) {
  const normalized = normalizePackageLatexCommands(source);
  assert.doesNotMatch(normalized, /\\braket(?![A-Za-z@])/u);
  assert.match(normalized, pattern);
}

const physicsAndUnitSource = String.raw`\mqty{a&b\\c&d}+\qtyrange{1}{10}{\milli\second}+\ang{30}+\unit{\joule\per\mole}`;
for (const [source, expected] of [
  [String.raw`\braket{\phi|\psi}`, String.raw`\left\langle \phi\middle\vert \psi\right\rangle`],
  [String.raw`\braket{\phi}{\psi}`, String.raw`\left\langle \phi\middle\vert \psi\right\rangle`],
  [String.raw`\braket{a||b}`, String.raw`\left\langle a\middle\Vert b\right\rangle`],
] as const) {
  assert.equal(
    normalizePackageLatexCommands(source),
    expected,
    `braket compatibility normalization failed for ${source}`,
  );
}

const qqBoundarySource = String.raw`x\qq{if}y>0`;
assert.equal(
  normalizePackageLatexCommands(qqBoundarySource),
  String.raw`x\quad\text{if}\quad{}y>0`,
  "compatibility expansion must preserve TeX control-word boundaries",
);

const normalizedPhysicsAndUnit = normalizePackageLatexCommands(physicsAndUnitSource);
assert.doesNotMatch(normalizedPhysicsAndUnit, /\\(?:mqty|qtyrange|ang|unit|milli|second|joule|per|mole)(?![A-Za-z@])/u);
assert.match(normalizedPhysicsAndUnit, /\\begin\{matrix\}/u);
assert.match(normalizedPhysicsAndUnit, /\\mathrm\{J\}/u);
assert.match(normalizedPhysicsAndUnit, /\\mathrm\{mol\}/u);
const normalizedAnticomm = normalizePackageLatexCommands(String.raw`\anticomm{A}{B}`);
assert.doesNotMatch(normalizedAnticomm, /\\anticomm(?![A-Za-z@])/u);
assert.match(normalizedAnticomm, /\\left\\\{A,B\\right\\\}/u);

for (const [source, expectedEnvironment] of [
  [String.raw`\mqty(a&b\\c&d)`, "pmatrix"],
  [String.raw`\mqty[a&b\\c&d]`, "bmatrix"],
  [String.raw`\mqty|a&b\\c&d|`, "vmatrix"],
  [String.raw`\mqty||a&b\\c&d||`, "Vmatrix"],
] as const) {
  const normalized = normalizePackageLatexCommands(source);
  assert.equal(
    normalized,
    `\\begin{${expectedEnvironment}}a&b\\\\c&d\\end{${expectedEnvironment}}`,
    `physics mqty delimiter normalization failed for ${source}`,
  );
}

for (const [name, [definition]] of Object.entries(VISUALTEX_PACKAGE_COMPATIBILITY_MACROS)) {
  assert.doesNotMatch(
    definition,
    /\\[A-Za-z@]+#\d/u,
    `${name} compatibility macro concatenates a TeX control word with an argument placeholder`,
  );
}

for (const command of ["bm", "abs", "bra", "ketbra", "dv", "pdv", "vb", "va", "vu", "anticomm"]) {
  assert.ok(command in VISUALTEX_PACKAGE_COMPATIBILITY_MACROS, `${command} is missing from renderer-independent package compatibility data`);
}

for (const command of ["\\arccos", "\\arctan"]) {
  assert.ok(commandRegistry.some((entry) => entry.command === command), `${command} is missing from the command registry`);
}

console.log("VisualTeX renderer-independent math variant/package compatibility regression passed");
