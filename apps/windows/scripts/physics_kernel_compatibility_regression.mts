import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { convertLatexToMarkup, validateLatex } from "mathlive/ssr";
import { VISUALTEX_CORE_MATHLIVE_MACROS } from "../src/math/coreLatexAliases.ts";
import { VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS } from "../src/math/mathLiveCompatibilityMacros.ts";
import { VISUALTEX_PACKAGE_COMPATIBILITY_MACROS } from "../src/math/packageMacroCompatibility.ts";
import { VISUALTEX_PHYSICS_KERNEL_MACROS } from "../src/math/physicsKernelMacros.ts";
import { parseLatexSourceDraft } from "../src/clipboard/LatexCopyService.ts";
import { patchVisualTexMathLiveCoreCompatibility } from "../vite.mathliveCoreCompatibility.ts";

const forbidden = new Set([
  "div", "qty", "Re", "Im", "real", "sin", "cos", "tan", "exp",
  "log", "ln", "det", "Pr", "arccsc", "arcsec",
]);
const seen = new Set<string>();

for (const macro of VISUALTEX_PHYSICS_KERNEL_MACROS) {
  assert.equal(seen.has(macro.name), false, `Duplicate physics command \\${macro.name}`);
  seen.add(macro.name);
  assert.equal(forbidden.has(macro.name), false, `Conflicting command \\${macro.name}`);
  assert.equal(
    VISUALTEX_CORE_MATHLIVE_MACROS.some((existing) => existing.name === macro.name),
    false,
    `\\${macro.name} replaces an existing core VisualTeX command`,
  );

  const args = ["{x}", "{y}", "{z}"].slice(0, macro.args);
  const latex = `\\${macro.name}${args.join("")}`;
  const actual = convertLatexToMarkup(latex, {
    macros: VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS,
  });
  assert.doesNotMatch(actual, /ML__error/, `${latex} renders as an error`);
  assert.deepEqual(
    VISUALTEX_PACKAGE_COMPATIBILITY_MACROS[macro.name],
    [macro.def, macro.args],
    `${latex} is missing from the renderer-independent export fallback`,
  );

  const validation = validateLatex(`\\${macro.name}`);
  // If MathLive learns a command natively in a future release that is fine as
  // long as VisualTeX does not override one of the explicitly forbidden names.
  if (!validation.some((error) => error.code === "unknown-command")) {
    assert.equal(forbidden.has(macro.name), false);
  }
}

for (const latex of [
  String.raw`\div`,
  String.raw`\Re`,
  String.raw`\Im`,
  String.raw`\sin`,
  String.raw`\det`,
]) {
  assert.equal(
    convertLatexToMarkup(latex, { macros: VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS }),
    convertLatexToMarkup(latex),
    `${latex} changed meaning`,
  );
}

const definition = (name: string) =>
  VISUALTEX_PHYSICS_KERNEL_MACROS.find((macro) => macro.name === name)?.def;
assert.equal(definition("derivative"), String.raw`\frac{\mathrm{d}#1}{\mathrm{d}#2}`);
assert.equal(definition("partialderivative"), String.raw`\frac{\partial #1}{\partial #2}`);
assert.equal(definition("functionalderivative"), String.raw`\frac{\delta #1}{\delta #2}`);
assert.deepEqual(
  VISUALTEX_PACKAGE_COMPATIBILITY_MACROS.dv,
  [String.raw`\frac{\mathrm{d}#1}{\mathrm{d}#2}`, 2],
  "The existing VisualTeX derivative shorthand changed",
);
assert.deepEqual(
  VISUALTEX_PACKAGE_COMPATIBILITY_MACROS.pdv,
  [String.raw`\frac{\partial #1}{\partial #2}`, 2],
  "The existing VisualTeX partial derivative shorthand changed",
);

for (const latex of [
  String.raw`\pmqty{1&0\\0&1}`,
  String.raw`\quantity{x}`,
  String.raw`\qq{if}`,
  String.raw`\derivative{f}{x}`,
]) {
  assert.equal(
    parseLatexSourceDraft(latex, "raw").valid,
    true,
    `${latex} was rejected as source`,
  );
}
assert.equal(
  parseLatexSourceDraft(String.raw`\pmqty{\notACommand&0\\0&1}`, "raw").error,
  "unknown-command",
  "Matrix validation masked an unknown command inside a cell",
);

const originalKernel = readFileSync("node_modules/mathlive/mathlive.mjs", "utf8");
const patchedKernel = patchVisualTexMathLiveCoreCompatibility(originalKernel);
const macroStart = patchedKernel.indexOf("var DEFAULT_MACROS = {");
const addedMacros = patchedKernel.slice(
  macroStart + "var DEFAULT_MACROS = {".length,
  patchedKernel.indexOf("\n  \"strut\":", macroStart),
);
for (const macro of VISUALTEX_PHYSICS_KERNEL_MACROS) {
  assert.ok(
    addedMacros.includes(`${JSON.stringify(macro.name)}: { def:`),
    `\\${macro.name} missing from the kernel`,
  );
}
assert.doesNotMatch(
  addedMacros,
  /"(?:div|qty|Re|Im)":/,
  "A conflicting command entered the kernel",
);

console.log(
  `Physics kernel compatibility passed (${seen.size} non-conflicting command names).`,
);
