import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { convertLatexToMarkup, validateLatex } from "mathlive/ssr";
import { VISUALTEX_CORE_MATHLIVE_MACROS } from "../src/math/coreLatexAliases.ts";
import { VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS } from "../src/math/mathLiveCompatibilityMacros.ts";
import {
  VISUALTEX_MATHJAX_PACKAGE_MACROS,
  VISUALTEX_MATHLIVE_PACKAGE_MACROS,
} from "../src/math/packageMacroCompatibility.ts";
import { VISUALTEX_PHYSICS_KERNEL_MACROS } from "../src/math/physicsKernelMacros.ts";
import { physicsToolbarCommands } from "../src/autocomplete/physicsToolbarCommands.ts";
import { parseLatexSourceDraft } from "../src/clipboard/LatexCopyService.ts";
import { patchVisualTexMathLiveCoreCompatibility } from "../vite.mathliveCoreCompatibility.ts";

const forbidden = new Set([
  // physics.sty replaces the meanings of these existing commands. \qty is
  // also used by siunitx for quantities with units.
  "div", "qty", "Re", "Im", "real", "sin", "cos", "tan", "exp",
  "log", "ln", "det", "Pr", "arccsc", "arcsec",
]);
const seen = new Set<string>();

for (const macro of VISUALTEX_PHYSICS_KERNEL_MACROS) {
  assert.equal(seen.has(macro.name), false, `Duplicate physics command \\${macro.name}`);
  seen.add(macro.name);
  assert.equal(forbidden.has(macro.name), false, `Conflicting command \\${macro.name}`);
  assert.equal(
    VISUALTEX_CORE_MATHLIVE_MACROS.some((existing) => existing.name === macro.name) ||
      macro.name in VISUALTEX_MATHLIVE_PACKAGE_MACROS,
    false,
    `\\${macro.name} replaces an existing VisualTeX command`,
  );
  assert.ok(
    validateLatex(`\\${macro.name}`).some((error) => error.code === "unknown-command"),
    `\\${macro.name} already has a MathLive meaning`,
  );

  const args = ["{x}", "{y}", "{z}"].slice(0, macro.args);
  const latex = `\\${macro.name}${args.join("")}`;
  const actual = convertLatexToMarkup(latex, {
    macros: VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS,
  });
  assert.doesNotMatch(actual, /ML__error/, `${latex} renders as an error`);
  assert.deepEqual(
    VISUALTEX_MATHJAX_PACKAGE_MACROS[macro.name],
    [macro.def, macro.args],
    `${latex} is missing from the MathJax export fallback`,
  );
}

assert.ok(physicsToolbarCommands.length >= 25, "Too few physics package toolbar commands");
const toolbarIds = new Set<string>();
for (const command of physicsToolbarCommands) {
  assert.equal(command.category, "physics");
  assert.equal(toolbarIds.has(command.id), false, `Duplicate toolbar tile ${command.id}`);
  toolbarIds.add(command.id);
  const name = command.command.slice(1);
  assert.ok(seen.has(name), `${command.command} has no non-conflicting kernel definition`);
  assert.ok(command.insertTemplate.startsWith(command.command), `${command.id} changed its source command`);
  for (const latex of [
    command.previewLatex,
    command.insertTemplate.replaceAll("\\placeholder{}", "x"),
  ]) {
    assert.doesNotMatch(
      convertLatexToMarkup(latex, { macros: VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS }),
      /ML__error/,
      `${command.id} cannot render ${latex}`,
    );
  }
}
assert.equal(toolbarIds.has("physics-div"), false, "The original \\div must not be replaced");
assert.equal(toolbarIds.has("physics-qty"), false, "Conflicting \\qty entered the toolbar");

for (const latex of [String.raw`\div`, String.raw`\Re`, String.raw`\Im`, String.raw`\sin`, String.raw`\det`]) {
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
assert.equal(
  VISUALTEX_MATHLIVE_PACKAGE_MACROS.dv.def,
  String.raw`\frac{\mathrm{d}#1}{\mathrm{d}#2}`,
  "The existing VisualTeX derivative shorthand changed",
);
assert.equal(
  VISUALTEX_MATHLIVE_PACKAGE_MACROS.pdv.def,
  String.raw`\frac{\partial #1}{\partial #2}`,
  "The existing VisualTeX partial derivative shorthand changed",
);
for (const latex of [
  String.raw`\pmqty{1&0\\0&1}`,
  String.raw`\quantity{x}`,
  String.raw`\qq{if}`,
  String.raw`\derivative{f}{x}`,
]) {
  assert.equal(parseLatexSourceDraft(latex, "raw").valid, true, `${latex} was rejected as source`);
}
assert.equal(
  parseLatexSourceDraft(String.raw`\pmqty{\notACommand&0\\0&1}`, "raw").error,
  "unknown-command",
  "Matrix validation masked an unknown command inside a cell",
);

const originalKernel = readFileSync("node_modules/mathlive/mathlive.mjs", "utf8");
const patchedKernel = patchVisualTexMathLiveCoreCompatibility(originalKernel);
const addedMacros = patchedKernel.slice(
  patchedKernel.indexOf("var DEFAULT_MACROS = {") + "var DEFAULT_MACROS = {".length,
  patchedKernel.indexOf("\n  \"strut\":", patchedKernel.indexOf("var DEFAULT_MACROS = {")),
);
for (const macro of VISUALTEX_PHYSICS_KERNEL_MACROS) {
  assert.ok(addedMacros.includes(`${JSON.stringify(macro.name)}: { def:`), `\\${macro.name} missing from the kernel`);
}
assert.doesNotMatch(addedMacros, /"(?:div|qty|Re|Im)":/, "A conflicting command entered the kernel");

console.log(`Physics kernel compatibility passed (${seen.size} non-conflicting command names).`);
