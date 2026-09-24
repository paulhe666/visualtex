import assert from "node:assert/strict";
import {
  compatibilityDirectPlaceholderWrappers,
  compatibilitySourceWrapperKinds,
  compatibilityWrapperCommands,
} from "../src/autocomplete/compatibilityCommands.ts";
import { commandRegistry } from "../src/autocomplete/commandRegistry.ts";
import {
  accentSourceWrapperCommands,
  primaryVariantSourceWrapperCommands,
  sourceWrapperCanonicalTargets,
  sourceWrapperKinds,
} from "../src/editor/sourceWrapperRegistry.ts";

for (const command of compatibilityDirectPlaceholderWrappers) {
  assert.ok(
    compatibilityWrapperCommands.has(command),
    `${command} direct-placeholder exception is not a wrapper command`,
  );
  assert.equal(
    sourceWrapperKinds.has(command),
    false,
    `${command} cannot be both source-first and direct-placeholder`,
  );
}

for (const [command, kind] of compatibilitySourceWrapperKinds) {
  assert.ok(
    compatibilityWrapperCommands.has(command),
    `${command} source-wrapper registration is not a wrapper command`,
  );
  assert.equal(
    compatibilityDirectPlaceholderWrappers.has(command),
    false,
    `${command} source-wrapper registration conflicts with direct-placeholder`,
  );
  assert.equal(
    sourceWrapperKinds.get(command),
    kind,
    `${command} did not reach the unified source-wrapper registry`,
  );
}

const allCompatibilityWrapperCommands = new Set([
  ...compatibilityDirectPlaceholderWrappers,
  ...compatibilitySourceWrapperKinds.keys(),
]);
assert.deepEqual(
  [...compatibilityWrapperCommands].filter(
    (command) => !allCompatibilityWrapperCommands.has(command),
  ),
  [],
  "A compatibility wrapper has no explicit input behavior classification",
);

for (const command of primaryVariantSourceWrapperCommands) {
  assert.equal(sourceWrapperKinds.get(command), "variant", `${command} variant missing`);
}
for (const command of accentSourceWrapperCommands) {
  assert.equal(sourceWrapperKinds.get(command), "accent", `${command} accent missing`);
  assert.ok(
    commandRegistry.some((entry) => entry.command === command),
    `${command} accent missing from command registry`,
  );
}
for (const [alias, canonical] of sourceWrapperCanonicalTargets) {
  assert.ok(sourceWrapperKinds.has(alias), `${alias} alias missing source-wrapper kind`);
  assert.ok(sourceWrapperKinds.has(canonical), `${alias} canonical target ${canonical} missing`);
}

for (const command of ["\\vb", "\\va", "\\vu"]) {
  assert.ok(sourceWrapperKinds.has(command), `${command} physics vector wrapper escaped source-first input`);
}
assert.equal(sourceWrapperKinds.get("\\va"), "accent");
assert.equal(sourceWrapperKinds.get("\\vu"), "accent");
assert.equal(sourceWrapperKinds.get("\\vb"), "variant");
assert.equal(sourceWrapperKinds.get("\\c"), "accent");

console.log(
  `Source-wrapper registry regression passed: ${sourceWrapperKinds.size} source-first commands, ${compatibilityDirectPlaceholderWrappers.size} explicit direct-placeholder exceptions.`,
);
