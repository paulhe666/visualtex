import assert from "node:assert/strict";
import { commandRegistry } from "../src/autocomplete/commandRegistry.ts";
import { hasBoundedOperatorPlaceholderOrder } from "../src/editor/boundedOperatorTemplate.ts";

const expectedEditableBoundedOperatorIds = [
  "int",
  "sum",
  "prod",
  "iint-bounds",
  "iiint-bounds",
  "oint-bounds",
  "sum-finite",
  "prod-finite",
  "coproduct",
  "fint",
  "dashint",
  "ddashint",
  "oiint",
  "oiiint",
  "ointctrclockwise",
  "varointclockwise",
  "varointctrclockwise",
] as const;

const byId = new Map(commandRegistry.map((command) => [command.id, command]));
for (const id of expectedEditableBoundedOperatorIds) {
  const command = byId.get(id);
  assert.ok(command, `Missing bounded-operator command: ${id}`);
  assert.equal(
    hasBoundedOperatorPlaceholderOrder(command.insertTemplate),
    true,
    `${id} must use lower-first bounded-operator navigation`,
  );
}

for (const id of [
  "scripts",
  "evalbar",
  "series",
  "productseries",
  "bigcap-limits",
  "bigcup-limits",
  "bigsqcup-limits",
  "bigvee-limits",
  "bigwedge-limits",
  "bigodot-limits",
  "bigoplus-limits",
  "bigotimes-limits",
  "biguplus-limits",
]) {
  const command = byId.get(id);
  assert.ok(command, `Missing comparison command: ${id}`);
  assert.equal(
    hasBoundedOperatorPlaceholderOrder(command.insertTemplate),
    false,
    `${id} must not be forced into editable lower->upper operator navigation`,
  );
}

const detected = commandRegistry
  .filter((command) => hasBoundedOperatorPlaceholderOrder(command.insertTemplate))
  .map((command) => command.id)
  .sort();
const expected = [...expectedEditableBoundedOperatorIds].sort();
assert.deepEqual(
  detected,
  expected,
  `Bounded-operator detector coverage drifted. Detected: ${detected.join(", ")}`,
);

console.log(
  `Bounded operator template audit passed: ${detected.length} editable lower+upper operator templates detected exactly.`,
);
console.log(detected.join(", "));
