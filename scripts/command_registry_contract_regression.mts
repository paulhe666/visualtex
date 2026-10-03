import assert from "node:assert/strict";

import { commandRegistry } from "../src/autocomplete/commandRegistry";
import {
  compatibilityCommandNames,
  compatibilityCommands,
} from "../src/autocomplete/compatibilityCommands";

const duplicateIds = [...new Set(
  commandRegistry
    .map((command) => command.id)
    .filter((id, index, ids) => ids.indexOf(id) !== index),
)];
assert.deepEqual(
  duplicateIds,
  [],
  `Command registry contains duplicate ids: ${duplicateIds.join(", ")}`,
);

const unicodeCommandEntries = commandRegistry.filter((command) =>
  [...command.command].some((character) => (character.codePointAt(0) ?? 0) > 127),
);
assert.deepEqual(
  unicodeCommandEntries.map(({ id, command }) => ({ id, command })),
  [],
  "Named toolbar/autocomplete commands should use LaTeX source spellings rather than literal Unicode glyphs",
);

const daggerEntries = commandRegistry.filter((command) => command.id === "dagger");
assert.equal(daggerEntries.length, 1);
assert.equal(daggerEntries[0]?.command, "\\dagger");
const inputOnlyCompatibilityAliases = compatibilityCommands
  .filter(
    (command) =>
      !compatibilityCommandNames.has(command.command.replace(/^\\/, "")),
  )
  .map((command) => command.command)
  .sort();
assert.deepEqual(inputOnlyCompatibilityAliases, ["\\simbfit"]);
assert.equal(
  commandRegistry.find(entry => entry.command === "\\simbfit")?.insertTemplate.split("{")[0],
  "\\symbfit",
);

console.log(
  `Command registry contract regression passed (${commandRegistry.length} commands)`,
);
