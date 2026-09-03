#!/usr/bin/env -S npx tsx

import assert from "node:assert/strict";
import { validateLatex } from "mathlive/ssr";
import {
  calculusCommandIds,
  commandRegistry,
} from "../src/autocomplete/commandRegistry.ts";

const byId = new Map(commandRegistry.map((command) => [command.id, command]));

const structureTemplates = new Map([
  ["left-upper-script", String.raw`{}^{\placeholder{}}\placeholder{}`],
  ["left-scripts", String.raw`{}_{\placeholder{}}^{\placeholder{}}\placeholder{}`],
  ["left-lower-script", String.raw`{}_{\placeholder{}}\placeholder{}`],
  ["linear-fraction", String.raw`\placeholder{}/\placeholder{}`],
  ["skewed-fraction", String.raw`\nicefrac{\placeholder{}}{\placeholder{}}`],
]);

for (const [id, expectedTemplate] of structureTemplates) {
  const command = byId.get(id);
  assert.ok(command, `missing structure toolbar command: ${id}`);
  assert.equal(command.category, "structure", `${id} is not in the structure category`);
  assert.equal(command.insertTemplate, expectedTemplate, `${id} insert template changed`);
  assert.deepEqual(
    validateLatex(command.previewLatex),
    [],
    `${id} preview is not valid MathLive LaTeX: ${command.previewLatex}`,
  );
}

for (const id of ["power", "scripts", "subscript"]) {
  assert.ok(byId.has(id), `existing right-side script template disappeared: ${id}`);
}

const bareIntegralIds = [
  "int-bare",
  "iint-bare",
  "iiint-bare",
  "oint-bare",
  "oiint-bare",
  "oiiint-bare",
];
const noDifferentialIntegralIds = [
  "intplain-no-d",
  "int-bounds-no-d",
  "iint-no-d",
  "iint-bounds-no-d",
  "iiint-no-d",
  "iiint-bounds-no-d",
  "oint-no-d",
  "oint-bounds-no-d",
  "oiint-no-d",
  "oiiint-no-d",
];

for (const id of [...bareIntegralIds, ...noDifferentialIntegralIds]) {
  assert.ok(calculusCommandIds.includes(id), `${id} is missing from the calculus toolbar order`);
  const command = byId.get(id);
  assert.ok(command, `missing calculus toolbar command: ${id}`);
  assert.equal(command.category, "calculus", `${id} is not in the calculus category`);
  assert.deepEqual(
    validateLatex(command.previewLatex),
    [],
    `${id} preview is not valid MathLive LaTeX: ${command.previewLatex}`,
  );
}

for (const id of bareIntegralIds) {
  const command = byId.get(id)!;
  assert.equal(
    command.insertTemplate,
    command.command,
    `${id} must insert only the bare integral operator`,
  );
  assert.doesNotMatch(command.insertTemplate, /\\placeholder/);
  assert.doesNotMatch(command.insertTemplate, /\\mathrm\{d\}|\\differentialD/);
}

for (const id of noDifferentialIntegralIds) {
  const command = byId.get(id)!;
  assert.match(command.insertTemplate, /\\placeholder\{\}/, `${id} must retain editable placeholders`);
  assert.doesNotMatch(
    command.insertTemplate,
    /\\mathrm\{d\}|\\differentialD/,
    `${id} must not force a differential element`,
  );
}

for (const id of [
  "intplain",
  "int",
  "iint-bounds",
  "iiint-bounds",
  "oint-bounds",
  "lineintegral",
  "surfaceintegral",
  "iiint",
  "volumeintegral",
  "oint",
  "closed-surface-integral",
  "closed-volume-integral",
]) {
  const command = byId.get(id);
  assert.ok(command, `existing integral template disappeared: ${id}`);
  assert.match(
    command.insertTemplate,
    /\\mathrm\{d\}/,
    `${id} should remain available as the differential-bearing variant`,
  );
}

console.log(
  `Toolbar template completion regression passed: ${structureTemplates.size} structure additions, ${bareIntegralIds.length} bare integrals, ${noDifferentialIntegralIds.length} no-differential integrals.`,
);
