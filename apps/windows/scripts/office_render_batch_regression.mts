import assert from "node:assert/strict";
import { decodeOfficeRenderBatch, decodeNextOfficeRenderBatch } from "../src/office/shared/renderBatchContract.ts";
import { renderOfficeBatch } from "../src/office/dialog/renderOfficeBatch.ts";
import { latexToMathMl } from "../src/export/runtime.ts";

const session = {
  id: "session-a", mode: "create", host: "word",
  formulaId: "4bf2217c-f29e-4f77-98b8-7258be8f63ae",
  sourceDocumentId: null, sourceObjectId: null, title: "",
  lines: [{id: "line-a", latex: "x^2"}], activeLineId: "line-a",
  codeFormat: "latex", displayMode: "inline", objectMode: "wordOmml",
  numbered: false, mathTypeNumberPosition: "right", fontSizePt: 12,
  exportWidth: 0, exportHeight: 0, exportResult: null, originalMetadata: null,
  dirty: false, status: "created", autoCommitOnClose: false, explicitCancel: false,
  error: null, createdAt: 1, updatedAt: 1, expiresAt: 3,
};
const input = { id: "batch-a", items: [
  { id: "angle", session: { ...session, lines: [{id: "line-a", latex: String.raw`\mathit{\angle}`}] } },
  { id: "invalid", session: { ...session, lines: [{id: "line-b", latex: String.raw`\unknownVisualTexCommand`}] } },
  { id: "upright", session: { ...session, lines: [{id: "line-c", latex: String.raw`\mathrm{α}`}] } },
]};
const frozenBefore = JSON.stringify(input);
const batch = decodeOfficeRenderBatch(input);
let active = 0, maxActive = 0;
const result = await renderOfficeBatch(batch, async (value) => {
  active++; maxActive = Math.max(maxActive, active);
  try {
    const mathMl = await latexToMathMl(value.lines[0].latex, false);
    return {svg: "", svgBase64: "", mathMl, width: 20, height: 15, baseline: 12};
  } finally { active--; }
});
assert.equal(maxActive, 1);
assert.deepEqual(result.map(item => item.id), ["angle", "invalid", "upright"]);
// A single-scalar mi has implicit italic style under MathML semantics.
assert.match(result[0].export!.mathMl!, /mathvariant="italic"|<mi>&#x2220;<\/mi>/);
assert.match(result[1].error!, /unknown|undefined/i);
assert.match(result[2].export!.mathMl!, /mathvariant="normal"/);
assert.equal(JSON.stringify(input), frozenBefore);
assert.throws(() => decodeOfficeRenderBatch({...input, items: [input.items[0], input.items[0]]}), /duplicate/);
assert.throws(() => decodeOfficeRenderBatch({...input, items: [{id:"a", session:{}}]}));
assert.deepEqual(decodeNextOfficeRenderBatch({id:null}), {id:null});
assert.throws(() => decodeNextOfficeRenderBatch({id:42}));
console.log("Office ephemeral render batch: ordered serial rendering, correlated failure, style fidelity and payload validation passed");
