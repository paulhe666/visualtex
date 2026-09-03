import assert from "node:assert/strict";
import {
  decodeMacosDocumentImportProgress,
  decodeMacosDocumentImportRequest,
  decodeMacosLatexRedrawFontSizes,
} from "../src/office/documentImport/documentImportPayloadValidation.ts";

const request = {
  protocolVersion: 1,
  sessionId: "12345678-1234-4234-9234-123456789abc",
  host: "word",
  sourceDocumentId: "doc-1",
  bookmarkName: "VisualTeXImport",
  defaultFontSizePt: 12,
  operation: "formulaRestore",
  redrawScope: "selection",
  outputKind: "image",
  sourceKind: "omml",
  source: "x",
  restoreTargets: [
    {
      sourceStart: 1,
      sourceEnd: 4,
      sourceText: "x^2",
      displayMode: "inline",
      fontSizePt: 12,
      sourceKind: "omml",
      mathMl: "<math/>",
      latex: "x^2",
    },
  ],
};
assert.equal(decodeMacosDocumentImportRequest(request), request);
assert.throws(
  () => decodeMacosDocumentImportRequest({ ...request, protocolVersion: 2 }),
  /protocolVersion/,
);
assert.throws(
  () =>
    decodeMacosDocumentImportRequest({
      ...request,
      restoreTargets: [{ ...request.restoreTargets[0], sourceEnd: 0 }],
    }),
  /sourceEnd/,
);
assert.throws(
  () =>
    decodeMacosDocumentImportRequest({
      ...request,
      restoreTargets: [{ ...request.restoreTargets[0], fontSizePt: Number.NaN }],
    }),
  /fontSizePt/,
);

assert.deepEqual(decodeMacosLatexRedrawFontSizes([10, 12.5, 20]), [10, 12.5, 20]);
assert.throws(() => decodeMacosLatexRedrawFontSizes([12, "14"]), /\[1\]/);

const progress = { current: 3, total: 10, stage: "inserting" };
assert.equal(decodeMacosDocumentImportProgress(progress), progress);
assert.throws(
  () => decodeMacosDocumentImportProgress({ ...progress, current: 11 }),
  /\.current/,
);

console.log("VisualTeX macOS document import payload safety regression passed");
