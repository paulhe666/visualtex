import assert from "node:assert/strict";
import {
  decodeCachedFormulaMetadata,
  decodeCompanionHealth,
  decodeNativePowerPointSelection,
  decodeNativePowerPointSlideSnapshot,
  decodeNativeWordInlineBaselineResult,
  decodePowerPointInteractionEvents,
} from "../src/office/api/companionPayloadValidation.ts";

const health = {
  ok: true,
  appVersion: "1.2.5",
  officeUiVersion: "1.2.5",
  protocolVersion: 1,
  ocrAvailable: true,
};
assert.equal(decodeCompanionHealth(health), health);
assert.throws(
  () => decodeCompanionHealth({ ...health, protocolVersion: "1" }),
  /health\.protocolVersion/,
);
assert.throws(() => decodeCompanionHealth(null), /health/);

const metadata = {
  schema: "visualtex-formula",
  schemaVersion: 1,
  formulaId: "4bf2217c-f29e-4f77-98b8-7258be8f63ae",
  title: "",
  latex: "x^2",
  lines: [{ id: "line-1", latex: "x^2" }],
  codeFormat: "raw",
  displayMode: "block",
  numbered: false,
  createdWithVersion: "1.2.5",
  updatedWithVersion: "1.2.5",
  createdAt: "2026-09-02T00:00:00.000Z",
  updatedAt: "2026-09-02T00:00:00.000Z",
} as const;
assert.equal(decodeCachedFormulaMetadata(metadata), metadata);
assert.throws(
  () => decodeCachedFormulaMetadata({ ...metadata, formulaId: "not-a-uuid" }),
  /formulaMetadata/,
);

const selection = {
  shapeName: "VisualTeX_1",
  slideIndex: 1,
  slideId: 256,
  presentationIdentity: "Deck.pptx",
  left: -5,
  top: 10,
  width: 120.5,
  height: 32,
};
assert.equal(decodeNativePowerPointSelection(selection), selection);
assert.throws(
  () => decodeNativePowerPointSelection({ ...selection, width: "120" }),
  /powerPointSelection\.width/,
);
assert.throws(
  () => decodeNativePowerPointSelection({ ...selection, slideIndex: 0 }),
  /powerPointSelection\.slideIndex/,
);

const snapshot = {
  presentationIdentity: "Deck.pptx",
  slideIndex: 1,
  slideId: 256,
  shapeCount: 2,
  shapeNames: ["Title 1", "VisualTeX_1"],
};
assert.equal(decodeNativePowerPointSlideSnapshot(snapshot), snapshot);
assert.throws(
  () =>
    decodeNativePowerPointSlideSnapshot({
      ...snapshot,
      shapeNames: ["Title 1", null],
    }),
  /shapeNames\[1\]/,
);

const baseline = {
  appliedPosition: -2.5,
  width: 80,
  height: 32,
  matchedShapeIndex: 1,
};
assert.equal(decodeNativeWordInlineBaselineResult(baseline), baseline);
assert.equal(
  decodeNativeWordInlineBaselineResult({ ...baseline, matchedShapeIndex: 0 })
    .matchedShapeIndex,
  0,
);
assert.throws(
  () =>
    decodeNativeWordInlineBaselineResult({
      ...baseline,
      matchedShapeIndex: -1,
    }),
  /matchedShapeIndex/,
);

const events = [
  {
    cursor: 4,
    host: "powerpoint",
    kind: "edit-selected",
    formulaId: metadata.formulaId,
    shapeName: selection.shapeName,
    slideIndex: 1,
    slideId: 256,
    presentationIdentity: "Deck.pptx",
    left: -5,
    top: 10,
    width: 120.5,
    height: 32,
    createdAt: 1_788_278_400_000,
  },
];
assert.deepEqual(decodePowerPointInteractionEvents(events), events);
assert.throws(
  () => decodePowerPointInteractionEvents({ events }),
  /powerPointInteractionEvents/,
);
assert.throws(
  () =>
    decodePowerPointInteractionEvents([
      { ...events[0], kind: "open", width: Number.NaN },
    ]),
  /\.kind/,
);
assert.throws(
  () =>
    decodePowerPointInteractionEvents([
      { ...events[0], width: Number.POSITIVE_INFINITY },
    ]),
  /\.width/,
);

console.log("VisualTeX companion payload safety regression passed");
