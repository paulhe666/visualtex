import assert from "node:assert/strict";
import {
  decodeNativeSystemMathFontProbes,
  decodeNativeSystemMathGlyphOutline,
} from "../src/math/systemMathGlyphPayloadValidation.ts";
import { decodeQuickOcrCapture } from "../src/ocr/quickOcrPayloadValidation.ts";

const probes = [
  {
    requestedFamily: "STIX Two Math",
    resolvedFamily: "STIX Two Math",
    available: true,
  },
];
assert.deepEqual(decodeNativeSystemMathFontProbes(probes), probes);
assert.throws(
  () => decodeNativeSystemMathFontProbes([{ requestedFamily: null }]),
  /probe\[0\]/,
);

const outline = {
  character: "∫",
  requestedFamily: "STIX Two Math",
  resolvedFamily: "STIX Two Math",
  fallbackUsed: false,
  glyphId: 123,
  path: "M0 0L1 1Z",
  metrics: {
    widthEm: 0.5,
    ascentEm: 0.8,
    descentEm: 0.2,
  },
};
assert.deepEqual(decodeNativeSystemMathGlyphOutline(outline), outline);
assert.throws(
  () =>
    decodeNativeSystemMathGlyphOutline({
      ...outline,
      metrics: { ...outline.metrics, descentEm: Infinity },
    }),
  /descentEm/,
);

assert.equal(decodeQuickOcrCapture(null), null);
assert.deepEqual(
  decodeQuickOcrCapture({ dataBase64: "YWJj", extension: "PNG" }),
  { dataBase64: "YWJj", extension: "png" },
);
assert.throws(
  () => decodeQuickOcrCapture({ dataBase64: "YWJj", extension: "../png" }),
  /extension/,
);
assert.throws(
  () => decodeQuickOcrCapture({ dataBase64: null, extension: "png" }),
  /image data/,
);

console.log("VisualTeX macOS native UI payload safety regression passed");
