import assert from "node:assert/strict";
import {
  decodeNativeSystemMathFontProbes,
  decodeNativeSystemMathGlyphOutline,
} from "../src/math/systemMathGlyphPayloadValidation.ts";
import {
  decodeQuickOcrCapture,
  decodeSilentOcrShortcut,
} from "../src/ocr/quickOcrPayloadValidation.ts";

const probes = [
  {
    requestedFamily: "Cambria Math",
    resolvedFamily: "Cambria Math",
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
  requestedFamily: "Cambria Math",
  resolvedFamily: "Cambria Math",
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
      metrics: { ...outline.metrics, widthEm: Number.NaN },
    }),
  /widthEm/,
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

assert.equal(decodeSilentOcrShortcut("Ctrl+Alt+O"), "Ctrl+Alt+O");
assert.throws(() => decodeSilentOcrShortcut({}), /shortcut/);

console.log("VisualTeX Windows native UI payload safety regression passed");
