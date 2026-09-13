import assert from "node:assert/strict";

import { reusableImageGeometry } from "../src/office/redraw/wordLatexRedrawGeometry.ts";
import type { VisualTeXFormulaMetadata } from "../src/office/shared/formulaMetadata.ts";

const metadata: VisualTeXFormulaMetadata = {
  schema: "visualtex-formula",
  schemaVersion: 1,
  formulaId: "11111111-1111-4111-8111-111111111111",
  title: "global conversion geometry fixture",
  latex: String.raw`\frac{-b\pm\sqrt{b^2-4ac}}{2a}`,
  lines: [
    {
      id: "22222222-2222-4222-8222-222222222222",
      latex: String.raw`\frac{-b\pm\sqrt{b^2-4ac}}{2a}`,
    },
  ],
  codeFormat: "raw",
  displayMode: "block",
  numbered: true,
  fontSizePt: 11,
  renderWidthPx: 175.32266666666663,
  renderHeightPx: 47.651999999999994,
  referenceWidthPt: 140.30196399999994,
  referenceHeightPt: 35.739,
  referenceBaselinePt: -11.244,
  formulaLetterFont: "times",
  formulaChineseFont: "songti",
  createdWithVersion: "1.2.6",
  updatedWithVersion: "1.2.6",
  createdAt: "2026-09-12T00:00:00.000Z",
  updatedAt: "2026-09-12T00:00:00.000Z",
};

const geometry = reusableImageGeometry(metadata);
assert.ok(geometry, "current image metadata must provide reusable OMML geometry");
assert.equal(geometry.width, metadata.renderWidthPx);
assert.equal(geometry.height, metadata.renderHeightPx);
assert.equal(geometry.referenceWidthPt, metadata.referenceWidthPt);
assert.equal(geometry.referenceHeightPt, metadata.referenceHeightPt);
assert.equal(geometry.referenceBaselinePt, metadata.referenceBaselinePt);
assert.ok(
  geometry.baseline! > 0 && geometry.baseline! < geometry.height!,
  "the recovered MathJax baseline must remain inside the original image bounds",
);

assert.equal(
  reusableImageGeometry({
    ...metadata,
    renderWidthPx: undefined,
    renderHeightPx: undefined,
  })?.referenceWidthPt,
  metadata.referenceWidthPt,
  "legacy metadata can recover render bounds from its stable Word reference geometry",
);

assert.equal(
  reusableImageGeometry({
    ...metadata,
    referenceWidthPt: undefined,
  }),
  null,
  "truly old metadata must use the vector-bounds compatibility render",
);

console.log("VisualTeX global conversion artifact regression: PASS");
