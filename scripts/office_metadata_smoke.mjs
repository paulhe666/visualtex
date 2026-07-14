import assert from "node:assert/strict";
import {
  decodePowerPointObjectReference,
  encodePowerPointObjectReference,
  formulaIdFromPowerPointShapeName,
  powerpointShapeName,
} from "../src/office/metadata/powerpointMetadata.ts";
import {
  createFormulaMetadata,
  decodeFormulaMetadata,
  encodeFormulaMetadata,
  formulaMetadataFromXml,
  formulaMetadataToXml,
  isVisualTeXFormulaMetadata,
  VISUALTEX_FORMULA_XML_NAMESPACE,
} from "../src/office/metadata/formulaMetadata.ts";

const formulaId = crypto.randomUUID();
const lines = [
  { id: crypto.randomUUID(), latex: String.raw`\frac{a+b}{c}` },
  { id: crypto.randomUUID(), latex: String.raw`\text{测试}+\sum_{i=1}^{n} i` },
];
const metadata = createFormulaMetadata({
  formulaId,
  title: "中文 Office 公式",
  lines,
  codeFormat: "align-star",
  displayMode: "block",
  numbered: true,
  appVersion: "1.0.6",
});

assert.equal(metadata.latex, lines.map((line) => line.latex).join("\n"));
assert.ok(isVisualTeXFormulaMetadata(metadata));
assert.equal(metadata.numbered, true);

const encoded = encodeFormulaMetadata(metadata);
assert.match(encoded, /^visualtex:v1:deflate:[A-Za-z0-9_-]+$/);
assert.deepEqual(decodeFormulaMetadata(encoded), metadata);
assert.equal(decodeFormulaMetadata(`${encoded.slice(0, -2)}!!`), null);
assert.ok(
  encoded.length < JSON.stringify(metadata).length * 1.3,
  "compressed metadata should not expand materially",
);

const xml = formulaMetadataToXml(metadata);
assert.match(xml, new RegExp(`xmlns="${VISUALTEX_FORMULA_XML_NAMESPACE}"`));
assert.deepEqual(formulaMetadataFromXml(xml), metadata);
assert.equal(
  formulaMetadataFromXml(xml.replace(formulaId, crypto.randomUUID())),
  null,
);

const updated = createFormulaMetadata({
  formulaId,
  title: "Updated",
  lines: [{ id: lines[0].id, latex: "x=y" }],
  codeFormat: "raw",
  displayMode: "inline",
  appVersion: "1.0.6",
  original: metadata,
});
assert.equal(updated.createdAt, metadata.createdAt);
assert.equal(updated.createdWithVersion, metadata.createdWithVersion);
assert.notEqual(updated.updatedAt, "");
assert.equal(updated.formulaId, metadata.formulaId);

assert.throws(() =>
  createFormulaMetadata({
    formulaId: "not-a-uuid",
    title: "Invalid",
    lines,
    codeFormat: "raw",
  }),
);
assert.equal(isVisualTeXFormulaMetadata({ ...metadata, lines: [] }), false);

const shapeName = powerpointShapeName(formulaId);
assert.equal(shapeName, `VisualTeX_${formulaId}`);
assert.equal(formulaIdFromPowerPointShapeName(shapeName), formulaId);
assert.equal(formulaIdFromPowerPointShapeName("Ordinary Picture"), null);

const objectReference = {
  slideId: "slide/id:包含空格",
  shapeId: "shape/id:αβ",
};
const encodedReference = encodePowerPointObjectReference(objectReference);
assert.match(encodedReference, /^visualtex-ppt:v1:[A-Za-z0-9_-]+$/);
assert.deepEqual(
  decodePowerPointObjectReference(encodedReference),
  objectReference,
);
assert.equal(decodePowerPointObjectReference(`${encodedReference}!!`), null);

console.log("Office formula metadata smoke test passed");
