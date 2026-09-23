import assert from "node:assert/strict";
import {
  createFormulaMetadata,
  decodeFormulaMetadata,
  encodeFormulaMetadata,
  isVisualTeXFormulaMetadata,
} from "../src/office/shared/formulaMetadata.ts";

const input = {
  formulaId: "11111111-1111-4111-8111-111111111111",
  title: "OMML identity compatibility",
  lines: [{ id: "22222222-2222-4222-8222-222222222222", latex: "x" }],
  codeFormat: "raw",
  displayMode: "inline" as const,
};
const base = createFormulaMetadata(input);
let checked = 0;
assert.ok(isVisualTeXFormulaMetadata(base));
checked++;
for (const fingerprint of ["a".repeat(64), "A".repeat(64), "omml2:" + "a".repeat(64), "OMML2:" + "A".repeat(64)]) {
  const metadata = { ...base, nativeOmmlFingerprint: fingerprint };
  assert.ok(isVisualTeXFormulaMetadata(metadata), "supported fingerprint rejected");
  assert.equal(decodeFormulaMetadata(encodeFormulaMetadata(metadata))?.nativeOmmlFingerprint, fingerprint,
    "metadata transport changed fingerprint or dropped its version");
  const reopened = createFormulaMetadata({ ...input, original: metadata });
  assert.equal(reopened.nativeOmmlFingerprint, fingerprint, "editor snapshot lost the captured identity");
  assert.equal(metadata.nativeOmmlFingerprint, fingerprint, "creating the snapshot mutated its source");
  checked += 4;
}
for (const fingerprint of ["", "x".repeat(64), "a".repeat(63), "a".repeat(65),
  "omml2:" + "a".repeat(63), "omml3:" + "a".repeat(64), "v2:" + "a".repeat(64),
  " " + "a".repeat(64), "omml2:" + "a".repeat(64) + " "]) {
  const metadata = { ...base, nativeOmmlFingerprint: fingerprint };
  assert.equal(isVisualTeXFormulaMetadata(metadata), false, "unknown/malformed fingerprint accepted");
  assert.throws(() => encodeFormulaMetadata(metadata), "invalid metadata reached the Office transport");
  checked += 2;
}
console.log(JSON.stringify({ suite: "omml_fingerprint_metadata_regression", checked, passed: true }));
