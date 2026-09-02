import assert from "node:assert/strict";
import {
  decodeOcrInstallProgress,
  decodeOcrRecognitionProgress,
  decodeOcrRecognitionResult,
  decodeOcrRuntimeStatus,
} from "../src/ocr/ocrPayloadValidation.ts";

const runtime = {
  installed: true,
  pythonPath: "/tmp/python",
  pythonVersion: "3.10",
  paddleVersion: "3.3.1",
  paddleocrVersion: "3.7.0",
  runtimePath: "/tmp/runtime",
  offlineBundleAvailable: true,
  installedModels: ["PP-FormulaNet_plus-M"],
  defaultModel: "PP-FormulaNet_plus-M",
  message: "ready",
};
assert.equal(decodeOcrRuntimeStatus(runtime), runtime);
assert.throws(
  () => decodeOcrRuntimeStatus({ ...runtime, installedModels: [null] }),
  /installedModels\[0\]/,
);

const recognition = {
  model: "PP-FormulaNet_plus-M",
  elapsedMs: 123,
  processedWidth: 640,
  processedHeight: 320,
  backgroundInverted: false,
  backgroundLuminance: 0.95,
  formulas: [{ latex: "x^2" }],
};
assert.equal(decodeOcrRecognitionResult(recognition), recognition);
assert.throws(
  () => decodeOcrRecognitionResult({ ...recognition, formulas: [{ latex: 1 }] }),
  /formulas\[0\]\.latex/,
);

const progress = {
  event: "progress",
  id: "run-1",
  stage: "inference",
  message: "running",
  model: "PP-FormulaNet_plus-M",
};
assert.equal(decodeOcrRecognitionProgress(progress), progress);
assert.throws(
  () => decodeOcrRecognitionProgress({ ...progress, model: "unknown" }),
  /\.model/,
);

const install = {
  stage: "download",
  percent: 50,
  message: "halfway",
  detail: null,
};
assert.equal(decodeOcrInstallProgress(install), install);
assert.throws(
  () => decodeOcrInstallProgress({ ...install, percent: Number.NaN }),
  /\.percent/,
);

console.log("VisualTeX macOS OCR payload safety regression passed");
