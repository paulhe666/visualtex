import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { formatOcrError } from "../src/ocr/ocrService.ts";

assert.equal(
  formatOcrError(
    { error: { message: "OCR_MODEL_INTEGRITY_ERROR: inference.json is empty" } },
    "Image OCR failed",
  ),
  "OCR_MODEL_INTEGRITY_ERROR: inference.json is empty",
);

const detailed = formatOcrError(
  {
    message: "OCR_MODEL_LOAD_ERROR: Paddle FormulaRecognition initialization failed",
    details: "[json.exception.parse_error.101] attempting to parse an empty input",
  },
  "Image OCR failed",
);
assert.match(detailed, /OCR_MODEL_LOAD_ERROR/);
assert.match(detailed, /json\.exception\.parse_error\.101/);
assert.notEqual(detailed, "Image OCR failed");

assert.equal(
  formatOcrError(
    { reason: { error: "OCR_MODEL_PATH_COMPATIBILITY_ERROR: no ASCII model path" } },
    "Quick OCR failed",
  ),
  "OCR_MODEL_PATH_COMPATIBILITY_ERROR: no ASCII model path",
);

assert.equal(formatOcrError(42, "Quick OCR failed"), "Quick OCR failed");

for (const relativePath of [
  "src/App.tsx",
  "src/components/OcrDialog.tsx",
  "src/office/dialog/OfficeDialogApp.tsx",
  "src/office/api/ocrHttpTransport.ts",
]) {
  const source = readFileSync(resolve(relativePath), "utf8");
  assert.match(
    source,
    /formatOcrError/,
    `${relativePath} must preserve structured OCR error details`,
  );
}

const officeTransport = readFileSync(
  resolve("src/office/api/ocrHttpTransport.ts"),
  "utf8",
);
assert.match(officeTransport, /JSON\.parse\(detail\)/);
assert.match(officeTransport, /throw new Error\(detail\.trim\(\)\)/);

console.log("VisualTeX OCR structured error-message regression passed");
