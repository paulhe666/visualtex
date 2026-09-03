import assert from "node:assert/strict";
import { isQuickOcrHudPayload } from "../src/ocr/QuickOcrHud.tsx";

assert.equal(
  isQuickOcrHudPayload({
    status: "running",
    message: "Recognizing…",
    progress: 42,
  }),
  true,
);
assert.equal(isQuickOcrHudPayload(null), false);
for (const progress of [Number.NaN, -1, 1.5, 101]) {
  assert.equal(
    isQuickOcrHudPayload({ status: "running", message: "x", progress }),
    false,
  );
}
assert.equal(
  isQuickOcrHudPayload({ status: "unknown", message: "x", progress: 10 }),
  false,
);
assert.equal(
  isQuickOcrHudPayload({ status: "success", message: null, progress: 100 }),
  false,
);

console.log("VisualTeX Quick OCR HUD payload safety regression passed");
