import assert from "node:assert/strict";
import { isSilentOcrHudPayload } from "../src/desktop/silentOcrHudPayload.ts";

assert.equal(
  isSilentOcrHudPayload({
    status: "running",
    message: "Recognizing…",
    progress: 42,
  }),
  true,
);
assert.equal(isSilentOcrHudPayload(null), false);
assert.equal(
  isSilentOcrHudPayload({ status: "success", message: "done", progress: 100 }),
  true,
);
assert.equal(
  isSilentOcrHudPayload({ status: "unknown", message: "x", progress: 20 }),
  false,
);
assert.equal(
  isSilentOcrHudPayload({ status: "error", message: 7, progress: 20 }),
  false,
);
assert.equal(
  isSilentOcrHudPayload({ status: "error", message: "x", progress: Number.NaN }),
  false,
);

console.log("VisualTeX silent OCR HUD payload safety regression passed");
