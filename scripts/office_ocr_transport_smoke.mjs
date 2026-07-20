import assert from "node:assert/strict";

const token = "a".repeat(64);
globalThis.window = {
  __VISUALTEX_INSTALL_TOKEN__: token,
  setTimeout,
  clearTimeout,
};
globalThis.document = {
  querySelector(selector) {
    return selector === 'meta[name="visualtex-install-token"]'
      ? { content: token }
      : null;
  },
};

const calls = [];
let eventPollCount = 0;
globalThis.fetch = async (input, init = {}) => {
  const url = String(input);
  calls.push({ url, init });

  if (url.startsWith("/api/v1/ocr/status")) {
    return new Response(
      JSON.stringify({
        installed: true,
        pythonPath: "/opt/visualtex/python",
        pythonVersion: "3.12.0",
        paddleVersion: "3.3.1",
        paddleocrVersion: "3.7.0",
        runtimePath: "/opt/visualtex/ocr",
        message: "ready",
      }),
      { status: 200, headers: { "Content-Type": "application/json" } },
    );
  }

  if (url === "/api/v1/ocr/warmup") {
    return new Response(null, { status: 204 });
  }

  if (url === "/api/v1/ocr/recognize") {
    return new Response(
      JSON.stringify({
        model: "PP-FormulaNet_plus-M",
        elapsedMs: 12,
        processedWidth: 100,
        processedHeight: 50,
        backgroundInverted: false,
        backgroundLuminance: 255,
        formulas: [{ latex: "x^2" }],
      }),
      { status: 200, headers: { "Content-Type": "application/json" } },
    );
  }

  if (url.startsWith("/api/v1/ocr/events")) {
    eventPollCount += 1;
    if (!url.includes("cursor=")) {
      return new Response(JSON.stringify({ cursor: 7, events: [] }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    }
    return new Response(
      JSON.stringify({
        cursor: 8,
        events:
          eventPollCount === 2
            ? [
                {
                  id: 8,
                  event: "ocr-recognition-progress",
                  payload: {
                    event: "progress",
                    id: "request-1",
                    stage: "model",
                    message: "loading",
                    model: "PP-FormulaNet_plus-M",
                  },
                },
              ]
            : [],
      }),
      { status: 200, headers: { "Content-Type": "application/json" } },
    );
  }

  throw new Error(`Unexpected fetch: ${url}`);
};

const { invoke, listen } = await import(
  "../src/office/api/ocrHttpTransport.ts"
);

const status = await invoke("get_ocr_runtime_status", {
  forceRefresh: true,
});
assert.equal(status.installed, true);
assert.equal(calls[0].url, "/api/v1/ocr/status?forceRefresh=true");
assert.equal(
  calls[0].init.headers["X-VisualTeX-Install-Token"],
  token,
);

const result = await invoke("recognize_formula_image", {
  request: {
    bytes: [1, 2, 3],
    extension: "jpg",
    model: "PP-FormulaNet_plus-M",
  },
});
assert.equal(result.formulas[0].latex, "x^2");
const recognitionCall = calls.find(
  (call) => call.url === "/api/v1/ocr/recognize",
);
assert.ok(recognitionCall);
assert.equal(
  recognitionCall.init.headers["Content-Type"],
  "application/octet-stream",
);
assert.equal(
  recognitionCall.init.headers["X-VisualTeX-Ocr-Model"],
  "PP-FormulaNet_plus-M",
);
assert.equal(
  recognitionCall.init.headers["X-VisualTeX-Ocr-Extension"],
  "jpg",
);
assert.deepEqual(Array.from(recognitionCall.init.body), [1, 2, 3]);

await invoke("warmup_ocr_model", {
  model: "PP-FormulaNet_plus-S",
});
const warmupCall = calls.find((call) => call.url === "/api/v1/ocr/warmup");
assert.ok(warmupCall);
assert.equal(
  warmupCall.init.headers["X-VisualTeX-Ocr-Model"],
  "PP-FormulaNet_plus-S",
);

const received = [];
const unlisten = await listen("ocr-recognition-progress", (event) => {
  received.push(event);
});
await new Promise((resolve) => setTimeout(resolve, 240));
unlisten();
assert.equal(received.length, 1);
assert.equal(received[0].id, 8);
assert.equal(received[0].payload.stage, "model");
assert.ok(
  calls.some(
    (call) =>
      call.url.includes("/api/v1/ocr/events?cursor=7") &&
      call.url.includes("event=ocr-recognition-progress"),
  ),
);

console.log("Office OCR HTTP transport smoke test passed");
