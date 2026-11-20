import assert from "node:assert/strict";
import { withAppBrowser } from "../../shared/visualtex-app/testing/browser_test_runtime.mjs";

const sessionId = "11111111-2222-4333-8444-555555555555";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

await withAppBrowser(async ({ baseUrl, send, evaluate }) => {
  await send("Page.addScriptToEvaluateOnNewDocument", { source: `(() => {
    const operation = new URLSearchParams(location.search).get("operation");
    // Model a hidden WebKit view: tasks run, but animation frames never arrive.
    let frameId = 0;
    window.requestAnimationFrame = () => ++frameId;
    window.cancelAnimationFrame = () => {};
    window.__conversionProbe = { commits: 0, closes: 0, exportReady: false };
    window.close = () => { window.__conversionProbe.closes += 1; };
    let session = {
      id: ${JSON.stringify(sessionId)}, mode: "edit", host: "word", operation,
      formulaId: "22222222-3333-4444-8555-666666666666", sourceDocumentId: "probe",
      sourceObjectId: "probe", title: "Hidden conversion", lines: [{ id: "line-1", latex: "x^2+y^2=z^2" }],
      activeLineId: "line-1", codeFormat: "raw", displayMode: "block", numbered: true,
      fontSizePt: 11, exportWidth: 0, exportHeight: 0, exportResult: null,
      originalMetadata: null, dirty: false, status: "created", autoCommitOnClose: true,
      explicitCancel: false, error: null, createdAt: Date.now(), updatedAt: Date.now(),
      expiresAt: Date.now() + 600000
    };
    const originalFetch = window.fetch.bind(window);
    window.fetch = async (input, init = {}) => {
      const url = new URL(typeof input === "string" ? input : input.url, location.href);
      if (url.pathname !== "/api/v1/sessions/${sessionId}") return originalFetch(input, init);
      if (String(init.method || "GET").toUpperCase() === "PATCH") {
        const patch = JSON.parse(String(init.body || "{}"));
        session = { ...session, ...patch };
        if (patch.status === "committing") {
          window.__conversionProbe.commits += 1;
          window.__conversionProbe.exportReady = Boolean(patch.exportResult?.svgBase64 && patch.exportResult?.ommlBase64);
          session.status = "completed";
        }
      }
      return new Response(JSON.stringify(session), { status: 200, headers: { "Content-Type": "application/json" } });
    };
  })();` });
  for (const operation of ["nativeToImage", "imageToNative"]) {
    await send("Page.navigate", { url: `${baseUrl}/office-native-dialog.html?sessionId=${sessionId}&officeHost=word&operation=${operation}` });
    let result;
    const deadline = Date.now() + 15000;
    while (Date.now() < deadline) {
      result = await evaluate("window.__conversionProbe");
      if (result?.closes) break;
      await sleep(100);
    }
    assert.equal(result?.commits, 1, `${operation} must commit exactly once without animation frames`);
    assert.equal(result?.closes, 1, `${operation} must finish without a visible editor`);
    assert.equal(result?.exportReady, true, `${operation} must submit complete image and OMML artifacts`);
    console.log(`${operation}: PASS with animation frames suspended`);
  }
});
