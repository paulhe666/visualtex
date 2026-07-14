import assert from "node:assert/strict";

let currentSessionId = crypto.randomUUID();
let currentFormulaId = crypto.randomUUID();
let currentLineId = crypto.randomUUID();
let dialogMessageHandler = null;
let dialogClosedHandler = null;
let dialogCloseCount = 0;
let applyCount = 0;
let nativeCommitCount = 0;
let nativeConfirmCount = 0;
let nativeFinalizeCount = 0;
const nativeWordBaselinePositions = [];
const nativeWordBaselineMarkers = [];
const appliedLatex = [];
const statusMessages = [];
const preparedInteractionTargets = [];

function createSession() {
  return {
    id: currentSessionId,
    mode: "create",
    host: "powerpoint",
    formulaId: currentFormulaId,
    sourceDocumentId: "presentation-1",
    sourceObjectId: null,
    title: "PowerPoint Formula",
    lines: [{ id: currentLineId, latex: "" }],
    activeLineId: currentLineId,
    codeFormat: "raw",
    displayMode: "block",
    exportWidth: 0,
    exportHeight: 0,
    exportResult: null,
    originalMetadata: null,
    dirty: false,
    status: "created",
    autoCommitOnClose: true,
    explicitCancel: false,
    error: null,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    expiresAt: Date.now() + 60_000,
  };
}

function setFormulaDraft(status, latex, includePng = true) {
  Object.assign(session, {
    lines: [{ id: currentLineId, latex }],
    activeLineId: currentLineId,
    dirty: true,
    status,
    exportWidth: 80,
    exportHeight: 32,
    exportResult: {
      svg: "<svg></svg>",
      svgBase64: "svg-base64",
      pngBase64: includePng ? "png-base64" : undefined,
      width: 80,
      height: 32,
      baseline: 24,
    },
    updatedAt: Date.now(),
  });
}

let session = createSession();

const fakeDialog = {
  addEventHandler(type, handler) {
    if (type === "dialogMessageReceived") dialogMessageHandler = handler;
    if (type === "dialogEventReceived") dialogClosedHandler = handler;
  },
  close() {
    dialogCloseCount += 1;
    dialogClosedHandler?.({ error: 12006 });
  },
};

globalThis.window = {
  __VISUALTEX_INSTALL_TOKEN__: "test-token",
  location: { origin: "https://127.0.0.1:43127" },
  setTimeout,
  clearTimeout,
  setInterval,
  clearInterval,
  alert() {},
};
globalThis.document = {
  querySelector() {
    return null;
  },
};

globalThis.Office = {
  AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
  EventType: {
    DialogMessageReceived: "dialogMessageReceived",
    DialogEventReceived: "dialogEventReceived",
  },
  context: {
    ui: {
      displayDialogAsync(_url, _options, callback) {
        callback({ status: "succeeded", value: fakeDialog });
      },
    },
  },
};

globalThis.fetch = async (input, init = {}) => {
  const url = String(input);
  if (url === "/health") {
    return new Response(
      JSON.stringify({
        ok: true,
        appVersion: "1.0.18",
        officeUiVersion: "1.0.18",
        protocolVersion: 1,
        ocrAvailable: true,
      }),
      { status: 200, headers: { "Content-Type": "application/json" } },
    );
  }

  if (url === "/api/v1/sessions" && init.method === "POST") {
    const inputSession = JSON.parse(init.body);
    Object.assign(session, inputSession, {
      id: currentSessionId,
      formulaId: currentFormulaId,
      status: "created",
      createdAt: session.createdAt,
      updatedAt: Date.now(),
    });
    return new Response(JSON.stringify(session), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  }

  if (url === `/api/v1/sessions/${currentSessionId}`) {
    if (init.method === "PATCH") {
      Object.assign(session, JSON.parse(init.body), { updatedAt: Date.now() });
    }
    return new Response(JSON.stringify(session), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  }

  if (
    url === `/api/v1/powerpoint/sessions/${currentSessionId}/commit` &&
    init.method === "POST"
  ) {
    nativeCommitCount += 1;
    return new Response(
      JSON.stringify({
        session,
        selection: {
          shapeName: `VisualTeX_${currentFormulaId}`,
          slideIndex: 1,
          slideId: 256,
          presentationIdentity: "Deck.pptx",
          left: 10,
          top: 20,
          width: 80,
          height: 32,
        },
      }),
      {
        status: 200,
        headers: { "Content-Type": "application/json" },
      },
    );
  }

  if (
    url === `/api/v1/powerpoint/sessions/${currentSessionId}/confirm` &&
    init.method === "POST"
  ) {
    nativeConfirmCount += 1;
    return new Response(JSON.stringify(session), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  }

  if (url === "/api/v1/word/inline-baseline" && init.method === "POST") {
    const { position, formulaMarker } = JSON.parse(init.body);
    nativeWordBaselinePositions.push(position);
    nativeWordBaselineMarkers.push(formulaMarker);
    return new Response(
      JSON.stringify({
        appliedPosition: position,
        width: 80,
        height: 32,
        matchedShapeIndex: 1,
      }),
      {
        status: 200,
        headers: { "Content-Type": "application/json" },
      },
    );
  }

  throw new Error(`Unexpected request: ${url}`);
};

const adapter = {
  host: "powerpoint",
  requiredExportFormat: "png",
  prepareInteractionTarget(target) {
    preparedInteractionTargets.push(target);
  },
  async readSelection() {
    return {
      sourceDocumentId: "presentation-1",
      sourceObjectId: null,
      sessionSeed: { displayMode: "block" },
    };
  },
  async applySession(value) {
    applyCount += 1;
    assert.equal(value.id, currentSessionId);
    assert.ok(value.lines[0].latex.trim());
    assert.ok(value.exportResult?.pngBase64);
    appliedLatex.push(value.lines[0].latex);
  },
  async openDesktopApp() {},
  showMessage(message) {
    statusMessages.push(message);
  },
};

const { OfficeBridge } = await import("../src/office/bridge/OfficeBridge.ts");
const { MacOfficeBridge } = await import("../src/office/macos/MacOfficeBridge.ts");

// Scenario 1: PowerPoint loses DialogMessageReceived, but the editor has already
// persisted status=committing. The bridge watcher must still insert the formula.
let commandCompletedCount = 0;
const bridge = new OfficeBridge(adapter);
await bridge.run("create", () => {
  commandCompletedCount += 1;
});

assert.equal(
  commandCompletedCount,
  1,
  "PowerPoint ribbon command must complete immediately after the editor opens",
);
assert.ok(dialogMessageHandler, "dialog message handler must be registered");
assert.ok(dialogClosedHandler, "dialog close handler must be registered");

setFormulaDraft("committing", String.raw`\alpha+x`);
await new Promise((resolve) => setTimeout(resolve, 500));

assert.equal(applyCount, 1, "session polling must apply a committing formula exactly once");
assert.equal(session.status, "completed");
assert.equal(dialogCloseCount, 1, "successful insertion must close the editor dialog");
assert.equal(
  commandCompletedCount,
  1,
  "formula completion must not complete the already released command twice",
);
assert.ok(statusMessages.includes("VisualTeX 公式已插入。"));

await new Promise((resolve) => setTimeout(resolve, 250));
assert.equal(applyCount, 1, "the stopped watcher must not apply the same session twice");
assert.equal(commandCompletedCount, 1, "command completion must be idempotent");

// Scenario 2: the user closes the editor directly. The bridge must remain alive,
// receive DialogEventReceived, and auto-commit the last persisted draft.
currentSessionId = crypto.randomUUID();
currentFormulaId = crypto.randomUUID();
currentLineId = crypto.randomUUID();
session = createSession();
let closeCompletedCount = 0;
const closeBridge = new OfficeBridge(adapter);
await closeBridge.run("create", () => {
  closeCompletedCount += 1;
});
assert.equal(
  closeCompletedCount,
  1,
  "manual-close Session must not keep the ribbon command pending",
);

setFormulaDraft("editing", String.raw`\sum_{i=1}^{n}x_i`, false);
dialogClosedHandler?.({ error: 12006 });
await new Promise((resolve) => setTimeout(resolve, 250));
assert.equal(
  applyCount,
  1,
  "closing must not submit the SVG-only draft to a PNG-backed Windows adapter",
);
session.exportResult.pngBase64 = "png-base64-after-close";
session.updatedAt = Date.now();
await new Promise((resolve) => setTimeout(resolve, 450));

assert.equal(
  applyCount,
  2,
  "closing the dialog must wait for and auto-commit the persisted PNG draft",
);
assert.equal(session.status, "completed");
assert.equal(
  closeCompletedCount,
  1,
  "manual close must not complete the same Office command twice",
);
assert.deepEqual(appliedLatex, [String.raw`\alpha+x`, String.raw`\sum_{i=1}^{n}x_i`]);

// Scenario 3: direct close must also apply a dirty edit Session, not only a
// newly-created formula.
currentSessionId = crypto.randomUUID();
currentFormulaId = crypto.randomUUID();
currentLineId = crypto.randomUUID();
session = createSession();
const editBridge = new OfficeBridge(adapter);
await editBridge.run("edit", () => undefined);
setFormulaDraft("editing", String.raw`\int_0^1 x^2\,dx`);
dialogClosedHandler?.({ error: 12006 });
await new Promise((resolve) => setTimeout(resolve, 350));
assert.equal(applyCount, 3, "closing a dirty edit dialog must update the formula");
assert.equal(session.status, "completed");

// Scenario 4: even if an Office dialog falls back to its parent bridge, a
// native macOS PowerPoint Session must still use the native commit endpoint.
currentSessionId = crypto.randomUUID();
currentFormulaId = crypto.randomUUID();
currentLineId = crypto.randomUUID();
session = createSession();
const nativeAdapter = {
  ...adapter,
  async readSelection() {
    return {
      sourceDocumentId: "visualtex-ppt-native-presentation:Deck.pptx",
      sourceObjectId: "visualtex-ppt-native-slide:256:1",
      sessionSeed: { displayMode: "block" },
    };
  },
  async finalizeNativePowerPointCommit(value, selection) {
    nativeFinalizeCount += 1;
    assert.equal(value.id, currentSessionId);
    assert.equal(selection.shapeName, `VisualTeX_${currentFormulaId}`);
  },
};
const nativeBridge = new MacOfficeBridge(nativeAdapter);
const nativeDialogCloseBefore = dialogCloseCount;
await nativeBridge.run("create", () => undefined);
setFormulaDraft("committing", String.raw`\beta+y`);
await new Promise((resolve) => setTimeout(resolve, 500));
assert.equal(nativeCommitCount, 1, "native PowerPoint Session must prepare natively");
assert.equal(
  nativeFinalizeCount,
  1,
  "Office.js must persist and verify the prepared PowerPoint shape metadata",
);
assert.equal(nativeConfirmCount, 1, "prepared PowerPoint commit must be confirmed once");
assert.equal(applyCount, 3, "native commit must not call the legacy adapter insertion path");
assert.equal(
  dialogCloseCount,
  nativeDialogCloseBefore + 1,
  "successful PowerPoint insertion must automatically close the editor window",
);

// Scenario 5: the macOS Word ribbon command delegates to the Word adapter and
// reports the number of refreshed equation labels.
const wordNumberingAdapter = {
  ...adapter,
  host: "word",
  async updateEquationNumbers() {
    return 3;
  },
};
const wordBridge = new MacOfficeBridge(wordNumberingAdapter);
assert.equal(await wordBridge.updateEquationNumbers(), 3);
assert.ok(statusMessages.includes("VisualTeX 已刷新 3 个公式编号。"));

// Scenario 6: Mac Word commits reapply the same mathematical descent through
// the native Font.Position API after the Office.js picture becomes durable.
currentSessionId = crypto.randomUUID();
currentFormulaId = crypto.randomUUID();
currentLineId = crypto.randomUUID();
session = createSession();
let wordApplyCount = 0;
const nativeWordAdapter = {
  ...adapter,
  host: "word",
  async readSelection() {
    return {
      sourceDocumentId: "document-1",
      sourceObjectId: null,
      sessionSeed: { displayMode: "inline" },
    };
  },
  async applySession() {
    wordApplyCount += 1;
  },
  getNativeWordFormulaMarker(sessionId) {
    assert.equal(sessionId, currentSessionId);
    return "visualtex:v1:deflate:test-marker";
  },
};
const nativeWordBridge = new MacOfficeBridge(nativeWordAdapter);
await nativeWordBridge.run("create", () => undefined);
assert.equal(session.host, "word");
setFormulaDraft("committing", String.raw`\frac{x}{y}`);
await new Promise((resolve) => setTimeout(resolve, 650));
assert.equal(wordApplyCount, 1);
assert.equal(
  nativeWordBaselinePositions.at(-1),
  -6,
  "native Word fallback must lower by h * (H - B) / H",
);
assert.equal(
  nativeWordBaselineMarkers.at(-1),
  "visualtex:v1:deflate:test-marker",
  "native Word fallback must target the durable formula marker, not the caret",
);

// Scenario 7: a double-click target must be prepared before selection is read,
// so PowerPoint's native format UI cannot steal the mutable selection.
currentSessionId = crypto.randomUUID();
currentFormulaId = crypto.randomUUID();
currentLineId = crypto.randomUUID();
session = createSession();
const interactionTarget = {
  host: "powerpoint",
  formulaId: currentFormulaId,
  shapeName: `VisualTeX_${currentFormulaId}`,
  slideIndex: 1,
  left: 10,
  top: 20,
  width: 80,
  height: 32,
};
const targetedBridge = new OfficeBridge(adapter);
await targetedBridge.run("edit", () => undefined, interactionTarget);
assert.deepEqual(preparedInteractionTargets.at(-1), interactionTarget);
session.status = "cancelled";
await new Promise((resolve) => setTimeout(resolve, 250));

console.log("Office command lifecycle, native commit and Word numbering command passed");
