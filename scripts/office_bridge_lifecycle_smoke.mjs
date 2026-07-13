import assert from "node:assert/strict";

let currentSessionId = crypto.randomUUID();
let currentFormulaId = crypto.randomUUID();
let currentLineId = crypto.randomUUID();
let dialogMessageHandler = null;
let dialogClosedHandler = null;
let dialogCloseCount = 0;
let applyCount = 0;
const appliedLatex = [];
const statusMessages = [];

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

function setFormulaDraft(status, latex) {
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
      pngBase64: "png-base64",
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
        appVersion: "1.0.16",
        officeUiVersion: "1.0.16",
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

  throw new Error(`Unexpected request: ${url}`);
};

const adapter = {
  host: "powerpoint",
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

// Scenario 1: PowerPoint loses DialogMessageReceived, but the editor has already
// persisted status=committing. The bridge watcher must still insert the formula.
let commandCompletedCount = 0;
const bridge = new OfficeBridge(adapter);
await bridge.run("create", () => {
  commandCompletedCount += 1;
});

assert.equal(commandCompletedCount, 0, "PowerPoint command must stay alive while the dialog is open");
assert.ok(dialogMessageHandler, "dialog message handler must be registered");
assert.ok(dialogClosedHandler, "dialog close handler must be registered");

setFormulaDraft("committing", String.raw`\alpha+x`);
await new Promise((resolve) => setTimeout(resolve, 500));

assert.equal(applyCount, 1, "session polling must apply a committing formula exactly once");
assert.equal(session.status, "completed");
assert.equal(dialogCloseCount, 1, "successful insertion must close the editor dialog");
assert.equal(commandCompletedCount, 1, "event.completed() must run only after insertion finishes");
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
assert.equal(closeCompletedCount, 0, "manual-close command must remain alive before the dialog closes");

setFormulaDraft("editing", String.raw`\sum_{i=1}^{n}x_i`);
dialogClosedHandler?.({ error: 12006 });
await new Promise((resolve) => setTimeout(resolve, 350));

assert.equal(applyCount, 2, "closing the dialog must auto-commit the persisted PowerPoint formula");
assert.equal(session.status, "completed");
assert.equal(closeCompletedCount, 1, "manual close must complete the Office command after auto-commit");
assert.deepEqual(appliedLatex, [String.raw`\alpha+x`, String.raw`\sum_{i=1}^{n}x_i`]);

console.log("PowerPoint Office command lifecycle, commit polling and close auto-commit passed");
