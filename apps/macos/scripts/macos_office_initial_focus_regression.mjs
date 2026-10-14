import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import ts from "typescript";

// Exercise the actual activation effect, including StrictMode's setup/cleanup
// replay, without opening Word or touching the user's Office Session.
const source = readFileSync(new URL("../src/office/dialog/OfficeDialogApp.tsx", import.meta.url), "utf8");
const file = ts.createSourceFile("OfficeDialogApp.tsx", source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
const effects = [];
function visit(node) {
  if (ts.isCallExpression(node) && node.expression.getText(file) === "useEffect") {
    const callback = node.arguments[0];
    if (callback && callback.getText(file).includes("const focusFirstLine =")) effects.push(callback.getText(file));
  }
  ts.forEachChild(node, visit);
}
visit(file);
assert.equal(effects.length, 1, "Expected exactly one initial focus controller");
const effectCode = ts.transpileModule(`globalThis.effect = ${effects[0]};`, {
  compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.None },
}).outputText;
const scenarios = [];
function fixture(overrides = {}) {
  let nextId = 0;
  const frames = new Map();
  const intervals = new Map();
  const timers = new Map();
  const windowListeners = new Map();
  const documentListeners = new Map();
  const calls = [];
  const field = {
    tagName: "MATH-FIELD",
    focus() { if (state.mountReady) document.activeElement = field; },
    classList: { contains: () => false },
  };
  const document = {
    activeElement: { tagName: "BODY", classList: { contains: () => false } },
    visibilityState: "visible",
    querySelector: () => state.mountReady ? field : null,
    addEventListener: (name, fn) => documentListeners.set(name, fn),
    removeEventListener: (name, fn) => {
      if (documentListeners.get(name) === fn) documentListeners.delete(name);
    },
  };
  const window = {
    focus() {},
    requestAnimationFrame: (fn) => { const id = ++nextId; frames.set(id, fn); return id; },
    cancelAnimationFrame: (id) => frames.delete(id),
    setInterval: (fn, ms) => { assert.equal(ms, 60); const id = ++nextId; intervals.set(id, fn); return id; },
    clearInterval: (id) => intervals.delete(id),
    setTimeout: (fn, ms) => { assert.equal(ms, 3000); const id = ++nextId; timers.set(id, fn); return id; },
    clearTimeout: (id) => timers.delete(id),
    addEventListener: (name, fn) => windowListeners.set(name, fn),
    removeEventListener: (name, fn) => {
      if (windowListeners.get(name) === fn) windowListeners.delete(name);
    },
  };
  const state = {
    sessionKey: "7:test-session",
    presentedSessionKey: "7:test-session",
    sessionHydrated: true,
    tauriResidentEditor: true,
    mountReady: true,
    initialEditorFocusSessionRef: { current: "" },
    activeSessionKeyRef: { current: "7:test-session" },
    editorRef: {
      current: { focus(options) {
        calls.push(JSON.parse(JSON.stringify(options)));
        if (state.mountReady) document.activeElement = field;
      } },
    },
    window, document,
    ...overrides,
  };
  const context = vm.createContext(state);
  new vm.Script(effectCode, { filename: fileURLToPath(new URL("../src/office/dialog/OfficeDialogApp.tsx", import.meta.url)) })
    .runInContext(context);
  return {
    state, calls, frames, intervals, timers, windowListeners, documentListeners,
    setup: () => context.effect(),
    frame() { const callbacks = [...frames.values()]; frames.clear(); callbacks.forEach((fn) => fn()); },
    tick() { [...intervals.values()].forEach((fn) => fn()); },
    deadline() { const callbacks = [...timers.values()]; timers.clear(); callbacks.forEach((fn) => fn()); },
  };
}
function run(name, fn) { fn(); scenarios.push(name); }
run("no focus before AppKit presentation", () => {
  const f = fixture({ presentedSessionKey: "" });
  assert.equal(f.setup(), undefined);
  assert.equal(f.frames.size, 0);
  assert.equal(f.calls.length, 0);
});
run("no focus before session hydration", () => {
  const f = fixture({ sessionHydrated: false });
  assert.equal(f.setup(), undefined);
  assert.equal(f.calls.length, 0);
});
run("StrictMode cleanup does not suppress the second activation", () => {
  const f = fixture();
  f.setup()();
  assert.equal(f.frames.size, 0);
  const cleanup = f.setup();
  assert.equal(typeof cleanup, "function");
  f.frame();
  assert.deepEqual(f.calls, [{ target: "first", moveToEnd: true }]);
  assert.equal(f.state.initialEditorFocusSessionRef.current, f.state.sessionKey);
  assert.equal(f.intervals.size, 0);
  assert.equal(f.windowListeners.size + f.documentListeners.size, 0);
  cleanup();
  assert.equal(f.setup(), undefined, "Successful activation must not repeat on a settings render");
});
run("late MathLive input mounts within the bounded repair window", () => {
  const f = fixture({ mountReady: false });
  const cleanup = f.setup();
  f.frame();
  assert.equal(f.state.initialEditorFocusSessionRef.current, "");
  assert.equal(f.intervals.size, 1);
  f.state.mountReady = true;
  f.tick();
  assert.equal(f.state.initialEditorFocusSessionRef.current, f.state.sessionKey);
  assert.equal(f.intervals.size, 0);
  cleanup();
  assert.equal(f.timers.size, 0);
});
run("failed activation stops after three seconds", () => {
  const f = fixture({ mountReady: false });
  const cleanup = f.setup();
  f.frame();
  f.deadline();
  assert.equal(f.intervals.size + f.windowListeners.size + f.documentListeners.size, 0);
  const calls = f.calls.length;
  f.tick();
  assert.equal(f.calls.length, calls);
  cleanup();
});
run("a stale generation cannot focus the next Office Session", () => {
  const f = fixture({ mountReady: false });
  const cleanup = f.setup();
  f.frame();
  f.state.activeSessionKeyRef.current = "8:next-session";
  f.state.mountReady = true;
  const calls = f.calls.length;
  f.tick();
  assert.equal(f.calls.length, calls);
  assert.equal(f.intervals.size, 0);
  assert.equal(f.state.initialEditorFocusSessionRef.current, "");
  cleanup();
});
run("non-resident Office view does not require the AppKit handshake", () => {
  const f = fixture({ tauriResidentEditor: false, presentedSessionKey: "" });
  const cleanup = f.setup();
  f.frame();
  assert.equal(f.calls.length, 1);
  cleanup();
});
run("disposed focus callbacks are harmless", () => {
  const f = fixture();
  const cleanup = f.setup();
  const pending = f.windowListeners.get("focus");
  cleanup();
  pending();
  assert.equal(f.calls.length, 0);
});
console.log(JSON.stringify({ passed: scenarios.length, scenarios }, null, 2));
