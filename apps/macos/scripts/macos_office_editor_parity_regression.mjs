import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";
import ts from "typescript";
import {
  DEFAULT_LATEX_FORMAT_PROFILE,
  normalizeLatexFormatProfile,
} from "../src/clipboard/latexFormatProfile.ts";

// Executes the actual synchronization function and activation effect extracted
// from the TSX source, in a process-local sandbox. No Office process, native
// command, persistent browser profile, or project file is touched by this test.
const filename = new URL("../src/office/dialog/OfficeDialogApp.tsx", import.meta.url);
const source = readFileSync(filename, "utf8");
const ast = ts.createSourceFile("OfficeDialogApp.tsx", source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
function collect(predicate) {
  const found = [];
  const visit = (node) => {
    if (predicate(node)) found.push(node);
    ts.forEachChild(node, visit);
  };
  visit(ast);
  return found;
}
function sandbox(code, globals) {
  const javascript = ts.transpileModule(code, {
    compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.None },
  }).outputText;
  const context = vm.createContext(globals);
  vm.runInContext(javascript, context, { timeout: 1000 });
  return context;
}
const functions = collect((node) => ts.isFunctionDeclaration(node) && node.name?.text === "syncOfficeEditorSystemSettings");
assert.equal(functions.length, 1, "Expected the real Office settings synchronizer");
const state = {
  latexFormatProfile: { ...DEFAULT_LATEX_FORMAT_PROFILE },
  inputBehavior: {
    autoEscapeShortcuts: true,
    autoExitSuperscript: true,
    autoExitSubscript: true,
    autoExitAccent: true,
    autoExitWrapperCommand: true,
    showStructuredCommandSuggestions: true,
    showOtherCommandSuggestions: true,
  },
  title: "untouched Office session",
  lines: [{ id: "unchanged", latex: "x+1" }],
  activeLineId: "unchanged",
};
const calls = [];
state.setLatexFormatProfile = (profile) => {
  calls.push("profile");
  state.latexFormatProfile = normalizeLatexFormatProfile(profile);
};
state.setInputBehavior = (key, value) => {
  calls.push(`input:${key}`);
  state.inputBehavior[key] = value;
};
// These setters intentionally throw: profile synchronization must never replace
// the loaded Office document or its legacy export-format metadata.
state.setLatexCodeFormat = state.replaceDocumentState = () => {
  throw new Error("Settings synchronization changed the Office session document");
};
const sync = sandbox(`${functions[0].getText(ast)}\nglobalThis.run = syncOfficeEditorSystemSettings;`, {
  normalizeLatexFormatProfile,
  useEditorStore: { getState: () => state },
  readLocalStorage: () => null,
}).run;
const completed = [];
for (const raw of [null, "", "{", "{}", '{"state":null}']) sync(raw);
assert.equal(calls.length, 0);
completed.push("missing/malformed persistence is ignored");
const profile = {
  inlineWrapper: "paren",
  inlineTextPolicy: "text-command",
  displayWrapper: "equation",
  numbered: true,
  multilineEnvironment: "align",
};
const payload = JSON.stringify({ state: { latexFormatProfile: profile } });
sync(payload);
assert.deepEqual(state.latexFormatProfile, profile);
assert.deepEqual(calls, ["profile"]);
completed.push("all five persisted profile fields synchronize");
sync(payload);
assert.deepEqual(calls, ["profile"]);
completed.push("identical profile does not trigger another store write");
sync(JSON.stringify({ state: { latexFormatProfile: { ...profile, inlineWrapper: "dollar" } } }));
assert.equal(state.latexFormatProfile.inlineWrapper, "dollar");
completed.push("explicit dollar wrapper can replace paren wrapper");
sync(JSON.stringify({ state: { inputBehavior: { autoExitAccent: false, autoExitSubscript: "false", unknown: true } } }));
assert.equal(state.inputBehavior.autoExitAccent, false);
assert.equal(state.inputBehavior.autoExitSubscript, true);
assert.deepEqual(state.lines, [{ id: "unchanged", latex: "x+1" }]);
assert.equal(state.activeLineId, "unchanged");
assert.equal(state.title, "untouched Office session");
completed.push("input settings reject wrong types and preserve session content");

const effects = collect((node) => ts.isCallExpression(node)
  && node.expression.getText(ast) === "useEffect"
  && node.arguments[0]?.getText(ast).includes("initialEditorFocusSessionRef.current === sessionKey"));
assert.equal(effects.length, 1, "Expected the actual initial-focus effect");
const effectCode = `globalThis.run = ${effects[0].arguments[0].getText(ast)};`;
function focusEnvironment(overrides = {}) {
  let identifier = 0;
  const frames = new Map();
  const intervals = new Map();
  const timeouts = new Map();
  const windowListeners = new Map();
  const documentListeners = new Map();
  const attempts = [];
  let fieldReady = true;
  let editorFocusWorks = true;
  let fallbackFocuses = 0;
  const document = {
    visibilityState: "visible",
    activeElement: { tagName: "BODY", classList: { contains: () => false } },
    querySelector: () => fieldReady ? field : null,
    addEventListener: (name, callback) => documentListeners.set(name, callback),
    removeEventListener: (name) => documentListeners.delete(name),
  };
  const field = { tagName: "MATH-FIELD", focus: () => {
    fallbackFocuses += 1;
    document.activeElement = field;
  } };
  const context = sandbox(effectCode, {
    sessionHydrated: true,
    sessionKey: "7:first-session",
    tauriResidentEditor: true,
    presentedSessionKey: "7:first-session",
    initialEditorFocusSessionRef: { current: "" },
    activeSessionKeyRef: { current: "7:first-session" },
    editorRef: { current: { focus: (options) => {
      attempts.push(JSON.parse(JSON.stringify(options)));
      if (fieldReady && editorFocusWorks) document.activeElement = field;
    } } },
    document,
    window: {
      focus: () => undefined,
      requestAnimationFrame: (callback) => { frames.set(++identifier, callback); return identifier; },
      cancelAnimationFrame: (id) => frames.delete(id),
      setInterval: (callback, milliseconds) => { intervals.set(++identifier, { callback, milliseconds }); return identifier; },
      clearInterval: (id) => intervals.delete(id),
      setTimeout: (callback, milliseconds) => { timeouts.set(++identifier, { callback, milliseconds }); return identifier; },
      clearTimeout: (id) => timeouts.delete(id),
      addEventListener: (name, callback) => windowListeners.set(name, callback),
      removeEventListener: (name) => windowListeners.delete(name),
    },
    ...overrides,
  });
  return {
    context, frames, intervals, timeouts, windowListeners, documentListeners, attempts,
    setFieldReady: (value) => { fieldReady = value; },
    setEditorFocusWorks: (value) => { editorFocusWorks = value; },
    fallbackFocuses: () => fallbackFocuses,
    flushFrame: () => {
      for (const [id, callback] of [...frames]) { frames.delete(id); callback(); }
    },
  };
}
for (const overrides of [
  { sessionHydrated: false },
  { sessionKey: "" },
  { presentedSessionKey: "" },
  { presentedSessionKey: "6:previous-session" },
]) {
  const env = focusEnvironment(overrides);
  assert.equal(env.context.run(), undefined);
  assert.equal(env.frames.size, 0);
  assert.equal(env.attempts.length, 0);
}
completed.push("parked/unhydrated/stale-generation windows never schedule focus");
{
  const env = focusEnvironment();
  const cleanup = env.context.run();
  assert.equal(env.attempts.length, 0, "Native presentation must precede the first focus frame");
  env.flushFrame();
  assert.deepEqual(env.attempts, [{ target: "first", moveToEnd: true }]);
  assert.equal(env.intervals.size, 0);
  assert.equal(env.windowListeners.size, 0);
  assert.equal(env.documentListeners.size, 0);
  assert.equal(env.context.run(), undefined, "An already-focused session must not restart activation repair");
  cleanup();
}
completed.push("presented session focuses first line once and releases activation listeners");
{
  const env = focusEnvironment();
  env.setEditorFocusWorks(false);
  const cleanup = env.context.run();
  env.flushFrame();
  assert.equal(env.fallbackFocuses(), 1);
  assert.equal(env.intervals.size, 0);
  cleanup();
}
completed.push("custom-element fallback handles delayed imperative focus");
{
  const env = focusEnvironment();
  env.setFieldReady(false);
  const cleanup = env.context.run();
  env.flushFrame();
  assert.equal(env.intervals.size, 1);
  const interval = [...env.intervals.values()][0];
  assert.equal(interval.milliseconds, 60);
  env.setFieldReady(true);
  interval.callback();
  assert.equal(env.intervals.size, 0);
  assert.equal(env.windowListeners.size, 0);
  assert.equal(env.documentListeners.size, 0);
  cleanup();
  assert.equal(env.timeouts.size, 0);
}
completed.push("shadow-input mounting retries stop immediately after focus succeeds");
{
  const env = focusEnvironment();
  env.setFieldReady(false);
  const cleanup = env.context.run();
  env.flushFrame();
  const deadline = [...env.timeouts.values()][0];
  assert.equal(deadline.milliseconds, 3000);
  deadline.callback();
  assert.equal(env.intervals.size, 0);
  assert.equal(env.windowListeners.size, 0);
  assert.equal(env.documentListeners.size, 0);
  cleanup();
  assert.equal(env.timeouts.size, 0);
}
completed.push("failed activation is bounded to three seconds");
{
  const env = focusEnvironment();
  const cleanup = env.context.run();
  cleanup();
  assert.equal(env.frames.size, 0);
  assert.equal(env.intervals.size, 0);
  assert.equal(env.windowListeners.size, 0);
  assert.equal(env.documentListeners.size, 0);
  env.flushFrame();
  assert.equal(env.attempts.length, 0);
}
completed.push("cleanup cancels pending focus before session replacement/unmount");
{
  const env = focusEnvironment();
  const firstCleanup = env.context.run();
  firstCleanup();
  assert.equal(env.context.initialEditorFocusSessionRef.current, "");
  const secondCleanup = env.context.run();
  env.flushFrame();
  assert.equal(env.attempts.length, 1, "StrictMode's second setup must still acquire focus");
  assert.equal(env.context.initialEditorFocusSessionRef.current, "7:first-session");
  secondCleanup();
}
completed.push("StrictMode setup/cleanup/setup does not suppress activation");
{
  const env = focusEnvironment();
  const cleanup = env.context.run();
  env.context.activeSessionKeyRef.current = "8:replacement-session";
  env.flushFrame();
  assert.equal(env.attempts.length, 0, "An old animation frame must not focus a replacement session");
  assert.equal(env.intervals.size, 0, "A stale animation frame must not start new repair timers");
  assert.equal(env.timeouts.size, 0);
  cleanup();
}
completed.push("stale activation frames cannot focus or restart timers for a new session");
{
  const env = focusEnvironment({ tauriResidentEditor: false, presentedSessionKey: "" });
  const cleanup = env.context.run();
  env.flushFrame();
  assert.equal(env.attempts.length, 1);
  cleanup();
}
completed.push("browser Office sessions do not require an AppKit presentation report");
console.log(JSON.stringify({ passed: completed.length, cases: completed }, null, 2));
console.log("macOS Office editor profile/focus source-execution regression passed (native UI acceptance still required)");
