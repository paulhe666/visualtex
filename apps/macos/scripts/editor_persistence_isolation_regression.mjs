import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
  createEditorPersistenceStorage,
  isOfficeEditorPersistenceScope,
} from "../src/runtime/editorPersistenceStorage.ts";

// All storage is a private Map in this process, not browser/user storage.
const key = "visualtex-editor";
const isolated = ["title", "lines", "activeLineId", "formulaAlignment", "latexCodeFormat", "history"];
const mainState = {
  title: "Main document", lines: [{ id: "main", latex: "a+b", mode: "inline" }],
  activeLineId: "main", formulaAlignment: "right", latexCodeFormat: "equation",
  history: [{ id: "history", latex: "y", createdAt: 1 }],
  latexFormatProfile: { displayWrapper: "equation", numbered: false },
  language: "cn", zoom: 0.5, futurePreference: "preserve",
};
const officeState = {
  title: "Office session", lines: [{ id: "office", latex: "x+1" }],
  activeLineId: "office", formulaAlignment: "left", latexCodeFormat: "raw", history: [],
  latexFormatProfile: { displayWrapper: "double-dollar", numbered: true }, language: "en", zoom: 0.6,
};
const envelope = (state) => JSON.stringify({ version: 0, state });
const completed = [];
function fixture(initial = envelope(mainState), office = true) {
  const data = new Map(initial === null ? [] : [[key, initial]]);
  const writes = [];
  const storage = {
    getItem: (name) => data.get(name) ?? null,
    setItem: (name, value) => { writes.push([name, value]); data.set(name, value); },
    removeItem: (name) => data.delete(name),
  };
  return { data, writes, storage, adapter: createEditorPersistenceStorage(storage, () => office) };
}
{
  const { adapter, data } = fixture();
  const loaded = JSON.parse(adapter.getItem(key)).state;
  for (const name of isolated) assert(!Object.hasOwn(loaded, name), `${name} leaked into Office hydration`);
  assert.deepEqual(loaded.latexFormatProfile, mainState.latexFormatProfile);
  assert.equal(loaded.language, "cn");
  assert.deepEqual(JSON.parse(data.get(key)).state, mainState, "Read must not modify storage");
}
completed.push("Office reads shared settings without loading the main document");
{
  const { adapter, data } = fixture();
  adapter.setItem(key, envelope(officeState));
  const saved = JSON.parse(data.get(key)).state;
  for (const name of isolated) assert.deepEqual(saved[name], mainState[name], `${name} was overwritten by Office`);
  assert.deepEqual(saved.latexFormatProfile, officeState.latexFormatProfile);
  assert.equal(saved.language, "en");
  assert.equal(saved.zoom, 0.6);
  assert.equal(saved.futurePreference, "preserve");
}
completed.push("all six document fields remain isolated while format/settings synchronize");
{
  const { adapter, data, writes } = fixture();
  adapter.setItem(key, envelope(officeState));
  const count = writes.length;
  adapter.setItem(key, envelope(officeState));
  assert.equal(writes.length, count, "Identical Office preference writes must not emit another storage event");
  const current = JSON.parse(data.get(key));
  current.state.title = "New main document";
  current.state.lines = [{ id: "latest", latex: "z^2" }];
  data.set(key, JSON.stringify(current));
  adapter.setItem(key, envelope({ ...officeState, language: "cn" }));
  const saved = JSON.parse(data.get(key)).state;
  assert.equal(saved.title, "New main document");
  assert.deepEqual(saved.lines, current.state.lines);
}
completed.push("every Office write preserves the newest stored main document, without duplicate events");
{
  const { adapter, data } = fixture(null);
  adapter.setItem(key, envelope(officeState));
  const saved = JSON.parse(data.get(key)).state;
  for (const name of isolated) assert(!Object.hasOwn(saved, name));
  assert.deepEqual(saved.latexFormatProfile, officeState.latexFormatProfile);
}
completed.push("first-run Office sessions persist settings only, never a session document");
{
  const { adapter, data, writes } = fixture();
  for (const invalid of ["{", "null", "[]", "{}", '{"state":[]}', '{"state":null}']) adapter.setItem(key, invalid);
  assert.equal(writes.length, 0);
  assert.equal(data.get(key), envelope(mainState));
  adapter.removeItem(key);
  assert.equal(data.get(key), envelope(mainState), "Office clearStorage must not erase the main document");
}
completed.push("invalid writes and Office clearStorage cannot overwrite or erase main data");
for (const invalid of ["{", "null", "[]", "{}", '{"state":[]}']) {
  const { adapter, data } = fixture(invalid);
  assert.equal(adapter.getItem(key), null);
  adapter.setItem(key, envelope(officeState));
  assert.equal(data.get(key), invalid, "Corrupt existing data must remain available for recovery");
}
completed.push("corrupt pre-existing storage is retained rather than replaced with Office defaults");
{
  const { adapter, data } = fixture(envelope(mainState), false);
  assert.equal(adapter.getItem(key), envelope(mainState));
  adapter.setItem(key, envelope(officeState));
  assert.equal(data.get(key), envelope(officeState));
  adapter.removeItem(key);
  assert.equal(data.has(key), false);
}
completed.push("main application persistence remains a transparent pass-through");
{
  const { adapter, data } = fixture();
  adapter.setItem("unrelated", "value");
  assert.equal(adapter.getItem("unrelated"), "value");
  adapter.removeItem("unrelated");
  assert.equal(data.has("unrelated"), false);
}
completed.push("unrelated storage keys retain ordinary behavior");
for (const location of [
  { pathname: "/office-native-dialog.html", search: "?sessionId=test" },
  { pathname: "/nested/office-native-dialog.html", search: "" },
  { pathname: "/", search: "?view=office-formula&officeHost=word" },
  { pathname: "/", search: "?view=office-document-import" },
  { pathname: "/", search: "?view=office-word-latex-redraw" },
]) assert.equal(isOfficeEditorPersistenceScope(location), true);
for (const location of [
  null, { pathname: "/", search: "" },
  { pathname: "/", search: "?view=desktop" },
  { pathname: "/office-native-dialog.html-unrelated", search: "" },
]) assert.equal(isOfficeEditorPersistenceScope(location), false);
completed.push("native-dialog and resident Office URLs are isolated, ordinary desktop URLs are not");
const storeSource = readFileSync(new URL("../src/stores/editorStore.ts", import.meta.url), "utf8");
assert.match(storeSource, /storage:\s*createJSONStorage\(\(\)\s*=>\s*editorPersistenceStorage\)/);
completed.push("the actual Zustand persistence entry is wired to the isolation adapter");
console.log(JSON.stringify({ passed: completed.length, cases: completed }, null, 2));
console.log("macOS editor persistence isolation regression passed");
