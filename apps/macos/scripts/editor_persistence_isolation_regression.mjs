import assert from "node:assert/strict";
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
const envelope = (state) => ({ version: 0, state });
const stored = (state) => JSON.stringify(envelope(state));
const completed = [];
function fixture(initial = stored(mainState), office = true) {
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
  const loaded = adapter.getItem(key).state;
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
  const { adapter, data } = fixture();
  adapter.removeItem(key);
  assert.equal(data.get(key), stored(mainState), "Office clearStorage must not erase the main document");
}
completed.push("Office clearStorage cannot overwrite or erase main data");
for (const invalid of ["{", "null", "[]", "{}", '{"state":[]}']) {
  const { adapter, data } = fixture(invalid);
  assert.equal(adapter.getItem(key), null);
  adapter.setItem(key, envelope(officeState));
  assert.equal(data.get(key), invalid, "Corrupt existing data must remain available for recovery");
}
completed.push("corrupt pre-existing storage is retained rather than replaced with Office defaults");
{
  const { adapter, data } = fixture(stored(mainState), false);
  assert.deepEqual(adapter.getItem(key), envelope(mainState));
  adapter.setItem(key, envelope(officeState));
  assert.equal(data.get(key), stored(officeState));
  adapter.removeItem(key);
  assert.equal(data.has(key), false);
}
completed.push("main application persistence remains a transparent pass-through");
{
  const { adapter, data } = fixture();
  adapter.setItem("unrelated", envelope({ value: 1 }));
  assert.deepEqual(adapter.getItem("unrelated"), envelope({ value: 1 }));
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
{
  for (const office of [false, true]) {
    const { adapter, writes } = fixture(undefined, office);
    let encodes = 0;
    const lines = [{ id: "main", latex: "x", toJSON() { encodes++; return { id: "main", latex: "x" }; } }];
    const state = { ...mainState, lines };
    adapter.setItem(key, envelope(state));
    const written = writes.length;
    const serialized = encodes;
    for (let i = 0; i < 100; i++) {
      adapter.setItem(key, envelope({ ...state, activeLineId: `line-${i}` }));
    }
    assert.equal(writes.length, written, "Caret changes must not write storage");
    assert.equal(encodes, serialized, "Caret changes must not serialize formulas");
    if (office) {
      adapter.setItem(key, envelope({ ...state, lines: [{ latex: "new session formula" }] }));
      assert.equal(writes.length, written, "Office typing must not write the main document");
    } else {
      adapter.setItem(key, envelope({ ...state, title: "Changed", activeLineId: "latest" }));
      assert.equal(writes.length, written + 1, "A document edit must still be saved immediately");
    }
  }
}
completed.push("caret moves and Office typing skip document serialization; real edits persist immediately");
console.log(JSON.stringify({ passed: completed.length, cases: completed }, null, 2));
console.log("macOS editor persistence isolation regression passed");
