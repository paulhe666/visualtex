import assert from "node:assert/strict";
import {
  prepareEditorCrashSafeRestart,
  VISUALTEX_EDITOR_CRASH_BACKUP_STORAGE_KEY,
} from "../src/runtime/editorCrashRecovery.ts";
import { createSafeStorage, safeStorage } from "../src/runtime/safeStorage.ts";
import { memoryStorage } from "./memory_storage.ts";

const storage = memoryStorage();
Object.assign(safeStorage, createSafeStorage(() => storage));

const editorKey = "visualtex-editor";

function reset() {
  safeStorage.removeItem(editorKey);
  safeStorage.removeItem(VISUALTEX_EDITOR_CRASH_BACKUP_STORAGE_KEY);
}

reset();

const original = JSON.stringify({
  state: {
    title: "Pathological document",
    lines: [{ id: "deep", latex: "\\begin{array}{c}x\\end{array}" }],
    activeLineId: "deep",
    history: [{ latex: "previous formula" }],
    theme: "dark",
    zoom: 0.7,
    editorLayout: "classic",
    formulaLetterFont: "stix",
    inputBehavior: {
      autoEscapeShortcuts: false,
      autoExitSuperscript: true,
    },
    usage: {
      frac: {
        commandId: "frac",
        useCount: 3,
        lastUsedAt: 123,
        recentUses: [123],
        acceptedPrefixes: {},
        contextCounts: {},
        pinned: false,
      },
    },
  },
  version: 0,
});

safeStorage.setItem(editorKey, original);
const result = prepareEditorCrashSafeRestart();
assert.deepEqual(result, { hadPersistedEditor: true, backupCreated: true });
assert.equal(
  safeStorage.getItem(VISUALTEX_EDITOR_CRASH_BACKUP_STORAGE_KEY),
  original,
  "safe restart must preserve the complete original persistence envelope",
);

const recoveredRaw = safeStorage.getItem(editorKey);
assert.ok(recoveredRaw);
const recovered = JSON.parse(recoveredRaw) as {
  state: Record<string, unknown>;
  version: number;
};
for (const key of ["title", "latex", "lines", "activeLineId", "history"]) {
  assert.equal(
    Object.prototype.hasOwnProperty.call(recovered.state, key),
    false,
    `safe restart must remove risky session field ${key}`,
  );
}
assert.equal(recovered.state.theme, "dark");
assert.equal(recovered.state.zoom, 0.7);
assert.equal(recovered.state.editorLayout, "classic");
assert.equal(recovered.state.formulaLetterFont, "stix");
assert.deepEqual(recovered.state.inputBehavior, {
  autoEscapeShortcuts: false,
  autoExitSuperscript: true,
});
assert.ok(recovered.state.usage, "recommendation preferences must be retained");

assert.deepEqual(prepareEditorCrashSafeRestart(), {
  hadPersistedEditor: true, backupCreated: false,
});
assert.equal(storage.getItem(VISUALTEX_EDITOR_CRASH_BACKUP_STORAGE_KEY), original);
const newDocument = JSON.stringify({ state: { lines: [{ latex: "new work" }] } });
storage.setItem(editorKey, newDocument);
assert.throws(() => prepareEditorCrashSafeRestart(), /已有待恢复/);
assert.equal(storage.getItem(editorKey), newDocument);
assert.equal(storage.getItem(VISUALTEX_EDITOR_CRASH_BACKUP_STORAGE_KEY), original);

reset();
storage.setItem(editorKey, original);
const storeWrite = storage.setItem;
storage.setItem = (key, value) => {
  if (key === editorKey) throw new Error("write failed");
  storeWrite(key, value);
};
assert.throws(() => prepareEditorCrashSafeRestart(), /write failed/);
assert.equal(storage.getItem(editorKey), original, "failed write must not delete active content");
assert.equal(storage.getItem(VISUALTEX_EDITOR_CRASH_BACKUP_STORAGE_KEY), original);
storage.setItem = storeWrite;
assert.deepEqual(prepareEditorCrashSafeRestart(), {
  hadPersistedEditor: true, backupCreated: false,
}, "retry can reuse the identical original backup after a failed active write");
assert.equal(storage.getItem(VISUALTEX_EDITOR_CRASH_BACKUP_STORAGE_KEY), original);
assert.equal(storage.getItem(editorKey), recoveredRaw);

reset();
safeStorage.setItem(editorKey, "{corrupt persistence");
const corruptResult = prepareEditorCrashSafeRestart();
assert.deepEqual(corruptResult, {
  hadPersistedEditor: true,
  backupCreated: true,
});
assert.equal(
  safeStorage.getItem(VISUALTEX_EDITOR_CRASH_BACKUP_STORAGE_KEY),
  "{corrupt persistence",
);
assert.equal(safeStorage.getItem(editorKey), null);

reset();
assert.deepEqual(prepareEditorCrashSafeRestart(), {
  hadPersistedEditor: false,
  backupCreated: false,
});

console.log("VisualTeX editor crash recovery regression passed");
