import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

class MemoryStorage {
  private values = new Map<string, string>();

  get length() {
    return this.values.size;
  }

  clear() {
    this.values.clear();
  }

  getItem(key: string) {
    return this.values.get(key) ?? null;
  }

  key(index: number) {
    return [...this.values.keys()][index] ?? null;
  }

  removeItem(key: string) {
    this.values.delete(key);
  }

  setItem(key: string, value: string) {
    this.values.set(key, String(value));
  }
}

const localStorage = new MemoryStorage();
const styleValues = new Map<string, string>();
const root = {
  dataset: {} as Record<string, string>,
  style: {
    setProperty(key: string, value: string) {
      styleValues.set(key, value);
    },
    removeProperty(key: string) {
      const previous = styleValues.get(key) ?? "";
      styleValues.delete(key);
      return previous;
    },
  },
};

Object.defineProperty(globalThis, "window", {
  configurable: true,
  value: {
    innerWidth: 1333,
    innerHeight: 777,
    localStorage,
    location: { search: "" },
    addEventListener() {},
    removeEventListener() {},
    dispatchEvent() {
      return true;
    },
  },
});
Object.defineProperty(globalThis, "document", {
  configurable: true,
  value: { documentElement: root },
});
Object.defineProperty(globalThis, "BroadcastChannel", {
  configurable: true,
  value: undefined,
});

const settingsDialogSource = readFileSync(
  new URL("../src/components/SettingsDialog.tsx", import.meta.url),
  "utf8",
);
assert.match(
  settingsDialogSource,
  /invoke\("write_export_file",\s*\{\s*request:\s*\{\s*path,\s*base64:/s,
  "configuration export must use the current nested write_export_file request contract",
);
assert.doesNotMatch(
  settingsDialogSource,
  /invoke\("write_export_file",\s*\{\s*path,\s*dataBase64:/s,
  "configuration export must not use the obsolete flat write_export_file contract",
);

const tauriLibSource = readFileSync(
  new URL("../src-tauri/src/lib.rs", import.meta.url),
  "utf8",
);
const windowsBackendSource = readFileSync(
  new URL("../src-tauri/src/office/windows_backend.rs", import.meta.url),
  "utf8",
);
assert.match(
  tauriLibSource,
  /get_word_numbering_user_configuration/,
  "configuration export must expose the Word numbering user preference read command",
);
assert.match(
  tauriLibSource,
  /apply_word_numbering_user_configuration/,
  "configuration import must expose the Word numbering user preference write command",
);
assert.match(
  windowsBackendSource,
  /HKCU\\Software\\VisualTeX\\Word/,
  "configuration Word preferences must use the same HKCU key as the Word VSTO add-in",
);

const {
  VISUALTEX_CONFIGURATION_SCHEMA,
  VISUALTEX_CONFIGURATION_VERSION,
  applyVisualTexConfiguration,
  buildVisualTexConfiguration,
  parseVisualTexConfiguration,
} = await import("../src/runtime/applicationConfiguration.ts");
const { useEditorStore } = await import("../src/stores/editorStore.ts");
const { safeStorage } = await import("../src/runtime/safeStorage.ts");

const editor = useEditorStore.getState();
editor.setTheme("raycast");
editor.setLanguage("en");
editor.setEditorLayout("classic");
editor.setZoom(1.25);
editor.setSourceOpen(true);
editor.setAutoPairDelimiters(false);
editor.setShowLineNumbers(true);
editor.setHighlightActiveLine(false);
editor.setFormulaInsetLeft(17);
editor.setFormulaInsetRight(29);
editor.setFormulaToolButtonSize(61);
editor.setFormulaToolButtonPadding(7);
editor.setFormulaRowVerticalInset(13);
editor.setPngExportBackground("#123456");
editor.setFormulaLetterFont("palatino");
editor.setFormulaChineseFont("songti");
editor.setPersonalize(false);
editor.setSuggestionCount(9);
editor.setCheckUpdatesOnStartup(false);
editor.setPowerPointDefaultFontSizePt(23.5);
editor.setClassicTileWidth(486);
editor.setClassicDockHeight(372);
editor.setKeypadMinimizeOnCopy(false);
editor.setInputBehavior("autoEscapeShortcuts", false);
editor.setInputBehavior("autoExitSuperscript", false);
editor.setInputBehavior("showOtherCommandSuggestions", true);
useEditorStore.setState({
  usage: {
    "mathlive-native:\\begin{cases}": {
      commandId: "mathlive-native:\\begin{cases}",
      useCount: 27,
      lastUsedAt: 1_765_000_000_000,
      recentUses: [1_764_999_999_000, 1_765_000_000_000],
      acceptedPrefixes: { b: 9, beg: 7 },
      contextCounts: { candidate: 18, toolbar: 9 },
      pinned: false,
    },
  },
  history: [
    {
      id: "history-a",
      latex: "\\begin{cases}x,&x>0\\\\0,&x\\le 0\\end{cases}",
      createdAt: 1_765_000_000_000,
    },
  ],
});

const expectedStorage = {
  "visualtex-custom-formula-tiles": JSON.stringify([{ id: "tile-a", latex: "x+y" }]),
  "visualtex-common-toolbar-command-ids-v1": JSON.stringify(["frac", "sqrt"]),
  "visualtex-common-toolbar-command-ids-v2": JSON.stringify(["sum", "int"]),
  "visualtex-formula-hotkeys-v1": JSON.stringify([{ commandId: "frac", code: "KeyF" }]),
  "visualtex-custom-formula-text-colors": JSON.stringify(["#123456"]),
  "visualtex-custom-formula-background-colors": JSON.stringify(["#fedcba"]),
  "visualtex-desktop-editor-toolbar-open": "false",
  "visualtex-desktop-editor-tiles-open": "true",
  "visualtex-desktop-editor-source-open": "false",
  "visualtex-office-editor-toolbar-open": "true",
  "visualtex-office-editor-tiles-open": "false",
  "visualtex-office-editor-source-open": "true",
  "visualtex.ocr.model": "PP-FormulaNet_plus-L",
  "visualtex.silent-ocr.enabled": "true",
  "visualtex.quick-ocr.capture-mode": "system-screenshot",
  "visualtex.custom-theme.v1": JSON.stringify({ mode: "light", colors: {} }),
  "visualtex.custom-symbols.v1": JSON.stringify({ version: 1, symbols: [] }),
} as const;
for (const [key, value] of Object.entries(expectedStorage)) {
  safeStorage.setItem(key, value);
}

const built = await buildVisualTexConfiguration();
assert.equal(built.schema, VISUALTEX_CONFIGURATION_SCHEMA);
assert.equal(built.version, VISUALTEX_CONFIGURATION_VERSION);
assert.deepEqual(built.windows, { main: { width: 1333, height: 777 } });
assert.equal(built.editorSettings.theme, "raycast");
assert.equal(built.editorSettings.zoom, 1.25);
assert.equal(built.editorSettings.formulaLetterFont, "palatino");
assert.equal(built.editorSettings.formulaChineseFont, "songti");
assert.equal(built.editorSettings.classicTileWidth, 486);
assert.equal(built.editorSettings.classicDockHeight, 372);
assert.equal(built.editorSettings.keypadMinimizeOnCopy, false);
assert.equal(built.editorSettings.inputBehavior?.autoEscapeShortcuts, false);
assert.equal(built.editorSettings.inputBehavior?.autoExitSuperscript, false);
assert.equal(built.editorSettings.inputBehavior?.showOtherCommandSuggestions, true);
assert.ok(
  built.capturedStorageKeys?.includes("visualtex-custom-formula-text-colors"),
  "current configuration exports must declare the storage keys they understand",
);
assert.equal(
  built.personalization?.usage?.["mathlive-native:\\begin{cases}"]?.useCount,
  27,
  "configuration export must preserve personalized command frequency",
);
assert.equal(
  built.personalization?.history?.[0]?.id,
  "history-a",
  "configuration export must preserve editor history",
);
assert.equal(built.word, undefined, "browser regression should leave native Word preferences absent");
for (const [key, value] of Object.entries(expectedStorage)) {
  assert.equal(built.storage[key], value, `configuration build must include ${key}`);
}

const withNativeWindows = {
  ...built,
  windows: {
    main: { width: 1440, height: 900 },
    officeEditor: { width: 1180, height: 760 },
  },
};
const parsed = parseVisualTexConfiguration(JSON.stringify(withNativeWindows));
assert.deepEqual(parsed.windows, {
  main: { width: 1440, height: 900 },
  keypad: null,
  officeEditor: { width: 1180, height: 760 },
});

editor.setTheme("light");
editor.setLanguage("cn");
editor.setEditorLayout("standard");
editor.setZoom(0.6);
editor.setFormulaLetterFont("katex");
editor.setFormulaChineseFont("system");
editor.setClassicTileWidth(300);
editor.setClassicDockHeight(240);
editor.setKeypadMinimizeOnCopy(true);
for (const key of Object.keys(expectedStorage)) safeStorage.removeItem(key);
safeStorage.setItem("visualtex.quick-ocr.capture-mode", "immediate");
useEditorStore.setState({ usage: {}, history: [] });

await applyVisualTexConfiguration(parsed);
const restored = useEditorStore.getState();
assert.equal(restored.theme, "raycast");
assert.equal(restored.language, "en");
assert.equal(restored.editorLayout, "classic");
assert.equal(restored.zoom, 1.25);
assert.equal(restored.formulaLetterFont, "palatino");
assert.equal(restored.formulaChineseFont, "songti");
assert.equal(restored.classicTileWidth, 486);
assert.equal(restored.classicDockHeight, 372);
assert.equal(restored.keypadMinimizeOnCopy, false);
assert.equal(restored.inputBehavior.autoEscapeShortcuts, false);
assert.equal(
  restored.usage["mathlive-native:\\begin{cases}"]?.useCount,
  27,
  "configuration import must restore personalized command frequency",
);
assert.equal(
  restored.history[0]?.id,
  "history-a",
  "configuration import must restore editor history",
);
const expectedRestoredStorage = {
  ...expectedStorage,
  "visualtex.quick-ocr.capture-mode": "clipboard",
};
for (const [key, value] of Object.entries(expectedRestoredStorage)) {
  assert.equal(safeStorage.getItem(key), value, `configuration import must restore ${key}`);
}

const sanitized = parseVisualTexConfiguration(
  JSON.stringify({
    schema: VISUALTEX_CONFIGURATION_SCHEMA,
    version: VISUALTEX_CONFIGURATION_VERSION,
    exportedAt: "invalid-but-allowed-label",
    editorSettings: {
      theme: "green",
      keypadMinimizeOnCopy: false,
      unknownFutureSetting: "ignored",
    },
    storage: {
      "visualtex.ocr.model": "not-a-model",
      "visualtex.silent-ocr.enabled": "maybe",
      "visualtex.quick-ocr.capture-mode": "invalid-mode",
      "visualtex-custom-formula-tiles": "not-json",
      "visualtex-formula-hotkeys-v1": "[]",
      unexpected: "ignored",
    },
    windows: {
      main: { width: 100, height: 100 },
      officeEditor: { width: 1200, height: 800 },
    },
  }),
);
assert.deepEqual(sanitized.editorSettings, {
  theme: "green",
  keypadMinimizeOnCopy: false,
});
assert.deepEqual(sanitized.storage, {
  "visualtex-formula-hotkeys-v1": "[]",
});
assert.deepEqual(sanitized.windows, {
  main: null,
  keypad: null,
  officeEditor: { width: 1200, height: 800 },
});

// Legacy v1 files predate capturedStorageKeys/personalization/Word preferences.
// Importing one must only touch fields it actually contains; newer user data
// must not be erased merely because the old exporter did not know those fields.
safeStorage.setItem("visualtex-custom-formula-text-colors", JSON.stringify(["#abcdef"]));
useEditorStore.setState({
  usage: {
    "future-command": {
      commandId: "future-command",
      useCount: 4,
      lastUsedAt: 123,
      recentUses: [123],
      acceptedPrefixes: { f: 4 },
      contextCounts: { candidate: 4 },
      pinned: false,
    },
  },
});
const legacyV1 = parseVisualTexConfiguration(
  JSON.stringify({
    schema: VISUALTEX_CONFIGURATION_SCHEMA,
    version: 1,
    exportedAt: "2026-01-01T00:00:00.000Z",
    editorSettings: { zoom: 1.1 },
    storage: { "visualtex.ocr.model": "PP-FormulaNet_plus-S" },
  }),
);
await applyVisualTexConfiguration(legacyV1);
assert.equal(useEditorStore.getState().zoom, 1.1);
assert.equal(safeStorage.getItem("visualtex.ocr.model"), "PP-FormulaNet_plus-S");
assert.equal(
  safeStorage.getItem("visualtex-custom-formula-text-colors"),
  JSON.stringify(["#abcdef"]),
  "legacy imports must preserve newer storage keys absent from the old file",
);
assert.equal(
  useEditorStore.getState().usage["future-command"]?.useCount,
  4,
  "legacy imports without personalization must preserve the current learned ranking",
);

// Current files explicitly list the storage keys captured by their exporter, so
// an omitted value can intentionally clear that setting without ambiguity.
const clearKnownStorage = parseVisualTexConfiguration(
  JSON.stringify({
    ...built,
    storage: Object.fromEntries(
      Object.entries(built.storage).filter(
        ([key]) => key !== "visualtex-custom-formula-text-colors",
      ),
    ),
  }),
);
safeStorage.setItem("visualtex-custom-formula-text-colors", JSON.stringify(["#112233"]));
await applyVisualTexConfiguration(clearKnownStorage);
assert.equal(
  safeStorage.getItem("visualtex-custom-formula-text-colors"),
  null,
  "current imports must be able to intentionally clear a captured setting",
);

// Additive future files use the same schema and are imported best-effort: known
// fields are restored and unknown fields are ignored instead of rejecting the file.
const futureVersion = parseVisualTexConfiguration(
  JSON.stringify({
    schema: VISUALTEX_CONFIGURATION_SCHEMA,
    version: 999,
    exportedAt: "2099-01-01T00:00:00.000Z",
    editorSettings: { theme: "green", unknownFutureSetting: "ignored" },
    storage: {},
    futureSection: { anything: true },
  }),
);
assert.equal(futureVersion.version, VISUALTEX_CONFIGURATION_VERSION);
assert.deepEqual(futureVersion.editorSettings, { theme: "green" });

const legacyWithoutEditorSettings = parseVisualTexConfiguration(
  JSON.stringify({
    schema: VISUALTEX_CONFIGURATION_SCHEMA,
    version: 1,
    storage: {},
  }),
);
assert.deepEqual(
  legacyWithoutEditorSettings.editorSettings,
  {},
  "missing settings from an older configuration must be treated as absent, not fatal",
);

assert.throws(
  () => parseVisualTexConfiguration("{}"),
  /not a VisualTeX configuration/i,
);
assert.throws(
  () =>
    parseVisualTexConfiguration(
      JSON.stringify({
        schema: VISUALTEX_CONFIGURATION_SCHEMA,
        version: 0,
        editorSettings: {},
        storage: {},
      }),
    ),
  /version is not supported/i,
);

console.log("VisualTeX Windows application configuration regression: PASS");
