import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
  normalizeOfficeFormulaShortcut,
  normalizePowerPointFormulaShortcutSettings,
  normalizeWordFormulaShortcutSettings,
  officeFormulaShortcutFromEvent,
} from "../src/office/officeFormulaShortcuts";
import {
  describeWordShortcutBinding,
  parseWordShortcutConflicts,
  WORD_FREE_SHORTCUT_SUGGESTIONS,
  wordDefaultBinding,
} from "../src/office/wordShortcutConflicts";

// Word keeps its own default bindings; the settings page must say so.
assert.deepEqual(wordDefaultBinding("Ctrl+D"), { category: 1, name: "FormatFont" });
assert.deepEqual(wordDefaultBinding("Ctrl+F"), { category: 1, name: "SmartFind" });
assert.equal(wordDefaultBinding("Ctrl+Shift+7")?.category, 7);
assert.equal(wordDefaultBinding("Ctrl+Alt+Q"), null);
assert.equal(wordDefaultBinding(null), null);
for (const suggestion of WORD_FREE_SHORTCUT_SUGGESTIONS) {
  assert.equal(wordDefaultBinding(suggestion), null, `${suggestion} must be free in Word`);
  assert.equal(normalizeOfficeFormulaShortcut(suggestion), suggestion);
}
assert.equal(
  describeWordShortcutBinding({ category: 1, name: "FormatFont" }, false),
  "Word 的「字体」",
);
assert.equal(
  describeWordShortcutBinding({ category: 1, name: "FormatFont" }, true),
  "the Word command “FormatFont”",
);
const reported = parseWordShortcutConflicts("Ctrl+Alt+Q=2:MTCommand_InsertInlineEqn;Ctrl+F7=0:;bad");
assert.deepEqual(reported.get("Ctrl+Alt+Q"), { category: 2, name: "MTCommand_InsertInlineEqn" });
assert.deepEqual(reported.get("Ctrl+F7"), { category: 0, name: "" });
assert.equal(reported.size, 2);
assert.equal(parseWordShortcutConflicts(undefined).size, 0);

assert.equal(normalizeOfficeFormulaShortcut("Ctrl+Alt+I"), "Ctrl+Alt+I");
assert.equal(normalizeOfficeFormulaShortcut("Alt+Shift+F12"), "Alt+Shift+F12");
assert.equal(normalizeOfficeFormulaShortcut("Shift+I"), null);
assert.equal(normalizeOfficeFormulaShortcut("Ctrl+Meta+I"), null);
assert.equal(normalizeOfficeFormulaShortcut("Ctrl+Alt+F25"), null);

const wordSettings = normalizeWordFormulaShortcutSettings({
  inline: { shortcut: "Ctrl+Alt+I", format: "omml" },
  display: { shortcut: "Ctrl+Alt+D", format: "ole" },
});
assert.deepEqual(wordSettings, {
  inline: { shortcut: "Ctrl+Alt+I", format: "omml" },
  display: { shortcut: "Ctrl+Alt+D", format: "ole" },
});

const powerPointSettings = normalizePowerPointFormulaShortcutSettings({
  svg: "Ctrl+Alt+I",
  ole: "Ctrl+Alt+P",
  inlineOmml: "Ctrl+Alt+M",
  displayOmml: "Ctrl+Alt+D",
});
assert.equal(
  powerPointSettings.svg,
  wordSettings.inline.shortcut,
  "separate Office hosts must be allowed to reuse a shortcut",
);

assert.deepEqual(
  normalizeWordFormulaShortcutSettings({
    inline: { shortcut: "Ctrl+Alt+I", format: "ole" },
    display: { shortcut: "Ctrl+Alt+I", format: "omml" },
  }),
  {
    inline: { shortcut: "Ctrl+Alt+I", format: "ole" },
    display: { shortcut: null, format: "omml" },
  },
  "two actions inside one host must not retain the same shortcut",
);

assert.deepEqual(
  normalizePowerPointFormulaShortcutSettings({
    svg: "Ctrl+Alt+S",
    ole: "Ctrl+Alt+S",
    inlineOmml: "Ctrl+Alt+I",
    displayOmml: "Ctrl+Alt+D",
  }),
  {
    svg: "Ctrl+Alt+S",
    ole: null,
    inlineOmml: "Ctrl+Alt+I",
    displayOmml: "Ctrl+Alt+D",
  },
  "PowerPoint's four real Ribbon actions must be conflict-free",
);

assert.deepEqual(
  normalizePowerPointFormulaShortcutSettings({
    inline: { shortcut: "Ctrl+Alt+P", format: "ole" },
    display: { shortcut: "Ctrl+Alt+D", format: "omml" },
  }),
  {
    svg: null,
    ole: "Ctrl+Alt+P",
    inlineOmml: null,
    displayOmml: "Ctrl+Alt+D",
  },
  "legacy PowerPoint settings must migrate by real action semantics",
);

assert.deepEqual(
  officeFormulaShortcutFromEvent({
    code: "KeyI",
    ctrlKey: true,
    altKey: true,
    shiftKey: false,
    metaKey: false,
  }),
  { shortcut: "Ctrl+Alt+I" },
);
assert.deepEqual(
  officeFormulaShortcutFromEvent({
    code: "KeyO",
    ctrlKey: true,
    altKey: true,
    shiftKey: false,
    metaKey: false,
  }),
  { error: "reserved" },
);
assert.deepEqual(
  officeFormulaShortcutFromEvent({
    code: "KeyD",
    ctrlKey: false,
    altKey: false,
    shiftKey: true,
    metaKey: false,
  }),
  { error: "modifier" },
);

const readSource = (relativePath: string) =>
  readFileSync(new URL(`../${relativePath}`, import.meta.url), "utf8");

const sharedHook = readSource(
  "src-windows/VisualTeX.WindowsOffice.VstoShared/OfficeFormulaShortcutKeyboardHook.cs",
);
const powerPointAddIn = readSource(
  "src-windows/VisualTeX.PowerPointVsto/ThisAddIn.cs",
);
const settingsUi = readSource("src/components/OfficeFormulaShortcutSettings.tsx");
const lifecycle = readSource("src-tauri/src/office/lifecycle.rs");
const windowsBackend = readSource("src-tauri/src/office/windows_backend.rs");

assert.match(sharedHook, /PowerPoint,/);
assert.match(sharedHook, /: "PowerPoint";/);
assert.match(sharedHook, /valuePrefix \+ "SvgFormulaShortcut"/);
assert.match(sharedHook, /valuePrefix \+ "InlineOmmlFormulaShortcut"/);
assert.match(powerPointAddIn, /OnPowerPointFormulaShortcut/);
assert.match(powerPointAddIn, /OfficeFormulaShortcutHost\.PowerPoint/);
assert.match(powerPointAddIn, /PowerPointSvg/);
assert.match(powerPointAddIn, /OnNewSvgFormula\(this\)/);
assert.match(settingsUi, /host: "word"/);
assert.match(settingsUi, /host: "powerPoint"/);
assert.match(settingsUi, /action: "svg"/);
assert.match(settingsUi, /action: "ole"/);
assert.match(settingsUi, /action: "inlineOmml"/);
assert.match(settingsUi, /action: "displayOmml"/);
assert.match(lifecycle, /\/settings\/powerPointFormulaShortcuts/);
assert.match(windowsBackend, /PowerPointSvgFormulaShortcut/);
assert.match(windowsBackend, /PowerPointOleFormulaShortcut/);
assert.match(windowsBackend, /PowerPointInlineOmmlFormulaShortcut/);
assert.match(windowsBackend, /PowerPointDisplayOmmlFormulaShortcut/);

console.log("Office formula shortcut regression passed.");
