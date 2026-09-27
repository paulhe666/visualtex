import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
  normalizeOfficeFormulaShortcut,
  normalizeOfficeFormulaShortcutSettings,
  officeFormulaShortcutFromEvent,
} from "../src/office/officeFormulaShortcuts";

assert.equal(normalizeOfficeFormulaShortcut("Ctrl+Alt+I"), "Ctrl+Alt+I");
assert.equal(normalizeOfficeFormulaShortcut("Alt+Shift+F12"), "Alt+Shift+F12");
assert.equal(normalizeOfficeFormulaShortcut("Shift+I"), null);
assert.equal(normalizeOfficeFormulaShortcut("Ctrl+Meta+I"), null);
assert.equal(normalizeOfficeFormulaShortcut("Ctrl+Alt+F25"), null);

const wordSettings = normalizeOfficeFormulaShortcutSettings({
  inline: { shortcut: "Ctrl+Alt+I", format: "omml" },
  display: { shortcut: "Ctrl+Alt+D", format: "ole" },
});
assert.deepEqual(wordSettings, {
  inline: { shortcut: "Ctrl+Alt+I", format: "omml" },
  display: { shortcut: "Ctrl+Alt+D", format: "ole" },
});

const powerPointSettings = normalizeOfficeFormulaShortcutSettings({
  inline: { shortcut: "Ctrl+Alt+I", format: "ole" },
  display: { shortcut: "Ctrl+Alt+P", format: "omml" },
});
assert.equal(
  powerPointSettings.inline.shortcut,
  wordSettings.inline.shortcut,
  "separate Office hosts must be allowed to reuse a shortcut",
);

assert.deepEqual(
  normalizeOfficeFormulaShortcutSettings({
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
assert.match(sharedHook, /valuePrefix \+ "InlineFormulaShortcut"/);
assert.match(powerPointAddIn, /OnPowerPointFormulaShortcut/);
assert.match(powerPointAddIn, /OfficeFormulaShortcutHost\.PowerPoint/);
assert.match(settingsUi, /host: "word"/);
assert.match(settingsUi, /host: "powerPoint"/);
assert.match(settingsUi, /onOpenFormulaHotkeys/);
assert.match(lifecycle, /\/settings\/powerPointFormulaShortcuts/);
assert.match(windowsBackend, /PowerPointInlineFormulaShortcut/);
assert.match(windowsBackend, /PowerPointDisplayFormulaShortcut/);

console.log("Office formula shortcut regression passed.");
