import assert from "node:assert/strict";
import {
  normalizeWordFormulaShortcut,
  normalizeWordFormulaShortcutSettings,
  wordFormulaShortcutFromEvent,
} from "../src/office/wordFormulaShortcuts";

assert.equal(normalizeWordFormulaShortcut("Ctrl+Alt+I"), "Ctrl+Alt+I");
assert.equal(normalizeWordFormulaShortcut("Alt+Shift+F12"), "Alt+Shift+F12");
assert.equal(normalizeWordFormulaShortcut("Shift+I"), null);
assert.equal(normalizeWordFormulaShortcut("Ctrl+Meta+I"), null);
assert.equal(normalizeWordFormulaShortcut("Ctrl+Alt+F25"), null);

assert.deepEqual(
  normalizeWordFormulaShortcutSettings({
    inline: { shortcut: "Ctrl+Alt+I", format: "omml" },
    display: { shortcut: "Ctrl+Alt+D", format: "ole" },
  }),
  {
    inline: { shortcut: "Ctrl+Alt+I", format: "omml" },
    display: { shortcut: "Ctrl+Alt+D", format: "ole" },
  },
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
);

assert.deepEqual(
  wordFormulaShortcutFromEvent({
    code: "KeyI",
    ctrlKey: true,
    altKey: true,
    shiftKey: false,
    metaKey: false,
  }),
  { shortcut: "Ctrl+Alt+I" },
);
assert.deepEqual(
  wordFormulaShortcutFromEvent({
    code: "KeyO",
    ctrlKey: true,
    altKey: true,
    shiftKey: false,
    metaKey: false,
  }),
  { error: "reserved" },
);
assert.deepEqual(
  wordFormulaShortcutFromEvent({
    code: "KeyD",
    ctrlKey: false,
    altKey: false,
    shiftKey: true,
    metaKey: false,
  }),
  { error: "modifier" },
);

console.log("Word formula shortcut regression passed.");
