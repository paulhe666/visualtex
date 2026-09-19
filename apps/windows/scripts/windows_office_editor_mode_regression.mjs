import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const read = async (path) => (await readFile(path, "utf8")).replace(/\r\n?/g, "\n");

const officeDialog = await read("src/office/dialog/OfficeDialogApp.tsx");
const workspace = await read("src/workspace/EditorWorkspace.tsx");
const mathEditor = await read("src/editor/MathEditor.tsx");

assert.ok(
  !officeDialog.includes('className="office-display-mode-setting"'),
  "Office editor must not expose a second inline/display selector",
);
assert.ok(
  workspace.includes("showLineModeControls={!isOfficeWorkspace}"),
  "Office workspaces must disable per-row line-mode controls",
);
assert.ok(
  mathEditor.includes("showLineModeControls = true"),
  "Desktop MathEditor must keep line-mode controls enabled by default",
);
assert.ok(
  mathEditor.includes("!interactionReadOnly && showLineModeControls"),
  "Per-row $/$$ controls must be gated by showLineModeControls",
);
assert.ok(
  mathEditor.includes("showLineModeControls &&\n        event.key === \"Enter\""),
  "Ctrl/Alt+Enter forced row-mode switching must be disabled in Office",
);
assert.ok(
  officeDialog.includes("displayMode,") &&
    officeDialog.includes("numbered: displayMode === \"block\" && numbered"),
  "Office session displayMode must still drive the actual Office commit",
);

console.log(
  "Windows Office editor mode regression passed: session displayMode retained; duplicate inline/display UI removed.",
);
