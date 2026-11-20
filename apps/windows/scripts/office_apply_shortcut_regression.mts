import assert from "node:assert/strict";
import { isOfficeApplyShortcut } from "../../shared/visualtex-app/src/office/dialog/officeApplyShortcut.ts";

const base = {
  key: "s",
  code: "KeyS",
  ctrlKey: true,
  altKey: false,
  shiftKey: false,
  metaKey: false,
  isComposing: false,
};

assert.equal(isOfficeApplyShortcut(base), true);
assert.equal(isOfficeApplyShortcut({ ...base, key: "S" }), true);
assert.equal(isOfficeApplyShortcut({ ...base, key: "ы", code: "KeyS" }), true);
assert.equal(isOfficeApplyShortcut({ ...base, ctrlKey: false }), false);
assert.equal(isOfficeApplyShortcut({ ...base, altKey: true }), false);
assert.equal(isOfficeApplyShortcut({ ...base, shiftKey: true }), false);
assert.equal(isOfficeApplyShortcut({ ...base, metaKey: true }), false);
assert.equal(isOfficeApplyShortcut({ ...base, isComposing: true }), false);
assert.equal(isOfficeApplyShortcut({ ...base, keyCode: 229 }), false);
assert.equal(isOfficeApplyShortcut({ ...base, key: "a", code: "KeyA" }), false);

console.log("Office Ctrl+S shortcut filter regression passed");
