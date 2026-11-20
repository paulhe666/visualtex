import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

// The PowerPoint hook pairs presses through the shared OfficeDoubleClickPairing (VstoShared), which holds the
// safety rules this regression guards.
const hook = await readFile(
  new URL("../src-windows/VisualTeX.PowerPointVsto/PowerPointDoubleClickHook.cs", import.meta.url),
  "utf8",
);
const pairing = await readFile(
  new URL("../src-windows/VisualTeX.WindowsOffice.VstoShared/OfficeDoubleClickGesture.cs", import.meta.url),
  "utf8",
);

assert.match(hook, /OfficeDoubleClickPairing _pairing = new\(\);/, "PowerPoint must pair presses through OfficeDoubleClickPairing");
assert.match(
  pairing,
  /var paired = _hasPrevious\b/,
  "the pairing must distinguish the first click from a real previous click",
);
assert.match(pairing, /Math\.Abs\(\(long\)screenX - _previousX\)/, "the X distance must use Int64 arithmetic");
assert.match(pairing, /Math\.Abs\(\(long\)screenY - _previousY\)/, "the Y distance must use Int64 arithmetic");
assert.doesNotMatch(
  pairing,
  /Math\.Abs\(screenX - _previousX\)/,
  "the pairing must not reintroduce the int.MinValue subtraction overflow",
);
assert.doesNotMatch(
  pairing,
  /Math\.Abs\(screenY - _previousY\)/,
  "the pairing must not reintroduce the int.MinValue subtraction overflow",
);

console.log("VisualTeX PowerPoint double-click hook safety regression passed");
