import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const source = await readFile(
  new URL(
    "../src-windows/VisualTeX.PowerPointVsto/PowerPointDoubleClickHook.cs",
    import.meta.url,
  ),
  "utf8",
);

assert.match(
  source,
  /var hasPreviousClick = previous != 0;/,
  "PowerPoint double-click hook must distinguish the first click from a real previous click",
);
assert.match(
  source,
  /Math\.Abs\(\(long\)input\.Pt\.X - _lastClickX\)/,
  "PowerPoint X coordinate distance must use Int64 arithmetic",
);
assert.match(
  source,
  /Math\.Abs\(\(long\)input\.Pt\.Y - _lastClickY\)/,
  "PowerPoint Y coordinate distance must use Int64 arithmetic",
);
assert.doesNotMatch(
  source,
  /Math\.Abs\(input\.Pt\.X - _lastClickX\)/,
  "PowerPoint hook must not reintroduce the int.MinValue subtraction overflow",
);
assert.doesNotMatch(
  source,
  /Math\.Abs\(input\.Pt\.Y - _lastClickY\)/,
  "PowerPoint hook must not reintroduce the int.MinValue subtraction overflow",
);

console.log("VisualTeX PowerPoint double-click hook safety regression passed");
