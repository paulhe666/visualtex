import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
  escapeMathLiveTextForMarkup,
  escapeMathLiveTextForXml,
  patchVisualTexMathLiveRuntimeSafety,
} from "../vite.mathliveRuntimeSafety.ts";

const source = readFileSync(
  new URL("../node_modules/mathlive/mathlive.mjs", import.meta.url),
  "utf8",
);
assert.match(source, /MathLive 0\.109\.2/);

const patched = patchVisualTexMathLiveRuntimeSafety(source);
assert.notEqual(patched, source);
assert.match(
  patched,
  /let body = this\.value \? visualTexEscapeMathLiveText\(this\.value\) : "";/,
);
assert.match(
  patched,
  /function visualTexEscapeMathLiveText\(value\)/,
);
assert.doesNotMatch(
  patched,
  /let body = \(_a3 = this\.value\) != null \? _a3 : "";/,
);
assert.match(
  patched,
  /return string\.replace\(\/&\/g, "&amp;"\).*"&quot;".*"&#39;"/,
);
assert.match(
  patched,
  /const visualTexRootFirstChild = this\.model\.root\.firstChild;/,
);
assert.doesNotMatch(
  patched,
  /this\.model\.root\.firstChild\.mode = mode;/,
);
assert.match(patched, /xmlEscape\(atom\.value\)/);
assert.match(patched, /xmlEscape\(mathML\)/);

const attack = '<img src=x onerror="globalThis.__probe=1"> & \'quoted\'';
assert.equal(
  escapeMathLiveTextForMarkup(attack),
  '&lt;img src=x onerror="globalThis.__probe=1"&gt; &amp; \'quoted\'',
);
assert.equal(
  escapeMathLiveTextForXml(attack),
  "&lt;img src=x onerror=&quot;globalThis.__probe=1&quot;&gt; &amp; &#39;quoted&#39;",
);
assert.throws(
  () => patchVisualTexMathLiveRuntimeSafety("unrelated source"),
  /patch anchor changed/,
  "a changed upstream source must fail closed",
);

console.log("VisualTeX MathLive runtime safety patch regression passed");
