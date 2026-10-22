import assert from "node:assert/strict";
import { readFileSync, readdirSync } from "node:fs";
import { formulaHtml } from "../../shared/tex-layout/src/painter.ts";
import { hpack } from "../../shared/tex-layout/src/pack.ts";

// The runtime-safety changes are written into VisualTeX's editing kernel
// (apps/shared/visualtex-core/src/kernel); its source must carry them.
const kernelDirectory = new URL("../../shared/visualtex-core/src/kernel/", import.meta.url);
const source = (readdirSync(kernelDirectory, { recursive: true }) as string[])
  .filter((name) => name.endsWith(".ts"))
  .map((name) => readFileSync(new URL(name, kernelDirectory), "utf8"))
  .join("\n");
assert.match(source, /mathlive: '0\.109\.2'/);
// Formula markup comes from the layout engine; MathLive's Box renderer (whose
// unescaped text body this used to fix) is gone from the kernel.
assert.doesNotMatch(source, /class Box\b/);
// MathML text is escaped, ampersands first.
assert.match(
  source,
  /function xmlEscape\(string: string\): string \{\s*return \(\s*string\s*\.replace\(\/&\/g, '&amp;'\)\s*\.replace\(\/"\/g, '&quot;'\)\s*\.replace\(\/'\/g, '&#39;'\)/,
);
assert.match(
  source,
  /const visualTexRootFirstChild = this\.model\.root\.firstChild;/,
);
assert.doesNotMatch(
  source,
  /this\.model\.root\.firstChild\.mode = mode;/,
);
assert.match(source, /xmlEscape\(atom\.value\)/);
assert.match(source, /xmlEscape\(mathML\)/);

// Text the engine puts into the markup (raw LaTeX input, symbols) is escaped.
const attack = '<img src=x onerror="globalThis.__probe=1"> & \'quoted\'';
const engineMarkup = formulaHtml(hpack([]), { glyphs: [], rules: [], spans: [], boxes: [] }, {
  fontFamily: () => "serif",
  elements: [{ x0: 0, y0: 0, x1: 65536, y1: 65536, text: attack }],
});
assert.doesNotMatch(engineMarkup, /<img/);
assert.match(engineMarkup, /&lt;img src=x onerror=&quot;globalThis\.__probe=1&quot;> &amp; 'quoted'/);

console.log("VisualTeX MathLive runtime safety regression passed");
