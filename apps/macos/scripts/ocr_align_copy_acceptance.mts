import assert from "node:assert/strict";
import {
  formatLatexLines,
  parseLatexSource,
} from "../src/clipboard/LatexCopyService.ts";
import { VISUALTEX_ALIGNMENT_MARKER_LATEX } from "../src/editor/alignmentMarkers.ts";

const rows = [
  `x${VISUALTEX_ALIGNMENT_MARKER_LATEX}=a+b+c${VISUALTEX_ALIGNMENT_MARKER_LATEX}${VISUALTEX_ALIGNMENT_MARKER_LATEX}u`,
  `y${VISUALTEX_ALIGNMENT_MARKER_LATEX}=\\frac{p}{q}${VISUALTEX_ALIGNMENT_MARKER_LATEX}${VISUALTEX_ALIGNMENT_MARKER_LATEX}v`,
  `z${VISUALTEX_ALIGNMENT_MARKER_LATEX}=r${VISUALTEX_ALIGNMENT_MARKER_LATEX}${VISUALTEX_ALIGNMENT_MARKER_LATEX}w`,
];

const copied = formatLatexLines(rows, "aligned");
assert.match(copied, /^\\\[/);
assert.match(copied, /\\begin\{aligned\}/);
assert.match(copied, /x&=a\+b\+c&&u\s*\\\\/);
assert.match(copied, /y&=\\frac\{p\}\{q\}&&v\s*\\\\/);
assert.match(copied, /z&=r&&w/);
assert.match(copied, /\\end\{aligned\}/);
assert.match(copied, /\\\]$/);
assert.doesNotMatch(copied, /visualtex-align-marker|\\class\{/);
assert.deepEqual(
  parseLatexSource(copied, "aligned"),
  rows,
  "Copying aligned OCR FormulaLines and parsing them again must restore the exact explicit alignment markers.",
);
assert.equal((copied.match(/&/g) ?? []).length, 9);

console.log(
  "macOS OCR aligned copy acceptance passed: FormulaLines serialize as one aligned environment with real '&' anchors and parse back to the exact VisualTeX marker rows.",
);
