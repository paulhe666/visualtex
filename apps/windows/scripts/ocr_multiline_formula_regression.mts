import assert from "node:assert/strict";
import {
  prepareOcrLatexForEditor,
  splitOcrLatexIntoFormulaLines,
} from "../src/ocr/multilineFormula.ts";
import { expandFormulaLinesForPendingOcr } from "../src/ocr/multilineFormulaLineExpansion.ts";
import { VISUALTEX_ALIGNMENT_MARKER_LATEX } from "../src/editor/alignmentMarkers.ts";

const aligned = String.raw`\[\begin{aligned}x&=1\\y&=2\\z&=3\end{aligned}\]`;
const alignedRows = ["x", "y", "z"].map(
  (name, index) => `${name}${VISUALTEX_ALIGNMENT_MARKER_LATEX}=${index + 1}`,
);
assert.deepEqual(splitOcrLatexIntoFormulaLines(aligned), alignedRows);
assert.equal(prepareOcrLatexForEditor(aligned), alignedRows.join("\n"));
assert.deepEqual(
  splitOcrLatexIntoFormulaLines("x=1\ny=2\r\nz=3"),
  ["x=1", "y=2", "z=3"],
);
assert.deepEqual(
  splitOcrLatexIntoFormulaLines(
    String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}`,
  ),
  [String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}`],
);
assert.deepEqual(splitOcrLatexIntoFormulaLines(String.raw`$E=mc^2$`), ["E=mc^2"]);

const expanded = expandFormulaLinesForPendingOcr(
  [
    { id: "before", latex: "a" },
    { id: "target", latex: "prefix x=1\ny=2\nz=3 suffix", alignment: "left" },
    { id: "after", latex: "b" },
  ],
  {
    joinedLatex: "x=1\ny=2\nz=3",
    lines: ["x=1", "y=2", "z=3"],
    expiresAt: Date.now() + 1_000,
  },
);
assert.ok(expanded);
assert.equal(expanded?.length, 5);
assert.equal(expanded?.[1].id, "target");
assert.equal(expanded?.[1].latex, "prefix x=1");
assert.equal(expanded?.[2].latex, "y=2");
assert.equal(expanded?.[3].latex, "z=3 suffix");
assert.equal(expanded?.[2].alignment, "left");
assert.notEqual(expanded?.[2].id, "target");

console.log("OCR multi-line FormulaLine splitting regression passed.");
