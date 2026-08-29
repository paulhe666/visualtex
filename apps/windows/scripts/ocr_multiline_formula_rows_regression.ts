import assert from "node:assert/strict";
import {
  normalizeOcrFormulaLines,
  splitOcrLatexIntoFormulaLines,
} from "../src/ocr/ocrService.ts";
import { VISUALTEX_ALIGNMENT_MARKER_LATEX } from "../src/editor/alignmentMarkers.ts";

assert.deepEqual(
  splitOcrLatexIntoFormulaLines(
    String.raw`\begin{aligned}x&=1\\y&=\frac{2}{3}\\[6pt]z&=4\end{aligned}`,
  ),
  [
    `x${VISUALTEX_ALIGNMENT_MARKER_LATEX}=1`,
    `y${VISUALTEX_ALIGNMENT_MARKER_LATEX}=\\frac{2}{3}`,
    `z${VISUALTEX_ALIGNMENT_MARKER_LATEX}=4`,
  ],
  "aligned OCR output must become independent rows while preserving the shared alignment point",
);

assert.deepEqual(
  splitOcrLatexIntoFormulaLines(
    String.raw`\[a+b\\c+d\]`,
  ),
  ["a+b", "c+d"],
  "outer display delimiters must not keep multiple OCR rows in one formula line",
);

assert.deepEqual(
  normalizeOcrFormulaLines([
    { latex: String.raw`\begin{gathered}p=1\\q=2\end{gathered}` },
    { latex: "r=3" },
  ]),
  ["p=1", "q=2", "r=3"],
  "provider formula arrays and multiline environments must flatten in visual order",
);

assert.deepEqual(
  splitOcrLatexIntoFormulaLines(
    String.raw`A=\begin{pmatrix}1&2\\3&4\end{pmatrix}`,
  ),
  [String.raw`A=\begin{pmatrix}1&2\\3&4\end{pmatrix}`],
  "matrix rows are one mathematical structure and must remain one editor formula line",
);

assert.deepEqual(
  splitOcrLatexIntoFormulaLines(
    String.raw`f(x)=\begin{cases}x,&x>0\\-x,&x\le0\end{cases}`,
  ),
  [String.raw`f(x)=\begin{cases}x,&x>0\\-x,&x\le0\end{cases}`],
  "cases rows must remain inside one editable cases structure",
);

assert.deepEqual(
  splitOcrLatexIntoFormulaLines("$x=1$\n$y=2$"),
  ["x=1", "y=2"],
  "separate physical OCR rows with inline delimiters must remain separate rows",
);

console.log("OCR multiline formula-row normalization passed.");
