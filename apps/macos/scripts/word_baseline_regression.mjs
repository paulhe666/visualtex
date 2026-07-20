import assert from "node:assert/strict";
import { latexToSvg } from "../src/export/latexToSvg.ts";
import {
  calculateInlineFormulaPosition,
  calculateInlineSessionPosition,
} from "../src/office/adapters/WordAdapter.ts";

const formulas = [
  String.raw`x`,
  String.raw`x_i`,
  String.raw`x^2+y^2`,
  String.raw`\alpha+\beta=\gamma`,
  String.raw`\frac{x+1}{y-1}`,
  String.raw`\sqrt{x^2+y^2}`,
  String.raw`\sum_{n=1}^{\infty}\frac{1}{n^2}`,
  String.raw`\int_0^1 x^2\,\mathrm{d}x`,
  String.raw`\left(\begin{matrix}a&b\\c&d\end{matrix}\right)`,
  String.raw`A_{i_1i_2\cdots i_n}=B^{j_1j_2\cdots j_m}`,
];

const results = formulas.map((latex) => {
  const exported = latexToSvg(latex, {
    displayMode: false,
    fontSizePt: 14,
    paddingPx: 1,
    background: "transparent",
  });
  const naturalWidthPt = exported.width * 0.75;
  const naturalHeightPt = exported.height * 0.75;
  const scale = Math.min(1, 500 / naturalWidthPt);
  const actualHeightPt = naturalHeightPt * scale;
  const descentPt =
    actualHeightPt * ((exported.height - exported.baseline) / exported.height);
  const position = calculateInlineFormulaPosition(
    actualHeightPt,
    exported.height,
    exported.baseline,
  );
  const sessionPosition = calculateInlineSessionPosition({
    exportWidth: exported.width,
    exportHeight: exported.height,
    exportResult: {
      svg: exported.svg,
      svgBase64: exported.base64,
      width: exported.width,
      height: exported.height,
      baseline: exported.baseline,
    },
  });

  assert.equal(sessionPosition, position, `${latex}: session and picture offsets differ`);
  assert.ok(position <= 0, `${latex}: Word baseline must never be raised`);
  // Word's native Font.Position is integer-valued. Rounding the exact SVG
  // descent is therefore allowed to leave at most half a point of residual.
  const residualPt = -position - descentPt;
  assert.ok(
    Math.abs(residualPt) <= 0.500_001,
    `${latex}: baseline residual ${residualPt.toFixed(4)}pt exceeds rounding tolerance`,
  );

  return {
    latex,
    heightPt: Number(actualHeightPt.toFixed(3)),
    descentPt: Number(descentPt.toFixed(3)),
    positionPt: position,
    residualPt: Number(residualPt.toFixed(3)),
  };
});

console.table(results);
console.log(`Word baseline regression passed for ${results.length} formula structures.`);
