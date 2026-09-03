import assert from "node:assert/strict";
import { latexToSvg } from "../src/export/latexToSvg.ts";
import {
  calculateInlineFormulaPosition,
  calculateInlineSessionPosition,
} from "../src/office/adapters/WordAdapter.ts";

// Coverage representatives only. Production alignment receives rendered
// geometry, never a formula string or structure category.
const formulas = [
  String.raw`x`,
  String.raw`L`,
  String.raw`L_z`,
  String.raw`L^2`,
  String.raw`L_z^2`,
  String.raw`L_zL^2`,
  String.raw`\frac{1}{2}`,
  String.raw`\cfrac{1}{1+\cfrac{1}{x}}`,
  String.raw`\sqrt{x}`,
  String.raw`\sqrt[3]{\frac{x+1}{y-1}}`,
  String.raw`\int_a^b c\,\mathrm{d}e`,
  String.raw`\sum_{n=1}^{\infty}\frac{1}{n^2}`,
  String.raw`\left(\begin{matrix}a&b\\c&d\end{matrix}\right)`,
  String.raw`\begin{cases}x^2,&x\ge 0\\-x,&x<0\end{cases}`,
  String.raw`\overbrace{a+b+\cdots+z}^{26\text{ terms}}`,
  String.raw`\underbrace{x_1+\cdots+x_n}_{n\text{ terms}}`,
  String.raw`\hat{x}+\vec{y}+\overline{AB}`,
  String.raw`\text{速度 }v=\frac{\mathrm{d}x}{\mathrm{d}t}`,
];
const fontSizesPt = [8, 10.5, 12, 14, 18, 24, 42, 72];

const results = [];
for (const fontSizePt of fontSizesPt) {
  for (const latex of formulas) {
    const exported = latexToSvg(latex, {
      displayMode: false,
      fontSizePt,
      paddingPx: 1,
      background: "transparent",
    });
    assert.ok(exported.width > 0 && Number.isFinite(exported.width));
    assert.ok(exported.height > 0 && Number.isFinite(exported.height));
    assert.ok(
      exported.baseline >= 0 && exported.baseline < exported.height,
      `${latex} @ ${fontSizePt}pt has an invalid mathematical baseline`,
    );

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
    const expectedMagnitude = Math.floor(descentPt + 0.0101);
    const expectedPosition = expectedMagnitude === 0 ? 0 : -expectedMagnitude;
    assert.equal(
      position,
      expectedPosition,
      `${latex} @ ${fontSizePt}pt did not apply the complete whole-point part of its rendered descent`,
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
    assert.equal(
      sessionPosition,
      position,
      `${latex} @ ${fontSizePt}pt: session and picture offsets differ`,
    );
    assert.ok(position <= 0, `${latex}: Word baseline must never be raised`);

    // Word's safe object-model setter stores this placement in whole points.
    // Visual parity tests show that lowering by the complete integer portion is
    // substantially closer to OMML than rounding or fixed optical lifts. The
    // remaining error is therefore the sub-point descent intentionally left above
    // the text baseline, except for the 0.01 pt near-integer float snap.
    const baselineErrorPt = -position - descentPt;
    assert.ok(
      baselineErrorPt >= -1.000_001 && baselineErrorPt <= 0.010_201,
      `${latex} @ ${fontSizePt}pt: baseline remainder ${baselineErrorPt.toFixed(4)}pt is outside the whole-point floor contract`,
    );

    results.push({
      latex,
      fontSizePt,
      heightPt: Number(actualHeightPt.toFixed(3)),
      descentPt: Number(descentPt.toFixed(3)),
      positionPt: position,
      baselineErrorPt: Number(baselineErrorPt.toFixed(3)),
    });
  }
}

const plainX = results.filter((row) => row.latex === "x");
assert.ok(
  plainX.some((row) => row.positionPt < 0),
  "Small ordinary symbols must not all collapse to Position=0 when their rendered descent rounds to one point",
);

console.table(results.filter((row) => row.fontSizePt === 10.5));
console.log(
  `Word baseline metric regression passed for ${results.length} combinations. `
  + "This verifies rendered geometry and whole-point floor quantisation only; "
  + "OMML visual parity is verified separately in real Word.",
);
