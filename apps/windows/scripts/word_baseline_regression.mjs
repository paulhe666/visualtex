import assert from "node:assert/strict";
import { latexToSvg } from "../src/export/latexToSvg.ts";
import {
  WORD_OMML_INLINE_MINIMUM_ASCENT_EM,
  WORD_OMML_INLINE_MINIMUM_DESCENT_EM,
} from "../src/export/wordOmmlInlineFrame.ts";
import {
  calculateInlineFormulaPosition,
  calculateInlineSessionPosition,
} from "../src/office/adapters/WordAdapter.ts";

// These expressions are coverage representatives only. The production path
// receives no formula name/type and applies the same metric calculation to all
// input. Keep this list structurally varied to guard that invariant.
const formulas = [
  String.raw`x`,
  String.raw`L`,
  String.raw`x_i`,
  String.raw`x^i`,
  String.raw`x_i^j`,
  String.raw`a=\frac{dv}{dt}`,
  String.raw`\cfrac{1}{1+\cfrac{1}{x}}`,
  String.raw`x^2+y^2`,
  String.raw`\alpha+\beta=\gamma`,
  String.raw`\frac{x+1}{y-1}`,
  String.raw`\sqrt{x^2+y^2}`,
  String.raw`\sqrt[3]{\frac{x+1}{y-1}}`,
  String.raw`\sum_{n=1}^{\infty}\frac{1}{n^2}`,
  String.raw`\int_0^1 x^2\,\mathrm{d}x`,
  String.raw`\left(\begin{matrix}a&b\\c&d\end{matrix}\right)`,
  String.raw`\begin{cases}x^2,&x\ge 0\\-x,&x<0\end{cases}`,
  String.raw`\overbrace{a+b+\cdots+z}^{26\text{ terms}}`,
  String.raw`A_{i_1i_2\cdots i_n}=B^{j_1j_2\cdots j_m}`,
];

const fontSizesPt = [8, 10.5, 12, 14, 18, 24, 42, 72];

function wordOmmlInlineFrame(fontSizePt) {
  return {
    displayMode: false,
    fontSizePt,
    paddingPx: 1,
    paddingYPx: 0,
    minimumAscentEm: WORD_OMML_INLINE_MINIMUM_ASCENT_EM,
    minimumDescentEm: WORD_OMML_INLINE_MINIMUM_DESCENT_EM,
    background: "transparent",
  };
}

const lSubscript = latexToSvg(String.raw`L_z`, wordOmmlInlineFrame(10.5));
const lSuperscript = latexToSvg(String.raw`L^2`, wordOmmlInlineFrame(10.5));
const lBothScripts = latexToSvg(String.raw`L_zL^2`, wordOmmlInlineFrame(10.5));
for (const [label, exported] of [
  ["L_z", lSubscript],
  ["L^2", lSuperscript],
  ["L_zL^2", lBothScripts],
]) {
  const ascentEm = exported.baseline / (10.5 * 96 / 72);
  const descentEm = (exported.height - exported.baseline) / (10.5 * 96 / 72);
  assert.ok(
    ascentEm >= WORD_OMML_INLINE_MINIMUM_ASCENT_EM - 1e-6,
    `${label}: Word inline OLE ascent ${ascentEm}em is below the OMML frame`,
  );
  assert.ok(
    descentEm >= WORD_OMML_INLINE_MINIMUM_DESCENT_EM - 1e-6,
    `${label}: Word inline OLE descent ${descentEm}em is below the OMML frame`,
  );
}
assert.ok(
  Math.abs(lSubscript.height - lSuperscript.height) < 0.5,
  "Simple subscript and superscript formulas must share the OMML-like inline frame",
);
assert.ok(
  Math.abs(lBothScripts.height - lSuperscript.height) < 0.5,
  "Combining ordinary scripts must not create a content-dependent tall OLE box",
);

const results = [];
for (const fontSizePt of fontSizesPt) {
  for (const latex of formulas) {
    const tight = latexToSvg(latex, {
      displayMode: false,
      fontSizePt,
      paddingPx: 0,
      background: "transparent",
    });
    const exported = latexToSvg(latex, wordOmmlInlineFrame(fontSizePt));
    const emPx = fontSizePt * 96 / 72;
    const tightDescent = tight.height - tight.baseline;
    const expectedAscent = Math.max(
      tight.baseline,
      WORD_OMML_INLINE_MINIMUM_ASCENT_EM * emPx,
    );
    const expectedDescent = Math.max(
      tightDescent,
      WORD_OMML_INLINE_MINIMUM_DESCENT_EM * emPx,
    );

    assert.ok(
      Math.abs(exported.baseline - expectedAscent) <= 1e-6,
      `${latex} @ ${fontSizePt}pt: ascent was not derived from formula geometry`,
    );
    assert.ok(
      Math.abs((exported.height - exported.baseline) - expectedDescent) <= 1e-6,
      `${latex} @ ${fontSizePt}pt: descent was not derived from formula geometry`,
    );
    assert.ok(
      Math.abs(exported.width - (tight.width + 2)) <= 1e-6,
      `${latex} @ ${fontSizePt}pt: vertical framing changed formula width`,
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
      `${latex}: session and picture offsets differ`,
    );
    assert.ok(position <= 0, `${latex}: Word baseline must never be raised`);
    // Word's native Font.Position is integer-valued. VisualTeX intentionally
    // leaves the OLE glyph baseline one point above the raw geometric result so
    // it optically matches native OMML on the same Word text line.
    const residualPt = -position - descentPt;
    assert.ok(
      Math.abs(residualPt + 1) <= 0.500_001,
      `${latex}: optical baseline residual ${residualPt.toFixed(4)}pt is outside tolerance`,
    );

    results.push({
      latex,
      fontSizePt,
      heightPt: Number(actualHeightPt.toFixed(3)),
      descentPt: Number(descentPt.toFixed(3)),
      positionPt: position,
      residualPt: Number(residualPt.toFixed(3)),
    });
  }
}

console.table(results.filter((row) => row.fontSizePt === 10.5));
console.log(
  `Word baseline regression passed for ${formulas.length} formula structures `
  + `at ${fontSizesPt.length} font sizes (${results.length} metric combinations).`,
);
