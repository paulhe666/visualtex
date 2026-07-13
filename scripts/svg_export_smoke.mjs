import assert from "node:assert/strict";
import { latexToSvg, svgToBase64 } from "../src/export/runtime.ts";

const matrixRows = Array.from({ length: 10 }, (_, row) =>
  Array.from({ length: 10 }, (_, column) => `a_{${row + 1}${column + 1}}`).join("&"),
).join("\\\\");

const cases = [
  ["fraction", String.raw`\frac{a+b}{c+d}`],
  ["root", String.raw`\sqrt[n]{x^2+y^2}`],
  ["integral", String.raw`\int_{-\infty}^{\infty} e^{-x^2}\,\mathrm{d}x`],
  ["sum", String.raw`\sum_{i=1}^{n} i^2`],
  ["matrix", String.raw`\begin{pmatrix}${matrixRows}\end{pmatrix}`],
  ["chinese", String.raw`\text{测试}+\alpha`],
  ["multiline", "a=b+c\nd=e-f\ng=h"],
  ["long", Array.from({ length: 25 }, (_, index) => `x_{${index + 1}}`).join("+")],
];

for (const [name, latex] of cases) {
  const result = await latexToSvg(latex, {
    displayMode: true,
    fontSizePt: 14,
    paddingPx: 10,
    background: name === "root" ? "white" : "transparent",
  });
  assert.match(result.svg, /^<svg\b/);
  assert.match(result.svg, /\bviewBox=/);
  assert.ok(result.width > 0, `${name} width`);
  assert.ok(result.height > 0, `${name} height`);
  assert.ok((result.baseline ?? -1) >= 0, `${name} baseline`);
  assert.ok(!/<foreignObject\b/i.test(result.svg), `${name} foreignObject`);
  assert.ok(!/<link\b|@import\b/i.test(result.svg), `${name} external CSS`);
  assert.ok(
    !/\b(?:href|xlink:href)=["'](?!#|data:)[^"']+/i.test(result.svg),
    `${name} external href`,
  );
  assert.ok(!/url\(\s*["']?https?:/i.test(result.svg), `${name} remote CSS URL`);
  assert.equal(result.base64, svgToBase64(result.svg));
  const decoded = new TextDecoder().decode(
    Uint8Array.from(atob(result.base64), (character) => character.charCodeAt(0)),
  );
  assert.equal(decoded, result.svg, `${name} UTF-8 base64 round trip`);
}

assert.throws(
  () =>
    latexToSvg("", {
      displayMode: true,
      fontSizePt: 12,
      paddingPx: 8,
      background: "transparent",
    }),
  /Cannot export an empty formula/,
);

console.log(`SVG export smoke test passed (${cases.length} formula classes)`);
