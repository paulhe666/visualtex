import assert from "node:assert/strict";
import { latexToMathMl, latexToSvg } from "../src/export/runtime.ts";
import { assertResolvedPresentationMathMl, assertResolvedMathJaxSvg } from "../src/math/latexCompatibility.ts";

// Production parser/export regression; no browser, Word, installation or mock
// conversion implementation. Run with tsx in a dependency-installed checkout.
let checked = 0;
for (const latex of [
  String.raw`\ThisCommandDoesNotExist{x}`,
  String.raw`\frac{x}`,
  String.raw`\begin{unknownenvironment}x\end{unknownenvironment}`,
]) {
  await assert.rejects(() => latexToMathMl(latex, true), undefined, `MathML export accepted invalid TeX: ${latex}`);
  await assert.rejects(() => latexToSvg(latex), undefined, `SVG export accepted invalid TeX: ${latex}`);
  checked += 2;
  // A failed conversion must not poison the next request in the shared jax.
  assert.match(await latexToMathMl(String.raw`x^2+1`, false), /<msup>/);
  checked++;
}
for (const latex of [
  String.raw`\verb|C:\Users|`,
  String.raw`\verb|\alpha|`,
  String.raw`\textcolor{red}{\text{important}}`,
  String.raw`\begin{pmatrix}\sqrt{x}&y\\0&1\end{pmatrix}`,
  String.raw`\begin{cases}\sqrt{x}&x>0\\0&x\le0\end{cases}`,
  String.raw`{}^{14}C`,
  String.raw`{}_{6}^{14}C`,
  String.raw`\bm{x}+\frac{a}{b}`,
]) {
  const mathMl = await latexToMathMl(latex, true);
  assert.match(mathMl, /^<math\b/);
  assert.doesNotThrow(() => assertResolvedPresentationMathMl(mathMl));
  const result = await latexToSvg(latex);
  assert.doesNotThrow(() => assertResolvedMathJaxSvg(result.svg));
  checked += 2;
}
for (const xml of [
  '<math xmlns="http://www.w3.org/1998/Math/MathML"><merror><mtext>bad</mtext></merror></math>',
  '<p:math xmlns:p="http://www.w3.org/1998/Math/MathML"><p:merror><p:mtext>bad</p:mtext></p:merror></p:math>',
]) {
  assert.throws(() => assertResolvedPresentationMathMl(xml), /error node/);
  checked++;
}
assert.doesNotThrow(() => assertResolvedPresentationMathMl('<math><mtext mathcolor="red">C:\\Users</mtext></math>'));
assert.doesNotThrow(() => assertResolvedMathJaxSvg('<svg><g data-mml-node="mtext" fill="red"><text>important</text></g></svg>'));
assert.throws(() => assertResolvedMathJaxSvg('<svg><g data-mml-node="merror"><text>bad</text></g></svg>'));
checked += 3;
console.log(JSON.stringify({ suite: "omml_core_export_regression", checked, passed: true }));
