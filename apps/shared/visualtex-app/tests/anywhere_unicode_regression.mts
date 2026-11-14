// VisualTeX Anywhere's Math Characters output (apps/shared/visualtex-app/src/anywhere/unicodeMath.ts):
// formulas as Unicode text, from the MathML of the export path.
//   tsx ../shared/visualtex-app/tests/anywhere_unicode_regression.mts
import assert from "node:assert/strict";
import { DOMParser } from "@xmldom/xmldom";
import { latexToMathMl } from "../src/export/runtime.ts";
import { mathMlToUnicode } from "../src/anywhere/unicodeMath.ts";

const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml") as never;
const unicode = (latex: string) => mathMlToUnicode(latexToMathMl(latex, false), parse);

const cases: [string, string][] = [
  ["x^2", "x²"],
  ["\\frac{a+b}{c+d}", "(a + b)/(c + d)"],
  ["x^{a+b}", "x^(a + b)"],
  ["\\int_a^b f(x)\\,dx", "∫ₐᵇ f(x) dx"],
  ["\\sqrt{x+1}", "√(x + 1)"],
  ["\\mathbb{R}", "ℝ"],
  ["\\sin x + \\cos^2 y", "sin x + cos² y"],
  ["\\lim_{x\\to 0} \\frac{\\sin x}{x} = 1", "lim_(x → 0) (sin x)/x = 1"],
  ["\\sum_{i=1}^{n} i^2", "∑ᵢ₌₁ⁿ i²"],
  ["a_{ij}", "aᵢⱼ"],
  ["x_{n+1}", "xₙ₊₁"],
  ["e^{-x^2}", "e^(−x²)"],
  ["\\alpha + \\beta = \\gamma", "α + β = γ"],
  ["\\hat{x}", "x\u0302"],
  ["\\vec{v}", "v\u20d7"],
  ["\\overline{AB}", "A\u0305B\u0305"],
  ["\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}", "(a, b; c, d)"],
  ["\\binom{n}{k}", "C(n, k)"],
  ["\\sqrt[3]{x}", "∛x"],
  ["\\sqrt[n]{x}", "ⁿ√x"],
  ["f'(x)", "f′(x)"],
  ["\\mathbf{v} \\cdot \\mathbf{w}", "𝐯 ⋅ 𝐰"],
  ["\\text{面积} = \\pi r^2", "面积 = πr²"],
  ["-x + y", "−x + y"],
  ["a \\le b", "a ≤ b"],
  ["\\left(\\frac{1}{2}\\right)^2", "(1/2)²"],
  ["\\frac{\\partial f}{\\partial x}", "∂f/∂x"],
  ["\\frac{1}{2x}", "1/(2x)"],
  ["\\sin\\left(x+1\\right)", "sin(x + 1)"],
  ["\\log_2 n", "log₂ n"],
  ["x^{-1}", "x⁻¹"],
  ["2^{10}", "2¹⁰"],
  ["\\nabla \\cdot \\mathbf{E} = \\frac{\\rho}{\\varepsilon_0}", "∇ ⋅ 𝐄 = ρ/ε₀"],
];

let failed = 0;
for (const [latex, expected] of cases) {
  const actual = unicode(latex).text;
  if (actual !== expected) {
    failed += 1;
    console.error(`FAIL ${latex}\n  expected ${expected}\n  actual   ${actual}`);
  }
}
assert.deepEqual(unicode("\\frac{a}{b}").notes, ["Fractions are written with a slash."]);
assert.deepEqual(unicode("x^2").notes, []);
assert.equal(failed, 0, `${failed} of ${cases.length} Unicode conversions differ`);
console.log(`anywhere unicode: ${cases.length} conversions OK`);
