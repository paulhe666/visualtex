// A big operator's operand in the MathML the Word writers get (export/operandMathMl.ts, the rule of
// tex-layout/src/operand-scope.ts): the writers make the row after the operator the body of their native n-ary object, and MathJax
// writes the items loose, so the operand is wrapped in an mrow: an integral's up to a sign after its differential, a sum's up to
// a top-level sign, both up to a relation. Without it the body was the single element after the operator (the f of ∫ f(x) dx).
import assert from "node:assert/strict";
import { latexToMathMl } from "../src/export/runtime.ts";

function compact(latex: string): string {
  return latexToMathMl(latex, true)
    .replace(/\sdata-[a-z-]+="[^"]*"/g, "")
    .replace(/\s*\n\s*/g, "")
    .replace(/^<math[^>]*>/, "")
    .replace(/<\/math>$/, "");
}

const SPACE = '<mstyle scriptlevel="0"><mspace width="0.167em"></mspace></mstyle>';
const cases: [string, string][] = [
  // the integrand through its differential is the body; the plus and c stand outside; the limits are untouched
  [
    "\\int_a^b f(x)\\,dx + c",
    `<msubsup><mo>&#x222B;</mo><mi>a</mi><mi>b</mi></msubsup><mrow><mi>f</mi><mo stretchy="false">(</mo><mi>x</mi><mo stretchy="false">)</mo>${SPACE}<mi>d</mi><mi>x</mi></mrow><mo>+</mo><mi>c</mi>`,
  ],
  // a plus before the differential belongs to the integrand
  ["\\int p+q\\,dr+s", `<mo>&#x222B;</mo><mrow><mi>p</mi><mo>+</mo><mi>q</mi>${SPACE}<mi>d</mi><mi>r</mi></mrow><mo>+</mo><mi>s</mi>`],
  // a sum ends at the top-level plus
  [
    "\\sum_{i=1}^n a_i + b",
    "<munderover><mo>&#x2211;</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></munderover><mrow><msub><mi>a</mi><mi>i</mi></msub></mrow><mo>+</mo><mi>b</mi>",
  ],
  // a relation ends any operand
  ["\\prod_i x_i = c", "<munder><mo>&#x220F;</mo><mi>i</mi></munder><mrow><msub><mi>x</mi><mi>i</mi></msub></mrow><mo>=</mo><mi>c</mi>"],
  // an operator inside an operand: the inner operand is inside the outer one
  [
    "\\sum_i\\sum_j a_{ij}+b",
    "<munder><mo>&#x2211;</mo><mi>i</mi></munder><mrow><munder><mo>&#x2211;</mo><mi>j</mi></munder><mrow><msub><mi>a</mi><mrow><mi>i</mi><mi>j</mi></mrow></msub></mrow></mrow><mo>+</mo><mi>b</mi>",
  ],
  // a bracketed sign does not end it
  [
    "\\sum_i (a_i+b_i)",
    '<munder><mo>&#x2211;</mo><mi>i</mi></munder><mrow><mo stretchy="false">(</mo><msub><mi>a</mi><mi>i</mi></msub><mo>+</mo><msub><mi>b</mi><mi>i</mi></msub><mo stretchy="false">)</mo></mrow>',
  ],
];

for (const [latex, expected] of cases) assert.equal(compact(latex), expected, latex);

// nothing to group where there is no big operator
assert.equal(compact("a+bc"), "<mi>a</mi><mo>+</mo><mi>b</mi><mi>c</mi>");
// an operator with nothing after it keeps its limits and gets no empty row
assert.ok(!/<mrow><\/mrow>/.test(compact("\\sum_{i=1}^n")), "an operator at the row's end has no operand to wrap");

console.log(`operand in the MathML for the Word writers: ${cases.length + 2} checks OK`);
