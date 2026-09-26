import assert from "node:assert/strict";
import { latexToMathMl } from "../src/export/runtime.ts";

const cases = [
  ["aligned", String.raw`\begin{aligned}f(x)&=x^2+1\\f'(x)&=2x\end{aligned}`, true],
  ["operatorname", String.raw`\operatorname{rank}(A)+\operatorname{tr}(B)`, false],
  ["overset-underset", String.raw`\overset{!}{=}+\underset{n\to\infty}{\lim}`, true],
  ["physics-qty", String.raw`\qty(\frac{a}{b})`, false],
  ["physics-dv", String.raw`\dv{f}{x}+\dv[2]{g}{t}`, false],
  ["physics-pdv", String.raw`\pdv{u}{x}+\pdv[2]{u}{y}`, false],
  ["physics-vqty", String.raw`\vqty{\frac{p^2}{2m}+V(x)}`, false],
  ["physics-vev", String.raw`\vev{\hat H}`, false],
  ["physics-flatfrac", String.raw`\flatfrac{a+b}{c+d}`, false],
  ["physics-qq", String.raw`x\qq{if}y>0`, false],
  ["physics-ket-bra", String.raw`\bra{\phi}\hat H\ket{\psi}`, false],
  ["physics-braket-one", String.raw`\braket{\phi|\psi}`, false],
  ["physics-braket-two", String.raw`\braket{\phi}{\psi}`, false],
  ["physics-braket-multi", String.raw`\braket{\phi|A|\psi}`, false],
  ["physics-commutator", String.raw`\comm{A}{B}+\anticomm{C}{D}`, false],
  ["physics-mqty", String.raw`\mqty(a&b\\c&d)`, true],
  ["siunitx", String.raw`v=\SI{3.00}{\meter\per\second}`, false],
  ["mathbbm", String.raw`\mathbbm{1}_{A}(x)`, false],
  ["middle-delimiter", String.raw`\left\langle x\middle|A\middle|y\right\rangle`, false],
  ["long-nested", String.raw`\sum_{n=1}^{\infty}\frac{(-1)^{n+1}}{n^2}\left[\sqrt{1+\frac{x_n^2}{1+x_n^4}}+\int_0^1\frac{t^n}{1+t^2}\,\mathrm dt\right]`, true],
] as const;

for (const [name, latex, display] of cases) {
  const mathMl = await latexToMathMl(latex, display);
  assert.match(mathMl, /<math\b/u, `${name}: no MathML root`);
  assert.doesNotMatch(mathMl, /<merror\b/u, `${name}: MathJax emitted merror`);
  console.log(`[MathType source corpus PASS] ${name}: mathml=${mathMl.length}`);
}

console.log(`MathType source corpus passed: ${cases.length} cases.`);
