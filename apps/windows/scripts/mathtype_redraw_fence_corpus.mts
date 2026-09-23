import fs from 'node:fs';
import path from 'node:path';
import { latexToMathMl } from '../src/export/runtime';

const [sourcePath, outputPath] = process.argv.slice(2);
if (!sourcePath || !outputPath) throw new Error('Usage: tsx scripts/mathtype_redraw_fence_corpus.mts <captured-source.json> <output.json>');
if (fs.existsSync(outputPath)) throw new Error(`Refusing to overwrite evidence: ${outputPath}`);
const source = JSON.parse(fs.readFileSync(sourcePath, 'utf8').replace(/^\uFEFF/, ''));
const formulas: {index: number; latex: string; displayMode: string; source?: string}[] = source.formulas.map((item: any) => ({...item, source: 'active-document'}));
const additional = [
  String.raw`\lvert x-\mu\rvert<\varepsilon`,
  String.raw`a+\lvert x\rvert+b`,
  String.raw`\lvert x\rvert+\lvert y\rvert`,
  String.raw`\lVert x\rVert<1`,
  String.raw`a+\lVert x\rVert+b`,
  String.raw`\langle x,y\rangle<1`,
  String.raw`\lceil x\rceil+\lfloor y\rfloor`,
  String.raw`\left|x-\mu\right|<\varepsilon`,
  String.raw`\left\lVert\frac{x}{y}\right\rVert+1`,
  String.raw`\lvert x\rvert^{2}+1`,
  String.raw`\lVert x\rVert_{2}+1`,
  String.raw`\lvert x\rvert_{i}^{2}+1`,
  String.raw`\left|x\right|^{2}+1`,
  String.raw`\lvert\lvert x\rvert\rvert+1`,
  String.raw`\left|\left|x\right|\right|+1`,
  String.raw`P(A\mid B)=\frac{P(A\cap B)}{P(B)}`,
  String.raw`a\mid b+c\mid d`,
  String.raw`\{x\mid x>0\}`,
  String.raw`\left\{x\mid x>0\right\}`,
  String.raw`\bra{x}A\ket{y}`,
  String.raw`\left.\frac{df}{dx}\right|_{x=0}`,
  String.raw`\left\{\begin{matrix}x\\y\end{matrix}\right.`,
  String.raw`(x+y)^{2}+1`,
  String.raw`\left(x+y\right)^{2}+1`,
  String.raw`f(x)+g(y)`,
  String.raw`\begin{pmatrix}x&y\\z&w\end{pmatrix}`,
];
for (const latex of additional) for (const displayMode of ['inline','block'])
  formulas.push({index: formulas.length+1, latex, displayMode, source: 'delimiter-regression'});
const results = [];
for (const formula of formulas) {
  try {
    const mathMl = await latexToMathMl(formula.latex, formula.displayMode === 'block');
    results.push({...formula, mathMl});
  } catch (error) {
    results.push({...formula, mathMl: '', exportError: String(error)});
    console.error(JSON.stringify({index:formula.index, latex:formula.latex, error:String(error)}));
  }
}
fs.mkdirSync(path.dirname(path.resolve(outputPath)), {recursive:true});
fs.writeFileSync(outputPath, JSON.stringify({document:source.name, textHash:source.textHash, formulas:results}, null, 2));
console.log(JSON.stringify({document:source.name, activeDocumentFormulas:source.formulas.length, additionalFormulas:additional.length*2, total:results.length, output:path.resolve(outputPath)}));
console.log(results.find(item => item.latex.includes('x_{8}-'))?.mathMl);
