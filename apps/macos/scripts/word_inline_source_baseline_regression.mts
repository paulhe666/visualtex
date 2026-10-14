import assert from 'node:assert/strict';
import { readFileSync, writeFileSync } from 'node:fs';
import { DOMParser } from '@xmldom/xmldom';
import { renderOfficeFormulaArtifacts } from '../src/office/shared/formulaRenderArtifacts.ts';
import { wordImageReferenceGeometry } from '../src/office/shared/wordImageGeometry.ts';
import { findLatexFormulaSpans } from '../src/office/documentImport/documentImportParser.ts';
import { findWindowsWordLatexRedrawSpans } from '../src/office/redraw/wordLatexRedrawParser.ts';
import { formatLatexLines, latexCodeFormats } from '../src/clipboard/LatexCopyService.ts';
import { normalizeChineseLatex } from '../src/editor/normalizeChineseLatex.ts';
import { readWordFormulaFontSize, writeWordFormulaFontSize } from '../src/office/shared/wordFormulaPreferences.ts';
globalThis.DOMParser ??= DOMParser;
const doc = new DOMParser().parseFromString('<root/>', 'application/xml');
Object.getPrototypeOf(doc).querySelector ??= function(name) { return this.getElementsByTagName(name)?.item(0) ?? null; };
if (!('children' in Object.getPrototypeOf(doc.documentElement))) Object.defineProperty(Object.getPrototypeOf(doc.documentElement), 'children', { get() { return Array.from(this.childNodes ?? []).filter(n => n.nodeType === 1); } });
const cases = [String.raw`N=150`, String.raw`n_{\max}=12`, 'x', 'x_i', 'x^2', String.raw`\frac{a}{b}`,String.raw`\frac{1}{1+\frac{x}{y}}`,String.raw`\sqrt{x}`,String.raw`\sqrt[3]{x_i}`,String.raw`\int_0^1 f(x)\,\mathrm{d}x`,String.raw`\sum_{i=1}^n a_i`,String.raw`\prod_i p_i`,String.raw`\lim_{x\to0}\frac{\sin x}{x}`,String.raw`\left(\frac{a}{b}\right)`,String.raw`\vec{x}+\hat y`,String.raw`\overline{x}+\underline y`,String.raw`\begin{matrix}a&b\\c&d\end{matrix}`,String.raw`\begin{cases}x&x>0\\-x&x<0\end{cases}`,String.raw`\text{中文}x_{\text{下标}}`,String.raw`\binom{n}{k}`,String.raw`\operatorname{rank}(A)=2`,String.raw`x_\alpha+\frac12+\sqrt x`];
let count = 0;
for(const font of ['times','katex','cambria','stix','palatino','helvetica'] as const) for (const size of [10.5,11,11.5,12,14,18,24]) for(const latex of cases) {
 const input={lines:[{id:'x',latex}],codeFormat:'raw' as const,displayMode:'inline' as const,host:'word' as const,formulaLetterFont:font};
 const original=renderOfficeFormulaArtifacts(input).svg;
 const aligned=renderOfficeFormulaArtifacts({...input,fontSizePt:size}).svg;
 const g=wordImageReferenceGeometry(aligned.width,aligned.height,aligned.baseline,font);
 const actualHeight=Math.round(g.referenceHeightPt*size/14*20)/20; // Word shape quantization
 const raw=g.referenceBaselinePt*actualHeight/g.referenceHeightPt;
 const position=-Math.floor(-raw+.51);
 const baselineError=position+(aligned.height-aligned.baseline)/aligned.height*actualHeight;
 assert.ok(Math.abs(baselineError)<.06,`${font} ${size} ${latex}: ${baselineError}`);
 assert.equal(aligned.width,original.width);
 assert.equal(aligned.baseline,original.baseline);
 assert.ok(aligned.height>=original.height);
 const box=aligned.svg.match(/viewBox="([^"]+)"/)![1].split(' ').map(Number);
 assert.ok(Math.abs(box[3]/aligned.height-1000/(14*96/72))<1e-6);
 count++;
}
const legacy=String.raw`前 $t_{\text{start}}=\max\left(a,$\ $b\right)$ 后 $N=150$`;
const recovered=findLatexFormulaSpans(legacy);
assert.equal(recovered.length,2);
assert.equal(recovered[0].latex,String.raw`t_{\text{start}}=\max\left(a,\ b\right)`);
assert.deepEqual(findWindowsWordLatexRedrawSpans(legacy).map(s=>[s.start,s.end,s.latex]),recovered.map(s=>[s.start,s.end,s.latex]));
for(const span of recovered)assert.equal(legacy.slice(span.start,span.end),span.sourceText);
assert.equal(findLatexFormulaSpans(String.raw`$\left(a$ 正文 $b\right)$`).length,2);
assert.equal(findLatexFormulaSpans(String.raw`$a$ $b$`).length,2);
assert.equal(findLatexFormulaSpans(String.raw`$\left(a$ \ $b$`).length,2);
assert.equal(normalizeChineseLatex(String.raw`\text{前}$\left(a,$\ $b\right)$\text{后}`),String.raw`\text{前}\left(a,\ b\right)\text{后}`);
for(const {id} of latexCodeFormats) {
 if(id==='raw')continue;
 const source=formatLatexLines([String.raw`\text{前}$\left(a,$\ $b\right)$\text{后}`],id);
 for(const span of findLatexFormulaSpans(source))renderOfficeFormulaArtifacts({lines:[{id:'x',latex:span.latex}],codeFormat:'raw',displayMode:span.displayMode,host:'word'});
}
const values=new Map<string,string>();
globalThis.window={ localStorage:{getItem:(k)=>values.get(k)??null,setItem:(k,v)=>values.set(k,v)} } as any;
assert.equal(readWordFormulaFontSize('image'),null);
values.set('visualtex.office.word.create.font-size-pt','11.5');
assert.equal(readWordFormulaFontSize('image'),11.5);
writeWordFormulaFontSize('image',12);writeWordFormulaFontSize('omml',14);
assert.equal(readWordFormulaFontSize('image'),12);assert.equal(readWordFormulaFontSize('omml'),14);
console.log(`PASS: ${count} inline geometry cases; unchanged glyph size/baseline; legacy split scopes; editor normalization and font preferences`);
const out='test-results/source-baseline-20260912';
writeFileSync(`${out}/baseline-cases.json`,JSON.stringify(cases,null,2));
writeFileSync(`${out}/baseline-word-source.txt`,cases.map((latex,i)=>`${String(i+1).padStart(2,'0')} 正文Hxy $${latex}$ Hxy正文。`).join('\n\n'));
