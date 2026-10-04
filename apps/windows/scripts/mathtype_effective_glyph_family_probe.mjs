import assert from 'node:assert/strict';
import {mkdir, writeFile} from 'node:fs/promises';
import {resolve, join} from 'node:path';
import {latexToMathMl, latexToSvg} from '../src/export/runtime.ts';

const root = resolve(process.argv[2] ?? (() => {throw new Error('An owned evidence directory is required.');})());
await mkdir(root, {recursive: true});
const cases = [
  ['plain', '-'], ['cal', String.raw`\mathcal{-}`], ['frak', String.raw`\mathfrak{-}`],
  ['sf', String.raw`\mathsf{-}`], ['tt', String.raw`\mathtt{-}`], ['bb', String.raw`\mathbb{-}`],
  ['bold', String.raw`\boldsymbol{-}`], ['italic', String.raw`\mathit{-}`],
  ['cal-letter-control', String.raw`\mathcal{A}`], ['plain-letter-control', 'A'],
  ['bold-ddot-contour', String.raw`\boldsymbol{\ddot{\oint}}`],
];
const observations = [];
for (const [id, latex] of cases) {
  const mathMl = await latexToMathMl(latex, true);
  const result = await latexToSvg(latex, {displayMode: true, fontSizePt:12, paddingPx:0, background:'transparent'});
  assert.doesNotMatch(mathMl, /<merror/);
  const paths = [...result.svg.matchAll(/<path\b[^>]*\bd="([^"]+)"[^>]*>/g)].map(m => m[1]);
  const uses = [...result.svg.matchAll(/<use\b[^>]*>/g)].map(m => m[0]);
  await writeFile(join(root, id+'.svg'), result.svg, 'utf8');
  observations.push({id, latex, mathMl, width:result.width, height:result.height, paths, uses});
}
await writeFile(join(root,'observations.json'),JSON.stringify(observations,null,2),'utf8');
const plain = observations[0];
for (const item of observations) {
  console.log(JSON.stringify({id:item.id, samePaths:JSON.stringify(item.paths)===JSON.stringify(plain.paths), width:item.width, height:item.height, paths:item.paths.length, uses:item.uses}));
}
