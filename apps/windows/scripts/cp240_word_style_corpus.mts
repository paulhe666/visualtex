import {latexToMathMl} from '../src/export/runtime.ts';
import {writeFileSync} from 'node:fs';
const sources=[String.raw`\mathrm{\hslash}`,String.raw`\mathbfit{\hslash}`,String.raw`\mathrm{\hbar}`,String.raw`\mathbfit{\hbar}`,String.raw`\mathrm{\alpha}`,String.raw`\mathbf{\alpha}`,String.raw`\mathit{\Gamma}`,String.raw`\mathbfit{\Gamma}`,String.raw`\mathrm{\ell}`,String.raw`\mathbf{\ell}`,String.raw`\mathrm{\angle}`,String.raw`\mathit{\angle}`,String.raw`\mathbfit{\hbar}^{2}`,String.raw`\mathrm{\hslash}_{i}`,String.raw`\mathbf{\alpha+\mathrm{\alpha}}`,String.raw`\mathrm{\Bbbk}`,String.raw`\mathit{\operatorname{sin}x}`,String.raw`\mathrm{\mathbb{R}+x}`,String.raw`\mathit{\mathcal{A}+x}`,String.raw`x_i^2+\frac{a}{b}`];
const formulas=[];for(let i=0;i<sources.length;i++)formulas.push({index:i+1,id:`S${String(i+1).padStart(2,'0')}`,latex:sources[i],mathMl:await latexToMathMl(sources[i],i%2===1),displayMode:i%2===1?'block':'inline',parseStatus:'passed'});
const base={schemaVersion:1,acceptanceCorpus:true,scope:'CP240 current source explicit-identifier-style Word integration; independent focused corpus, original F02/F09 failures remain recorded',sourceExport:'production latexToMathMl',formulas};
writeFileSync('../../artifacts/conversion-symbol-audit-20261002/cp240-font-open/style20.json',JSON.stringify(base,null,2));
writeFileSync('../../artifacts/conversion-symbol-audit-20261002/cp240-font-open/style4.json',JSON.stringify({...base,formulas:formulas.slice(0,4)},null,2));
console.log('Generated 20 current-source formulas and four-case focused subset');
