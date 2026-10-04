import {latexToMathMl} from '../src/export/runtime.ts';
import {writeFileSync,existsSync} from 'node:fs';
const root='../../artifacts/conversion-symbol-audit-20261002/cp251-preserve-structure/';
const sources=[String.raw`(x+y)^2`,String.raw`\left[\frac{x}{y}+(x+y)\right]`,String.raw`\mathbf{(x)}`,String.raw`x^2`,String.raw`x_i^j`,String.raw`x^{a_b}`,String.raw`x_{a^b}`,String.raw`\frac{a}{b}`,String.raw`\sqrt{x_i^2+y_i^2}`,String.raw`\sqrt[3]{x^{a_b}}`,String.raw`\sum_{i=1}^{n}x_i^2`,String.raw`\int_0^1x^2\,dx`,String.raw`\hat{x}^2+\vec{v}_i`,String.raw`\begin{matrix}a&b\\c&d\end{matrix}`,String.raw`\left(\frac{a}{b}\right)^2`,String.raw`\left.\frac{df}{dx}\right|_{x=0}`,String.raw`\mathbb{R}+\mathfrak{g}`,String.raw`\mathrm{\alpha}+\mathit{\Gamma}`,String.raw`\symtt{x09}`,String.raw`a\quad b`];
const formulas=[];for(let i=0;i<sources.length;i++)formulas.push({index:i+1,id:`P${String(i+1).padStart(2,'0')}`,latex:sources[i],mathMl:await latexToMathMl(sources[i],i%2===1),displayMode:i%2===1?'block':'inline',parseStatus:'passed'});
const base={schemaVersion:1,acceptanceCorpus:true,scope:'CP251 structured import controls; known original F02/F09/hbar failures retained separately',sourceExport:'production latexToMathMl',formulas};
for(const count of [3,20]){const path=root+'import'+count+'.json';if(existsSync(path))throw new Error('Refusing overwrite');writeFileSync(path,JSON.stringify({...base,formulas:formulas.slice(0,count)},null,2));}
console.log('Generated3focused and20daily structured import cases');
