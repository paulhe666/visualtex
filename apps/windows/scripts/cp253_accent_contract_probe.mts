import {latexToMathMl} from '../src/export/runtime.ts';
import {writeFileSync,existsSync} from 'node:fs';
const sources=[String.raw`\hat{xyz}`,String.raw`\widehat{xyz}`,String.raw`\tilde{xyz}`,String.raw`\widetilde{xyz}`,String.raw`\vec{xyz}`,String.raw`\overrightarrow{xyz}`,String.raw`\mathbf{\hat{xyz}}`,String.raw`\hat{\mathbf{xyz}}`];
const rows=[];for(let i=0;i<sources.length;i++)rows.push({id:`A${i+1}`,latex:sources[i],mathMl:await latexToMathMl(sources[i],false)});
const p='../../artifacts/conversion-symbol-audit-20261002/cp253-accent-input.json';if(existsSync(p))throw new Error('Refusing overwrite');writeFileSync(p,JSON.stringify(rows,null,2));console.log('Prepared8');
