import {latexToMathMl} from '../src/export/runtime.ts';
import {writeFileSync,existsSync} from 'node:fs';
const sources=[String.raw`\sin x`,String.raw`\operatorname{sin}x`,String.raw`\operatorname{foo}^{2}⁡x`,String.raw`\underset{x\to0}{\operatorname{lim}}`];
const formulas=[];for(let i=0;i<sources.length;i++)formulas.push({index:i+1,id:`F${i+1}`,latex:sources[i],mathMl:await latexToMathMl(sources[i],i%2===1),displayMode:i%2===1?'block':'inline',parseStatus:'passed'});
const path='../../artifacts/conversion-symbol-audit-20261002/cp252-function-recovery/word4.json';if(existsSync(path))throw new Error('Refusing overwrite');writeFileSync(path,JSON.stringify({schemaVersion:1,acceptanceCorpus:true,scope:'CP252 four strict function-role/application controls; no native identity exemption',sourceExport:'production latexToMathMl',formulas},null,2));console.log('Prepared4');
