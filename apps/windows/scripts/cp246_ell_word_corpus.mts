import {latexToMathMl} from '../src/export/runtime.ts';
import {writeFileSync,existsSync} from 'node:fs';
const path='../../artifacts/conversion-symbol-audit-20261002/cp246-native-style/ell-word4.json';
if(existsSync(path))throw new Error('Refusing overwrite');
const commands=['mathrm','mathit','mathbf','mathbfit'];
const formulas=[];
for(let i=0;i<commands.length;i++){const latex='\\'+commands[i]+'{\\ell}';formulas.push({index:i+1,id:'E0'+(i+1),latex,mathMl:await latexToMathMl(latex,i%2===1),displayMode:i%2===1?'block':'inline',parseStatus:'passed'});}
writeFileSync(path,JSON.stringify({schemaVersion:1,acceptanceCorpus:true,scope:'CP246 independent ell four basic styles; no script-l substitution',sourceExport:'production latexToMathMl',formulas},null,2));
console.log('Generated four ell styles');
