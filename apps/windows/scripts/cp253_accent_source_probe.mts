import {latexToMathMl} from '../src/export/runtime.ts';
import {writeFileSync,existsSync} from 'node:fs';
const sources=[String.raw`\overset{\mathbf{\hat{}}}{x}`,String.raw`\overset{\mathbf{ˆ}}{x}`,String.raw`\accentset{\mathbf{ˆ}}{x}`,String.raw`\mathbf{\hat{\mathnormal{x}}}`,String.raw`\mathbf{\hat{\mathrm{x}}}`,String.raw`\mathit{\hat{\mathnormal{x}}}`];
const rows=[];for(const latex of sources){try{rows.push({latex,mathMl:await latexToMathMl(latex,false)})}catch(e){rows.push({latex,error:String(e)})}}
const p='../../artifacts/conversion-symbol-audit-20261002/cp253-accent-source-probe.json';if(existsSync(p))throw new Error('Refusing overwrite');writeFileSync(p,JSON.stringify(rows,null,2));console.log(JSON.stringify(rows));
