import {latexToMathMl} from '../src/export/runtime.ts';
import {writeFileSync,existsSync} from 'node:fs';
const sources=[String.raw`\sin x`,String.raw`\operatorname{sin}⁡x`,String.raw`\operatorname{sin}x`,String.raw`\operatorname{foo}x`,String.raw`\mathrm{foo}x`,String.raw`\operatorname{foo}^{2}⁡x`,String.raw`\operatorname{sin}⁡+x`,String.raw`\operatorname{sin}⁡\,x`,String.raw`\mathit{\operatorname{sin}x}`,String.raw`\mathop{\mathrm{foo}}x`,String.raw`\mathbf{\operatorname{foo}}x`,String.raw`\mathrm{foo1}x`];
const rows=[];for(const latex of sources)rows.push({latex,mathMl:await latexToMathMl(latex,false)});
const path='../../artifacts/conversion-symbol-audit-20261002/cp252-function-spellings.json';if(existsSync(path))throw new Error('Refusing overwrite');writeFileSync(path,JSON.stringify(rows,null,2));console.log(JSON.stringify(rows));
