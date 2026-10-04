import {convertLatexToMathMl,convertLatexToMarkup} from 'mathlive';
import {readFileSync,writeFileSync} from 'node:fs';
const r=JSON.parse(readFileSync('../../artifacts/conversion-symbol-audit-20261002/cp239-identifier-scope-candidate.json','utf8'));
const rows=r.rows.map((x:any)=>{try{const html=convertLatexToMarkup(x.latex);return{latex:x.latex,mathMl:convertLatexToMathMl(x.latex),classes:[...new Set([...html.matchAll(/class="([^"]+)"/g)].map(m=>m[1]))]};}catch(e){return{latex:x.latex,error:String(e)}}});
writeFileSync('../../artifacts/conversion-symbol-audit-20261002/cp239-mathlive-scope.json',JSON.stringify({rows},null,2));console.log(JSON.stringify({rows}));
