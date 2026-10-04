import {latexToMathMl} from '../src/export/runtime.ts';
import {readFileSync,writeFileSync,existsSync} from 'node:fs';
const root='../../artifacts/conversion-symbol-audit-20261002/';
const input=JSON.parse(readFileSync(root+'cp257-editor-color.json','utf8'));
const rows=[];for(const [index,row] of input.rows.entries()) rows.push({id:'color-'+index,latex:row.initial,mathMl:await latexToMathMl(row.initial,false),before:row.before,source:row.source,range:row.range});
const p=root+'cp257-color-payload.json';if(existsSync(p))throw new Error('Refusing overwrite');writeFileSync(p,JSON.stringify(rows,null,2));console.log(JSON.stringify(rows));