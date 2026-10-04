import {latexToMathMl} from '../src/export/runtime.ts';
import {readFileSync,writeFileSync,existsSync} from 'node:fs';
const root='../../artifacts/conversion-symbol-audit-20261002/';
const payload=JSON.parse(readFileSync(root+'cp257-color-payload.json','utf8'));
const cases=[];for(const row of payload)cases.push({name:row.id,sourceMathMl:await latexToMathMl(row.before,false),mathMl:row.mathMl,expectedText:row.id==='color-3'?'xy':row.id==='color-4'?'x':'x+y',expectedColors:row.id==='color-1'?['DC2626',null,null]:Array(row.id==='color-3'?2:row.id==='color-4'?1:3).fill('2563EB'),expectedControlColors:row.id==='color-3'||row.id==='color-4'?['2563EB']:[]});
const p=root+'cp257-color/word-cases.json';if(existsSync(p))throw new Error('Refusing overwrite');writeFileSync(p,JSON.stringify({cases},null,2));console.log('Prepared '+cases.length+' Word color cases');