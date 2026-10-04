import assert from 'node:assert/strict';
import {readFileSync,writeFileSync} from 'node:fs';
import {spawnSync} from 'node:child_process';
import {latexToMathMl} from '../src/export/runtime.ts';
const path='../../artifacts/conversion-symbol-audit-20261002/';
const expected=JSON.parse(readFileSync(path+'cp239-identifier-scope-candidate.json','utf8'));
const rows=[];for(const row of expected.rows){const actual=await latexToMathMl(row.latex,false);assert.equal(actual,row.candidate,row.latex);rows.push({latex:row.latex,actual,passed:true});}
const run=spawnSync(process.execPath,['--import','tsx','scripts/mathml_hslash_hbar_source_regression.mts','candidate'],{encoding:'utf8',maxBuffer:20*1024*1024});
writeFileSync(path+'cp239-hslash-hbar-production.json',run.stdout);assert.equal(run.status,0,run.stderr);const source=JSON.parse(run.stdout);assert.equal(source.summary.passed,20);
const report={passed:rows.length,failed:0,scope:'Explicit identifier style contract with function/intrinsic/nested/text-italic/combining controls',rows};writeFileSync(path+'cp239-identifier-production.json',JSON.stringify(report,null,2));console.log(JSON.stringify({identifierControls:rows.length,hslashHbarSource:source.summary}));
