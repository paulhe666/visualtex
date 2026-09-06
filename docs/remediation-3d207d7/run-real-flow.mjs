// Orchestrates the existing visible Word/Ribbon UI driver only. These actions
// collect evidence; they do not declare acceptance or create a service object.
import {spawnSync} from 'node:child_process';
import {readFileSync} from 'node:fs';
import {dirname,join,resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
const directory=dirname(fileURLToPath(import.meta.url));
const steps=JSON.parse(readFileSync(resolve(directory,process.argv[2]),'utf8').replace(/^\uFEFF/,''));
for(const step of steps){
 if(!Array.isArray(step)||!step.every(value=>typeof value==='string'))throw new Error('Expected string argument arrays.');
 console.log('REAL_UI_STEP',JSON.stringify(step));
 const result=spawnSync(process.execPath,[join(directory,'ui-flow.mjs'),...step],{stdio:'inherit',cwd:resolve(directory,'../..'),timeout:240000});
 if(result.error)throw result.error;
 if(result.status!==0)throw new Error('Real UI step stopped. Inspect its actual document/dialog before continuing.');
}
console.log('UI_STEPS_FINISHED_REVIEW_EVIDENCE_REQUIRED');
