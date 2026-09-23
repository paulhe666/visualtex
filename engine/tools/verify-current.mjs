// Reproducible non-Git verification receipts. Writes only owned test outputs
// below engine/target/verification; never changes production, index or refs.
import {spawnSync} from 'node:child_process';
import {createHash} from 'node:crypto';
import {readFileSync,readdirSync,mkdirSync,writeFileSync} from 'node:fs';
import {resolve,dirname,relative} from 'node:path';
import {fileURLToPath} from 'node:url';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const stamp=new Date().toISOString().replaceAll(':','-').replaceAll('.','-');
const output=resolve(root,'target/verification',stamp);mkdirSync(output,{recursive:true});
const sha=bytes=>createHash('sha256').update(bytes).digest('hex');
function inputs(){
 const files=[];
 const walk=dir=>{for(const entry of readdirSync(dir,{withFileTypes:true}).sort((a,b)=>a.name.localeCompare(b.name))){
  const path=resolve(dir,entry.name);if(entry.isDirectory())walk(path);else if(entry.isFile())files.push({path:relative(root,path).replaceAll('\\','/'),sha256:sha(readFileSync(path))});
 }};
 for(const folder of ['crates','web/host','web/sandbox','specs','tools','fixtures'])walk(resolve(root,folder));
 for(const file of ['Cargo.toml','Cargo.lock','rust-toolchain.toml','package.json'])files.push({path:file,sha256:sha(readFileSync(resolve(root,file)))});
 return {files,sha256:sha(JSON.stringify(files))};
}
const before=inputs();const steps=[];
const commands=[
 ['build-and-native-wasm',['tools/build.mjs']],
 ['sandbox',['tools/sandbox-browser.mjs']],
 ['roots',['tools/radical-browser.mjs']],
 ['fences',['tools/fence-browser.mjs']],
 ['accents',['tools/accent-browser.mjs']],
 ['operators',['tools/operator-browser.mjs']],
 ['styles',['tools/style-browser.mjs']],
 ['spacing-phantom-boxed',['tools/spacing-browser.mjs']],
 ['colors',['tools/color-browser.mjs']],
 ['symbols',['tools/symbol-browser.mjs']],
 ['input-fault-injection',['tools/input-bridge-contract-review.mjs']],
 ['svg-geometry',['tools/svg-view-review.mjs']],
 ['protection',['../docs/engine-rebuild/tools/inspect-baseline.mjs','--check','docs/engine-rebuild/checkpoints/20260913-P00-provenance.json','--git-transition','docs/engine-rebuild/checkpoints/20260913-P00-git-transition.json']],
];
let failed=false;
for(const [name,args] of commands){
 const startedAt=new Date().toISOString();const start=performance.now();
 const run=spawnSync(process.execPath,args,{cwd:root,encoding:'utf8',timeout:100000,maxBuffer:32*1024*1024});
 const log=`$ node ${args.join(' ')}\nSTART ${startedAt}\n${run.stdout??''}\nSTDERR\n${run.stderr??''}\nERROR ${run.error?.message??''}\nEXIT ${run.status}\n`;
 writeFileSync(resolve(output,`${name}.log`),log);
 const step={name,command:['node',...args],startedAt,durationMs:Math.round(performance.now()-start),exitCode:run.status,
  error:run.error?.message??null,status:run.status===0&&!run.error?'PASS':'FAILED',log:`${name}.log`,logSha256:sha(log)};
 if(name==='build-and-native-wasm')step.rustPassed=[...(run.stdout??'').matchAll(/test result: ok\. (\d+) passed;/g)].reduce((sum,m)=>sum+Number(m[1]),0);
 steps.push(step);console.log(JSON.stringify(step));
 if(step.status!=='PASS'){failed=true;console.error(log.slice(-18000));break;}
}
const after=inputs();if(before.sha256!==after.sha256){failed=true;console.error('Inputs changed during verification; this run cannot qualify a source checkpoint.');}
const report={schema:'visualtex-engine-verification-v1',status:failed?'FAILED':'PASS',startedAt:stamp,sourceSha256:before.sha256,
 sourceUnchanged:before.sha256===after.sha256,inputs:before.files,steps,physicalWindowsIme:false,officeActions:0,gitMutations:0};
writeFileSync(resolve(output,'report.json'),JSON.stringify(report,null,2));
console.log(JSON.stringify({status:report.status,receipt:relative(root,resolve(output,'report.json')).replaceAll('\\','/'),sourceSha256:before.sha256,steps:steps.length}));
if(failed)process.exitCode=1;
