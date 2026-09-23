// Reproducible P08 grid-environment acceptance bundle. This is not the P00
// protection runner and cannot override its red state. Writes only owned target receipts.
import {spawnSync} from 'node:child_process';
import {createHash} from 'node:crypto';
import {readFileSync,readdirSync,mkdirSync,writeFileSync} from 'node:fs';
import {resolve,dirname,relative} from 'node:path';
import {fileURLToPath} from 'node:url';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const output=resolve(root,'target/matrix-verification',new Date().toISOString().replaceAll(':','-').replaceAll('.','-'));
mkdirSync(output,{recursive:true});const sha=bytes=>createHash('sha256').update(bytes).digest('hex');
function inputs(){
 const files=[];const walk=dir=>{for(const e of readdirSync(dir,{withFileTypes:true}).sort((a,b)=>a.name.localeCompare(b.name))){
  const path=resolve(dir,e.name);if(e.isDirectory())walk(path);else if(e.isFile())files.push({path:relative(root,path).replaceAll('\\','/'),sha256:sha(readFileSync(path))});
 }};
 for(const folder of ['crates','web/host','web/sandbox','specs','tools','fixtures'])walk(resolve(root,folder));
 for(const path of ['Cargo.toml','Cargo.lock','rust-toolchain.toml','package.json'])files.push({path,sha256:sha(readFileSync(resolve(root,path)))});
 return {files,sha256:sha(JSON.stringify(files))};
}
const before=inputs(),steps=[];
const commands=[['build','tools/build.mjs'],['matrix-parity','tools/matrix-parity-review.mjs'],['matrix-browser','tools/matrix-browser-review.mjs'],
 ['rectangle-browser','tools/grid-rectangle-browser-review.mjs'],['virtual-viewport-browser','tools/virtual-viewport-browser-review.mjs'],
 ['grid-native-ime','tools/grid-native-ime-review.mjs'],['gathered-parity','tools/contained-multiline-parity-review.mjs'],['gathered-browser','tools/contained-multiline-browser-review.mjs'],['grid-boundary','tools/grid-view-contract-review.mjs']];
let failed=false;
for(const [name,tool] of commands){
 const startedAt=new Date().toISOString(),start=performance.now();
 const result=spawnSync(process.execPath,[tool],{cwd:root,encoding:'utf8',timeout:150000,maxBuffer:32*1024*1024});
 const log=`START ${startedAt}\nnode ${tool}\n${result.stdout??''}\nSTDERR\n${result.stderr??''}\nERROR ${result.error?.message??''}\nEXIT ${result.status}\n`;
 writeFileSync(resolve(output,name+'.log'),log);
 const step={name,status:result.status===0&&!result.error?'PASS':'FAILED',startedAt,durationMs:Math.round(performance.now()-start),exitCode:result.status,error:result.error?.message??null,log:name+'.log',logSha256:sha(log)};
 if(name==='build')step.rustPassed=[...(result.stdout??'').matchAll(/test result: ok\. (\d+) passed;/g)].reduce((n,m)=>n+Number(m[1]),0);
 const receipt=(result.stdout??'').match(/"receipt":\s*"([^"]+)"/);if(receipt)step.receipt=receipt[1];
 steps.push(step);console.log(JSON.stringify(step));
 if(step.status!=='PASS'){failed=true;console.error(log.slice(-14000));break;}
}
const after=inputs();if(before.sha256!==after.sha256)failed=true;
const physicalWindowsIme=steps.some(step=>step.name==='grid-native-ime'&&step.status==='PASS');
const report={schema:'visualtex-p08-grid-verification-v1',status:failed?'FAILED':'PASS',scope:'P08 isolated engine/sandbox candidate: Gathered, Matrix, identified cell rectangles, Cases, Aligned/align/align*, long-document viewport virtualization and physical Windows IME; P00 remains a separate gate',sourceSha256:before.sha256,
 sourceUnchanged:before.sha256===after.sha256,inputs:before.files,steps,physicalWindowsIme,officeActions:0,gitMutations:0};
writeFileSync(resolve(output,'report.json'),JSON.stringify(report,null,2));
console.log(JSON.stringify({status:report.status,receipt:relative(root,resolve(output,'report.json')).replaceAll('\\','/'),sourceSha256:before.sha256,steps:steps.length}));
if(failed)process.exitCode=1;
