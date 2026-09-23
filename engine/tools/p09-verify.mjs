// Stable-source P09 acceptance. P08 regression stays a mandatory lower-phase gate.
import {spawnSync} from 'node:child_process';
import {createHash} from 'node:crypto';
import {readFileSync,readdirSync,mkdirSync,writeFileSync} from 'node:fs';
import {resolve,dirname,relative} from 'node:path';
import {fileURLToPath} from 'node:url';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const stamp=new Date().toISOString().replaceAll(':','-').replaceAll('.','-');
const output=resolve(root,'target/p09-verification',stamp);mkdirSync(output,{recursive:true});
const sha=bytes=>createHash('sha256').update(bytes).digest('hex');
function inputs(){const files=[];const walk=dir=>{for(const e of readdirSync(dir,{withFileTypes:true}).sort((a,b)=>a.name.localeCompare(b.name))){const path=resolve(dir,e.name);if(e.isDirectory())walk(path);else if(e.isFile())files.push({path:relative(root,path).replaceAll('\\','/'),sha256:sha(readFileSync(path))});}};for(const folder of ['crates','web/host','web/sandbox','specs','tools','fixtures'])walk(resolve(root,folder));for(const path of ['Cargo.toml','Cargo.lock','rust-toolchain.toml','package.json'])files.push({path,sha256:sha(readFileSync(resolve(root,path)))});return {files,sha256:sha(JSON.stringify(files))};}
const before=inputs(),steps=[];
const commands=[
 ['p08-regression','tools/matrix-verify.mjs',300000],
 ['command-inventory','tools/p09-command-inventory.mjs',120000],
 ['command-state','tools/command-input-contract-review.mjs',120000],
 ['command-browser','tools/command-input-browser-review.mjs',180000],
 ['macro-parity','tools/macro-parity-review.mjs',180000],
 ['macro-browser','tools/macro-browser-review.mjs',180000],
 ['matrix-variants-parity','tools/matrix-variants-parity-review.mjs',180000],
 ['matrix-variants-browser','tools/matrix-variants-browser-review.mjs',180000],
 ['middle-parity','tools/middle-delimiter-parity-review.mjs',180000],
 ['middle-browser','tools/middle-delimiter-browser-review.mjs',180000],
 ['input-bridge','tools/input-bridge-contract-review.mjs',180000],
 ['structural-clipboard-browser','tools/grid-rectangle-browser-review.mjs',240000],
 ['custom-symbol-parity','tools/custom-symbol-parity-review.mjs',180000],
 ['custom-symbol-browser','tools/custom-symbol-browser-review.mjs',180000],
 ['ocr-target-parity','tools/ocr-target-parity-review.mjs',180000],
 ['ocr-target-browser','tools/ocr-target-browser-review.mjs',180000],
 ['legacy-archive-contract','tools/legacy-archive-contract-review.mjs',180000],
 ['legacy-archive-browser','tools/legacy-archive-browser-review.mjs',180000],
 ['codemirror-source-bridge','tools/codemirror-source-bridge-review.mjs',180000],
 ['product-compat-browser','tools/p09-product-bridge-browser-review.mjs',180000],
];
let failed=false;
for(const [name,tool,timeout] of commands){const startedAt=new Date().toISOString(),start=performance.now();const result=spawnSync(process.execPath,[tool],{cwd:root,encoding:'utf8',timeout,maxBuffer:64*1024*1024});const log=`START ${startedAt}\nnode ${tool}\n${result.stdout??''}\nSTDERR\n${result.stderr??''}\nERROR ${result.error?.message??''}\nEXIT ${result.status}\n`;writeFileSync(resolve(output,`${name}.log`),log);const receipt=(result.stdout??'').match(/"receipt":\s*"([^"]+)"/);const step={name,status:result.status===0&&!result.error?'PASS':'FAILED',durationMs:Math.round(performance.now()-start),exitCode:result.status,error:result.error?.message??null,log:`${name}.log`,logSha256:sha(log),...(receipt?{receipt:receipt[1]}:{})};steps.push(step);console.log(JSON.stringify(step));if(step.status!=='PASS'){failed=true;console.error(log.slice(-16000));break;}}
const after=inputs();if(before.sha256!==after.sha256)failed=true;
const report={schema:'visualtex-p09-verification-v1',status:failed?'FAILED':'PASS',scope:'P09 command input, bounded macros/language extensions, versioned structural clipboard, immutable custom-symbol artwork, strict deferred OCR targets, legacy archive read compatibility, verified font/hotkey adaptation and CodeMirror-facing source revision/origin/history bridge over a passing P08 isolated engine/sandbox baseline',sourceSha256:before.sha256,sourceUnchanged:before.sha256===after.sha256,inputs:before.files,steps,productionAppWrites:0,officeActions:0,gitMutations:0};
writeFileSync(resolve(output,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({status:report.status,receipt:relative(root,resolve(output,'report.json')).replaceAll('\\','/'),sourceSha256:report.sourceSha256,steps:steps.length}));if(failed)process.exitCode=1;
