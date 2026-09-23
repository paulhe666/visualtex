// Synthetic boundary checks for the current TypeScript grid geometry validator.
// Native/WASM scenes are genuine; corruptions below are explicit fault injection,
// not keyboard, OS IME or user-document acceptance.
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import { readFileSync, mkdirSync, writeFileSync } from 'node:fs';
import { resolve, relative } from 'node:path';
import { createSandboxAssets } from './sandbox-assets.mjs';
const {assets,manifest,root}=createSandboxAssets();
const module=assets.get('/host/grid-geometry.js');
assert.ok(module,'Current grid-geometry module is missing');
const {validateGridGeometry}=await import('data:text/javascript;base64,'+module.bytes.toString('base64'));
const {EngineRuntime}=createRequire(import.meta.url)(resolve(root,'web/pkg-node/visualtex_engine_wasm.js'));
const runtime=new EngineRuntime();
const fontPath=resolve(root,'../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf');
const registered=JSON.parse(runtime.registerFont(readFileSync(fontPath),0));assert.equal(registered.status,'ok');
const font=registered.result.font;
function send(command){const r=JSON.parse(runtime.dispatch(JSON.stringify({protocol:1,command})));assert.equal(r.status,'ok',JSON.stringify(r));return r.result;}
const output=resolve(root,'target/grid-view-contract-review',new Date().toISOString().replaceAll(':','-').replaceAll('.','-'));
mkdirSync(output,{recursive:true});const cases=[];
function check(name,test){try{test();cases.push({name,status:'PASS'});}catch(error){cases.push({name,status:'FAILED',error:String(error)});}}
function scene(source,id){const h=send({op:'create',documentId:id,source}).handle;
 const s=send({op:'layout',handle:h,expectedRevision:0,options:{mathFont:font.id,textFont:null,style:'display',fontSizePx:16,paddingEm:.125},includeSvg:false}).snapshot;
 send({op:'destroy',handle:h,expectedRevision:0});return s;}
try{
 const s=scene(String.raw`\begin{gathered}x\\y\end{gathered}`,'view-grid');
 check('actual current core scene is accepted',()=>validateGridGeometry(s));
 const nested=scene(String.raw`\begin{gathered}\begin{gathered}x\\y\end{gathered}\\z\end{gathered}`,'view-nested');
 check('nested owners retain distinct valid column anchors',()=>validateGridGeometry(nested));
 for(const command of ['phantom','hphantom','vphantom']){
  const hidden=scene(`\\${command}{\\begin{gathered}a+b+c+d\\\\e\\end{gathered}}x`,'view-'+command);
  check('hidden '+command+' grid can exceed export bounds but not owner bounds',()=>validateGridGeometry(hidden));
 }
 const faults=[
  ['missing required grid metadata',c=>{delete c.gridColumns;}],
  ['empty metadata on a real grid',c=>{c.gridColumns=[];}],
  ['NaN x',c=>{c.gridColumns[0].x=NaN;}],
  ['infinite width',c=>{c.gridColumns[0].width=Infinity;}],
  ['negative width',c=>{c.gridColumns[0].width=-1;}],
  ['unknown owner',c=>{c.gridColumns[0].gridId=0xfffffffe;}],
  ['root row used as grid owner',c=>{c.gridColumns[0].gridId=c.regions[0].nodeId;}],
  ['zero column identity',c=>{c.gridColumns[0].columnId=0;}],
  ['column identity aliases node',c=>{c.gridColumns[0].columnId=c.regions[0].nodeId;}],
  ['duplicate anchor',c=>{c.gridColumns.push(structuredClone(c.gridColumns[0]));}],
  ['noncontiguous index',c=>{c.gridColumns[0].index=1;}],
  ['overbudget index',c=>{c.gridColumns[0].index=64;}],
  ['column to right of owner',c=>{c.gridColumns[0].x+=100;}],
  ['oversized column',c=>{c.gridColumns[0].width+=100;}],
  ['string instead of numeric coordinate',c=>{c.gridColumns[0].x='0';}],
  ['duplicate grid region',c=>{c.regions.push(structuredClone(c.regions.find(r=>r.role==='grid')));}],
 ];
 for(const [name,mutate] of faults)check('reject '+name,()=>{const bad=structuredClone(s);mutate(bad);assert.throws(()=>validateGridGeometry(bad));});
 const noGrid=scene('x+1','view-plain');check('ordinary scene has no forced fake alignment data',()=>validateGridGeometry(noGrid));
 assert.equal(runtime.activeSessions,0);
}finally{runtime.free();}
const report={status:cases.every(c=>c.status==='PASS')?'PASS':'FAILED',cases,passed:cases.filter(c=>c.status==='PASS').length,
 failed:cases.filter(c=>c.status==='FAILED').length,hostBuild:manifest.buildId,
 moduleSha256:createHash('sha256').update(module.bytes).digest('hex'),syntheticFaultInjection:true,
 physicalWindowsIme:false,officeActions:0,productionWrites:0};
writeFileSync(resolve(output,'report.json'),JSON.stringify(report,null,2));
console.log(JSON.stringify({receipt:relative(root,resolve(output,'report.json')).replaceAll('\\','/'),...report},null,2));
if(report.status!=='PASS')process.exitCode=1;
