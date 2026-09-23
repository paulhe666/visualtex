import assert from 'node:assert/strict';
import {mkdirSync,writeFileSync} from 'node:fs';
import {performance} from 'node:perf_hooks';
import {relative,resolve} from 'node:path';
import {createSandboxAssets} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';

const {assets,manifest,root}=createSandboxAssets();
const stamp=new Date().toISOString().replaceAll(':','-').replaceAll('.','-');
const out=resolve(root,'target/p14-browser-performance-review',stamp);mkdirSync(out,{recursive:true});
const terms=28;
const source=Array.from({length:terms},(_,index)=>`x_{${index+1}}`).join('+');
const bootstrap=String.raw`import init,{EngineRuntime} from '/assets/core.js';
import {EngineEditorHost} from '/host/editor-host.js';import {VerifiedAssetLoader} from '/host/resources.js';
const metrics={dispatch:{},actions:{},frameSequence:0,lastFrameAt:0,errors:[]};let raw,host,loader;
const bump=(map,key)=>map[key]=(map[key]??0)+1;
class CountingRuntime{dispatch(request){let command=null;try{command=JSON.parse(request).command}catch{};if(command){bump(metrics.dispatch,command.op);if(command.op==='editorAction')bump(metrics.actions,command.action?.kind??'unknown');}
 const reply=raw.dispatch(request);if(command?.op==='editorAction'){requestAnimationFrame(t=>{metrics.frameSequence++;metrics.lastFrameAt=t;});}return reply;}}
window.ready=(async()=>{const manifest=await(await fetch('/manifest.json')).json();loader=new VerifiedAssetLoader(manifest,{baseUrl:new URL('/assets/',location.href).href,expectedBuildId:manifest.buildId,expectedEngineProtocol:1});await init({module_or_path:(await loader.load('core')).bytes});raw=new EngineRuntime();const fonts={};for(const asset of manifest.assets.filter(a=>a.kind==='font')){const r=JSON.parse(raw.registerFont((await loader.load(asset.id)).bytes,0));if(r.status!=='ok')throw Error(JSON.stringify(r));fonts[asset.id]=r.result.font.id;}const runtime=new CountingRuntime();host=new EngineEditorHost(runtime,document.getElementById('mount'),{documentId:'p14-browser-performance',initialSource:${JSON.stringify(source)},layout:{mathFont:fonts.math,textFont:fonts.text,mathFallbacks:[],style:'display',fontSizePx:28,paddingEm:.125},onError:e=>metrics.errors.push(String(e))});host.focus();window.review={inspect:()=>host.inspect(),metrics:()=>structuredClone(metrics),focus:()=>host.focus(),dispose:()=>{host.dispose();const active=raw.activeSessions;raw.free();loader.dispose();return{active}}};return true;})();`;
assets.set('/p14-performance.js',{type:'text/javascript; charset=utf-8',bytes:Buffer.from(bootstrap)});
const html='<!doctype html><html><head><meta charset="utf-8"><title>P14 browser performance</title><style>body{font:16px sans-serif;margin:24px}.editor{position:relative;min-height:150px;padding:16px;border:1px solid #888;overflow:auto}</style></head><body><div id="mount" class="editor"></div><script type="module" src="/p14-performance.js"></script></body></html>';
const percentile=(values,p)=>{const sorted=[...values].sort((a,b)=>a-b);return sorted[Math.min(sorted.length-1,Math.floor(sorted.length*p))]};
const summarize=values=>({min:Math.min(...values),median:percentile(values,.5),p95:percentile(values,.95),max:Math.max(...values)});
let report;
try{report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
 await evaluate('ready');await send('Performance.enable');
 const inspect=()=>evaluate('review.inspect()'),counts=()=>evaluate('review.metrics()');
 const pageHeap=async()=>{try{await send('HeapProfiler.collectGarbage')}catch{};const m=await send('Performance.getMetrics');const map=Object.fromEntries(m.metrics.map(x=>[x.name,x.value]));return{jsHeapUsedSize:map.JSHeapUsedSize??null,jsHeapTotalSize:map.JSHeapTotalSize??null,nodes:map.Nodes??null,documents:map.Documents??null,layoutCount:map.LayoutCount??null,recalcStyleCount:map.RecalcStyleCount??null};};
 async function key(key,code,vk,text){for(const type of ['keyDown','keyUp'])await send('Input.dispatchKeyEvent',{type,key,code,windowsVirtualKeyCode:vk,nativeVirtualKeyCode:vk,...(type==='keyDown'&&text!==undefined?{text}:{})});}
 async function operation(action){const before=(await counts()).frameSequence;const start=performance.now();await action();await evaluate(`(async()=>{const before=${before};while(review.metrics().frameSequence===before)await new Promise(requestAnimationFrame);return review.metrics().lastFrameAt;})()`);return performance.now()-start;}
 const initial=await inspect();assert.ok(initial.document.analysis.facts.nodes>=150&&initial.document.analysis.facts.nodes<=220,`unexpected fixture node count ${initial.document.analysis.facts.nodes}`);
 await key('End','End',35);for(let i=0;i<10;i++){await operation(()=>key('q','KeyQ',81,'q'));await operation(()=>key('Backspace','Backspace',8));}
 const beforeHeap=await pageHeap();const editMs=[];for(let i=0;i<60;i++){editMs.push(await operation(()=>key('q','KeyQ',81,'q')));editMs.push(await operation(()=>key('Backspace','Backspace',8)));}
 const afterEdits=await inspect();assert.equal(afterEdits.document.source,initial.document.source);const beforeMoves=await counts();const moveRevision=afterEdits.document.revision,moveSource=afterEdits.document.source;const moveMs=[];for(let i=0;i<100;i++){moveMs.push(await operation(()=>key(i%2?'ArrowRight':'ArrowLeft',i%2?'ArrowRight':'ArrowLeft',i%2?39:37)));}
 const afterMoves=await inspect(),afterMoveCounts=await counts();assert.equal(afterMoves.document.revision,moveRevision);assert.equal(afterMoves.document.source,moveSource);assert.equal(afterMoveCounts.dispatch.layout??0,beforeMoves.dispatch.layout??0,'cursor movement requested a fresh layout');assert.equal(afterMoveCounts.errors.length,0,JSON.stringify(afterMoveCounts.errors));
 const afterHeap=await pageHeap();const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});writeFileSync(resolve(out,'p14-performance.png'),Buffer.from(screenshot.data,'base64'));
 const dispose=await evaluate('review.dispose()');assert.equal(dispose.active,0);const disposedHeap=await pageHeap();
 const editStats=summarize(editMs),moveStats=summarize(moveMs);assert.ok(editStats.p95<=32,`local edit p95 ${editStats.p95.toFixed(2)}ms exceeds 32ms target`);assert.ok(moveStats.p95<=20,`browser event-to-frame cursor p95 ${moveStats.p95.toFixed(2)}ms exceeds one-frame guard`);
 return{status:'PASS',browser:browserVersion.product,hostBuild:manifest.buildId,wasmSha256:manifest.assets.find(a=>a.id==='core').sha256,fixture:{terms,nodes:initial.document.analysis.facts.nodes,sourceBytes:new TextEncoder().encode(source).length},localEditEventToFrameMs:editStats,cursorEventToFrameMs:moveStats,localEditTargetP95Ms:32,cursorCoreTargetMs:8,browserFrameGuardP95Ms:20,cursorMoveLayoutDelta:(afterMoveCounts.dispatch.layout??0)-(beforeMoves.dispatch.layout??0),cursorMoveRevisionDelta:afterMoves.document.revision-moveRevision,runtimeCalls:{beforeMoves,afterMoves:afterMoveCounts},heap:{before:beforeHeap,afterEditsAndMoves:afterHeap,afterDispose:disposedHeap},disposeActiveSessions:dispose.active,actualBrowserKeyboard:true,screenshot:relative(root,resolve(out,'p14-performance.png')).replaceAll('\\','/'),interpretation:'Cursor timing includes browser input dispatch and the next animation frame; the <=8ms core-only target is not inferred from this frame-paced measurement. Zero layout calls/revision changes during cursor movement is asserted directly.'};
});}catch(error){report={status:'FAILED',error:String(error)};throw error;}finally{if(report){writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({receipt:relative(root,resolve(out,'report.json')).replaceAll('\\','/'),...report},null,2));}}
