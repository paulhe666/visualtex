// P08 long-document viewport acceptance. One thousand passive formula shells
// are scrolled through a real Edge viewport; only nearby formulas own SVG/input
// DOM. Core sessions park/resume without source reconstruction.
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {mkdirSync,writeFileSync} from 'node:fs';
import {resolve,relative} from 'node:path';
import {createSandboxAssets} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';

const {assets,manifest,root}=createSandboxAssets();
const stamp=new Date().toISOString().replaceAll(':','-').replaceAll('.','-');
const out=resolve(root,'target/virtual-viewport-browser-review',stamp);mkdirSync(out,{recursive:true});
const sha=b=>createHash('sha256').update(b).digest('hex');
const bootstrap=String.raw`
import init,{EngineRuntime} from '/assets/core.js';
import {VerifiedAssetLoader} from '/host/resources.js';
import {VirtualFormulaViewport} from '/host/virtual-formula-viewport.js';
let runtime,loader,viewport;const errors=[];
const formulaSource=i=>{
 if(i>0&&i%100===10)return '\\begin{matrix}a_{'+i+'}&b\\\\c&d\\end{matrix}';
 if(i>0&&i%100===20)return '\\begin{cases}x_{'+i+'}&x>0\\\\-x&x<0\\end{cases}';
 if(i>0&&i%100===30)return '\\begin{aligned}a_{'+i+'}&=b+c\\\\d&=e+f\\end{aligned}';
 if(i>0&&i%100===40)return '\\frac{x_{'+i+'}+1}{y+2}';
 return i===0?'x':'x_{'+i+'}+1';
};
const ready=(async()=>{
 const manifest=await(await fetch('/manifest.json',{cache:'no-store'})).json();
 loader=new VerifiedAssetLoader(manifest,{baseUrl:new URL('/assets/',location.href).href,expectedBuildId:manifest.buildId,expectedEngineProtocol:1});
 const core=await loader.load('core');await init({module_or_path:core.bytes});runtime=new EngineRuntime();
 const fonts={};for(const asset of manifest.assets.filter(a=>a.kind==='font')){const verified=await loader.load(asset.id);const r=JSON.parse(runtime.registerFont(verified.bytes,0));if(r.status!=='ok')throw Error(JSON.stringify(r));fonts[asset.id]=r.result.font.id;}
 const records=Array.from({length:1000},(_,i)=>({documentId:'virtual-formula-'+i,source:formulaSource(i)}));
 viewport=new VirtualFormulaViewport(runtime,document.getElementById('scroll'),records,{
  layout:{mathFont:fonts.math,textFont:fonts.text,mathFallbacks:[],style:'display',fontSizePx:22,paddingEm:.125},
  overscanPx:120,minHeightPx:76,maxMounted:16,onError:(index,error)=>errors.push({index,message:String(error)})
 });
 await new Promise(r=>setTimeout(r,120));
})();
const wait=ms=>new Promise(r=>setTimeout(r,ms));
window.__virtualReview={ready,wait,
 stats:()=>({...viewport.stats,activeSessions:runtime.activeSessions,parkedSessions:runtime.parkedSessions,parkedWeight:runtime.parkedWeight,
  visibleCarets:[...document.querySelectorAll('.vt-core-caret')].filter(e=>getComputedStyle(e).visibility!=='hidden').length,errors:errors.slice()}),
 inspect:i=>viewport.inspect(i),identity:i=>viewport.sessionIdentity(i),source:i=>viewport.source(i),focus:i=>viewport.focus(i),
 shell:i=>{const e=viewport.shell(i);return {top:e.offsetTop,height:e.getBoundingClientRect().height,mounted:e.dataset.mounted??null};},
 scrollTo:async i=>{viewport.shell(i).scrollIntoView({block:'center'});await wait(80);viewport.refresh();await wait(30);return viewport.stats;},
 scan:async()=>{
  const scroller=document.getElementById('scroll');scroller.focus();await wait(30);
  const max=Math.max(0,scroller.scrollHeight-scroller.clientHeight),step=Math.max(120,Math.floor(scroller.clientHeight*.72));
  let peakMounted=0,peakInputs=0,peakSvgs=0,peakActive=0,peakParked=0,samples=0;
  const started=performance.now();
  for(let y=0;y<=max;y+=step){scroller.scrollTop=Math.min(y,max);await wait(18);viewport.refresh();
   const s=viewport.stats;peakMounted=Math.max(peakMounted,s.mounted);peakInputs=Math.max(peakInputs,s.inputDom);peakSvgs=Math.max(peakSvgs,s.svgDom);
   peakActive=Math.max(peakActive,runtime.activeSessions);peakParked=Math.max(peakParked,runtime.parkedSessions);samples++;
   if(s.mounted>16||s.inputDom!==s.mounted||s.svgDom!==s.mounted||runtime.activeSessions!==s.mounted)throw Error('Virtual owner/DOM budget diverged during scan');
  }
  scroller.scrollTop=max;await wait(100);viewport.refresh();await wait(30);
  const missed=[];for(let i=0;i<1000;i++)if(viewport.sessionIdentity(i)===null)missed.push(i);
  for(const i of missed){viewport.shell(i).scrollIntoView({block:'center'});await wait(35);viewport.refresh();await wait(15);
   const s=viewport.stats;peakMounted=Math.max(peakMounted,s.mounted);peakInputs=Math.max(peakInputs,s.inputDom);peakSvgs=Math.max(peakSvgs,s.svgDom);
   peakActive=Math.max(peakActive,runtime.activeSessions);peakParked=Math.max(peakParked,runtime.parkedSessions);
   if(s.mounted>16||s.inputDom!==s.mounted||s.svgDom!==s.mounted||runtime.activeSessions!==s.mounted)throw Error('Virtual owner/DOM budget diverged while verifying skipped shells');
  }
  return {durationMs:Math.round(performance.now()-started),peakMounted,peakInputs,peakSvgs,peakActive,peakParked,samples,maxScroll:max,discreteScrollMisses:missed,final:window.__virtualReview.stats()};
 },
 dispose:()=>{viewport.dispose();return {activeSessions:runtime.activeSessions,parkedSessions:runtime.parkedSessions,parkedWeight:runtime.parkedWeight,inputs:document.querySelectorAll('.vt-core-input').length,svgs:document.querySelectorAll('.vt-core-svg').length};}
};
window.addEventListener('pagehide',()=>{try{viewport?.dispose();}finally{runtime?.free();loader?.dispose();}},{once:true});
ready.catch(error=>{errors.push({index:-1,message:String(error)});document.getElementById('error').textContent=String(error);});
`;
assets.set('/virtual-review.js',{type:'text/javascript; charset=utf-8',bytes:Buffer.from(bootstrap)});
const html=`<!doctype html><html><head><meta charset="utf-8"><title>Virtual formula viewport review</title>
<style>body{font:14px Arial;margin:18px}#scroll{height:480px;overflow:auto;border:1px solid #888;position:relative;outline:none}.vt-core-virtual-formula{border-bottom:1px solid #eee;padding:4px 12px;box-sizing:border-box}.vt-core-scene{min-height:36px}.vt-core-position{font-size:11px;color:#555}</style></head>
<body><h1>1000 formula virtual viewport</h1><p id="error"></p><div id="scroll" tabindex="0"></div><script type="module" src="/virtual-review.js"></script></body></html>`;
const steps=[];let report;
try{
 report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
  await evaluate(`new Promise((resolve,reject)=>{const until=Date.now()+5000;const poll=()=>{if(window.__virtualReview)return window.__virtualReview.ready.then(()=>resolve(true),reject);if(Date.now()>until)return reject(new Error(document.getElementById('error')?.textContent||'Virtual review bootstrap did not load'));setTimeout(poll,25)};poll()})`);
  const stats=()=>evaluate('window.__virtualReview.stats()');
  const inspect=i=>evaluate(`window.__virtualReview.inspect(${i})`);
  const identity=i=>evaluate(`window.__virtualReview.identity(${i})`);
  async function key(key,code,n,modifiers=0,text){
   await send('Input.dispatchKeyEvent',{type:'keyDown',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers,...(text?{text}:{})});
   await send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers});
  }
  async function wait(ms=80){await evaluate(`window.__virtualReview.wait(${ms}).then(()=>true)`);}
  let s=await stats();assert.equal(s.formulas,1000);assert.ok(s.mounted>0&&s.mounted<=16,s);assert.equal(s.inputDom,s.mounted);assert.equal(s.svgDom,s.mounted);assert.equal(s.activeSessions,s.mounted);assert.equal(s.parkedSessions,0);assert.equal(s.visibleCarets,0);assert.deepEqual(s.errors,[]);
  steps.push({label:'1000 passive shells initialize with bounded live editor DOM',stats:s});

  assert.equal(await evaluate('window.__virtualReview.focus(0)'),true);await key('+','Equal',187,8,'+');await key('1','Digit1',49,0,'1');
  const edited0=await inspect(0),id0=await identity(0);assert.equal(edited0.document.source,'x+1');assert.ok(id0&&!id0.parked);assert.ok(edited0.undoDepth>=1);assert.equal((await stats()).visibleCarets,1);
  steps.push({label:'Formula 0 edited before leaving viewport',handle:id0.handle,revision:edited0.document.revision,undoDepth:edited0.undoDepth});

  await evaluate('window.__virtualReview.scrollTo(500).then(()=>true)');assert.equal(await evaluate('window.__virtualReview.focus(500)'),true);await key('q','KeyQ',81,0,'q');
  const edited500=await inspect(500),id500=await identity(500);assert.ok(edited500.document.source.endsWith('q'));assert.ok(id500&&!id500.parked);
  await evaluate('document.getElementById("scroll").focus()');await wait(100);
  const parked0=await identity(0),parked500=await identity(500);assert.ok(parked0?.parked);assert.ok(!parked500?.parked||await evaluate('document.getElementById("scroll").contains(document.activeElement)'));
  assert.deepEqual(await inspect(0),edited0,'Formula 0 snapshot changed while parked');
  s=await stats();assert.ok(s.mounted<=16);assert.equal(s.inputDom,s.mounted);assert.equal(s.activeSessions,s.mounted);assert.deepEqual(s.errors,[]);
  steps.push({label:'Offscreen formula parks exact editor state after focus moves',formula0:parked0,formula500:id500,stats:s});

  const scan=await evaluate('window.__virtualReview.scan()');assert.ok(scan.peakMounted<=16,scan);assert.equal(scan.peakInputs,scan.peakMounted);assert.equal(scan.peakSvgs,scan.peakMounted);assert.ok(scan.peakActive<=16,scan);
  assert.equal(scan.final.materialized,1000,'Full viewport scan did not visit every formula shell');assert.ok(scan.final.parkedSessions>=980,scan.final);assert.ok(scan.final.mounted<=16);assert.equal(scan.final.activeSessions,scan.final.mounted);assert.deepEqual(scan.final.errors,[]);
  steps.push({label:'Full document scroll materializes all 1000 formulas without accumulating editable DOM',scan});

  await evaluate('window.__virtualReview.scrollTo(0).then(()=>true)');await wait(50);const resumed0=await inspect(0),resumedId0=await identity(0);
  assert.equal(resumedId0.handle,id0.handle,'Formula handle changed after viewport remount');assert.equal(resumedId0.parked,false);assert.deepEqual(resumed0,edited0,'Formula 0 source/tree/selection/history changed across park/resume');
  assert.equal(await evaluate('window.__virtualReview.focus(0)'),true);await key('z','KeyZ',90,2);assert.equal((await inspect(0)).document.source,'x');await key('y','KeyY',89,2);
  const redone0=await inspect(0);assert.equal(redone0.document.source,'x+1');assert.deepEqual(redone0.document.analysis.tree,edited0.document.analysis.tree,'Redo after resume changed mathematical identities');
  steps.push({label:'Remounted formula keeps handle, full snapshot and undo/redo identities',handle:resumedId0.handle,revision:redone0.document.revision});

  // Protocol composition is not physical Windows IME. It proves only that a
  // live preedit pins the formula instead of being parked underneath the OS UI.
  await send('Input.imeSetComposition',{text:'sudu',selectionStart:4,selectionEnd:4});assert.ok((await inspect(0)).composition);
  await evaluate('window.__virtualReview.scrollTo(900).then(()=>true)');await wait(80);const duringIme=await identity(0);assert.ok(duringIme&&!duringIme.parked,'Composing formula was parked');
  assert.equal((await inspect(0)).document.source,'x+1');await send('Input.imeSetComposition',{text:'',selectionStart:0,selectionEnd:0});await evaluate('window.__virtualReview.scrollTo(900).then(()=>true)');assert.equal(await evaluate('window.__virtualReview.focus(900)'),true);await wait(100);
  const afterIme=await evaluate(`({identity:window.__virtualReview.identity(0),snapshot:window.__virtualReview.inspect(0),stats:window.__virtualReview.stats(),active:{tag:document.activeElement?.tagName,id:document.activeElement?.id,className:document.activeElement?.className,formula:document.activeElement?.closest?.('.vt-core-virtual-formula')?.dataset?.formulaIndex??null},scrollTop:document.getElementById('scroll').scrollTop})`);
  steps.push({label:'Protocol composition pin release diagnostic',duringIme,afterIme});
  assert.ok(afterIme.identity?.parked,'Formula did not park after composition ended and focus left');
  steps.push({label:'Protocol composition pins an offscreen formula until preedit is cancelled',identity:duringIme});

  s=await stats();assert.ok(s.mounted<=16);assert.equal(s.inputDom,s.mounted);assert.equal(s.svgDom,s.mounted);assert.equal(s.activeSessions,s.mounted);assert.equal(s.visibleCarets,1);assert.deepEqual(s.errors,[]);
  const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:false});writeFileSync(resolve(out,'virtual-1000-formulas.png'),Buffer.from(screenshot.data,'base64'));
  const disposed=await evaluate('window.__virtualReview.dispose()');assert.deepEqual(disposed,{activeSessions:0,parkedSessions:0,parkedWeight:0,inputs:0,svgs:0});
  return {status:'PASS',browser:browserVersion.product,hostBuild:manifest.buildId,wasmSha256:manifest.assets.find(a=>a.id==='core').sha256,
   formulaCount:1000,maxMountedConfigured:16,steps,scan,finalBeforeDispose:s,disposed,actualBrowserScroll:true,actualTrustedKeyboard:true,
   protocolComposition:true,physicalWindowsIme:false,screenshot:relative(root,resolve(out,'virtual-1000-formulas.png')).replaceAll('\\','/')};
 });
}catch(error){report={status:'FAILED',error:String(error),steps,hostBuild:manifest.buildId};throw error;}
finally{if(report){writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({receipt:relative(root,resolve(out,'report.json')).replaceAll('\\','/'),...report},null,2));}}
