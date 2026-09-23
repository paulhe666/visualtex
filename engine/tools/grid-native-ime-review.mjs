// P08 physical Windows IME acceptance for contained Cases and Aligned grids.
// Keystrokes are guarded Win32 SendInput into one UUID/PID-owned Edge window.
// No CDP imeSetComposition, Unicode injection, clipboard, Word or user browser.
import assert from 'node:assert/strict';
import {createHash,randomUUID} from 'node:crypto';
import {mkdirSync,writeFileSync} from 'node:fs';
import {spawnSync} from 'node:child_process';
import {resolve,relative} from 'node:path';
import {setTimeout as sleep} from 'node:timers/promises';
import {createSandboxAssets} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';

const {assets,manifest,root}=createSandboxAssets();
const title=`VisualTeX-IME-${randomUUID()}`;
const stamp=new Date().toISOString().replaceAll(':','-').replaceAll('.','-');
const out=resolve(root,'target/grid-native-ime-review',stamp);mkdirSync(out,{recursive:true});
const sha=b=>createHash('sha256').update(b).digest('hex');
const casesSource=String.raw`\begin{cases}\text{}&x>0\\b&x<0\end{cases}`;
const alignedSource=String.raw`\begin{aligned}\text{}&=b\\c&=d\end{aligned}`;
const steps=[];
const bootstrap=`import init,{EngineRuntime} from '/assets/core.js';
import {EngineEditorHost} from '/host/editor-host.js';
import {VerifiedAssetLoader} from '/host/resources.js';
const errors=[];let runtime,loader,host;window.enterCalls=0;
window.ready=(async()=>{const manifest=await(await fetch('/manifest.json',{cache:'no-store'})).json();loader=new VerifiedAssetLoader(manifest,{baseUrl:new URL('/assets/',location.href).href,expectedBuildId:manifest.buildId,expectedEngineProtocol:1});
 const core=await loader.load('core');await init({module_or_path:core.bytes});runtime=new EngineRuntime();const fonts={};
 for(const asset of manifest.assets.filter(a=>a.kind==='font')){const verified=await loader.load(asset.id);const r=JSON.parse(runtime.registerFont(verified.bytes,0));if(r.status!=='ok')throw Error(JSON.stringify(r));fonts[asset.id]=r.result.font.id;}
 host=new EngineEditorHost(runtime,document.getElementById('mount'),{documentId:'p08-grid-native-ime',initialSource:${JSON.stringify(casesSource)},layout:{mathFont:fonts.math,textFont:fonts.text,mathFallbacks:[],style:'display',fontSizePx:28,paddingEm:.125},onChange:s=>document.getElementById('source').textContent=s.document.source,onError:e=>errors.push(String(e)),onEnter:()=>window.enterCalls++});window.host=host;return true;})();
window.review={inspect:()=>host.inspect(),scene:()=>host.layoutSnapshot,events:()=>host.input.events,errors:()=>errors.slice(),enters:()=>window.enterCalls,replace:s=>host.replaceSource(s),focus:()=>host.focus(),dispose:()=>{host.dispose();runtime.free();loader.dispose();}};
window.addEventListener('pagehide',()=>{try{host?.dispose();}finally{runtime?.free();loader?.dispose();}},{once:true});`;
assets.set('/grid-native-ime.js',{type:'text/javascript; charset=utf-8',bytes:Buffer.from(bootstrap)});
const html=`<!doctype html><html><head><meta charset="utf-8"><title>${title}</title>
<style>body{margin:44px;font:18px sans-serif;color:#171717}#mount{margin:60px 24px;min-height:180px}.vt-core-position{margin-top:28px;font-size:14px}.vt-core-preedit{background:white;border-bottom:1px solid black;color:black}pre{padding:12px;background:#eee}</style></head>
<body><h2>P08 Windows IME — contained grids</h2><p>Only this owned Edge window receives test keys. Word is not involved.</p><div id="mount"></div><pre id="source"></pre><script type="module" src="/grid-native-ime.js"></script></body></html>`;

let report;
try{
 report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion,ownedProcessId})=>{
  let initialLayout;let shiftToggleUsed=false;
  const native=(action,extra=[])=>{
    const p=spawnSync('powershell.exe',['-NoProfile','-File',resolve(root,'tools/owned-ime-input.ps1'),'-OwnedProcessId',String(ownedProcessId),'-ExpectedTitle',title,'-Action',action,...extra],{encoding:'utf8',windowsHide:true,timeout:action==='activate'?35000:20000,maxBuffer:1024*1024});
    assert.equal(p.error,undefined);assert.equal(p.status,0,p.stderr||p.stdout);const result=JSON.parse(p.stdout.replace(/^\uFEFF/,''));steps.push({nativeAction:action,args:extra,result});return result;
  };
  const inspect=()=>evaluate(`(()=>{const editor=review.inspect(),scene=review.scene(),input=host.input.textarea,focus=editor.selection.kind==='range'?editor.selection.focus:null;const stop=focus?scene?.carets.find(c=>c.position.kind===focus.kind&&c.position.nodeId===focus.nodeId&&(focus.kind==='text'?c.position.grapheme===focus.grapheme:c.position.offset===focus.offset)):null;return {editor,events:review.events(),errors:review.errors(),enterCalls:review.enters(),inputFocused:document.activeElement===input,expected:stop?host.view.layoutToClient(stop.x,stop.top):null,proxy:(()=>{const r=input.getBoundingClientRect();return {left:r.left,top:r.top,width:r.width,height:r.height,scrollLeft:input.scrollLeft,scrollTop:input.scrollTop}})()};})()`);
  const grid=s=>s.editor.document.analysis.tree.entries.find(e=>e.node.kind==='grid')?.node;
  const activate=()=>{native('activate');};
  async function focusEmptyText(source){
    await evaluate(`review.replace(${JSON.stringify(source)});true`);
    const point=await evaluate(`(()=>{const s=review.inspect(),scene=review.scene();let id=null;s.document.analysis.tree.entries[0].node.rows[0].cells[0].entries[0].node.children;
      s.document.analysis.tree.entries[0].node.rows[0].cells[0].entries[0].node.kind==='text'&&(id=s.document.analysis.tree.entries[0].node.rows[0].cells[0].entries[0].node.id);
      if(id===null)throw Error('Expected empty text node');const c=scene.carets.find(c=>c.position.kind==='text'&&c.position.nodeId===id&&c.position.grapheme===0);if(!c)throw Error('Missing text caret');return host.view.layoutToClient(c.x,(c.top+c.bottom)/2);})()`);
    await send('Page.bringToFront');activate();
    await send('Input.dispatchMouseEvent',{type:'mousePressed',...point,button:'left',clickCount:1});await send('Input.dispatchMouseEvent',{type:'mouseReleased',...point,button:'left',clickCount:1});
    const s=await inspect();assert.equal(s.inputFocused,true);assert.equal(s.editor.selection.focus.kind,'text');return s;
  }
  async function beginPinyin(source){
    let before=await focusEmptyText(source);native('type',['-Text','sudu']);await sleep(220);let composing=await inspect();
    if(!composing.editor.composition){
      steps.push({label:'IME initially in Latin mode; toggling owned Chinese IME mode',source:composing.editor.document.source});
      before=await focusEmptyText(source);native('key',['-Key','shift']);shiftToggleUsed=true;native('type',['-Text','sudu']);await sleep(220);composing=await inspect();
    }
    assert.ok(composing.editor.composition,'Native pinyin did not enter Windows IME composition');
    assert.equal(composing.editor.document.source,before.editor.document.source,'Preedit mutated committed formula');
    assert.ok(composing.events.some(e=>e.type==='compositionstart'&&e.trusted));assert.ok(composing.events.some(e=>e.type==='compositionupdate'&&e.trusted));
    assert.equal(composing.proxy.scrollLeft,0);assert.ok(composing.proxy.width>32);assert.ok(Math.abs(composing.proxy.left-composing.expected.x)<1);
    return {before,composing};
  }
  async function ensureChineseLayout(){
    const env=native('inspect');initialLayout=env.beforeLayout;const chinese=env.loadedLayouts.find(l=>(l.languageId&1023)===4);assert.ok(chinese,'An already-loaded Chinese input locale is required');if(chinese.handle!==env.beforeLayout)native('layout',['-KeyboardLayout',String(chinese.handle)]);return env;
  }
  async function physicalGridIme(source,environment){
    const {before,composing}=await beginPinyin(source);const initialGrid=grid(before),rows=initialGrid.rows.length,columns=initialGrid.columns.slice(),sourceBefore=before.editor.document.source,outerBefore=before.enterCalls;
    native('key',['-Key','shiftenter']);await sleep(220);let afterChord=await inspect();
    assert.equal(grid(afterChord).rows.length,rows,`${environment}: Shift+Enter during preedit inserted a row`);assert.deepEqual(grid(afterChord).columns,columns);assert.equal(afterChord.enterCalls,outerBefore,`${environment}: preedit Shift+Enter reached outer Enter`);
    assert.equal(afterChord.editor.document.analysis.facts.grids,1);assert.ok(afterChord.editor.document.analysis.complete&&afterChord.editor.document.analysis.supported&&afterChord.editor.document.analysis.exportable,`${environment}: IME chord damaged the contained grid source`);
    steps.push({label:`${environment} native IME preedit owns Shift+Enter`,before:sourceBefore,after:afterChord.editor.document.source,compositionStillActive:Boolean(afterChord.editor.composition),rows,selection:afterChord.editor.selection});
    if(afterChord.editor.composition){native('key',['-Key','escape']);await sleep(140);} // reset deterministically for explicit commit proof
    const fresh=await beginPinyin(source),freshGrid=grid(fresh.before),freshRows=freshGrid.rows.length,freshColumns=freshGrid.columns.slice();native('key',['-Key','space']);await sleep(180);const committed=await inspect();
    assert.equal(committed.editor.composition,null);assert.ok(committed.editor.document.source.includes(String.raw`\text{速度}`));assert.equal(grid(committed).rows.length,freshRows);assert.deepEqual(grid(committed).columns,freshColumns);assert.equal(committed.editor.undoDepth,fresh.before.editor.undoDepth+1);assert.equal(committed.errors.length,0);
    steps.push({label:`${environment} native IME commit is one grid-local edit`,source:committed.editor.document.source,rows,undoDepth:committed.editor.undoDepth});
    return committed;
  }
  await evaluate('window.ready.then(()=>true)');await send('Page.bringToFront');await ensureChineseLayout();
  try {
    const casesCommitted=await physicalGridIme(casesSource,'cases');
    // Cases ordinary Enter is an internal row command after IME is complete.
    activate();native('key',['-Key','enter']);await sleep(140);const casesEnter=await inspect();assert.equal(grid(casesEnter).rows.length,grid(casesCommitted).rows.length+1);assert.equal(casesEnter.enterCalls,casesCommitted.enterCalls);assert.equal(casesEnter.editor.document.analysis.facts.grids,1);
    steps.push({label:'cases physical Enter adds one internal row only',rows:grid(casesEnter).rows.length});

    const alignedCommitted=await physicalGridIme(alignedSource,'aligned');
    // Aligned ordinary Enter remains the outer callback; Shift+Enter is internal.
    activate();native('key',['-Key','enter']);await sleep(140);const alignedEnter=await inspect();assert.equal(grid(alignedEnter).rows.length,grid(alignedCommitted).rows.length);assert.equal(alignedEnter.enterCalls,alignedCommitted.enterCalls+1);
    activate();native('key',['-Key','shiftenter']);await sleep(140);const alignedBreak=await inspect();assert.equal(grid(alignedBreak).rows.length,grid(alignedCommitted).rows.length+1);assert.equal(alignedBreak.enterCalls,alignedEnter.enterCalls);assert.deepEqual(grid(alignedBreak).columns,grid(alignedCommitted).columns);
    steps.push({label:'aligned physical Enter stays outer while Shift+Enter adds one internal row',rows:grid(alignedBreak).rows.length,enterCalls:alignedBreak.enterCalls});

    const final=await inspect();assert.equal(final.errors.length,0);assert.ok(final.events.some(e=>e.type==='keydown'&&e.key==='Enter'&&e.trusted));
    report={status:'PASS',browser:browserVersion.product,hostBuild:manifest.buildId,wasmSha256:manifest.assets.find(a=>a.id==='core').sha256,actualWindowsIme:true,keySource:'Win32 SendInput ASCII pinyin and guarded Shift+Enter into UUID/PID-owned Edge',humanPhysicalKeyboard:false,cdpCompositionUsed:false,unicodeInjectionUsed:false,clipboardUsed:false,cases:true,aligned:true,stepsCount:steps.length,steps,notCovered:['WebView2 embedding','Office UI','macOS IME','production default switch']};
    return report;
  } finally {
    if(shiftToggleUsed){try{const s=await inspect();if(s.editor.composition)native('key',['-Key','escape']);native('key',['-Key','shift']);}catch{}}
    if(initialLayout!==undefined){try{native('layout',['-KeyboardLayout',String(initialLayout)]);}catch{}}
    try{await evaluate('review.dispose();true');}catch{}
  }
 },{visible:true});
}catch(error){report={status:'FAILED',error:String(error),steps};throw error;}
finally{writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({receipt:relative(root,resolve(out,'report.json')).replaceAll('\\','/'),status:report?.status,error:report?.error??null,steps:steps.length},null,2));}
