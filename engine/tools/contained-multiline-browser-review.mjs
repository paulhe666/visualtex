// Independent P08 acceptance using two real editor hosts and browser keys.
// Only setup/inspection is scripted. All document edits occur through visible
// controls, keyboard or pointer input. No user browser or Office is attached.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve, relative } from 'node:path';
import { createSandboxAssets } from './sandbox-assets.mjs';
import { withOwnedBrowser } from './owned-browser-review.mjs';

const { assets, manifest, root } = createSandboxAssets();
const stamp = new Date().toISOString().replaceAll(':', '-').replaceAll('.', '-');
const out = resolve(root, 'target/contained-multiline-review', stamp);
mkdirSync(out, { recursive: true });
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const bootstrap = String.raw`
import init, { EngineRuntime } from '/assets/core.js';
import { EngineEditorHost } from '/host/editor-host.js';
import { VerifiedAssetLoader } from '/host/resources.js';
const hosts = {}, errors = [], outerEnter = { A:0, B:0 }, exports = {};
let runtime, loader;
const ready = (async () => {
  const manifest = await (await fetch('/manifest.json', { cache:'no-store' })).json();
  loader = new VerifiedAssetLoader(manifest, {
    baseUrl: new URL('/assets/',location.href).href,
    expectedBuildId: manifest.buildId, expectedEngineProtocol:1,
  });
  const core = await loader.load('core');
  await init({ module_or_path:core.bytes });
  runtime = new EngineRuntime();
  const fontIds = {};
  for (const asset of manifest.assets.filter(a => a.kind === 'font')) {
    const verified = await loader.load(asset.id);
    const reply = JSON.parse(runtime.registerFont(verified.bytes,0));
    if (reply.status !== 'ok' || reply.result.font.sha256 !== asset.sha256) throw new Error('Font identity mismatch');
    fontIds[asset.id] = reply.result.font.id;
  }
  for (const id of ['A','B']) {
    const parent = document.getElementById('editor-'+id);
    const source = document.getElementById('source-'+id);
    const host = new EngineEditorHost(runtime,parent,{
      documentId:'contained-review-'+id, initialSource:id==='A'?'':'z_2',
      layout:{ mathFont:fontIds.math,textFont:fontIds.text,
        mathFallbacks:fontIds.symbols?[fontIds.symbols]:[],style:'display',fontSizePx:25,paddingEm:.125 },
      onChange:snapshot => { source.value=snapshot.document.source; },
      onError:error => { errors.push({ frame:id,message:String(error) }); },
      onEnter:() => { outerEnter[id]++; },
    });
    hosts[id]=host; source.value=host.inspect().document.source;
    source.addEventListener('input',() => {
      try { host.replaceSource(source.value); } catch (error) { errors.push({frame:id,message:String(error)}); }
    });
    for (const button of document.querySelectorAll('[data-frame="'+id+'"]')) {
      button.addEventListener('pointerdown',event => { if(event.button===0)event.preventDefault(); });
      button.addEventListener('click',() => {
        try {
          if(button.dataset.action==='focus')host.focus();
          if(button.dataset.action==='undo'){host.history(false);host.focus();}
          if(button.dataset.action==='redo'){host.history(true);host.focus();}
          if(button.dataset.action==='export'){
            exports[id]=host.exportSvg(); document.getElementById('export-'+id).textContent=exports[id];
          }
        } catch (error) { errors.push({frame:id,message:String(error)}); }
      });
    }
    document.getElementById('zoom-'+id).addEventListener('change',event => {
      try { host.setZoom(Number(event.target.value)); } catch(error) { errors.push({frame:id,message:String(error)}); }
    });
  }
  hosts.A.focus();
})();
window.__containedReview = {
  ready,
  inspect:id=>hosts[id]?.inspect(),
  scene:id=>hosts[id]?.layoutSnapshot,
  stats:()=>({ sessions:runtime?.activeSessions,enters:{...outerEnter},errors:errors.slice() }),
  events:id=>hosts[id]?.input.events??[],
  exported:id=>exports[id]??null,
};
window.addEventListener('pagehide',()=>{
  for (const host of Object.values(hosts))host.dispose(); runtime?.free();loader?.dispose();
},{once:true});
ready.catch(error=>{document.getElementById('error').textContent=String(error);});
`;
assets.set('/review.js', {type:'text/javascript; charset=utf-8',bytes:Buffer.from(bootstrap)});
const section = id => `<section><h2>Independent formula ${id}</h2>
<div><button data-frame="${id}" data-action="focus" id="focus-${id}">Focus ${id}</button>
<button data-frame="${id}" data-action="undo" id="undo-${id}">Undo</button>
<button data-frame="${id}" data-action="redo" id="redo-${id}">Redo</button>
<button data-frame="${id}" data-action="export" id="svg-${id}">Export SVG</button>
<select id="zoom-${id}"><option value=".5">50%</option><option value="1" selected>100%</option><option value="2">200%</option></select></div>
<div id="editor-${id}" class="editor"></div><textarea id="source-${id}" aria-label="Formula ${id} source"></textarea>
<pre id="export-${id}"></pre></section>`;
const html = `<!doctype html><html><head><meta charset="utf-8"><title>Contained multiline independent review</title>
<style>body{font:16px Arial;margin:24px}section{border:1px solid #888;padding:16px;margin:18px 0}.editor{position:relative;min-height:130px;padding:14px 0;overflow:auto}textarea{width:96%;height:72px}button,select{margin-right:8px;padding:8px}pre{max-height:80px;overflow:auto;white-space:pre-wrap}</style></head>
<body><h1>One formula frame, multiple internal rows</h1><p id="error"></p>${section('A')}${section('B')}<script type="module" src="/review.js"></script></body></html>`;
const steps=[];
let report;
try {
 report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
  const inspect=id=>evaluate(`window.__containedReview.inspect(${JSON.stringify(id)})`);
  const stats=()=>evaluate('window.__containedReview.stats()');
  async function key(key,code,n,modifiers=0,text){
    await send('Input.dispatchKeyEvent',{type:'keyDown',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers,...(text?{text}:{})});
    await send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers});
  }
  async function click(selector){
    const p=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e)throw new Error('Missing review control');e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();const x=r.x+r.width/2,y=r.y+r.height/2;if(!e.contains(document.elementFromPoint(x,y)))throw new Error('Control outside actual viewport');return {x,y};})()`);
    for(const type of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type,...p,button:'left',clickCount:1});
  }
  async function type(text){for(const ch of text){
    if(/^[a-z]$/.test(ch))await key(ch,'Key'+ch.toUpperCase(),ch.toUpperCase().charCodeAt(0),0,ch);
    else if(/^[0-9]$/.test(ch))await key(ch,'Digit'+ch,ch.charCodeAt(0),0,ch);
    else if(ch==='+')await key('+','Equal',187,8,'+');
    else if(ch==='-')await key('-','Minus',189,0,'-');
    else if(ch==='=')await key('=','Equal',187,0,'=');
    else throw new Error('Unsupported test key '+ch);
  }}
  async function setSource(id,text){await click('#source-'+id);await key('a','KeyA',65,2);await send('Input.insertText',{text});}
  async function checkpoint(label){
    const a=await inspect('A');
    assert.equal(await evaluate('document.getElementById("source-A").value'),a.document.source);
    assert.deepEqual(await inspect('B'),bBefore,'Frame B changed during '+label);
    assert.equal((await stats()).sessions,2);
    steps.push({label,source:a.document.source,revision:a.document.revision,undoDepth:a.undoDepth,selection:a.selection,
      runtime:await stats(),activeInput:await evaluate('({tag:document.activeElement?.tagName,id:document.activeElement?.id,className:document.activeElement?.className})')});return a;
  }
  const getGrid=s=>s.document.analysis.tree.entries.find(e=>e.node.kind==='grid')?.node;
  const contentState=s=>({source:s.document.source,tree:s.document.analysis.tree,selection:s.selection,documentId:s.document.documentId});
  async function pointerCaret(id,nodeId,offset){
    const p=await evaluate(`(()=>{const s=window.__containedReview.scene(${JSON.stringify(id)});const c=s.carets.find(c=>c.position.kind==='row'&&c.position.nodeId===${nodeId}&&c.position.offset===${offset});if(!c)throw new Error('Missing grid caret');const editor=document.getElementById('editor-'+${JSON.stringify(id)});const svg=editor.querySelector('svg');svg.scrollIntoView({block:'center'});const p=new DOMPoint(c.x,(c.top+c.bottom)/2).matrixTransform(svg.getScreenCTM());if(!editor.contains(document.elementFromPoint(p.x,p.y)))throw new Error('Caret is not inside editor viewport');return{x:p.x,y:p.y};})()`);
    for(const kind of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type:kind,...p,button:'left',clickCount:1});
  }
  await evaluate('window.__containedReview.ready.then(()=>true)');
  const bBefore=await inspect('B');assert.equal(bBefore.document.source,'z_2');
  assert.deepEqual((await stats()).errors,[],'Review host failed to initialize');
  assert.ok(await evaluate('window.__containedReview.scene("A")'),'Review input must have current layout before keys are sent');
  await click('#focus-A');await type('x+y=1');const single=await checkpoint('typed initial single formula');
  assert.equal(single.document.source,'x+y=1','Initial trusted key input did not reach frame A');
  await key('Enter','Enter',13,8);const promoted=await checkpoint('M01 Shift+Enter creates one contained group');
  assert.ok(getGrid(promoted),'Shift+Enter did not create a contained grid');
  assert.equal(promoted.document.analysis.facts.grids,1);assert.equal(getGrid(promoted).rows.length,2);
  assert.equal(promoted.undoDepth,single.undoDepth+1);assert.equal((await stats()).enters.A,0);
  assert.equal(promoted.document.documentId,single.document.documentId);
  assert.equal(promoted.document.analysis.tree.id,single.document.analysis.tree.id);
  await click('#undo-A');const undone=await checkpoint('M08 one undo reverses whole promotion');
  assert.deepEqual(contentState(undone),contentState(single));
  await click('#redo-A');const redone=await checkpoint('M08 redo restores the same internal identities');
  assert.deepEqual(contentState(redone),contentState(promoted));
  await type('x-y=0');const two=await checkpoint('M01 two completed rows');
  assert.equal(getGrid(two).rows.length,2);assert.ok(two.document.analysis.exportable);
  await key('Enter','Enter',13,8);await type('z=2');const three=await checkpoint('M02 existing group gains a third row');
  assert.equal(three.document.analysis.facts.grids,1);assert.equal(getGrid(three).rows.length,3);
  assert.equal(await evaluate('document.querySelectorAll("#editor-A svg").length'),1);
  assert.equal((await stats()).enters.A,0);
  await click('#svg-A');const vector=await evaluate('window.__containedReview.exported("A")');
  assert.ok(vector?.startsWith('<svg'));assert.ok(!vector.includes('foreignObject')&&!vector.includes('<text'));
  const layout=await evaluate('window.__containedReview.scene("A")');
  for(const field of ['width','height'])assert.ok(Math.abs(Number(new RegExp('\\b'+field+'="([^"]+)"').exec(vector)?.[1])-layout[field])<1e-5);
  await click('#focus-A');const beforeEnter=await inspect('A');await key('Enter','Enter',13);
  assert.equal((await stats()).enters.A,1,'Ordinary Enter must remain a separate host action');
  assert.deepEqual(await inspect('A'),beforeEnter);await checkpoint('M11 ordinary Enter does not insert an environment row');

  await setSource('A','abc');
  const middleRoot=(await inspect('A')).document.analysis.tree.id;
  await pointerCaret('A',middleRoot,1);
  const middleBefore=await inspect('A');assert.equal(middleBefore.selection.focus.offset,1);
  const originalIds=middleBefore.document.analysis.tree.entries.map(e=>e.node.id);
  await key('Enter','Enter',13,8);const split=await checkpoint('M03 split at a root boundary preserves child identities');
  const cells=getGrid(split).rows.map(r=>r.cells[0]);
  assert.deepEqual(cells[0].entries.map(e=>e.node.id),originalIds.slice(0,1));
  assert.deepEqual(cells[1].entries.map(e=>e.node.id),originalIds.slice(1));
  assert.equal(split.selection.focus.nodeId,cells[1].id);assert.equal(split.selection.focus.offset,0);
  await key('Backspace','Backspace',8);const merged=await checkpoint('M10 row-start Backspace merges within the same frame');
  assert.equal(getGrid(merged).rows.length,1);assert.equal((await stats()).enters.A,1);

  await setSource('A',String.raw`\frac{x}{y}`);
  const fraction=await inspect('A');const numerator=fraction.document.analysis.tree.entries[0].node.numerator;
  await pointerCaret('A',numerator.id,1);const interior=await inspect('A');
  await key('Enter','Enter',13,8);assert.deepEqual(await inspect('A'),interior);
  await checkpoint('M04 unsupported interior break leaves the whole fraction unchanged');

  // Chromium-protocol composition, not physical Windows IME. Keep this claim
  // distinct from the separate native-IME checkpoint.
  await setSource('A','q');await click('#focus-A');const preeditBefore=await inspect('A');
  await send('Input.imeSetComposition',{text:'sudu',selectionStart:4,selectionEnd:4});
  const composing=await inspect('A');assert.ok(composing.composition,'Composition did not start');
  await key('Enter','Enter',13,8);const pending=await inspect('A');
  assert.equal(pending.document.source,preeditBefore.document.source);
  assert.equal(pending.document.revision,preeditBefore.document.revision);
  assert.equal(pending.document.analysis.facts.grids,0);
  assert.equal((await stats()).enters.A,1);
  await send('Input.imeSetComposition',{text:'',selectionStart:0,selectionEnd:0});
  await checkpoint('M07 pending protocol composition creates no extra formula row');

  await setSource('A',three.document.source);const reopened=await checkpoint('M09 full multiline source reopens as one formula');
  assert.equal(reopened.document.analysis.facts.grids,1);assert.equal(getGrid(reopened).rows.length,3);
  const naturalScene=await evaluate('window.__containedReview.scene("A")');
  const zoomCases=[];
  for(const dpr of [1,1.25,2]){
    await send('Emulation.setDeviceMetricsOverride',{width:1100,height:760,deviceScaleFactor:dpr,mobile:false});
    for(const [zoomIndex,zoom] of [.5,1,2].entries()){
      await click('#zoom-A');await key('Home','Home',36);
      for(let i=0;i<zoomIndex;i++)await key('ArrowDown','ArrowDown',40);
      await key('Enter','Enter',13);
      assert.equal(await evaluate('Number(document.getElementById("zoom-A").value)'),zoom);
      for(const rowIndex of [0,2]){
        const current=await inspect('A');const rowCell=getGrid(current).rows[rowIndex].cells[0];
        await pointerCaret('A',rowCell.id,1);
        const positioned=await inspect('A');
        assert.equal(positioned.selection.focus.nodeId,rowCell.id);assert.equal(positioned.selection.focus.offset,1);
        assert.equal(positioned.document.source,reopened.document.source);assert.equal(positioned.document.revision,reopened.document.revision);
        const currentScene=await evaluate('window.__containedReview.scene("A")');
        assert.equal(currentScene.width,naturalScene.width);assert.equal(currentScene.height,naturalScene.height);
        assert.deepEqual(currentScene.gridColumns,naturalScene.gridColumns);
        assert.deepEqual(await inspect('B'),bBefore,'Zoom/pointer input changed frame B');
        zoomCases.push({dpr,zoom,rowIndex,position:positioned.selection.focus});
      }
    }
  }
  await send('Emulation.clearDeviceMetricsOverride');
  await click('#zoom-A');await key('Home','Home',36);await key('ArrowDown','ArrowDown',40);await key('Enter','Enter',13);
  await checkpoint('M12 18 zoom/DPR/row pointer checks leave source and other formula unchanged');
  const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});
  writeFileSync(resolve(out,'two-independent-frames.png'),Buffer.from(screenshot.data,'base64'));
  const aEvents=await evaluate('window.__containedReview.events("A")');
  return {status:'PASS',browser:browserVersion.product,hostBuild:manifest.buildId,
    bootstrapSha256:sha(bootstrap),wasmSha256:manifest.assets.find(a=>a.id==='core').sha256,
    steps,zoomCases,frameBUnchanged:true,actualBrowserKeysAndPointer:true,events:aEvents.length,
    trustedEvents:aEvents.filter(e=>e.trusted).length,physicalWindowsIme:false,
    protocolComposition:true,notCovered:['Matrix/cases/aligned','Native Windows IME','P09 structured clipboard','Office or installed application'],
    screenshot:relative(root,resolve(out,'two-independent-frames.png')).replaceAll('\\','/')};
 });
} catch (error) {
 report={status:'FAILED',error:String(error),steps,hostBuild:manifest.buildId,bootstrapSha256:sha(bootstrap),physicalWindowsIme:false};
 writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));
 console.error(JSON.stringify({receipt:relative(root,resolve(out,'report.json')), ...report},null,2));
 throw error;
}
writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));
console.log(JSON.stringify({receipt:relative(root,resolve(out,'report.json')), ...report},null,2));
