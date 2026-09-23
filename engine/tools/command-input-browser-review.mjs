// P09 real Edge keyboard acceptance for the self-owned transient command input.
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {mkdirSync,writeFileSync} from 'node:fs';
import {resolve,relative} from 'node:path';
import {createSandboxAssets} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';

const {assets,manifest,root}=createSandboxAssets();
const stamp=new Date().toISOString().replaceAll(':','-').replaceAll('.','-');
const out=resolve(root,'target/command-input-browser-review',stamp);mkdirSync(out,{recursive:true});
const sha=b=>createHash('sha256').update(b).digest('hex');
const bootstrap=`import init,{EngineRuntime} from '/assets/core.js';
import {EngineEditorHost} from '/host/editor-host.js';
import {VerifiedAssetLoader} from '/host/resources.js';
import {CORE_COMMAND_DESCRIPTORS} from '/host/command-input.js';
let runtime,loader;const hosts={},errors=[],accepted=[],states=[],enters={A:0,B:0};
window.ready=(async()=>{const manifest=await(await fetch('/manifest.json',{cache:'no-store'})).json();loader=new VerifiedAssetLoader(manifest,{baseUrl:new URL('/assets/',location.href).href,expectedBuildId:manifest.buildId,expectedEngineProtocol:1});const core=await loader.load('core');await init({module_or_path:core.bytes});runtime=new EngineRuntime();const fonts={};
for(const asset of manifest.assets.filter(a=>a.kind==='font')){const verified=await loader.load(asset.id);const r=JSON.parse(runtime.registerFont(verified.bytes,0));if(r.status!=='ok')throw Error(JSON.stringify(r));fonts[asset.id]=r.result.font.id;}
for(const id of ['A','B']){const host=new EngineEditorHost(runtime,document.getElementById('editor-'+id),{documentId:'command-'+id,initialSource:id==='A'?'x':'z_2',commands:CORE_COMMAND_DESCRIPTORS,layout:{mathFont:fonts.math,textFont:fonts.text,mathFallbacks:[],style:'display',fontSizePx:26,paddingEm:.125},onError:e=>errors.push({id,message:String(e)}),onEnter:()=>enters[id]++,onCommandState:s=>states.push({id,state:s}),onCommandAccepted:(commandId,query)=>accepted.push({id,commandId,query})});hosts[id]=host;}hosts.A.focus();return true;})();
window.review={inspect:id=>hosts[id].inspect(),command:id=>hosts[id].inspectCommandInput(),events:id=>hosts[id].input.events,stats:()=>({active:runtime.activeSessions,errors:errors.slice(),accepted:accepted.slice(),states:states.slice(),enters:{...enters},popupA:{hidden:hosts.A.commandPopup.hidden,text:hosts.A.commandPopup.textContent,selected:hosts.A.commandPopup.querySelector('[aria-selected="true"]')?.dataset.commandId??null}}),focus:id=>hosts[id].focus(),replace:(id,s)=>hosts[id].replaceSource(s),dispose:()=>{for(const h of Object.values(hosts))h.dispose();runtime.free();loader.dispose();}};
window.addEventListener('pagehide',()=>{try{for(const h of Object.values(hosts))h.dispose();}finally{runtime?.free();loader?.dispose();}},{once:true});`;
assets.set('/command-review.js',{type:'text/javascript; charset=utf-8',bytes:Buffer.from(bootstrap)});
const html=`<!doctype html><html><head><meta charset="utf-8"><title>P09 command input</title><style>body{font:16px Arial;margin:22px}.editor{position:relative;border:1px solid #aaa;min-height:100px;margin:12px;padding:12px}.vt-core-command-popup{background:white;border:1px solid #888;padding:6px;min-width:180px}.vt-core-command-candidate[aria-selected="true"]{font-weight:700}</style></head><body><h1>Self-owned command input</h1><div id="editor-A" class="editor"></div><div id="editor-B" class="editor"></div><script type="module" src="/command-review.js"></script></body></html>`;
const steps=[];let report;
try{
 report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
  await evaluate('window.ready.then(()=>true)');
  const inspect=id=>evaluate(`review.inspect(${JSON.stringify(id)})`),command=id=>evaluate(`review.command(${JSON.stringify(id)})`),stats=()=>evaluate('review.stats()');
  async function key(key,code,vk,modifiers=0,text){await send('Input.dispatchKeyEvent',{type:'keyDown',key,code,windowsVirtualKeyCode:vk,nativeVirtualKeyCode:vk,modifiers,...(text!==undefined?{text}:{})});await send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:vk,nativeVirtualKeyCode:vk,modifiers});}
  async function typeCommand(name){await key('\\','Backslash',220,0,'\\');for(const ch of name){const upper=ch.toUpperCase();await key(ch,`Key${upper}`,upper.charCodeAt(0),0,ch);}}
  const before=await inspect('A'),bBefore=await inspect('B');assert.equal(before.document.source,'x');assert.equal((await stats()).active,2);
  await typeCommand('partial');let draft=await inspect('A'),draftState=await command('A'),runtimeStats=await stats();
  assert.deepEqual(draft,before,'Command draft mutated editor snapshot');assert.equal(draftState.active,true);assert.equal(draftState.query,'partial');assert.equal(draftState.candidates[0].id,'symbol-partial');assert.equal(runtimeStats.popupA.hidden,false);assert.equal(runtimeStats.popupA.selected,'symbol-partial');assert.deepEqual(await inspect('B'),bBefore);
  steps.push({label:'trusted \\partial draft changes no source/revision/history',command:draftState,popup:runtimeStats.popupA});
  await key(' ','Space',32,0,' ');let partial=await inspect('A');runtimeStats=await stats();
  assert.equal(partial.document.revision,before.document.revision+1);assert.equal(partial.undoDepth,before.undoDepth+1);assert.equal(partial.document.analysis.facts.fractions,0);assert.equal(partial.document.analysis.tree.entries.at(-1).node.command,'\\partial');assert.equal((await command('A')).active,false);assert.equal(runtimeStats.popupA.hidden,true);assert.deepEqual(runtimeStats.accepted.at(-1),{id:'A',commandId:'symbol-partial',query:'partial'});assert.deepEqual(await inspect('B'),bBefore);
  steps.push({label:'Space confirms exact partial as one symbol transaction',source:partial.document.source,revision:partial.document.revision,undoDepth:partial.undoDepth});
  await key('z','KeyZ',90,2);let restored=await inspect('A');assert.equal(restored.document.source,before.document.source);assert.deepEqual(restored.document.analysis.tree,before.document.analysis.tree);assert.deepEqual(restored.selection,before.selection);

  await typeCommand('partialderivative');const templateDraft=await inspect('A');assert.deepEqual(templateDraft,restored);assert.equal((await command('A')).candidates[0].id,'template-partial-derivative');const entersBefore=(await stats()).enters.A;
  await key('Enter','Enter',13);let templated=await inspect('A');assert.equal((await stats()).enters.A,entersBefore);assert.equal(templated.document.analysis.facts.fractions,1);assert.equal(templated.document.revision,restored.document.revision+1);assert.equal(templated.undoDepth,restored.undoDepth+1);assert.match(templated.document.source,/\\frac/u);assert.deepEqual(await inspect('B'),bBefore);
  steps.push({label:'Enter confirms explicit derivative template without outer Enter',source:templated.document.source,undoDepth:templated.undoDepth});
  await key('z','KeyZ',90,2);restored=await inspect('A');assert.equal(restored.document.source,before.document.source);assert.deepEqual(restored.document.analysis.tree,before.document.analysis.tree);assert.deepEqual(restored.selection,before.selection);

  const escapeBaseline=restored;await typeCommand('sqrt');assert.equal((await command('A')).active,true);await key('Escape','Escape',27);assert.deepEqual(await inspect('A'),escapeBaseline);assert.equal((await command('A')).active,false);assert.equal((await stats()).popupA.hidden,true);
  steps.push({label:'Escape cancels command draft without history'});

  const backspaceBaseline=await inspect('A');await typeCommand('frac');await key('Backspace','Backspace',8);assert.equal((await command('A')).query,'fra');assert.deepEqual(await inspect('A'),backspaceBaseline);for(let i=0;i<3;i++)await key('Backspace','Backspace',8);assert.equal((await command('A')).query,'');await key('Backspace','Backspace',8);assert.equal((await command('A')).active,false);assert.deepEqual(await inspect('A'),backspaceBaseline);
  steps.push({label:'Backspace edits and finally cancels transient slash state'});

  await typeCommand('partial');let collision=await command('A');assert.equal(collision.candidates[0].id,'symbol-partial');assert.ok(collision.candidates.some(c=>c.id==='template-partial-derivative'));await key('ArrowDown','ArrowDown',40);collision=await command('A');assert.equal(collision.candidates[collision.selectedIndex].id,'template-partial-derivative');await key('Tab','Tab',9);templated=await inspect('A');assert.equal(templated.document.analysis.facts.fractions,1);assert.deepEqual(await inspect('B'),bBefore);await key('z','KeyZ',90,2);restored=await inspect('A');assert.equal(restored.document.source,'x');
  steps.push({label:'Arrow selection may explicitly choose template; Tab confirms one action'});

  await typeCommand('partial');assert.equal((await command('A')).active,true);await evaluate(`review.replace('A','y');true`);const replaced=await inspect('A');assert.equal(replaced.document.source,'y');assert.equal((await command('A')).active,false);assert.equal((await stats()).popupA.hidden,true);assert.deepEqual(await inspect('B'),bBefore);
  steps.push({label:'source owner replacement cancels stale command target before it can confirm'});

  await evaluate(`review.replace('A','x');review.focus('A');true`);await typeCommand('par');const imeBefore=await inspect('A');assert.equal((await command('A')).active,true);await send('Input.imeSetComposition',{text:'sudu',selectionStart:4,selectionEnd:4});const composing=await inspect('A');assert.equal((await command('A')).active,false);assert.ok(composing.composition);assert.equal(composing.document.source,imeBefore.document.source);await send('Input.imeSetComposition',{text:'',selectionStart:0,selectionEnd:0});assert.equal((await inspect('A')).composition,null);assert.deepEqual(await inspect('B'),bBefore);
  steps.push({label:'browser composition has priority and cancels transient command draft'});

  const events=await evaluate('review.events("A")');assert.ok(events.some(e=>e.type==='keydown'&&e.key==='\\'&&e.trusted));assert.ok(events.some(e=>e.type==='keydown'&&e.key==='Enter'&&e.trusted));assert.equal((await stats()).errors.length,0);
  const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});writeFileSync(resolve(out,'command-input.png'),Buffer.from(screenshot.data,'base64'));
  const accepted=(await stats()).accepted;const disposed=await evaluate('review.dispose();({inputs:document.querySelectorAll(".vt-core-input").length,popups:document.querySelectorAll(".vt-core-command-popup").length})');assert.deepEqual(disposed,{inputs:0,popups:0});
  return {status:'PASS',browser:browserVersion.product,hostBuild:manifest.buildId,wasmSha256:manifest.assets.find(a=>a.id==='core').sha256,actualBrowserKeyboard:true,trustedEvents:events.filter(e=>e.trusted).length,frameBUnchanged:true,steps,accepted,screenshot:relative(root,resolve(out,'command-input.png')).replaceAll('\\','/')};
 });
}catch(error){report={status:'FAILED',error:String(error),steps};throw error;}
finally{if(report){writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({receipt:relative(root,resolve(out,'report.json')).replaceAll('\\','/'),status:report.status,error:report.error??null,steps:steps.length},null,2));}}
