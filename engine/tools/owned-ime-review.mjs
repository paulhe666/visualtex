// Actual Windows IME processing, driven by guarded ASCII SendInput keystrokes.
// No CDP imeSetComposition, Unicode injection or clipboard is used for text.
import assert from 'node:assert/strict';
import {createHash,randomUUID} from 'node:crypto';
import {readFileSync,mkdirSync,writeFileSync} from 'node:fs';
import {spawnSync} from 'node:child_process';
import {dirname,resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import {setTimeout as sleep} from 'node:timers/promises';
import {withOwnedBrowser} from './owned-browser-review.mjs';
import {createSandboxAssets} from './sandbox-assets.mjs';
// Always test the current host source. A historical target/svg-review build
// may be older than the core or input bridge, even when its file still exists.
const currentHostAssets=createSandboxAssets();
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const sha=b=>createHash('sha256').update(b).digest('hex');
const wasm=readFileSync(resolve(root,'web/pkg/visualtex_engine_wasm_bg.wasm'));
const math=readFileSync(process.env.VISUALTEX_TEST_MATH_FONT??resolve(root,'../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf'));
const text=readFileSync(process.env.VISUALTEX_TEST_TEXT_FONT??'C:/Windows/Fonts/NotoSansSC-VF.ttf');
const title=`VisualTeX-IME-${randomUUID()}`;
const assets=new Map([
 ['/core.js',{type:'text/javascript',bytes:readFileSync(resolve(root,'web/pkg/visualtex_engine_wasm.js'))}],
 ['/core.wasm',{type:'application/wasm',bytes:wasm}],['/math.otf',{type:'font/otf',bytes:math}],['/text.ttf',{type:'font/ttf',bytes:text}],
]);
for(const [path,asset] of currentHostAssets.assets){
 if(!path.startsWith('/host/')||!path.endsWith('.js'))continue;
 assets.set(path.slice('/host'.length),asset);
}
for(const name of ['editor-host','input-bridge','svg-view']){
 assert.ok(assets.get(`/${name}.js`),`Current compiled host module is missing: ${name}`);
}
const html=`<!doctype html><meta charset="utf-8"><title>${title}</title><style>body{margin:44px;font:18px sans-serif;color:#171717}#mount{margin:60px 24px}.vt-core-position{margin-top:32px;font-size:14px}pre{padding:16px;background:#eee}.vt-core-preedit{background:white;border-bottom:1px solid black;color:black}</style><h2>VisualTeX 系统输入法验收</h2><p>仅此窗口接收测试拼音按键；原 Word 文档不参与测试。</p><div id="mount"></div><pre id="source"></pre>`;
const out=resolve(root,'target/owned-ime-review');mkdirSync(out,{recursive:true});
const source=String.raw`\frac{x_1+\text{}}{y^2}`;
const steps=[];
const report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion,ownedProcessId})=>{
 let initialLayout;
 let shiftToggleUsed=false;
 const native=(action,extra=[])=>{
   const p=spawnSync('powershell.exe',['-NoProfile','-File',resolve(root,'tools/owned-ime-input.ps1'),
     '-OwnedProcessId',String(ownedProcessId),'-ExpectedTitle',title,'-Action',action,...extra],
     {encoding:'utf8',windowsHide:true,timeout:action==='activate'?35000:20000,maxBuffer:1024*1024});
   assert.equal(p.error,undefined);assert.equal(p.status,0,p.stderr||p.stdout);
   const result=JSON.parse(p.stdout.replace(/^\uFEFF/,''));steps.push({native:result});return result;
 };
 const inspect=()=>evaluate(`(()=>{const input=host.input.textarea;const r=input.getBoundingClientRect();const editor=host.inspect();const focus=editor.selection.kind==='range'?editor.selection.focus:null;const stop=focus?host.layoutSnapshot?.carets.find(c=>c.position.kind===focus.kind&&c.position.nodeId===focus.nodeId&&(focus.kind==='text'?c.position.grapheme===focus.grapheme:c.position.offset===focus.offset)):null;return {editor,events:host.input.events,errors:host.diagnostics,inputFocused:document.activeElement===input,proxy:{left:r.left,top:r.top,width:r.width,height:r.height,scrollLeft:input.scrollLeft,scrollTop:input.scrollTop,expected:stop?host.view.layoutToClient(stop.x,stop.top):null}};})()`);
 async function prepare(zoom){
   const point=await evaluate(`(()=>{host.replaceSource(${JSON.stringify(source)});host.setZoom(${zoom});
    const c=host.layoutSnapshot.carets.find(c=>c.position.kind==='text'&&c.position.grapheme===0);return host.view.layoutToClient(c.x,(c.top+c.bottom)/2);})()`);
   await send('Input.dispatchMouseEvent',{type:'mousePressed',...point,button:'left',clickCount:1});
   await send('Input.dispatchMouseEvent',{type:'mouseReleased',...point,button:'left',clickCount:1});
   await send('Page.bringToFront');native('activate');
   // Restore only the owned mathematical caret after a manual titlebar click.
   await send('Input.dispatchMouseEvent',{type:'mousePressed',...point,button:'left',clickCount:1});
   await send('Input.dispatchMouseEvent',{type:'mouseReleased',...point,button:'left',clickCount:1});
   const s=await inspect();assert.equal(s.inputFocused,true);assert.equal(s.editor.selection.focus.kind,'text');return s;
 }
 try{
   await evaluate(`(async()=>{
    const {default:init,EngineRuntime}=await import('/core.js');const {EngineEditorHost}=await import('/editor-host.js');
    const data=await Promise.all(['/core.wasm','/math.otf','/text.ttf'].map(p=>fetch(p).then(r=>r.arrayBuffer())));await init({module_or_path:data[0]});
    window.runtime=new EngineRuntime();const f=data.slice(1).map(b=>JSON.parse(runtime.registerFont(new Uint8Array(b),0)).result.font.id);
    window.enterCalls=0;window.host=new EngineEditorHost(runtime,document.getElementById('mount'),{documentId:'owned-windows-ime',initialSource:${JSON.stringify(source)},
     layout:{mathFont:f[0],textFont:f[1],style:'display',fontSizePx:28,paddingEm:0.125},
     onChange:s=>document.getElementById('source').textContent=s.document.source,onEnter:()=>window.enterCalls++});return true;
   })()`);
   await send('Page.bringToFront');
   const env=native('inspect');initialLayout=env.beforeLayout;
   if(process.argv.includes('--inspect-only'))return {status:'INSPECT_ONLY',browser:browserVersion.product,ownedProcessId,title,environment:env,actualWindowsIme:false,steps};
   const chinese=env.loadedLayouts.find(l=>(l.languageId&1023)===4);
   assert.ok(chinese,'An already-loaded Chinese input locale is required; no IME is installed by this test');
   if(chinese.handle!==env.beforeLayout)native('layout',['-KeyboardLayout',String(chinese.handle)]);
   let before=await prepare(1);
   native('type',['-Text','sudu']);await sleep(200);
   let composing=await inspect();
   if(!composing.editor.composition){
     steps.push({englishModeObserved:composing.editor.document.source});
     before=await prepare(1);native('key',['-Key','shift']);shiftToggleUsed=true;native('type',['-Text','sudu']);await sleep(200);composing=await inspect();
   }
   assert.ok(composing.editor.composition,'Native pinyin keys did not enter Windows IME composition');
   assert.equal(composing.editor.document.source,before.editor.document.source,'Preedit changed committed formula');
   assert.ok(composing.events.some(e=>e.type==='compositionstart'&&e.trusted));
   assert.ok(composing.events.some(e=>e.type==='compositionupdate'&&e.trusted));
   assert.equal(composing.proxy.scrollLeft,0,'Native composition must not scroll the candidate anchor away from the core caret');
   assert.ok(composing.proxy.width>32,'Native input proxy needs composition room');
   assert.ok(Math.abs(composing.proxy.left-composing.proxy.expected.x)<1);
   steps.push({label:'real native composition',snapshot:composing});
   native('capture',['-OutputPath',resolve(out,'native-candidate.png')]);
   native('key',['-Key','backspace']);await sleep(150);
   const erased=await inspect();assert.equal(erased.editor.document.source,before.editor.document.source);assert.ok(erased.editor.composition);
   native('key',['-Key','escape']);await sleep(150);
   const cancelled=await inspect();assert.equal(cancelled.editor.document.source,before.editor.document.source);assert.equal(cancelled.editor.composition,null);
   steps.push({label:'native Backspace and Escape',snapshot:cancelled});
   native('type',['-Text','sudu']);native('key',['-Key','space']);await sleep(150);
   const committed=await inspect();assert.equal(committed.editor.composition,null);
   assert.equal(committed.editor.document.source,String.raw`\frac{x_1+\text{速度}}{y^2}`);
   assert.equal(committed.editor.undoDepth,before.editor.undoDepth+1,'One IME commit must produce one undo record');
   assert.equal(committed.errors.length,0);steps.push({label:'native space commit',snapshot:committed});
   native('capture',['-OutputPath',resolve(out,'native-committed.png')]);
   const zoomCases=[];
   for(const zoom of [0.5,2]){
     const start=await prepare(zoom);native('type',['-Text','sudu']);await sleep(100);
     const candidate=await inspect();assert.ok(candidate.editor.composition);assert.equal(candidate.editor.document.source,start.editor.document.source);
     assert.equal(candidate.proxy.scrollLeft,0,'Zoomed composition range remains anchored');assert.ok(Math.abs(candidate.proxy.left-candidate.proxy.expected.x)<1);
     steps.push({label:'native zoomed candidate anchor',zoom,proxy:candidate.proxy});
     native('capture',['-OutputPath',resolve(out,`native-candidate-zoom-${zoom}.png`)]);
     native('key',['-Key','space']);await sleep(100);const end=await inspect();
     assert.equal(end.editor.document.source,String.raw`\frac{x_1+\text{速度}}{y^2}`);
     assert.equal(end.editor.undoDepth,start.editor.undoDepth+1);zoomCases.push({zoom,revision:end.editor.document.revision,oneCommit:true});
   }
   before=await prepare(1);native('type',['-Text','sudu']);await sleep(100);
   assert.ok((await inspect()).editor.composition);native('key',['-Key','enter']);await sleep(150);
   const entered=await inspect();assert.equal(entered.editor.composition,null);assert.equal(await evaluate('window.enterCalls'),0);
   assert.equal(entered.editor.document.source.includes('\n'),false);assert.equal(entered.editor.undoDepth,before.editor.undoDepth+1);
   steps.push({label:'native IME Enter did not split formula',snapshot:entered});
   const result={status:'PASS',browser:browserVersion.product,ownedProcessId,title,actualWindowsIme:true,
     hostBuildId:currentHostAssets.manifest.buildId,hostModuleHashes:Object.fromEntries(['editor-host','input-bridge','svg-view'].map(name=>[name,sha(assets.get(`/${name}.js`).bytes)])),
     keySource:'Win32 SendInput ASCII virtual keys through the installed OS input method',humanPhysicalKeyboard:false,
     cdpCompositionUsed:false,unicodeInjectionUsed:false,clipboardUsed:false,zoomCases,
     wasmSha256:sha(wasm),steps,notCovered:['WebView2 embedding','macOS IME','Office UI','production default switch']};
   return result;
 }catch(error){let snapshot;try{snapshot=await inspect();}catch{}
   const failure=JSON.stringify({error:String(error),steps,snapshot},null,2);
   writeFileSync(resolve(out,'failure.json'),failure);writeFileSync(resolve(out,`${title}-failure.json`),failure);throw error;
 }finally{
   if(shiftToggleUsed){try{const current=await inspect();if(current.editor.composition)native('key',['-Key','escape']);native('key',['-Key','shift']);}catch{}}
   if(initialLayout!==undefined){try{native('layout',['-KeyboardLayout',String(initialLayout)]);}catch{}}
   try{await evaluate('window.host?.dispose();window.runtime?.free();true');}catch{}
 }
},{visible:true});
writeFileSync(resolve(out,process.argv.includes('--inspect-only')?'environment.json':'report.json'),JSON.stringify(report,null,2));
console.log(JSON.stringify({...report,steps:`${steps.length} steps retained in report.json`},null,2));
