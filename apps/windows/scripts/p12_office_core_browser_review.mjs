// P12 isolated production-Office-dialog browser acceptance. No Word/COM/Session files.
import assert from 'node:assert/strict';
import {readFileSync,readdirSync,statSync,mkdirSync,writeFileSync} from 'node:fs';
import {resolve,dirname,relative,extname,join} from 'node:path';
import {fileURLToPath} from 'node:url';
import {withOwnedBrowser} from '../../../engine/tools/owned-browser-review.mjs';

const here=dirname(fileURLToPath(import.meta.url));
const appRoot=resolve(here,'..');
const repoRoot=resolve(appRoot,'../..');
const dist=resolve(appRoot,'dist-office-windows-native');
const stamp=new Date().toISOString().replaceAll(':','-').replaceAll('.','-');
const out=resolve(repoRoot,'engine/target/p12-office-browser-review',stamp);mkdirSync(out,{recursive:true});
const mime=path=>({'.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.wasm':'application/wasm','.otf':'font/otf','.woff2':'font/woff2'}[extname(path)]??'application/octet-stream');
const assets=new Map();
function collect(directory,prefix=''){
  for(const name of readdirSync(directory)){
    const file=join(directory,name),rel=prefix?`${prefix}/${name}`:name;
    if(statSync(file).isDirectory())collect(file,rel);else if(rel!=='dialog/index.html')assets.set('/'+rel.replaceAll('\\','/'),{type:mime(file),bytes:readFileSync(file)});
  }
}
collect(dist);
const originalHtml=readFileSync(resolve(dist,'dialog/index.html'),'utf8');
const mock=String.raw`(()=>{
const realFetch=window.fetch.bind(window);const now=Date.now();
const base={id:'p12-session',mode:'edit',host:'word',formulaId:'p12-formula',sourceDocumentId:'owned-doc',sourceObjectId:'owned-object',title:'P12 isolated formula',lines:[{id:'line-1',latex:'x+1'}],activeLineId:'line-1',codeFormat:'raw',displayMode:'inline',objectMode:'nativeOle',numbered:false,mathTypeNumberPosition:'right',fontSizePt:10.5,exportWidth:0,exportHeight:0,exportResult:null,originalMetadata:null,dirty:false,status:'editing',autoCommitOnClose:true,explicitCancel:false,error:null,createdAt:now,updatedAt:now,expiresAt:now+600000};
const state=window.__p12Mock={session:structuredClone(base),patches:[],requests:[],closeCalls:0,commitPending:false,theme:'light',delayNextPngMs:0};
const originalToBlob=HTMLCanvasElement.prototype.toBlob;
HTMLCanvasElement.prototype.toBlob=function(callback,...args){const delay=state.delayNextPngMs;state.delayNextPngMs=0;return originalToBlob.call(this,value=>delay?setTimeout(()=>callback(value),delay):callback(value),...args);};
const json=(value,status=200)=>new Response(JSON.stringify(value),{status,headers:{'Content-Type':'application/json'}});
window.fetch=async(input,init={})=>{const url=new URL(typeof input==='string'?input:input.url,location.href);const method=(init.method||((typeof input==='object'&&input.method)||'GET')).toUpperCase();
 if(!url.pathname.startsWith('/api/v1/'))return realFetch(input,init);state.requests.push({method,path:url.pathname});
 if(url.pathname==='/api/v1/preferences'&&method==='GET')return json({powerpointDefaultFontSizePt:20,editorPreferences:{settings:{theme:state.theme,editorLayout:'standard',language:'cn',zoom:.6,formulaLetterFont:'katex',formulaChineseFont:'system'}}});
 if(url.pathname==='/api/v1/theme'&&method==='GET')return json({theme:state.theme,editorLayout:'standard'});
 if(url.pathname==='/api/v1/app/converter/next-batch'&&method==='GET')return json({sessionIds:[]});
 if(url.pathname==='/api/v1/sessions/p12-session'&&method==='GET'){if(state.commitPending){state.session={...state.session,status:'completed',updatedAt:Date.now()};state.commitPending=false;}return json(state.session);}
 if(url.pathname==='/api/v1/sessions/p12-session'&&method==='PATCH'){const body=JSON.parse(String(init.body||'{}'));state.patches.push(structuredClone(body));state.session={...state.session,...body,updatedAt:Date.now()};const response=structuredClone(state.session);if(body.status==='committing')state.commitPending=true;return json(response);}
 if(url.pathname==='/api/v1/app/sessions/p12-session/close'&&method==='POST'){state.closeCalls++;return new Response(null,{status:204});}
 return json({error:'P12 isolated mock has no route'},404);
};})();`;
assets.set('/p12-mock.js',{type:'text/javascript; charset=utf-8',bytes:Buffer.from(mock)});
const html=originalHtml.replace('<script type="module"','<script src="/p12-mock.js"></script>\n    <script type="module"');
const coreChunk=[...assets.keys()].find(p=>/\/assets\/CoreMathEditor-[^/]+\.js$/u.test(p));
const wasmAsset=[...assets.keys()].find(p=>/\/assets\/visualtex_engine_wasm_bg-[^/]+\.wasm$/u.test(p));
assert.ok(coreChunk,'Office production bundle is missing CoreMathEditor chunk');assert.ok(wasmAsset,'Office production bundle is missing core WASM');

async function waitFor(evaluate,expression,timeout=12000){const start=Date.now();let last;while(Date.now()-start<timeout){try{last=await evaluate(expression);if(last)return last;}catch(error){last=String(error);}await new Promise(r=>setTimeout(r,80));}throw new Error(`P12 browser wait timed out: ${expression}; last=${String(last)}`);}
let report;const checks=[];
try{
 report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion,origin,requests})=>{
  await send('Network.enable');await send('Network.setCacheDisabled',{cacheDisabled:true});
  const navigate=async query=>{const loaded=send('Page.navigate',{url:`${origin}/?${query}`});await loaded;await new Promise(r=>setTimeout(r,50));};
  // Hidden converter must remain legacy and must not even load the core dynamic chunk.
  const beforeCore=requests.get(coreChunk)??0,beforeWasm=requests.get(wasmAsset)??0;
  await navigate('runtime=vsto-convert&sessionId=p12-session');
  await waitFor(evaluate,'window.__p12Mock && window.__p12Mock.closeCalls>0');
  assert.equal(await evaluate('document.querySelectorAll(\'[data-engine="visualtex-core"]\').length'),0);
  assert.equal(requests.get(coreChunk)??0,beforeCore);assert.equal(requests.get(wasmAsset)??0,beforeWasm);
  checks.push('vsto-convert completes without mounting or loading the core editor/WASM');
  // Fresh interactive navigation resets the mock Session and is allowed to select core.
  await navigate('runtime=vsto-desktop&sessionId=p12-session');
  await waitFor(evaluate,'document.querySelector(\'[data-engine="visualtex-core"]\')!==null');
  await waitFor(evaluate,'document.querySelector(\'.vt-core-input\')!==null');
  assert.equal(await evaluate('document.querySelectorAll("math-field").length'),0);
  assert.ok((requests.get(coreChunk)??0)>beforeCore);assert.ok((requests.get(wasmAsset)??0)>beforeWasm);
  checks.push('interactive production Office dialog mounts the core editor and loads verified WASM/font resources');
  // Initial Session hydration is not a user edit and must not commit/redraw anything.
  await new Promise(r=>setTimeout(r,250));
  const initialPatches=await evaluate('window.__p12Mock.patches');
  assert.ok(initialPatches.every(p=>p.dirty!==true&&p.status!=="committing"&&(!p.lines||p.lines[0]?.latex==="x+1")));
  assert.ok(initialPatches.every(p=>!p.exportResult));
  checks.push('initial Session to core hydration remains clean and produces no export or commit transaction');
  // Delay the first PNG completion, then edit again. The old core revision must be rejected rather than overwriting the newer source/export.
  await evaluate('window.__p12Mock.delayNextPngMs=650;document.querySelector(".vt-core-input").focus();true');
  await send('Input.insertText',{text:'a'});
  await waitFor(evaluate,'window.__p12Mock.patches.some(p=>p.lines?.[0]?.latex?.includes("a"))');
  await send('Input.insertText',{text:'b'});
  await waitFor(evaluate,'window.__p12Mock.patches.some(p=>p.lines?.[0]?.latex?.includes("ab"))');
  await waitFor(evaluate,'window.__p12Mock.patches.some(p=>p.lines?.[0]?.latex?.includes("ab") && p.exportResult?.mathMl?.startsWith("<math") && p.exportResult?.svg && p.exportResult?.pngBase64)');
  await new Promise(r=>setTimeout(r,750));
  const staleAudit=await evaluate('window.__p12Mock.patches.map(p=>({latex:p.lines?.[0]?.latex||null,hasExport:Boolean(p.exportResult?.pngBase64)}))');
  assert.equal(staleAudit.some(p=>p.latex==="x+1a"&&p.hasExport),false);
  assert.equal(staleAudit.some(p=>p.latex==="x+1ab"&&p.hasExport),true);
  const settledExportAudit=await evaluate(`(()=>{const rows=window.__p12Mock.patches.filter(p=>p.lines?.[0]?.latex==="x+1ab");const first=rows.findIndex(p=>Boolean(p.exportResult?.pngBase64));return{first,after:rows.slice(first+1).map(p=>p.exportResult===null?"null":p.exportResult?.pngBase64?"export":"other")}})()`);
  assert.ok(settledExportAudit.first>=0);
  assert.equal(settledExportAudit.after.includes("null"),false,"settled core export must not be cleared by a Session-object identity rerender");
  checks.push('rapid edits retire stale asynchronous core exports and settled exports are never cleared by Session response identity changes');
  const edited=await evaluate('window.__p12Mock.patches.filter(p=>p.lines?.[0]?.latex?.includes("ab")).at(-1)');
  assert.match(edited.lines[0].latex,/ab/u);
  checks.push('real browser text input updates the retained core owner and immediately persists source to the existing Session PATCH route');
  const exported=await evaluate('window.__p12Mock.patches.filter(p=>p.lines?.[0]?.latex?.includes("ab") && p.exportResult?.pngBase64).at(-1)');
  assert.ok(exported.exportResult.width>0&&exported.exportResult.height>0);assert.ok(exported.exportResult.baseline>=0&&exported.exportResult.baseline<=exported.exportResult.height);
  checks.push('same interactive core revision produces MathML, vector SVG, PNG and natural baseline geometry before export PATCH');
  // Live appearance sync remains outside core ownership.
  await evaluate('window.__p12Mock.theme="dark";true');
  await waitFor(evaluate,'document.documentElement.dataset.theme==="dark" || document.documentElement.classList.contains("dark") || document.body.classList.contains("dark")',2500);
  checks.push('theme synchronization remains owned by the existing Office shell while the core editor stays mounted');
  // Product undo/redo buttons must route to core history, not a second legacy history owner.
  await waitFor(evaluate,'!document.querySelector("[data-office-undo-action]").disabled');
  await evaluate('document.querySelector("[data-office-undo-action]").click();true');
  await waitFor(evaluate,'window.__p12Mock.patches.some(p=>p.lines?.[0]?.latex==="x+1a")');
  await waitFor(evaluate,'!document.querySelector("[data-office-redo-action]").disabled');
  await evaluate('document.querySelector("[data-office-redo-action]").click();true');
  await waitFor(evaluate,'window.__p12Mock.patches.filter(p=>p.lines?.[0]?.latex?.includes("ab")).length>=2');
  checks.push('Office undo/redo controls use the core history owner and persist restored source');
  // Explicit Apply waits for a current core export, then uses the unchanged existing Session commit handshake.
  await evaluate('document.querySelector("[data-office-primary-action]").click();true');
  await waitFor(evaluate,'window.__p12Mock.patches.some(p=>p.status==="committing" && p.exportResult?.mathMl?.startsWith("<math"))');
  await waitFor(evaluate,'window.__p12Mock.closeCalls>0');
  const committing=await evaluate('window.__p12Mock.patches.filter(p=>p.status==="committing").at(-1)');
  assert.match(committing.lines[0].latex,/ab/u);assert.ok(committing.exportResult.svg&&committing.exportResult.pngBase64);
  checks.push('explicit Apply exports the current core revision and preserves the existing Session/host commit handshake');
  // Fresh interactive session: Cancel must retain the existing Session path and never synthesize a commit.
  await navigate('runtime=vsto-desktop&sessionId=p12-session');
  await waitFor(evaluate,'document.querySelector(\'[data-engine="visualtex-core"]\')!==null');
  await evaluate('document.querySelector(".vt-core-input").focus();true');await send('Input.insertText',{text:'c'});
  await waitFor(evaluate,'window.__p12Mock.patches.some(p=>p.lines?.[0]?.latex?.includes("c"))');
  await evaluate('document.querySelector("[data-office-cancel-action]").click();true');
  await waitFor(evaluate,'window.__p12Mock.patches.some(p=>p.status==="cancelled"&&p.explicitCancel===true)');
  const cancelAudit=await evaluate('({commits:window.__p12Mock.patches.filter(p=>p.status==="committing").length,cancel:window.__p12Mock.patches.filter(p=>p.status==="cancelled").at(-1),closeCalls:window.__p12Mock.closeCalls})');
  assert.equal(cancelAudit.commits,0);assert.equal(cancelAudit.cancel.explicitCancel,true);assert.ok(cancelAudit.closeCalls>0);
  checks.push('Cancel keeps the original Session cancellation/close mechanism and never creates a commit');
  const allPatches=await evaluate('window.__p12Mock.patches');
  return{status:'PASS',browser:browserVersion.product,coreChunk,wasmAsset,patches:allPatches.length,checks,wordAutomation:0,officeSessionFiles:0,notCovered:['real Word Ribbon/document mutation','native window close interception','IME candidate UI','window reuse with a real native companion']};
 });
}catch(error){report={status:'FAILED',error:String(error),checks};throw error;}finally{if(report){writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({receipt:relative(repoRoot,resolve(out,'report.json')).replaceAll('\\','/'),...report},null,2));}}
