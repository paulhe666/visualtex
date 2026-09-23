// Component boundary tests with a fault-injectable InputPort. These are explicit
// synthetic DOM traces, not mathematical-core, physical IME or Office acceptance.
import assert from 'node:assert/strict';
import {readFileSync,mkdirSync,writeFileSync} from 'node:fs';
import {dirname,resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import {withOwnedBrowser} from './owned-browser-review.mjs';
import {createSandboxAssets} from './sandbox-assets.mjs';
const currentHostAssets=createSandboxAssets();
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const bridge=currentHostAssets.assets.get('/host/input-bridge.js');
assert.ok(bridge,'Current input bridge did not compile');
const assets=new Map([['/input-bridge.js',bridge]]);
const results=await withOwnedBrowser(assets,'<!doctype html><meta charset="utf-8"><title>VisualTeX input boundary tests</title><div id="mount"></div>',async({evaluate})=>
 evaluate(`(async()=>{
  const {BrowserInputBridge,FRAGMENT_MIME,LATEX_MIME}=await import('/input-bridge.js');
  const tests=[];
  function equal(a,b,message){if(JSON.stringify(a)!==JSON.stringify(b))throw Error(message+': '+JSON.stringify(a)+' != '+JSON.stringify(b));}
  function check(value,message){if(!value)throw Error(message);}
  function harness(overrides={}){
   const calls=[],errors=[];let next=1;
   const port={key:e=>{calls.push(['key',e.key]);return false;},insert:t=>{calls.push(['insert',t]);return true;},
     remove:b=>{calls.push(['remove',b]);return true;},history:r=>{calls.push(['history',r]);return true;},
     beginComposition:()=>{const id=next++;calls.push(['begin',id]);return id;},
     updateComposition:(id,t)=>{calls.push(['update',id,t]);return true;},
     commitComposition:(id,t)=>{calls.push(['commit',id,t]);return true;},
     cancelComposition:id=>{calls.push(['cancel',id]);return true;},
     copy:()=>({schemaVersion:1,latex:'x',revision:3,selectionEpoch:4}),cut:()=>true,
     paste:(text,latex,structure)=>{calls.push(['paste',text,latex,structure]);return true;},focus:()=>{},error:e=>errors.push(String(e)),...overrides};
   const bridge=new BrowserInputBridge(document.getElementById('mount'),port),input=bridge.textarea;
   const comp=(type,data='')=>input.dispatchEvent(new CompositionEvent(type,{data,bubbles:true}));
   const edit=(type,inputType,data='',cancelable=true,isComposing=false)=>{const e=new InputEvent(type,{inputType,data,cancelable,isComposing,bubbles:true});
    // This is an explicitly synthetic cross-host trace: Chromium normalizes
    // unrecognized inputType constructor values to empty strings. Preserve the
    // requested boundary field rather than accidentally testing another event.
    if(e.inputType!==inputType)Object.defineProperty(e,'inputType',{value:inputType});
    input.dispatchEvent(e);return e;};
   return {bridge,input,port,calls,errors,comp,edit,dispose:()=>bridge.dispose()};
  }
  async function test(name,run){try{await run();tests.push({name,status:'PASS'});}catch(error){tests.push({name,status:'FAILED',error:String(error)});}}
  await test('cancelable normal insertion is forwarded once',()=>{const h=harness();try{
   const e=h.edit('beforeinput','insertText','x');check(e.defaultPrevented,'DOM default was not cancelled');equal(h.calls,[['insert','x']],'one insertion');equal(h.input.value,'','mirror cleared');
  }finally{h.dispose();}});
  await test('noncancelable insertion is owned by input not beforeinput',()=>{const h=harness();try{
   h.edit('beforeinput','insertText','x',false);equal(h.calls,[],'too-early insertion');h.input.value='x';h.edit('input','insertText','x',false);equal(h.calls,[['insert','x']],'one native insertion');
  }finally{h.dispose();}});
  await test('terminal composition input does not duplicate committed text',()=>{const h=harness();try{
   h.comp('compositionstart');h.comp('compositionupdate','速度');h.input.value='速度';h.comp('compositionend','速度');
   h.edit('beforeinput','insertFromComposition','速度');h.edit('input','insertText','速度',false);
   equal(h.calls.filter(c=>c[0]==='commit'),[['commit',1,'速度']],'commit count');equal(h.calls.filter(c=>c[0]==='insert'),[],'terminal duplicate');
  }finally{h.dispose();}});
  await test('a later real key boundary permits identical text again',()=>{const h=harness();try{
   h.comp('compositionstart');h.input.value='速';h.comp('compositionend','速');
   h.input.dispatchEvent(new KeyboardEvent('keydown',{key:'x'}));h.edit('beforeinput','insertText','速');
   equal(h.calls.filter(c=>c[0]==='insert'),[['insert','速']],'same-character input incorrectly suppressed');
  }finally{h.dispose();}});
  await test('empty compositionend cancels rather than inserting or deleting',()=>{const h=harness();try{
   h.comp('compositionstart');h.comp('compositionupdate','temp');h.comp('compositionend','');
   equal(h.calls.filter(c=>c[0]==='cancel'),[['cancel',1]],'cancel');equal(h.calls.filter(c=>c[0]==='insert'||c[0]==='remove'||c[0]==='commit'),[],'cancel changed text');
  }finally{h.dispose();}});
  for(const failure of ['false','throw'])await test('failed composition '+failure+' keeps a retryable owner',()=>{
   let reject=true;const h=harness({commitComposition:(id,t)=>{h.calls.push(['commit',id,t]);if(reject){if(failure==='throw')throw Error('injected commit failure');return false;}return true;}});
   try{
    h.comp('compositionstart');h.comp('compositionupdate','速度');h.input.value='速度';h.comp('compositionend','速度');
    check(h.bridge.composing,'bridge forgot an uncommitted composition');check(h.input.value==='速度','uncommitted text was erased');
    reject=false;check(h.bridge.retryPendingInput(),'explicit retry did not succeed');check(!h.bridge.composing,'successful retry did not finish');
    equal(h.calls.filter(c=>c[0]==='commit'),[['commit',1,'速度'],['commit',1,'速度']],'retry must use same composition owner');
   }finally{h.dispose();}
  });
  await test('failed plain insertion leaves a recoverable input mirror',()=>{const h=harness({insert:()=>false});try{
   h.input.value='user input';h.edit('input','insertText','user input',false);equal(h.input.value,'user input','failed input was erased');
  }finally{h.dispose();}});
  await test('pending text retries on Enter and is not mixed with another key',()=>{
   let reject=true,accepted=0;const h=harness({insert:t=>{h.calls.push(['insert',t]);if(reject)return false;accepted++;return true;}});
   try{
    h.edit('beforeinput','insertText','x');check(h.bridge.pendingInput?.text==='x','recovery payload missing');
    const k=new KeyboardEvent('keydown',{key:'y',cancelable:true});h.input.dispatchEvent(k);check(k.defaultPrevented,'pending text allowed unrelated key');
    equal(h.calls,[['insert','x']],'new key forwarded while recovery pending');
    reject=false;h.input.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',cancelable:true}));
    equal(accepted,1,'retry committed twice');equal(h.bridge.pendingInput,null,'retry not cleared');equal(h.input.value,'','accepted mirror not cleared');
   }finally{h.dispose();}
  });
  await test('pending text may be explicitly discarded without deleting core content',()=>{const h=harness({insert:()=>false});try{
    h.edit('beforeinput','insertText','pending');h.input.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',cancelable:true}));
    equal(h.bridge.pendingInput,null,'pending text not discarded');equal(h.calls.filter(c=>c[0]==='remove'),[],'discard deleted committed content');
  }finally{h.dispose();}});
  await test('another native composition cannot overwrite a failed commit',()=>{const h=harness({commitComposition:()=>false});try{
    h.comp('compositionstart');h.comp('compositionupdate','old');h.comp('compositionend','old');
    h.comp('compositionstart');h.comp('compositionupdate','new');h.input.value='new';h.comp('compositionend','new');
    equal(h.bridge.pendingInput.text,'old','failed original preedit lost');equal(h.input.value,'old','new native mirror displaced original recovery text');
    check(h.bridge.discardPendingInput(),'explicit cancel failed');equal(h.calls.filter(c=>c[0]==='cancel'),[['cancel',1]],'original owner not cancelled exactly once');
  }finally{h.dispose();}});
  await test('corrupt rich clipboard falls back to plain text without executing HTML',()=>{const h=harness();try{
   const data=new DataTransfer();data.setData(FRAGMENT_MIME,'{invalid');data.setData('text/plain','x+y');data.setData('text/html','<script>window.clipboardExecuted=true</script>');
   h.input.dispatchEvent(new ClipboardEvent('paste',{clipboardData:data,cancelable:true}));
   equal(h.calls,[['paste','x+y',null,null]],'plain fallback');check(!window.clipboardExecuted,'HTML executed');
  }finally{h.dispose();}});
  await test('copy does not expose session identity or live node identity',()=>{const h=harness();try{
   const data=new DataTransfer();h.input.dispatchEvent(new ClipboardEvent('copy',{clipboardData:data,cancelable:true}));
   equal(JSON.parse(data.getData(FRAGMENT_MIME)),{schemaVersion:1,latex:'x'},'unexpected copy payload');equal(data.getData(LATEX_MIME),'x','LaTeX fallback missing');equal(data.getData('text/plain'),'x','plain fallback missing');
  }finally{h.dispose();}});
  await test('versioned structural flavor is forwarded as data while HTML stays inert',()=>{const structure={schema:'visualtex-engine-fragment',version:1,sourceSha256:'0'.repeat(64),shape:{kind:'row',entries:[]},assets:[]};const h=harness();try{
   const data=new DataTransfer();data.setData('text/plain','fallback');data.setData(LATEX_MIME,'latex-fallback');data.setData(FRAGMENT_MIME,JSON.stringify({schemaVersion:2,latex:'x',structure}));data.setData('text/html','<img src=x onerror="window.clipboardExecuted=true">');
   h.input.dispatchEvent(new ClipboardEvent('paste',{clipboardData:data,cancelable:true}));equal(h.calls,[['paste','fallback','x',structure]],'structured paste forwarding');check(!window.clipboardExecuted,'HTML executed');
  }finally{h.dispose();}});
  await test('legacy v1 custom flavor and application/x-latex remain readable',()=>{const h=harness();try{
   let data=new DataTransfer();data.setData('text/plain','plain');data.setData(LATEX_MIME,'\\\\alpha');h.input.dispatchEvent(new ClipboardEvent('paste',{clipboardData:data,cancelable:true}));equal(h.calls,[['paste','plain','\\\\alpha',null]],'LaTeX MIME fallback');
   h.calls.length=0;data=new DataTransfer();data.setData('text/plain','plain');data.setData(FRAGMENT_MIME,JSON.stringify({schemaVersion:1,latex:'\\\\beta'}));h.input.dispatchEvent(new ClipboardEvent('paste',{clipboardData:data,cancelable:true}));equal(h.calls,[['paste','plain','\\\\beta',null]],'legacy custom fallback');
  }finally{h.dispose();}});
  await test('oversized clipboard is rejected before port dispatch',()=>{const h=harness();try{
   const data=new DataTransfer();data.setData('text/plain','x'.repeat(1048577));h.input.dispatchEvent(new ClipboardEvent('paste',{clipboardData:data,cancelable:true}));
   equal(h.calls,[],'oversize paste reached port');check(h.errors.length===1,'missing budget diagnostic');
  }finally{h.dispose();}});
  await test('disposal disconnects input listeners and DOM',()=>{const h=harness();h.dispose();h.edit('beforeinput','insertText','x');equal(h.calls,[],'disposed listener ran');check(!h.input.isConnected,'input retained');});
  return {syntheticDomEvents:true,scope:'Input bridge component boundary only; fault-injectable InputPort',tests,passed:tests.filter(t=>t.status==='PASS').length,failed:tests.filter(t=>t.status==='FAILED').length};
 })()`));
const out=resolve(root,'target/input-bridge-contract-review');mkdirSync(out,{recursive:true});writeFileSync(resolve(out,'report.json'),JSON.stringify(results,null,2));
console.log(JSON.stringify(results,null,2));assert.equal(results.failed,0,'Input bridge contract review has preserved failing cases');
