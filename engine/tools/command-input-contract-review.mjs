// P09 pure transient command-state contract. The browser host is tested separately.
import assert from 'node:assert/strict';
import {mkdirSync,readFileSync,writeFileSync} from 'node:fs';
import {resolve,relative,dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
import {createSandboxAssets} from './sandbox-assets.mjs';
const {root}=createSandboxAssets();
const stamp=new Date().toISOString().replaceAll(':','-').replaceAll('.','-');
const out=resolve(root,'target/command-input-contract-review',stamp);mkdirSync(out,{recursive:true});
const compiled=readFileSync(resolve(root,'target/sandbox/host/command-input.js'));
const moduleUrl=`data:text/javascript;base64,${compiled.toString('base64')}`;
const {CommandInputState,CORE_COMMAND_DESCRIPTORS}=await import(moduleUrl);
const p={kind:'row',nodeId:1,offset:1,affinity:'after'};
const target=(revision=0,selectionEpoch=0)=>({documentId:'command-contract',revision,selectionEpoch,selection:{kind:'range',anchor:p,focus:p}});
const checks=[];
function check(name,fn){fn();checks.push(name);}
check('draft and candidate movement are transient',()=>{
 const state=new CommandInputState(CORE_COMMAND_DESCRIPTORS);state.begin(target());
 for(const ch of 'partial')assert.equal(state.append(ch,target()),true);
 const snapshot=state.snapshot();assert.equal(snapshot.active,true);assert.equal(snapshot.query,'partial');assert.equal(snapshot.candidates[0].id,'symbol-partial');
 assert.equal(snapshot.candidates[0].action.kind,'insertSymbol');assert.equal(snapshot.candidates[0].action.command,'\\partial');
 assert.equal(state.move(1,target()),true);assert.equal(state.snapshot().candidates[state.snapshot().selectedIndex].id,'template-partial-derivative');
});
check('exact partial can never silently become derivative template',()=>{
 const state=new CommandInputState(CORE_COMMAND_DESCRIPTORS);state.begin(target());for(const ch of 'partial')state.append(ch,target());
 const confirmation=state.confirm(target());assert.equal(confirmation.id,'symbol-partial');assert.deepEqual(confirmation.action,{kind:'insertSymbol',command:'\\partial'});assert.equal(state.snapshot().active,false);
});
check('explicit derivative template is a separate action',()=>{
 const state=new CommandInputState(CORE_COMMAND_DESCRIPTORS);state.begin(target());for(const ch of 'partialderivative')state.append(ch,target());
 const confirmation=state.confirm(target());assert.equal(confirmation.id,'template-partial-derivative');assert.equal(confirmation.action.kind,'pasteLatex');assert.match(confirmation.action.latex,/^\\frac/u);
});
check('backspace edits draft then cancels empty slash',()=>{
 const state=new CommandInputState(CORE_COMMAND_DESCRIPTORS);state.begin(target());state.append('sqrt',target());
 assert.equal(state.backspace(target()),'updated');assert.equal(state.snapshot().query,'sqr');
 state.backspace(target());state.backspace(target());state.backspace(target());assert.equal(state.snapshot().query,'');assert.equal(state.backspace(target()),'cancelled');assert.equal(state.active,false);
});
check('stale revision selection epoch or selection refuses confirmation',()=>{
 for(const stale of [target(1,0),target(0,1),{...target(),selection:{kind:'range',anchor:{...p,offset:0},focus:{...p,offset:0}}}]){
  const state=new CommandInputState(CORE_COMMAND_DESCRIPTORS);state.begin(target());state.append('partial',target());assert.equal(state.confirm(stale),null);assert.equal(state.active,false);
 }
});
check('usage may reorder prefix matches but not mutate descriptors',()=>{
 const descriptors=[
  {id:'alpha',trigger:'alpha',aliases:[],priority:0,action:{kind:'insertSymbol',command:'\\alpha'}},
  {id:'approx',trigger:'approx',aliases:[],priority:0,action:{kind:'insertSymbol',command:'\\approx'}},
 ];
 const state=new CommandInputState(descriptors,{approx:{useCount:100,pinned:true,acceptedPrefixes:{a:9}}});state.begin(target());state.append('a',target());
 assert.equal(state.snapshot().candidates[0].id,'approx');assert.equal(descriptors[0].trigger,'alpha');
});
check('duplicate ids and unsafe triggers fail closed',()=>{
 assert.throws(()=>new CommandInputState([{id:'x',trigger:'x',action:{kind:'insertText',text:'x',mode:'math'}},{id:'x',trigger:'y',action:{kind:'insertText',text:'y',mode:'math'}}]));
 assert.throws(()=>new CommandInputState([{id:'bad',trigger:'\\frac{}',action:{kind:'insertText',text:'x',mode:'math'}}]));
});
const report={schema:'visualtex-p09-command-state-v1',status:'PASS',checks,descriptorCount:CORE_COMMAND_DESCRIPTORS.length,partialExactAction:'insertSymbol',templateAction:'pasteLatex',documentWrites:0};
writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({receipt:relative(root,resolve(out,'report.json')).replaceAll('\\','/'),...report},null,2));
