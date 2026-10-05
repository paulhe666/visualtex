const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const assert = require('node:assert/strict');
const ts = require('typescript');
const source = fs.readFileSync(path.join(__dirname, '../src/office/dialog/useOfficeSession.ts'), 'utf8');
const compiled = ts.transpileModule(source, { compilerOptions: {
  target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS,
}}).outputText;
function deferred() { let resolve, reject; const promise = new Promise((a,b) => { resolve=a; reject=b; }); return { promise, resolve, reject }; }
async function flush() { for (let i=0; i<12; i++) await Promise.resolve(); }
function fixture() {
  const slots=[], effects=[], events=new EventTarget(), gets=[], patches=[];
  let cursor=0;
  const changed=(a,b)=>!a||!b||a.length!==b.length||a.some((v,i)=>v!==b[i]);
  const react={
    useState(initial) { const i=cursor++; if (!(i in slots)) slots[i]=typeof initial==='function'?initial():initial;
      return [slots[i], value=>{ slots[i]=typeof value==='function'?value(slots[i]):value; }]; },
    useRef(initial) { const i=cursor++; return slots[i]??(slots[i]={current:initial}); },
    useCallback(fn,deps) { const i=cursor++; if(changed(slots[i]?.deps,deps)) slots[i]={fn,deps}; return slots[i].fn; },
    useEffect(fn,deps) { const i=cursor++; if(changed(slots[i]?.deps,deps)) effects.push(()=>{ slots[i]?.cleanup?.(); slots[i]={deps,cleanup:fn()}; }); },
  };
  const context={ exports:{}, console, URLSearchParams,
    window:{ location:{search:'?sessionId=A',pathname:'/dialog'},
      addEventListener:events.addEventListener.bind(events), removeEventListener:events.removeEventListener.bind(events),
      setTimeout:()=>1,clearTimeout:()=>{} },
    require(id) { if(id==='react') return react;
      if(id.includes('readErrorMessage')) return {readErrorMessage:String};
      return { getOfficeSession(id) { const d=deferred(); gets.push({id,...d}); return d.promise; },
        updateOfficeSession(id,update) { const d=deferred(); patches.push({id,update,...d}); return d.promise; } };
    },
  };
  vm.runInNewContext(compiled,context);
  const render=()=>{cursor=0;const result=context.exports.useOfficeSession();while(effects.length)effects.shift()();return result;};
  const switchTo=id=>{events.dispatchEvent(new CustomEvent('visualtex-office-session',{detail:{sessionId:id}}));return render();};
  const loaded=async (id,marker=id)=>{gets.at(-1).resolve({id,title:marker,lines:[{latex:marker}]});await flush();return render();};
  return {render,switchTo,loaded,gets,patches,unmount(){for(const s of slots)s?.cleanup?.();}};
}
(async()=>{
  let checks=0;
  for(const target of ['B','','A']) {
    const f=fixture(); f.render(); let current=await f.loaded('A');
    const oldActivation=current.activation;
    const oldIsActive=current.isActive;
    const saving=current.save({title:'old'}); await flush();
    current=f.switchTo(target);
    assert.equal(oldIsActive(), false, 'old async UI continuations lose their owner synchronously');
    assert.equal(current.isActive(), true);
    if(target) current=await f.loaded(target,'new activation');
    f.patches[0].resolve({id:'A',title:'old'}); await saving; current=f.render();
    assert.equal(current.sessionId,target);
    assert.equal(current.session?.title,target?'new activation':undefined);
    if(current.activation) assert.notEqual(current.activation,oldActivation);
    f.unmount(); checks++;
  }
  const f=fixture(); f.render(); const firstGet=f.gets[0]; f.switchTo('B');
  await f.loaded('B'); firstGet.resolve({id:'A'}); await flush(); assert.equal(f.render().session.id,'B'); checks++;
  let current=f.render(); const first=current.save({title:'first'}); const rejected=first.catch(()=>{});
  const second=current.save({title:'second'}); await flush();
  assert.equal(f.patches.length,1); f.patches[0].reject(new Error('first failed')); await rejected; await flush();
  assert.equal(f.patches.length,2); assert.equal(f.patches[1].id,'B');
  f.patches[1].resolve({id:'B',title:'second'}); await second; assert.equal(f.render().session.title,'second'); checks++;
  current=f.render(); const pending=current.save({title:'after unmount'}); await flush(); f.unmount();
  f.patches[2].resolve({id:'B',title:'after unmount'}); await pending;
  assert.equal(f.render().session.title,'second'); checks++;
  console.log(`${checks} Office session ownership cases passed (real hook, mocked React lifecycle/HTTP; no host acceptance)`);
})().catch(error=>{console.error(error);process.exitCode=1;});
