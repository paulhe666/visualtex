// Real installed WebView input workflows. DOM/model access is observation only;
// every formula edit goes through keyboard, IME, mouse or the source editor.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
const {chromium}=await import(process.env.VISUALTEX_PLAYWRIGHT_MODULE||'playwright');
const [port='9227',output,filter='all']=process.argv.slice(2);
assert.ok(output);await fs.mkdir(output,{recursive:true});
const browser=await chromium.connectOverCDP(`http://127.0.0.1:${port}`);
const page=browser.contexts()[0].pages().find(p=>p.url()==='http://tauri.localhost/');
assert.ok(page,'Must test the installed Tauri WebView');
const cdp=await page.context().newCDPSession(page), errors=[], steps=[];
const savedStorage=await page.evaluate(()=>Object.fromEntries(Object.entries(localStorage)));
const savedFormulas=await page.locator('math-field').evaluateAll(fields=>fields.map(f=>f.value));
await fs.writeFile(path.join(output,'session-backup.json'),JSON.stringify({storage:savedStorage,formulas:savedFormulas},null,2));
page.on('pageerror',e=>errors.push({message:e.message,stack:e.stack}));
const settle=()=>page.waitForTimeout(100);
const compact=s=>s.replace(/\s+/g,'').replace(/([_^])\{([a-zA-Z0-9])\}/g,'$1$2');
const state=()=>page.evaluate(()=>({
  source:document.querySelector('.cm-content')?.innerText,
  sourceError:document.querySelector('.source-error-chip')?.textContent,
  nativeVisible:!!document.querySelector('#mathlive-suggestion-popover.is-visible'),
  outerPanels:document.querySelectorAll('.suggestion-popup,#visualtex-native-input-suggestion-popover').length,
  candidates:[...document.querySelectorAll('#mathlive-suggestion-popover li')].map(e=>({latex:e.dataset.command,text:e.textContent,current:e.classList.contains('ML__popover__current')})),
  fields:[...document.querySelectorAll('math-field')].map(f=>({value:f.value,mode:f.mode,readOnly:f.readOnly,focus:f.hasFocus(),position:f.position,selection:f.selection,
    atoms:f._mathfield.model.root.children.map(a=>({offset:f._mathfield.model.offsetOf(a),type:a.type,value:a.value,branch:a.parentBranch,parent:a.parent?.type,parentCommand:a.parent?.command,command:a.command,selected:a.isSelected})),
    phantomCarets:f.shadowRoot.querySelectorAll('.visualtex-structural-placeholder-caret').length,
  }))
}));
const check=async(name,verify)=>{await settle();const s=await state();steps.push({name,...s});await fs.writeFile(path.join(output,'steps.json'),JSON.stringify(steps,null,2));verify(s);assert.deepEqual(errors,[],name+': runtime exception');assert.equal(s.outerPanels,0,'Only native candidate UI');assert.ok(s.fields.every(f=>!f.readOnly));console.log('PASS '+name);return s;};
const key=async k=>{await page.keyboard.press(k);await settle();};
const type=async t=>{await page.keyboard.type(t);await settle();};
const fresh=async()=>{await page.getByRole('button',{name:'新建',exact:true}).click();await settle();};
const seed=async latex=>{await fresh();await page.locator('[data-classic-bottom-view="source"]').click();await page.locator('.cm-content').click();await page.keyboard.press('Control+a');await page.keyboard.insertText('$$\n'+latex+'\n$$');await settle();await page.locator('.mathfield-host').first().click({position:{x:5,y:20}});await settle();};
const atom=async(predicate,side='center',count=1)=>{
  const rect=await page.locator('math-field').first().evaluate((f,{expression,side})=>{
    const match=Function('a','return ('+expression+')');const a=f._mathfield.model.root.children.find(match);if(!a)throw Error('No matching atom: '+expression);
    const node=f.shadowRoot.querySelector('[data-atom-id="'+a.id+'"]');if(!node)throw Error('Atom has no rendered target');return node.getBoundingClientRect().toJSON();
  },{expression:predicate,side});
  await page.mouse.click(rect.x+rect.width*(side==='after'?.8:side==='before'?.2:.5),rect.y+rect.height/2,{clickCount:count});await settle();
};
const ime=async text=>{await cdp.send('Input.imeSetComposition',{text,selectionStart:text.length,selectionEnd:text.length});await page.waitForTimeout(40);await cdp.send('Input.insertText',{text});await settle();};
const choose=async(predicate,method='keyboard')=>{
  const candidates=(await state()).candidates;const index=candidates.findIndex(predicate);assert.ok(index>=0,'Candidate exists');
  if(method==='mouse'){await page.locator('#mathlive-suggestion-popover li').nth(index).click();await settle();return;}
  const current=candidates.findIndex(x=>x.current);for(let i=0;i<(index-current+candidates.length)%candidates.length;i++)await key('ArrowDown');await key('Enter');
};
const branch=(s,expected,value)=>{const f=s.fields[0];const a=f.atoms.find(a=>a.value===value&&a.type!=='first');assert.ok(a,`Missing ${value}: ${f.value}`);assert.equal(a.branch,expected,`${value} must stay in ${expected}: ${f.value}`);};
const same=(actual,expected)=>assert.equal(compact(actual),compact(expected));
const runs=[];
const run=async(name,body)=>{if(filter!=='all'&&!name.includes(filter))return;console.log('CASE '+name);await body();runs.push(name);};
try {
 await fs.writeFile(path.join(output,'before.json'),JSON.stringify(await state(),null,2));
 await run('semantic-candidates',async()=>{
  for(const query of ['int','del','frac','lim','sum','sqrt','mat','vec']){
   await fresh();await type('\\'+query);
   await check(query+' native semantic candidates',s=>{
    assert.ok(s.nativeVisible);assert.ok(s.candidates.length>0);
    assert.equal(new Set(s.candidates.map(c=>c.latex)).size,s.candidates.length);
    if(query==='int')for(const part of ['\\int','\\int_{','\\oint'])assert.ok(s.candidates.some(c=>c.latex.startsWith(part)));
    if(query==='del')for(const latex of ['\\delta','\\Delta'])assert.ok(s.candidates.some(c=>c.latex===latex));
   });
   if(query==='int'||query==='del')await page.screenshot({path:path.join(output,query+'-candidates.png')});
   if(query==='del'){await choose(c=>c.latex==='\\Delta');await check('del accepts uppercase Delta',s=>same(s.fields[0].value,'\\Delta'));await type('x');await check('semantic non-prefix has no raw residue',s=>same(s.fields[0].value,'\\Delta x'));}
   else await key('Escape');
  }
 });
 await run('bounded-integral-input-paths',async()=>{
  for(const method of ['keyboard','mouse'])for(const input of ['key','text','ime']){
   await fresh();await type('\\int');await choose(c=>c.latex.startsWith('\\int_{')&&c.latex.includes('\\mathrm{d}'),method);
   await check(`${method}/${input}: initial lower selection`,s=>{const f=s.fields[0];assert.equal(f.atoms.filter(a=>a.command==='\\int').length,1);assert.ok(f.atoms.some(a=>a.selected&&a.branch==='subscript'));});
   if(input==='key')await type('a');else if(input==='text')await page.keyboard.insertText('a');else await ime('a');
   await check(`${method}/${input}: first input lower`,s=>branch(s,'subscript','a'));
   await type('q');await check(`${method}/${input}: consecutive input stays lower`,s=>branch(s,'subscript','q'));
   await key('Tab');await type('b');await check(`${method}/${input}: Tab input upper`,s=>branch(s,'superscript','b'));
  }
 });
 await run('fraction-placeholder-delete',async()=>{
  for(const input of ['key','text','ime'])for(const deletion of ['Backspace','Delete','selected']){
   await fresh();await type('\\frac');await key('Enter');
   if(input==='key')await type('a');else if(input==='text')await page.keyboard.insertText('a');else await ime('a');
   await check(`${input}/${deletion}: numerator input`,s=>branch(s,'above','a'));
   await atom('a.value === "a"','after');
   if(deletion==='Delete')await key('ArrowLeft');if(deletion==='selected')await key('Shift+ArrowLeft');
   await key(deletion==='selected'?'Backspace':deletion);
   await check(`${input}/${deletion}: placeholder retained in numerator`,s=>{same(s.fields[0].value,'\\frac{\\placeholder{}}{\\placeholder{}}');assert.ok(s.fields[0].atoms.some(a=>a.type==='placeholder'&&a.selected&&a.branch==='above'));});
   await type('x');await check(`${input}/${deletion}: replacement stays numerator`,s=>branch(s,'above','x'));
   await key('Control+z');await check(`${input}/${deletion}: undo typing`,s=>assert.ok(!s.fields[0].atoms.some(a=>a.value==='x')));
   await key('Control+y');await check(`${input}/${deletion}: redo typing`,s=>branch(s,'above','x'));
  }
 });
 await run('differential-backslash',async()=>{
  for(const placement of ['selected','collapsed']){
   await seed('\\int_0^1 f\\,\\mathrm{d}\\placeholder{}');await atom('a.type === "placeholder"');if(placement==='collapsed')await key('ArrowRight');
   await type('\\');await check(placement+': backslash consumes differential placeholder',s=>{assert.ok(!s.fields[0].atoms.some(a=>a.type==='placeholder'));assert.equal(s.fields[0].mode,'latex');assert.equal(s.fields[0].phantomCarets,0);});
   await type('del');await choose(c=>c.latex==='\\Delta');await check(placement+': semantic command stays after differential',s=>same(s.fields[0].value,'\\int_0^1 f\\,\\mathrm{d}\\Delta'));
  }
 });
 await run('nested-branches',async()=>{
  for(const latex of ['\\frac{\\sqrt{\\placeholder{}}}{y}+z','\\sqrt{\\frac{\\placeholder{}}{b}}+z','\\vec{\\frac{\\placeholder{}}{b}}+z','\\int_{\\frac{\\placeholder{}}{b}}^1 f\\,\\mathrm{d}x','\\begin{pmatrix}\\frac{\\placeholder{}}{b}&c\\\\d&e\\end{pmatrix}']){
   await seed(latex);await atom('a.type === "placeholder"');await type('\\del');await choose(c=>c.latex==='\\Delta','mouse');
   await check('nested semantic replacement '+latex,s=>{assert.equal(s.fields[0].atoms.filter(a=>a.value==='Δ').length,1);assert.equal(s.fields[0].atoms.filter(a=>a.type==='latexgroup').length,0);});
   await atom('a.value === "Δ"','after');await key('Backspace');await check('nested deletion retains branch '+latex,s=>{same(s.fields[0].value,latex);});
  }
 });
 await run('history-and-fast-completion',async()=>{
  await fresh();await type('\\int');await choose(c=>c.latex.startsWith('\\int_{')&&c.latex.includes('\\mathrm{d}'));
  const original=(await state()).fields[0].value;await key('Control+z');await check('completion undo is atomic',s=>same(s.fields[0].value,''));await key('Control+y');await check('completion redo retains lower selection',s=>{same(s.fields[0].value,original);assert.ok(s.fields[0].atoms.some(a=>a.selected&&a.branch==='subscript'));});
  await type('x');await check('immediate post-redo typing',s=>branch(s,'subscript','x'));
  for(let i=0;i<3;i++){await fresh();await type('\\frac');await page.keyboard.press('Enter');await page.keyboard.type('a');await check('fast completion then input '+i,s=>branch(s,'above','a'));}
  await seed('\\frac{\\placeholder{}}{b}');await atom('a.type === "placeholder"');await type('\\sqrt');await key('Enter');await type('x');await check('nested command keyboard confirmation',s=>assert.ok(s.fields[0].value.includes('\\sqrt{x}')));
 });
 await run('required-arguments-and-command-families',async()=>{
  for(const query of ['sqrt','vec','hat','widehat','overline','underline','binom','overset','underset','matrix','cases','ket','bra','braket','left','partial','oint','iint','prod']){
   await fresh();await type('\\'+query);await check(query+' list renders without exception',s=>assert.ok(s.nativeVisible&&s.candidates.length));
   const first=(await state()).candidates[0].latex;await key('Enter');
   await check(query+' accepts a complete structure',s=>{assert.equal(s.fields[0].mode,'math');assert.ok(s.fields[0].value,'completion must not silently disappear');if(first.includes('\\placeholder'))assert.ok(s.fields[0].atoms.some(a=>a.type==='placeholder'));});
   await type('x');await check(query+' accepts a character once',s=>{assert.equal(s.fields[0].atoms.filter(a=>a.value==='x').length,1);assert.ok(!s.fields[0].atoms.some(a=>a.type==='latexgroup'));});
  }
 });
 await run('source-toggle-multiline-persistence',async()=>{
  await fresh();await type('a');await key('Shift+Enter');await type('b');
  await check('Shift Enter retains one multiline field',s=>{assert.equal(s.fields.length,1);assert.match(s.fields[0].value,/gathered/);assert.ok(s.fields[0].atoms.some(a=>a.value==='a')&&s.fields[0].atoms.some(a=>a.value==='b'));});
  await key('Enter');await type('c');await key('Control+a');await type('x');
  await check('all visual rows replace together',s=>{assert.equal(s.fields.length,1);same(s.fields[0].value,'x');});
  await key('Control+z');await check('undo restores every visual row',s=>{assert.equal(s.fields.length,2);assert.ok(s.fields[0].value.includes('gathered'));same(s.fields[1].value,'c');});
  await seed('\\frac{a}{b}+z');
  for(let i=0;i<3;i++){
   await page.locator('.cm-content').click();await page.locator('[data-classic-bottom-view="source"]').click();
   await atom('a.value === "a"','after');await type('q');await check('source focus release '+i,s=>branch(s,'above','q'));await key('Backspace');
   await page.locator('[data-classic-bottom-view="source"]').click();
  }
  await check('source panel cycles preserve formula',s=>same(s.fields[0].value,'\\frac{a}{b}+z'));
  await page.reload();await page.waitForSelector('math-field');await check('WebView reload restores document',s=>same(s.fields[0].value,'\\frac{a}{b}+z'));
  await atom('a.value === "a"','after');await type('q');await check('input works after restored state',s=>branch(s,'above','q'));
 });
 await run('software-keyboard-controls',async()=>{
  const input=async(key,vk,modifiers=0)=>{for(const type of ['keyDown','keyUp'])await cdp.send('Input.dispatchKeyEvent',{type,key,code:'',windowsVirtualKeyCode:vk,modifiers});await settle();};
  await fresh();await type('\\frac');await input('Enter',13);await check('code-less Enter accepts native candidate',s=>{assert.equal(s.fields[0].mode,'math');assert.ok(s.fields[0].value.startsWith('\\frac'));});
  await type('a');await input('Backspace',8);await check('code-less Backspace retains numerator',s=>same(s.fields[0].value,'\\frac{\\placeholder{}}{\\placeholder{}}'));
  await type('x');await input('Tab',9);await type('b');await check('code-less Tab advances to denominator',s=>branch(s,'below','b'));
  await input('ArrowLeft',37);await input('Delete',46);await check('code-less Delete removes denominator locally',s=>same(s.fields[0].value,'\\frac{x}{\\placeholder{}}'));
  await fresh();await type('\\int');await input('ArrowDown',40);await input('ArrowUp',38);await choose(c=>c.latex.startsWith('\\int_{')&&c.latex.includes('\\mathrm{d}'));
  await type('q');await input('Tab',9);await type('b');await check('code-less Tab preserves bounded operator order',s=>{branch(s,'subscript','q');branch(s,'superscript','b');});
 });
 await run('integral-in-existing-complex-formulas',async()=>{
  for(const prefix of ['\\ket{a}a\\ket{a}\\Delta=','\\frac{\\int_0^1 f(t)\\,dt}{\\sqrt{\\vec{v}^2+1}}+','\\int_0^1\\frac{f(t)}{1+t^2}\\,dt+'])for(const input of ['key','text','ime']){
   await seed(prefix);await key('Control+End');await type('\\int');await choose(c=>c.latex.startsWith('\\int_{')&&c.latex.includes('\\mathrm{d}'),'mouse');
   await check(input+' complex integral initially selects lower',s=>assert.ok(s.fields[0].atoms.some(a=>a.selected&&a.branch==='subscript'&&a.parentCommand==='\\int')));
   if(input==='ime')await ime('q');else if(input==='text')await page.keyboard.insertText('q');else await type('q');
   await check(input+' complex integral first input lower',s=>branch(s,'subscript','q'));await type('r');await check(input+' complex integral second input lower',s=>branch(s,'subscript','r'));
   await key('Tab');await type('B');await check(input+' complex integral upper input',s=>branch(s,'superscript','B'));
   await key('Control+z');await key('Control+y');await check(input+' complex integral redo preserves both limits',s=>{branch(s,'subscript','q');branch(s,'subscript','r');branch(s,'superscript','B');});
  }
 });
 await fs.writeFile(path.join(output,'result.json'),JSON.stringify({status:'PASS',runs,checks:steps.length,errors},null,2));
 console.log(JSON.stringify({status:'PASS',runs:runs.length,checks:steps.length}));
}catch(error){await fs.writeFile(path.join(output,'failure.json'),JSON.stringify({error:String(error),stack:error.stack,errors},null,2));await page.screenshot({path:path.join(output,'failure.png')});throw error;}
finally{
 try {
  // Restore even after an assertion fails. Retaining a failed fixture for manual
  // diagnosis is explicit, and the backup stays beside the evidence either way.
  if(!process.argv.includes('--keep-fixture')){
   await page.evaluate(storage=>{for(const key of ['visualtex-editor','visualtex-desktop-editor-source-open','visualtex-desktop-editor-toolbar-open']){if(storage[key]===undefined)localStorage.removeItem(key);else localStorage.setItem(key,storage[key]);}},savedStorage);
   await page.reload();await page.waitForSelector('math-field');await settle();
   const restored=await page.locator('math-field').evaluateAll(fields=>fields.map(f=>f.value));
   assert.deepEqual(restored,savedFormulas,'Restore pre-test formulas');
   await fs.writeFile(path.join(output,'restored.json'),JSON.stringify({status:'PASS',formulasRestored:restored.length}));
  }
 }finally{await browser.close();}
}
