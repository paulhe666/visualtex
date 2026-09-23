// Real installed WebView2 workflows. Preserve and restore the open source buffer.
// Usage: VISUALTEX_PLAYWRIGHT_MODULE=<module URL> node ... PORT OUTPUT CASE|all
// Only trusted keyboard/mouse edits; model/selection/store access is read-only.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
const {chromium}=await import(process.env.VISUALTEX_PLAYWRIGHT_MODULE||'playwright');
const [port='9227',output,scenario='all']=process.argv.slice(2);
assert.ok(output,'An evidence output directory is required');
await fs.mkdir(output,{recursive:true});
const browser=await chromium.connectOverCDP(`http://127.0.0.1:${Number(port)}`);
const page=browser.contexts()[0].pages().find(p=>p.url()==='http://tauri.localhost/');
assert.ok(page,'Expected installed VisualTeX');
page.setDefaultTimeout(6000);
const events=[],results=[],failures=[];
const actions=[];
let hung=false;
const debug=await browser.contexts()[0].newCDPSession(page);
await debug.send('Debugger.enable');
debug.on('Debugger.paused',async params=>{
  await fs.writeFile(path.join(output,'paused-stack.json'),JSON.stringify(params,null,2));
  const frame=params.callFrames[0];
  const source=await debug.send('Debugger.getScriptSource',{scriptId:frame.location.scriptId});
  await fs.writeFile(path.join(output,'paused-script.js'),source.scriptSource);
  console.log('Saved debugger pause stack');
});
const action=async(name,run)=>{
  actions.push({name,at:new Date().toISOString()});await fs.writeFile(path.join(output,'actions.json'),JSON.stringify(actions,null,2));
  let timer;
  try{return await Promise.race([run(),new Promise((_,reject)=>{timer=setTimeout(()=>{hung=true;void debug.send('Debugger.pause');reject(Error(`Unresponsive during ${name}`));},6000);})]);}
  finally{clearTimeout(timer);}
};
page.on('pageerror',error=>events.push({type:'pageerror',message:error.message,stack:error.stack}));
const settle=()=>page.waitForTimeout(160);
const state=()=>page.evaluate(()=>({
  active:document.activeElement?.tagName,
  source:document.querySelector('.cm-content')?.innerText,
  error:document.querySelector('.source-error-chip')?.textContent,
  title:document.querySelector('input[aria-label="公式文档标题"]')?.value,
  scripts:[...document.scripts].map(s=>s.src),
  fields:[...document.querySelectorAll('math-field')].map(f=>({
    value:f.value,readOnly:f.readOnly,focused:f.hasFocus(),position:f.position,
    lastOffset:f.lastOffset,parentEnvironment:f._mathfield?.model.parentEnvironment?.environmentName,
    selection:f.selection,selected:f.getValue(f.selection,'latex'),
    caret:[...f.shadowRoot.querySelectorAll('.ML__caret')].map(e=>e.getBoundingClientRect().toJSON()),
    errors:f.errors,
  })),
}));
let currentCase='prepare';
const capture=async(name,verify)=>{
  await settle();const s=await action('read '+name,state);results.push({case:currentCase,name,...s});
  await fs.writeFile(path.join(output,'steps.json'),JSON.stringify(results,null,2));
  console.log(JSON.stringify({case:currentCase,name,fields:s.fields.map(f=>({value:f.value,readOnly:f.readOnly,selected:f.selected,position:f.position})),error:s.error}));
  if(verify)verify(s);
  assert.deepEqual(events,[],`${name}: runtime exception`);
  return s;
};
const shot=async name=>page.screenshot({path:path.join(output,`${currentCase}-${name}.png`)});
const sourcePanel=async()=>{await page.locator('[data-classic-bottom-view="source"]').click();await settle();};
const setSource=async text=>{await sourcePanel();await page.locator('.cm-content').click();await page.keyboard.press('Control+a');await page.keyboard.insertText(text);await settle();};
const seed=async(...values)=>{await page.getByRole('button',{name:'新建',exact:true}).click();await setSource(values.map(v=>`$$\n${v}\n$$`).join('\n\n'));};
const visual=async(index=0)=>{await page.locator('.mathfield-host').nth(index).click({position:{x:5,y:20}});await settle();};
const command=async id=>{
  await page.locator('[data-classic-bottom-view="tools"]').click();
  const button=page.locator(`.classic-bottom-toolbar [data-command-id="${id}"]`).first();
  await button.scrollIntoViewIfNeeded();await button.click();await settle();
};
const key=async name=>{await action('key '+name,()=>page.keyboard.press(name));await settle();};
const type=async text=>{await action('type '+text,()=>page.keyboard.type(text));await settle();};
const values=s=>s.fields.map(f=>f.value);
const compact=s=>s.replace(/\s+/g,'').replace(/([_^])\{([A-Za-z0-9])\}/g,'$1$2');
const same=(actual,expected)=>assert.equal(compact(actual),compact(expected));
const editable=s=>{assert.ok(s.fields.some(f=>f.focused));assert.ok(s.fields.every(f=>!f.readOnly));assert.ok(!s.error);};
const clickAtom=async(text,index=0,side='after')=>{
  const candidates=await page.locator('math-field').nth(index).evaluate((f,text)=>[...f.shadowRoot.querySelectorAll('span')]
    .filter(e=>!e.querySelector('span')&&e.textContent===text)
    .map(e=>({text:e.textContent,...e.getBoundingClientRect().toJSON()})).filter(r=>r.width&&r.height),text);
  if(candidates.length!==1){await fs.writeFile(path.join(output,`${currentCase}-atom-inspection.json`),JSON.stringify(await page.locator('math-field').nth(index).evaluate(f=>[...f.shadowRoot.querySelectorAll('span')].filter(e=>e.textContent.length<8).map(e=>({html:e.outerHTML,text:e.textContent,rect:e.getBoundingClientRect().toJSON()}))),null,2));}
  assert.equal(candidates.length,1,`Expected a unique visible ${text}: ${JSON.stringify(candidates)}`);
  const r=candidates[0],x=r.x+r.width*(side==='after'?0.85:0.15),y=r.y+r.height/2;
  actions.push({name:'click glyph '+text,rect:r,x,y,nativeOffset:await page.locator('math-field').nth(index).evaluate((f,{x,y})=>f.getOffsetFromPoint(x,y,{bias:0}),{x,y})});
  await fs.writeFile(path.join(output,'actions.json'),JSON.stringify(actions,null,2));
  await page.mouse.click(x,y);await settle();
};
const replay=async(expected,redo=false)=>{await key(redo?'Control+y':'Control+z');const s=await capture(redo?'Redo':'Undo');assert.deepEqual(values(s),expected);editable(s);};
const selectAccentInk=async()=>{
  await page.waitForTimeout(550);
  const ink=page.locator('math-field').first().locator('.ML__visualtex-accent-mark').first();
  const box=await ink.evaluate(node=>{const svg=node.querySelector('svg');if(svg){const b=svg.getBBox(),m=svg.getScreenCTM();if(m&&b.width&&b.height){const tl=new DOMPoint(b.x,b.y).matrixTransform(m),br=new DOMPoint(b.x+b.width,b.y+b.height).matrixTransform(m);return {x:tl.x,y:tl.y,width:br.x-tl.x,height:br.y-tl.y};}}const r=node.getBoundingClientRect();return {x:r.x,y:r.y,width:r.width,height:r.height};});assert.ok(box);
  actions.push({name:'accent ink hit',box,hit:await page.locator('math-field').first().evaluate((f,box)=>{const hit=f.shadowRoot.elementFromPoint(box.x+box.width/2,box.y+box.height/2);return {html:hit?.outerHTML,parent:hit?.parentElement?.outerHTML};},box)});await fs.writeFile(path.join(output,'actions.json'),JSON.stringify(actions,null,2));
  await page.mouse.dblclick(box.x+box.width/2,box.y+box.height/2);await settle();
};
const fieldEnd=async(index=0)=>{
  const box=await page.locator('.mathfield-host').nth(index).boundingBox();
  await page.mouse.click(box.x+box.width-4,box.y+box.height/2);await settle();
  const f=await page.locator('math-field').nth(index).evaluate(f=>({position:f.position,last:f.lastOffset}));
  assert.equal(f.position,f.last,'Clicking row end should reach the root end');
};
const quiescence=async()=>{
  const observations=await action('idle DOM and event-loop check',()=>page.evaluate(()=>new Promise(resolve=>{
    let ticks=0,mutations=0;const observer=new MutationObserver(records=>mutations+=records.length);
    for(const f of document.querySelectorAll('math-field'))observer.observe(f.shadowRoot,{subtree:true,attributes:true,attributeFilter:['class'],childList:true});
    const timer=setInterval(()=>ticks++,40);setTimeout(()=>{observer.disconnect();clearInterval(timer);resolve({ticks,mutations});},320);
  })));
  results.push({case:currentCase,name:'idle DOM and event-loop check',...observations});
  assert.ok(observations.ticks>=3);assert.ok(observations.mutations<10,JSON.stringify(observations));
};
const cases={
  async history_fork_and_selection(){
    const original=[String.raw`\frac{a+\hat{q}}{\sqrt{1+b}}`,String.raw`\vec{z}+\Delta`];
    await seed(...original);await visual();await clickAtom('q');await key('Shift+ArrowLeft');
    await page.locator('[data-formula-selection-bold]').click();await settle();
    const styled=values(await capture('style before source undo'));
    assert.notEqual(styled[0],original[0]);assert.equal(styled[1],original[1]);
    await sourcePanel();await page.locator('.cm-content').click();await page.keyboard.press('Control+a');await page.keyboard.insertText(String.raw`\frac{`);await settle();
    await page.keyboard.press('Control+z');await settle();await visual();
    await capture('source-local undo returns styled document',s=>assert.deepEqual(values(s),styled));
    await replay(original);
    await fieldEnd(1);await type('w');
    const branched=values(await capture('edit after undo creates new history branch',s=>{same(s.fields[0].value,original[0]);same(s.fields[1].value,original[1]+' w');}));
    await key('Control+y');await capture('Redo cannot resurrect discarded style branch',s=>assert.deepEqual(values(s),branched));
    await replay(original);
    await clickAtom('b');await key('Shift+ArrowLeft');await capture('select denominator glyph before delete');
    await key('Backspace');
    await action('Undo selected deletion and replace immediately',async()=>{await page.keyboard.press('Control+z');await page.keyboard.type('c');});await settle();
    await capture('Undo restores replacement selection at correct depth',s=>{same(s.fields[0].value,String.raw`\frac{a+\hat{q}}{\sqrt{1+c}}`);same(s.fields[1].value,original[1]);});
    await shot('history-fork');
  },
  async accent_geometry(){
    const originalZoom=parseInt(await page.locator('.canvas-controls [aria-live]').innerText(),10);
    const setZoom=async target=>{let current=parseInt(await page.locator('.canvas-controls [aria-live]').innerText(),10);while(current!==target){await page.getByRole('button',{name:current<target?'放大公式':'缩小公式',exact:true}).click();const next=parseInt(await page.locator('.canvas-controls [aria-live]').innerText(),10);assert.notEqual(next,current);current=next;}await settle();};
    const setFont=async value=>{await page.getByRole('button',{name:'设置',exact:true}).click();try{await page.locator('[data-interface-customization-trigger]').click();const select=page.locator('[data-formula-letter-font-setting]');await select.scrollIntoViewIfNeeded();const original=await select.inputValue();await select.selectOption(value);return original;}finally{const customClose=page.getByRole('button',{name:'关闭界面自定义',exact:true});if(await customClose.count())await customClose.click();await page.getByRole('button',{name:'关闭设置',exact:true}).click();await settle();}};
    const latex=String.raw`\vec{\placeholder{}}+\hat{\placeholder{}}+\dddot{\placeholder{}}+\ddddot{\placeholder{}}+\textcolor{red}{\vec{q}}+\widehat{ab}+\widetilde{xy}`;
    await seed(latex);await visual();let originalFont;
    try {
      for(const font of ['katex','times','cambria']){
        const previous=await setFont(font);originalFont??=previous;
        for(const zoom of [40,100,160]){
          await setZoom(zoom);await visual();
          const geometry=await page.locator('math-field').first().evaluate(f=>[...f.shadowRoot.querySelectorAll('.visualtex-accent-structural-placeholder')].map(p=>{
            const l=p.closest('.visualtex-accent-placeholder-layout'),a=l.querySelector(':scope > .ML__center .ML__accent-body'),pr=p.getBoundingClientRect(),ar=a.getBoundingClientRect();
            const svg=a.querySelector('svg');let ink=ar.toJSON();
            if(svg){const b=svg.getBBox(),m=svg.getScreenCTM(),tl=new DOMPoint(b.x,b.y).matrixTransform(m),br=new DOMPoint(b.x+b.width,b.y+b.height).matrixTransform(m);ink={x:tl.x,y:tl.y,width:br.x-tl.x,height:br.y-tl.y};}
            return {placeholder:pr.toJSON(),accent:ar.toJSON(),ink,centerError:Math.abs(pr.x+pr.width/2-(ink.x+ink.width/2)),fontSize:parseFloat(getComputedStyle(p).fontSize)};
          }));
          results.push({case:currentCase,name:`${font} ${zoom}% geometry`,geometry});
          assert.equal(geometry.length,4);
          for(const g of geometry){assert.ok(g.centerError<1.5,JSON.stringify({font,zoom,...g}));assert.ok(g.ink.width>0&&g.ink.height>0);}
          const red=await page.locator('math-field').first().evaluate(f=>({ink:[...f.shadowRoot.querySelectorAll('.ML__visualtex-accent-ink svg')].map(svg=>getComputedStyle(svg).color),glyph:getComputedStyle([...f.shadowRoot.querySelectorAll('span')].find(e=>e.textContent==='q'&&!e.querySelector('span'))).color}));
          assert.equal(red.ink.at(-1),red.glyph);assert.notEqual(red.glyph,red.ink[0]);
          await shot(`${font}-${zoom}`);
        }
      }
    }finally {if(originalFont)await setFont(originalFont);await setZoom(originalZoom);}
  },
  async accent_native(){
    for(const accent of (process.env.VISUALTEX_ACCENT_CASES?.split(',')||['vec','hat','acute','grave','dot','ddot','dddot','ddddot','breve','check','bar','tilde','widehat','widetilde'])) {
      const template=`\\frac{\\sqrt{1+\\${accent}{\\placeholder{}}}}{1+y}`;
      await seed(template);await visual();await key('Tab');
      await capture(`${accent} placeholder selected`,s=>same(s.fields[0].selected,String.raw`\placeholder{}`));
      const layout=await page.locator('math-field').first().evaluate(f=>({owners:f.shadowRoot.querySelectorAll('.visualtex-accent-placeholder-layout').length,markers:f.shadowRoot.querySelectorAll('.visualtex-accent-structural-placeholder').length,overlays:f.shadowRoot.querySelectorAll('.visualtex-combining-accent-overlay').length,shifts:[...f.shadowRoot.querySelectorAll('[style]')].filter(e=>e.style.getPropertyValue('--visualtex-accent-placeholder-shift')).length}));
      assert.deepEqual(layout,{owners:1,markers:1,overlays:0,shifts:0});
      await type('x');
      const typed=values(await capture(`${accent} finalized without placeholder geometry`,s=>same(s.fields[0].value,template.replace(String.raw`\placeholder{}`,'x'))));
      assert.equal(await page.locator('math-field').first().evaluate(f=>f.shadowRoot.querySelectorAll('.visualtex-accent-placeholder-layout,.visualtex-accent-structural-placeholder').length),0);
      await replay([template]);await replay(typed,true);
      await clickAtom('x');await key('Shift+ArrowLeft');
      await capture(`${accent} body selection`,s=>same(s.fields[0].selected,'x'));
      await key('Backspace');await type('x');
      await capture(`${accent} replace body preserves accent`,s=>same(s.fields[0].value,template.replace(String.raw`\placeholder{}`,'x')));
      await selectAccentInk();
      await capture(`${accent} select complete structure by its mark`,s=>assert.ok(s.fields[0].selected.includes(`\\${accent}`)));
      await key('Backspace');await type('z');
      await capture(`${accent} remove selected accent and replace`,s=>same(s.fields[0].value,String.raw`\frac{\sqrt{1+z}}{1+y}`));
    }
    await shot('accent-native');
  },
  async vec_persistence(){
    const latex=String.raw`\vec{}\vec{}+\frac{\vec{}}{\hat{}}`;
    await seed(latex);await visual();
    for(let i=0;i<8;i++){
      await fieldEnd();await type('w');
      const typed=values(await capture(`vec cycle ${i+1} append`,s=>same(s.fields[0].value,latex+'w')));
      await replay([latex]);await replay(typed,true);await replay([latex]);
      await sourcePanel();await page.locator('.cm-content').click();await visual();
      if(i%2===1){await page.reload();await page.waitForSelector('math-field');await visual();}
      await capture(`vec cycle ${i+1} stable after focus or reload`,s=>same(s.fields[0].value,latex));
    }
    await shot('vec-persistence');
  },
  async empty_structures(){
    for(const [i,latex] of [String.raw`\vec{}\vec{}`,String.raw`\hat{}+\vec{}+\sqrt{}+\frac{}{}`,String.raw`a+\vec{}+b`,String.raw`\begin{gathered}\vec{}\\\hat{}\end{gathered}`].entries()) {
      await seed(latex);await visual();await key('Control+a');
      await capture(`empty ${i+1} select all`);
      await key('Backspace');
      await capture(`empty ${i+1} delete all structures`,s=>same(s.fields[0].value,''));
      await type('x');await capture(`empty ${i+1} replace with x`,s=>same(s.fields[0].value,'x'));
      await replay(['']);await replay([latex]);
      await key('Control+a');await type('y');
      await capture(`empty ${i+1} selected replacement`,s=>same(s.fields[0].value,'y'));
      await replay([latex]);await replay(['y'],true);
    }
    await shot('empty-structures');
  },
  async toolbar(){
    await page.getByRole('button',{name:'新建',exact:true}).click();
    await command('frac');await capture('outer fraction');
    await command('sqrt');await capture('nested square root');
    await command('vector');await capture('nested vector');
    await type('v');await capture('fill vector',s=>same(s.fields[0].value,String.raw`\frac{\sqrt{\vec{v}}}{\placeholder{}}`));
    await key('Tab');await command('nthroot');await capture('cube root denominator');
    await type('3');await key('Tab');await type('x');
    const complete=await capture('complete nested fraction',s=>{editable(s);same(s.fields[0].value,String.raw`\frac{\sqrt{\vec{v}}}{\sqrt[3]{x}}`);});
    await sourcePanel();await visual();await capture('source visual roundtrip',s=>assert.deepEqual(values(s),values(complete)));
    await shot('complete');
  },
  async nested_edit(){
    const latex=String.raw`\left(\frac{\sqrt{1+\hat{q}_{i}^{2}}}{1+\frac{x}{y}}\right)^{n}`;
    await seed(latex);await visual();const original=values(await capture('seed deep formula',editable));
    await clickAtom('q');await capture('click nested q');await selectAccentInk();
    await capture('select nested accented term',s=>same(s.fields[0].selected,String.raw`\hat{q}`));
    await command('vector');const wrapped=values(await capture('wrap selected accented term in vector',s=>{editable(s);assert.match(s.fields[0].value,/\\vec\{\\hat\{q\}\}/);}));
    await replay(original);await replay(wrapped,true);
    await clickAtom('y');await key('Backspace');await type('r');
    await capture('replace deep denominator with Backspace then type',s=>{editable(s);same(s.fields[0].value,wrapped[0].replace('{y}','{r}'));});
    await shot('edited');
  },
  async selection_style(){
    await seed(String.raw`\frac{\hat{\Delta}+q}{\sqrt{x_{i}^{2}+1}}`,String.raw`\vec{z}+\Delta`);
    await visual();const original=values(await capture('two formulas'));
    await clickAtom('Δ');await key('Shift+ArrowLeft');
    await capture('select nested Delta inside accent',s=>same(s.fields[0].selected,String.raw`\Delta`));
    await page.getByRole('button',{name:'切换选中内容的粗体状态',exact:true}).click();
    const styled=values(await capture('bold nested Delta only',s=>{assert.match(s.fields[0].value,/\\mathbf\{\\Delta\}/);assert.equal(s.fields[1].value,original[1]);}));
    await page.locator('[data-formula-selection-italic]').click();
    await capture('bold italic nested Delta',s=>{assert.match(s.fields[0].value,/\\mathbfit\{\\Delta\}/);assert.equal(s.fields[1].value,original[1]);});
    await page.locator('[data-formula-selection-italic]').click();
    await capture('remove italic but keep bold nested Delta',s=>assert.deepEqual(values(s),styled));
    await page.getByRole('button',{name:'切换选中内容的粗体状态',exact:true}).click();
    await capture('remove bold preserves nested accent',s=>assert.deepEqual(values(s),original));
    await page.getByRole('button',{name:'切换选中内容的粗体状态',exact:true}).click();
    await visual(1);await fieldEnd(1);await type('w');
    const second=values(await capture('edit unselected second formula',s=>{assert.equal(s.fields[0].value,styled[0]);assert.equal(s.fields[1].value,original[1]+' w');}));
    await replay(styled);await replay(second,true);await shot('styled');
  },
  async multiline(){
    await seed(String.raw`\frac{a}{\sqrt{b+1}}`);await visual();await fieldEnd();
    await key('Shift+Enter');await command('parentheses');await command('frac');
    await type('c');await key('Tab');await command('vector');await type('d');
    const two=values(await capture('two different nested rows',s=>{editable(s);assert.equal(s.fields.length,1);same(s.fields[0].value,String.raw`\begin{gathered}\frac{a}{\sqrt{b+1}}\\\left(\frac{c}{\vec{d}}\right)\end{gathered}`);}));
    await sourcePanel();await page.locator('.cm-content').click();await visual();
    await capture('multiline source focus release',s=>{editable(s);assert.deepEqual(values(s),two);});
    await fieldEnd();await key('Shift+Enter');await capture('empty third row before typing');
    await page.locator('math-field').evaluate(f=>{window.__nestedEventTrace=[];for(const name of ['beforeinput','input','selection-change'])f.addEventListener(name,e=>window.__nestedEventTrace.push({name,value:f.value,position:f.position,selection:f.selection,data:e.data,inputType:e.inputType}),true);});
    await type('z');
    await capture('third row',s=>{editable(s);assert.equal((s.fields[0].value.match(/\\begin\{gathered\}/g)||[]).length,1);assert.match(s.fields[0].value,/\\\\\s*z\\end\{gathered\}$/);});
    await key('Control+z');await capture('Undo third row input');await type('w');
    await fs.writeFile(path.join(output,'multiline-events.json'),JSON.stringify(await page.evaluate(()=>window.__nestedEventTrace),null,2));
    await capture('Undo then immediate third row edit',s=>{editable(s);assert.match(s.fields[0].value,/\\\\\s*w\\end\{gathered\}$/);assert.doesNotMatch(s.fields[0].value,/z/);});
    await shot('multiline');
  },
  async source_draft(){
    const latex=String.raw`\frac{\hat{q}_{i}^{2}}{\sqrt{1+\frac{x}{y}}}`;
    await seed(latex);await visual();const original=values(await capture('valid nested source'));
    await setSource('$$\n'+latex.slice(0,-1)+'\n$$');await capture('temporarily missing outer brace');
    await page.keyboard.press('Control+z');await visual();await capture('source undo repairs brace',s=>{editable(s);assert.deepEqual(values(s),original);});
    await clickAtom('q');await type('r');const changed=values(await capture('visual edit after source repair'));
    assert.notDeepEqual(changed,original);
    await sourcePanel();await page.locator('.cm-content').click();await page.locator('[data-classic-bottom-collapse]').click();
    await visual();await key('Control+z');await capture('global undo after closing source',s=>assert.deepEqual(values(s),original));
    await sourcePanel();await visual();await shot('reopened');
  },
  async matrix_edit(){
    await seed(String.raw`\begin{pmatrix}\frac{q}{r}&\sqrt{\placeholder{}}\\\hat{s}&t_i\end{pmatrix}`);
    await visual();const original=values(await capture('matrix with empty radical',editable));
    await clickAtom('q');await key('Shift+ArrowLeft');await command('vector');
    const wrapped=values(await capture('wrap numerator inside matrix',s=>same(s.fields[0].value,original[0].replace('{q}',String.raw`{\vec{q}}`))));
    await replay(original);await replay(wrapped,true);
    await fs.writeFile(path.join(output,'matrix-geometry.json'),JSON.stringify(await page.locator('math-field').evaluate(f=>({atoms:Array.from({length:f.lastOffset+1},(_,i)=>({i,info:f.getElementInfo(i)})),dom:f.shadowRoot.innerHTML})),null,2));
    await shot('before-denominator-click');await clickAtom('r');await capture('click matrix fraction denominator');await shot('after-denominator-click');await key('Backspace');await type('2');
    const denominator=values(await capture('replace denominator inside matrix',s=>same(s.fields[0].value,wrapped[0].replace('{r}','{2}'))));
    await key('Tab');await type('b');
    await capture('fill radical in next matrix cell',s=>same(s.fields[0].value,denominator[0].replace(String.raw`\placeholder{}`,'b')));
    await key('Control+z');await type('c');
    const complete=values(await capture('undo and replace inside matrix radical',s=>{editable(s);same(s.fields[0].value,denominator[0].replace(String.raw`\placeholder{}`,'c'));}));
    await sourcePanel();await page.locator('.cm-content').click();await page.locator('[data-classic-bottom-collapse]').click();await visual();
    await sourcePanel();await visual();await capture('matrix survives panel close and reopen',s=>assert.deepEqual(values(s),complete));await shot('matrix');
  },
  async drag_delete(){
    await seed(String.raw`\frac{a+b}{\sqrt{c+d}}+\vec{e}`);await visual();
    const original=values(await capture('drag selection seed'));
    const boxes=await page.locator('math-field').evaluate(f=>Object.fromEntries(['a','b'].map(text=>{const e=[...f.shadowRoot.querySelectorAll('span')].find(e=>!e.querySelector('span')&&e.textContent===text);return [text,e.getBoundingClientRect().toJSON()];})));
    await page.mouse.move(boxes.a.x+1,boxes.a.y+boxes.a.height/2);await page.mouse.down();
    await page.mouse.move(boxes.b.right-1,boxes.b.y+boxes.b.height/2,{steps:12});await page.mouse.up();
    await capture('drag numerator only',s=>same(s.fields[0].selected,'a+b'));
    await command('hat');const wrapped=values(await capture('accent dragged numerator only',s=>same(s.fields[0].value,String.raw`\frac{\hat{a+b}}{\sqrt{c+d}}+\vec{e}`)));
    await replay(original);await replay(wrapped,true);
    await clickAtom('c',0,'before');await key('Delete');await type('q');
    const edited=values(await capture('forward delete in nested radical',s=>same(s.fields[0].value,String.raw`\frac{\hat{a+b}}{\sqrt{q+d}}+\vec{e}`)));
    await key('Control+z');await type('r');
    await capture('undo and replace deletion input at correct depth',s=>same(s.fields[0].value,edited[0].replace('q+d','r+d')));await shot('drag-delete');
  },
  async corpus(){
    const corpus=[
      String.raw`\overbrace{\frac{\sqrt{x_{i}^{2}+1}}{\hat{y}_{j}}+\vec{z}}^{n}`,
      String.raw`\int_{0}^{1}\frac{\sin\left(\sqrt{x^{2}+1}\right)}{1+\left(\frac{x}{2}\right)^{2}}\,\mathrm{d}x`,
      String.raw`\sum_{i=1}^{n}\frac{\vec{a}_{i}\cdot\hat{b}}{\sqrt[3]{1+x_{i}^{2}}}`,
      String.raw`\begin{pmatrix}\frac{q}{1+x}&\sqrt{b}\\\hat{c}&d_{i}^{2}\end{pmatrix}`,
      String.raw`f(x)=\begin{cases}\frac{x^{2}}{1+\sqrt{x}}&x>0\\\vec{a}_{i}&x\leq0\end{cases}`,
      String.raw`\left\langle\psi\middle|\frac{\hat{H}}{\sqrt{1+\hat{H}^{2}}}\middle|\phi\right\rangle`,
      String.raw`\begin{gathered}\frac{\sqrt{q}}{\vec{x}}\\\hat{y}_{i}^{2}+\left(\frac{a}{b}\right)\end{gathered}`,
      String.raw`\frac{\sqrt{\vec{}}}{\left(\hat{}+\frac{}{}\right)}`,
    ];
    for(let i=0;i<corpus.length;i++){
      await seed(corpus[i]);await visual();const baseline=values(await capture(`corpus ${i+1} source accepts`,editable));
      same(baseline[0],corpus[i]);
      await fieldEnd();await type('z');
      const edited=values(await capture(`corpus ${i+1} append z`,s=>{editable(s);same(s.fields[0].value,baseline[0]+'z');}));
      await replay(baseline);await replay(edited,true);
      await sourcePanel();await visual();await capture(`corpus ${i+1} synchronized`,s=>assert.deepEqual(values(s),edited));
      await page.reload();await page.waitForSelector('math-field');await visual();
      await capture(`corpus ${i+1} survives reload`,s=>{editable(s);assert.deepEqual(values(s),edited);});
      await shot(`corpus-${i+1}`);
    }
  },
  async scripts_and_matrix_power(){
    await seed(String.raw`\left(\frac{q_i^2+\sqrt{a}}{\vec{b}}\right)^n`);await visual();
    const original=values(await capture('nested scripts seed'));
    await clickAtom('2');await key('Backspace');await type('3');
    await capture('replace superscript inside numerator',s=>same(s.fields[0].value,original[0].replace('^2','^3')));
    await clickAtom('i');await key('Shift+ArrowLeft');await type('j');
    const edited=values(await capture('replace subscript without changing superscript',s=>same(s.fields[0].value,original[0].replace('^2','^3').replace('_i','_j'))));
    await clickAtom('n');await key('Shift+ArrowLeft');await command('frac');
    const fraction=values(await capture('wrap selected outer exponent in fraction',s=>assert.match(s.fields[0].value,/\^\{\\frac\{n\}\{\\placeholder\{\}\}\}/)));
    await type('2');
    const powered=values(await capture('fill fraction denominator in outer exponent',s=>same(s.fields[0].value,edited[0].replace(/\^\{?n\}?$/,String.raw`^{\frac{n}{2}}`))));
    await replay(fraction);await replay(edited);await replay(fraction,true);await replay(powered,true);
    await seed(String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}^{n}`);await visual();
    await clickAtom('n');await key('Shift+ArrowLeft');await type('2');
    await capture('matrix-level superscript remains clickable',s=>same(s.fields[0].value,String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}^{2}`));
    await shot('scripts');
  },
  async multiline_start(){
    await seed(String.raw`\begin{gathered}\frac{q}{r}\\\sqrt{s}\end{gathered}`);await visual();
    await capture('multiline root start',s=>assert.equal(s.fields[0].position,0));
    await key('Shift+Enter');await type('a');
    await capture('split at root start moves caret with original first row',s=>same(s.fields[0].value,String.raw`\begin{gathered}\\a\frac{q}{r}\\\sqrt{s}\end{gathered}`));
    await shot('multiline-start');
  },
  async local_row_break(){
    const original=String.raw`\frac{ab}{\hat{c}}+\sqrt{d}`;
    await seed(original);await visual();await clickAtom('a');
    await key('Shift+Enter');
    const split=values(await capture('split numerator without duplicating fraction',s=>same(s.fields[0].value,String.raw`\frac{\begin{gathered}a\\b\end{gathered}}{\hat{c}}+\sqrt{d}`)));
    await type('x');
    const typed=values(await capture('type in new numerator row',s=>same(s.fields[0].value,String.raw`\frac{\begin{gathered}a\\xb\end{gathered}}{\hat{c}}+\sqrt{d}`)));
    await replay(split);await replay([original]);await replay(split,true);await replay(typed,true);
    await sourcePanel();await page.locator('.cm-content').click();await visual();
    await capture('local multiline survives source focus round trip',s=>assert.deepEqual(values(s),typed));
    await shot('local-row-break');
  },
  async rapid_history(){
    const templates=[
      String.raw`\begin{gathered}\frac{a}{b}\\\placeholder{}\end{gathered}`,
      String.raw`\begin{pmatrix}\frac{q}{r}&\sqrt{\placeholder{}}\\\hat{s}&t_i\end{pmatrix}`,
      String.raw`\frac{\sqrt{\vec{\placeholder{}}}}{\sqrt[3]{x}}`,
    ];
    for(const [i,latex] of templates.entries()){
      await seed(latex);await visual();await key('Tab');
      await capture(`rapid ${i+1} selected placeholder`,s=>same(s.fields[0].selected,String.raw`\placeholder{}`));
      await type('z');
      await action(`rapid ${i+1} Undo then w without settle`,async()=>{await page.keyboard.press('Control+z');await page.keyboard.type('w');});
      await capture(`rapid ${i+1} correct branch immediately`,s=>{editable(s);same(s.fields[0].value,latex.replace(String.raw`\placeholder{}`,'w'));});
      await page.waitForTimeout(250);
      await capture(`rapid ${i+1} deferred focus repair preserves w`,s=>same(s.fields[0].value,latex.replace(String.raw`\placeholder{}`,'w')));
    }
    await shot('rapid-history');
  },
};
await sourcePanel();const before=await state();
await fs.writeFile(path.join(output,'before.json'),JSON.stringify(before,null,2));
await fs.writeFile(path.join(output,'before-source.tex'),before.source);
try{
  for(const name of scenario==='all'?Object.keys(cases):[scenario]){
    assert.ok(cases[name],`Unknown case ${name}`);currentCase=name;
    try{await cases[name]();await quiescence();console.log(`PASS ${name}`);}
    catch(error){failures.push({case:name,message:error.message,stack:error.stack});console.log(`FAIL ${name}: ${error.message}`);if(hung){await new Promise(r=>setTimeout(r,1000));break;}await capture('FAILURE STATE');await shot('failure');}
  }
}finally{
  await fs.writeFile(path.join(output,'result.json'),JSON.stringify({results,failures,events},null,2));
  try{if(!hung){await setSource(before.source);await visual();const restored=await state();assert.deepEqual(values(restored),values(before));await fs.writeFile(path.join(output,'restored.json'),JSON.stringify(restored,null,2));}}
  finally{await debug.send('Debugger.resume').catch(()=>{});await debug.detach();await browser.close();}
}
assert.deepEqual(failures,[],'Nested user workflows failed; inspect result.json and screenshots');
