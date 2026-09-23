// Real WebView2 pointer/keyboard selection workflows; model access is read-only.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
const {chromium}=await import(process.env.VISUALTEX_PLAYWRIGHT_MODULE||'playwright');
const [port='9227',output,scenario='probe']=process.argv.slice(2);
assert.ok(output);await fs.mkdir(output,{recursive:true});
const browser=await chromium.connectOverCDP(`http://127.0.0.1:${Number(port)}`);
const page=browser.contexts()[0].pages().find(p=>p.url()==='http://tauri.localhost/');assert.ok(page);
page.setDefaultTimeout(7000);
if(process.env.VISUALTEX_SELECTION_RELOAD==='1'){await page.reload();await page.waitForSelector('math-field');}
const errors=[],steps=[],failures=[];page.on('pageerror',e=>errors.push(e.message));
const settle=()=>page.waitForTimeout(180);
const sourcePanel=async()=>{await page.locator('[data-classic-bottom-view="source"]').click();await settle();};
const visual=async(index=0)=>{await page.locator('.mathfield-host').nth(index).click({position:{x:5,y:20}});await settle();};
const setSource=async text=>{await sourcePanel();await page.locator('.cm-content').click();await page.keyboard.press('Control+a');await page.keyboard.insertText(text);await settle();await visual();};
const seed=async(...values)=>{await page.getByRole('button',{name:'新建',exact:true}).click();await setSource(values.map(v=>`$$\n${v}\n$$`).join('\n\n'));};
const state=()=>page.evaluate(()=>({zoom:Number.parseInt(document.querySelector('.canvas-controls span')?.textContent),source:document.querySelector('.cm-content')?.innerText,title:document.querySelector('input[aria-label="公式文档标题"]')?.value,scripts:[...document.scripts].map(s=>s.src),fields:[...document.querySelectorAll('math-field')].map(f=>({value:f.value,position:f.position,lastOffset:f.lastOffset,selection:f.selection,selected:f.getValue(f.selection,'latex'),readOnly:f.readOnly,focused:f.hasFocus(),multi:f.classList.contains('has-visualtex-multi-line-selection'),rects:[...f.shadowRoot.querySelectorAll('.ML__selection')].map(e=>({rect:e.getBoundingClientRect().toJSON(),background:getComputedStyle(e).backgroundColor,display:getComputedStyle(e).display}))}))}));
const setZoom=async value=>{for(let n=0;n<100;n++){const zoom=(await state()).zoom;if(zoom===value)return;await page.getByRole('button',{name:zoom<value?'放大公式':'缩小公式',exact:true}).click();await settle();}throw Error('Cannot reach zoom '+value);};
const record=async name=>{await settle();const s=await state();steps.push({name,...s});await fs.writeFile(path.join(output,'steps.json'),JSON.stringify(steps,null,2));return s;};
const snapshot=async name=>{const s=await record(name);await page.screenshot({path:path.join(output,name+'.png')});return s;};
const glyph=async(text,index=0)=>{
  const result=await page.locator('math-field').nth(index).evaluate((f,text)=>{
    const nodes=[...f.shadowRoot.querySelectorAll('span')].filter(e=>!e.querySelector('span')&&e.textContent===text&&e.getBoundingClientRect().width);
    if(nodes.length!==1)return {count:nodes.length,nodes:nodes.map(e=>e.outerHTML)};
    const node=nodes[0],atom=f._mathfield.model.root.children.find(a=>a.id===node.dataset.atomId);
    return {count:1,rect:node.getBoundingClientRect().toJSON(),offset:atom?f._mathfield.model.offsetOf(atom):null,parent:atom?.parent.type,branch:atom?.parentBranch};
  },text);
  assert.equal(result.count,1,'unique glyph '+text+' '+JSON.stringify(result));return result;
};
const clickGlyph=async(text,index=0,side=.8,count=1)=>{await page.waitForTimeout(550);const g=await glyph(text,index);await page.mouse[count===2?'dblclick':'click'](g.rect.x+g.rect.width*side,g.rect.y+g.rect.height/2);await settle();return g;};
const formulas=[String.raw`\frac{a+\sqrt{b_i^2+c}}{\hat{d}+e_f}+g`,String.raw`\begin{pmatrix}\frac{h}{j}&\sqrt{k}\\\vec{m}&n_p\end{pmatrix}+q`];
await sourcePanel();const before=await state();await fs.writeFile(path.join(output,'before.json'),JSON.stringify(before,null,2));
try {
  if(scenario==='probe'||scenario==='verify'){
    await seed(...formulas);await clickGlyph('c');await page.keyboard.press('Control+a');
    const all=await snapshot('select-all-before');
    if(scenario==='verify')assert.ok(all.fields.every(f=>f.multi&&f.selection.ranges[0][0]===0&&f.selection.ranges[0][1]===f.lastOffset&&f.rects.length===1),'all rows and one coherent highlight per row');
    await fs.writeFile(path.join(output,'geometry.json'),JSON.stringify(await page.locator('math-field').first().evaluate(f=>f._mathfield.model.root.children.map(a=>({offset:f._mathfield.model.offsetOf(a),id:a.id,type:a.type,value:a.value,parent:a.parent?.type,branch:a.parentBranch,bounds:f.getElementInfo(f._mathfield.model.offsetOf(a))?.bounds,nodes:[...f.shadowRoot.querySelectorAll('[data-atom-id="'+a.id+'"]')].map(e=>({html:e.outerHTML,rect:e.getBoundingClientRect().toJSON()}))}))),null,2));
    await visual();
    for(const [i,texts] of [[0,['a','b','i','2','c','d','e','f','g']],[1,['h','j','k','m','n','p','q']]]){
      for(const text of texts){
        for(const side of [.2,.8]){const target=await clickGlyph(text,i,side);const s=await record(`point-${text}-${side}`);steps.at(-1).target=target;if(scenario==='verify')assert.equal(s.fields[i].position,target.offset-(side<.5?1:0),`caret ${text} ${side}`);}
        await clickGlyph(text,i,.5,2);const selected=await snapshot(`double-${text}`);
        if(scenario==='verify'){assert.equal(selected.fields[i].selected,text,'double click '+text);assert.equal(selected.fields[i].rects.length,1,'single glyph rectangle '+text);}
      }
    }
    await seed(String.raw`\mathbf{a+b}+\frac{\mathbf{c}}{\mathbf{d}}+\mathbf{e}`);
    await clickGlyph('c',0,.5,2);await snapshot('double-styled-numerator');
  } else if(scenario==='many-lines'){
    await seed(...Array.from({length:20},(_,i)=>i%2?formulas[1]:formulas[0]));await visual(9);
    const original=(await state()).fields.map(f=>f.value);await page.keyboard.press('Control+a');
    const selected=await snapshot('twenty-lines-select-all');assert.equal(selected.fields.length,20);assert.ok(selected.fields.every(f=>f.multi&&f.selection.ranges[0][0]===0&&f.selection.ranges[0][1]===f.lastOffset));
    await page.keyboard.type('z');assert.deepEqual((await record('twenty-lines-replaced')).fields.map(f=>f.value),['z']);
    await page.keyboard.press('Control+z');assert.deepEqual((await record('twenty-lines-undo')).fields.map(f=>f.value),original);
  } else if(scenario==='zoom'){
    await seed(String.raw`\frac{a+\hat{b}}{\sqrt{c_d}}+e`);
    for(const zoom of [40,100,160]){
      await setZoom(zoom);await visual();
      for(const text of ['a','b','c','d','e']){
        const target=await clickGlyph(text,0,.8);assert.equal((await record(`zoom-${zoom}-point-${text}`)).fields[0].position,target.offset);
        await clickGlyph(text,0,.2,2);assert.equal((await record(`zoom-${zoom}-double-${text}`)).fields[0].selected,text);
      }
      await page.keyboard.press('Control+a');const s=await snapshot('zoom-'+zoom+'-all');assert.equal(s.fields[0].rects.length,1);
    }
  } else if(scenario==='structure'){
    const drag=async(a,b,index=0,reverse=false)=>{
      await page.waitForTimeout(550);const g=await glyph(a,index),h=await glyph(b,index);
      const p={x:g.rect.left+g.rect.width*.1,y:g.rect.top+g.rect.height*.5},q={x:h.rect.left+h.rect.width*.9,y:h.rect.top+h.rect.height*.5};
      const from=reverse?q:p,to=reverse?p:q;
      await page.mouse.move(from.x,from.y);await page.mouse.down();await page.mouse.move(to.x,to.y,{steps:20});await page.mouse.up();await settle();
    };
    await seed(...formulas);
    for(const reverse of [false,true]){
      await drag('a','c',0,reverse);const numerator=await snapshot('drag-numerator-'+reverse);assert.equal(numerator.fields[0].selected,String.raw`a+\sqrt{b_i^2+c}`);assert.equal(numerator.fields[0].rects.length,1);
      await drag('a','f',0,reverse);const fraction=await snapshot('drag-fraction-'+reverse);assert.equal(fraction.fields[0].selected,String.raw`\frac{a+\sqrt{b_i^2+c}}{\hat{d}+e_f}`);assert.equal(fraction.fields[0].rects.length,1);
    }
    await drag('h','k',1);await snapshot('drag-matrix-row');
    await clickGlyph('a',0,.1);await page.keyboard.down('Shift');await clickGlyph('c',0,.9);await page.keyboard.up('Shift');assert.equal((await record('shift-click-numerator')).fields[0].selected,String.raw`a+\sqrt{b_i^2+c}`);
    await seed(String.raw`\frac{123}{456}+\mathbf{ab}+\hat{\vec{r_s}}+t`);
    await clickGlyph('2',0,.2,2);assert.equal((await record('double-number')).fields[0].selected,'123');
    await clickGlyph('a',0,.2,2);assert.equal((await record('double-bold-glyph')).fields[0].selected,String.raw`\mathbf{a}`);
    await clickGlyph('r',0,.2,2);assert.equal((await snapshot('double-nested-accent')).fields[0].selected,'r');
    const compact=s=>s.replace(/([_^])\{([A-Za-z0-9])\}/g,'$1$2');
    await page.keyboard.type('x');const changed=await record('replace-nested-accent-body');assert.ok(compact(changed.fields[0].value).includes(String.raw`\hat{\vec{x_s}}`));await page.keyboard.press('Control+z');assert.ok(compact((await record('undo-nested-accent-body')).fields[0].value).includes(String.raw`\hat{\vec{r_s}}`));
  } else if(scenario==='document'){
    const values=s=>s.fields.map(f=>f.value);
    const original=async()=>{await seed(...formulas);return values(await state());};
    const undo=async expected=>{await page.keyboard.press('Control+z');assert.deepEqual(values(await record('undo')),expected);};
    for(const index of [0,1]){
      const expected=await original();await clickGlyph(index?'q':'g',index);await page.keyboard.press('Control+a');
      await page.keyboard.type('x');assert.deepEqual(values(await record('replace-all-from-'+index)),['x']);
      await undo(expected);await page.keyboard.press('Control+y');assert.deepEqual(values(await record('redo-replacement')),['x']);
    }
    for(const key of ['Backspace','Delete']){
      const expected=await original();await clickGlyph('q',1);await page.keyboard.press('Control+a');await page.keyboard.press(key);
      assert.deepEqual(values(await record('delete-all-'+key)),['']);await undo(expected);
    }
    for(const [key,index,atEnd] of [['ArrowLeft',0,false],['ArrowRight',1,true],['Home',0,false],['End',1,true],['Escape',1,true]]){
      await original();await clickGlyph('g');await page.keyboard.press('Control+a');await page.keyboard.press(key);
      const s=await record('cancel-'+key);assert.ok(s.fields.every(f=>!f.multi&&!f.selected));assert.ok(s.fields[index].focused);assert.equal(s.fields[index].position,atEnd?s.fields[index].lastOffset:0);
    }
    const expected=await original();await clickGlyph('g');await page.keyboard.press('Control+a');
    await page.locator('math-field').first().evaluate(f=>{const data=new DataTransfer();f.dispatchEvent(new ClipboardEvent('copy',{bubbles:true,composed:true,cancelable:true,clipboardData:data}));window.__selectionClipboard=[...data.types].map(type=>[type,data.getData(type)]);});
    const clipboard=await page.evaluate(()=>window.__selectionClipboard);assert.ok(clipboard.find(([type,data])=>type==='text/plain'&&data.includes('pmatrix')&&data.includes('frac')));
    await record('copy-document');
    await page.locator('math-field').first().evaluate(f=>{const data=new DataTransfer();f.dispatchEvent(new ClipboardEvent('cut',{bubbles:true,composed:true,cancelable:true,clipboardData:data}));});
    assert.deepEqual(values(await record('cut-document')),['']);await undo(expected);
    await clickGlyph('g');await page.keyboard.press('Control+a');
    await page.locator('math-field').first().evaluate(f=>{const data=new DataTransfer();for(const [type,value] of window.__selectionClipboard)data.setData(type,value);f.dispatchEvent(new ClipboardEvent('paste',{bubbles:true,composed:true,cancelable:true,clipboardData:data}));});
    const canonical=s=>s.replace(/\s+/g,'').replace(/([_^])\{([A-Za-z0-9])\}/g,'$1$2');
    assert.deepEqual(values(await record('paste-over-document')).map(canonical),expected.map(canonical));
    await original();await visual(1);await page.keyboard.press('Control+End');await page.keyboard.press('Enter');await settle();
    assert.equal((await state()).fields.length,3);await page.keyboard.press('Control+a');await page.keyboard.type('z');assert.deepEqual(values(await record('replace-including-empty-row')),['z']);
    await seed('a+b','c+d');await clickGlyph('b');await page.keyboard.press('Control+a');await page.locator('[data-formula-selection-bold]').click();
    const styled=await record('bold-whole-document');assert.equal(styled.fields.length,2);assert.ok(styled.fields.every(f=>f.value.includes('mathbf')));await undo(['a+b','c+d']);
    for(const key of ['Backspace','x']){
      await seed('a+b','c+d');await clickGlyph('b');await page.keyboard.press('Control+a');await page.keyboard.press('Shift+ArrowLeft');
      await record('shrink-document-selection-'+key);await page.keyboard.press(key);assert.deepEqual(values(await record('replace-shrunken-document-'+key)),[key==='x'?'xd':'d']);await undo(['a+b','c+d']);
    }
    await sourcePanel();await page.locator('.cm-content').click();await page.keyboard.press('Control+a');await record('source-select-all-is-local');assert.deepEqual(values(await state()),['a+b','c+d']);
  } else throw Error('Unknown selection scenario '+scenario);
}catch(e){failures.push({message:e.message,stack:e.stack});await snapshot('failure').catch(()=>{});}
finally{
  try {
    if(before.zoom&&(await state()).zoom!==before.zoom)await setZoom(before.zoom);
    await setSource(before.source);
    // The source serializer omits empty visual rows. Restore these through Enter.
    let restored=await state();
    while(restored.fields.length<before.fields.length && before.fields[restored.fields.length].value===''){
      await visual();await page.keyboard.press('Control+a');await page.keyboard.press('ArrowRight');await settle();await page.keyboard.press('Enter');await settle();restored=await state();
    }
    assert.deepEqual(restored.fields.map(f=>f.value),before.fields.map(f=>f.value));
    await fs.writeFile(path.join(output,'restored.json'),JSON.stringify(restored,null,2));
  }catch(e){failures.push({phase:'restore',message:e.message,stack:e.stack});}
  await fs.writeFile(path.join(output,'result.json'),JSON.stringify({steps,failures,errors},null,2));await browser.close();
}
assert.deepEqual(failures,[]);assert.deepEqual(errors,[]);console.log(`PASS ${scenario}: ${steps.length} observations`);
