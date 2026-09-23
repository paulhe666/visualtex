// Prepare/verify/restore a real installed-app restart. The caller closes and
// starts VisualTeX between prepare and verify; this script does not simulate it.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
const {chromium}=await import(process.env.VISUALTEX_PLAYWRIGHT_MODULE||'playwright');
const [port,output,mode]=process.argv.slice(2);
assert.ok(port&&output&&['prepare','verify','restore'].includes(mode));
await fs.mkdir(output,{recursive:true});
const browser=await chromium.connectOverCDP(`http://127.0.0.1:${Number(port)}`);
const page=browser.contexts()[0].pages().find(p=>p.url()==='http://tauri.localhost/');
assert.ok(page,'Expected installed VisualTeX');
const errors=[];page.on('pageerror',e=>errors.push(e.message));
const state=()=>page.evaluate(()=>({
  source:document.querySelector('.cm-content')?.innerText,
  title:document.querySelector('input[aria-label="公式文档标题"]')?.value,
  zoom:document.querySelector('.canvas-controls [aria-live]')?.textContent,
  scripts:[...document.scripts].map(s=>s.src),
  fields:[...document.querySelectorAll('math-field')].map(f=>({value:f.value,readOnly:f.readOnly,focused:f.hasFocus(),position:f.position})),
}));
const write=(name,value)=>fs.writeFile(path.join(output,name),JSON.stringify(value,null,2));
const read=async name=>JSON.parse(await fs.readFile(path.join(output,name),'utf8'));
const visual=async()=>{await page.locator('.mathfield-host').first().click({position:{x:5,y:20}});await page.waitForTimeout(300);};
const source=async text=>{
  await page.locator('[data-classic-bottom-view="source"]').click();
  await page.locator('.cm-content').click();await page.keyboard.press('Control+a');await page.keyboard.insertText(text);
  await page.waitForTimeout(250);await visual();await page.waitForTimeout(750);
};
const values=s=>s.fields.map(f=>f.value);
try {
  if(mode==='prepare'){
    await page.locator('[data-classic-bottom-view="source"]').click();
    const before=await state();
    // Never replace a prior backup with a temporary test document.
    await fs.writeFile(path.join(output,'before.json'),JSON.stringify(before,null,2),{flag:'wx'});
    const formulas=[
      String.raw`\vec{}\vec{}+\frac{\vec{}}{\hat{}}`,
      String.raw`\begin{gathered}\frac{\sqrt{\vec{v}}}{\sqrt[3]{x}}\\\left(\frac{c}{\hat{\mathbf{\Delta}}}\right)\end{gathered}`,
      String.raw`\begin{pmatrix}\frac{\dddot{q}}{2}&\sqrt{c}\\\widehat{s}&t_i\end{pmatrix}^{n}`,
    ];
    await source(formulas.map(v=>`$$\n${v}\n$$`).join('\n\n'));
    const prepared=await state();assert.deepEqual(values(prepared),formulas);await write('prepared.json',prepared);
  }else if(mode==='verify'){
    const prepared=await read('prepared.json');
    await page.locator('[data-classic-bottom-view="source"]').click();await visual();
    const before=await state();assert.deepEqual(values(before),values(prepared));assert.equal(before.zoom,prepared.zoom);assert.ok(before.fields.every(f=>!f.readOnly));
    const box=await page.locator('.mathfield-host').first().boundingBox();await page.mouse.click(box.x+box.width-4,box.y+box.height/2);
    await page.keyboard.type('w');await page.waitForTimeout(180);
    const typed=await state();assert.equal(typed.fields[0].value,prepared.fields[0].value+'w');
    await page.keyboard.press('Control+z');await page.waitForTimeout(200);
    const after=await state();assert.deepEqual(values(after),values(prepared));
    await write('verified.json',{before,typed,after,errors});
  }else{
    const before=await read('before.json');await source(before.source);
    const after=await state();assert.deepEqual(values(after),values(before));assert.equal(after.title,before.title);assert.equal(after.zoom,before.zoom);
    await write('restored.json',{same:true,state:after,errors});
  }
  assert.deepEqual(errors,[]);await page.screenshot({path:path.join(output,`${mode}.png`)});
  console.log(`PASS installed process restart ${mode}`);
}finally{await browser.close();}
