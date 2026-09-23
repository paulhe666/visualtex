// Core contract tests in an isolated MathField in the installed WebView.
// Programmatic setup is intentional here; user workflows are tested separately.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
const {chromium}=await import(process.env.VISUALTEX_PLAYWRIGHT_MODULE||'playwright');
const [port='9227',output]=process.argv.slice(2);
assert.ok(output);
await fs.mkdir(output,{recursive:true});
const browser=await chromium.connectOverCDP(`http://127.0.0.1:${Number(port)}`);
try {
  const page=browser.contexts()[0].pages().find(p=>p.url()==='http://tauri.localhost/');
  const results=await page.evaluate(async()=>{
    const before=[...document.querySelectorAll('math-field')].map(f=>f.value);
    const field=document.createElement('math-field');
    field.style.cssText='position:fixed;bottom:0;right:0;z-index:100000;background:white';
    document.body.append(field);
    await new Promise(resolve=>requestAnimationFrame(resolve));
    const results=[];
    const check=(name,ok,detail)=>{results.push({name,ok,detail});if(!ok)throw Error(name+': '+JSON.stringify(detail));};
    const setup=latex=>{
      field.setValue(latex);
      field.resetUndo();
      return field._mathfield.model;
    };
    const guarded=(command,...args)=>{
      const setValue=field.setValue,insert=field.insert;
      field.setValue=field.insert=()=>{throw Error('Public LaTeX reparse called by core command');};
      try {return field.executeCommand(command,...args);}
      finally {field.setValue=setValue;field.insert=insert;}
    };
    const compact=s=>s.replace(/\s+/g,'');
    try {
      for(const latex of [String.raw`\hat{\Delta}`,String.raw`\vec{\hat{\Delta}}`,String.raw`\frac{\hat{\Delta}+q}{\sqrt{x_i^2+1}}`,String.raw`\mathbb{R}+\mathfrak{g}`,String.raw`\mathrm{ab}+x`,String.raw`\mathbf{\hat{\Delta}+q}`]) {
        const model=setup(latex), original=field.value;
        field.selection={ranges:[[0,field.lastOffset]],direction:'forward'};
        const root=model.root, atoms=[...root.children], selection=JSON.stringify(field.selection);
        guarded('visualTexToggleSelectionStyle','bold');
        const bold=field.value;
        check('style preserves atoms: '+latex,model.root===root&&atoms.every(a=>root.children.includes(a)),bold);
        check('style preserves selection: '+latex,JSON.stringify(field.selection)===selection,field.selection);
        guarded('visualTexToggleSelectionStyle','italic');
        guarded('visualTexToggleSelectionStyle','italic');
        if(latex===String.raw`\hat{\Delta}`||latex===String.raw`\vec{\hat{\Delta}}`)
          check('italic double-toggle restores bold: '+latex,compact(field.value)===compact(bold),{bold,actual:field.value});
        guarded('visualTexToggleSelectionStyle','bold');
        // Mixed-style selections normalize on toggle; identity/inverse only for uniform glyph styles.
        check('no tree growth: '+latex,root.children.length===atoms.length&&atoms.every(a=>root.children.includes(a)),{original,actual:field.value});
      }
      for(const item of [
        {latex:'ab+cd',split:'b',env:'gathered',expected:String.raw`\begin{gathered}ab\\+cd\end{gathered}`},
        {latex:String.raw`\frac{ab}{\hat{c}}`,split:'a',env:'gathered',expected:String.raw`\frac{\begin{gathered}a\\b\end{gathered}}{\hat{c}}`},
        {latex:String.raw`\sqrt{a+b}`,split:'a',env:'aligned',expected:String.raw`\sqrt{\begin{aligned}a&\\+b&\end{aligned}}`},
        {latex:String.raw`\begin{gathered}a\\b\end{gathered}`,split:'end',env:'gathered',expected:String.raw`\begin{gathered}a\\b\\\end{gathered}`},
        {latex:String.raw`\begin{gathered}a\\b\end{gathered}`,split:'start',env:'gathered',expected:String.raw`\begin{gathered}\\a\\b\end{gathered}`},
      ]) {
        const model=setup(item.latex),root=model.root,atoms=root.children.filter(a=>a.type!=='first'),original=field.value;
        field.position=item.split==='end'?field.lastOffset:item.split==='start'?0:model.offsetOf(atoms.find(a=>a.value===item.split));
        guarded('visualTexInsertRowBreak',item.env);
        check('row split semantics: '+item.latex+' '+item.split,compact(field.value)===compact(item.expected),field.value);
        check('row split preserves atoms: '+item.latex,model.root===root&&atoms.every(a=>root.children.includes(a)),field.value);
        field.executeCommand('undo');
        check('native undo row break: '+item.latex,compact(field.value)===compact(original),field.value);
        field.executeCommand('redo');
        check('native redo row break: '+item.latex,compact(field.value)===compact(item.expected),field.value);
      }
      for(const latex of [String.raw`\mathrm{ab}+x`,String.raw`\hat{\Delta}+x+1`,String.raw`\mathbb{R}+\mathfrak{g}`]) {
        const model=setup(latex);
        field.selection={ranges:[[0,field.lastOffset]]};
        guarded('visualTexToggleSelectionStyle','bold');
        const styled=field.value;
        const glyphStyles=()=>model.root.children.filter(a=>a.value&&a.type!=='first').map(a=>({value:a.value,variant:a.style.variant||'normal',variantStyle:a.style.variantStyle}));
        const before=glyphStyles();
        field.setValue('');
        field.setValue(styled);
        check('style survives serialization: '+latex,JSON.stringify(before)===JSON.stringify(glyphStyles()),{styled,before,after:glyphStyles()});
      }
      {
        const model=setup(String.raw`\frac{\hat{\Delta}+q}{\sqrt{x_i^2+1}}+z`);
        const original=field.value;
        const delta=model.root.children.find(a=>a.value==='Δ');
        const z=model.root.children.find(a=>a.value==='z');
        const selected=new Set([delta,z]);
        const untouched=model.root.children.filter(a=>a.type!=='first'&&!selected.has(a));
        const untouchedStyle=untouched.map(a=>JSON.stringify(a.style));
        const requestedSelection={ranges:[delta,z].map(a=>[model.offsetOf(a)-1,model.offsetOf(a)]),direction:'backward'};
        field.selection=requestedSelection;
        const selection=JSON.stringify(field.selection);
        guarded('visualTexToggleSelectionStyle','bold',requestedSelection);
        const styled=field.value;
        check('disjoint selection preserves untouched atoms',untouched.every((a,i)=>model.root.children.includes(a)&&JSON.stringify(a.style)===untouchedStyle[i]),styled);
        check('explicit ranges do not move current selection',JSON.stringify(field.selection)===selection,field.selection);
        field.executeCommand('undo');
        check('native undo selected font',field.value===original,field.value);
        field.executeCommand('redo');
        check('native redo selected font',field.value===styled,field.value);
      }
      {
        setup('ab');field.position=1;
        let beforeCount=0;const inputTypes=[];
        const before=()=>beforeCount++;const after=event=>inputTypes.push(event.inputType);
        // Drain setup notifications; inspect only the command transaction.
        await new Promise(resolve=>setTimeout(resolve,20));
        field.addEventListener('beforeinput',before);field.addEventListener('input',after);
        guarded('visualTexInsertRowBreak','gathered');
        await new Promise(resolve=>setTimeout(resolve,20));
        field.removeEventListener('beforeinput',before);field.removeEventListener('input',after);
        check('one row-break input transaction',beforeCount===1&&JSON.stringify(inputTypes)===JSON.stringify(['insertLineBreak']),{beforeCount,inputTypes});
      }
      {
        const model=setup(String.raw`\begin{aligned}a & \\ b & \end{aligned}`);
        const array=model.root.body?.find(atom=>atom.type==='array');
        const b=array?.getCell?.(1,0)?.find(atom=>atom.value==='b');
        check('aligned array fixture available',Boolean(array&&b),field.value);
        const original=field.value;
        model.position=model.offsetOf(array.getCell(1,0)[0]);
        guarded('visualTexInsertAlignmentPoint');
        const alignedValue=field.value;
        const branch=model.at(model.position)?.parentBranch;
        check('alignment point becomes array column',Array.isArray(branch)&&branch[0]===1&&branch[1]===1,{branch,alignedValue});
        check('alignment point preserves existing atoms',array.getCell(1,1).includes(b),alignedValue);
        field.executeCommand('undo');
        check('native undo alignment point',compact(field.value)===compact(original),field.value);
        field.executeCommand('redo');
        check('native redo alignment point',compact(field.value)===compact(alignedValue),field.value);
      }
      {
        setup(String.raw`\vec{x}+\dddot{q}+\ddddot{r}+\hat{\placeholder{}}`);
        await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
        const shadow=field.shadowRoot;
        check('native renderer emits SVG without application observer',shadow.querySelectorAll('.ML__visualtex-accent-ink svg').length===3,shadow.querySelectorAll('.ML__visualtex-accent-ink').length);
        check('native renderer owns accent placeholder markers',shadow.querySelectorAll('.visualtex-accent-placeholder-layout').length===1&&shadow.querySelectorAll('.visualtex-accent-structural-placeholder').length===1,shadow.querySelectorAll('.visualtex-accent-placeholder-layout').length);
        check('no overlay or pixel shift in native renderer',shadow.querySelectorAll('.visualtex-combining-accent-overlay,[style*="--visualtex-accent-placeholder-shift"]').length===0,field.value);
      }
      for(const latex of [String.raw`\vec{}\vec{}`,String.raw`\hat{}+\frac{}{}+\sqrt{}`,String.raw`\begin{gathered}\vec{}\\\hat{}\end{gathered}`]) {
        setup(latex);
        const original=field.value;
        for(let i=0;i<12;i++)field.macros={...field.macros};
        check('macro reparsing is idempotent: '+latex,field.value===original,field.value);
        field.selection={ranges:[[0,field.lastOffset]]};
        check('empty structures included in selection: '+latex,compact(field.getValue(field.selection,'latex'))===compact(original),field.getValue(field.selection,'latex'));
        field.executeCommand('deleteBackward');
        check('native selection deletion removes empty structures: '+latex,field.value==='',field.value);
      }
      for(const command of ['visualTexInsertRowBreak','visualTexInsertAlignmentPoint','visualTexToggleSelectionStyle']) {
        const alignmentCommand=command==='visualTexInsertAlignmentPoint';
        setup(alignmentCommand?String.raw`\begin{aligned}a & \\ b & \end{aligned}`:'ab');
        if(alignmentCommand) {
          const model=field._mathfield.model;
          const array=model.root.body?.find(atom=>atom.type==='array');
          model.position=model.offsetOf(array.getCell(1,0)[0]);
        } else field.selection={ranges:[[0,field.lastOffset]]};
        const original=field.value;
        const cancel=event=>event.preventDefault();field.addEventListener('beforeinput',cancel);
        if(alignmentCommand) guarded(command);
        else guarded(command,command.includes('Style')?'bold':'gathered');
        field.removeEventListener('beforeinput',cancel);
        check('cancellable beforeinput '+command,field.value===original,field.value);
        field.readOnly=true;
        if(alignmentCommand) guarded(command);
        else guarded(command,command.includes('Style')?'bold':'gathered');
        field.readOnly=false;
        check('readOnly '+command,field.value===original,field.value);
      }
    } catch(error) {results.push({name:'exception',ok:false,detail:error.stack});}
    finally {field.remove();}
    check('user document unchanged',JSON.stringify(before)===JSON.stringify([...document.querySelectorAll('math-field')].map(f=>f.value)),before.length);
    return results;
  });
  await fs.writeFile(path.join(output,'result.json'),JSON.stringify(results,null,2));
  for(const r of results)console.log(JSON.stringify(r));
  assert.equal(results.filter(r=>!r.ok).length,0);
} finally {await browser.close();}
