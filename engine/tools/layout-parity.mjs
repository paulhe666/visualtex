// Real native/WASM geometry comparison using the same injected font bytes.
// Fonts are read-only local fixtures. No source font or production asset is copied.
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { once } from 'node:events';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const root=resolve(dirname(fileURLToPath(import.meta.url)), '..');
const paths={
  math:process.env.VISUALTEX_TEST_MATH_FONT ?? resolve(root,'../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf'),
  text:process.env.VISUALTEX_TEST_TEXT_FONT ?? 'C:\\Windows\\Fonts\\NotoSansSC-VF.ttf',
};
const expected={
  math:'6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f',
  text:'763146584cf0710223441356b4395e279021b0806c196614377a7a0174ae074a',
};
const sha=bytes=>createHash('sha256').update(bytes).digest('hex');
const bytes=Object.fromEntries(Object.entries(paths).map(([kind,path])=>[kind,readFileSync(path)]));
for (const kind of Object.keys(bytes)) assert.equal(sha(bytes[kind]),expected[kind],`Fixed ${kind} font fixture differs; do not refresh a golden hash to conceal this.`);
const nativePath=resolve(root,`target/debug/visualtex-engine-test-cli${process.platform==='win32'?'.exe':''}`);
const wasmPath=resolve(root,'web/pkg-node/visualtex_engine_wasm_bg.wasm');
const {EngineRuntime}=createRequire(import.meta.url)(resolve(root,'web/pkg-node/visualtex_engine_wasm.js'));
const wasm=new EngineRuntime();
assert.equal(JSON.parse(wasm.registerFont(new Uint8Array([0,1,2]),0)).status,'error');
const fontIds={};const registrations={};
for (const kind of Object.keys(bytes)) {
  const response=JSON.parse(wasm.registerFont(bytes[kind],0));assert.equal(response.status,'ok',JSON.stringify(response));
  assert.equal(response.result.kind,'fontRegistered');
  const font=response.result.font;assert.equal(font.sha256,expected[kind]);assert.equal(font.faceIndex,0);
  fontIds[kind]=font.id;registrations[kind]=font;
  assert.deepEqual(JSON.parse(wasm.registerFont(bytes[kind],0)),response,'Identical bytes must deduplicate');
}
const native=spawn(nativePath,['--font',paths.math,'--font',paths.text],{stdio:['pipe','pipe','pipe']});
const lines=createInterface({input:native.stdout,crlfDelay:Infinity})[Symbol.asyncIterator]();
let stderr='';native.stderr.on('data',data=>{stderr=(stderr+data.toString()).slice(-12000);});
const timer=setTimeout(()=>native.kill(),90000);
const transcript=createHash('sha256');let cases=0,scenes=0,maxNumericDifference=0;
function equivalent(a,b,path='response') {
  if(typeof a==='number' && typeof b==='number') {
    assert.ok(Number.isFinite(a) && Number.isFinite(b),`nonfinite ${path}`);
    const difference=Math.abs(a-b);maxNumericDifference=Math.max(maxNumericDifference,difference);
    assert.ok(difference<=1e-9,`${path}: ${a} != ${b}`);return;
  }
  if(a===null || b===null || typeof a!=='object' || typeof b!=='object') {assert.equal(a,b,path);return;}
  assert.equal(Array.isArray(a),Array.isArray(b),path);assert.deepEqual(Object.keys(a),Object.keys(b),path);
  for(const key of Object.keys(a)) equivalent(a[key],b[key],`${path}.${key}`);
}
async function send(command,error=null) {
  const input=JSON.stringify({protocol:1,command});const a=JSON.parse(wasm.dispatch(input));
  if(!native.stdin.write(input+'\n')) await once(native.stdin,'drain');
  const next=await lines.next();assert.equal(next.done,false,stderr);const b=JSON.parse(next.value);
  equivalent(a,b,`case${cases}:${command.op}`);transcript.update(input+'\n'+next.value+'\n');cases++;
  if(error) {assert.equal(a.status,'error');assert.equal(a.error.code,error);return a;}
  assert.equal(a.status,'ok',JSON.stringify(a));return a.result;
}
const defaults={mathFont:fontIds.math,textFont:fontIds.text,style:'display',fontSizePx:16,paddingEm:0.125};
const requestLayout=(handle,revision=0,options={},includeSvg=true,error=null)=>send({op:'layout',handle,expectedRevision:revision,options:{...defaults,...options},includeSvg},error);
function validate(scene,svg) {
  scenes++;assert.ok(scene.width>0&&scene.height>0&&scene.baseline>=0&&scene.baseline<=scene.height);
  assert.equal(scene.width,scene.widthEm*scene.fontSizePx);assert.equal(scene.height,scene.heightEm*scene.fontSizePx);
  assert.ok(scene.carets.length>0);assert.equal(new Set(scene.regions.map(r=>r.nodeId)).size,scene.regions.length);
  for(const g of scene.glyphs) {const resource=scene.resources[g.resource];assert.ok(resource);assert.ok(resource.shape.glyphId>0);
    if(resource.shape.ink) {const [a,b,c,d]=resource.shape.ink;
      assert.ok(a*g.scale+g.x>=-1e-9&&b*g.scale+g.y>=-1e-9);
      assert.ok(c*g.scale+g.x<=scene.widthEm+1e-9&&d*g.scale+g.y<=scene.heightEm+1e-9);
    }
  }
  if(svg!==null) {assert.ok(scene.exportReady);assert.ok(svg.startsWith('<svg ')&&svg.endsWith('</svg>'));
    for(const forbidden of ['<text','<image','<script','foreignObject','href=','NaN','Infinity']) assert.ok(!svg.includes(forbidden),forbidden);
  }
}
const sources=[
  ...['',String.raw`\limits`,String.raw`\nolimits`,String.raw`\displaylimits`].flatMap(mode=>[
    `\\mathop{f(x)}${mode}_a^b`, `\\mathop{\\binom{n}{k}}${mode}_i^n`,
  ]),
  String.raw`\mathop{\sum\sum}_i^n`,String.raw`\mathop{\text{算子}}_a^b`,
  String.raw`\textcolor{blue}{\mathop{f}_a^b}`,String.raw`\phantom{\mathop{f}_a^b}+x`,
  ...['overgroup','undergroup','overleftharpoon','overrightharpoon','overlinesegment','underlinesegment'].flatMap(c=>[
    `\\${c}{x+y}`,`\\${c}{${'x+'.repeat(24)}x}`,`\\textcolor{red}{\\${c}{\\binom{n}{k}}}`,
  ]),
  String.raw`a\color{red}b{\color{blue}c}d\textcolor{green}{e}f`,
  String.raw`x\color{red}+y`,String.raw`x\textcolor{blue}{+}y`,
  String.raw`\textcolor{red}{\boxed{\frac{\sqrt{x}}{\overline{y}}}}\phantom{\textcolor{blue}{\boxed{z}}}`,
  String.raw`\sqrt[{\textcolor[HTML]{FF0000}{n}}]{x}`,
  ...['red','green','blue','cyan','magenta','yellow','black','white','gray','darkgray','lightgray','brown','lime','olive','orange','pink','purple','teal','violet','red!50!blue','red!25!blue!30']
    .map(c=>`\\textcolor{${c}}{\\frac{x_1}{y^2}}`),
  String.raw`\textcolor[RGB]{1,128,255}{x}+\textcolor[rgb]{1,.5,0}{y}+\textcolor[gray]{.25}{z}`,
  ...[',',':',';','!','thinspace','medspace','thickspace','negthinspace','negmedspace','negthickspace','enspace','enskip','quad','qquad']
    .map(command=>`x\\${command} +y`),
  ...['phantom','hphantom','vphantom'].flatMap(command=>[
    `\\${command}{\\frac{x}{y}}`,`x+\\${command}{\\boxed{\\sqrt{\\frac{a}{b}}}}+y`]),
  String.raw`\boxed{\frac{x}{y}}`,String.raw`\boxed{\phantom{\mathbf{x}}+\quad y}`,
  String.raw`\!\!`,String.raw`x\!\!\!\!`,String.raw`\vphantom{\frac{x}{y}}`,
  String.raw`\hphantom{\frac{x}{y}}`,String.raw`\boxed{\!}`,
  'x+1',String.raw`\frac{x}{y}`,String.raw`\frac{x_1+\text{速度}}{y^2}`,
  String.raw`\text{中文 Café office AV}`,String.raw`\frac{\frac{a}{b}}{x^{y^2}}`,
  String.raw`x^{\dfrac{a}{b}}+\frac{\tfrac{c}{d}}{y}`,
  String.raw`x_1^2+\alpha=\beta`,String.raw`\text{<svg \& x>}`,
  String.raw`\sqrt{x}`,String.raw`\sqrt[n+1]{x^2}`,
  String.raw`\sqrt[{\sqrt[3]{n}}]{\frac{x_1+\text{速度}}{y^2}}`,
  String.raw`\sqrt{\dfrac{1}{\dfrac{2}{\dfrac{3}{4}}}}`,
  String.raw`\left(\frac{x}{y}\right)`,String.raw`\left.\sqrt[n]{x}\right]`,
  String.raw`\left\{\dfrac{x}{\dfrac{y}{z}}\right.`,
  String.raw`\left[\left|\sqrt[3]{x^2}\right|+\text{速度}\right]`,
  ...['hat','widehat','tilde','widetilde','vec','bar','dot','ddot','dddot','ddddot','breve','check','acute','grave','mathring',
      'overline','underline','overrightarrow','overleftarrow','overleftrightarrow','underrightarrow','underleftarrow',
      'underleftrightarrow','overbrace','underbrace','overparen','underparen','overarc','underarc'].map(c=>`\\${c}{x+y}`),
  ...['overset','underset','stackrel','stackbin'].map(c=>`\\${c}{a+1}{x}`),
  String.raw`\hat{\vec{x}}`,String.raw`\overbrace{x+y}^{n}`,String.raw`\underbrace{x+y}_{n}`,
  String.raw`\overset{a}{\underset{b}{\hat{x}}}`,String.raw`\overline{\text{速度}}`,
  `\\overbrace{${'x+'.repeat(40)}y}`,`\\overrightarrow{${'x+'.repeat(40)}y}`,
  ...['sum','prod','coprod','bigcup','bigcap','bigsqcup','biguplus','bigvee','bigwedge','bigodot','bigoplus','bigotimes',
      'int','iint','iiint','oint','oiint','oiiint','lim','limsup','liminf','min','max','inf','sup'].flatMap(c=>
      ['',String.raw`\limits`,String.raw`\nolimits`,String.raw`\displaylimits`].map(l=>`\\${c}${l}_{i=1}^{n}x`)),
  String.raw`\sum^{n}_{i=1}\limits\frac{x_i}{\sum_j^n y_j}`,String.raw`\hat{\sum_i^n}`,
  String.raw`\int\limits_{\text{下限}}^{\text{上限}}x`,String.raw`\sqrt[{\sum_]}]{x}`,
  String.raw`\mathbf{A+x}`,String.raw`\mathrm{d}x+\mathit{ABC}`,String.raw`\mathbb{R}+\mathcal{A}+\mathfrak{H}`,
  String.raw`\boldsymbol{\alpha}+\bm{\Gamma}+\mathsf{ABC}+\mathtt{x1}`,
  String.raw`\sin x+\cos y+\log z+\exp t`,String.raw`\operatorname{tr}(A)+\operatorname{rank}(B)`,String.raw`\mathbf{\frac{x_1}{\mathbb{R}}}+\mathrm{\sin}`,
];
try {
  for(const [i,source] of sources.entries()) {
    const created=await send({op:'create',documentId:`layout-parity-${i}`,source});const h=created.handle;
    for(const style of ['display','text','script','scriptScript']) {
      const first=await requestLayout(h,0,{style});validate(first.snapshot,first.svg);
      const again=await requestLayout(h,0,{style});validate(again.snapshot,again.svg);
      assert.equal(again.snapshot.rebuiltFragments,0);assert.equal(again.svg,first.svg);
      const scaled=await requestLayout(h,0,{style,fontSizePx:32});validate(scaled.snapshot,scaled.svg);
      assert.equal(scaled.snapshot.width,first.snapshot.width*2);assert.equal(scaled.snapshot.baseline,first.snapshot.baseline*2);
      assert.deepEqual(scaled.snapshot.glyphs,first.snapshot.glyphs);assert.equal(scaled.snapshot.rebuiltFragments,0);
    }
    await requestLayout(h,99,{},true,'staleRevision');
    await requestLayout(h,0,{mathFont:'unregistered'},true,'invalidInput');
    await send({op:'destroy',handle:h,expectedRevision:0});
    await requestLayout(h,0,{},true,'invalidHandle');
  }
  const source=String.raw`\frac{x_1+\text{速度}}{y^2}`;
  const h=(await send({op:'create',documentId:'layout-edit-parity',source})).handle;
  const initial=await requestLayout(h);validate(initial.snapshot,initial.svg);
  const at=source.indexOf('1');
  await send({op:'patchSource',handle:h,expectedRevision:0,edit:{startUtf16:at,endUtf16:at+1,replacement:'3'}});
  const edited=await requestLayout(h,1);validate(edited.snapshot,edited.svg);
  assert.notEqual(edited.svg,initial.svg);assert.ok(edited.snapshot.reusedFragments>0);
  assert.ok(edited.snapshot.rebuiltFragments<initial.snapshot.rebuiltFragments);
  await send({op:'editorAction',handle:h,expectedRevision:1,eventSequence:1,action:{kind:'undo'}});
  const restored=await requestLayout(h,2);validate(restored.snapshot,restored.svg);assert.equal(restored.svg,initial.svg);
  const editor=(await send({op:'inspectEditor',handle:h,expectedRevision:2})).snapshot;
  assert.equal(editor.document.source,source);
  const chinese=editor.document.analysis.tree.entries[0].node.numerator.entries.find(e=>e.node.kind==='text').node;
  const caret={kind:'text',nodeId:chinese.id,grapheme:2,affinity:'after'};
  await send({op:'editorAction',handle:h,expectedRevision:2,eventSequence:2,action:{kind:'setSelection',selection:{kind:'range',anchor:caret,focus:caret}}});
  const begin=await send({op:'editorAction',handle:h,expectedRevision:2,eventSequence:3,action:{kind:'beginComposition',mode:'text'}});
  const id=begin.result.editor.composition.id;
  await send({op:'editorAction',handle:h,expectedRevision:2,eventSequence:4,action:{kind:'updateComposition',id,preedit:'率'}});
  const preview=await requestLayout(h,2);assert.equal(preview.svg,initial.svg,'Composition preview may not rewrite committed math');
  await send({op:'editorAction',handle:h,expectedRevision:2,eventSequence:5,action:{kind:'commitComposition',id,text:null}});
  const committed=await requestLayout(h,3);validate(committed.snapshot,committed.svg);assert.notEqual(committed.svg,initial.svg);
  assert.equal(committed.snapshot.fonts[1].id,fontIds.text);
  await send({op:'destroy',handle:h,expectedRevision:3});

  {
    const h=(await send({op:'create',documentId:'root-editor-parity',source:''})).handle;
    let revision=0, sequence=0;
    const action=async action=>{const r=await send({op:'editorAction',handle:h,expectedRevision:revision,eventSequence:++sequence,action});revision=r.result.editor.document.revision;return r.result.editor;};
    const edit=command=>action({kind:'edit',command,groupId:null});
    await edit({kind:'insertStructure',structure:'indexedRoot'});
    await edit({kind:'insertText',text:'3',mode:'math'});
    await action({kind:'moveHorizontal',direction:'right',extend:false});
    const done=await edit({kind:'insertText',text:'x',mode:'math'});
    assert.equal(done.document.source,String.raw`\sqrt[3]{x}`);
    const initial=await requestLayout(h,revision);validate(initial.snapshot,initial.svg);
    const beforeIds=done.document.analysis.tree.entries[0].node;
    const row={kind:'row',nodeId:done.document.analysis.tree.id,offset:0,affinity:'after'};
    await action({kind:'setSelection',selection:{kind:'range',anchor:row,focus:{...row,offset:1}}});
    const copy=await edit({kind:'duplicateSelection'});
    assert.equal(copy.document.analysis.facts.radicals,2);
    assert.notEqual(copy.document.analysis.tree.entries[1].node.radicand.id,beforeIds.radicand.id);
    await action({kind:'undo'});
    const restored=await requestLayout(h,revision);validate(restored.snapshot,restored.svg);assert.equal(restored.svg,initial.svg);
    await send({op:'destroy',handle:h,expectedRevision:revision});
  }
  {
    const h=(await send({op:'create',documentId:'fence-editor-parity',source:''})).handle;
    let revision=0,sequence=0;
    const action=async action=>{const r=await send({op:'editorAction',handle:h,expectedRevision:revision,eventSequence:++sequence,action});revision=r.result.editor.document.revision;return r.result;};
    const edit=command=>action({kind:'edit',command,groupId:null});
    await edit({kind:'typeDelimiter',text:'('});await edit({kind:'typeDelimiter',text:'['});
    await edit({kind:'insertText',text:'x',mode:'math'});
    await edit({kind:'typeDelimiter',text:')'});
    const before=await requestLayout(h,revision);validate(before.snapshot,before.svg);
    const closedInner=await edit({kind:'typeDelimiter',text:']'});
    assert.equal(closedInner.contentChanged,false);
    const closedOuter=await edit({kind:'typeDelimiter',text:')'});
    assert.equal(closedOuter.contentChanged,false);
    assert.equal(closedOuter.editor.document.source,String.raw`\left(\left[x)\right]\right)`);
    const after=await requestLayout(h,revision);assert.equal(after.svg,before.svg);assert.equal(after.snapshot.rebuiltFragments,0);
    await action({kind:'undo'});const undone=await requestLayout(h,revision);validate(undone.snapshot,undone.svg);
    await send({op:'destroy',handle:h,expectedRevision:revision});
  }
  {
    const h=(await send({op:'create',documentId:'accent-editor-parity',source:''})).handle;
    let revision=0,sequence=0;
    const action=async action=>{const r=await send({op:'editorAction',handle:h,expectedRevision:revision,eventSequence:++sequence,action});revision=r.result.editor.document.revision;return r.result;};
    const edit=command=>action({kind:'edit',command,groupId:null});
    await edit({kind:'insertAccent',command:String.raw`\hat`});await edit({kind:'insertAccent',command:String.raw`\vec`});
    const typed=await edit({kind:'insertText',text:'x',mode:'math'});
    const before=await requestLayout(h,revision);validate(before.snapshot,before.svg);
    const q=await send({op:'queryScene',handle:h,expectedRevision:revision,sceneToken:before.sceneToken,
      query:{kind:'selection',selection:typed.editor.selection}});
    assert.deepEqual(q.result.description.path,['root','accent','accentBody','accent','accentBody']);
    const exited=await action({kind:'moveHorizontal',direction:'right',extend:false});assert.equal(exited.contentChanged,false);
    assert.equal(exited.editor.selection.focus.nodeId,exited.editor.document.analysis.tree.id);
    const after=await requestLayout(h,revision);assert.equal(after.svg,before.svg);assert.equal(after.snapshot.rebuiltFragments,0);
    await edit({kind:'insertText',text:'+y',mode:'math'});await action({kind:'undo'});
    assert.equal((await requestLayout(h,revision)).svg,before.svg);
    await send({op:'destroy',handle:h,expectedRevision:revision});
  }
  {
    const h=(await send({op:'create',documentId:'style-editor-parity',source:'x+y'})).handle;
    let revision=0,sequence=0;
    const action=async action=>{const r=await send({op:'editorAction',handle:h,expectedRevision:revision,eventSequence:++sequence,action});revision=r.result.editor.document.revision;return r.result;};
    const edit=command=>action({kind:'edit',command,groupId:null});
    const initial=await send({op:'inspectEditor',handle:h,expectedRevision:0});
    const root=initial.snapshot.document.analysis.tree.id;
    await action({kind:'setSelection',selection:{kind:'range',anchor:{kind:'row',nodeId:root,offset:0,affinity:'after'},focus:{kind:'row',nodeId:root,offset:3,affinity:'after'}}});
    const styled=await edit({kind:'setStyle',style:String.raw`\mathbf`});
    assert.equal(styled.editor.document.source,String.raw`\mathbf{x+y}`);assert.equal(styled.editor.document.analysis.facts.styles,1);
    const bold=await requestLayout(h,revision);validate(bold.snapshot,bold.svg);
    const body=styled.editor.document.analysis.tree.entries[0].node.body;
    const q=await send({op:'queryScene',handle:h,expectedRevision:revision,sceneToken:bold.sceneToken,
      query:{kind:'selection',selection:{kind:'range',anchor:{kind:'row',nodeId:body.id,offset:0,affinity:'after'},focus:{kind:'row',nodeId:body.id,offset:0,affinity:'after'}}}});
    assert.deepEqual(q.result.description.path,['root','style','styleBody']);
    await action({kind:'undo'});const plain=await requestLayout(h,revision);validate(plain.snapshot,plain.svg);assert.notEqual(plain.svg,bold.svg);
    await action({kind:'redo'});const restored=await requestLayout(h,revision);assert.equal(restored.svg,bold.svg);
    await send({op:'destroy',handle:h,expectedRevision:revision});
  }
  {
    const h=(await send({op:'create',documentId:'nary-editor-parity',source:''})).handle;
    let revision=0,sequence=0;
    const action=async action=>{const r=await send({op:'editorAction',handle:h,expectedRevision:revision,eventSequence:++sequence,action});revision=r.result.editor.document.revision;return r.result;};
    const edit=command=>action({kind:'edit',command,groupId:null});
    await edit({kind:'insertOperator',command:String.raw`\sum`,withLower:true,withUpper:true});
    await edit({kind:'insertText',text:'i=1',mode:'math'});await action({kind:'moveHorizontal',direction:'right',extend:false});
    await edit({kind:'insertText',text:'n',mode:'math'});const typed=await action({kind:'moveHorizontal',direction:'right',extend:false});
    assert.equal(typed.editor.document.source,String.raw`\sum_{i=1}^{n}`);
    const before=await requestLayout(h,revision);validate(before.snapshot,before.svg);
    const part=typed.editor.document.analysis.tree.entries[0].node;
    const lower=part.parts.find(p=>p.kind==='script'&&p.slot.script==='subscript').slot.argument.id;
    const upper=part.parts.find(p=>p.kind==='script'&&p.slot.script==='superscript').slot.argument.id;
    const p={kind:'row',nodeId:lower,offset:0,affinity:'after'};
    const up=await send({op:'queryScene',handle:h,expectedRevision:revision,sceneToken:before.sceneToken,
      query:{kind:'vertical',from:p,down:false,desiredX:null}});assert.equal(up.result.position.nodeId,upper);
    await action({kind:'setSelection',selection:{kind:'range',anchor:p,focus:p}});
    const changed=await edit({kind:'setOperatorLimits',mode:'noLimits'});assert.equal(changed.contentChanged,true);
    const side=await requestLayout(h,revision);validate(side.snapshot,side.svg);assert.notEqual(side.svg,before.svg);
    await action({kind:'undo'});const restored=await requestLayout(h,revision);assert.equal(restored.svg,before.svg);
    await send({op:'destroy',handle:h,expectedRevision:revision});
  }
  {
    const h=(await send({op:'create',documentId:'spacing-phantom-editor-parity',source:''})).handle;
    let revision=0,sequence=0;
    const action=async action=>{const r=await send({op:'editorAction',handle:h,expectedRevision:revision,eventSequence:++sequence,action});revision=r.result.editor.document.revision;return r.result;};
    const edit=command=>action({kind:'edit',command,groupId:null});
    await edit({kind:'insertPhantom',command:String.raw`\phantom`});
    const typed=await edit({kind:'insertText',text:'x',mode:'math'});
    const hidden=await requestLayout(h,revision);validate(hidden.snapshot,hidden.svg);
    assert.equal(hidden.snapshot.glyphs.length,0);assert.equal(hidden.snapshot.rules.length,0);
    const q=await send({op:'queryScene',handle:h,expectedRevision:revision,sceneToken:hidden.sceneToken,query:{kind:'selection',selection:typed.editor.selection}});
    assert.ok(q.result.description.path.includes('phantomBody'));
    await action({kind:'moveHorizontal',direction:'right',extend:false});
    await edit({kind:'insertSpace',command:String.raw`\!`});
    await edit({kind:'insertBoxed'});const completed=await edit({kind:'insertText',text:'y',mode:'math'});
    assert.equal(completed.editor.document.source,String.raw`\phantom{x}\!\boxed{y}`);
    const drawn=await requestLayout(h,revision);validate(drawn.snapshot,drawn.svg);
    assert.equal(drawn.snapshot.glyphs.length,1);assert.equal(drawn.snapshot.rules.length,4);
    await action({kind:'undo'});await action({kind:'redo'});
    assert.equal((await requestLayout(h,revision)).svg,drawn.svg);
    const changed=await send({op:'replaceSource',handle:h,expectedRevision:revision,source:String.raw`\mathbf{x}+y`});revision=changed.snapshot.revision;
    const before=await requestLayout(h,revision);
    const source=changed.snapshot.source;const at=source.indexOf('{x')+1;
    const patched=await send({op:'patchSource',handle:h,expectedRevision:revision,edit:{startUtf16:at,endUtf16:at+1,replacement:'z'}});revision=patched.snapshot.revision;
    assert.equal(patched.snapshot.analysis.tree.entries[0].node.body.entries[0].node.value,'z','Style source patch must change real mathematical content');
    const after=await requestLayout(h,revision);validate(after.snapshot,after.svg);assert.notEqual(before.svg,after.svg);
    await send({op:'destroy',handle:h,expectedRevision:revision});
  }
  for(const command of ['overgroup','undergroup','overleftharpoon','overrightharpoon','overlinesegment','underlinesegment']) {
    const h=(await send({op:'create',documentId:`extended-accent-${command}`,source:''})).handle;
    let revision=0,sequence=0;
    const action=async action=>{const r=await send({op:'editorAction',handle:h,expectedRevision:revision,eventSequence:++sequence,action});revision=r.result.editor.document.revision;return r.result;};
    const edit=command=>action({kind:'edit',command,groupId:null});
    await edit({kind:'insertAccent',command:`\\${command}`});const typed=await edit({kind:'insertText',text:'x+y',mode:'math'});
    assert.equal(typed.editor.document.source,`\\${command}{x+y}`);
    const initial=await requestLayout(h,revision);validate(initial.snapshot,initial.svg);
    const mark=typed.editor.document.analysis.tree.entries[0].node;
    await action({kind:'setSelection',selection:{kind:'node',nodeId:mark.id}});
    const duplicate=await edit({kind:'duplicateSelection'});assert.notEqual(duplicate.editor.document.analysis.tree.entries[1].node.body.id,mark.body.id);
    await action({kind:'undo'});assert.equal((await requestLayout(h,revision)).svg,initial.svg);
    const next=await action({kind:'redo'});const copiedSource=next.editor.document.source;const at=copiedSource.lastIndexOf('{x')+1;
    const patched=await send({op:'patchSource',handle:h,expectedRevision:revision,edit:{startUtf16:at,endUtf16:at+1,replacement:'z'}});revision=patched.snapshot.revision;
    assert.equal(patched.snapshot.analysis.tree.entries[0].node.body.entries[0].node.value,'x');
    assert.equal(patched.snapshot.analysis.tree.entries[1].node.body.entries[0].node.value,'z');
    const final=await requestLayout(h,revision);validate(final.snapshot,final.svg);
    await send({op:'destroy',handle:h,expectedRevision:revision});
  }
  {
    const h=(await send({op:'create',documentId:'general-operator-editor',source:''})).handle;
    let revision=0,sequence=0;
    const action=async action=>{const r=await send({op:'editorAction',handle:h,expectedRevision:revision,eventSequence:++sequence,action});revision=r.result.editor.document.revision;return r.result;};
    const edit=command=>action({kind:'edit',command,groupId:null});
    await edit({kind:'insertGeneralOperator',withLower:true,withUpper:true});
    await edit({kind:'insertText',text:'a',mode:'math'});await action({kind:'moveHorizontal',direction:'right',extend:false});
    await edit({kind:'insertText',text:'b',mode:'math'});await action({kind:'moveHorizontal',direction:'right',extend:false});
    const done=await edit({kind:'insertText',text:'f',mode:'math'});assert.equal(done.editor.document.source,String.raw`\mathop{f}_{a}^{b}`);
    const op=done.editor.document.analysis.tree.entries[0].node;assert.equal(op.kind,'generalOperator');
    const first=await requestLayout(h,revision);validate(first.snapshot,first.svg);
    const bodyCaret={kind:'row',nodeId:op.body.id,offset:0,affinity:'after'};
    const selection={kind:'range',anchor:bodyCaret,focus:bodyCaret};
    const description=await send({op:'queryScene',handle:h,expectedRevision:revision,sceneToken:first.sceneToken,query:{kind:'selection',selection}});
    assert.deepEqual(description.result.description.path,['root','generalOperator','operatorBody']);
    await action({kind:'setSelection',selection});await edit({kind:'setOperatorLimits',mode:'noLimits'});
    const side=await requestLayout(h,revision);validate(side.snapshot,side.svg);assert.notEqual(side.svg,first.svg);
    await action({kind:'undo'});assert.equal((await requestLayout(h,revision)).svg,first.svg);
    await action({kind:'setSelection',selection:{kind:'node',nodeId:op.id}});
    const duplicate=await edit({kind:'duplicateSelection'});const copy=duplicate.editor.document.analysis.tree.entries[1].node;
    assert.notEqual(copy.body.id,op.body.id);assert.notEqual(copy.parts[0].slot.argument.id,op.parts[0].slot.argument.id);
    await action({kind:'undo'});assert.equal((await requestLayout(h,revision)).svg,first.svg);
    await send({op:'destroy',handle:h,expectedRevision:revision});
  }
  {
    const h=(await send({op:'create',documentId:'color-editor-parity',source:'x+y'})).handle;
    let revision=0,sequence=0;
    const action=async action=>{const r=await send({op:'editorAction',handle:h,expectedRevision:revision,eventSequence:++sequence,action});revision=r.result.editor.document.revision;return r.result;};
    const plain=await requestLayout(h,revision);validate(plain.snapshot,plain.svg);
    const p={kind:'row',nodeId:1,offset:0,affinity:'after'};
    await action({kind:'setSelection',selection:{kind:'range',anchor:p,focus:{...p,offset:3}}});
    const colored=await action({kind:'edit',command:{kind:'setColor',color:{r:255,g:0,b:0}}});
    assert.equal(colored.editor.document.source,String.raw`\textcolor[HTML]{FF0000}{x+y}`);
    const red=await requestLayout(h,revision);validate(red.snapshot,red.svg);
    assert.equal(red.snapshot.width,plain.snapshot.width);assert.equal(red.snapshot.height,plain.snapshot.height);
    assert.ok(red.snapshot.glyphs.every(g=>g.color?.r===255&&g.color?.g===0&&g.color?.b===0));
    const query=await send({op:'queryScene',handle:h,expectedRevision:revision,sceneToken:red.sceneToken,query:{kind:'selection',selection:colored.editor.selection}});
    assert.ok(query.result.description.path.includes('colorBody'));
    await action({kind:'undo'});assert.equal((await requestLayout(h,revision)).svg,plain.svg);
    await action({kind:'redo'});assert.equal((await requestLayout(h,revision)).svg,red.svg);
    await send({op:'destroy',handle:h,expectedRevision:revision});
  }
  for(const [i,draft] of [String.raw`\color{red}`,String.raw`\textcolor{red}{}`,String.raw`\phantom{}`,String.raw`\hphantom{}`,String.raw`\vphantom{}`,String.raw`\boxed{}`, String.raw`\sum_{}^n`,String.raw`\hat{}`,String.raw`\overset{}{x}`,String.raw`\underset{a}{}`, '', '{}','%only comment',String.raw`\frac{}{} `,String.raw`\frac{{}}{y}`,String.raw`x^{{}}`,String.raw`\sqrt[]{x}`,String.raw`\sqrt{}`,String.raw`\left(\right)`].entries()) {
    const h=(await send({op:'create',documentId:`blank-${i}`,source:draft})).handle;
    const result=await requestLayout(h,0,{},false);validate(result.snapshot,result.svg);assert.equal(result.snapshot.exportReady,false);
    await requestLayout(h,0,{},true,'incompleteSource');await send({op:'destroy',handle:h,expectedRevision:0});
  }
  for(const [i,[source,error]] of [[String.raw`\futureaccent{x}`,'unsupportedSyntax'],[String.raw`\frac{x}{`,'incompleteSource']].entries()) {
    const h=(await send({op:'create',documentId:`unsupported-${i}`,source})).handle;
    await requestLayout(h,0,{},true,error);const doc=await send({op:'inspectDocument',handle:h});assert.equal(doc.snapshot.source,source);
    await send({op:'destroy',handle:h,expectedRevision:0});
  }
  assert.equal((await send({op:'stats'})).activeSessions,0);
  console.log(JSON.stringify({status:'PASS',actualNativeExecution:true,actualWasmExecution:true,requestCases:cases,validatedScenes:scenes,
    maxNumericDifference,geometricToleranceEm:1e-9,exactSvgParity:true,fonts:registrations,
    nativeSha256:sha(readFileSync(nativePath)),wasmSha256:sha(readFileSync(wasmPath)),transcriptSha256:transcript.digest('hex'),
    scope:'Headless K0/P07 geometry, real font shaping, accents/annotations/large operators/math styles, SVG, cache and composition-state parity. No actual OS IME, browser input or Office UI acceptance.'},null,2));
} finally {
  clearTimeout(timer);wasm.free();native.stdin.end();lines.return?.();
  if(native.exitCode===null) native.kill();
}
