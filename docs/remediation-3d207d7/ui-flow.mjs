import {spawnSync} from 'node:child_process';
import {appendFileSync, mkdirSync, existsSync, readFileSync} from 'node:fs';
import {join, dirname, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
const dir=dirname(fileURLToPath(import.meta.url));
const root=resolve(dir,'../..');const output=join(dir,'evidence');mkdirSync(output,{recursive:true});
const ps=join(process.env.SystemRoot??'C:\\Windows','System32','WindowsPowerShell','v1.0','powershell.exe');
const targetFile=join(output,'target-word.json');
const recordedTarget=existsSync(targetFile)?JSON.parse(readFileSync(targetFile,'utf8').replace(/^\uFEFF/, '')):{};
const targetWordPid=Number(process.env.VISUALTEX_PERF_WORD_PID??recordedTarget.pid??0);
const pinnedDocument=process.env.VISUALTEX_PERF_DOCUMENT;
const pinnedStage=process.env.VISUALTEX_PERF_STAGE??recordedTarget.stage;
if(process.env.VISUALTEX_PERF_WORD_PID&&(!pinnedDocument||!process.env.VISUALTEX_PERF_STAGE))throw new Error('Pinned performance runs require the exact PID, document and stage.');
function exec(script,params={},quiet=false){if(script==='word-ui.ps1'&&pinnedDocument)params={ExpectedDocument:pinnedDocument,...params};if(targetWordPid && ['word-ui.ps1','inspect-word.ps1'].includes(script))params={WordProcessId:targetWordPid,...params};const args=['-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',join(dir,script)];for(const[k,v]of Object.entries(params)){if(v===false||v===undefined)continue;args.push('-'+k);if(v!==true)args.push(String(v));}const t=Date.now();const r=spawnSync(ps,args,{cwd:root,encoding:'utf8',timeout:40000,maxBuffer:4*1024*1024});const record={time:new Date().toISOString(),script,params,elapsed:Date.now()-t,status:r.status,stdout:r.stdout,stderr:r.stderr,error:r.error?.message};appendFileSync(join(output,'ui-actions.ndjson'),JSON.stringify(record)+'\n');if(!quiet)process.stdout.write((r.stdout??'')+(r.stderr??''));if(r.error||r.status!==0)throw new Error(r.error?.message??r.stderr??`PowerShell exit ${r.status}`);return r.stdout;}
function ui(Action,params={},quiet=false){return exec('word-ui.ps1',{WindowHandle:-1,Action,...params},quiet)}
const pause=ms=>new Promise(r=>setTimeout(r,ms));
async function waitForCreateEditor(format, display) {
 const deadline=Date.now()+15000;
 while(Date.now()<deadline) {
  assertNoWordErrorDialog('create-editor-ready');
  const windows=JSON.parse(ui('windows',{},true));
  const editors=windows.filter(w=>w.process==='visualtex'&&w.title.startsWith('Office'));
  if(editors.length>1)throw new Error('Multiple Office editors are visible; refusing ambiguous UI input.');
  if(editors.length===1){
   ui('probe',{Target:'editor',Label:'create-editor-ready'},true);
   const controls=JSON.parse(readFileSync(join(output,'create-editor-ready-uia.json'),'utf8'));
   const id=controls.map(c=>c.value??'').join('\n').match(/sessionId=([0-9a-f-]+)/i)?.[1];
   if(id){
    const path=join(process.env.APPDATA,'com.visualtex.studio','office','sessions',id,'session.json');
    const session=JSON.parse(readFileSync(path,'utf8').replace(/^\uFEFF/,''));
    const documentName=(session.sourceDocumentId??'').split(/[\\/]/).pop();
    if(session.mode!=='create'||session.displayMode!==display||(pinnedDocument&&documentName!==pinnedDocument))
     throw new Error(`Unexpected real editor session: ${id} ${session.mode} ${session.displayMode} ${documentName}`);
    if(format==='omml'&&session.objectMode!=='wordOmml')throw new Error('The real editor is not the requested OMML session.');
    if(controls.some(c=>c.type==='ControlType.Button'&&c.name==='完成并插入'&&c.enabled))return;
   }
  }
  await pause(400);
 }
 throw new Error('The exact real create editor was not ready in 15 seconds; do not click a second Ribbon action.');
}
async function readReadySourceEditor(label) {
 const deadline=Date.now()+15000;
 while(Date.now()<deadline) {
  const visibleEditors=JSON.parse(ui('windows',{},true)).filter(w=>w.process==='visualtex'&&w.title.startsWith('Office'));
  if(visibleEditors.length>1)throw new Error('Multiple real Office editors are visible; no input was sent.');
  if(!visibleEditors.length){assertNoWordErrorDialog(label);await pause(400);continue;}
  ui('probe',{Target:'editor',Label:label},true);
  const controls=JSON.parse(readFileSync(join(output,label+'-uia.json'),'utf8'));
  const edits=controls.filter(c=>c.type==='ControlType.Edit');
  const sources=edits.filter(c=>!c.name&&c.enabled&&!c.offscreen&&c.rect!=='Empty');
  if(sources.length===1)return {edits,source:sources[0]};
  if(sources.length>1)throw new Error('Multiple real source editors are visible; no input was sent.');
  assertNoWordErrorDialog(label);
  await pause(600);
 }
 throw new Error('The existing editor did not expose its source input in 15 seconds; no input was sent.');
}
async function waitForWordDialog() {
 const deadline=Date.now()+30000;
 while(Date.now()<deadline) {
  const windows=JSON.parse(ui('windows',{},true));
  if(windows.some(w=>w.process==='WINWORD'&&w.title.startsWith('VisualTeX')&&w.pid===targetWordPid))return;
  await pause(700);
 }
 throw new Error('The real Word dialog did not appear within 30 seconds; inspect Word before continuing.');
}
function operationLog() {
 const file=join(output,pinnedStage+'-word-hook.log');
 return existsSync(file)?readFileSync(file,'utf8'):'';
}
function redrawLog() {
 const file=join(output,pinnedStage+'-word-redraw.log');
 return existsSync(file)?readFileSync(file,'utf8'):'';
}
async function waitForRedraw(label,start,startedAt) {
 const deadline=Date.now()+180000;
 while(Date.now()<deadline) {
  const log=redrawLog().slice(start);
  assertNoWordErrorDialog(label);
  if(/redraw-(failed|cancelled) /.test(log))throw new Error('The installed redraw operation failed or was cancelled. Inspect its real log/document.');
  const complete=log.match(/redraw-complete formulas=.+/);
  if(complete){
   appendFileSync(join(output,'ui-actions.ndjson'),JSON.stringify({time:new Date().toISOString(),action:'observed-real-redraw-complete',label,elapsedMs:Date.now()-startedAt,record:complete[0]})+'\n');
   console.log(complete[0]);return;
  }
  await pause(1200);
 }
 throw new Error('The installed redraw operation has not reported completion; do not repeat it.');
}
async function waitForWordOperation(kind,label,start) {
 const deadline=Date.now()+180000;
 const completed=kind==='conversion'?'format-conversion-complete source=':kind==='reference'?'reference-insertion-complete target=':'bulk-import-complete blocks=';
 const failed=kind==='conversion'?'format-conversion-failed sourceMode=':kind==='reference'?'reference-insertion-failed':'bulk-import-failed';
 while(Date.now()<deadline) {
  const log=operationLog().slice(start);
  if(log.includes(failed)) {
   await waitForWordDialog();assertNoWordErrorDialog(label);
   throw new Error('The installed plugin reported an operation failure; inspect its real error.');
  }
  if(log.includes(completed)) {assertNoWordErrorDialog(label);return;}
  // Observe actual windows while waiting; a late error must not be missed merely
  // because a fixed delay elapsed. Word COM/XML and screenshot checks still follow.
  assertNoWordErrorDialog(label);
  await pause(1200);
 }
 throw new Error('The installed Word operation has not reported completion in 180 seconds; inspect without restarting it.');
}
async function waitForEditorRelease(label,start) {
 const deadline=Date.now()+90000;
 while(Date.now()<deadline) {
  assertNoWordErrorDialog(label);
  const released=operationLog().slice(start).match(/ribbon-session-operation-released sessionId=([0-9a-f-]+)/i);
  if(released) {
   const file=join(process.env.APPDATA,'com.visualtex.studio','office','sessions',released[1],'session.json');
   const session=JSON.parse(readFileSync(file,'utf8').replace(/^\uFEFF/,''));
   appendFileSync(join(output,'ui-actions.ndjson'),JSON.stringify({time:new Date().toISOString(),action:'observed-editor-release',label,id:released[1],status:session.status,error:session.error})+'\n');
   if(session.status==='failed')throw new Error(`Real editor session failed: ${session.error}`);
   return;
  }
  await pause(1000);
 }
 throw new Error('The real editor operation has not released; inspect without repeating the edit.');
}
function snapshot(label) {
 const windows=JSON.parse(ui('windows',{},true));
 const expectedDocument=pinnedDocument??JSON.parse(readFileSync(targetFile,'utf8').replace(/^\uFEFF/, '')).document;
 const win=windows.find(w=>w.process==='WINWORD'&&w.title.endsWith(' - Word')&&(!targetWordPid||w.pid===targetWordPid)&&(!expectedDocument||w.title===expectedDocument+' - Word'));
 const number=win?.title.match(/(\d+) - Word$/)?.[1];
 if(!win || !expectedDocument)throw new Error('The recorded live test document is unavailable');
 exec('inspect-word.ps1',{DocumentName:expectedDocument,Label:label});
}
function assertNoWordErrorDialog(label) {
 const windows=JSON.parse(ui('windows',{},true));
 const dialog=windows.find(w=>w.process==='WINWORD'&&w.title.startsWith('VisualTeX')&&(!targetWordPid||w.pid===targetWordPid));
 if(!dialog)return;
 const controls=ui('native',{Target:'dialog',Label:label+'-result'});
 ui('screenshot',{Target:'dialog',Label:label+'-result'});
 throw new Error(`Real Word dialog requires inspection; flow stopped: ${dialog.title}\n${controls}`);
}
const [action,...argv]=process.argv.slice(2);
if(action==='insert'||action==='resume-insert'){
 const logStart=operationLog().length;
 const [format='native',display='block',side='right',preset='勾股定理']=argv;
 if(action==='insert'){
  ui('ribbon',{},true);
  ui('click',{Name:(format==='omml'?'OMML':'OLE')+(display==='block'?' 行间公式':' 行内公式')});
 }
 await waitForCreateEditor(format,display);
 if(format!=='omml') ui('combo',{Target:'editor',Name:'公式对象格式',Choice:format==='native'?'VisualTeX OLE':'MathType OLE'});
 if(display==='block')ui('check',{Target:'editor',Name:'编号',Checked:true});
 if(format==='mathType'&&display==='block')ui('combo',{Target:'editor',Name:'MathType 公式编号位置',Choice:side==='left'?'左侧':'右侧'});
 ui('click',{Target:'editor',Name:preset});
 ui('click',{Target:'editor',Name:'完成并插入'});
 await waitForEditorRelease('insert-result',logStart);
 console.log('INSERT_FLOW_COMPLETED',format,display,side,preset);
}else if(action==='edit'||action==='edit-ole'){
 const logStart=operationLog().length;
 const [index='2',preset='keep',numbered='keep',completion='update']=argv;
 ui(action==='edit-ole'?'select-ole':'select-math',{WindowHandle:0,Index:index});
 ui('ribbon',{},true);ui('click',{Name:'编辑所选公式'});await pause(1600);
 if(preset!=='keep')ui('click',{Target:'editor',Name:preset});
 if(numbered!=='keep')ui('check',{Target:'editor',Name:'编号',Checked:numbered==='on'});
 if(completion==='close')ui('keys',{Target:'editor',Value:'%{F4}'});
 else ui('click',{Target:'editor',Name:'更新公式'});
 await waitForEditorRelease('edit-result',logStart);
 assertNoWordErrorDialog('edit-result');
}else if(action==='edit-source'||action==='resume-edit-source'){
 const logStart=operationLog().length;
 const [kind='ole',index='3',sourceArgument='E=mc^2',label='edit-source']=argv;
 const source=sourceArgument.startsWith('@')?readFileSync(resolve(dir,sourceArgument.slice(1)),'utf8').replace(/^\uFEFF/,'').trim():sourceArgument;
 if(!['ole','math'].includes(kind))throw new Error('Expected ole or math editor target');
 if(action==='edit-source'){
  ui(kind==='ole'?'select-ole':'select-math',{WindowHandle:0,Index:index});
  ui('ribbon',{},true);ui('click',{Name:'编辑所选公式'});await pause(1600);
 }
 const editor=await readReadySourceEditor(label+'-before');
 ui('paste',{Target:'editor',Index:editor.edits.indexOf(editor.source)+1,Value:source});
 ui('probe',{Target:'editor',Label:label+'-prepared'},true);
 const actual=JSON.parse(readFileSync(join(output,label+'-prepared-uia.json'),'utf8'))
   .filter(c=>c.type==='ControlType.Edit'&&!c.name&&c.enabled&&!c.offscreen&&c.rect!=='Empty');
 if(actual.length!==1||actual[0].value!==source)throw new Error('Real source editor did not retain the requested text');
 if(process.env.VISUALTEX_PERF_REQUIRE_CHANGE==='1'){
  const controls=JSON.parse(readFileSync(join(output,label+'-prepared-uia.json'),'utf8'));
  const id=controls.map(c=>c.value??'').join('\n').match(/sessionId=([0-9a-f-]+)/i)?.[1];
  if(!id)throw new Error('No exact editor session identity for the changed-content measurement.');
  const path=join(process.env.APPDATA,'com.visualtex.studio','office','sessions',id,'session.json');
  const deadline=Date.now()+7000;let ready=false;
  while(Date.now()<deadline){
   const draft=JSON.parse(readFileSync(path,'utf8').replace(/^\uFEFF/,''));
   const text=draft.lines?.map(line=>line.latex).join('\n')??'';
   const original=draft.originalMetadata?.lines?.map(line=>line.latex).join('\n')??draft.originalMetadata?.latex??'';
   if(draft.status==='editing'&&draft.dirty&&text&&text!==original&&draft.exportResult?.mathMl){ready=true;break;}
   await pause(100);
  }
  if(!ready)throw new Error('The source remains an unchanged/incomplete draft; refusing to count a no-op as a fast edit.');
 }
 ui('click',{Target:'editor',Name:'更新公式'});
 await waitForEditorRelease(label,logStart);
 assertNoWordErrorDialog(label);
}else if(action==='edit-size'){
 const logStart=operationLog().length;
 const [kind='math',index='3',choice='小四（12 磅）',label='edit-size']=argv;
 if(!['ole','math'].includes(kind))throw new Error('Expected ole or math editor target');
 ui(kind==='ole'?'select-ole':'select-math',{WindowHandle:0,Index:index});
 ui('ribbon',{},true);ui('click',{Name:'编辑所选公式'});
 await readReadySourceEditor(label+'-before');
 ui('combo',{Target:'editor',Name:'公式字号',Choice:choice});
 ui('probe',{Target:'editor',Label:label+'-prepared'},true);
 const controls=JSON.parse(readFileSync(join(output,label+'-prepared-uia.json'),'utf8'));
 const actual=controls.find(c=>c.type==='ControlType.ComboBox'&&c.name==='公式字号')?.value;
 if(actual!==choice)throw new Error('The actual editor size does not match the selected option: '+actual);
 ui('click',{Target:'editor',Name:'更新公式'});
 await waitForEditorRelease(label,logStart);
 assertNoWordErrorDialog(label);
}else if(action==='reference'){
 const logStart=operationLog().length;
 ui('position',{WindowHandle:0});ui('enter',{WindowHandle:0});
 ui('ribbon',{},true);ui('click',{Name:'插入公式引用'});await pause(500);
 ui('dlglist',{Target:'dialog',Index:argv[0]??'2'});
 ui('dlgclick',{Target:'dialog',Name:'插入引用'});await pause(700);
 await waitForWordOperation('reference',argv[1]??'reference-insert',logStart);
 assertNoWordErrorDialog(argv[1]??'reference-insert');
}else if(action==='number-format'){
 const [format='continuous',label='number-format']=argv;
 const names={'continuous':'全文连续编号（1）','heading1-dot':'按章编号（1.1）','heading1-dash':'按章编号（1-1）','heading2-dot':'按节编号（1.1.1）','heading2-dash':'按节编号（1.1-1）'};
 if(!names[format])throw new Error('Unknown number format');
 ui('ribbon',{},true);ui('click',{Name:'编号格式'});ui('click',{Name:names[format]});await pause(700);
 assertNoWordErrorDialog(label);console.log('NUMBER_FORMAT_NO_ERROR_DIALOG');
}else if(action==='convert'){
 const [pair='vt-omml',scope='full',label='convert']=argv;
 const logStart=operationLog().length;
 const names={'vt-omml':'VisualTeX → OMML','vt-mt':'VisualTeX → MathType','omml-mt':'OMML → MathType','mt-omml':'MathType → OMML','mt-vt':'MathType → VisualTeX','omml-vt':'OMML → VisualTeX'};
 ui('ribbon',{},true);ui('click',{Name:names[pair]});ui('click',{Name:scope==='full'?'全文批量转换':'转换选中部分'});await pause(800);
 if(scope==='full'){await waitForWordDialog();ui('native',{Target:'dialog',Label:label+'-confirm'});ui('dlgclick',{Target:'dialog',Name:'是'});} await pause(800);
 await waitForWordOperation('conversion',label,logStart);
 assertNoWordErrorDialog(label);console.log('CONVERSION_NO_ERROR_DIALOG');
}else if(action==='redraw'||action==='redraw-existing'){
 const [format='omml',breaks='soft-breaks',label='redraw',input='redraw-source.tex',latinFont='',numbered='true']=argv;
 const redrawStart=redrawLog().length;
 if(action==='redraw') {
  exec('word-ui.ps1',{Action:'text',InputFile:input,Name:breaks,Index:12,LatinFont:latinFont});
  snapshot(label+'-source');
 }
 ui('ribbon',{},true);ui('click',{Name:'重绘全文'});
 ui('click',{Name:'全文 LaTeX 重绘为 '+(format==='omml'?'Word OMML':format==='native'?'VisualTeX OLE':'MathType')});
 ui('native',{Target:'dialog',Label:label+'-confirm'});
 ui('dlgcheck',{Target:'dialog',Name:'为全部',Checked:numbered==='true'});
 const startedAt=Date.now();ui('dlgclick',{Target:'dialog',Name:'开始重绘'});
 await waitForRedraw(label,redrawStart,startedAt);
 assertNoWordErrorDialog(label);console.log('REDRAW_NO_ERROR_DIALOG');
}else if(action==='import'||action==='resume-import'){
 const [format='omml',numbered='true',label='import',input='import-source.tex']=argv;
 if(!existsSync(resolve(dir,input)))throw new Error('Import source file does not exist: '+resolve(dir,input));
 const logStart=operationLog().length;
 if(action==='import'){ui('ribbon',{},true);ui('click',{Name:'批量导入'});await pause(2200);}
 ui('paste',{Target:'desktop',Name:'文档源码',InputFile:input});
 if(format!=='keep') ui('combo',{Target:'desktop',Name:'公式格式',Choice:{omml:'Word 原生 OMML',native:'VisualTeX OLE',mathType:'MathType OLE'}[format]});
 ui('check',{Target:'desktop',Name:'所有行间公式添加编号',Checked:numbered==='true'});
 ui('probe',{Target:'desktop',Label:label+'-preview'},true);
 const controls=JSON.parse(readFileSync(join(output,label+'-preview-uia.json'),'utf8'));
 const actualFormat=controls.find(c=>c.type==='ControlType.ComboBox'&&c.name==='公式格式')?.value;
 const expectedFormat={omml:'Word 原生 OMML',native:'VisualTeX OLE',mathType:'MathType OLE'}[format];
 if(expectedFormat&&actualFormat!==expectedFormat)throw new Error(`Import format was not selected: actual=${actualFormat} expected=${expectedFormat}`);
 const actualNumbered=controls.find(c=>c.type==='ControlType.CheckBox'&&c.name==='所有行间公式添加编号')?.value;
 if(actualNumbered!==(numbered==='true'?'On':'Off'))throw new Error('Import numbering checkbox does not match requested state.');
 ui('click',{Target:'desktop',Name:'导入到 Word'});await pause(800);
 await waitForWordOperation('import',label,logStart);
 assertNoWordErrorDialog(label);console.log('IMPORT_NO_ERROR_DIALOG');
}else if(action==='new'){
 exec('word-ui.ps1',{Action:'new',Label:argv[0]??'repro'});await pause(800);ui('ribbon',{},true);
}else if(action==='snapshot'){
 snapshot(argv[0]??'snapshot');
}else if(action==='lists'){
 exec('inspect-lists.ps1',{Label:argv[0]??'lists'});
}else if(action==='activate-redraw'){
 exec('activate-recorded-test.ps1',{EvidenceName:argv[0]});
}else if(action==='jump'){
 exec('jump-reference.ps1',{Label:argv[0],Index:argv[1]??'1'});
 ui('screenshot',{Label:argv[0]+'-visible'});
}else if(action==='next'){
 ui('position',{WindowHandle:0});ui('enter',{WindowHandle:0});
}else if(action==='ui'){
 const spec=JSON.parse(argv.join(' '));ui(spec.Action,spec);
}else throw new Error('Unknown flow '+action);
