import {spawnSync} from 'node:child_process';
import {appendFileSync, mkdirSync} from 'node:fs';
import {join, dirname, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
const dir=dirname(fileURLToPath(import.meta.url));
const root=resolve(dir,'../..');const output=join(dir,'evidence');mkdirSync(output,{recursive:true});
const ps=join(process.env.SystemRoot??'C:\\Windows','System32','WindowsPowerShell','v1.0','powershell.exe');
function exec(script,params={},quiet=false){const args=['-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',join(dir,script)];for(const[k,v]of Object.entries(params)){if(v===false||v===undefined)continue;args.push('-'+k);if(v!==true)args.push(String(v));}const t=Date.now();const r=spawnSync(ps,args,{cwd:root,encoding:'utf8',timeout:40000,maxBuffer:4*1024*1024});const record={time:new Date().toISOString(),script,params,elapsed:Date.now()-t,status:r.status,stdout:r.stdout,stderr:r.stderr,error:r.error?.message};appendFileSync(join(output,'ui-actions.ndjson'),JSON.stringify(record)+'\n');if(!quiet)process.stdout.write((r.stdout??'')+(r.stderr??''));if(r.error||r.status!==0)throw new Error(r.error?.message??r.stderr??`PowerShell exit ${r.status}`);return r.stdout;}
function ui(Action,params={},quiet=false){return exec('word-ui.ps1',{WindowHandle:-1,Action,...params},quiet)}
const pause=ms=>new Promise(r=>setTimeout(r,ms));
const [action,...argv]=process.argv.slice(2);
if(action==='insert'){
 const [format='native',display='block',side='right',preset='勾股定理']=argv;
 ui('ribbon',{},true);
 ui('click',{Name:(format==='omml'?'OMML':'OLE')+(display==='block'?' 行间公式':' 行内公式')});await pause(1600);
 if(format!=='omml') ui('combo',{Target:'editor',Name:'公式对象格式',Value:format==='native'?'{HOME}{ENTER}':'{HOME}{DOWN}{ENTER}'});
 if(display==='block')ui('check',{Target:'editor',Name:'编号',Checked:true});
 if(format==='mathType'&&display==='block')ui('combo',{Target:'editor',Name:'MathType 公式编号位置',Value:side==='left'?'{HOME}{ENTER}':'{END}{ENTER}'});
 ui('click',{Target:'editor',Name:preset});
 ui('click',{Target:'editor',Name:'完成并插入'});await pause(1600);
 console.log('INSERT_FLOW_COMPLETED',format,display,side,preset);
}else if(action==='convert'){
 const [pair='vt-omml',scope='full',label='convert']=argv;
 const names={'vt-omml':'VisualTeX → OMML','vt-mt':'VisualTeX → MathType','omml-mt':'OMML → MathType','mt-omml':'MathType → OMML','mt-vt':'MathType → VisualTeX','omml-vt':'OMML → VisualTeX'};
 ui('ribbon',{},true);ui('click',{Name:names[pair]});ui('click',{Name:scope==='full'?'全文批量转换':'转换选中部分'});await pause(800);
 if(scope==='full'){ui('native',{Target:'dialog',Label:label+'-confirm'});ui('dlgclick',{Target:'dialog',Name:'是'});} await pause(3500);
 const windows=JSON.parse(ui('windows',{},true));
 if(windows.some(w=>w.process==='WINWORD'&&w.title.startsWith('VisualTeX'))){ui('native',{Target:'dialog',Label:label+'-result'});ui('screenshot',{Target:'dialog',Label:label+'-result'});}else console.log('CONVERSION_NO_ERROR_DIALOG');
}else if(action==='redraw'){
 const [format='omml',breaks='soft-breaks',label='redraw']=argv;
 exec('word-ui.ps1',{Action:'text',InputFile:'redraw-source.tex',Name:breaks,Index:12});
 ui('ribbon',{},true);ui('click',{Name:'重绘全文'});
 ui('click',{Name:'全文 LaTeX 重绘为 '+(format==='omml'?'Word OMML':format==='native'?'VisualTeX OLE':'MathType')});
 ui('native',{Target:'dialog',Label:label+'-confirm'});
 ui('dlgcheck',{Target:'dialog',Name:'为全部',Checked:true});ui('dlgclick',{Target:'dialog',Name:'开始重绘'});await pause(4000);
 const windows=JSON.parse(ui('windows',{},true));if(windows.some(w=>w.process==='WINWORD'&&w.title.startsWith('VisualTeX')))ui('native',{Target:'dialog',Label:label+'-result'});else console.log('REDRAW_NO_ERROR_DIALOG');
}else if(action==='import'){
 const [format='omml',numbered='true',label='import',input='import-source.tex']=argv;
 ui('ribbon',{},true);ui('click',{Name:'批量导入'});await pause(2200);
 ui('paste',{Target:'desktop',Name:'文档源码',InputFile:input});
 if(format!=='keep') ui('combo',{Target:'desktop',Name:'公式格式',Value:format==='omml'?'{HOME}{ENTER}':format==='native'?'{HOME}{DOWN}{ENTER}':'{END}{ENTER}'});
 ui('check',{Target:'desktop',Name:'所有行间公式添加编号',Checked:numbered==='true'});
 ui('probe',{Target:'desktop',Label:label+'-preview'},true);ui('click',{Target:'desktop',Name:'导入到 Word'});await pause(4500);
 const windows=JSON.parse(ui('windows',{},true));
 if(windows.some(w=>w.process==='WINWORD'&&w.title.startsWith('VisualTeX'))){ui('native',{Target:'dialog',Label:label+'-result'});ui('screenshot',{Target:'dialog',Label:label+'-result'});}else console.log('IMPORT_NO_ERROR_DIALOG');
}else if(action==='new'){
 exec('word-ui.ps1',{Action:'new',Label:argv[0]??'repro'});await pause(800);ui('ribbon',{},true);
}else if(action==='snapshot'){
 const windows=JSON.parse(ui('windows',{},true));const win=windows.find(w=>w.process==='WINWORD'&&w.title.endsWith(' - Word'));
 const number=win?.title.match(/(\d+) - Word$/)?.[1];if(!number)throw new Error('No numbered live test document');
 exec('inspect-word.ps1',{DocumentNumber:number,Label:argv[0]??'snapshot'});
}else if(action==='next'){
 ui('keys',{Value:'^{END}{ENTER}'});await pause(300);
}else if(action==='ui'){
 const spec=JSON.parse(argv.join(' '));ui(spec.Action,spec);
}else throw new Error('Unknown flow '+action);
