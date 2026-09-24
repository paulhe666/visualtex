// This checks the consistency of recorded audit evidence. It does not run or certify product acceptance.
import { readFileSync, writeFileSync, existsSync, readdirSync } from 'node:fs';
import { join, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
const root = dirname(fileURLToPath(import.meta.url));
const repo = resolve(root, '../..');
const evidence = join(root, 'evidence');
const read = name => JSON.parse(readFileSync(join(evidence,name),'utf8').replace(/^\uFEFF/,''));
const stable = x => Array.isArray(x) ? x.map(stable) : x && typeof x === 'object' ? Object.fromEntries(Object.keys(x).sort().map(k=>[k,stable(x[k])])) : x;
const equal = (a,b) => JSON.stringify(stable(a)) === JSON.stringify(stable(b));
const checks=[];
function check(name, ok, detail) { checks.push({name,ok:Boolean(ok),detail}); console.log(`${ok?'PASS':'FAIL'} | ${name} | ${JSON.stringify(detail)}`); }
const before=read('original-inventory.json'), after=read('final-original-inventory.json');
check('original document count',before.documents.length===8&&after.documents.length===8,after.documents.map(x=>x.name));
for(const d of before.documents){const current=after.documents.find(x=>x.name===d.name);const changed=Object.keys(d).filter(k=>k!=='key'&&!equal(d[k],current?.[k]));check('original preserved '+d.name,current&&!current.saved&&changed.length===0,{changed,saved:current?.saved});}
const cases=[
 ['repro01-converted',3,0,3],['repro01-edit-failed',4,0,2],
 ['audit-close-before',3,0,3],['audit-close-after',4,0,2],
 ['audit-single-omml-control-after',1,0,1],
 ['repro02-after-format',0,1,1],['audit-ref-final-after',0,1,1],
 ['repro03-left-vt-after',0,3,2],['repro03-right-mt-after',0,3,2],['audit-left-mt-after',0,3,2],['repro03-control-after',2,0,2],
 ['repro04-omml-import-unumbered',6,0,0],['repro04b-omml-numbered-after',6,0,5],
 ['repro04c-native-numbered-after',0,6,5],['repro04d-mathtype-cases-after',0,0,0],
 ['repro04e-seeded-omml-after',1,0,1],['repro04f-seeded-native-after',0,7,1],['repro04f-after-edit',0,7,2],
 ['repro05-omml-after',2,0,2],['audit-omml-soft-clean-after',2,0,1],['audit-vt-soft-clean-after',0,2,2],['audit-mt-soft-clean-after',0,2,2]
];
for(const [label,math,ole,number] of cases){const d=read(label+'-inventory.json').documents[0];const observed=[d.maths.length,d.shapes.length,d.fields.filter(f=>/SEQ VisualTeXEquation|MACROBUTTON MTPlaceRef/.test(f.code)).length];check('snapshot '+label,equal(observed,[math,ole,number]),{document:d.name,observed});}
const messages=[
 ['repro03-left-vt-result-native.json','non-empty boundary'],
 ['repro03-right-mt-result-native.json','exposed 2/3'],
 ['audit-left-mt-result-native.json','exposed 2/3'],
 ['repro04e-seeded-omml-result-native.json','disappeared before row grouping'],
 ['repro04d-mathtype-cases-result-native.json','invalid standalone MathType MTEF']
];
for(const [file,text]of messages){const data=read(file);check('recorded dialog '+file,JSON.stringify(data).includes(text),text);}
const rt=read('codec-roundtrip.json');
check('cases signature difference',rt.find(x=>x.case==='actual')?.expectedSignature===rt.find(x=>x.case==='actual')?.actualSignature+'o()', 'only trailing empty operator differs');
check('cases pure-data control',rt.find(x=>x.case==='actual_without_empty_close')?.equal===true&&rt.find(x=>x.case==='actual_with_empty_second_column')?.equal===false,'not a missing second-column fix');
const fonts=read('repro04-omml-import-unumbered-inventory.json').documents[0].maths.map(x=>x.font);
check('import font evidence',fonts.length===6&&fonts.every(x=>x==='Cambria Math'),fonts);
const head=execFileSync('git',['rev-parse','HEAD'],{cwd:repo,encoding:'utf8'}).trim();
const logicalDiff=execFileSync('git',['diff','--numstat'],{cwd:repo,encoding:'utf8'}).trim();
check('baseline HEAD',head==='3d207d7a81ff40014fd861b242cff96ea77060a2',head);
check('no tracked logical changes',logicalDiff==='',logicalDiff);
const sha=p=>createHash('sha256').update(readFileSync(p)).digest('hex');
const installedRoot=join(process.env.ProgramFiles??'C:\\Program Files','VisualTeX','WindowsOffice','VSTO');
const binaries=[
 ['desktop',join(process.env.LOCALAPPDATA??'','VisualTeX','visualtex.exe'),join(repo,'apps/windows/src-tauri/target/release/visualtex.exe')],
 ['Word VSTO',join(installedRoot,'VisualTeX.WordVsto.dll'),join(repo,'apps/windows/src-windows/VisualTeX.WordVsto/bin/x64/Release/net472/VisualTeX.WordVsto.dll')]
];
for(const[name,installed,built]of binaries){
 if(existsSync(installed)&&existsSync(built)){
  const a=sha(installed),b=sha(built);
  if(name==='desktop'){
   // NSIS packaging patches the Tauri bundle-type marker. The packed payload,
   // not the post-bundle raw target executable, is the installation identity.
   const installer=join(repo,'apps/windows/src-tauri/target/release/bundle/nsis/VisualTeX_1.2.6_x64-setup.exe');
   const sevenZip=join(process.env.ProgramFiles??'C:\\Program Files','7-Zip','7z.exe');
   const payload=execFileSync(sevenZip,['e','-so',installer,'visualtex.exe'],{maxBuffer:40*1024*1024,timeout:120000,windowsHide:true});
   const packedSha=createHash('sha256').update(payload).digest('hex');
   const installedBytes=readFileSync(installed),buildBytes=readFileSync(built);let differentBytes=0;const offsets=[];
   for(let i=0;i<Math.max(installedBytes.length,buildBytes.length);i++){if(installedBytes[i]!==buildBytes[i]){differentBytes++;if(offsets.length<10)offsets.push(i);}}
   check('installed/NSIS payload hash '+name,a===packedSha,{installed,installer,installedSha256:a,packagedSha256:packedSha,rawBuildSha256:b,rawBuildDifferentBytes:differentBytes,rawBuildDifferenceOffsets:offsets});
  }else check('installed/build hash '+name,a===b,{installed,built,installedSha256:a,buildSha256:b});
 }else check('installed/build paths '+name,false,{installed,built,installedExists:existsSync(installed),buildExists:existsSync(built)});
}
const records=readFileSync(join(evidence,'ui-actions.ndjson'),'utf8').trim().split('\n').map(x=>JSON.parse(x));
const summary={scope:'Audit evidence consistency, not repaired-product acceptance',time:new Date().toISOString(),head,checks,recordedUiActions:records.length,failedAutomationActions:records.filter(x=>x.status!==0).length,evidenceFiles:readdirSync(evidence).length,allChecksPassed:checks.every(x=>x.ok)};
writeFileSync(join(evidence,'final-audit-verification.json'),JSON.stringify(summary,null,2)+'\n','utf8');
console.log('AUDIT_EVIDENCE_CHECKS',checks.length,'PASSED',checks.filter(x=>x.ok).length);
if(!summary.allChecksPassed)process.exitCode=1;
