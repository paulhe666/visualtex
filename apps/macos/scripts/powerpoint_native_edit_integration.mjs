// Exercises native MathZones resolution, native source export and replacement
// through the reviewed VBA/Rust code. Only NativeOmmlAcceptance.pptm is mutated.
import { execFileSync } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync, rmSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const output=resolve(root,'build-logs/powerpoint-native-omml');
const sessions=resolve(homedir(),'Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXRuntime/OfficeSessions');
const pointer=resolve(sessions,'powerpoint-active-session.txt');
const previous=existsSync(pointer)?readFileSync(pointer):null;
const payload=resolve(output,'native-inline.bin');
function as(lines){ return execFileSync('/usr/bin/osascript',lines.flatMap(l=>['-e',l]),{encoding:'utf8',timeout:25000}).trim(); }
function macro(name,args=[]){return as(['tell application "Microsoft PowerPoint"',`run VB macro macro name ${JSON.stringify(name)} list of parameters {${args.map(JSON.stringify).join(',')}}`,'end tell']);}
function nativeClipboard(macroName,args=[],dataPath=payload){return as([
 'use framework "AppKit"','use scripting additions',
 `set dataBytes to current application's NSData's dataWithContentsOfFile:${JSON.stringify(dataPath)}`,
 'if dataBytes is missing value then error "Missing native fixture"',
 "set pb to current application's NSPasteboard's generalPasteboard()", "set saved to current application's NSMutableArray's array()",
 "repeat with sourceItem in pb's pasteboardItems()", "set newItem to current application's NSPasteboardItem's alloc()'s init()",
 "repeat with t in sourceItem's types()", "newItem's setData:(sourceItem's dataForType:t) forType:t",'end repeat',"saved's addObject:newItem",'end repeat',
 "pb's clearContents()", "pb's setData:dataBytes forType:\"com.microsoft.Art--Text-ClipFormat\"", "set ownedCount to (pb's changeCount()) as integer",
 'try','tell application "Microsoft PowerPoint"',`run VB macro macro name ${JSON.stringify(macroName)} list of parameters {${args.map(JSON.stringify).join(',')}}`,'end tell',
 'on error errorText number errorNumber',"if ((pb's changeCount()) as integer) is ownedCount then", "pb's clearContents()", "pb's writeObjects:saved",'end if','error errorText number errorNumber','end try',
 "if ((pb's changeCount()) as integer) is ownedCount then", "pb's clearContents()", "pb's writeObjects:saved",'end if',
]);}
function capture(id,recheck){
 const args=['test','--quiet','--lib','office::macos_offline::tests::live_powerpoint_native_edit_source','--','--ignored','--exact'];
 execFileSync('cargo',args,{cwd:resolve(root,'src-tauri'),encoding:'utf8',timeout:90000,env:{...process.env,VISUALTEX_NATIVE_EDIT_SESSION:id,VISUALTEX_NATIVE_EDIT_RECHECK:recheck?'1':'0',VISUALTEX_NATIVE_EDIT_REPLACEMENT_JSON:resolve(output,'edit-corpus/sources.json')},stdio:['ignore','pipe','pipe']});
 return JSON.parse(readFileSync(resolve(sessions,id,'native-edit-test-result.json'),'utf8'));
}
if(process.argv.includes('--body-probe')){
 nativeClipboard('VTNativeReplaceBodyProbe',['whole']);
 nativeClipboard('VTNativeReplaceBodyProbe',['inner']);
 nativeClipboard('VTNativeReplaceBodyProbe',['delete']);
 nativeClipboard('VTNativeReplaceBodyProbe',['context']);
 process.exit(0);
}
const index=process.argv.indexOf('--case');
const corpus=process.argv.includes('--corpus');
const cases=index<0?(corpus?JSON.parse(readFileSync(resolve(output,'edit-corpus/sources.json'),'utf8')).map(x=>'corpus_'+x.name):['caret','partial','shape','multiple','ambiguous','ordinary','point','miss','merged','merged_inside','display','style_changed']):process.argv[index+1].split(',');
const results=[];
try{
 for(const name of cases){
  const id=randomUUID(); const dir=resolve(sessions,id);mkdirSync(dir,{recursive:true});
  try{
   const sourcePayload=name.startsWith('corpus_')?resolve(output,'edit-corpus',name.slice(7)+'.bin'):payload;
   nativeClipboard('VTNativeEditPrepare',[name,id],sourcePayload);
   const report=readFileSync(resolve(dir,'native-edit-prepare.txt'),'utf8');
   if(report.includes('error='))throw Error(name+': '+report);
   macro('VTNativeAcceptanceSave');
   const slide=report.match(/slide=(\d+)/)?.[1];if(!slide)throw Error('Missing owned slide index');
   const beforeSnapshot=JSON.parse(execFileSync('/usr/bin/python3',[resolve(root,'scripts/powerpoint_native_snapshot.py'),slide],{encoding:'utf8'}));
   if(['ambiguous','ordinary','miss'].includes(name)){
    if(!report.includes('resolved=none')||existsSync(resolve(dir,'request.json')))throw Error(name+' should not resolve any equation');
    results.push({name,report,beforeSnapshot,rejected:true});console.log('PASS '+name+' left ordinary/ambiguous content untouched');continue;
   }
   if(report.includes('resolved=none'))throw Error(name+' did not resolve a native equation');
   const request=JSON.parse(readFileSync(resolve(dir,'request.json'),'utf8'));
   if(request.formulaId!==null||request.encodedMetadata!==null||request.pendingMarker!==null||!request.nativeEquation||request.mode!=='edit')throw Error('Native edit depends on an identity or placeholder');
   const source=capture(id,false);
   if(!source.ok||!source.latex)throw Error(name+': '+JSON.stringify(source));
   const original=readFileSync(resolve(dir,'native-edit-original.omml'),'utf8');
   const check=capture(id,true);if(!check.ok)throw Error('Unchanged source failed recheck: '+JSON.stringify(check));
   if(name==='style_changed'){
    macro('VTNativeEditMutateStyle');const changed=capture(id,true);
    if(changed.ok||!changed.error?.includes('changed'))throw Error('Concurrent native change was not detected');
    results.push({name,report,beforeSnapshot,source,changed});console.log('PASS '+name+' rejected concurrent native edit');continue;
   }
   const dispatch=`protocolVersion=1\nsessionId=${id}\nhost=powerpoint\naction=commit\nmode=edit\nnativeEquation=1\nleadingParagraph=0\ntrailingParagraph=0\ndisplayMode=${request.displayMode}\nfontSizePt=${request.powerPoint.fontSizePt}\n`;
   writeFileSync(resolve(dir,'dispatch.txt'),dispatch);writeFileSync(pointer,id);
   const replacement=resolve(dir,'native-edit-replacement.bin');
   if(!existsSync(replacement))throw Error('Native contextual edit fragment was not produced');
   nativeClipboard('VTNativeAcceptanceApply',[],replacement);
   const applied=readFileSync(resolve(dir,'native-result.txt'),'utf8').trim();if(applied!=='ok')throw Error('Native edit failed: '+applied);
   macro('VTNativeAcceptanceReport',[id]);const first=readFileSync(resolve(dir,'acceptance-report.txt'),'utf8');
   macro('VTNativeAcceptanceApply');macro('VTNativeAcceptanceReport',[id]);
   if(readFileSync(resolve(dir,'acceptance-report.txt'),'utf8')!==first)throw Error('Repeated commit inserted duplicate math');
   results.push({name,report,beforeSnapshot,source,request,original,after:first});console.log('PASS '+name+' -> '+source.latex);
  }finally{
   try{macro('VTCancelPowerPointNativeEquation',[id]);}catch{}
   rmSync(dir,{recursive:true,force:true});
  }
 }
 macro('VTNativeAcceptanceSave');
 const path=resolve(output,'native-edit-integration.json');
 const old=existsSync(path)?JSON.parse(readFileSync(path,'utf8')):[];
 writeFileSync(path,JSON.stringify([...old.filter(r=>!cases.includes(r.name)),...results],null,2));
}finally{if(previous)writeFileSync(pointer,previous);else if(existsSync(pointer))rmSync(pointer);}
console.log(`PASS ${results.length} native editing cases.`);
