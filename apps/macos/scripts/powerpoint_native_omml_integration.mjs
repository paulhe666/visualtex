// Real PowerPoint regression of the production native-math insertion callbacks.
// Operates exclusively on the owned NativeOmmlAcceptance.pptm and temporary sessions.
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync, rmSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { homedir } from 'node:os';
import { randomUUID } from 'node:crypto';
import { fileURLToPath } from 'node:url';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const out=resolve(root,'build-logs/powerpoint-native-omml');
const sessionRoot=resolve(homedir(),'Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXRuntime/OfficeSessions');
const pointer=resolve(sessionRoot,'powerpoint-active-session.txt');
const originalPointer=existsSync(pointer)?readFileSync(pointer):null;
function as(lines){return execFileSync('/usr/bin/osascript',lines.flatMap(x=>['-e',x]),{encoding:'utf8',timeout:25000}).trim();}
function macro(name,args=[]){return as(['tell application "Microsoft PowerPoint"',`run VB macro macro name ${JSON.stringify(name)} list of parameters {${args.map(x=>JSON.stringify(x)).join(',')}}`,'end tell']);}
function apply(payload){return as([
 'use framework "AppKit"','use scripting additions',
 `set payload to current application's NSData's dataWithContentsOfFile:${JSON.stringify(payload)}`,
 'if payload is missing value then error "Missing native clipboard fixture"',
 "set pb to current application's NSPasteboard's generalPasteboard()", "set saved to current application's NSMutableArray's array()",
 "repeat with sourceItem in pb's pasteboardItems()", "set newItem to current application's NSPasteboardItem's alloc()'s init()",
 "repeat with t in sourceItem's types()", "newItem's setData:(sourceItem's dataForType:t) forType:t",'end repeat',"saved's addObject:newItem",'end repeat',
 "pb's clearContents()", 'set didSet to pb\'s setData:payload forType:"com.microsoft.Art--Text-ClipFormat"',
 "set ownedCount to (pb's changeCount()) as integer",
 'try','if didSet is false then error "Cannot set the native clipboard"','tell application "Microsoft PowerPoint"',
 'run VB macro macro name "VTNativeAcceptanceApply" list of parameters {}','end tell',
 'on error messageText number errorNumber',"if ((pb's changeCount()) as integer) is ownedCount then", "pb's clearContents()","pb's writeObjects:saved",'end if','error messageText number errorNumber','end try',
 "if ((pb's changeCount()) as integer) is ownedCount then", "pb's clearContents()","pb's writeObjects:saved",'end if',
]);}
function captureContext(destination){return as([
 'use framework "AppKit"','use scripting additions',
 "set pb to current application's NSPasteboard's generalPasteboard()", "set saved to current application's NSMutableArray's array()",
 "repeat with sourceItem in pb's pasteboardItems()", "set newItem to current application's NSPasteboardItem's alloc()'s init()",
 "repeat with t in sourceItem's types()", "newItem's setData:(sourceItem's dataForType:t) forType:t",'end repeat',"saved's addObject:newItem",'end repeat',
 "set ownedCount to (pb's changeCount()) as integer",'try',
 'tell application "Microsoft PowerPoint" to run VB macro macro name "VisualTeX_CopyNativeParagraphContext" list of parameters {}',
 "set ownedCount to (pb's changeCount()) as integer",'set payload to pb\'s dataForType:"com.microsoft.Art--GVML-ClipFormat"',
 'if payload is missing value then error "Missing native paragraph data"',`set wrote to payload's writeToFile:${JSON.stringify(destination)} atomically:true`,
 'if wrote is false then error "Cannot preserve native paragraph"',
 'on error messageText number errorNumber',"if ((pb's changeCount()) as integer) is ownedCount then", "pb's clearContents()","pb's writeObjects:saved",'end if','error messageText number errorNumber','end try',
 "if ((pb's changeCount()) as integer) is ownedCount then", "pb's clearContents()","pb's writeObjects:saved",'end if',
]);}
function slideSummary(id){
 macro('VTNativeAcceptanceReport',[id]);
 return readFileSync(resolve(sessionRoot,id,'acceptance-report.txt'),'utf8');
}
const allCases=['inline_mid','inline_caret','inline_start','inline_end','inline_empty','unicode','multiline_replace','font_size','format_preservation','display_middle','display_existing','display_start','display_end','display_empty','display_bullets','standalone_inline','standalone_display','stale','cancel'];
const selected=process.argv.indexOf('--case');
const cases=selected>=0?process.argv[selected+1].split(','):allCases;
const results=[];
try{
 for(const name of cases){
  if(!allCases.includes(name)) throw Error('Unknown test case '+name);
  const id=randomUUID(); const dir=resolve(sessionRoot,id);mkdirSync(dir,{recursive:true});
  const display=name.startsWith('display_')||name==='standalone_display';
  const leading=display && ['display_middle','display_end','display_bullets'].includes(name);
  const trailing=display && ['display_middle','display_start','display_bullets'].includes(name);
  const variant=display?`display${leading?1:0}${trailing?1:0}`:(name==='font_size'?'inline36':'inline');
  let payload=resolve(out,`native-${variant}.bin`);
  const dispatch=`protocolVersion=1\nsessionId=${id}\nhost=powerpoint\naction=${name==='cancel'?'cancel':'commit'}\nmode=create\nnativeEquation=1\nleadingParagraph=${leading?1:0}\ntrailingParagraph=${trailing?1:0}\ndisplayMode=${display?'block':'inline'}\nfontSizePt=${name==='font_size'?36:24}\n`;
  writeFileSync(resolve(dir,'dispatch.txt'),dispatch);writeFileSync(pointer,id);
  try{
   macro('VTNativeAcceptancePrepareCase',[name,id]);
   if(existsSync(resolve(dir,'native-prepare.txt'))){const p=readFileSync(resolve(dir,'native-prepare.txt'),'utf8');if(p.trim()!=='ok')throw Error('Prepare '+name+': '+p);}
   const before=slideSummary(id);
   if(leading){
    const context=resolve(out,`paragraph-${name}.bin`);
    captureContext(context);
    if(readFileSync(resolve(dir,'native-context-status.txt'),'utf8').trim()!=='ok')throw Error('Cannot capture original paragraph context');
    payload=resolve(out,`native-styled-${name}.bin`);
    execFileSync('cargo',['test','--quiet','--lib','office::powerpoint_omml::tests::style_owned_ui_fixture_when_requested','--','--exact'],{
     cwd:resolve(root,'src-tauri'),encoding:'utf8',timeout:90000,
     env:{...process.env,VISUALTEX_PPT_CONTEXT_FIXTURE_IN:context,VISUALTEX_PPT_CONTEXT_FIXTURE_OUT:payload,VISUALTEX_PPT_CONTEXT_TRAILING:trailing?'1':'0',VISUALTEX_PPT_CONTEXT_PARAGRAPH:readFileSync(resolve(dir,'native-context-paragraph.txt'),'utf8').trim()},
     stdio:['ignore','pipe','pipe'],
    });
   }
   if(name==='cancel') macro('VTNativeAcceptanceApply'); else apply(payload);
   const result=readFileSync(resolve(dir,'native-result.txt'),'utf8').trim();
   if(name==='stale') {if(!result.startsWith('error\n')||!result.includes('changed while'))throw Error('Stale target result: '+result);}
   else if(result!=='ok')throw Error(name+': '+result);
   const after=slideSummary(id);
   if(name==='stale'||name==='cancel'){if(after!==before)throw Error(name+' mutated the original text');}
   else {
    // Second callback must be a no-op even after restoring the user's clipboard.
    macro('VTNativeAcceptanceApply');
    if(slideSummary(id)!==after)throw Error(name+' duplicated or changed the formula on retry');
    if(readFileSync(resolve(dir,'native-result.txt'),'utf8').trim()!=='ok')throw Error(name+' idempotent callback failed');
   }
   const [slideIndex,shapeCount,...text]=after.split('|');
   if(Number(shapeCount)!==1)throw Error(name+' created an unexpected number of shapes: '+after);
   console.log(name+': '+JSON.stringify(after));results.push({name,slideIndex:Number(slideIndex),shapeCount:Number(shapeCount),text:text.join('|'),result,before});
  }finally{
   try{macro('VTCancelPowerPointNativeEquation',[id]);}catch{}
   rmSync(dir,{recursive:true,force:true});
  }
 }
 macro('VTNativeAcceptanceSave');
 const resultPath=resolve(out,'native-integration-results.json');
 const previous=existsSync(resultPath)?JSON.parse(readFileSync(resultPath,'utf8')):[];
 writeFileSync(resultPath,JSON.stringify([...previous.filter(r=>!cases.includes(r.name)),...results],null,2));
}finally{
 if(originalPointer!==null)writeFileSync(pointer,originalPointer);else if(existsSync(pointer))rmSync(pointer);
}
console.log(`Completed ${results.length} real PowerPoint callback cases.`);
