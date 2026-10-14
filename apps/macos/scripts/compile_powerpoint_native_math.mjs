// Compile the reviewed modules inside the real macOS PowerPoint VBE.
// Scope is one owned acceptance presentation; never quit Office or close users' files.
import { execFileSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const root=resolve(dirname(fileURLToPath(import.meta.url)), '..');
const out=resolve(root,'build-logs/powerpoint-native-omml');
const projectIndex=process.argv.indexOf('--project');
const name=projectIndex>=0?process.argv[projectIndex+1]:'NativeOmmlAcceptance';
if(!['NativeOmmlAcceptance','NativeOmmlRelease'].includes(name)) throw Error('Only the two owned native-math projects are permitted');
const path=resolve(out,name+'.pptm');
const office=resolve(root,'office/macos-offline');
const modules=[
 ['VTOfficePaths','shared/VTOfficePaths.bas'],['VTProtocol','shared/VTProtocol.bas'],
 ['VTMetadata','shared/VTMetadata.bas'],['VTErrorHandling','shared/VTErrorHandling.bas'],
 ['VTLauncher','shared/VTLauncher.bas'],['VTPowerPointNativeTarget','powerpoint/VTPowerPointNativeTarget.cls'],
 ['VTPowerPointNativeMath','powerpoint/VTPowerPointNativeMath.bas'],['VTPowerPointEvents','powerpoint/VTPowerPointEvents.cls'],
 ['VTPowerPointAdapter','powerpoint/VTPowerPointAdapter.bas'],['VTRibbonCallbacks','powerpoint/VTRibbonCallbacks.bas'],
];
function as(lines){return execFileSync('/usr/bin/osascript',lines.flatMap(x=>['-e',x]),{encoding:'utf8',timeout:60000}).trim();}
function ui(lines){return as(['tell application "Microsoft PowerPoint" to activate','tell application "System Events"','tell process "Microsoft PowerPoint"','set frontmost to true',...lines,'end tell','end tell']);}
const ownWindow=[`set w to first window whose name starts with ${JSON.stringify('Microsoft Visual Basic - '+name)}`,'perform action "AXRaise" of w','delay 0.2'];
function selectOwnProject(){
 return ui([...ownWindow,
  'set projectOutline to first UI element of w whose role is "AXOutline"',
  'set matched to false',
  'repeat with r in rows of projectOutline',
  'set rowCell to UI element 1 of r',
  'set rowText to name of every UI element of rowCell as text',
  `if rowText contains ${JSON.stringify(name)} then`,
  'select r','set matched to true','exit repeat','end if','end repeat',
  'if matched is false then error "The owned VBA project is not in the outline"']);
}
function removeModule(moduleName){
 selectOwnProject();
 ui([...ownWindow,'set o to first UI element of w whose role is "AXOutline"',
 'repeat with pass from 1 to 4','set insideOwn to false','set idx to 1','repeat while idx <= count of rows of o',
 'set c to UI element 1 of row idx of o','set t to name of every UI element of c as text',
 `if t contains "VBAProject (" then set insideOwn to (t contains ${JSON.stringify(name)})`,
 'if insideOwn then','try','set d to first UI element of c whose role is "AXDisclosureTriangle"','if value of d is false then click d','end try','end if','set idx to idx + 1','end repeat','end repeat',
 'set insideOwn to false','set found to false','repeat with r in rows of o','set c to UI element 1 of r','set t to name of every UI element of c as text',
 `if t contains "VBAProject (" then set insideOwn to (t contains ${JSON.stringify(name)})`,
 `if insideOwn and t is ${JSON.stringify(moduleName)} then`,'select r','set found to true','exit repeat','end if','end repeat',
 'if found is false then error "The module is not in the owned project"',
 'set removed to false','repeat with mn in {"文件", "File"}','if removed is false then','try',
 'set fm to menu 1 of menu bar item (mn as text) of menu bar 1','repeat with e in menu items of fm','set t to name of e as text',
 `if t is "删除" or t is "Remove" or t starts with ${JSON.stringify('删除 '+moduleName)} or t starts with ${JSON.stringify('Remove '+moduleName)} then`,
 'click e','set removed to true','exit repeat','end if','end repeat','end try','end if','end repeat',
 'if removed is false then error "Cannot remove the owned VBA module"','delay 0.4',
 'repeat with b in buttons of window 1','set t to name of b as text','if t starts with "否" or t starts with "No" then','click b','exit repeat','end if','end repeat','delay 0.3']);
}
function importModule(moduleName,modulePath){
 selectOwnProject();
 ui([...ownWindow,
  'set opened to false','repeat with menuName in {"文件", "File"}',
  'if opened is false then','try',
  'set fm to menu 1 of menu bar item (menuName as text) of menu bar 1',
  'repeat with entry in menu items of fm','set t to name of entry as text',
  'if t starts with "导入文件" or t starts with "Import File" then','click entry','set opened to true','exit repeat','end if','end repeat','end try','end if','end repeat',
  'if opened is false then error "The VBE import command is unavailable"',
  'delay 0.5','keystroke "g" using {command down, shift down}','delay 0.4',
  'if exists text field 1 of sheet 1 of sheet 1 of window 1 then',
  `set value of text field 1 of sheet 1 of sheet 1 of window 1 to ${JSON.stringify(modulePath)}`,
  'else', `set value of text field 1 of sheet 1 of window 1 to ${JSON.stringify(modulePath)}`, 'end if',
  'key code 36','delay 0.6','key code 36','delay 1.1']);
 console.log('Imported '+moduleName);
}
const args=process.argv.slice(2);
if(args.includes('--release-copy')){
 if(name!=='NativeOmmlRelease')throw Error('--release-copy requires the release project');
 if(existsSync(path))throw Error('The release copy already exists; refusing overwrite');
 copyFileSync(resolve(out,'NativeOmmlAcceptance.pptm'),path);
 console.log('Created owned release copy '+path);
}
if(args.includes('--prepare')){
 mkdirSync(out,{recursive:true});
 if(existsSync(path))throw Error('Owned build already exists; refusing to overwrite it.');
 as(['tell application "Microsoft PowerPoint"','set p to make new presentation',
 'if (count of slides of p) is 0 then make new slide at end of p',
 `save p in POSIX file ${JSON.stringify(path)} as save as Open XML presentation macro enabled`,'end tell']);
 ui(['set tm to menu 1 of menu bar item "工具" of menu bar 1','set mm to menu 1 of menu item "宏" of tm','click menu item "Visual Basic 编辑器" of mm','delay 1','keystroke "r" using {command down}']);
 console.log('Prepared '+path);
}
if(args.includes('--vbe')) ui([`perform action "AXRaise" of window ${JSON.stringify(name)}`, 'set tm to menu 1 of menu bar item "工具" of menu bar 1','set mm to menu 1 of menu item "宏" of tm','click menu item "Visual Basic 编辑器" of mm','delay 1','keystroke "r" using {command down}']);
if(args.includes('--outline'))console.log(ui([...ownWindow,'set o to first UI element of w whose role is "AXOutline"','set texts to ""','repeat with r in rows of o','set c to UI element 1 of r','set texts to texts & (name of every UI element of c as text) & linefeed','end repeat','return texts']));
const removeIndex=args.indexOf('--remove');
if(removeIndex>=0) for(const moduleName of args[removeIndex+1].split(',')) removeModule(moduleName);
const replaceIndex=args.indexOf('--replace');
if(replaceIndex>=0){for(const moduleName of args[replaceIndex+1].split(',')){
 const entry=modules.find(x=>x[0]===moduleName);
 if(!entry && moduleName!=='VTNativeMathAcceptance') throw Error('Unknown module '+moduleName);
 removeModule(moduleName); importModule(moduleName,entry?resolve(office,entry[1]):resolve(root,'scripts/VTNativeMathAcceptance.bas'));
}}
const i=args.indexOf('--import');
if(i>=0){
 const names=args[i+1]?.split(',')??[];
 for(const name of names){
  const entry=modules.find(x=>x[0]===name);
  if(name==='VTNativeMathAcceptance') importModule(name,resolve(root,'scripts/VTNativeMathAcceptance.bas'));
  else { if(!entry)throw Error('Unknown module '+name); importModule(name,resolve(office,entry[1])); }
 }
}
if(args.includes('--compile')){
 const result=ui([...ownWindow,'set compiled to false','repeat with menuName in {"调试", "Debug"}','if compiled is false then','try',
 'set dm to menu 1 of menu bar item (menuName as text) of menu bar 1','repeat with entry in menu items of dm','set t to name of entry as text',
 'if t starts with "编译 " or t starts with "Compile " then','if enabled of entry then click entry','set compiled to true','exit repeat','end if','end repeat','end try','end if','end repeat','delay 1',
 'set messageText to ""','repeat with candidate in windows','try','if description of candidate is "警告" or description of candidate is "Warning" then set messageText to messageText & (value of every static text of candidate as text)','end try','end repeat','return messageText']);
 if(result)throw Error('VBE compilation: '+result);
 as(['tell application "Microsoft PowerPoint"',`save presentation ${JSON.stringify(name+'.pptm')}`,'end tell']);
 console.log('VBE compiled and saved '+path);
}
if(args.includes('--verify')){
 const verify=String.raw`
import sys,pathlib,re
from decimal import Decimal
from oletools.olevba import VBA_Parser
p=VBA_Parser(sys.argv[1]);macros={pathlib.Path(n).stem:code for _,_,n,code in p.extract_macros()};p.close()
for filename in sys.argv[2:]:
 name=pathlib.Path(filename).stem
 if name not in macros: raise SystemExit('Missing compiled module '+name)
 source=pathlib.Path(filename).read_text()
 # Keep literal strings exact. Only canonicalize VBE's case/number formatting.
 def normalize(s):
  s=re.sub(r'(?im)^Attribute .*\n?', '', s.replace('\r\n','\n').replace('\r','\n'))
  s=s[s.lower().find('option explicit'):].strip()
  number=re.compile(r'(?<![\w&])((?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)(?:[#%!@&^])?(?![\w])',re.I)
  def num(m):
   v=format(Decimal(m.group(1)),'f')
   return v.rstrip('0').rstrip('.') if '.' in v else v
  parts=re.split(r'("(?:[^"]|"")*")',s)
  return ''.join(p if i%2 else re.sub(r'\s+','',number.sub(num,p.lower())) for i,p in enumerate(parts))
 a,b=normalize(source),normalize(macros[name])
 if a!=b:
  pos=next((i for i,(x,y) in enumerate(zip(a,b)) if x!=y), min(len(a),len(b)))
  raise SystemExit('Compiled source mismatch '+name+' at '+str(pos)+'\nSOURCE: '+repr(a[max(0,pos-70):pos+100])+'\nSAVED: '+repr(b[max(0,pos-70):pos+100]))
 print('Verified '+name)
`;
 console.log(execFileSync('/usr/bin/python3',['-c',verify,path,...modules.map(x=>resolve(office,x[1]))],{encoding:'utf8'}));
}
