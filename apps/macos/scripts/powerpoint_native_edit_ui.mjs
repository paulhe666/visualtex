// Real editor UI actions restricted to our owned PowerPoint acceptance session.
import {execFileSync} from 'node:child_process';
import {readFileSync,readdirSync,statSync,writeFileSync} from 'node:fs';
import {homedir} from 'node:os';
import {resolve} from 'node:path';
const root=resolve(import.meta.dirname,'..');
const sessions=resolve(homedir(),'Library/Application Support/com.visualtex.studio/office/sessions');
const recent=readdirSync(sessions).map(id=>resolve(sessions,id,'session.json')).filter(p=>{try{return statSync(p).isFile()}catch{return false}}).sort((a,b)=>statSync(b).mtimeMs-statSync(a).mtimeMs);
const state=recent.map(p=>JSON.parse(readFileSync(p,'utf8'))).find(s=>s.host==='powerpoint'&&s.nativeEquation&&['editing','created'].includes(s.status));
if(!state||!state.sourceDocumentId?.includes('/visualtexmac/apps/macos/build-logs/powerpoint-native-omml/NativeOmmlAcceptance.pptm'))throw Error('No owned native PowerPoint UI session is active');
const command=process.argv[2];
const x=Number(process.argv[3]),y=Number(process.argv[4]);
if(!['replace-source','click'].includes(command)||!Number.isFinite(x)||!Number.isFinite(y))throw Error('Usage: ... replace-source|click x y');
const lines=['use framework "AppKit"','use scripting additions',
 "set pb to current application's NSPasteboard's generalPasteboard()", "set saved to current application's NSMutableArray's array()",
 "repeat with sourceItem in pb's pasteboardItems()", "set newItem to current application's NSPasteboardItem's alloc()'s init()",
 "repeat with t in sourceItem's types()", "newItem's setData:(sourceItem's dataForType:t) forType:t",'end repeat',"saved's addObject:newItem",'end repeat',
 'tell application "System Events" to tell process "visualtex"', 'set frontmost to true',`click at {${x}, ${y}}`,'delay 0.2',
 ...(command==='replace-source'?['keystroke "a" using {command down}']:[]),'end tell',
 ...(command==='replace-source'?["pb's clearContents()",`pb's setString:${JSON.stringify('$$\n\\frac{y^2+3}{5}\n$$'.replaceAll('\\n','\n'))} forType:"public.utf8-plain-text"`,'tell application "System Events" to tell process "visualtex"','keystroke "v" using {command down}','delay 0.6','end tell',"pb's clearContents()", "pb's writeObjects:saved"]:[]),
];
execFileSync('/usr/bin/osascript',lines.flatMap(l=>['-e',l]),{encoding:'utf8',timeout:20000});
console.log('Operated on owned session '+state.id+' via '+command);
writeFileSync(resolve(root,'build-logs/powerpoint-native-omml/native-ui-session.json'),JSON.stringify({id:state.id,sourceDocumentId:state.sourceDocumentId,command},null,2));
