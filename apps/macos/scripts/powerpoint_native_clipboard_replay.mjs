// Tests independently replayable PowerPoint native text data, without an OLE
// source handle, live source document, picture preview or a staging window.
import {execFileSync} from 'node:child_process';
import {resolve,dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const arg=process.argv.indexOf('--payload');
const payload=arg>=0?resolve(process.argv[arg+1]):resolve(root,'build-logs/powerpoint-native-omml/clipboard/0.bin');
const lines=['use framework "AppKit"','use scripting additions',
"set pb to current application's NSPasteboard's generalPasteboard()","set saved to current application's NSMutableArray's array()",
"repeat with sourceItem in pb's pasteboardItems()","set newItem to current application's NSPasteboardItem's alloc()'s init()","repeat with t in sourceItem's types()","newItem's setData:(sourceItem's dataForType:t) forType:t",'end repeat',"saved's addObject:newItem",'end repeat',
`set payload to current application's NSData's dataWithContentsOfFile:${JSON.stringify(payload)}`,
"pb's clearContents()","pb's setData:payload forType:\"com.microsoft.Art--Text-ClipFormat\"",
 'try','tell application "Microsoft PowerPoint"','set p to presentation "NativeOmmlGuardProbe.pptx"','set tr to text range of text frame of shape "NativeMathTarget" of slide 1 of p','set content of tr to "before X after"','paste text range (character 8 of tr)','save p','set resultText to content of text range of text frame of shape "NativeMathTarget" of slide 1 of p','end tell',
 'on error messageText number errorNumber',"pb's clearContents()","pb's writeObjects:saved",'error messageText number errorNumber','end try',"pb's clearContents()","pb's writeObjects:saved",'return resultText'];
console.log(execFileSync('/usr/bin/osascript',lines.flatMap(l=>['-e',l]),{encoding:'utf8',timeout:25000}));
console.log(execFileSync('/usr/bin/python3',['-c',String.raw`import zipfile,sys,xml.etree.ElementTree as E
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','m':'http://schemas.openxmlformats.org/officeDocument/2006/math'}
with zipfile.ZipFile(sys.argv[1]) as z:root=E.fromstring(z.read('ppt/slides/slide1.xml'))
for s in root.findall('.//p:sp',ns):
 if s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='NativeMathTarget' and s.findall('.//m:oMath',ns):
  print('math',len(s.findall('.//m:oMath',ns)),'paragraphs',len(s.findall('p:txBody/a:p',ns)))
  if '--verbose' in sys.argv: print(E.tostring(s.find('p:txBody',ns),encoding='unicode'))
`,resolve(root,'build-logs/powerpoint-native-omml/NativeOmmlGuardProbe.pptx')],{encoding:'utf8'}));
