// Inspect only the mathematical selection in our own fixture. Restores clipboard.
import {execFileSync} from 'node:child_process';
import {mkdirSync,readFileSync} from 'node:fs';
import {dirname,resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
const out=resolve(dirname(fileURLToPath(import.meta.url)),'../build-logs/powerpoint-native-omml/clipboard');mkdirSync(out,{recursive:true});
if(process.argv.includes('--package-variants')){
 console.log(execFileSync('/usr/bin/python3',['-c',String.raw`import zipfile,pathlib,sys,xml.etree.ElementTree as E
out=pathlib.Path(sys.argv[1]);a=out/'0.bin';b=out.parent/'native-inline.bin'
with zipfile.ZipFile(a) as z:original={n:z.read(n) for n in z.namelist()}
with zipfile.ZipFile(b) as z:generated={n:z.read(n) for n in z.namelist()}
def save(name,data,compression=zipfile.ZIP_DEFLATED):
 p=out/(name+'.bin')
 with zipfile.ZipFile(p,'w',compression) as z:
  for n,d in data.items():z.writestr(n,d)
 print(p)
save('captured-stored',original,zipfile.ZIP_STORED)
save('generated-deflated',generated)
for name,part in [('generated-drawing','clipboard/drawings/drawing1.xml'),('generated-theme','clipboard/theme/theme1.xml')]:
 data=dict(original);data[part]=generated[part];save(name,data)
ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','m':'http://schemas.openxmlformats.org/officeDocument/2006/math','a14':'http://schemas.microsoft.com/office/drawing/2010/main','lc':'http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas'}
for p,u in ns.items():E.register_namespace(p,u)
for name,xpath in [('generated-textbody','.//a:txBody'),('generated-math','.//m:oMath'),('empty-liststyle','.//a:lstStyle')]:
 d='clipboard/drawings/drawing1.xml';tree=E.fromstring(original[d]);gt=E.fromstring(generated[d]);old=tree.find(xpath,ns);new=gt.find(xpath,ns)
 parent=next(p for p in tree.iter() if old in list(p));idx=list(parent).index(old);parent.remove(old);parent.insert(idx,new)
 data=dict(original);data[d]=E.tostring(tree,encoding='utf-8',xml_declaration=True);save(name,data)
`,out],{encoding:'utf8'})); process.exit(0);
}
const lines=['use framework "AppKit"','use scripting additions',
"set pb to current application's NSPasteboard's generalPasteboard()","set saved to current application's NSMutableArray's array()",
"repeat with sourceItem in pb's pasteboardItems()","set newItem to current application's NSPasteboardItem's alloc()'s init()","repeat with t in sourceItem's types()","newItem's setData:(sourceItem's dataForType:t) forType:t",'end repeat',"saved's addObject:newItem",'end repeat',
 'tell application "Microsoft PowerPoint"','set p to presentation "NativeOmmlGuardProbe.pptx"','set tr to text range of text frame of shape "NativeMathSource" of slide 1 of p','copy text range (characters 2 thru ((text length of tr) - 1) of tr)','end tell',
 'set resultText to ""','set i to 0',"repeat with t in pb's types()", "set dataBytes to pb's dataForType:t", `set dest to ${JSON.stringify(out+'/')} & (i as text) & ".bin"`, "dataBytes's writeToFile:dest atomically:true",
 'set resultText to resultText & (i as text) & "|" & (t as text) & "|" & ((dataBytes\'s |length|()) as text) & linefeed','set i to i + 1','end repeat',"pb's clearContents()","pb's writeObjects:saved",'return resultText'];
console.log(execFileSync('/usr/bin/osascript',lines.flatMap(l=>['-e',l]),{encoding:'utf8',timeout:20000}));
console.log(execFileSync('/usr/bin/python3',['-c',String.raw`import pathlib,sys,zipfile
for p in pathlib.Path(sys.argv[1]).glob('*.bin'):
 b=p.read_bytes();print(p.name,len(b),repr(b[:200]))
 if zipfile.is_zipfile(p):
  with zipfile.ZipFile(p) as z:print(z.namelist())
`,out],{encoding:'utf8'}));
