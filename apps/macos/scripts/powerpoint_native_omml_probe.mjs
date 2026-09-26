// Isolated native Office Math probe. Never closes or saves user presentations.
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const out = resolve(root, 'build-logs/powerpoint-native-omml');
mkdirSync(out, { recursive: true });
function as(lines) {
  return execFileSync('/usr/bin/osascript', lines.flatMap(x => ['-e', x]), { encoding: 'utf8', timeout: 60000 }).trim();
}
const template = resolve(out, 'NativeOmmlTemplate.pptx');
if (!existsSync(template)) {
  console.log(as([
    'tell application "Microsoft PowerPoint"',
    'set p to make new presentation',
    'if (count of slides of p) is 0 then make new slide at end of p',
    `save p in POSIX file ${JSON.stringify(template)} as save as Open XML presentation`,
    'close p saving no',
    'end tell',
  ]));
}
const probeName = process.argv.includes('--minimal') ? 'NativeOmmlMinimalProbe.pptx' : process.argv.includes('--guards') ? 'NativeOmmlGuardProbe.pptx' : 'NativeOmmlProbe.pptx';
const fixture = resolve(out, probeName);
if (existsSync(fixture) && !process.argv.includes('--resume')) throw new Error('Probe exists; use --resume to inspect the owned fixture.');
if (!process.argv.includes('--resume')) {
  const generator = String.raw`
import sys, zipfile, xml.etree.ElementTree as E
p='http://schemas.openxmlformats.org/presentationml/2006/main'
a='http://schemas.openxmlformats.org/drawingml/2006/main'
m='http://schemas.openxmlformats.org/officeDocument/2006/math'
a14='http://schemas.microsoft.com/office/drawing/2010/main'
for prefix,uri in [('p',p),('a',a),('m',m),('a14',a14)]: E.register_namespace(prefix,uri)
with zipfile.ZipFile(sys.argv[1]) as z: entries={n:z.read(n) for n in z.namelist()}
s=E.fromstring(entries['ppt/slides/slide1.xml']); tree=s.find('.//{'+p+'}spTree')
for node in list(tree)[2:]: tree.remove(node)
def shape(id,name,y,body):
 return E.fromstring(f'''<p:sp xmlns:p="{p}" xmlns:a="{a}" xmlns:m="{m}" xmlns:a14="{a14}"><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="{y}"/><a:ext cx="7315200" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="2400"/></a:pPr>{body}<a:endParaRPr sz="2400"/></a:p></p:txBody></p:sp>''')
r=lambda t:f'<m:r><a:rPr sz="2400"><a:latin typeface="Cambria Math"/></a:rPr><m:t>{t}</m:t></m:r>'
math='<a14:m><m:oMath><m:f><m:fPr/><m:num>'+r('x+1')+'</m:num><m:den>'+r('2')+'</m:den></m:f></m:oMath></a14:m>'
if len(sys.argv)>3 and sys.argv[3]=='guards': math='<a:r><a:rPr sz="2400"/><a:t>A</a:t></a:r>'+math+'<a:r><a:rPr sz="2400"/><a:t>B</a:t></a:r>'
tree.append(shape(2,'NativeMathSource',914400,math))
tree.append(shape(3,'NativeMathTarget',2743200,'<a:r><a:rPr sz="2400"/><a:t>before X after</a:t></a:r>'))
entries['ppt/slides/slide1.xml']=E.tostring(s,encoding='utf-8',xml_declaration=True)
if len(sys.argv)>4 and sys.argv[4]=='minimal':
 entries={
 '[Content_Types].xml':'''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>''',
 '_rels/.rels':'''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>''',
 'ppt/presentation.xml':f'''<p:presentation xmlns:p="{p}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>''',
 'ppt/_rels/presentation.xml.rels':'''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>''',
 'ppt/slides/slide1.xml':entries['ppt/slides/slide1.xml']}
with zipfile.ZipFile(sys.argv[2],'w',zipfile.ZIP_DEFLATED) as z:
 for n,data in entries.items(): z.writestr(n,data)
`;
  execFileSync('/usr/bin/python3', ['-c', generator, template, fixture, process.argv.includes('--guards') ? 'guards' : '', process.argv.includes('--minimal') ? 'minimal' : '']);
  console.log(as(['tell application "Microsoft PowerPoint"', `open POSIX file ${JSON.stringify(fixture)}`, 'activate', 'end tell']));
}
if (process.argv.includes('--paste')) {
  console.log(as([
    'use framework "AppKit"', 'use scripting additions',
    'set pb to current application\'s NSPasteboard\'s generalPasteboard()',
    'set savedItems to current application\'s NSMutableArray\'s array()',
    'repeat with sourceItem in (pb\'s pasteboardItems())',
    'set savedItem to current application\'s NSPasteboardItem\'s alloc()\'s init()',
    'repeat with t in sourceItem\'s types()', 'savedItem\'s setData:(sourceItem\'s dataForType:t) forType:t', 'end repeat',
    'savedItems\'s addObject:savedItem', 'end repeat',
    'tell application "Microsoft PowerPoint"',
    `set p to presentation ${JSON.stringify(probeName)}`,
    'set s to slide 1 of p',
    'set sourceText to text range of text frame of shape "NativeMathSource" of s',
    ...(process.argv.includes('--guards') ? [
      'set sourceLength to text length of sourceText',
      'copy text range (characters 2 thru (sourceLength - 1) of sourceText)',
    ] : ['copy text range sourceText']),
    'set target to text range of text frame of shape "NativeMathTarget" of s',
    'paste text range (character 8 of target)',
    'save p',
    'set resultText to content of text range of text frame of shape "NativeMathTarget" of s',
    'end tell',
    'pb\'s clearContents()', 'pb\'s writeObjects:savedItems',
    'return resultText',
  ]));
  const inspect = String.raw`
import sys,zipfile,xml.etree.ElementTree as E
with zipfile.ZipFile(sys.argv[1]) as z: root=E.fromstring(z.read('ppt/slides/slide1.xml'))
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','m':'http://schemas.openxmlformats.org/officeDocument/2006/math'}
for s in root.findall('.//p:sp',ns):
 print(s.find('p:nvSpPr/p:cNvPr',ns).attrib, 'math=',len(s.findall('.//m:oMath',ns)), E.tostring(s,encoding='unicode'))
`;
  console.log(execFileSync('/usr/bin/python3', ['-c', inspect, fixture], { encoding: 'utf8' }));
}
console.log(`Owned probe: ${fixture}`);
