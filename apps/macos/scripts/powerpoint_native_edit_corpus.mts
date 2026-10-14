// Source fixtures use the same renderer as the real Office editor.
import { DOMParser } from '@xmldom/xmldom';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { renderOfficeFormulaArtifacts } from '../src/office/shared/formulaRenderArtifacts';
import { normalizeFormulaEditorDocument, serializeFormulaEditorDocument } from '../src/office/shared/formulaEditorDocument';
(globalThis as any).DOMParser ??= DOMParser;
const document = new DOMParser().parseFromString('<root/>', 'application/xml');
const docProto=Object.getPrototypeOf(document);
const elementProto=Object.getPrototypeOf(document.documentElement);
if (!docProto.querySelector) docProto.querySelector=function(name:string){return this.getElementsByTagName(name)?.item(0)??null;};
if (!('children' in elementProto)) Object.defineProperty(elementProto,'children',{configurable:true,get(){return Array.from(this.childNodes??[]).filter((n:any)=>n.nodeType===1);}});
const outputRoot=resolve(dirname(fileURLToPath(import.meta.url)),'../build-logs/powerpoint-native-omml');
if(process.argv.includes('--roundtrip')){
 const records=JSON.parse(readFileSync(resolve(outputRoot,'native-edit-integration.json'),'utf8')).filter((x:any)=>x.name.startsWith('corpus_'));
 const exported=[];
 for(const record of records){
  const latex=record.source.latex;
  try {
   const artifact=renderOfficeFormulaArtifacts({lines:[{id:'roundtrip',latex}],codeFormat:'raw',displayMode:record.source.displayMode,host:'powerpoint',includeWordOmml:true,fontSizePt:24,formulaLetterFont:record.source.font});
   exported.push({name:record.name.slice(7),latex,ommlBase64:artifact.omml!.ommlBase64});
   console.log('Roundtrip rendered '+record.name);
  } catch(error){console.error('Roundtrip FAILED '+record.name+': '+String(error));console.error(JSON.stringify(normalizeFormulaEditorDocument([{id:'x',latex}], 'raw')));console.error(serializeFormulaEditorDocument(normalizeFormulaEditorDocument([{id:'x',latex}], 'raw')));process.exitCode=1;}
 }
 writeFileSync(resolve(outputRoot,'edit-corpus/roundtrip.json'),JSON.stringify(exported,null,2));
 process.exit(process.exitCode??0);
}
const samples=[
 ['replacement',String.raw`\frac{y^2+3}{5}`],
 ['root',String.raw`\sqrt[3]{x^2+1}`],
 ['sum',String.raw`\sum_{i=1}^{n}a_i^2`],
 ['matrix',String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}`],
 ['cases',String.raw`\begin{cases}x^2 & x>0\\ -x & x\le 0\end{cases}`],
 ['accents',String.raw`\vec{x}+\hat{y}+\overline{z}+\underline{w}`],
 ['styles',String.raw`\mathbf{x}+\mathbb{R}+\mathcal{L}+\mathrm{d}`],
 ['integral',String.raw`\int_0^\infty e^{-x^2}\,\mathrm{d}x`],
 ['functions',String.raw`\sin x+\lim_{n\to\infty}a_n`],
 ['prescripts',String.raw`{}^{14}_{6}\mathrm{C}`],
 ['chinese',String.raw`\text{速度}=\frac{\text{距离}}{\text{时间}}`],
 ['relations',String.raw`\alpha+\beta\le\gamma`],
 ['aligned',String.raw`\begin{aligned}a&=b+c\\d&=e\end{aligned}`],
] as const;
const results=samples.map(([name,latex])=>{
 console.log('Rendering '+name);
 const lines=name==='aligned'?[{id:'row-1',latex:'a=b+c'},{id:'row-2',latex:'d=e'}]:[{id:'fixture-line',latex}];
 const artifacts=renderOfficeFormulaArtifacts({lines,codeFormat:name==='aligned'?'align':'raw',displayMode:name==='aligned'?'block':'inline',host:'powerpoint',includeWordOmml:true,fontSizePt:24});
 if(!artifacts.omml)throw Error('Missing OMML for '+name);
 return {name,latex,ommlBase64:artifacts.omml.ommlBase64};
});
const out=resolve(dirname(fileURLToPath(import.meta.url)),'../build-logs/powerpoint-native-omml/edit-corpus');
mkdirSync(out,{recursive:true});
writeFileSync(resolve(out,'sources.json'),JSON.stringify(results,null,2));
console.log('Rendered '+results.length+' native edit fixtures at '+out);
