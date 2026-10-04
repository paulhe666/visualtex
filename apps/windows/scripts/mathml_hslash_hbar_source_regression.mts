import assert from "node:assert/strict";
import {readFileSync,statSync} from "node:fs";
import {spawnSync} from "node:child_process";
import {fileURLToPath} from "node:url";
import {createHash} from "node:crypto";
import {DOMParser} from "@xmldom/xmldom";
import {latexToMathMl,latexToSvg} from "../src/export/runtime.ts";

// Source-only U+210F recovery test. argv[2] labels baseline or candidate.
// JSON goes to stdout; callers retain it with workspace write, not shell redirection.
const repo=new URL("../../../",import.meta.url);
const mode=process.argv[2]||"candidate";
assert.ok(["baseline","candidate"].includes(mode));
const converter=new URL("apps/windows/src-windows/VisualTeX.WindowsOffice.VstoShared/MathMlToLatexConverter.cs",repo);
const probe=new URL("artifacts/conversion-symbol-audit-20261002/spacing-source-probe/bin/Debug/net9.0/SpacingSourceProbe.dll",repo);
const hash=(x:string|Buffer)=>createHash("sha256").update(x).digest("hex");
const sourceHash=hash(readFileSync(converter)),probeHash=hash(readFileSync(probe));
if(mode==="candidate") assert.ok(statSync(probe).mtimeMs>=statSync(converter).mtimeMs,"Rebuild isolated C# harness after converter release");
const styles=["normal","italic","bold-italic"] as const;
type Style=(typeof styles)[number];
type Item={id:string;alternate:boolean;style:Style;mathMl:string;reference:string};
const wrap=(command:string,style:Style)=>"\\"+({normal:"mathrm",italic:"mathit","bold-italic":"mathbfit"}[style])+"{\\"+command+"}";
const cases:Item[]=[];
for(const alternate of [false,true]){
 const command=alternate?"hbar":"hslash",alt=alternate?' data-mjx-alternate="1"':"";
 for(const style of styles){
  cases.push({id:command+"-direct-"+style,alternate,style,mathMl:'<math><mi mathvariant="'+style+'"'+alt+'>ℏ</mi></math>',reference:wrap(command,style)});
  cases.push({id:command+"-inherited-"+style,alternate,style,mathMl:'<math><mstyle mathvariant="'+style+'"><mi'+alt+'>ℏ</mi></mstyle></math>',reference:wrap(command,style)});
 }
 cases.push({id:command+"-subscript-italic",alternate,style:"italic",mathMl:'<math><msub><mi mathvariant="italic"'+alt+'>&#x210F;</mi><mi>i</mi></msub></math>',reference:wrap(command,"italic")+"_{i}"});
 cases.push({id:command+"-superscript-normal",alternate,style:"normal",mathMl:'<math><msup><mi mathvariant="normal"'+alt+'>&#x210F;</mi><mn>2</mn></msup></math>',reference:wrap(command,"normal")+"^{2}"});
 cases.push({id:command+"-implicit-italic",alternate,style:"italic",mathMl:'<math><mi'+alt+'>ℏ</mi></math>',reference:"\\"+command});
}
cases.push({id:"hslash-container-alternate-normal",alternate:false,style:"normal",mathMl:'<math data-mjx-alternate="1"><mi mathvariant="normal">ℏ</mi></math>',reference:wrap("hslash","normal")});
cases.push({id:"hslash-container-and-token-zero-italic",alternate:false,style:"italic",mathMl:'<math><mstyle data-mjx-alternate="1" mathvariant="italic"><mi data-mjx-alternate="0">ℏ</mi></mstyle></math>',reference:wrap("hslash","italic")});
assert.equal(cases.length,20);
const run=spawnSync("dotnet",[fileURLToPath(probe)],{input:cases.map(c=>JSON.stringify({mathMl:c.mathMl})).join("\n")+"\n",encoding:"utf8",maxBuffer:1024*1024});
assert.equal(run.status,0,run.stderr);
const recovered=run.stdout.trim().split(/\r?\n/).map(s=>JSON.parse(s).latex);
assert.equal(recovered.length,20);
const parse=(xml:string)=>new DOMParser().parseFromString(xml,"application/xml");
function tokens(xml:string){
 return Array.from(parse(xml).getElementsByTagName("*")).filter(n=>["mi","mo","mn","mtext"].includes(n.tagName)).map(n=>{
  let declared="";
  for(let p:any=n;p?.nodeType===1;p=p.parentNode){declared=p.getAttribute("mathvariant")||"";if(declared)break;}
  const text=n.textContent||"";
  return {role:n.tagName,text,scalars:Array.from(text,c=>c.codePointAt(0)!),style:declared||(n.tagName==="mi"&&Array.from(text).length===1?"italic":"normal"),alternate:n.getAttribute("data-mjx-alternate")==="1"};
 });
}
function visual(rendered:Awaited<ReturnType<typeof latexToSvg>>){
 const doc=parse(rendered.svg),paths=new Map(Array.from(doc.getElementsByTagName("path")).map(n=>[n.getAttribute("id"),n.getAttribute("d")]));
 const glyphs=Array.from(doc.getElementsByTagName("use")).map(n=>{
  const path=paths.get((n.getAttribute("xlink:href")||n.getAttribute("href")||"").slice(1));
  assert.ok(path,"Every actual SVG use must resolve to a glyph path");
  const transforms:string[]=[];
  for(let p:any=n;p?.nodeType===1;p=p.parentNode)if(p.getAttribute("transform"))transforms.unshift(p.getAttribute("transform"));
  return {scalar:n.getAttribute("data-c"),outline:hash(path!),transforms,x:n.getAttribute("x")||"",y:n.getAttribute("y")||""};
 });
 assert.ok(glyphs.length);
 return {width:rendered.width,height:rendered.height,baseline:rendered.baseline,viewBox:doc.documentElement.getAttribute("viewBox"),glyphs};
}
const eq=(a:unknown,b:unknown)=>JSON.stringify(a)===JSON.stringify(b);
const rows:any[]=[];
for(let i=0;i<cases.length;i++){
 const item=cases[i],latex=recovered[i],checks:Record<string,boolean>={},errors:string[]=[];
 let actualMathMl="",referenceMathMl="",actual:any,reference:any,actualSvg:any,referenceSvg:any;
 try{
  const source=tokens(item.mathMl),expectedSource=source.find(x=>x.text==="ℏ");
  actualMathMl=await latexToMathMl(latex,false);
  referenceMathMl=await latexToMathMl(item.reference,false);
  actual=tokens(actualMathMl);reference=tokens(referenceMathMl);
  const target=actual.find((x:any)=>x.text==="ℏ");
  checks.scalar=eq(target?.scalars,[0x210F]);
  checks.miRole=target?.role==="mi";
  checks.effectiveStyle=target?.style===item.style;
  checks.alternate=target?.alternate===item.alternate;
  checks.sourceDeclaredContract=expectedSource?.style===item.style&&expectedSource?.alternate===item.alternate;
  checks.allTokensMatchCanonical=eq(actual,reference);
  checks.commandIdentity=new RegExp("\\\\"+(item.alternate?"hbar":"hslash")+"(?:[^A-Za-z]|$)").test(latex);
  const scriptCounts=(xml:string)=>["msub","msup"].map(tag=>parse(xml).getElementsByTagName(tag).length);
  checks.scriptStructure=eq(scriptCounts(item.mathMl),scriptCounts(actualMathMl));
  actualSvg=visual(await latexToSvg(latex));referenceSvg=visual(await latexToSvg(item.reference));
  checks.canonicalSvg=eq(actualSvg,referenceSvg);
 }catch(error){errors.push(error instanceof Error?error.message:String(error));}
 errors.push(...Object.entries(checks).filter(([,ok])=>!ok).map(([key])=>key));
 rows.push({...item,referenceLatex:item.reference,latex,actualMathMl,referenceMathMl,actual,reference,actualSvg,referenceSvg,checks,errors,passed:errors.length===0});
}
const pairControls=styles.map(style=>{
 const ordinary=rows.find(r=>r.id==="hslash-direct-"+style),alternate=rows.find(r=>r.id==="hbar-direct-"+style);
 return {style,canonicalOutlinesDistinct:!eq(ordinary.referenceSvg?.glyphs.map((g:any)=>g.outline),alternate.referenceSvg?.glyphs.map((g:any)=>g.outline)),actualOutlinesDistinct:!eq(ordinary.actualSvg?.glyphs.map((g:any)=>g.outline),alternate.actualSvg?.glyphs.map((g:any)=>g.outline)),styleUnchanged:ordinary.actual?.[0]?.style===style&&alternate.actual?.[0]?.style===style};
});
const intrinsicControls=["hslash","hbar"].map(command=>{
 const normal=rows.find(r=>r.id===command+"-direct-normal"),italic=rows.find(r=>r.id===command+"-direct-italic");
 return {command,canonicalNormalItalicSameIntrinsicOutline:eq(normal.referenceSvg?.glyphs.map((g:any)=>g.outline),italic.referenceSvg?.glyphs.map((g:any)=>g.outline)),actualNormalItalicSameIntrinsicOutline:eq(normal.actualSvg?.glyphs.map((g:any)=>g.outline),italic.actualSvg?.glyphs.map((g:any)=>g.outline)),canonicalFullSvgEqual:eq(normal.referenceSvg,italic.referenceSvg),actualFullSvgEqual:eq(normal.actualSvg,italic.actualSvg)};
});
assert.equal(hash(readFileSync(probe)),probeHash,"Harness changed during test");
if(mode==="candidate")assert.equal(hash(readFileSync(converter)),sourceHash,"Converter changed during test");
const checkPassCounts=Object.fromEntries(Object.keys(rows[0].checks).map(key=>[key,rows.filter(r=>r.checks[key]===true).length]));
const summary={total:20,checkPassCounts,passed:rows.filter(r=>r.passed).length,failed:rows.filter(r=>!r.passed).length,ordinaryFlippedToAlternate:rows.filter(r=>!r.alternate&&r.actual?.find((t:any)=>t.text==="ℏ")?.alternate===true).map(r=>r.id),failedCases:rows.filter(r=>!r.passed).map(r=>({id:r.id,latex:r.latex,errors:r.errors}))};
const report={generatedAt:new Date().toISOString(),mode,sourceHashAtRun:sourceHash,probeSha256:probeHash,probeModifiedAt:statSync(probe).mtime.toISOString(),frontendSha256:hash(readFileSync(new URL("apps/windows/src/export/runtime.ts",repo))),summary,pairControls,intrinsicControls,scope:"Owned source-only U+210F corpus, C# MathMlToLatexConverter to production latexToMathMl and latexToSvg. Checks effective style, mi/scalar/alternate, named source identity, inherited and scripted controls, resolved actual SVG outlines and placement. Normal and italic share intrinsic artwork; no false visual-difference requirement. Baseline label uses a previously built source-only harness and does not assert it is the current product binary. No native, Word, screenshot or production edits by this test.",rows};
console.log(JSON.stringify(process.argv[3]==="--summary"?{...report,rows:undefined}:report));
if(summary.failed||pairControls.some(x=>!x.actualOutlinesDistinct||!x.styleUnchanged))process.exitCode=1;
