import assert from "node:assert/strict";
import {readFileSync,statSync} from "node:fs";
import {spawnSync} from "node:child_process";
import {fileURLToPath} from "node:url";
import {resolve} from "node:path";
import {createHash} from "node:crypto";
import {DOMParser} from "@xmldom/xmldom";
import {latexToMathMl,latexToSvg} from "../src/export/runtime.ts";

// Run from the repository root:
// node --import ./apps/windows/node_modules/tsx/dist/loader.mjs apps/windows/scripts/mathml_subset_four_source_regression.mts <cases.json>
// This test writes only JSON to stdout. The caller may retain that JSON as an artifact.
// C# recovery uses an isolated source-only harness; no Word or native renderer is invoked.
type Variant="normal"|"italic"|"bold"|"bold-italic";
type Case={id:string;command:string;alternate:boolean;variant:Variant;source:string;decoded:string;checks:Record<string,boolean>;expected:{scalar:number;tokenClass?:string;style?:string;alternate?:boolean};actualAttributes?:unknown;rootHex:string};
type Token={role:string;text:string;scalars:number[];alternate:boolean;style:string};
type Glyph={code:string;outline:string;transforms:string[]};
type Row={id:string;command:string;variant:Variant;alternate:boolean;latex:string;referenceLatex:string;passed:boolean;checks:Record<string,boolean>;errors:string[];source?:Token;decoded?:Token;exported?:Token;reference?:Token;glyphs?:Glyph[];referenceGlyphs?:Glyph[];geometry?:unknown;referenceGeometry?:unknown;mathMl?:string;referenceMathMl?:string;rootHex:string};
const repo=new URL("../../../",import.meta.url);
const audit=new URL("artifacts/conversion-symbol-audit-20261002/",repo);
assert.ok(process.argv[2],"Expected a 32-case JSON path as argv[2]");
const inputPath=resolve(process.argv[2]);
const inputText=readFileSync(inputPath,"utf8").replace(/^\uFEFF/,"");
const cases=JSON.parse(inputText) as Case[];
const variants:Variant[]=["normal","italic","bold","bold-italic"];
const families=[
 {ordinary:"subsetneqq",alternate:"varsubsetneqq",scalar:0x2ACB},
 {ordinary:"supsetneqq",alternate:"varsupsetneqq",scalar:0x2ACC},
 {ordinary:"nsubseteq",alternate:"nsubseteqq",scalar:0x2288},
 {ordinary:"nsupseteq",alternate:"nsupseteqq",scalar:0x2289},
];
const expectedIds=families.flatMap(f=>[f.ordinary,f.alternate].flatMap(c=>variants.map(v=>c+"-"+v))).sort();
assert.equal(cases.length,32,"This regression requires exactly 32 input cases");
assert.deepEqual(cases.map(c=>c.id).sort(),expectedIds,"Case IDs must cover precisely four pairs and four styles");
const sourcePaths=[
 "apps/windows/src-windows/VisualTeX.WindowsOffice.VstoShared/MathMlToLatexConverter.cs",
 "apps/windows/src-windows/VisualTeX.WindowsOffice.VstoShared/MathMlTokenStyleContract.cs",
 "apps/windows/src-windows/VisualTeX.WindowsOffice.VstoShared/MathSpaceContract.cs",
];
const probe=fileURLToPath(new URL("spacing-source-probe/bin/Debug/net9.0/SpacingSourceProbe.dll",audit));
const hash=(value:string|Buffer)=>createHash("sha256").update(value).digest("hex");
for(const path of sourcePaths) assert.ok(statSync(probe).mtimeMs>=statSync(new URL(path,repo)).mtimeMs,"Rebuild the isolated harness after changing "+path);
const beforeHashes=Object.fromEntries(sourcePaths.map(p=>[p,hash(readFileSync(new URL(p,repo)))]));
const run=spawnSync("dotnet",[probe],{input:cases.map(c=>JSON.stringify({mathMl:c.decoded})).join("\n")+"\n",encoding:"utf8",maxBuffer:4*1024*1024});
assert.equal(run.status,0,run.stderr);
const recovered=run.stdout.trim().split(/\r?\n/).map(s=>JSON.parse(s).latex as string);
assert.equal(recovered.length,32,"C# recovery must return one line per case");
const parse=(xml:string)=>new DOMParser().parseFromString(xml,"application/xml");
function token(xml:string):Token {
 const all=Array.from(parse(xml).getElementsByTagName("*")).filter(e=>["mi","mo"].includes(e.tagName));
 assert.equal(all.length,1,"Expected one mi/mo token: "+xml);
 const e=all[0],text=e.textContent||"";
 return {role:e.tagName,text,scalars:Array.from(text,c=>c.codePointAt(0)!),alternate:e.getAttribute("data-mjx-alternate")==="1",style:e.getAttribute("mathvariant")||(e.tagName==="mi"?"italic":"normal")};
}
function glyphSignature(svg:string):Glyph[] {
 const doc=parse(svg),paths=new Map(Array.from(doc.getElementsByTagName("path")).map(e=>[e.getAttribute("id"),e.getAttribute("d")]));
 const glyphs=Array.from(doc.getElementsByTagName("use")).map(e=>{
  const path=paths.get((e.getAttribute("xlink:href")||e.getAttribute("href")||"").slice(1));
  assert.ok(path,"Every SVG use must resolve to a real path");
  const transforms:string[]=[];
  for(let n:any=e;n?.nodeType===1;n=n.parentNode) if(n.getAttribute("transform")) transforms.unshift(n.getAttribute("transform"));
  return {code:e.getAttribute("data-c")||"",outline:hash(path!),transforms};
 });
 assert.ok(glyphs.length>0,"Expected resolved SVG glyphs");
 return glyphs;
}
const wrappers:Record<Variant,string>={normal:"",italic:"mathit",bold:"mathbf","bold-italic":"mathbfit"};
const eq=(a:unknown,b:unknown)=>JSON.stringify(a)===JSON.stringify(b);
const rows:Row[]=[];
for(let i=0;i<cases.length;i++){
 const item=cases[i],canonical="\\"+item.command,referenceLatex=wrappers[item.variant]?"\\"+wrappers[item.variant]+"{"+canonical+"}":canonical;
 const family=families.find(f=>[f.ordinary,f.alternate].includes(item.command));
 assert.ok(family,item.id);
 const expected:Token={role:"mo",text:String.fromCodePoint(family.scalar),scalars:[family.scalar],alternate:item.command===family.alternate,style:item.variant};
 const row:Row={id:item.id,command:item.command,variant:item.variant,alternate:item.alternate,latex:recovered[i],referenceLatex,passed:false,checks:{},errors:[],rootHex:item.rootHex};
 try {
  row.source=token(item.source);row.decoded=token(item.decoded);
  row.checks.inputExpectations=item.expected.scalar===family.scalar&&item.alternate===expected.alternate;
  row.checks.codecReportedChecks=Object.values(item.checks).every(value=>value===true);
  for(const [name,value] of Object.entries(expected)){
   row.checks["source."+name]=eq(row.source[name as keyof Token],value);
   row.checks["decoded."+name]=eq(row.decoded[name as keyof Token],value);
  }
  row.mathMl=await latexToMathMl(row.latex,false);
  row.referenceMathMl=await latexToMathMl(referenceLatex,false);
  row.exported=token(row.mathMl);row.reference=token(row.referenceMathMl);
  for(const [name,value] of Object.entries(expected)){
   row.checks["exported."+name]=eq(row.exported[name as keyof Token],value);
   row.checks["canonical."+name]=eq(row.reference[name as keyof Token],value);
  }
  row.checks.canonicalTokenMatch=eq(row.exported,row.reference);
  const svg=await latexToSvg(row.latex),reference=await latexToSvg(referenceLatex);
  row.glyphs=glyphSignature(svg.svg);row.referenceGlyphs=glyphSignature(reference.svg);
  row.geometry=[svg.width,svg.height,svg.baseline];row.referenceGeometry=[reference.width,reference.height,reference.baseline];
  row.checks.canonicalSvgGlyphs=eq(row.glyphs,row.referenceGlyphs);
  row.checks.canonicalSvgGeometry=eq(row.geometry,row.referenceGeometry);
 } catch(error){row.errors.push(error instanceof Error?error.message:String(error));}
 for(const [check,ok] of Object.entries(row.checks)) if(!ok) row.errors.push(check);
 row.passed=row.errors.length===0;rows.push(row);
}
const pairs=families.flatMap(f=>variants.map(variant=>{
 const ordinary=rows.find(r=>r.command===f.ordinary&&r.variant===variant)!;
 const alternate=rows.find(r=>r.command===f.alternate&&r.variant===variant)!;
 const checks={
  styleUnchanged:ordinary.exported?.style===variant&&alternate.exported?.style===variant,
  roleUnchanged:ordinary.exported?.role==="mo"&&alternate.exported?.role==="mo",
  scalarUnchanged:eq(ordinary.exported?.scalars,[f.scalar])&&eq(alternate.exported?.scalars,[f.scalar]),
  alternateDistinguished:ordinary.exported?.alternate===false&&alternate.exported?.alternate===true,
  resolvedGlyphOutlinesDistinct:!!ordinary.glyphs?.length&&!!alternate.glyphs?.length&&!eq(ordinary.glyphs.map(g=>g.outline),alternate.glyphs.map(g=>g.outline)),
  bothCanonicalGlyphControlsMatch:ordinary.checks.canonicalSvgGlyphs===true&&alternate.checks.canonicalSvgGlyphs===true,
 };
 return {ordinary:ordinary.id,alternate:alternate.id,variant,passed:Object.values(checks).every(Boolean),checks};
}));
const afterHashes=Object.fromEntries(sourcePaths.map(p=>[p,hash(readFileSync(new URL(p,repo)))]));
assert.deepEqual(afterHashes,beforeHashes,"Source files changed during the run; rerun against a stable snapshot");
const summary={
 total:32,passed:rows.filter(r=>r.passed).length,failed:rows.filter(r=>!r.passed).length,
 alternateCases:16,ordinaryControls:16,pairs:16,pairsPassed:pairs.filter(p=>p.passed).length,pairsFailed:pairs.filter(p=>!p.passed).length,
 firstErrors:rows.filter(r=>!r.passed).slice(0,4).map(r=>({id:r.id,latex:r.latex,errors:r.errors})),
};
const report={
 generatedAt:new Date().toISOString(),inputPath,inputSha256:hash(inputText),sourceSha256:beforeHashes,
 frontendRuntimeSha256:hash(readFileSync(new URL("apps/windows/src/export/runtime.ts",repo))),
 probeSha256:hash(readFileSync(probe)),summary,
 scope:"Owned 32-case codec corpus through current C# source recovery and production latexToMathMl/latexToSvg. Checks source/readback/export role, Unicode scalar, style, alternate identity, named canonical command reference, resolved SVG path outlines/transforms and geometry. Sixteen ordinary/alternate comparisons keep style fixed. SVG evidence is frontend-only and does not prove MathType/native pixels. No Word, native renderer, screenshot, installed module edit or production edit.",
 rows,pairs
};
const output=process.argv[3]==="--summary"?{...report,rows:undefined,pairs:undefined}:process.argv[3]==="--compact"?{...report,rows:rows.map(({mathMl,referenceMathMl,referenceGlyphs,referenceGeometry,reference,checks,...row})=>({...row,checked:Object.keys(checks),failedChecks:Object.entries(checks).filter(([,ok])=>!ok).map(([key])=>key)}))}:report;
console.log(JSON.stringify(output));
if(summary.failed||summary.pairsFailed) process.exitCode=1;
