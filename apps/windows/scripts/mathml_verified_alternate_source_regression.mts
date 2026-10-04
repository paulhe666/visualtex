import assert from "node:assert/strict";
import {readFileSync} from "node:fs";
import {spawnSync} from "node:child_process";
import {fileURLToPath} from "node:url";
import {createHash} from "node:crypto";
import {DOMParser} from "@xmldom/xmldom";
import {latexToMathMl,latexToSvg} from "../src/export/runtime.ts";

const root = new URL("../../../artifacts/conversion-symbol-audit-20261002/",import.meta.url);
const cases = JSON.parse(readFileSync(new URL("alternate-five-contract-20261002T142900Z/cases.json",root),"utf8")) as Array<{id:string;command:string;alternate:boolean;variant:string;source:string;decoded:string;passed:boolean}>;
assert.equal(cases.length,40); assert.ok(cases.every(x=>x.passed));
const probe = fileURLToPath(new URL("spacing-source-probe/bin/Debug/net9.0/SpacingSourceProbe.dll",root));
const run = spawnSync("dotnet",[probe],{input:cases.map(x=>JSON.stringify({mathMl:x.decoded})).join("\n")+"\n",encoding:"utf8"});
assert.equal(run.status,0,run.stderr);
const recovered = run.stdout.trim().split(/\r?\n/).map(x=>JSON.parse(x).latex as string);
const parse=(xml:string)=>new DOMParser().parseFromString(xml,"application/xml");
function token(xml:string) {
 const all=Array.from(parse(xml).getElementsByTagName("*")).filter(e=>["mi","mo"].includes(e.tagName));
 assert.equal(all.length,1,xml); const e=all[0];
 return {role:e.tagName,text:e.textContent,alternate:e.getAttribute("data-mjx-alternate")==="1",style:e.getAttribute("mathvariant") || (e.tagName==="mi"?"italic":"normal")};
}
function glyphSignature(svg:string) {
 const doc=parse(svg), paths=new Map(Array.from(doc.getElementsByTagName("path")).map(e=>[e.getAttribute("id"),e.getAttribute("d")]));
 return Array.from(doc.getElementsByTagName("use")).map(e=>{
  const path=paths.get((e.getAttribute("xlink:href")||e.getAttribute("href")||"").slice(1));
  assert.ok(path,svg); const transforms=[];
  for(let n:any=e;n?.nodeType===1;n=n.parentNode) if(n.getAttribute("transform")) transforms.unshift(n.getAttribute("transform"));
  return {code:e.getAttribute("data-c"),outline:createHash("sha256").update(path!).digest("hex"),transforms};
 });
}
const wrappers:Record<string,string>={normal:"",italic:"mathit",bold:"mathbf","bold-italic":"mathbfit"};
const rows=[];
for(let i=0;i<cases.length;i++){
 const item=cases[i], latex=recovered[i], from=token(item.decoded);
 const canonical="\\"+item.command;
 const expected=wrappers[item.variant] ? "\\"+wrappers[item.variant]+"{"+canonical+"}" : canonical;
 const mathMl=await latexToMathMl(latex,false), actual=token(mathMl), intended=token(await latexToMathMl(expected,false));
 assert.equal(actual.text,from.text,item.id);
 assert.equal(actual.alternate,item.alternate,item.id);
 assert.equal(actual.style,item.variant,item.id);
 assert.deepEqual(actual,intended,item.id);
 assert.equal(actual.role,"mo",item.id);
 const svg=await latexToSvg(latex), reference=await latexToSvg(expected);
 assert.deepEqual(glyphSignature(svg.svg),glyphSignature(reference.svg),item.id);
 assert.deepEqual([svg.width,svg.height,svg.baseline],[reference.width,reference.height,reference.baseline],item.id);
 rows.push({id:item.id,latex,decoded:from,exported:actual,glyphs:glyphSignature(svg.svg),width:svg.width,height:svg.height});
}
const ordinaryRoleLimits=rows.filter(x=>x.decoded.role!==x.exported.role).map(x=>x.id);
assert.deepEqual(ordinaryRoleLimits.sort(),["nmid-italic","nmid-bold","nmid-bold-italic","subsetneq-italic","subsetneq-bold","subsetneq-bold-italic","supsetneq-italic","supsetneq-bold","supsetneq-bold-italic"].sort());
for(const alt of rows.filter(x=>x.exported.alternate)){
 const family=alt.id.replace(/-(normal|italic|bold|bold-italic)$/,"");
 const ordinary={shortmid:"mid",nshortmid:"nmid",smallsetminus:"setminus",varsubsetneq:"subsetneq",varsupsetneq:"supsetneq"}[family]!;
 const other=rows.find(x=>x.id===ordinary+"-"+alt.exported.style)!;
 assert.notDeepEqual(alt.glyphs.map(x=>x.outline),other.glyphs.map(x=>x.outline),alt.id+" must retain a real alternate outline");
}
console.log(JSON.stringify({passed:rows.length,alternates:20,ordinaryControls:20,ordinaryRoleLimits,scope:"Current source recovery from verified managed codec readback into production MathJax; all40 roles/styles/alternate flags and resolved SVG outlines/geometry checked. Native runtime not invoked.",rows}));
