import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { DOMParser } from "@xmldom/xmldom";
import { latexToMathMl, latexToSvg } from "../src/export/runtime.ts";
const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml");
const hash = (value: string) => createHash("sha256").update(value).digest("hex");
const probe = fileURLToPath(new URL("../../../artifacts/conversion-symbol-audit-20261002/spacing-source-probe/bin/Debug/net9.0/SpacingSourceProbe.dll", import.meta.url));
function recover(mathMl: string) {
  const run = spawnSync("dotnet", [probe], {input:JSON.stringify({mathMl}) + "\n",encoding:"utf8"});
  assert.equal(run.status, 0, run.stderr);
  return JSON.parse(run.stdout.trim()).latex as string;
}
function tokens(xml: string) {
  return Array.from(parse(xml).getElementsByTagName("*"))
    .filter(n => ["mi", "mn", "mo", "mtext"].includes(n.tagName))
    .map(n => {
      let variant = "";
      for (let cursor:any = n; cursor?.nodeType === 1; cursor = cursor.parentNode) {
        variant = cursor.getAttribute("mathvariant");
        if (variant) break;
      }
      return {kind:n.tagName,text:n.textContent,variant:variant || (n.tagName === "mi" && Array.from(n.textContent ?? "").length === 1 ? "italic" : "normal")};
    });
}
function visual(rendered: Awaited<ReturnType<typeof latexToSvg>>) {
  const doc = parse(rendered.svg);
  const paths = new Map(Array.from(doc.getElementsByTagName("path")).map(n => [n.getAttribute("id"), n.getAttribute("d")]));
  const marks:any[] = [];
  function visit(node:any, transforms:string[]) {
    if (node.nodeType !== 1 || node.tagName === "defs") return;
    const next = node.getAttribute("transform") ? [...transforms,node.getAttribute("transform")] : transforms;
    if (node.tagName === "use") {
      const outline = paths.get((node.getAttribute("xlink:href") || node.getAttribute("href")).slice(1));
      assert.ok(outline, "Every used glyph must resolve to its real SVG outline");
      marks.push({kind:"glyph",code:node.getAttribute("data-c"),outline:hash(outline!),x:node.getAttribute("x"),y:node.getAttribute("y"),transforms:next});
    } else if (node.tagName === "text") {
      marks.push({kind:"text",text:node.textContent,style:node.getAttribute("style"),x:node.getAttribute("x"),y:node.getAttribute("y"),transforms:next});
    } else if (node.tagName === "rect") {
      marks.push({kind:"rect",x:node.getAttribute("x"),y:node.getAttribute("y"),width:node.getAttribute("width"),height:node.getAttribute("height"),transforms:next});
    }
    for (let child=node.firstChild;child;child=child.nextSibling) visit(child,next);
  }
  visit(doc.documentElement,[]);
  return {width:rendered.width,height:rendered.height,baseline:rendered.baseline,viewBox:doc.documentElement.getAttribute("viewBox"),marks};
}
const priorPi = "100\\text{\\%}\\mathrm{e}^{\\mathrm{i}\\pi}+1=0";
const explicitPi = "100\\text{\\%}\\mathrm{e}^{\\mathrm{i}\\mathit{π}}+1=0";
let passed = 0;
let geometryEvidence: unknown;
for (const displayMode of [false,true]) for (const fontSizePt of [10.5,12]) {
  const options = {displayMode,fontSizePt,paddingPx:8};
  const normalA = await latexToSvg("A",options);
  const mathitA = await latexToSvg("\\mathit{A}",options);
  assert.deepEqual(tokens(await latexToMathMl("A",displayMode)),tokens(await latexToMathMl("\\mathit{A}",displayMode)));
  assert.notDeepEqual(visual(normalA),visual(mathitA),"Effective italic alone cannot prove glyph-family equivalence");
  const first = await latexToSvg(priorPi,options);
  const second = await latexToSvg(explicitPi,options);
  assert.deepEqual(tokens(await latexToMathMl(priorPi,displayMode)),tokens(await latexToMathMl(explicitPi,displayMode)));
  assert.deepEqual(visual(first),visual(second),"Compare resolved glyph outlines and every accumulated placement transform, not just widths or cache ids");
  assert.notDeepEqual(visual(normalA),visual(await latexToSvg("\\mathrm{A}",options)));
  if (displayMode && fontSizePt === 12) geometryEvidence = {
    A:{mathItalic:visual(normalA),texMathit:visual(mathitA)},
    pi:visual(first),
  };
  passed += 3;
}
const recoveryCases = [
  {mathMl:"<math><mi mathvariant='italic' data-mtef-character-style='2'>A</mi></math>",source:"A"},
  {mathMl:"<math><mi mathvariant='italic'>π</mi></math>",source:"\\pi"},
  {mathMl:await latexToMathMl("\\mathit{A}",false),source:"\\mathit{A}"},
  {mathMl:await latexToMathMl("\\mathit{π}",false),source:"\\mathit{π}"},
  {mathMl:"<math><mi mathvariant='normal'>A</mi></math>",source:"\\mathrm{A}"},
];
for(const item of recoveryCases) {
  const source = recover(item.mathMl);
  assert.equal(source,item.source);
  assert.deepEqual(visual(await latexToSvg(source)),visual(await latexToSvg(item.source)));
  passed++;
}
console.log(JSON.stringify({
  passed,
  result:"Mathematical italic and TeX mathit stay distinct; equivalent pi spelling verified by resolved SVG path hashes and placement transforms.",
  geometryEvidence,
  scope:"Production source export/SVG geometry evidence; no screenshot or native/Word runtime acceptance."
}));
