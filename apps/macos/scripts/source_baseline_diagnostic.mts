import assert from "node:assert/strict";
import { DOMParser } from "@xmldom/xmldom";
import { formatLatexLines, latexCodeFormats, parseLatexSourceDraft } from "../src/clipboard/LatexCopyService.ts";
import { findLatexFormulaSpans, parseLatexMarkdownDocument } from "../src/office/documentImport/documentImportParser.ts";
import { findWindowsWordLatexRedrawSpans } from "../src/office/redraw/wordLatexRedrawParser.ts";
import { renderOfficeFormulaArtifacts } from "../src/office/shared/formulaRenderArtifacts.ts";
import { normalizeMathModeSource } from "../src/math/mathModeSource.ts";
import { stripVisualTexAlignmentMarkers } from "../src/editor/alignmentMarkers.ts";

globalThis.DOMParser ??= DOMParser;
const doc = new DOMParser().parseFromString("<root/>", "application/xml");
Object.getPrototypeOf(doc).querySelector ??= function(name) { return this.getElementsByTagName(name)?.item(0) ?? null; };
if (!("children" in Object.getPrototypeOf(doc.documentElement))) Object.defineProperty(Object.getPrototypeOf(doc.documentElement), "children", { get() { return Array.from(this.childNodes ?? []).filter(n => n.nodeType === 1); } });


import { readFileSync, writeFileSync } from "node:fs";
const base="test-results/source-baseline-20260912";
const editor=JSON.parse(readFileSync(`${base}/current-editor.json`,"utf8"));
const emitted=formatLatexLines(editor.lines.slice(-1).map(l=>l.latex),editor.latexCodeFormat);
writeFileSync(`${base}/last-line-emitted.txt`,emitted);
for (const [name, source] of [["last-line-emitted",emitted],["import-source",readFileSync(`${base}/import-source.txt`,"utf8")]]) {
 const spans=findLatexFormulaSpans(source);
 console.log(name,spans.length);
 for(const [i,span] of spans.entries()) {
  try { renderOfficeFormulaArtifacts({lines:[{id:"x",latex:span.latex}],codeFormat:"raw",displayMode:span.displayMode,host:"word",formulaLetterFont:editor.formulaLetterFont,formulaChineseFont:editor.formulaChineseFont}); }
  catch(e){console.log(i+1,span.latex,String(e));}
 }
}
for(const [name,formula] of [["N",String.raw`N=150`],["n",String.raw`n_{\max}=12`],["x","x"],["sub","x_i"],["sup","x^2"]]) {
 const a=renderOfficeFormulaArtifacts({lines:[{id:"x",latex:formula}],codeFormat:"raw",displayMode:"inline",host:"word",formulaLetterFont:editor.formulaLetterFont,formulaChineseFont:editor.formulaChineseFont});
 console.log(formula,a.svg.width,a.svg.height,a.svg.baseline,"descent at 11.5:",(a.svg.height-a.svg.baseline)*.75*1.1*11.5/14);
 writeFileSync(`${base}/geometry-${name}.svg`,a.svg.svg);
}
