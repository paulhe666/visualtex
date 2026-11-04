// The document as Word (.docx) and PDF (DESIGN.md §12): the shared writers (docx.ts, printCopy.ts) with the app's
// formula renderings — OMML through the shared Word writer (Rust omml_write), VisualTeX formula pictures from the
// editor's engine (the same pictures the Office add-ins insert, their metadata in the alt text).
import { invoke, isTauri } from "@tauri-apps/api/core";
import { zipSync } from "fflate";
import type { Node } from "prosemirror-model";
import {
  docxFormulaRequests,
  documentSettings,
  documentToDocx,
  type DocxFormula,
  type DocxFormulaRequest,
} from "../../../shared/visualtex-document/src/index.ts";
import { latexToMathMl, latexToSvg } from "../export/runtime";
import { svgToPng } from "../export/svgToPng";
import { createFormulaMetadata, encodeFormulaMetadata } from "../office/shared/formulaMetadata";
import { useEditorStore } from "../stores/editorStore";
import { portableFormulaLatex } from "./DocumentEditor";

import type { WordFormulaFormat } from "../stores/documentFormatStore";
export type { WordFormulaFormat };

const PX_TO_PT = 0.75;

/** VisualTeX formulas: the editor engine's SVG (and a 3× PNG for programs without SVG), sized at the text's size. */
async function pictureFormula(request: DocxFormulaRequest, sizePt: number): Promise<DocxFormula> {
  const editor = useEditorStore.getState();
  const latex = portableFormulaLatex(request.latex);
  const svg = await latexToSvg(latex, {
    displayMode: request.display,
    fontSizePt: sizePt,
    paddingPx: request.display ? 2 : 1,
    background: "transparent",
    formulaLetterFont: editor.formulaLetterFont,
    formulaChineseFont: editor.formulaChineseFont,
    renderer: "engine",
    color: "#000000",
  });
  const png = await svgToPng(svg, { scale: 3, background: "transparent" });
  const formulaId = crypto.randomUUID();
  const metadata = createFormulaMetadata({
    formulaId,
    title: "",
    lines: latex.split("\n").filter((line) => line.trim()).map((line) => ({ id: crypto.randomUUID(), latex: line })),
    codeFormat: "latex",
    displayMode: request.display ? "block" : "inline",
    numbered: !!request.number,
    renderWidthPx: svg.width,
    renderHeightPx: svg.height,
    baseline: svg.baseline,
    fontSizePt: sizePt,
    renderFontSizePt: sizePt,
    formulaLetterFont: editor.formulaLetterFont,
    formulaChineseFont: editor.formulaChineseFont,
  });
  return {
    kind: "picture",
    png: base64Bytes(png.base64),
    svg: new TextEncoder().encode(svg.svg),
    widthPt: svg.width * PX_TO_PT,
    heightPt: svg.height * PX_TO_PT,
    depthPt: Math.max(0, svg.height - svg.baseline) * PX_TO_PT,
    descr: encodeFormulaMetadata(metadata),
    title: `visualtex:formula-ref:v1:${formulaId}:${request.display ? "block" : "inline"}:${request.number ? 1 : 0}`,
  };
}

function base64Bytes(base64: string): Uint8Array {
  const binary = atob(base64.replace(/^data:[^,]*,/, ""));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

interface OmmlAnswer {
  omml?: string;
  error?: string;
}

/** Office equations: MathML from the formula, OMML (numbered ones with Word's own "#" number) from the app. */
async function ommlFormulas(requests: DocxFormulaRequest[], mathFont: string): Promise<Map<string, DocxFormula>> {
  const items = await Promise.all(
    requests.map(async (request) => ({
      mathml: await latexToMathMl(portableFormulaLatex(request.latex), request.display),
      number: request.number ?? null,
    })),
  );
  const answers = await invoke<OmmlAnswer[]>("document_export_omml", { items, mathFont });
  const out = new Map<string, DocxFormula>();
  requests.forEach((request, index) => {
    const omml = answers[index]?.omml;
    if (omml) out.set(request.key, { kind: "omml", xml: omml });
  });
  return out;
}

export interface WordExportResult {
  bytes: Uint8Array;
  /** formulas that could not be written in the chosen form (written as pictures, or as their LaTeX) */
  fallbacks: number;
}

/** The document as a .docx file, its formulas in the chosen form. */
export async function documentToWord(doc: Node, format: WordFormulaFormat, labels: { bibliographyTitle: string; title?: string }, onProgress?: (done: number, total: number) => void): Promise<WordExportResult> {
  const requests = docxFormulaRequests(doc);
  const settings = documentSettings(doc);
  const bodySize = settings.fontSize;
  const formulas = new Map<string, DocxFormula>();
  let fallbacks = 0;
  if (format === "omml" && isTauri()) {
    const omml = await ommlFormulas(requests, "Cambria Math");
    for (const [key, value] of omml) formulas.set(key, value);
  } else if (format === "mathtype" && isTauri()) {
    const objects = await invoke<Array<{ key: string; runXml: string; parts: Array<{ path: string; base64: string; contentType: string; extension: string }> } | null>>("document_export_mathtype", {
      items: await Promise.all(requests.map(async (request) => ({ key: request.key, mathml: await latexToMathMl(portableFormulaLatex(request.latex), request.display), display: request.display, number: request.number ?? null, fontSizePt: bodySize }))),
    });
    for (const object of objects) {
      if (!object) continue;
      formulas.set(object.key, {
        kind: "object",
        runXml: (rel) => object.runXml.replace(/\{rel:([^|}]+)\|([^}]+)\}/g, (_m, target: string, type: string) => rel(target, type)),
        parts: object.parts.map((part) => ({ path: part.path, bytes: base64Bytes(part.base64), contentType: part.contentType, extension: part.extension })),
      });
    }
  }
  // the rest (and every formula as a VisualTeX picture): drawn by the engine
  let done = 0;
  for (const request of requests) {
    if (!formulas.has(request.key)) {
      if (format !== "visualtex") fallbacks++;
      try {
        formulas.set(request.key, await pictureFormula(request, request.key.startsWith("fn:") ? bodySize * 0.8 : bodySize));
      } catch {
        // left as its LaTeX
      }
    }
    onProgress?.(++done, requests.length);
  }
  const files = documentToDocx(doc, formulas, { bibliographyTitle: labels.bibliographyTitle, title: labels.title, mathFont: "Cambria Math" });
  const ordered: Record<string, Uint8Array> = { "[Content_Types].xml": files["[Content_Types].xml"] };
  for (const [path, bytes] of Object.entries(files)) if (path !== "[Content_Types].xml") ordered[path] = bytes;
  return { bytes: zipSync(ordered, { level: 6 }), fallbacks };
}
