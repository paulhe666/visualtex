import assert from "node:assert/strict";
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { spawnSync } from "node:child_process";
import { DOMParser } from "@xmldom/xmldom";
import { zipSync, unzipSync, strToU8 } from "fflate";
import { renderOfficeFormulaArtifacts } from "../src/office/shared/formulaRenderArtifacts";
import { createFormulaMetadata, encodeFormulaMetadata } from "../src/office/shared/formulaMetadata";
import { wordImageReferenceGeometry } from "../src/office/shared/wordImageGeometry";
globalThis.DOMParser ??= DOMParser as any;
const domProbe = new DOMParser().parseFromString("<root/>", "application/xml");
Object.getPrototypeOf(domProbe).querySelector ??= function(name: string) { return this.getElementsByTagName(name).item(0); };
if (!("children" in domProbe.documentElement)) Object.defineProperty(Object.getPrototypeOf(domProbe.documentElement), "children", { get() { return Array.from(this.childNodes).filter((node: any) => node.nodeType === 1); } });
const root = join(homedir(), "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime/Tests/numbering-edit");
mkdirSync(root, { recursive: true });
function run(args: string[]) {
  const result = spawnSync("/usr/bin/osascript", args.flatMap(line => ["-e", line]), { encoding: "utf8", timeout: 1200000 });
  assert.equal(result.status, 0, result.stderr); return result.stdout;
}
if (process.argv.includes("--run") || process.argv.includes("--verify")) {
  if (process.argv.includes("--run")) run(['tell application "Microsoft Word"', 'with timeout of 1200 seconds', 'activate', 'run VB macro macro name "VisualTeX_RunWordNumberingEditRegression"', 'end timeout', 'end tell']);
  const result = readFileSync(join(root, "result.txt"), "utf8");
  console.log(result); assert.ok(result.startsWith("PASS"), result);
  const samples = result.split("\n").filter(line => line.startsWith("sample|")).map(line => line.split("|"));
  assert.equal(samples.length, 120);
  for (const native of ["False", "True"]) for (const chapter of ["False", "True"]) {
    const durations = samples.filter(row => row[1] === native && row[2] === chapter).map(row => Number(row[5])).sort((a, b) => a - b);
    assert.equal(durations.length, 30);
    console.log(JSON.stringify({ host: native === "True" ? "OMML" : "image", numbering: chapter === "True" ? "chapter" : "sequence", samples: 30, medianMs: durations[15], p95Ms: durations[28], maxMs: durations[29] }));
  }
} else {
  const lines = [{ id: "81818181-8181-4181-8181-818181818181", latex: String.raw`x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}` }];
  const on = renderOfficeFormulaArtifacts({ lines, codeFormat: "raw", displayMode: "block", host: "word", numbered: true, formulaLetterFont: "times", fontSizePt: 14 });
  const off = renderOfficeFormulaArtifacts({ lines, codeFormat: "raw", displayMode: "block", host: "word", numbered: false, formulaLetterFont: "times", fontSizePt: 14 });
  assert.ok(on.omml && off.omml);
  const geometry = wordImageReferenceGeometry(on.svg.width, on.svg.height, on.svg.baseline, "times");
  writeFileSync(join(root, "on.docx"), Buffer.from(on.omml.ommlDocxBase64, "base64url"));
  writeFileSync(join(root, "off.docx"), Buffer.from(off.omml.ommlDocxBase64, "base64url"));
  writeFileSync(join(root, "image.svg"), on.svg.svg);
  const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAEAQH/8l0Z8QAAAABJRU5ErkJggg==", "base64");
  writeFileSync(join(root, "image.png"), png);
  // Use the production OMML package as the Word document shell and replace its
  // body with a DrawingML SVG/PNG image. Word performs the actual import.
  const files = unzipSync(Buffer.from(off.omml.ommlDocxBase64, "base64url"));
  const w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const a = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
  const wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const cx = Math.round(geometry.referenceWidthPt * 12700), cy = Math.round(geometry.referenceHeightPt * 12700);
  files["word/document.xml"] = strToU8(`<w:document xmlns:w="${w}" xmlns:r="${r}" xmlns:a="${a}" xmlns:wp="${wp}" xmlns:pic="${pic}" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="Formula"/><a:graphic><a:graphicData uri="${pic}"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="formula.svg"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="png"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="svg"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`);
  const packageNs = "http://schemas.openxmlformats.org/package/2006/relationships";
  files["word/_rels/document.xml.rels"] = strToU8(`<Relationships xmlns="${packageNs}"><Relationship Id="svg" Type="${r}/image" Target="media/formula.svg"/><Relationship Id="png" Type="${r}/image" Target="media/formula.png"/></Relationships>`);
  files["word/media/formula.svg"] = strToU8(on.svg.svg); files["word/media/formula.png"] = png;
  files["[Content_Types].xml"] = strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="svg" ContentType="image/svg+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  writeFileSync(join(root, "image.docx"), zipSync(files));
  for (let index = 1; index <= 50; index++) for (const numbered of [false, true]) {
    const formulaId = `91919191-9191-4191-8191-${String(index).padStart(12, "0")}`;
    const artifacts = numbered ? on : off;
    const metadata = encodeFormulaMetadata(createFormulaMetadata({ formulaId, title: "Numbering regression", lines, codeFormat: "raw", displayMode: "block", numbered, fontSizePt: 14, formulaLetterFont: "times", renderWidthPx: on.svg.width, renderHeightPx: on.svg.height, imageInkCenterYRatio: 0.5, ...geometry }));
    const dispatch = { formulaId, displayMode: "block", numbered: numbered ? "1" : "0", metadata, latexBase64: Buffer.from(artifacts.canonicalLatex).toString("base64url"), ommlBase64: artifacts.omml!.ommlBase64, nativeDocumentPath: join(root, numbered ? "on.docx" : "off.docx"), imagePath: join(root, "image.svg"), vectorDocumentPath: join(root, "image.docx"), fallbackImagePath: join(root, "image.png"), widthPoints: geometry.referenceWidthPt, heightPoints: geometry.referenceHeightPt, fontSizePt: 14, ...geometry, wordMathFontName: "Times New Roman", inkCenterYRatio: 0.5 };
    writeFileSync(join(root, `${formulaId}-${numbered ? "on" : "off"}.txt`), Object.entries(dispatch).map(([key, value]) => `${key}=${value}`).join("\n"));
  }
  const scratch = join(homedir(), "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch");
  const blank = { ...files, "word/document.xml": strToU8(`<w:document xmlns:w="${w}"><w:body><w:p/></w:body></w:document>`) };
  for (let representation = 0; representation < 2; representation++) for (let format = 0; format < 2; format++) writeFileSync(join(scratch, `numbering-acceptance-${representation}-${format}.docx`), zipSync(blank));
  console.log(`Prepared 50 formula pairs in ${root}`);
}
