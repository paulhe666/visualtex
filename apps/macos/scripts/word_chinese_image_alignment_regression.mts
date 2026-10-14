import assert from "node:assert/strict";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { homedir } from "node:os";
import { join } from "node:path";
import { DOMParser } from "@xmldom/xmldom";
import { strToU8, strFromU8, unzipSync, zipSync } from "fflate";
import { renderOfficeFormulaArtifacts } from "../src/office/shared/formulaRenderArtifacts";
import { createFormulaMetadata, encodeFormulaMetadata } from "../src/office/shared/formulaMetadata";
import { wordImageReferenceGeometry } from "../src/office/shared/wordImageGeometry";

globalThis.DOMParser ??= DOMParser as any;
const root = join(homedir(), "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch");
const fixturePath = join(root, "word-chinese-image-alignment-fixture.docx");
const casesPath = join(root, "word-chinese-image-alignment-cases.json");
const vectors = join(process.cwd(), ".tmp-chinese-alignment-vectors");
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const xml = (value: unknown) => String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll('"', "&quot;");
const variable = (prefix: string, id: string) => `${prefix}${id.replaceAll("-", "_")}`;
const size = 18;
const text = (value: string, font = "Arial") => `<w:r><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="Heiti SC"/><w:sz w:val="36"/></w:rPr><w:t xml:space="preserve">${xml(value)}</w:t></w:r>`;
const modes = ["baseline", "top", "center", "bottom", "auto"];
function run(command: string, args: string[], env = process.env) {
  const result = spawnSync(command, args, { encoding: "utf8", timeout: 660000, env, maxBuffer: 16 * 1024 * 1024 });
  assert.equal(result.status, 0, result.stderr || result.stdout || result.error?.message);
  return result.stdout.trim();
}
async function waitForWordExit() {
  const deadline = Date.now() + 20000;
  while (run("/usr/bin/osascript", ["-e", 'application "Microsoft Word" is running']) === "true") {
    assert.ok(Date.now() < deadline, "Word did not finish closing");
    await new Promise(resolve => setTimeout(resolve, 500));
  }
}
if (process.argv.includes("--run")) {
  const cases = JSON.parse(readFileSync(casesPath, "utf8"));
  const metrics = run("/Applications/VisualTeX.app/Contents/MacOS/visualtex", ["--word-image-alignment-metrics", JSON.stringify({ bodyFont: "Heiti SC", bodySample: "中文", metadata: cases[0].metadata })]);
  assert.equal(metrics.split("|").length, 6, "Install the matching r108 client before running Word pixel acceptance.");
  run("/usr/bin/osascript", ["-e", 'if application "Microsoft Word" is running then',
    "-e", 'tell application "Microsoft Word" to quit saving no', "-e", "end if"]);
  await waitForWordExit();
  run("/usr/bin/osascript", ["-e", 'tell application "Microsoft Word"',
    "-e", "with timeout of 600 seconds",
    "-e", "activate", "-e", `open file name ${JSON.stringify(fixturePath)} read only false`,
    "-e", 'run VB macro macro name "VisualTeX_RunWordChineseImageAlignmentRegression"',
    "-e", "end timeout", "-e", "end tell"]);
  const status = readFileSync(join(homedir(), "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime/Tests/word-chinese-image-alignment-result.txt"), "utf8");
  console.log(status);
  assert.ok(status.startsWith("PASS"), status);
} else if (process.argv.includes("--live")) {
  // Exercise automatic startup and the actual idle callback. This path never
  // invokes the initializer, positioning macro, or watcher directly.
  const cases = JSON.parse(readFileSync(casesPath, "utf8"));
  const files = unzipSync(readFileSync(fixturePath));
  const doc = new DOMParser().parseFromString(strFromU8(files["word/document.xml"]), "application/xml");
  const positions = new Map<string, number>();
  for (const drawing of Array.from(doc.getElementsByTagNameNS(W, "drawing"))) {
    const id = drawing.getElementsByTagNameNS(WP, "docPr")[0].getAttribute("title")?.split(":")[3];
    const item = cases.find((candidate: any) => candidate.formulaId === id);
    let run: any = drawing.parentNode;
    while (run.localName !== "r") run = run.parentNode;
    positions.set(`${item.alignment}/${item.kind}`, Number(run.getElementsByTagNameNS(W, "position")[0]?.getAttributeNS(W, "val") ?? 0) / 2);
  }
  run("/usr/bin/osascript", ["-e", 'if application "Microsoft Word" is running then',
    "-e", 'tell application "Microsoft Word" to quit saving no', "-e", "end if"]);
  await waitForWordExit();
  run("/usr/bin/osascript", ["-e", 'tell application "Microsoft Word"',
    "-e", "activate", "-e", `open file name ${JSON.stringify(fixturePath)} read only false`, "-e", "end tell"]);
  const openDeadline = Date.now() + 20000;
  while (run("/usr/bin/osascript", ["-e", 'tell application "Microsoft Word"',
    "-e", "if count of documents = 0 then return 0", "-e", "return count of inline shapes of active document", "-e", "end tell"]) !== "45") {
    assert.ok(Date.now() < openDeadline, "Word did not open the fixture after startup");
    await new Promise(resolve => setTimeout(resolve, 500));
  }
  const enums = ["baseline", "top", "center", "east asian50", "auto"];
  try {
    for (const [index, mode] of modes.entries()) {
      run("/usr/bin/osascript", ["-e", 'tell application "Microsoft Word"',
        "-e", `set base line alignment of paragraph format of text object of active document to baseline align ${enums[index]}`,
        "-e", "end tell"]);
      const deadline = Date.now() + 20000;
      let actual: number[] = [];
      do {
        await new Promise(resolve => setTimeout(resolve, 1000));
        const value = run("/usr/bin/osascript", ["-e", 'tell application "Microsoft Word"',
          "-e", 'set resultText to ""', "-e", "repeat with shapeIndex from 1 to (count of inline shapes of active document)",
          "-e", "set formulaShape to inline shape shapeIndex of active document",
          "-e", 'set resultText to resultText & (font position of font object of text object of formulaShape) & linefeed',
          "-e", "end repeat", "-e", "return resultText", "-e", "end tell"]);
        actual = value.split(/\s+/).map(Number);
        if (cases.every((item: any, i: number) => actual[i] === positions.get(`${mode}/${item.kind}`))) break;
      } while (Date.now() < deadline);
      assert.equal(actual.length, 45);
      for (const [i, item] of cases.entries()) assert.equal(actual[i], positions.get(`${mode}/${item.kind}`), `${mode}/image ${i + 1}: automatic idle repair did not run`);
      console.log(`PASS actual idle observer: ${mode}, 45 images, caret outside formula paragraphs`);
    }
  } finally {
    try {
      run("/usr/bin/osascript", ["-e", 'tell application "Microsoft Word" to quit saving no']);
      await waitForWordExit();
      run("/usr/bin/osascript", ["-e", 'tell application "Microsoft Word"',
        "-e", `open file name ${JSON.stringify(fixturePath)} read only false`, "-e", "end tell"]);
    } catch (error) {
      console.error("Word fixture cleanup:", String(error));
    }
  }
} else if (process.argv.includes("--verify")) {
  const expected = JSON.parse(readFileSync(casesPath, "utf8"));
  const path = process.argv[process.argv.indexOf("--verify") + 1] ?? fixturePath;
  const files = unzipSync(readFileSync(path));
  const doc = new DOMParser().parseFromString(strFromU8(files["word/document.xml"]), "application/xml");
  const settings = new DOMParser().parseFromString(strFromU8(files["word/settings.xml"]), "application/xml");
  const variables = new Map(Array.from(settings.getElementsByTagNameNS(W, "docVar")).map(v => [v.getAttributeNS(W, "name"), v.getAttributeNS(W, "val")]));
  const rows = Array.from(doc.getElementsByTagNameNS(W, "drawing")).map(drawing => {
    const props = drawing.getElementsByTagNameNS(WP, "docPr")[0];
    const id = props.getAttribute("title")?.split(":")[3];
    const item = expected.find((candidate: any) => candidate.formulaId === id);
    assert.ok(item, `Unexpected formula ${id}`);
    let run: any = drawing.parentNode;
    while (run.localName !== "r") run = run.parentNode;
    let paragraph = run.parentNode;
    while (paragraph.localName !== "p") paragraph = paragraph.parentNode;
    const alignment = paragraph.getElementsByTagNameNS(W, "textAlignment")[0]?.getAttributeNS(W, "val") ?? "auto";
    const position = Number(run.getElementsByTagNameNS(W, "position")[0]?.getAttributeNS(W, "val") ?? 0) / 2;
    const extent = drawing.getElementsByTagNameNS(WP, "extent")[0];
    const height = Number(extent.getAttribute("cy")) / 12700;
    const width = Number(extent.getAttribute("cx")) / 12700;
    assert.equal(alignment, item.alignment, `User typography changed: ${id}`);
    const canvas = variables.get(variable("VT_ImageCanvas_", id))?.split("|").map(Number);
    const glyphHeight = canvas ? canvas[1] * height / canvas[2] : height;
    assert.ok(Math.abs(glyphHeight - item.heightPt) < 0.1 && Math.abs(width - item.widthPt) < 0.1, `Image resize changed painted geometry: ${id}`);
    if (canvas) {
      assert.equal(alignment, "top", "Transparent canvas belongs only to Top typography");
      assert.ok(height > glyphHeight && position > 0, "Top canvas must use a positive rendered position");
    }
    assert.ok(Number.isInteger(position) && Math.abs(position) <= 256);
    return { index: item.index, kind: item.kind, alignment, position, height, glyphHeight, canvas: Boolean(canvas) };
  });
  assert.equal(rows.length, 45);
  for (const kind of ["quadratic-display", "integral-display", "quadratic-text", "integral-text", "sum-text", "subscript", "superscript", "block", "numbered-block"]) {
    const subset = rows.filter(row => row.kind === kind);
    assert.equal(subset.length, 5);
    // Distinct Word anchors may converge on the same integral Position after
    // optical/raster rounding, especially for short letters. Require an actual
    // mode-dependent result, not three artificially different offsets.
    assert.ok(new Set(subset.map(row => row.position)).size >= 2, `${kind}: host alignment has no dedicated compensation`);
  }
  console.log(JSON.stringify({ result: "PASS", count: rows.length, rows }, null, 2));
} else {
  mkdirSync(root, { recursive: true });
  mkdirSync(vectors, { recursive: true });
  const files: Record<string, Uint8Array> = {};
  const cases: any[] = [];
  let relationships = "", variables = "", paragraphs = "", index = 0;
  const quadratic = String.raw`x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}`;
  const integral = String.raw`\int_{-\infty}^{\infty}e^{-x^2}\,\mathrm{d}x=\sqrt{\pi}`;
  function picture(latex: string, style: "text" | "display", alignment: string, kind: string, block = false, numbered = false) {
    index++;
    const formulaId = `90909090-9090-4090-8090-${String(index).padStart(12, "0")}`;
    const displayMode = block ? "block" : "inline";
    const lines = [{ id: "81818181-8181-4181-8181-818181818181", latex }];
    const svg = renderOfficeFormulaArtifacts({ lines, host: "word", codeFormat: "raw", displayMode, inlineImageMathStyle: style, numbered, fontSizePt: size, formulaLetterFont: "times", includeWordOmml: false }).svg;
    const geometry = wordImageReferenceGeometry(svg.width, svg.height, svg.baseline, "times");
    const widthPt = geometry.referenceWidthPt * size / 14, heightPt = geometry.referenceHeightPt * size / 14;
    const metadata = encodeFormulaMetadata(createFormulaMetadata({ formulaId, title: kind, lines, codeFormat: "raw", displayMode, inlineImageMathStyle: style, numbered, fontSizePt: size, formulaLetterFont: "times", renderWidthPx: svg.width, renderHeightPx: svg.height, imageInkCenterYRatio: 0.5, ...geometry }));
    const title = `visualtex:formula-ref:v1:${formulaId}:${displayMode}:${numbered ? 1 : 0}`;
    // Deliberately start with the old baseline-only policy. The production
    // Word event/timer path must correct every image without changing its mode.
    const position = -Math.floor(-geometry.referenceBaselinePt * size / 14 + 0.51);
    const name = `f${index}.svg`;
    writeFileSync(join(vectors, name), svg.svg);
    files[`word/media/${name}`] = strToU8(svg.svg);
    files[`word/media/f${index}.png`] = new Uint8Array(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAEAQH/8l0Z8QAAAABJRU5ErkJggg==", "base64"));
    relationships += `<Relationship Id="s${index}" Type="${R}/image" Target="media/${name}"/><Relationship Id="p${index}" Type="${R}/image" Target="media/f${index}.png"/>`;
    for (const [prefix, value] of [["VT_ImageScale_", `${size}|${geometry.referenceWidthPt}|${geometry.referenceHeightPt}|${geometry.referenceBaselinePt}|${size}`], ["VT_ImageInkCenter_", "0.5"], ["VT_Format_", `${displayMode}|${numbered ? 1 : 0}`], ["VT_Metadata_", metadata], ["VT_Latex_", Buffer.from(latex).toString("base64url")]]) {
      variables += `<w:docVar w:name="${variable(prefix, formulaId)}" w:val="${xml(value)}"/>`;
    }
    cases.push({ index, formulaId, kind, displayMode, numbered, style, alignment, widthPt, heightPt, metadata, ...geometry });
    const cx = Math.round(widthPt * 12700), cy = Math.round(heightPt * 12700);
    return `<w:r><w:rPr><w:sz w:val="36"/><w:position w:val="${position * 2}"/></w:rPr><w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${index}" name="${kind}" title="${title}" descr="${xml(metadata)}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="${index}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="p${index}"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="s${index}"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  }
  function page(alignment: string, block: boolean) {
    paragraphs += `<w:p><w:pPr>${paragraphs ? "<w:pageBreakBefore/>" : ""}</w:pPr>${text(`${alignment} / 18 pt / ${block ? "block" : "reference inline"}`)}</w:p>`;
    const rows = block ? [picture(quadratic, "display", alignment, "block", true), picture(integral, "display", alignment, "numbered-block", true, true) + text("\t(1)", "Cambria Math")] : [
      text("Displaystyle 样式 ") + picture(quadratic, "display", alignment, "quadratic-display") + text(" 也是这样的 ") + picture(integral, "display", alignment, "integral-display"),
      text("紧凑样式 ") + picture(quadratic, "text", alignment, "quadratic-text") + text(" 是这样的 ") + picture(integral, "text", alignment, "integral-text"),
      text("分短发短发 ") + picture(String.raw`(a+b)^n=\sum_{k=0}^{n}\binom{n}{k}a^{n-k}b^k`, "text", alignment, "sum-text"),
      text("上标 ") + picture("L_z", "text", alignment, "subscript") + picture("L^2", "text", alignment, "superscript") + text(" 哈哈哈"),
    ];
    for (const row of rows) paragraphs += `<w:p><w:pPr><w:textAlignment w:val="${alignment}"/><w:spacing w:before="400" w:after="400"/><w:rPr><w:sz w:val="36"/></w:rPr></w:pPr>${row}</w:p>`;
  }
  for (const mode of modes) page(mode, false);
  for (const mode of modes) page(mode, true);
  files["word/document.xml"] = strToU8(`<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}" xmlns:asvg="${ASVG}"><w:body>${paragraphs}<w:sectPr><w:pgSz w:w="18000" w:h="14000"/><w:pgMar w:top="600" w:bottom="600" w:left="600" w:right="600"/></w:sectPr></w:body></w:document>`);
  files["word/settings.xml"] = strToU8(`<w:settings xmlns:w="${W}"><w:zoom w:percent="100"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:docVars>${variables}</w:docVars></w:settings>`);
  files["word/styles.xml"] = strToU8(`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Heiti SC"/><w:sz w:val="36"/></w:rPr></w:style></w:styles>`);
  files["word/_rels/document.xml.rels"] = strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}<Relationship Id="settings" Type="${R}/settings" Target="settings.xml"/><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`);
  files["_rels/.rels"] = strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  files["[Content_Types].xml"] = strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="svg" ContentType="image/svg+xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  run("cargo", ["test", "--manifest-path", "src-tauri/Cargo.toml", "--lib", "freeze_alignment_fixture_vectors", "--", "--ignored"], { ...process.env, VISUALTEX_ALIGNMENT_SVG_DIRECTORY: vectors });
  for (const item of cases) files[`word/media/f${item.index}.svg`] = new Uint8Array(readFileSync(join(vectors, `f${item.index}.svg`)));
  writeFileSync(fixturePath, zipSync(files));
  writeFileSync(casesPath, JSON.stringify(cases, null, 2));
  console.log(fixturePath);
}
