import { spawn, spawnSync } from "node:child_process";
import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { join, resolve } from "node:path";
import { strToU8, zipSync } from "fflate";
import { latexToSvg } from "../src/export/runtime.ts";

const repositoryRoot = resolve(new URL("..", import.meta.url).pathname);
const scratchRoot = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch",
);
const runtimeRoot = join(
  homedir(),
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime",
);
const formulaFont = process.env.VT_WORD_FORMULA_FONT || "times";
const expectedFixtureDocumentName = `word-image-omml-baseline-${formulaFont}-${process.pid}-fixture.docx`;
const docxPath = join(scratchRoot, expectedFixtureDocumentName);
const pdfPath = join(scratchRoot, `word-image-omml-baseline-${formulaFont}-fixture.pdf`);
const conversionPdfPath = join(
  scratchRoot,
  `word-image-omml-baseline-${formulaFont}-real-conversion.pdf`,
);
const reportPath = join(scratchRoot, `word-image-omml-baseline-${formulaFont}-fixture.json`);
const runRealConversion =
  process.env.VT_WORD_REAL_OMML_TO_IMAGE === "1";
const runRoundTrip = process.env.VT_WORD_REAL_IMAGE_OMML_ROUNDTRIP === "1";
const skipPdf = process.env.VT_WORD_SKIP_PDF === "1";
const pdfRequestPath = join(runtimeRoot, "document-import-regression-pdf-path.txt");
const pdfStatusPath = join(runtimeRoot, "document-import-regression-pdf-status.txt");
const transparentPng = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAEAQH/8l0Z8QAAAABJRU5ErkJggg==",
  "base64",
);
const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const ASVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const FONT_SIZE_PT = 11;
const REFERENCE_FONT_SIZE_PT = 14;
const TIMES_WIDTH_SCALE = 1.067;
const TIMES_HEIGHT_SCALE = 1.0;
const TEX_VISUAL_SCALE = 1.1;
const TEX_SHALLOW_DESCENT_FLOOR_PT = 1.91;
const [widthScale, heightScale] = formulaFont === "times"
  ? [TIMES_WIDTH_SCALE, TIMES_HEIGHT_SCALE]
  : [TEX_VISUAL_SCALE, TEX_VISUAL_SCALE];
const portOffset = process.pid % 1000;
const vitePort = 6500 + portOffset;
const debugPort = 12500 + portOffset;
const baseUrl = `http://127.0.0.1:${vitePort}`;
const chromeProfile = `/tmp/visualtex-word-baseline-${process.pid}-${Date.now()}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const cases = [
  { id: "x2", latex: String.raw`x^2` },
  { id: "frac", latex: String.raw`\frac{a}{b}` },
  { id: "int", latex: String.raw`\int_0^x f(t)dt` },
  { id: "sum", latex: String.raw`\sum_i a_i` },
  { id: "sqrt", latex: String.raw`\sqrt{x}` },
];

mkdirSync(scratchRoot, { recursive: true });
mkdirSync(runtimeRoot, { recursive: true });

function escapeXml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function run(command, args, timeout = 120_000) {
  const result = spawnSync(command, args, {
    encoding: "utf8",
    timeout,
    maxBuffer: 32 * 1024 * 1024,
  });
  if (result.status !== 0) {
    throw new Error(
      result.stderr.trim() || result.stdout.trim() || `${command} failed`,
    );
  }
  return result.stdout.trimEnd();
}

function appleScript(lines, timeout = 120_000) {
  return run(
    "/usr/bin/osascript",
    lines.flatMap((line) => ["-e", line]),
    timeout,
  );
}

const sleep = (milliseconds) =>
  new Promise((resolvePromise) => setTimeout(resolvePromise, milliseconds));

async function waitFor(url, timeoutMs = 20_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return response;
    } catch {}
    await sleep(100);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

class CdpClient {
  constructor(url) {
    this.url = url;
    this.nextId = 1;
    this.pending = new Map();
  }

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (!message.id) return;
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      if (message.error) pending.reject(new Error(message.error.message));
      else pending.resolve(message.result);
    });
  }

  send(method, params = {}) {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }

  close() {
    this.socket?.close();
  }
}

async function renderProductionOmml(entries) {
  const vite = spawn(
    process.execPath,
    [
      "node_modules/vite/bin/vite.js",
      "--host",
      "127.0.0.1",
      "--port",
      String(vitePort),
      "--strictPort",
    ],
    { cwd: repositoryRoot, stdio: "ignore" },
  );
  let chrome;
  let client;
  try {
    await waitFor(baseUrl);
    chrome = spawn(
      chromePath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${chromeProfile}`,
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const target = targets.find(
      (candidate) =>
        candidate.type === "page" &&
        typeof candidate.webSocketDebuggerUrl === "string",
    );
    if (!target) throw new Error("Headless Chrome exposed no debuggable page");
    client = new CdpClient(target.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    const expression = `(async () => {
      const module = await import(${JSON.stringify(`${baseUrl}/src/office/omml/latexToOmml.ts`)});
      const entries = ${JSON.stringify(entries.map(({ id, latex }) => ({ id, latex })))};
      return entries.map((entry) => ({
        id: entry.id,
        omml: module.latexLinesToOmml(
          [entry.latex],
          "inline",
          "raw",
          { formulaLetterFont: ${JSON.stringify(formulaFont)} },
        ),
      }));
    })()`;
    const evaluated = await client.send("Runtime.evaluate", {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (evaluated.exceptionDetails) {
      throw new Error(
        evaluated.exceptionDetails.exception?.description ||
          evaluated.exceptionDetails.text ||
          "Headless OMML generation failed",
      );
    }
    const value = evaluated.result?.value;
    if (!Array.isArray(value) || value.length !== entries.length) {
      throw new Error("Headless OMML generation returned an incomplete result");
    }
    return value;
  } finally {
    client?.close();
    if (chrome && !chrome.killed) chrome.kill("SIGTERM");
    if (!vite.killed) vite.kill("SIGTERM");
    rmSync(chromeProfile, {
      recursive: true,
      force: true,
      maxRetries: 8,
      retryDelay: 100,
    });
  }
}

function wordInlineBaseline(value) {
  // SVG and Word picture dimensions quantize independently. Apply the same
  // 0.01 pt boundary tolerance as production, not a formula-wide offset.
  return value < 0 ? -Math.floor(-value + 0.51) : 0;
}

function calculateWordGeometry(svg) {
  const referenceWidthPt = svg.width * 0.75 * widthScale;
  const referenceHeightPt = svg.height * 0.75 * heightScale;
  const descentRatio = Math.max(
    0,
    Math.min(1, (svg.height - svg.baseline) / svg.height),
  );
  const measuredReferenceBaselinePt = -referenceHeightPt * descentRatio;
  const referenceBaselinePt =
    formulaFont !== "times" &&
    measuredReferenceBaselinePt < 0 &&
    measuredReferenceBaselinePt > -TEX_SHALLOW_DESCENT_FLOOR_PT
      ? -TEX_SHALLOW_DESCENT_FLOOR_PT
      : measuredReferenceBaselinePt;
  const pointScale = FONT_SIZE_PT / REFERENCE_FONT_SIZE_PT;
  const widthPt = referenceWidthPt * pointScale;
  const heightPt = referenceHeightPt * pointScale;
  const baselineRawPt = referenceBaselinePt * heightPt / referenceHeightPt;
  const baselinePt = wordInlineBaseline(baselineRawPt);
  return {
    widthPt,
    heightPt,
    baselinePt,
    baselineRawPt,
    referenceWidthPt,
    referenceHeightPt,
    referenceBaselinePt,
  };
}

function pictureRun(item, index) {
  const widthEmu = Math.round(item.geometry.widthPt * 12_700);
  const heightEmu = Math.round(item.geometry.heightPt * 12_700);
  const positionHalfPoints = item.geometry.baselinePt * 2;
  const pngRel = `rIdPng${index}`;
  const svgRel = `rIdSvg${index}`;
  return `<w:r><w:rPr><w:noProof/><w:position w:val="${positionHalfPoints}"/></w:rPr><w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${widthEmu}" cy="${heightEmu}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${index}" name="VT-${escapeXml(item.id)}" descr="VT-${escapeXml(item.id)}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="${index}" name="${escapeXml(item.id)}.svg"/>` +
    `<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>` +
    `<a:blip r:embed="${pngRel}" cstate="print"><a:extLst>` +
    `<a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">` +
    `<asvg:svgBlip r:embed="${svgRel}"/></a:ext></a:extLst></a:blip>` +
    `<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

function createFixturePackage(items) {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
  const packageRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  const documentRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${items.map((item, index) => {
  const id = index + 1;
  return `  <Relationship Id="rIdPng${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${item.id}.png"/>\n` +
    `  <Relationship Id="rIdSvg${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${item.id}.svg"/>`;
}).join("\n")}
</Relationships>`;
  const paragraphs = items.map((item, index) => {
    const label = `VT${index + 1}`;
    return `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr>` +
      `<w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>${label}</w:t></w:r>` +
      `<w:r><w:t xml:space="preserve">  </w:t></w:r>` +
      item.omml +
      `<w:r><w:t xml:space="preserve">    </w:t></w:r>` +
      pictureRun(item, index + 1) +
      `</w:p>`;
  }).join("");
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${WORD_NS}" xmlns:m="${MATH_NS}" xmlns:r="${REL_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}" xmlns:asvg="${ASVG_NS}">
  <w:body>${paragraphs}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body>
</w:document>`;
  const files = {
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(packageRelationships),
    "word/document.xml": strToU8(documentXml),
    "word/_rels/document.xml.rels": strToU8(documentRelationships),
  };
  items.forEach((item) => {
    files[`word/media/${item.id}.png`] = new Uint8Array(transparentPng);
    files[`word/media/${item.id}.svg`] = strToU8(item.svg.svg);
  });
  return zipSync(files, { level: 6 });
}

function parseScalar(value) {
  if (value === "missing value" || value === "") return null;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : value;
}

function collectHiddenWordLayout(documentName) {
  const raw = appleScript([
    'tell application "Microsoft Word"',
    `set fixtureDocument to document ${JSON.stringify(documentName)}`,
    'set outputText to "DOC" & tab & (name of fixtureDocument) & tab & (count of inline shapes of fixtureDocument) & tab & (count of math objects of fixtureDocument) & linefeed',
    "repeat with shapeIndex from 1 to (count of inline shapes of fixtureDocument)",
    "set formulaShape to inline shape shapeIndex of fixtureDocument",
    "set formulaRange to text object of formulaShape",
    "set paragraphRange to text object of paragraph 1 of formulaRange",
    'set outputText to outputText & "IMAGE" & tab & shapeIndex & tab & (start of content of formulaRange) & tab & (end of content of formulaRange) & tab & (start of content of paragraphRange) & tab & (end of content of paragraphRange) & tab & (font position of font object of formulaRange) & tab & (width of formulaShape) & tab & (height of formulaShape) & tab & (get range information formulaRange information type horizontal position relative to page) & tab & (get range information formulaRange information type vertical position relative to page) & tab & (count of fields of formulaRange) & tab & (count of math objects of formulaRange) & linefeed',
    "end repeat",
    "repeat with mathIndex from 1 to (count of math objects of fixtureDocument)",
    "set formulaMath to math object mathIndex of fixtureDocument",
    "set formulaRange to text range of formulaMath",
    "set paragraphRange to text object of paragraph 1 of formulaRange",
    "set formulaEndPosition to end of content of formulaRange",
    "set formulaEndRange to create range fixtureDocument start formulaEndPosition end formulaEndPosition",
    'set outputText to outputText & "OMML" & tab & mathIndex & tab & (start of content of formulaRange) & tab & (end of content of formulaRange) & tab & (start of content of paragraphRange) & tab & (end of content of paragraphRange) & tab & (display type of formulaMath) & tab & (get range information formulaRange information type horizontal position relative to page) & tab & (get range information formulaEndRange information type horizontal position relative to page) & tab & (get range information formulaRange information type vertical position relative to page) & linefeed',
    "end repeat",
    "return outputText",
    "end tell",
  ]);
  const lines = raw.split(/\r?\n/).filter(Boolean);
  const [header, ...rows] = lines;
  const headerFields = header.split("\t");
  const layout = {
    document: {
      name: headerFields[1],
      imageCount: parseScalar(headerFields[2]),
      ommlCount: parseScalar(headerFields[3]),
    },
    images: [],
    omml: [],
  };
  for (const row of rows) {
    const fields = row.split("\t");
    if (fields[0] === "IMAGE") {
      layout.images.push({
        index: parseScalar(fields[1]),
        rangeStart: parseScalar(fields[2]),
        rangeEnd: parseScalar(fields[3]),
        paragraphStart: parseScalar(fields[4]),
        paragraphEnd: parseScalar(fields[5]),
        fontPositionPt: parseScalar(fields[6]),
        widthPt: parseScalar(fields[7]),
        heightPt: parseScalar(fields[8]),
        pageXPt: parseScalar(fields[9]),
        pageYPt: parseScalar(fields[10]),
        fieldCount: parseScalar(fields[11]),
        ommlCount: parseScalar(fields[12]),
      });
    } else if (fields[0] === "OMML") {
      layout.omml.push({
        index: parseScalar(fields[1]),
        rangeStart: parseScalar(fields[2]),
        rangeEnd: parseScalar(fields[3]),
        paragraphStart: parseScalar(fields[4]),
        paragraphEnd: parseScalar(fields[5]),
        displayType: fields[6],
        pageXPt: parseScalar(fields[7]),
        pageEndXPt: parseScalar(fields[8]),
        pageYPt: parseScalar(fields[9]),
      });
    }
  }
  return layout;
}

function collectRasterGeometry(targetPdfPath = pdfPath) {
  const raw = run(
    "/usr/bin/swift",
    [
      join(repositoryRoot, "scripts/pdf_formula_geometry.swift"),
      targetPdfPath,
      "--raster-only",
    ],
    120_000,
  );
  return JSON.parse(raw);
}

function objectInkBounds(
  rasterGeometry,
  minX,
  maxX,
  targetCenterY,
  verticalRadiusPt,
) {
  if (
    ![minX, maxX, targetCenterY, verticalRadiusPt].every(Number.isFinite) ||
    maxX <= minX ||
    verticalRadiusPt <= 0
  ) return null;
  const components = (
    rasterGeometry.rasterComponents ??
    (rasterGeometry.rasterBands ?? []).flatMap((band) => band.components ?? [])
  ).filter(
    (component) =>
      component.centerX >= minX - 0.5 &&
      component.centerX <= maxX + 0.5 &&
      component.centerY >= targetCenterY - verticalRadiusPt &&
      component.centerY <= targetCenterY + verticalRadiusPt,
  );
  if (components.length === 0) return null;
  const minY = Math.min(...components.map((component) => component.minY));
  const maxY = Math.max(...components.map((component) => component.maxY));
  return {
    minY,
    maxY,
    height: maxY - minY,
    centerY: (minY + maxY) / 2,
    componentCount: components.length,
    targetCenterY,
    verticalRadiusPt,
  };
}

function exportHiddenWordPdf(documentName, targetPdfPath) {
  rmSync(targetPdfPath, { force: true });
  rmSync(pdfStatusPath, { force: true });
  writeFileSync(pdfRequestPath, targetPdfPath, { mode: 0o600 });
  appleScript([
    'tell application "Microsoft Word"',
    `set fixtureDocument to document ${JSON.stringify(documentName)}`,
    "activate object fixtureDocument",
    'run VB macro macro name "VisualTeX_ExportActiveDocumentPdfForRegression"',
    "end tell",
  ]);
  appleScript([
    `tell application ${JSON.stringify(frontmostBefore)} to activate`,
  ]);
  const status = readFileSync(pdfStatusPath, "utf8").trim();
  if (!status.startsWith("ok|") || !readFileSync(targetPdfPath).length) {
    throw new Error(`Hidden Word PDF export failed: ${status}`);
  }
}

async function convertHiddenImagesToOmml(documentName, itemCount) {
  for (let convertedCount = 0; convertedCount < itemCount; convertedCount += 1) {
    const shapeIndex = convertedCount + 1;
    appleScript([
      'tell application "Microsoft Word"',
      `set fixtureDocument to document ${JSON.stringify(documentName)}`,
      `set formulaRange to text object of inline shape ${shapeIndex} of fixtureDocument`,
      "activate object fixtureDocument",
      "select formulaRange",
      'run VB macro macro name "VisualTeX_ConvertSelectedToNativeEquation"',
      "end tell",
    ]);
    const expectedImages = itemCount * 2 - convertedCount - 1;
    const expectedOmml = convertedCount + 1;
    const deadline = Date.now() + 45_000;
    while (Date.now() < deadline) {
      const layout = collectHiddenWordLayout(documentName);
      if (
        layout.document.imageCount === expectedImages &&
        layout.document.ommlCount === expectedOmml
      ) break;
      await sleep(250);
    }
    const layout = collectHiddenWordLayout(documentName);
    if (
      layout.document.imageCount !== expectedImages ||
      layout.document.ommlCount !== expectedOmml
    ) {
      throw new Error(
        `Timed out converting image ${convertedCount + 1}/${itemCount} to OMML: images=${layout.document.imageCount}, omml=${layout.document.ommlCount}`,
      );
    }
  }
}

async function convertHiddenOmmlToImages(documentName, itemCount) {
  for (let convertedCount = 0; convertedCount < itemCount; convertedCount += 1) {
    appleScript([
      'tell application "Microsoft Word"',
      `set fixtureDocument to document ${JSON.stringify(documentName)}`,
      "set formulaRange to text range of math object 1 of fixtureDocument",
      "activate object fixtureDocument",
      "select formulaRange",
      'run VB macro macro name "VisualTeX_ConvertSelectedToImageFormula"',
      "end tell",
    ]);
    const expectedImages = itemCount + convertedCount + 1;
    const expectedOmml = itemCount - convertedCount - 1;
    const deadline = Date.now() + 45_000;
    while (Date.now() < deadline) {
      const raw = appleScript([
        'tell application "Microsoft Word"',
        `set fixtureDocument to document ${JSON.stringify(documentName)}`,
        "return ((count of inline shapes of fixtureDocument) as text) & tab & ((count of math objects of fixtureDocument) as text)",
        "end tell",
      ]);
      const [imageCount, ommlCount] = raw.split("\t").map(Number);
      if (imageCount === expectedImages && ommlCount === expectedOmml) break;
      await sleep(250);
    }
    const layout = collectHiddenWordLayout(documentName);
    if (
      layout.document.imageCount !== expectedImages ||
      layout.document.ommlCount !== expectedOmml
    ) {
      throw new Error(
        `Timed out converting OMML ${convertedCount + 1}/${itemCount}: images=${layout.document.imageCount}, omml=${layout.document.ommlCount}`,
      );
    }
  }
}

const frontmostBefore = appleScript([
  'tell application "System Events" to return name of first application process whose frontmost is true',
]);
if (frontmostBefore === "Microsoft Word") {
  throw new Error(
    "Refusing to run the hidden Word fixture while Word is frontmost; this probe must not disturb the user's Word page.",
  );
}

const renderedOmml = await renderProductionOmml(cases);
const ommlById = new Map(renderedOmml.map((entry) => [entry.id, entry.omml]));
const items = cases.map((entry) => {
  const svg = latexToSvg(entry.latex, {
    displayMode: false,
    fontSizePt: REFERENCE_FONT_SIZE_PT,
    paddingPx: 1,
    background: "transparent",
    forceExplicitBlack: true,
    formulaLetterFont: formulaFont,
  });
  const omml = ommlById.get(entry.id);
  if (typeof omml !== "string" || !omml.includes("<m:oMath")) {
    throw new Error(`Production OMML is missing for ${entry.id}`);
  }
  return {
    ...entry,
    svg,
    omml,
    geometry: calculateWordGeometry(svg),
  };
});

rmSync(docxPath, { force: true });
rmSync(pdfPath, { force: true });
rmSync(pdfStatusPath, { force: true });
writeFileSync(docxPath, createFixturePackage(items));
writeFileSync(pdfRequestPath, pdfPath, { mode: 0o600 });

let sourceDocumentName = "";
let fixtureDocumentName = expectedFixtureDocumentName;
try {
  const openResult = appleScript([
    'tell application "Microsoft Word"',
    'if not (exists active document) then error "Microsoft Word has no source document to restore"',
    "set sourceDocument to active document",
    "set sourceDocumentName to name of sourceDocument",
    `open POSIX file ${JSON.stringify(docxPath)} read only false add to recent files false`,
    "set fixtureDocument to active document",
    "set fixtureWindow to active window of fixtureDocument",
    "set visible of fixtureWindow to false",
    `set font size of font object of text object of fixtureDocument to ${FONT_SIZE_PT}`,
    'return sourceDocumentName & tab & (name of fixtureDocument)',
    "end tell",
  ]);
  [sourceDocumentName, fixtureDocumentName] = openResult.split("\t");
  if (!fixtureDocumentName) fixtureDocumentName = expectedFixtureDocumentName;
  const layout = collectHiddenWordLayout(fixtureDocumentName);
  if (!skipPdf) exportHiddenWordPdf(fixtureDocumentName, pdfPath);
  const rasterGeometry = skipPdf ? null : collectRasterGeometry(pdfPath);
  const measurements = items.map((item, index) => {
    const omml = layout.omml[index];
    const image = layout.images[index];
    const imageTargetCenterY = image
      ? image.pageYPt + image.heightPt / 2
      : null;
    const imageInk = image && rasterGeometry
      ? objectInkBounds(
          rasterGeometry,
          image.pageXPt,
          image.pageXPt + image.widthPt,
          imageTargetCenterY,
          Math.max(3, image.heightPt / 2 + 2),
        )
      : null;
    const ommlInk = omml && image && rasterGeometry
      ? objectInkBounds(
          rasterGeometry,
          omml.pageXPt,
          omml.pageEndXPt,
          imageInk?.centerY ?? imageTargetCenterY,
          Math.max(4, image.heightPt / 2 + 3),
        )
      : null;
    return {
      id: item.id,
      latex: item.latex,
      expectedGeometry: item.geometry,
      omml,
      image,
      ommlInk,
      imageInk,
      inkDelta: ommlInk && imageInk
        ? {
            topPt: imageInk.minY - ommlInk.minY,
            bottomPt: imageInk.maxY - ommlInk.maxY,
            centerPt: imageInk.centerY - ommlInk.centerY,
            heightPt: imageInk.height - ommlInk.height,
          }
        : null,
    };
  });
  if (
    layout.document.imageCount !== items.length ||
    layout.document.ommlCount !== items.length
  ) {
    throw new Error(
      `Hidden Word fixture structure mismatch: images=${layout.document.imageCount}, omml=${layout.document.ommlCount}`,
    );
  }
  for (const [index, measurement] of measurements.entries()) {
    const expectedPosition = items[index].geometry.baselinePt;
    if (!measurement.image || !measurement.omml) {
      throw new Error(`Missing Word layout object for ${measurement.id}`);
    }
    if (measurement.image.fontPositionPt !== expectedPosition) {
      throw new Error(
        `Word baseline mismatch for ${measurement.id}: expected ${expectedPosition} pt, got ${measurement.image.fontPositionPt} pt`,
      );
    }
    if (!skipPdf && (!measurement.imageInk || !measurement.ommlInk || !measurement.inkDelta)) {
      throw new Error(`Missing PDF ink geometry for ${measurement.id}`);
    }
    if (!skipPdf && Math.abs(measurement.inkDelta.centerPt) > 1) {
      throw new Error(
        `PDF ink center mismatch for ${measurement.id}: ${measurement.inkDelta.centerPt.toFixed(3)} pt`,
      );
    }
    if (
      !skipPdf && (
      Math.abs(measurement.inkDelta.topPt) > 4 ||
      Math.abs(measurement.inkDelta.bottomPt) > 4)
    ) {
      throw new Error(
        `PDF ink bounds mismatch for ${measurement.id}: top=${measurement.inkDelta.topPt.toFixed(3)} pt, bottom=${measurement.inkDelta.bottomPt.toFixed(3)} pt`,
      );
    }
  }

  let realConversion = null;
  if (runRealConversion) {
    await convertHiddenOmmlToImages(fixtureDocumentName, items.length);
    const convertedLayout = collectHiddenWordLayout(fixtureDocumentName);
    if (
      convertedLayout.document.imageCount !== items.length * 2 ||
      convertedLayout.document.ommlCount !== 0
    ) {
      throw new Error(
        `Real OMML conversion structure mismatch: images=${convertedLayout.document.imageCount}, omml=${convertedLayout.document.ommlCount}`,
      );
    }
    if (!skipPdf) exportHiddenWordPdf(fixtureDocumentName, conversionPdfPath);
    const convertedRaster = skipPdf ? null : collectRasterGeometry(conversionPdfPath);
    const comparisons = items.map((item, index) => {
      const convertedImage = convertedLayout.images[index * 2];
      const referenceImage = convertedLayout.images[index * 2 + 1];
      if (!convertedImage || !referenceImage) {
        throw new Error(`Missing converted/reference image pair for ${item.id}`);
      }
      for (const [kind, image] of [
        ["converted", convertedImage],
        ["reference", referenceImage],
      ]) {
        if (image.fieldCount !== 0 || image.ommlCount !== 0) {
          throw new Error(
            `${kind} image for ${item.id} is not a pure InlineShape: fields=${image.fieldCount}, omml=${image.ommlCount}`,
          );
        }
      }
      if (
        convertedImage.fontPositionPt !== referenceImage.fontPositionPt ||
        Math.abs(convertedImage.widthPt - referenceImage.widthPt) > 0.1 ||
        Math.abs(convertedImage.heightPt - referenceImage.heightPt) > 0.1
      ) {
        throw new Error(
          `Real OMML conversion geometry drift for ${item.id}: converted=${convertedImage.widthPt}x${convertedImage.heightPt}@${convertedImage.fontPositionPt}, reference=${referenceImage.widthPt}x${referenceImage.heightPt}@${referenceImage.fontPositionPt}`,
        );
      }
      const convertedInk = convertedRaster ? objectInkBounds(
        convertedRaster,
        convertedImage.pageXPt,
        convertedImage.pageXPt + convertedImage.widthPt,
        convertedImage.pageYPt + convertedImage.heightPt / 2,
        Math.max(3, convertedImage.heightPt / 2 + 2),
      ) : null;
      const referenceInk = convertedRaster ? objectInkBounds(
        convertedRaster,
        referenceImage.pageXPt,
        referenceImage.pageXPt + referenceImage.widthPt,
        referenceImage.pageYPt + referenceImage.heightPt / 2,
        Math.max(3, referenceImage.heightPt / 2 + 2),
      ) : null;
      if (!skipPdf && (!convertedInk || !referenceInk)) {
        throw new Error(`Missing real-conversion PDF ink for ${item.id}`);
      }
      const inkDelta = convertedInk && referenceInk ? {
        topPt: convertedInk.minY - referenceInk.minY,
        bottomPt: convertedInk.maxY - referenceInk.maxY,
        centerPt: convertedInk.centerY - referenceInk.centerY,
        heightPt: convertedInk.height - referenceInk.height,
      } : null;
      if (
        !skipPdf && (
        Math.abs(inkDelta.topPt) > 0.5 ||
        Math.abs(inkDelta.bottomPt) > 0.5 ||
        Math.abs(inkDelta.centerPt) > 0.5)
      ) {
        throw new Error(
          `Real OMML conversion PDF ink drift for ${item.id}: top=${inkDelta.topPt.toFixed(3)}, bottom=${inkDelta.bottomPt.toFixed(3)}, center=${inkDelta.centerPt.toFixed(3)} pt`,
        );
      }
      return {
        id: item.id,
        convertedImage,
        referenceImage,
        convertedInk,
        referenceInk,
        inkDelta,
      };
    });
    let roundTrip = null;
    if (runRoundTrip) {
      await convertHiddenImagesToOmml(fixtureDocumentName, items.length);
      const nativeLayout = collectHiddenWordLayout(fixtureDocumentName);
      if (
        nativeLayout.document.imageCount !== items.length ||
        nativeLayout.document.ommlCount !== items.length
      ) {
        throw new Error(
          `Image-to-OMML round-trip structure mismatch: images=${nativeLayout.document.imageCount}, omml=${nativeLayout.document.ommlCount}`,
        );
      }
      await convertHiddenOmmlToImages(fixtureDocumentName, items.length);
      const imageLayout = collectHiddenWordLayout(fixtureDocumentName);
      if (
        imageLayout.document.imageCount !== items.length * 2 ||
        imageLayout.document.ommlCount !== 0
      ) {
        throw new Error(
          `OMML-to-image round-trip structure mismatch: images=${imageLayout.document.imageCount}, omml=${imageLayout.document.ommlCount}`,
        );
      }
      roundTrip = { nativeLayout, imageLayout };
    }
    realConversion = {
      pdfPath: conversionPdfPath,
      layout: convertedLayout,
      comparisons,
      roundTrip,
    };
  }

  const report = {
    generatedAt: new Date().toISOString(),
    formulaFont,
    frontmostBefore,
    sourceDocumentName,
    fixtureDocumentName,
    paths: { docxPath, pdfPath, conversionPdfPath, reportPath },
    layout,
    measurements,
    realConversion,
  };
  writeFileSync(reportPath, `${JSON.stringify(report, null, 2)}\n`, "utf8");
  console.log(JSON.stringify({
    reportPath,
    direct: measurements.map((measurement) => ({
      id: measurement.id,
      fontPositionPt: measurement.image?.fontPositionPt,
      inkDelta: measurement.inkDelta,
    })),
    realConversion: realConversion
      ? realConversion.comparisons.map((comparison) => ({
          id: comparison.id,
          fontPositionPt: comparison.convertedImage.fontPositionPt,
          widthDeltaPt:
            comparison.convertedImage.widthPt -
            comparison.referenceImage.widthPt,
          heightDeltaPt:
            comparison.convertedImage.heightPt -
            comparison.referenceImage.heightPt,
          inkDelta: comparison.inkDelta,
        }))
      : null,
  }, null, 2));
} finally {
  if (fixtureDocumentName) {
    try {
      appleScript([
        'tell application "Microsoft Word"',
        `if exists document ${JSON.stringify(fixtureDocumentName)} then close document ${JSON.stringify(fixtureDocumentName)} saving no`,
        fixtureDocumentName !== expectedFixtureDocumentName
          ? `if exists document ${JSON.stringify(expectedFixtureDocumentName)} then close document ${JSON.stringify(expectedFixtureDocumentName)} saving no`
          : "",
        sourceDocumentName
          ? `if exists document ${JSON.stringify(sourceDocumentName)} then activate object document ${JSON.stringify(sourceDocumentName)}`
          : "",
        "end tell",
      ].filter(Boolean), 30_000);
    } catch {}
  }
  try {
    appleScript([
      `tell application ${JSON.stringify(frontmostBefore)} to activate`,
    ], 30_000);
  } catch {}
}
