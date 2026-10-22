import { execFileSync, spawnSync } from "node:child_process";
import { randomUUID } from "node:crypto";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { DOMParser } from "@xmldom/xmldom";
import { strToU8, unzipSync, zipSync } from "fflate";
import { parseLatexMarkdownDocument } from "../src/office/documentImport/documentImportParser";
import { normalizeFormulaEditorDocument } from "../src/office/shared/formulaEditorDocument";
import {
  createFormulaMetadata,
  encodeFormulaMetadata,
} from "../src/office/shared/formulaMetadata";
import { renderOfficeFormulaArtifacts } from "../src/office/shared/formulaRenderArtifacts";

const root = resolve(new URL("..", import.meta.url).pathname);
const fixturePath = join(root, "scripts/fixtures/document-import-mixed-math-20260810.txt");
const wordAddinPath = resolve(
  process.argv[process.argv.indexOf("--word-addin") + 1] ||
    join(root, "office/macos-offline/resources/VisualTeX.dotm"),
);
const runtimeRoot = join(
  homedir(),
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime",
);
const sessionsRoot = join(runtimeRoot, "OfficeSessions");
const nativeRoot = join(runtimeRoot, "NativeDocuments");
const wordStartup = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word",
);
const installedAddin = join(wordStartup, "VisualTeX.dotm");
const addinBackup = join(tmpdir(), `visualtex-mixed-word-${process.pid}.dotm`);
const mathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const unitSeparator = "|||VISUALTEX|||";
const officeScratchRoot = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch",
);

globalThis.DOMParser ??= DOMParser;
const domProbe = new DOMParser().parseFromString("<root/>", "application/xml");
const documentPrototype = Object.getPrototypeOf(domProbe);
const elementPrototype = Object.getPrototypeOf(domProbe.documentElement);
if (typeof documentPrototype.querySelector !== "function") {
  documentPrototype.querySelector = function querySelector(name: string) {
    return this.getElementsByTagName(name)?.item(0) ?? null;
  };
}
if (!("children" in elementPrototype)) {
  Object.defineProperty(elementPrototype, "children", {
    configurable: true,
    get() {
      return Array.from(this.childNodes ?? []).filter((node: any) => node.nodeType === 1);
    },
  });
}

function run(program: string, args: string[], timeout = 60_000) {
  return execFileSync(program, args, {
    encoding: "utf8",
    timeout,
    maxBuffer: 32 * 1024 * 1024,
  }).trim();
}

function runAppleScript(lines: string[], timeout = 60_000) {
  return run(
    "/usr/bin/osascript",
    lines.flatMap((line) => ["-e", line]),
    timeout,
  );
}

function manifestText(entries: Array<[string, string | number]>) {
  const seen = new Set<string>();
  return `${entries
    .map(([key, rawValue]) => {
      if (!/^[A-Za-z0-9]+$/.test(key) || seen.has(key)) {
        throw new Error(`invalid manifest key ${key}`);
      }
      seen.add(key);
      const value = String(rawValue);
      if (/[\r\n\0]/.test(value)) throw new Error(`invalid manifest value ${key}`);
      return `${key}=${value}`;
    })
    .join("\n")}\n`;
}

function base64Url(value: string) {
  return Buffer.from(value, "utf8").toString("base64url");
}

function docxBytes(ommlFragments: string[]) {
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    "</Types>";
  const relationships =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    "</Relationships>";
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${wordNamespace}" xmlns:m="${mathNamespace}"><w:body>` +
    ommlFragments.map((omml) => `<w:p>${omml}</w:p>`).join("") +
    "<w:sectPr/></w:body></w:document>";
  return zipSync(
    {
      "[Content_Types].xml": strToU8(contentTypes),
      "_rels/.rels": strToU8(relationships),
      "word/document.xml": strToU8(documentXml),
    },
    { level: 6 },
  );
}

function currentSessions() {
  if (!existsSync(sessionsRoot)) return new Set<string>();
  return new Set(
    readdirSync(sessionsRoot).filter((name) =>
      /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(name),
    ),
  );
}

async function waitForSession(before: Set<string>, timeoutMs = 30_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const fresh = [...currentSessions()].filter((id) => !before.has(id));
    if (fresh.length === 1 && existsSync(join(sessionsRoot, fresh[0], "request.json"))) {
      return fresh[0];
    }
    await new Promise((resolvePromise) => setTimeout(resolvePromise, 100));
  }
  throw new Error("Word did not create one documentImport session");
}

function paragraphEntries(block: any, itemIndex: number): Array<[string, string | number]> {
  if (!block.paragraphId) return [];
  const prefix = `item${itemIndex}`;
  return [
    [`${prefix}paragraphId`, block.paragraphId],
    [`${prefix}paragraphStyle`, block.paragraphStyle || "normal"],
    [`${prefix}paragraphAlignment`, block.paragraphAlignment || "left"],
    [`${prefix}listKind`, block.listKind || "none"],
    [`${prefix}listLevel`, block.listLevel || 0],
    [`${prefix}paragraphStart`, block.paragraphStart ? 1 : 0],
    [`${prefix}paragraphEnd`, block.paragraphEnd ? 1 : 0],
  ];
}

function formulaMetadata(
  formulaId: string,
  canonicalLatex: string,
  displayMode: "inline" | "block",
  lines: Array<{ id: string; latex: string }>,
  codeFormat: any,
) {
  return encodeFormulaMetadata(
    createFormulaMetadata({
      formulaId,
      title: displayMode === "inline" ? "Mixed inline formula" : "Mixed display formula",
      lines,
      codeFormat,
      sourceLatex: canonicalLatex,
      displayMode,
      numbered: false,
      fontSizePt: 11,
    }),
  );
}

function assertStructuredOmml(
  xml: string,
  label: string,
  allowLiteralLatex = false,
) {
  const required = [
    ["fraction", "<m:f>"],
    ["radical", "<m:rad>"],
    ["superscript", "<m:sSup>"],
    ["n-ary", "<m:nary>"],
    ["matrix", "<m:m>"],
  ] as const;
  for (const [feature, token] of required) {
    if (!xml.includes(token)) {
      throw new Error(`${label} is missing real OMML ${feature} structure ${token}`);
    }
  }
  if (
    !allowLiteralLatex &&
    (xml.includes("\\frac") ||
      xml.includes("\\sqrt") ||
      xml.includes("\\begin{pmatrix}"))
  ) {
    throw new Error(`${label} still contains literal LaTeX commands inside Word OMML`);
  }
}

const source = readFileSync(fixturePath, "utf8");
const blocks = parseLatexMarkdownDocument(source, "auto", 11);
const formulaCount = blocks.filter((block) => block.kind === "formula").length;
if (formulaCount < 50) throw new Error(`fixture parsed only ${formulaCount} formulas`);

let addinBackedUp = false;
const createdNativeFiles: string[] = [];
const createdSessions: string[] = [];

async function runCase(label: string, corruptFormulaOrdinal = 0) {
  const before = currentSessions();
  const documentName = runAppleScript([
    'tell application "Microsoft Word"',
    "make new document",
    "set documentObject to active document",
    "activate object documentObject",
    'run VB macro macro name "VisualTeX_InsertLatexMarkdownDocument"',
    "return name of documentObject",
    "end tell",
  ], 60_000);
  const sessionId = await waitForSession(before);
  createdSessions.push(sessionId);
  const sessionDir = join(sessionsRoot, sessionId);
  const request = JSON.parse(readFileSync(join(sessionDir, "request.json"), "utf8"));
  if (request.operation !== "documentImport") throw new Error(`${label}: unexpected request`);
  spawnSync("/usr/bin/pkill", ["-x", "visualtex"], { timeout: 10_000 });

  const entries: Array<[string, string | number]> = [];
  const ommlFragments: string[] = [];
  const formulas: Array<{ latex: string; displayMode: "inline" | "block" }> = [];
  let formulaOrdinal = 0;
  for (let itemIndex = 0; itemIndex < blocks.length; itemIndex += 1) {
    const block: any = blocks[itemIndex];
    const prefix = `item${itemIndex}`;
    if (block.kind === "text") {
      entries.push([`${prefix}kind`, "text"], [`${prefix}textBase64`, base64Url(block.text)]);
      entries.push(...paragraphEntries(block, itemIndex));
      continue;
    }
    formulaOrdinal += 1;
    const formulaId = randomUUID();
    const displayMode = block.displayMode as "inline" | "block";
    const normalized = normalizeFormulaEditorDocument(
      [{ id: randomUUID(), latex: block.latex.trim() }],
      "raw",
    );
    const artifacts = await renderOfficeFormulaArtifacts({
      lines: normalized.lines,
      codeFormat: normalized.codeFormat,
      displayMode,
      host: "word",
    });
    if (!artifacts.omml) {
      throw new Error(`production renderer returned no OMML for ${block.latex}`);
    }
    const latex = artifacts.canonicalLatex;
    const omml = Buffer.from(artifacts.omml.ommlBase64, "base64url").toString("utf8");
    const nativePath = join(nativeRoot, `${formulaId}.docx`);
    writeFileSync(
      nativePath,
      Buffer.from(artifacts.omml.ommlDocxBase64, "base64url"),
      { mode: 0o600 },
    );
    createdNativeFiles.push(nativePath);
    ommlFragments.push(omml);
    formulas.push({ latex, displayMode });
    const batchIndex = formulaOrdinal === corruptFormulaOrdinal ? formulaCount + 19 : formulaOrdinal;
    entries.push(
      [`${prefix}kind`, "formula"],
      [`${prefix}formulaId`, formulaId],
      [`${prefix}latexBase64`, base64Url(latex)],
      [`${prefix}displayMode`, displayMode],
      [`${prefix}numbered`, 0],
      [`${prefix}fontSizePt`, "11.000000"],
      [
        `${prefix}metadata`,
        formulaMetadata(
          formulaId,
          latex,
          displayMode,
          artifacts.lines,
          artifacts.codeFormat,
        ),
      ],
      [`${prefix}ommlBase64`, artifacts.omml.ommlBase64],
      [`${prefix}nativeDocumentPath`, nativePath],
      [`${prefix}nativeBatchDocumentIndex`, batchIndex],
      [`${prefix}widthPoints`, "11.000000"],
      [`${prefix}heightPoints`, "19.800000"],
      [`${prefix}baseline`, "0.000000"],
      [`${prefix}referenceWidthPt`, "14.000000"],
      [`${prefix}referenceHeightPt`, "14.000000"],
      [`${prefix}referenceBaselinePt`, "0.000000"],
    );
    entries.push(...paragraphEntries(block, itemIndex));
  }

  assertStructuredOmml(
    ommlFragments.join("\n"),
    `${label} production OMML batch`,
  );
  const batchPath = join(sessionDir, "document-native-batch.docx");
  writeFileSync(batchPath, docxBytes(ommlFragments), { mode: 0o600 });
  const manifestPath = join(sessionDir, "document-import.txt");
  writeFileSync(
    manifestPath,
    manifestText([
      ["protocolVersion", 1],
      ["sessionId", sessionId],
      ["operation", "documentImport"],
      ["outputKind", "omml"],
      ["sourceDocumentId", request.sourceDocumentId],
      ["bookmarkName", request.documentImport.bookmarkName],
      ["itemCount", blocks.length],
      ["nativeBatchDocumentPath", batchPath],
      ["nativeBatchFormulaCount", formulaCount],
      ...entries,
    ]),
    { mode: 0o600 },
  );
  writeFileSync(
    join(sessionDir, "dispatch.txt"),
    manifestText([
      ["protocolVersion", 1],
      ["sessionId", sessionId],
      ["action", "documentCommit"],
      ["host", "word"],
      ["sourceDocumentId", request.sourceDocumentId],
      ["bookmarkName", request.documentImport.bookmarkName],
      ["documentImportPath", manifestPath],
    ]),
    { mode: 0o600 },
  );
  writeFileSync(join(sessionsRoot, "word-active-session.txt"), sessionId, { mode: 0o600 });

  const statusPath = join(sessionDir, "word-callback-status.txt");
  rmSync(statusPath, { force: true });
  const startedAt = process.hrtime.bigint();
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(documentName)}`,
    'run VB macro macro name "VisualTeX_ApplyPendingResultForRegression"',
    "end tell",
  ], 180_000);
  const elapsedMs = Number(process.hrtime.bigint() - startedAt) / 1_000_000;
  const status = existsSync(statusPath) ? readFileSync(statusPath, "utf8") : "missing";
  if (!status.startsWith("PASS")) throw new Error(`${label} callback failed:\n${status}`);

  const snapshot = runAppleScript([
    `set separatorText to ${JSON.stringify(unitSeparator)}`,
    'tell application "Microsoft Word"',
    `set documentObject to document ${JSON.stringify(documentName)}`,
    "set documentText to content of text object of documentObject",
    'return ((count of math objects of documentObject) as text) & separatorText & ((count of inline shapes of documentObject) as text) & separatorText & documentText',
    "end tell",
  ], 30_000).split(unitSeparator);
  const mathCount = Number(snapshot[0]);
  const imageCount = Number(snapshot[1]);
  const documentText = snapshot.slice(2).join(unitSeparator);
  const expectedMathCount = formulaCount - (corruptFormulaOrdinal ? 1 : 0);
  if (mathCount !== expectedMathCount || imageCount !== 0) {
    throw new Error(`${label} object count mismatch ${JSON.stringify({ mathCount, imageCount, expectedMathCount })}`);
  }
  let fallbackPreserved = true;
  if (corruptFormulaOrdinal) {
    const failed = formulas[corruptFormulaOrdinal - 1];
    const fallback = failed.displayMode === "block" ? `$$${failed.latex}$$` : `$${failed.latex}$`;
    fallbackPreserved = documentText.includes(fallback);
    if (!fallbackPreserved) throw new Error(`${label} did not preserve failed LaTeX`);
  }

  const savedDocxPath = join(
    officeScratchRoot,
    `document-import-mixed-${label}-${process.pid}.docx`,
  );
  rmSync(savedDocxPath, { force: true });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(documentName)}`,
    "set documentObject to active document",
    `save as documentObject file name ${JSON.stringify(savedDocxPath)} file format format document default add to recent files false`,
    "end tell",
  ], 30_000);
  if (!existsSync(savedDocxPath)) {
    throw new Error(`${label} did not save a DOCX for OpenXML inspection`);
  }
  const savedPackage = unzipSync(new Uint8Array(readFileSync(savedDocxPath)));
  const documentXmlBytes = savedPackage["word/document.xml"];
  if (!documentXmlBytes) {
    throw new Error(`${label} saved DOCX has no word/document.xml`);
  }
  const savedDocumentXml = Buffer.from(documentXmlBytes).toString("utf8");
  assertStructuredOmml(
    savedDocumentXml,
    `${label} saved Word document`,
    Boolean(corruptFormulaOrdinal),
  );

  runAppleScript([
    'tell application "Microsoft Word"',
    "close active document saving no",
    "end tell",
  ], 20_000);
  rmSync(savedDocxPath, { force: true });
  return {
    label,
    elapsedMs,
    itemCount: blocks.length,
    formulaCount,
    mathCount,
    fallbackPreserved,
    structuredOmmlVerified: true,
  };
}

try {
  mkdirSync(sessionsRoot, { recursive: true });
  mkdirSync(nativeRoot, { recursive: true });
  mkdirSync(wordStartup, { recursive: true });
  mkdirSync(officeScratchRoot, { recursive: true });
  try { runAppleScript(['tell application "Microsoft Word" to quit saving no'], 20_000); } catch {}
  spawnSync("/usr/bin/killall", ["Microsoft Word"], { timeout: 10_000 });
  spawnSync("/bin/sleep", ["2"], { timeout: 5_000 });
  if (existsSync(installedAddin)) {
    copyFileSync(installedAddin, addinBackup);
    addinBackedUp = true;
  }
  copyFileSync(wordAddinPath, installedAddin);
  run("/usr/bin/open", ["-a", "Microsoft Word"], 10_000);
  const waitStarted = Date.now();
  while (Date.now() - waitStarted < 30_000) {
    try {
      if (runAppleScript(['tell application "Microsoft Word" to return "ready"'], 5_000) === "ready") break;
    } catch {}
    await new Promise((resolvePromise) => setTimeout(resolvePromise, 250));
  }
  const normal = await runCase("normal");
  const isolatedFailure = await runCase("one-formula-failure", 10);
  console.log(JSON.stringify({ status: "PASS", formulaCount, normal, isolatedFailure }, null, 2));
} finally {
  try { runAppleScript(['tell application "Microsoft Word" to quit saving no'], 20_000); } catch {}
  spawnSync("/usr/bin/killall", ["Microsoft Word"], { timeout: 10_000 });
  spawnSync("/usr/bin/pkill", ["-x", "visualtex"], { timeout: 10_000 });
  for (const path of createdNativeFiles) rmSync(path, { force: true });
  for (const sessionId of createdSessions) rmSync(join(sessionsRoot, sessionId), { recursive: true, force: true });
  rmSync(join(sessionsRoot, "word-active-session.txt"), { force: true });
  if (addinBackedUp && existsSync(addinBackup)) copyFileSync(addinBackup, installedAddin);
  rmSync(addinBackup, { force: true });
}
