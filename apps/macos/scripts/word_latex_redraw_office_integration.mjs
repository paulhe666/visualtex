import { execFileSync } from "node:child_process";
import {
  existsSync,
  readFileSync,
  readdirSync,
} from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { decodeFormulaMetadata } from "../src/office/shared/formulaMetadata.ts";

const runtimeRoot = join(
  homedir(),
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime",
);
const sessionsRoot = join(runtimeRoot, "OfficeSessions");
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const unitSeparator = "\x1f";
const recordSeparator = "\x1e";

function run(program, args, timeout = 60_000) {
  return execFileSync(program, args, {
    encoding: "utf8",
    timeout,
    maxBuffer: 16 * 1024 * 1024,
  }).trim();
}

function runAppleScript(lines, timeout = 60_000) {
  return run(
    "/usr/bin/osascript",
    lines.flatMap((line) => ["-e", line]),
    timeout,
  );
}

function sessionIds() {
  if (!existsSync(sessionsRoot)) return new Set();
  return new Set(
    readdirSync(sessionsRoot, { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => entry.name),
  );
}

function requestForSession(sessionId) {
  const requestPath = join(sessionsRoot, sessionId, "request.json");
  if (!existsSync(requestPath)) return null;
  try {
    const request = JSON.parse(readFileSync(requestPath, "utf8"));
    const sourcePath = join(sessionsRoot, sessionId, "latex-redraw-source.txt");
    if (request.operation === "latexRedraw" && existsSync(sourcePath)) {
      request.source = readFileSync(sourcePath, "utf8");
    }
    return request;
  } catch {
    return null;
  }
}

async function waitForNewRedrawSession(before, expected, timeoutMs = 30_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    for (const sessionId of sessionIds()) {
      if (before.has(sessionId)) continue;
      const request = requestForSession(sessionId);
      if (
        request?.operation === "latexRedraw" &&
        request?.host === "word" &&
        request?.documentImport?.redrawScope === expected.scope &&
        request?.documentImport?.outputKind === expected.output
      ) {
        return { sessionId, request };
      }
    }
    await sleep(50);
  }
  throw new Error(`Word did not create the expected redraw Session: ${JSON.stringify(expected)}`);
}

async function waitForWordUi(timeoutMs = 30_000) {
  const deadline = Date.now() + timeoutMs;
  let lastError = "";
  while (Date.now() < deadline) {
    try {
      const result = runAppleScript([
        'tell application "Microsoft Word" to activate',
        'tell application "System Events"',
        'tell process "Microsoft Word"',
        "set visible to true",
        "set frontmost to true",
        'if (count of menu bars) is 0 then error "Word UI is not ready"',
        'return "READY"',
        "end tell",
        "end tell",
      ], 5_000);
      if (result === "READY") return;
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error);
    }
    await sleep(250);
  }
  throw new Error(`Word UI did not become ready: ${lastError}`);
}

async function waitForCommit(sessionId, timeoutMs = 120_000) {
  const progressPath = join(
    sessionsRoot,
    sessionId,
    "document-import-progress.txt",
  );
  const deadline = Date.now() + timeoutMs;
  let lastProgress = "";
  while (Date.now() < deadline) {
    if (existsSync(progressPath)) {
      lastProgress = readFileSync(progressPath, "utf8");
      if (/stage=complete(?:\r?\n|$)/.test(lastProgress)) return lastProgress;
      if (/stage=error(?:\r?\n|$)/.test(lastProgress)) {
        throw new Error(`Word redraw reported an error: ${lastProgress}`);
      }
    }
    await sleep(100);
  }
  throw new Error(`Word redraw did not complete: ${lastProgress || "no progress"}`);
}

function createDocument(fixture) {
  const name = runAppleScript([
    'tell application "Microsoft Word"',
    "activate",
    "make new document",
    "set documentObject to active document",
    `set content of text object of documentObject to ${JSON.stringify(fixture.text)}`,
    ...fixture.fontRanges.map(
      ({ start, end, size }) =>
        `set font size of font object of (create range documentObject start ${start} end ${end}) to ${size}`,
    ),
    `select (create range documentObject start ${fixture.selectionStart} end ${fixture.selectionEnd})`,
    "return name of documentObject as text",
    "end tell",
  ]);
  return name;
}

function closeDocument(name) {
  try {
    runAppleScript([
      'tell application "Microsoft Word"',
      `if exists document ${JSON.stringify(name)} then close document ${JSON.stringify(name)} saving no`,
      "end tell",
    ]);
  } catch {
    // Best-effort cleanup; a later case creates and targets its own document.
  }
}

async function invokeRedraw(caseInfo, documentName) {
  const macro = caseInfo.scope === "selection"
    ? caseInfo.output === "image"
      ? "VisualTeX_RedrawSelectionToImage"
      : "VisualTeX_RedrawSelectionToOmml"
    : caseInfo.output === "image"
      ? "VisualTeX_RedrawDocumentToImage"
      : "VisualTeX_RedrawDocumentToOmml";
  const before = sessionIds();
  const invocation = [
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(documentName)}`,
    `run VB macro macro name ${JSON.stringify(macro)}`,
    "end tell",
  ];
  runAppleScript(invocation, 60_000);
  return waitForNewRedrawSession(before, caseInfo);
}

function inspectImageDocument(documentName) {
  const raw = runAppleScript([
    'tell application "Microsoft Word"',
    `set documentObject to document ${JSON.stringify(documentName)}`,
    `set reportText to (content of text object of documentObject as text) & (ASCII character 31) & ((count of inline shapes of documentObject) as text)`,
    "repeat with shapeIndex from 1 to count of inline shapes of documentObject",
    "set formulaShape to inline shape shapeIndex of documentObject",
    'set reportText to reportText & (ASCII character 30) & (alternative text of formulaShape as text)',
    "end repeat",
    "return reportText",
    "end tell",
  ]);
  const [summary, ...metadataValues] = raw.split(recordSeparator);
  const separator = summary.lastIndexOf(unitSeparator);
  return {
    text: summary.slice(0, separator),
    count: Number(summary.slice(separator + 1)),
    metadata: metadataValues.map((value) => decodeFormulaMetadata(value)),
  };
}

function inspectParagraphCount(documentName) {
  return Number(runAppleScript([
    'tell application "Microsoft Word"',
    `set documentObject to document ${JSON.stringify(documentName)}`,
    "return count paragraphs of documentObject",
    "end tell",
  ]));
}

function inspectOmmlDocument(documentName) {
  const raw = runAppleScript([
    'tell application "Microsoft Word"',
    `set documentObject to document ${JSON.stringify(documentName)}`,
    "set formulaCount to 0",
    'set formulaRecords to ""',
    "set bookmarkCount to count of bookmarks of documentObject",
    "repeat with bookmarkIndex from 1 to bookmarkCount",
    "set candidateBookmark to bookmark bookmarkIndex of documentObject",
    "set bookmarkName to name of candidateBookmark as text",
    'if bookmarkName starts with "VT_F_" then',
    "set formulaCount to formulaCount + 1",
    "set formulaSize to font size of font object of text object of candidateBookmark",
    'set formulaRecords to formulaRecords & (ASCII character 30) & bookmarkName & (ASCII character 31) & (formulaSize as text)',
    "end if",
    "end repeat",
    "set reportText to (content of text object of documentObject as text) & (ASCII character 31) & (formulaCount as text) & formulaRecords",
    "return reportText",
    "end tell",
  ]);
  const [summary, ...formulaRecords] = raw.split(recordSeparator);
  const separator = summary.lastIndexOf(unitSeparator);
  return {
    text: summary.slice(0, separator),
    count: Number(summary.slice(separator + 1)),
    formulas: formulaRecords.map((record) => {
      const [name, size] = record.split(unitSeparator);
      return { name, size: Number(size) };
    }),
  };
}

function assertContext(caseInfo, report) {
  for (const sentinel of caseInfo.sentinels) {
    if (!report.text.includes(sentinel)) {
      throw new Error(`${caseInfo.name} lost surrounding text ${sentinel}: ${JSON.stringify(report)}`);
    }
  }
  for (const literal of caseInfo.removedLiterals) {
    if (report.text.includes(literal)) {
      throw new Error(`${caseInfo.name} left a LaTeX source literal in Word: ${literal}`);
    }
  }
  if (report.count !== caseInfo.expectedFormulaCount) {
    throw new Error(`${caseInfo.name} inserted ${report.count} formulas instead of ${caseInfo.expectedFormulaCount}`);
  }
}

async function runCase(caseInfo) {
  const documentName = createDocument(caseInfo);
  try {
    const { sessionId, request } = await invokeRedraw(caseInfo, documentName);
    const expectedSource = caseInfo.scope === "selection"
      ? caseInfo.text.slice(caseInfo.selectionStart, caseInfo.selectionEnd)
      : `${caseInfo.text}\r`;
    if (
      request.operation !== "latexRedraw" ||
      request.documentImport?.redrawScope !== caseInfo.scope ||
      request.documentImport?.outputKind !== caseInfo.output ||
      request.source !== expectedSource
    ) {
      throw new Error(`Unexpected redraw request: ${JSON.stringify(request)}`);
    }
    await waitForCommit(sessionId);
    const report = caseInfo.output === "image"
      ? inspectImageDocument(documentName)
      : inspectOmmlDocument(documentName);
    assertContext(caseInfo, report);
    if (Number.isInteger(caseInfo.expectedParagraphCount)) {
      const paragraphCount = inspectParagraphCount(documentName);
      if (paragraphCount !== caseInfo.expectedParagraphCount) {
        throw new Error(
          `${caseInfo.name} changed Word paragraph topology: expected ${caseInfo.expectedParagraphCount}, got ${paragraphCount}`,
        );
      }
    }
    if (caseInfo.output === "image") {
      const fontSizes = report.metadata.map((metadata) => metadata?.fontSizePt);
      if (
        fontSizes.length !== caseInfo.expectedFontSizes.length ||
        fontSizes.some((size, index) => Math.abs(size - caseInfo.expectedFontSizes[index]) > 0.01)
      ) {
        throw new Error(`${caseInfo.name} image metadata did not retain source font sizes: ${JSON.stringify(fontSizes)}`);
      }
    } else {
      const actualSizes = report.formulas.map((formula) => formula.size).sort((a, b) => a - b);
      const expectedSizes = [...caseInfo.expectedFontSizes].sort((a, b) => a - b);
      if (
        actualSizes.length !== expectedSizes.length ||
        actualSizes.some((size, index) => Math.abs(size - expectedSizes[index]) > 0.1)
      ) {
        throw new Error(`${caseInfo.name} OMML did not retain source font sizes: ${JSON.stringify(actualSizes)}`);
      }
    }
    return {
      name: caseInfo.name,
      sessionId,
      scope: caseInfo.scope,
      output: caseInfo.output,
      formulaCount: report.count,
      expectedFontSizes: caseInfo.expectedFontSizes,
      finalText: report.text,
    };
  } finally {
    closeDocument(documentName);
  }
}

await waitForWordUi();
runAppleScript([
  'tell application "Microsoft Word"',
  'run VB macro macro name "VisualTeX_InitializeWordHost"',
  "end tell",
]);

const selectionImageText = "OUTSIDE-A PRE $x+1$ MID $$y^2$$ POST OUTSIDE-B";
const selectionImageStart = selectionImageText.indexOf("PRE");
const selectionImageEnd = selectionImageText.indexOf(" OUTSIDE-B");
const selectionOmmlText = "KEEP-L alpha \\(z_1+z_2\\) omega KEEP-R";
const selectionOmmlStart = selectionOmmlText.indexOf("alpha");
const selectionOmmlEnd = selectionOmmlText.indexOf(" KEEP-R");
const fullImageText = "DOC-I before $a/b$ after\r$$c^2=a^2+b^2$$\rDOC-I tail";
const fullOmmlText = "DOC-O before $p+q$ middle \\[r^2\\] after";
const fullUnicodeImageText = "DOC-U 😀 before $$𝑥+𝑦=𝑧$$ after";
const paragraphTopologyText =
  "DOC-P head\r$$a=1$$\r$$b=2$$\r\r$$c=3$$\rDOC-P tail";

const cases = [
  {
    name: "selection-image-right-to-left",
    scope: "selection",
    output: "image",
    text: selectionImageText,
    selectionStart: selectionImageStart,
    selectionEnd: selectionImageEnd,
    fontRanges: [
      { start: 0, end: selectionImageText.length, size: 11 },
      { start: selectionImageText.indexOf("$x+1$"), end: selectionImageText.indexOf("$x+1$") + 5, size: 10 },
      { start: selectionImageText.indexOf("$$y^2$$"), end: selectionImageText.indexOf("$$y^2$$") + 7, size: 10 },
    ],
    expectedFontSizes: [11, 11],
    expectedFormulaCount: 2,
    sentinels: ["OUTSIDE-A", "PRE", "MID", "POST", "OUTSIDE-B"],
    removedLiterals: ["$x+1$", "$$y^2$$"],
  },
  {
    name: "selection-omml",
    scope: "selection",
    output: "omml",
    text: selectionOmmlText,
    selectionStart: selectionOmmlStart,
    selectionEnd: selectionOmmlEnd,
    fontRanges: [
      { start: 0, end: selectionOmmlText.length, size: 13 },
      { start: selectionOmmlText.indexOf("\\("), end: selectionOmmlText.indexOf("\\)") + 2, size: 10 },
    ],
    expectedFontSizes: [13],
    expectedFormulaCount: 1,
    sentinels: ["KEEP-L", "alpha", "omega", "KEEP-R"],
    removedLiterals: ["\\(z_1+z_2\\)"],
  },
  {
    name: "document-image",
    scope: "document",
    output: "image",
    text: fullImageText,
    selectionStart: 0,
    selectionEnd: fullImageText.length,
    fontRanges: [
      { start: 0, end: fullImageText.length, size: 11 },
      { start: fullImageText.indexOf("$a/b$"), end: fullImageText.indexOf("$a/b$") + 5, size: 10 },
      { start: fullImageText.indexOf("$$c^2=a^2+b^2$$"), end: fullImageText.indexOf("$$c^2=a^2+b^2$$") + "$$c^2=a^2+b^2$$".length, size: 10 },
    ],
    expectedFontSizes: [11, 11],
    expectedFormulaCount: 2,
    sentinels: ["DOC-I", "before", "after", "tail"],
    removedLiterals: ["$a/b$", "$$c^2=a^2+b^2$$"],
  },
  {
    name: "document-image-preserves-paragraph-topology",
    scope: "document",
    output: "image",
    text: paragraphTopologyText,
    selectionStart: 0,
    selectionEnd: paragraphTopologyText.length,
    fontRanges: [],
    expectedFontSizes: [11, 11, 11],
    expectedFormulaCount: 3,
    expectedParagraphCount: paragraphTopologyText.split("\r").length,
    sentinels: ["DOC-P head", "DOC-P tail"],
    removedLiterals: ["$$a=1$$", "$$b=2$$", "$$c=3$$"],
  },
  {
    name: "document-image-supplementary-unicode",
    scope: "document",
    output: "image",
    text: fullUnicodeImageText,
    selectionStart: 0,
    selectionEnd: fullUnicodeImageText.length,
    fontRanges: [],
    expectedFontSizes: [11],
    expectedFormulaCount: 1,
    sentinels: ["DOC-U", "😀", "before", "after"],
    removedLiterals: ["$$𝑥+𝑦=𝑧$$"],
  },
  {
    name: "document-omml",
    scope: "document",
    output: "omml",
    text: fullOmmlText,
    selectionStart: 0,
    selectionEnd: fullOmmlText.length,
    fontRanges: [
      { start: 0, end: fullOmmlText.length, size: 12 },
      { start: fullOmmlText.indexOf("$p+q$"), end: fullOmmlText.indexOf("$p+q$") + 5, size: 10 },
      { start: fullOmmlText.indexOf("\\["), end: fullOmmlText.indexOf("\\]") + 2, size: 10 },
    ],
    expectedFontSizes: [12, 12],
    expectedFormulaCount: 2,
    sentinels: ["DOC-O", "before", "middle", "after"],
    removedLiterals: ["$p+q$", "\\[r^2\\]"],
  },
];

const results = [];
for (const caseInfo of cases) {
  results.push(await runCase(caseInfo));
  console.log(`PASS ${caseInfo.name}`);
}

console.log(JSON.stringify({ status: "PASS", results }, null, 2));
