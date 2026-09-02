#!/usr/bin/env node

import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

const sessionsRoot = join(
  homedir(),
  "Library",
  "Application Scripts",
  "com.microsoft.Word",
  "VisualTeXRuntime",
  "OfficeSessions",
);
const safePerformancePath = join(
  homedir(),
  "Library",
  "Group Containers",
  "UBF8T346G9.Office",
  "VisualTeX",
  "Scratch",
  "word-safe-edit-performance.json",
);

const limitMs = Number(process.env.VISUALTEX_WORD_EDIT_APPLY_LIMIT_MS ?? "1000");
if (!Number.isFinite(limitMs) || limitMs <= 0) {
  throw new Error("VISUALTEX_WORD_EDIT_APPLY_LIMIT_MS must be a positive number");
}
if (!existsSync(sessionsRoot)) {
  throw new Error(`VisualTeX Office Session root is missing: ${sessionsRoot}`);
}

function readJson(path) {
  try {
    return JSON.parse(readFileSync(path, "utf8"));
  } catch {
    return undefined;
  }
}

function readJsonLines(path) {
  if (!existsSync(path)) return [];
  const rows = [];
  for (const line of readFileSync(path, "utf8").split(/\r?\n/u)) {
    if (!line.trim()) continue;
    try {
      rows.push(JSON.parse(line));
    } catch {
      // Ignore a final partially-written diagnostic line.
    }
  }
  return rows;
}

const ommlOnly = process.argv.includes("--omml-only");
const caseMatchers = [
  {
    key: "image-inline",
    matches: (request) =>
      request.mode === "edit" &&
      request.nativeEquation === false &&
      request.displayMode === "inline" &&
      String(request.sourceDocumentId ?? "").includes("word-safe-edit-image-inline-"),
  },
  {
    key: "omml-inline",
    matches: (request) =>
      request.mode === "edit" &&
      request.nativeEquation === true &&
      request.displayMode === "inline" &&
      String(request.sourceDocumentId ?? "").includes(
        "word-safe-edit-omml-inline-",
      ),
  },
  {
    key: "omml-display",
    matches: (request) =>
      request.mode === "edit" &&
      request.nativeEquation === true &&
      request.displayMode === "block" &&
      String(request.sourceDocumentId ?? "").includes("word-safe-edit-omml-display-"),
  },
  {
    key: "omml-numbered",
    matches: (request) =>
      request.mode === "edit" &&
      request.nativeEquation === true &&
      request.displayMode === "block" &&
      request.numbered === true &&
      String(request.sourceDocumentId ?? "").includes(
        "word-safe-edit-omml-numbered-",
      ),
  },
].filter(({ key }) => !ommlOnly || key.startsWith("omml-"));

const candidates = [];
for (const sessionId of readdirSync(sessionsRoot)) {
  const sessionPath = join(sessionsRoot, sessionId);
  const requestPath = join(sessionPath, "request.json");
  const performancePath = join(sessionPath, "editor-performance.jsonl");
  if (!existsSync(requestPath) || !existsSync(performancePath)) continue;
  const request = readJson(requestPath);
  if (!request) continue;
  const matchedCase = caseMatchers.find(({ matches }) => matches(request));
  if (!matchedCase) continue;
  const events = readJsonLines(performancePath);
  const backendComplete = [...events]
    .reverse()
    .find(({ stage }) => stage === "apply-backend-complete");
  if (!backendComplete || !Number.isFinite(Number(backendComplete.elapsedMs))) continue;
  candidates.push({
    key: matchedCase.key,
    sessionId,
    elapsedMs: Number(backendComplete.elapsedMs),
    modifiedMs: statSync(performancePath).mtimeMs,
    vbaTracePath: join(sessionPath, "word-vba-performance.txt"),
    clickToOfficeCompleteMs: undefined,
  });
}

const safePerformance = readJson(safePerformancePath);
if (safePerformance?.status === "PASS" && Array.isArray(safePerformance.cases)) {
  for (const testCase of safePerformance.cases) {
    const matcher = caseMatchers.find(({ key }) => key === testCase?.label);
    const editApply = testCase?.edit?.apply;
    if (
      !matcher ||
      !Number.isFinite(Number(editApply?.backendElapsedMs)) ||
      !Number.isFinite(Number(editApply?.clickToOfficeCompleteMs))
    ) {
      continue;
    }
    candidates.push({
      key: matcher.key,
      sessionId: String(testCase.edit.sessionId ?? ""),
      elapsedMs: Number(editApply.backendElapsedMs),
      clickToOfficeCompleteMs: Number(editApply.clickToOfficeCompleteMs),
      modifiedMs: statSync(safePerformancePath).mtimeMs,
      vbaTrace: Boolean(testCase.edit.vba?.raw),
    });
  }
}

const latest = new Map();
for (const candidate of candidates.sort((left, right) => right.modifiedMs - left.modifiedMs)) {
  if (!latest.has(candidate.key)) latest.set(candidate.key, candidate);
}

const results = caseMatchers.map(({ key }) => {
  const result = latest.get(key);
  if (!result) {
    throw new Error(`No completed safe Word edit performance Session found for ${key}`);
  }
  return result;
});

const failures = results.filter(
  ({ elapsedMs, clickToOfficeCompleteMs }) =>
    elapsedMs >= limitMs ||
    (Number.isFinite(clickToOfficeCompleteMs) &&
      clickToOfficeCompleteMs >= limitMs),
);
const reportPath = join(
  process.cwd(),
  "test-results",
  "word-edit-performance",
  "latest.json",
);
mkdirSync(join(process.cwd(), "test-results", "word-edit-performance"), {
  recursive: true,
});
writeFileSync(
  reportPath,
  `${JSON.stringify(
    {
      generatedAt: new Date().toISOString(),
      limitMs,
      passed: failures.length === 0,
      results: results.map(({ key, sessionId, elapsedMs, clickToOfficeCompleteMs, vbaTracePath, vbaTrace }) => ({
        key,
        sessionId,
        elapsedMs,
        clickToOfficeCompleteMs,
        vbaTrace: vbaTrace ?? existsSync(vbaTracePath ?? ""),
      })),
    },
    null,
    2,
  )}\n`,
  "utf8",
);
for (const result of results) {
  const clickText = Number.isFinite(result.clickToOfficeCompleteMs)
    ? `; click ${result.clickToOfficeCompleteMs.toFixed(1)} ms`
    : "";
  console.log(
    `${result.key}: backend ${result.elapsedMs.toFixed(1)} ms${clickText} ` +
      `(limit < ${limitMs.toFixed(1)} ms; session ${result.sessionId}; ` +
      `vbaTrace=${result.vbaTrace ?? existsSync(result.vbaTracePath ?? "") ? "yes" : "no"})`,
  );
}
if (failures.length > 0) {
  throw new Error(
    `Word edit Apply performance budget exceeded: ${failures
      .map(({ key, elapsedMs, clickToOfficeCompleteMs }) =>
        `${key}=backend ${elapsedMs.toFixed(1)}ms` +
        (Number.isFinite(clickToOfficeCompleteMs)
          ? `/click ${clickToOfficeCompleteMs.toFixed(1)}ms`
          : ""),
      )
      .join(", ")}`,
  );
}
console.log(`Word edit Apply performance regression PASS (${reportPath})`);
