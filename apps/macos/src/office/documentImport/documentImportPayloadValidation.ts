import type {
  MacosDocumentImportProgress,
  MacosDocumentImportRequest,
} from "./documentImportClient";

type JsonRecord = Record<string, unknown>;

const SESSION_ID_PATTERN =
  /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

function invalid(path: string, expectation: string): never {
  throw new Error(
    `VisualTeX macOS document import returned invalid data at ${path}; expected ${expectation}.`,
  );
}

function record(value: unknown, path: string): JsonRecord {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    invalid(path, "an object");
  }
  return value as JsonRecord;
}

function stringValue(value: unknown, path: string) {
  if (typeof value !== "string") invalid(path, "a string");
  return value;
}

function nonEmptyString(value: unknown, path: string) {
  const result = stringValue(value, path);
  if (!result.trim()) invalid(path, "a non-empty string");
  return result;
}

function finiteNumber(value: unknown, path: string) {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    invalid(path, "a finite number");
  }
  return value;
}

function nonNegativeInteger(value: unknown, path: string) {
  if (
    typeof value !== "number" ||
    !Number.isSafeInteger(value) ||
    value < 0
  ) {
    invalid(path, "a non-negative integer");
  }
  return value;
}

function fontSize(value: unknown, path: string) {
  const result = finiteNumber(value, path);
  if (result < 5 || result > 200) {
    invalid(path, "a font size from 5 through 200 pt");
  }
  return result;
}

function optionalString(value: unknown, path: string) {
  if (value === undefined || value === null) return;
  stringValue(value, path);
}

function displayMode(value: unknown, path: string) {
  if (value !== "inline" && value !== "block") {
    invalid(path, '"inline" or "block"');
  }
  return value;
}

function sourceKind(value: unknown, path: string) {
  if (value !== "omml" && value !== "image") {
    invalid(path, '"omml" or "image"');
  }
  return value;
}

export function decodeMacosDocumentImportRequest(
  value: unknown,
): MacosDocumentImportRequest {
  const request = record(value, "documentImportRequest");
  if (request.protocolVersion !== 1) {
    invalid("documentImportRequest.protocolVersion", "protocol version 1");
  }
  const sessionId = nonEmptyString(request.sessionId, "documentImportRequest.sessionId");
  if (!SESSION_ID_PATTERN.test(sessionId)) {
    invalid("documentImportRequest.sessionId", "a VisualTeX UUID session id");
  }
  if (request.host !== "word") invalid("documentImportRequest.host", '"word"');
  stringValue(request.sourceDocumentId, "documentImportRequest.sourceDocumentId");
  nonEmptyString(request.bookmarkName, "documentImportRequest.bookmarkName");
  fontSize(request.defaultFontSizePt, "documentImportRequest.defaultFontSizePt");
  if (
    request.operation !== "documentImport" &&
    request.operation !== "latexRedraw" &&
    request.operation !== "formulaRestore"
  ) {
    invalid(
      "documentImportRequest.operation",
      '"documentImport", "latexRedraw" or "formulaRestore"',
    );
  }
  if (
    request.redrawScope !== undefined &&
    request.redrawScope !== "selection" &&
    request.redrawScope !== "document"
  ) {
    invalid("documentImportRequest.redrawScope", '"selection" or "document"');
  }
  if (
    request.outputKind !== undefined &&
    request.outputKind !== "omml" &&
    request.outputKind !== "image" &&
    request.outputKind !== "latex"
  ) {
    invalid("documentImportRequest.outputKind", '"omml", "image" or "latex"');
  }
  if (request.sourceKind !== undefined) {
    sourceKind(request.sourceKind, "documentImportRequest.sourceKind");
  }
  optionalString(request.source, "documentImportRequest.source");
  if (request.restoreTargets !== undefined) {
    if (!Array.isArray(request.restoreTargets)) {
      invalid("documentImportRequest.restoreTargets", "an array");
    }
    request.restoreTargets.forEach((entry, index) => {
      const path = `documentImportRequest.restoreTargets[${index}]`;
      const target = record(entry, path);
      const start = nonNegativeInteger(target.sourceStart, `${path}.sourceStart`);
      const end = nonNegativeInteger(target.sourceEnd, `${path}.sourceEnd`);
      if (end < start) invalid(`${path}.sourceEnd`, "an offset not before sourceStart");
      stringValue(target.sourceText, `${path}.sourceText`);
      displayMode(target.displayMode, `${path}.displayMode`);
      fontSize(target.fontSizePt, `${path}.fontSizePt`);
      sourceKind(target.sourceKind, `${path}.sourceKind`);
      optionalString(target.mathMl, `${path}.mathMl`);
      optionalString(target.latex, `${path}.latex`);
    });
  }
  return request as unknown as MacosDocumentImportRequest;
}

export function decodeMacosLatexRedrawFontSizes(value: unknown): number[] {
  if (!Array.isArray(value)) {
    invalid("latexRedrawFontSizes", "an array of font sizes");
  }
  return value.map((entry, index) =>
    fontSize(entry, `latexRedrawFontSizes[${index}]`),
  );
}

export function decodeMacosDocumentImportProgress(
  value: unknown,
): MacosDocumentImportProgress {
  const progress = record(value, "documentImportProgress");
  const current = nonNegativeInteger(progress.current, "documentImportProgress.current");
  const total = nonNegativeInteger(progress.total, "documentImportProgress.total");
  if (current > total && total !== 0) {
    invalid("documentImportProgress.current", "a value not greater than total");
  }
  stringValue(progress.stage, "documentImportProgress.stage");
  return progress as unknown as MacosDocumentImportProgress;
}
