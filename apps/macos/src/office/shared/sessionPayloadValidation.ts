import type {
  NativePowerPointCommitSelection,
  OfficeExportResult,
  OfficeFormulaSession,
  PreparedPowerPointCommit,
} from "./sessionClient";
import { isVisualTeXFormulaMetadata } from "./formulaMetadata";

const SESSION_MODES = ["create", "edit"] as const;
const OFFICE_HOSTS = ["word", "powerpoint"] as const;
const DISPLAY_MODES = ["inline", "block"] as const;
const OPERATIONS = ["nativeToImage", "imageToNative"] as const;
const SESSION_STATUSES = [
  "created",
  "editing",
  "committing",
  "completed",
  "cancelled",
  "failed",
] as const;
const LETTER_FONTS = [
  "katex",
  "times",
  "cambria",
  "stix",
  "palatino",
  "helvetica",
] as const;
const CHINESE_FONTS = [
  "system",
  "pingfang",
  "songti",
  "kaiti",
  "heiti",
] as const;

type JsonRecord = Record<string, unknown>;

function invalid(path: string, expectation: string): never {
  throw new Error(
    `VisualTeX Office companion returned invalid data at ${path}; expected ${expectation}.`,
  );
}

function record(value: unknown, path: string): JsonRecord {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    invalid(path, "an object");
  }
  return value as JsonRecord;
}

function stringValue(value: unknown, path: string): string {
  if (typeof value !== "string") invalid(path, "a string");
  return value;
}

function nonEmptyString(value: unknown, path: string): string {
  const result = stringValue(value, path);
  if (!result.trim()) invalid(path, "a non-empty string");
  return result;
}

function nullableString(value: unknown, path: string): string | null {
  if (value === null) return null;
  return stringValue(value, path);
}

function optionalNullableString(value: unknown, path: string) {
  if (value === undefined || value === null) return;
  stringValue(value, path);
}

function booleanValue(value: unknown, path: string): boolean {
  if (typeof value !== "boolean") invalid(path, "a boolean");
  return value;
}

function finiteNumber(value: unknown, path: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    invalid(path, "a finite number");
  }
  return value;
}

function nonNegativeNumber(value: unknown, path: string): number {
  const result = finiteNumber(value, path);
  if (result < 0) invalid(path, "a non-negative finite number");
  return result;
}

function nonNegativeInteger(value: unknown, path: string): number {
  const result = nonNegativeNumber(value, path);
  if (!Number.isInteger(result)) invalid(path, "a non-negative integer");
  return result;
}

function enumValue<T extends string>(
  value: unknown,
  path: string,
  choices: readonly T[],
): T {
  if (typeof value !== "string" || !choices.includes(value as T)) {
    invalid(path, choices.map((choice) => JSON.stringify(choice)).join(" or "));
  }
  return value as T;
}

function optionalEnum<T extends string>(
  value: unknown,
  path: string,
  choices: readonly T[],
) {
  if (value === undefined || value === null) return;
  enumValue(value, path, choices);
}

function validateFormulaLines(value: unknown, path: string) {
  if (!Array.isArray(value)) invalid(path, "an array of formula lines");
  value.forEach((entry, index) => {
    const line = record(entry, `${path}[${index}]`);
    nonEmptyString(line.id, `${path}[${index}].id`);
    stringValue(line.latex, `${path}[${index}].latex`);
  });
}

function validateOriginalMetadata(value: unknown, path: string) {
  if (value === null) return;
  if (!isVisualTeXFormulaMetadata(value)) {
    invalid(path, "valid VisualTeX formula metadata or null");
  }
}

function validateExportResult(value: unknown, path: string): OfficeExportResult {
  const result = record(value, path);
  stringValue(result.svg, `${path}.svg`);
  stringValue(result.svgBase64, `${path}.svgBase64`);
  optionalNullableString(result.pngBase64, `${path}.pngBase64`);
  optionalNullableString(result.ommlBase64, `${path}.ommlBase64`);
  optionalNullableString(result.ommlDocxBase64, `${path}.ommlDocxBase64`);
  nonNegativeNumber(result.width, `${path}.width`);
  nonNegativeNumber(result.height, `${path}.height`);
  for (const key of [
    "baseline",
    "inkTopRatio",
    "inkBottomRatio",
    "inkCenterYRatio",
  ] as const) {
    const entry = result[key];
    if (entry !== undefined && entry !== null) {
      finiteNumber(entry, `${path}.${key}`);
    }
  }
  return result as unknown as OfficeExportResult;
}

export function decodeOfficeFormulaSession(
  value: unknown,
): OfficeFormulaSession {
  const session = record(value, "session");
  nonEmptyString(session.id, "session.id");
  enumValue(session.mode, "session.mode", SESSION_MODES);
  enumValue(session.host, "session.host", OFFICE_HOSTS);
  optionalEnum(session.operation, "session.operation", OPERATIONS);
  nonEmptyString(session.formulaId, "session.formulaId");
  nullableString(session.sourceDocumentId, "session.sourceDocumentId");
  nullableString(session.sourceObjectId, "session.sourceObjectId");
  stringValue(session.title, "session.title");
  validateFormulaLines(session.lines, "session.lines");
  nullableString(session.activeLineId, "session.activeLineId");
  stringValue(session.codeFormat, "session.codeFormat");
  enumValue(session.displayMode, "session.displayMode", DISPLAY_MODES);
  booleanValue(session.numbered, "session.numbered");
  if (session.fontSizePt !== undefined && session.fontSizePt !== null) {
    const fontSizePt = finiteNumber(session.fontSizePt, "session.fontSizePt");
    if (fontSizePt < 5 || fontSizePt > 200) {
      invalid("session.fontSizePt", "a number from 5 through 200");
    }
  }
  optionalEnum(session.formulaLetterFont, "session.formulaLetterFont", LETTER_FONTS);
  optionalEnum(
    session.formulaChineseFont,
    "session.formulaChineseFont",
    CHINESE_FONTS,
  );
  nonNegativeNumber(session.exportWidth, "session.exportWidth");
  nonNegativeNumber(session.exportHeight, "session.exportHeight");
  if (session.exportResult !== null) {
    validateExportResult(session.exportResult, "session.exportResult");
  }
  validateOriginalMetadata(session.originalMetadata, "session.originalMetadata");
  booleanValue(session.dirty, "session.dirty");
  enumValue(session.status, "session.status", SESSION_STATUSES);
  booleanValue(session.autoCommitOnClose, "session.autoCommitOnClose");
  booleanValue(session.explicitCancel, "session.explicitCancel");
  nullableString(session.error, "session.error");
  nonNegativeInteger(session.createdAt, "session.createdAt");
  nonNegativeInteger(session.updatedAt, "session.updatedAt");
  nonNegativeInteger(session.expiresAt, "session.expiresAt");
  return session as unknown as OfficeFormulaSession;
}

function decodePowerPointSelection(
  value: unknown,
  path: string,
): NativePowerPointCommitSelection {
  const selection = record(value, path);
  nonEmptyString(selection.shapeName, `${path}.shapeName`);
  nonNegativeInteger(selection.slideIndex, `${path}.slideIndex`);
  if (selection.slideId !== undefined && selection.slideId !== null) {
    nonNegativeInteger(selection.slideId, `${path}.slideId`);
  }
  optionalNullableString(
    selection.presentationIdentity,
    `${path}.presentationIdentity`,
  );
  finiteNumber(selection.left, `${path}.left`);
  finiteNumber(selection.top, `${path}.top`);
  nonNegativeNumber(selection.width, `${path}.width`);
  nonNegativeNumber(selection.height, `${path}.height`);
  return selection as unknown as NativePowerPointCommitSelection;
}

export function decodePreparedPowerPointCommit(
  value: unknown,
): PreparedPowerPointCommit {
  const prepared = record(value, "preparedPowerPointCommit");
  decodeOfficeFormulaSession(prepared.session);
  decodePowerPointSelection(
    prepared.selection,
    "preparedPowerPointCommit.selection",
  );
  return prepared as unknown as PreparedPowerPointCommit;
}
