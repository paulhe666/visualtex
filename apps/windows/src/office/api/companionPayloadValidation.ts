import type { VisualTeXFormulaMetadata } from "../metadata/formulaMetadata";
import { isVisualTeXFormulaMetadata } from "../metadata/formulaMetadata";
import type {
  CompanionHealth,
  NativePowerPointSelection,
  NativePowerPointSlideSnapshot,
  NativeWordInlineBaselineResult,
  PowerPointInteractionEvent,
} from "./companionClient";

type JsonRecord = Record<string, unknown>;

function invalid(path: string, expectation: string): never {
  throw new Error(
    `VisualTeX companion returned invalid data at ${path}; expected ${expectation}.`,
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

function positiveInteger(value: unknown, path: string): number {
  const result = nonNegativeInteger(value, path);
  if (result < 1) invalid(path, "a positive integer");
  return result;
}

function optionalString(value: unknown, path: string) {
  if (value === undefined || value === null) return;
  stringValue(value, path);
}

function optionalFiniteNumber(value: unknown, path: string) {
  if (value === undefined || value === null) return;
  finiteNumber(value, path);
}

function optionalPositiveInteger(value: unknown, path: string) {
  if (value === undefined || value === null) return;
  positiveInteger(value, path);
}

export function decodeCompanionHealth(value: unknown): CompanionHealth {
  const health = record(value, "health");
  booleanValue(health.ok, "health.ok");
  nonEmptyString(health.appVersion, "health.appVersion");
  nonEmptyString(health.officeUiVersion, "health.officeUiVersion");
  positiveInteger(health.protocolVersion, "health.protocolVersion");
  booleanValue(health.ocrAvailable, "health.ocrAvailable");
  return health as unknown as CompanionHealth;
}

export function decodeCachedFormulaMetadata(
  value: unknown,
): VisualTeXFormulaMetadata {
  if (!isVisualTeXFormulaMetadata(value)) {
    invalid("formulaMetadata", "valid VisualTeX formula metadata");
  }
  return value;
}

export function decodeNativePowerPointSelection(
  value: unknown,
  path = "powerPointSelection",
): NativePowerPointSelection {
  const selection = record(value, path);
  nonEmptyString(selection.shapeName, `${path}.shapeName`);
  positiveInteger(selection.slideIndex, `${path}.slideIndex`);
  optionalPositiveInteger(selection.slideId, `${path}.slideId`);
  optionalString(selection.presentationIdentity, `${path}.presentationIdentity`);
  finiteNumber(selection.left, `${path}.left`);
  finiteNumber(selection.top, `${path}.top`);
  nonNegativeNumber(selection.width, `${path}.width`);
  nonNegativeNumber(selection.height, `${path}.height`);
  return selection as unknown as NativePowerPointSelection;
}

export function decodeNativePowerPointSlideSnapshot(
  value: unknown,
): NativePowerPointSlideSnapshot {
  const snapshot = record(value, "powerPointSlideSnapshot");
  stringValue(
    snapshot.presentationIdentity,
    "powerPointSlideSnapshot.presentationIdentity",
  );
  positiveInteger(snapshot.slideIndex, "powerPointSlideSnapshot.slideIndex");
  positiveInteger(snapshot.slideId, "powerPointSlideSnapshot.slideId");
  nonNegativeInteger(snapshot.shapeCount, "powerPointSlideSnapshot.shapeCount");
  if (!Array.isArray(snapshot.shapeNames)) {
    invalid("powerPointSlideSnapshot.shapeNames", "an array of strings");
  }
  snapshot.shapeNames.forEach((entry, index) =>
    stringValue(entry, `powerPointSlideSnapshot.shapeNames[${index}]`),
  );
  return snapshot as unknown as NativePowerPointSlideSnapshot;
}

export function decodeNativeWordInlineBaselineResult(
  value: unknown,
): NativeWordInlineBaselineResult {
  const result = record(value, "wordInlineBaseline");
  finiteNumber(result.appliedPosition, "wordInlineBaseline.appliedPosition");
  nonNegativeNumber(result.width, "wordInlineBaseline.width");
  nonNegativeNumber(result.height, "wordInlineBaseline.height");
  nonNegativeInteger(
    result.matchedShapeIndex,
    "wordInlineBaseline.matchedShapeIndex",
  );
  return result as unknown as NativeWordInlineBaselineResult;
}

function decodePowerPointInteractionEvent(
  value: unknown,
  path: string,
): PowerPointInteractionEvent {
  const event = record(value, path);
  nonNegativeInteger(event.cursor, `${path}.cursor`);
  if (event.host !== "word" && event.host !== "powerpoint") {
    invalid(`${path}.host`, '"word" or "powerpoint"');
  }
  if (event.kind !== "edit-selected" && event.kind !== "edit-requested") {
    invalid(`${path}.kind`, '"edit-selected" or "edit-requested"');
  }
  nonEmptyString(event.formulaId, `${path}.formulaId`);
  nonEmptyString(event.shapeName, `${path}.shapeName`);
  optionalPositiveInteger(event.slideIndex, `${path}.slideIndex`);
  optionalPositiveInteger(event.slideId, `${path}.slideId`);
  optionalString(event.presentationIdentity, `${path}.presentationIdentity`);
  optionalFiniteNumber(event.left, `${path}.left`);
  optionalFiniteNumber(event.top, `${path}.top`);
  if (event.width !== undefined && event.width !== null) {
    nonNegativeNumber(event.width, `${path}.width`);
  }
  if (event.height !== undefined && event.height !== null) {
    nonNegativeNumber(event.height, `${path}.height`);
  }
  nonNegativeInteger(event.createdAt, `${path}.createdAt`);
  return event as unknown as PowerPointInteractionEvent;
}

export function decodePowerPointInteractionEvents(
  value: unknown,
): PowerPointInteractionEvent[] {
  if (!Array.isArray(value)) {
    invalid("powerPointInteractionEvents", "an array");
  }
  return value.map((entry, index) =>
    decodePowerPointInteractionEvent(
      entry,
      `powerPointInteractionEvents[${index}]`,
    ),
  );
}
