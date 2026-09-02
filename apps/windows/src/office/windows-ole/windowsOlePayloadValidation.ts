import {
  OFFICE_BRIDGE_PROTOCOL_VERSION,
  type OfficeBridgeEvent,
  type OfficeBridgeResponse,
  type OfficeSelectionResult,
} from "../shared/protocol";
import {
  isVisualTeXFormulaMetadata,
  validFormulaId,
} from "../shared/formulaMetadata";

type JsonRecord = Record<string, unknown>;

function invalid(path: string, expectation: string): never {
  throw new Error(
    `VisualTeX Windows Office Bridge returned invalid data at ${path}; expected ${expectation}.`,
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

function nonNegativeInteger(value: unknown, path: string): number {
  const result = finiteNumber(value, path);
  if (!Number.isInteger(result) || result < 0) {
    invalid(path, "a non-negative integer");
  }
  return result;
}

function validateProtocolVersion(value: unknown, path: string) {
  if (value !== OFFICE_BRIDGE_PROTOCOL_VERSION) {
    invalid(path, `protocol version ${OFFICE_BRIDGE_PROTOCOL_VERSION}`);
  }
}

export function decodeOfficeBridgeResponseEnvelope<TResult>(
  value: unknown,
  requestId: string,
): OfficeBridgeResponse<TResult> {
  const response = record(value, "response");
  validateProtocolVersion(response.protocolVersion, "response.protocolVersion");
  const responseId = nonEmptyString(response.id, "response.id");
  if (responseId !== requestId) {
    invalid("response.id", `the current request id ${JSON.stringify(requestId)}`);
  }
  const ok = booleanValue(response.ok, "response.ok");
  if (!ok && response.error !== undefined) {
    const error = record(response.error, "response.error");
    nonEmptyString(error.code, "response.error.code");
    nonEmptyString(error.message, "response.error.message");
    if (error.retryable !== undefined) {
      booleanValue(error.retryable, "response.error.retryable");
    }
  }
  return response as unknown as OfficeBridgeResponse<TResult>;
}

export function decodeOfficeSelectionResult(value: unknown): OfficeSelectionResult {
  const selection = record(value, "selection");
  if (selection.host !== "word" && selection.host !== "powerpoint") {
    invalid("selection.host", '"word" or "powerpoint"');
  }
  nullableString(selection.documentId, "selection.documentId");
  nullableString(selection.objectId, "selection.objectId");
  booleanValue(selection.readOnly, "selection.readOnly");

  const formulaId = nullableString(selection.formulaId, "selection.formulaId");
  if (formulaId !== null && !validFormulaId(formulaId)) {
    invalid("selection.formulaId", "a VisualTeX formula UUID or null");
  }

  if (selection.metadata !== null) {
    if (!isVisualTeXFormulaMetadata(selection.metadata)) {
      invalid("selection.metadata", "valid VisualTeX formula metadata or null");
    }
    if (formulaId !== null && selection.metadata.formulaId !== formulaId) {
      invalid("selection.metadata.formulaId", "the selected formulaId");
    }
  }
  return selection as unknown as OfficeSelectionResult;
}

export function decodeUpdatedEquationNumberResult(
  value: unknown,
): { updated: number } {
  const result = record(value, "updateEquationNumbers");
  nonNegativeInteger(result.updated, "updateEquationNumbers.updated");
  return result as { updated: number };
}

function decodeOfficeBridgeEvent(
  value: unknown,
  path: string,
): OfficeBridgeEvent & { cursor: number } {
  const event = record(value, path);
  validateProtocolVersion(event.protocolVersion, `${path}.protocolVersion`);
  nonEmptyString(event.event, `${path}.event`);
  nonNegativeInteger(event.cursor, `${path}.cursor`);
  if (!("payload" in event)) {
    invalid(`${path}.payload`, "an event payload (which may be null)");
  }
  return event as unknown as OfficeBridgeEvent & { cursor: number };
}

export function decodeOfficeBridgeEvents(
  value: unknown,
): Array<OfficeBridgeEvent & { cursor: number }> {
  if (!Array.isArray(value)) {
    invalid("events", "an array");
  }
  return value.map((event, index) =>
    decodeOfficeBridgeEvent(event, `events[${index}]`),
  );
}
