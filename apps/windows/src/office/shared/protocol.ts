import type { VisualTeXFormulaMetadata } from "./formulaMetadata";

export const OFFICE_BRIDGE_PROTOCOL_VERSION = 1 as const;

export type OfficeBridgeMethod =
  | "health"
  | "office.detect"
  | "powerpoint.getSelection"
  | "powerpoint.insertFormula"
  | "powerpoint.replaceFormula"
  | "powerpoint.markFormula"
  | "powerpoint.deleteFormula"
  | "word.getSelection"
  | "word.insertInlineFormula"
  | "word.insertDisplayFormula"
  | "word.replaceFormula"
  | "word.updateEquationNumbers"
  | "office.openWord"
  | "office.openPowerPoint"
  | "shutdown";

export interface OfficeBridgeRequest<TParams = Record<string, unknown>> {
  protocolVersion: typeof OFFICE_BRIDGE_PROTOCOL_VERSION;
  id: string;
  method: OfficeBridgeMethod;
  params: TParams;
}

export interface OfficeBridgeError {
  code: string;
  message: string;
  retryable?: boolean;
  details?: Record<string, unknown> | null;
}

export interface OfficeBridgeResponse<TResult = unknown> {
  protocolVersion: typeof OFFICE_BRIDGE_PROTOCOL_VERSION;
  id: string;
  ok: boolean;
  result?: TResult;
  error?: OfficeBridgeError;
}

export interface OfficeBridgeEvent<TPayload = unknown> {
  protocolVersion: typeof OFFICE_BRIDGE_PROTOCOL_VERSION;
  event: string;
  payload: TPayload;
}

export interface OfficeSelectionResult {
  host: "word" | "powerpoint";
  documentId: string | null;
  objectId: string | null;
  readOnly: boolean;
  formulaId: string | null;
  metadata: VisualTeXFormulaMetadata | null;
}

export interface FormulaFileParams {
  sessionId: string;
  formulaId: string;
  imagePath: string;
  metadata: VisualTeXFormulaMetadata;
  width: number;
  height: number;
}

export interface ReplaceFormulaParams extends FormulaFileParams {
  sourceDocumentId: string | null;
  sourceObjectId: string | null;
}
