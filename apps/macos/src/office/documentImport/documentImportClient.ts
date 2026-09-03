import type { VisualTeXFormulaMetadata } from "../shared/formulaMetadata";
import { invokeTauri } from "../shared/tauriTransport";
import {
  decodeMacosDocumentImportProgress,
  decodeMacosDocumentImportRequest,
  decodeMacosLatexRedrawFontSizes,
} from "./documentImportPayloadValidation";
import type {
  DocumentFormulaDisplayMode,
  DocumentFormulaOutputKind,
  DocumentListKind,
  DocumentParagraphAlignment,
  DocumentParagraphStyle,
} from "./documentImportParser";

export interface MacosFormulaRestoreTarget {
  sourceStart: number;
  sourceEnd: number;
  sourceText: string;
  displayMode: DocumentFormulaDisplayMode;
  fontSizePt: number;
  sourceKind: "omml" | "image";
  mathMl?: string;
  latex?: string;
}

export interface MacosDocumentImportRequest {
  protocolVersion: number;
  sessionId: string;
  host: "word";
  sourceDocumentId: string;
  bookmarkName: string;
  defaultFontSizePt: number;
  operation: "documentImport" | "latexRedraw" | "formulaRestore";
  redrawScope?: "selection" | "document";
  outputKind?: DocumentFormulaOutputKind | "latex";
  sourceKind?: "omml" | "image";
  source?: string;
  restoreTargets?: MacosFormulaRestoreTarget[];
}

export interface DocumentImportParagraphCommitMetadata {
  paragraphId?: string;
  paragraphStyle?: DocumentParagraphStyle;
  paragraphAlignment?: DocumentParagraphAlignment;
  listKind?: DocumentListKind;
  listLevel?: number;
  paragraphStart?: boolean;
  paragraphEnd?: boolean;
}

export interface DocumentImportTextCommitItem
  extends DocumentImportParagraphCommitMetadata {
  kind: "text";
  text: string;
  sourceStart?: number;
  sourceEnd?: number;
  sourceText?: string;
}

export interface DocumentImportFormulaCommitItem
  extends DocumentImportParagraphCommitMetadata {
  kind: "formula";
  formulaId: string;
  latex: string;
  displayMode: DocumentFormulaDisplayMode;
  numbered: boolean;
  fontSizePt: number;
  metadata: VisualTeXFormulaMetadata;
  ommlBase64: string;
  ommlDocxBase64: string;
  svgBase64?: string;
  pngBase64?: string;
  width?: number;
  height?: number;
  baseline?: number;
  inkCenterYRatio?: number;
  sourceStart?: number;
  sourceEnd?: number;
  sourceText?: string;
}

export type DocumentImportCommitItem =
  | DocumentImportTextCommitItem
  | DocumentImportFormulaCommitItem;

export interface CommitMacosDocumentImportInput {
  outputKind: DocumentFormulaOutputKind | "latex";
  items: DocumentImportCommitItem[];
}

export interface MacosDocumentImportProgress {
  current: number;
  total: number;
  stage: "preparing" | "inserting" | "complete" | "error" | string;
}

export interface MacosLatexRedrawFontRangeInput {
  sourceStart: number;
  sourceEnd: number;
  sourceText: string;
  displayMode: DocumentFormulaDisplayMode;
}

export async function getMacosDocumentImportRequest(sessionId: string) {
  return decodeMacosDocumentImportRequest(
    await invokeTauri<unknown>("get_macos_offline_document_import_request", {
      sessionId,
    }),
  );
}

export function reportMacosLatexRedrawStage(
  sessionId: string,
  stage: string,
  elapsedMs: number,
  itemCount: number,
) {
  return invokeTauri<void>("report_macos_offline_latex_redraw_stage", {
    sessionId,
    stage,
    elapsedMs,
    itemCount,
  });
}

export async function resolveMacosLatexRedrawFontSizes(
  sessionId: string,
  ranges: MacosLatexRedrawFontRangeInput[],
) {
  return decodeMacosLatexRedrawFontSizes(
    await invokeTauri<unknown>("resolve_macos_offline_latex_redraw_font_sizes", {
      sessionId,
      input: { ranges },
    }),
  );
}

export function focusMacosDocumentImportTarget(
  operation: "documentImport" | "latexRedraw" | "formulaRestore" = "documentImport",
) {
  return invokeTauri<void>("focus_macos_offline_document_import_target", {
    operation,
  });
}

export function restoreMacosDocumentImportWindow() {
  return invokeTauri<void>("restore_macos_offline_document_import_window", {});
}

export async function getMacosDocumentImportProgress(sessionId: string) {
  return decodeMacosDocumentImportProgress(
    await invokeTauri<unknown>("get_macos_offline_document_import_progress", {
      sessionId,
    }),
  );
}

export function commitMacosDocumentImport(
  sessionId: string,
  input: CommitMacosDocumentImportInput,
) {
  return invokeTauri<void>("commit_macos_offline_document_import", {
    sessionId,
    input,
  });
}

export function cancelMacosDocumentImport(sessionId: string) {
  return invokeTauri<void>("cancel_macos_offline_document_import", {
    sessionId,
  });
}

export function closeMacosDocumentImportWindow() {
  return invokeTauri<void>("close_macos_offline_office_editor_window", {});
}
