import type { FormulaLine } from "../types/formula";

export type EditKind =
  | "insert"
  | "delete-backward"
  | "delete-forward"
  | "composition"
  | "replace";

export type SelectionDirection = "forward" | "backward" | "none";

export interface MathSelectionSnapshot {
  ranges: Array<[number, number]>;
  direction: SelectionDirection;
}

export type FormulaEditSource =
  | "keyboard"
  | "toolbar"
  | "candidate"
  | "ocr"
  | "paste";

export interface PendingFormulaEditTransaction {
  kind: "formula";
  lineId: string;
  beforeLatex: string;
  afterLatex: string;
  beforeSelection: MathSelectionSnapshot;
  afterSelection: MathSelectionSnapshot;
  beforeActiveLineId: string | null;
  afterActiveLineId: string | null;
  editKind: EditKind;
  source: FormulaEditSource;
  startedAt: number;
  updatedAt: number;
}

export interface PendingTitleTransaction {
  kind: "title";
  beforeTitle: string;
  afterTitle: string;
  startedAt: number;
  updatedAt: number;
}

export type PendingEditTransaction =
  | PendingFormulaEditTransaction
  | PendingTitleTransaction;

export interface ReplaceFormulaEntry {
  type: "replace-formula";
  lineId: string;
  beforeLatex: string;
  afterLatex: string;
  beforeSelection: MathSelectionSnapshot;
  afterSelection: MathSelectionSnapshot;
  beforeActiveLineId: string | null;
  afterActiveLineId: string | null;
  timestamp: number;
  source: FormulaEditSource;
}

export interface AddLineEntry {
  type: "add-line";
  line: FormulaLine;
  index: number;
  beforeActiveLineId: string | null;
  afterActiveLineId: string | null;
  beforeSelection: MathSelectionSnapshot | null;
  afterSelection: MathSelectionSnapshot | null;
  timestamp: number;
}

export interface RemoveLineEntry {
  type: "remove-line";
  line: FormulaLine;
  index: number;
  beforeActiveLineId: string | null;
  afterActiveLineId: string | null;
  beforeSelection: MathSelectionSnapshot | null;
  afterSelection: MathSelectionSnapshot | null;
  timestamp: number;
}

export interface DocumentSnapshot {
  title: string;
  lines: FormulaLine[];
  activeLineId: string | null;
  selectionByLineId: Record<string, MathSelectionSnapshot>;
}

export type ReplaceDocumentSource =
  | "source-apply"
  | "history-restore"
  | "new-document"
  | "open-document"
  | "ocr";

export interface ReplaceDocumentEntry {
  type: "replace-document";
  before: DocumentSnapshot;
  after: DocumentSnapshot;
  source: ReplaceDocumentSource;
  timestamp: number;
}

export interface ChangeTitleEntry {
  type: "change-title";
  beforeTitle: string;
  afterTitle: string;
  timestamp: number;
}

export type HistoryEntry =
  | ReplaceFormulaEntry
  | AddLineEntry
  | RemoveLineEntry
  | ReplaceDocumentEntry
  | ChangeTitleEntry;

export interface DocumentCheckpoint {
  id: string;
  createdAt: number;
  operationIndex: number;
  reason: string;
  document: DocumentSnapshot;
}

export interface HistoryStateSnapshot {
  undoStack: readonly HistoryEntry[];
  redoStack: readonly HistoryEntry[];
  pendingTransaction: PendingEditTransaction | null;
  isReplaying: boolean;
  operationIndex: number;
  checkpoints: readonly DocumentCheckpoint[];
  canUndo: boolean;
  canRedo: boolean;
}

export type ReplayDirection = "undo" | "redo";

export interface FormulaEditInput {
  lineId: string;
  beforeLatex: string;
  afterLatex: string;
  beforeSelection: MathSelectionSnapshot;
  afterSelection: MathSelectionSnapshot;
  beforeActiveLineId: string | null;
  afterActiveLineId: string | null;
  editKind: EditKind;
  source: FormulaEditSource;
  timestamp?: number;
}

export interface TitleEditInput {
  beforeTitle: string;
  afterTitle: string;
  timestamp?: number;
}

export interface HistoryAdapter {
  applyEntry: (
    entry: HistoryEntry,
    direction: ReplayDirection,
  ) => void | Promise<void>;
  getDocumentSnapshot: () => DocumentSnapshot;
}
