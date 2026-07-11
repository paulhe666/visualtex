import { useSyncExternalStore } from "react";
import { loadRecentCheckpoints, persistCheckpoint } from "./checkpointStore";
import type {
  ChangeTitleEntry,
  DocumentCheckpoint,
  FormulaEditInput,
  HistoryAdapter,
  HistoryEntry,
  HistoryStateSnapshot,
  MathSelectionSnapshot,
  PendingEditTransaction,
  PendingFormulaEditTransaction,
  ReplaceFormulaEntry,
  TitleEditInput,
} from "./historyTypes";

export const EDIT_GROUP_TIMEOUT_MS = 500;
export const MAX_HISTORY_ENTRIES = 300;
export const CHECKPOINT_INTERVAL = 30;
export const MAX_MEMORY_CHECKPOINTS = 10;
const HISTORY_TRIM_COUNT = 50;

type Listener = () => void;

function cloneSelection(
  selection: MathSelectionSnapshot,
): MathSelectionSnapshot {
  return {
    ranges: selection.ranges.map(([start, end]) => [start, end]),
    direction: selection.direction,
  };
}

export function clampSelection(
  selection: MathSelectionSnapshot,
  lastOffset: number,
): MathSelectionSnapshot {
  return {
    ranges: selection.ranges.map(([start, end]) => [
      Math.max(0, Math.min(start, lastOffset)),
      Math.max(0, Math.min(end, lastOffset)),
    ]),
    direction: selection.direction,
  };
}

function cloneHistoryEntry(entry: HistoryEntry): HistoryEntry {
  switch (entry.type) {
    case "replace-formula":
      return {
        ...entry,
        beforeSelection: cloneSelection(entry.beforeSelection),
        afterSelection: cloneSelection(entry.afterSelection),
      };
    case "add-line":
    case "remove-line":
      return {
        ...entry,
        line: { ...entry.line },
        beforeSelection: entry.beforeSelection
          ? cloneSelection(entry.beforeSelection)
          : null,
        afterSelection: entry.afterSelection
          ? cloneSelection(entry.afterSelection)
          : null,
      };
    case "replace-document":
      return {
        ...entry,
        before: {
          ...entry.before,
          lines: entry.before.lines.map((line) => ({ ...line })),
          selectionByLineId: Object.fromEntries(
            Object.entries(entry.before.selectionByLineId).map(
              ([lineId, selection]) => [lineId, cloneSelection(selection)],
            ),
          ),
        },
        after: {
          ...entry.after,
          lines: entry.after.lines.map((line) => ({ ...line })),
          selectionByLineId: Object.fromEntries(
            Object.entries(entry.after.selectionByLineId).map(
              ([lineId, selection]) => [lineId, cloneSelection(selection)],
            ),
          ),
        },
      };
    case "change-title":
      return { ...entry };
  }
}

interface TextDifference {
  start: number;
  beforeEnd: number;
  afterEnd: number;
}

function textDifference(before: string, after: string): TextDifference {
  let start = 0;
  const sharedLength = Math.min(before.length, after.length);
  while (start < sharedLength && before[start] === after[start]) start += 1;

  let beforeEnd = before.length;
  let afterEnd = after.length;
  while (
    beforeEnd > start &&
    afterEnd > start &&
    before[beforeEnd - 1] === after[afterEnd - 1]
  ) {
    beforeEnd -= 1;
    afterEnd -= 1;
  }
  return { start, beforeEnd, afterEnd };
}

function normalizeFormulaEditKind(
  input: FormulaEditInput,
): FormulaEditInput["editKind"] {
  if (input.editKind === "composition") return "composition";
  const difference = textDifference(input.beforeLatex, input.afterLatex);
  const removedLength = difference.beforeEnd - difference.start;
  const insertedLength = difference.afterEnd - difference.start;
  if (removedLength === 0 && insertedLength > 0) return "insert";
  if (insertedLength === 0 && removedLength > 0) {
    return input.editKind === "delete-forward"
      ? "delete-forward"
      : "delete-backward";
  }
  return "replace";
}

function formulaEditsAreAdjacent(
  pending: PendingFormulaEditTransaction,
  next: FormulaEditInput,
): boolean {
  const pendingDifference = textDifference(
    pending.beforeLatex,
    pending.afterLatex,
  );
  const nextDifference = textDifference(next.beforeLatex, next.afterLatex);
  switch (pending.editKind) {
    case "insert":
      return nextDifference.start === pendingDifference.afterEnd;
    case "delete-backward":
      return nextDifference.beforeEnd === pendingDifference.start;
    case "delete-forward":
      return nextDifference.start === pendingDifference.start;
    default:
      return false;
  }
}

function isMergeableFormulaKind(
  transaction: PendingFormulaEditTransaction,
): boolean {
  return (
    transaction.editKind === "insert" ||
    transaction.editKind === "delete-backward" ||
    transaction.editKind === "delete-forward"
  );
}

function canMergeFormulaEdit(
  pending: PendingFormulaEditTransaction,
  next: FormulaEditInput,
  timestamp: number,
): boolean {
  return (
    isMergeableFormulaKind(pending) &&
    pending.lineId === next.lineId &&
    pending.editKind === next.editKind &&
    pending.source === next.source &&
    timestamp - pending.updatedAt <= EDIT_GROUP_TIMEOUT_MS &&
    pending.afterLatex === next.beforeLatex &&
    pending.afterActiveLineId === next.beforeActiveLineId &&
    formulaEditsAreAdjacent(pending, next)
  );
}

export class HistoryManager {
  private undoStack: HistoryEntry[] = [];
  private redoStack: HistoryEntry[] = [];
  private pendingTransaction: PendingEditTransaction | null = null;
  private isReplaying = false;
  private operationIndex = 0;
  private checkpoints: DocumentCheckpoint[] = [];
  private adapter: HistoryAdapter | null = null;
  private listeners = new Set<Listener>();
  private commitTimer: ReturnType<typeof setTimeout> | null = null;
  private snapshot: HistoryStateSnapshot = this.createStateSnapshot();

  constructor() {
    void this.loadPersistedCheckpoints();
  }

  configure(adapter: HistoryAdapter | null) {
    this.adapter = adapter;
  }

  subscribe = (listener: Listener) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  getSnapshot = () => this.snapshot;

  getState(): HistoryStateSnapshot {
    return this.snapshot;
  }

  recordFormulaEdit(input: FormulaEditInput) {
    if (this.isReplaying || input.beforeLatex === input.afterLatex) return;

    this.redoStack = [];
    const timestamp = input.timestamp ?? Date.now();
    const normalizedInput: FormulaEditInput = {
      ...input,
      editKind: normalizeFormulaEditKind(input),
    };
    const pending = this.pendingTransaction;
    if (
      pending?.kind === "formula" &&
      canMergeFormulaEdit(pending, normalizedInput, timestamp)
    ) {
      this.pendingTransaction = {
        ...pending,
        afterLatex: normalizedInput.afterLatex,
        afterSelection: cloneSelection(normalizedInput.afterSelection),
        afterActiveLineId: normalizedInput.afterActiveLineId,
        updatedAt: timestamp,
      };
    } else {
      this.commitPendingTransaction();
      this.pendingTransaction = {
        kind: "formula",
        lineId: normalizedInput.lineId,
        beforeLatex: normalizedInput.beforeLatex,
        afterLatex: normalizedInput.afterLatex,
        beforeSelection: cloneSelection(normalizedInput.beforeSelection),
        afterSelection: cloneSelection(normalizedInput.afterSelection),
        beforeActiveLineId: normalizedInput.beforeActiveLineId,
        afterActiveLineId: normalizedInput.afterActiveLineId,
        editKind: normalizedInput.editKind,
        source: normalizedInput.source,
        startedAt: timestamp,
        updatedAt: timestamp,
      };
    }

    if (
      normalizedInput.editKind === "composition" ||
      normalizedInput.editKind === "replace" ||
      normalizedInput.source === "paste"
    ) {
      this.commitPendingTransaction();
      return;
    }

    this.schedulePendingCommit();
    this.emit();
  }

  recordTitleEdit(input: TitleEditInput) {
    if (this.isReplaying || input.beforeTitle === input.afterTitle) return;

    this.redoStack = [];
    const timestamp = input.timestamp ?? Date.now();
    const pending = this.pendingTransaction;
    if (
      pending?.kind === "title" &&
      timestamp - pending.updatedAt <= EDIT_GROUP_TIMEOUT_MS &&
      pending.afterTitle === input.beforeTitle
    ) {
      this.pendingTransaction = {
        ...pending,
        afterTitle: input.afterTitle,
        updatedAt: timestamp,
      };
    } else {
      this.commitPendingTransaction();
      this.pendingTransaction = {
        kind: "title",
        beforeTitle: input.beforeTitle,
        afterTitle: input.afterTitle,
        startedAt: timestamp,
        updatedAt: timestamp,
      };
    }
    this.schedulePendingCommit();
    this.emit();
  }

  commitPendingTransaction() {
    this.clearCommitTimer();
    const pending = this.pendingTransaction;
    if (!pending) return;
    this.pendingTransaction = null;

    if (pending.kind === "formula") {
      if (pending.beforeLatex !== pending.afterLatex) {
        const entry: ReplaceFormulaEntry = {
          type: "replace-formula",
          lineId: pending.lineId,
          beforeLatex: pending.beforeLatex,
          afterLatex: pending.afterLatex,
          beforeSelection: cloneSelection(pending.beforeSelection),
          afterSelection: cloneSelection(pending.afterSelection),
          beforeActiveLineId: pending.beforeActiveLineId,
          afterActiveLineId: pending.afterActiveLineId,
          timestamp: pending.updatedAt,
          source: pending.source,
        };
        this.pushInternal(entry, true);
        return;
      }
    } else if (pending.beforeTitle !== pending.afterTitle) {
      const entry: ChangeTitleEntry = {
        type: "change-title",
        beforeTitle: pending.beforeTitle,
        afterTitle: pending.afterTitle,
        timestamp: pending.updatedAt,
      };
      this.pushInternal(entry, true);
      return;
    }

    this.emit();
  }

  push(entry: HistoryEntry) {
    if (this.isReplaying) return;
    this.commitPendingTransaction();
    this.pushInternal(entry, true);
  }

  async undo(): Promise<boolean> {
    this.commitPendingTransaction();
    if (this.isReplaying || !this.adapter) return false;

    const entry = this.undoStack.pop();
    if (!entry) {
      this.emit();
      return false;
    }

    this.isReplaying = true;
    this.emit();
    try {
      await this.adapter.applyEntry(entry, "undo");
      this.redoStack.push(entry);
      return true;
    } catch (error) {
      this.undoStack.push(entry);
      throw error;
    } finally {
      this.isReplaying = false;
      this.emit();
    }
  }

  async redo(): Promise<boolean> {
    this.commitPendingTransaction();
    if (this.isReplaying || !this.adapter) return false;

    const entry = this.redoStack.pop();
    if (!entry) {
      this.emit();
      return false;
    }

    this.isReplaying = true;
    this.emit();
    try {
      await this.adapter.applyEntry(entry, "redo");
      this.undoStack.push(entry);
      return true;
    } catch (error) {
      this.redoStack.push(entry);
      throw error;
    } finally {
      this.isReplaying = false;
      this.emit();
    }
  }

  clear() {
    this.clearCommitTimer();
    this.undoStack = [];
    this.redoStack = [];
    this.pendingTransaction = null;
    this.isReplaying = false;
    this.operationIndex = 0;
    this.emit();
  }

  async createCheckpoint(reason: string): Promise<DocumentCheckpoint | null> {
    if (!this.adapter) return null;
    const previous = this.checkpoints[0];
    if (
      previous &&
      previous.operationIndex === this.operationIndex &&
      reason !== "before-unload"
    ) {
      return previous;
    }

    const checkpoint: DocumentCheckpoint = {
      id: crypto.randomUUID(),
      createdAt: Date.now(),
      operationIndex: this.operationIndex,
      reason,
      document: this.adapter.getDocumentSnapshot(),
    };
    this.checkpoints = [checkpoint, ...this.checkpoints].slice(
      0,
      MAX_MEMORY_CHECKPOINTS,
    );
    this.emit();
    await persistCheckpoint(checkpoint).catch(() => undefined);
    return checkpoint;
  }

  private pushInternal(entry: HistoryEntry, clearRedo: boolean) {
    this.undoStack.push(cloneHistoryEntry(entry));
    if (clearRedo) this.redoStack = [];
    this.operationIndex += 1;

    const isLargeOperation = entry.type === "replace-document";
    if (
      isLargeOperation ||
      this.operationIndex % CHECKPOINT_INTERVAL === 0 ||
      this.undoStack.length > MAX_HISTORY_ENTRIES
    ) {
      void this.createCheckpoint(
        isLargeOperation ? entry.source : "history-capacity",
      );
    }

    if (this.undoStack.length > MAX_HISTORY_ENTRIES) {
      this.undoStack.splice(0, HISTORY_TRIM_COUNT);
    }
    this.emit();
  }

  private schedulePendingCommit() {
    this.clearCommitTimer();
    this.commitTimer = setTimeout(
      () => this.commitPendingTransaction(),
      EDIT_GROUP_TIMEOUT_MS,
    );
  }

  private clearCommitTimer() {
    if (this.commitTimer !== null) {
      clearTimeout(this.commitTimer);
      this.commitTimer = null;
    }
  }

  private async loadPersistedCheckpoints() {
    const persisted = await loadRecentCheckpoints().catch(() => []);
    if (!persisted.length) return;
    const merged = new Map<string, DocumentCheckpoint>();
    [...this.checkpoints, ...persisted].forEach((checkpoint) => {
      const previous = merged.get(checkpoint.id);
      if (!previous || checkpoint.createdAt > previous.createdAt) {
        merged.set(checkpoint.id, checkpoint);
      }
    });
    this.checkpoints = [...merged.values()]
      .sort((left, right) => right.createdAt - left.createdAt)
      .slice(0, MAX_MEMORY_CHECKPOINTS);
    this.emit();
  }

  private createStateSnapshot(): HistoryStateSnapshot {
    return {
      undoStack: this.undoStack,
      redoStack: this.redoStack,
      pendingTransaction: this.pendingTransaction,
      isReplaying: this.isReplaying,
      operationIndex: this.operationIndex,
      checkpoints: this.checkpoints,
      canUndo: Boolean(this.pendingTransaction || this.undoStack.length),
      canRedo: this.redoStack.length > 0,
    };
  }

  private emit() {
    this.snapshot = this.createStateSnapshot();
    this.listeners.forEach((listener) => listener());
  }
}

export const historyManager = new HistoryManager();

export function useHistorySnapshot(): HistoryStateSnapshot {
  return useSyncExternalStore(
    historyManager.subscribe,
    historyManager.getSnapshot,
    historyManager.getSnapshot,
  );
}
