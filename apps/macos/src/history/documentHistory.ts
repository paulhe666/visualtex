import type { HistoryManager } from "./HistoryManager";
import { cloneSelectionMap } from "./documentSnapshot";
import type { FormulaLine } from "../types/formula";
import { createUuid } from "../runtime/browserCompatibility";
import {
  cloneFormulaLines,
  createFormulaLine,
  useEditorStore,
} from "../stores/editorStore";
import type {
  DocumentSnapshot,
  FormulaEditInput,
  HistoryEntry,
  MathSelectionSnapshot,
  ReplayDirection,
} from "./historyTypes";

export interface FocusRestoreTarget {
  lineId: string;
  latex: string;
  selection: MathSelectionSnapshot | null;
  document?: DocumentSnapshot;
}

export function getEditorDocumentSnapshot(
  selectionByLineId: Record<string, MathSelectionSnapshot> = {},
): DocumentSnapshot {
  const state = useEditorStore.getState();
  return {
    title: state.title,
    lines: cloneFormulaLines(state.lines),
    activeLineId: state.activeLineId,
    formulaAlignment: state.formulaAlignment,
    selectionByLineId: cloneSelectionMap(selectionByLineId),
  };
}

// All field events and explicit commands commit the same normalized document
// value. Selection metadata belongs to history, not to the persisted store.
export function commitFormulaEdit(
  history: HistoryManager,
  edit: FormulaEditInput,
  mode: "grouped" | "discrete" | "silent" = "grouped",
) {
  const state = useEditorStore.getState();
  const line = state.lines.find(item => item.id === edit.lineId);
  if (!line) return;
  const beforeActiveLineId = state.activeLineId;
  state.replaceFormulaLine(edit.lineId, edit.afterLatex, edit.lineId);
  const afterLatex = useEditorStore.getState().lines.find(item => item.id === edit.lineId)!.latex;
  if (mode === "silent" || history.getState().isReplaying || line.latex === afterLatex) return;
  const committed = { ...edit, beforeLatex: line.latex, afterLatex,
    beforeActiveLineId, afterActiveLineId: edit.lineId };
  if (mode === "grouped") history.recordFormulaEdit(committed);
  else history.push({ ...committed, type: "replace-formula", timestamp: edit.timestamp ?? Date.now() });
}

export function reconcileFormulaLines(
  values: readonly string[],
  currentLines: readonly FormulaLine[],
  modes?: readonly FormulaLine["mode"][],
  displayStyles?: readonly FormulaLine["displayStyle"][],
): FormulaLine[] {
  const normalizedValues = values.length ? values : [""];
  return normalizedValues.map((latex, index) => ({
    id: currentLines[index]?.id ?? createUuid(),
    latex,
    mode:
      modes?.[index] === "inline"
        ? "inline"
        : modes?.[index] === "display"
          ? "display"
          : currentLines[index]?.mode ?? "display",
    displayStyle:
      displayStyles?.[index] ??
      currentLines[index]?.displayStyle ??
      "default",
  }));
}

export function createBlankDocumentSnapshot(title: string): DocumentSnapshot {
  const line = createFormulaLine("");
  return {
    title,
    lines: [line],
    activeLineId: line.id,
    formulaAlignment: "left",
    selectionByLineId: {
      [line.id]: { ranges: [[0, 0]], direction: "none" },
    },
  };
}

function targetFromSnapshot(snapshot: DocumentSnapshot): FocusRestoreTarget | null {
  const lineId = snapshot.activeLineId;
  if (!lineId) return null;
  const line = snapshot.lines.find((item) => item.id === lineId);
  if (!line) return null;
  return {
    lineId,
    latex: line.latex,
    selection: snapshot.selectionByLineId[lineId] ?? null,
    document: snapshot,
  };
}

export function applyHistoryEntryToEditor(
  entry: HistoryEntry,
  direction: ReplayDirection,
): FocusRestoreTarget | null {
  const undoing = direction === "undo";
  const store = useEditorStore.getState();

  switch (entry.type) {
    case "replace-formula": {
      const latex = undoing ? entry.beforeLatex : entry.afterLatex;
      const activeLineId = undoing
        ? entry.beforeActiveLineId
        : entry.afterActiveLineId;
      const selection = undoing
        ? entry.beforeSelection
        : entry.afterSelection;
      store.replaceFormulaLine(entry.lineId, latex, activeLineId ?? entry.lineId);
      return {
        lineId: activeLineId ?? entry.lineId,
        latex:
          useEditorStore
            .getState()
            .lines.find((line) => line.id === (activeLineId ?? entry.lineId))
            ?.latex ?? latex,
        selection,
      };
    }

    case "add-line":
    case "remove-line": {
      const inserting = (entry.type === "add-line") !== undoing;
      const lines = store.lines.filter(line => line.id !== entry.line.id);
      if (inserting) lines.splice(Math.max(0, Math.min(entry.index, lines.length)), 0, entry.line);
      const activeLineId = (undoing ? entry.beforeActiveLineId : entry.afterActiveLineId)
        ?? (inserting ? entry.line.id : null);
      store.replaceDocumentState({ ...getEditorDocumentSnapshot(), lines, activeLineId });
      const next = useEditorStore.getState();
      const line = next.lines.find(item => item.id === next.activeLineId);
      return line ? {
        lineId: line.id,
        latex: line.latex,
        selection: undoing ? entry.beforeSelection : entry.afterSelection,
      } : null;
    }

    case "replace-document": {
      const snapshot = undoing ? entry.before : entry.after;
      store.replaceDocumentState(snapshot);
      return targetFromSnapshot(snapshot);
    }

    case "change-title":
      store.setTitle(undoing ? entry.beforeTitle : entry.afterTitle);
      return null;
  }
}
