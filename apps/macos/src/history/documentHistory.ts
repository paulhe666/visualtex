import type { FormulaLine } from "../types/formula";
import {
  cloneFormulaLines,
  createFormulaLine,
  normalizeFormulaLines,
  useEditorStore,
} from "../stores/editorStore";
import type {
  DocumentSnapshot,
  HistoryEntry,
  MathSelectionSnapshot,
  ReplayDirection,
} from "./historyTypes";

export interface FocusRestoreTarget {
  lineId: string;
  latex: string;
  selection: MathSelectionSnapshot | null;
}

export function cloneSelectionMap(
  selectionByLineId: Record<string, MathSelectionSnapshot>,
): Record<string, MathSelectionSnapshot> {
  return Object.fromEntries(
    Object.entries(selectionByLineId).map(([lineId, selection]) => [
      lineId,
      {
        ranges: selection.ranges.map(([start, end]) => [start, end]),
        direction: selection.direction,
      },
    ]),
  );
}

export function getEditorDocumentSnapshot(
  selectionByLineId: Record<string, MathSelectionSnapshot> = {},
): DocumentSnapshot {
  const state = useEditorStore.getState();
  return {
    title: state.title,
    lines: cloneFormulaLines(state.lines),
    activeLineId: state.activeLineId,
    selectionByLineId: cloneSelectionMap(selectionByLineId),
  };
}

export function documentSnapshotsEquivalent(
  left: DocumentSnapshot,
  right: DocumentSnapshot,
): boolean {
  return (
    left.title === right.title &&
    left.activeLineId === right.activeLineId &&
    left.lines.length === right.lines.length &&
    left.lines.every(
      (line, index) =>
        line.id === right.lines[index]?.id &&
        line.latex === right.lines[index]?.latex,
    )
  );
}

export function reconcileFormulaLines(
  values: readonly string[],
  currentLines: readonly FormulaLine[],
): FormulaLine[] {
  const normalizedValues = values.length ? values : [""];
  return normalizedValues.map((latex, index) => ({
    id: currentLines[index]?.id ?? crypto.randomUUID(),
    latex,
  }));
}

export function createBlankDocumentSnapshot(title: string): DocumentSnapshot {
  const line = createFormulaLine("");
  return {
    title,
    lines: [line],
    activeLineId: line.id,
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
      store.replaceFormulaLine(entry.lineId, latex);
      store.setActiveLineId(activeLineId ?? entry.lineId);
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

    case "add-line": {
      if (undoing) {
        store.removeFormulaLine(entry.line.id);
        store.setActiveLineId(entry.beforeActiveLineId);
        const targetLineId =
          entry.beforeActiveLineId ?? useEditorStore.getState().lines[0]?.id;
        if (!targetLineId) return null;
        const line = useEditorStore
          .getState()
          .lines.find((item) => item.id === targetLineId);
        return line
          ? {
              lineId: targetLineId,
              latex: line.latex,
              selection: entry.beforeSelection,
            }
          : null;
      }

      store.insertFormulaLine(entry.line, entry.index);
      store.setActiveLineId(entry.afterActiveLineId ?? entry.line.id);
      return {
        lineId: entry.afterActiveLineId ?? entry.line.id,
        latex: entry.line.latex,
        selection: entry.afterSelection,
      };
    }

    case "remove-line": {
      if (undoing) {
        store.insertFormulaLine(entry.line, entry.index);
        store.setActiveLineId(entry.beforeActiveLineId ?? entry.line.id);
        return {
          lineId: entry.beforeActiveLineId ?? entry.line.id,
          latex:
            useEditorStore
              .getState()
              .lines.find(
                (line) => line.id === (entry.beforeActiveLineId ?? entry.line.id),
              )?.latex ?? entry.line.latex,
          selection: entry.beforeSelection,
        };
      }

      store.removeFormulaLine(entry.line.id);
      store.setActiveLineId(entry.afterActiveLineId);
      const targetLineId =
        entry.afterActiveLineId ?? useEditorStore.getState().lines[0]?.id;
      if (!targetLineId) return null;
      const line = useEditorStore
        .getState()
        .lines.find((item) => item.id === targetLineId);
      return line
        ? {
            lineId: targetLineId,
            latex: line.latex,
            selection: entry.afterSelection,
          }
        : null;
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
