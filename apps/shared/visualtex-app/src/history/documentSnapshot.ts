import type { DocumentSnapshot, MathSelectionSnapshot } from "./historyTypes";

export function cloneSelection(selection: MathSelectionSnapshot): MathSelectionSnapshot {
  return {
    ranges: selection.ranges.map(([start, end]) => [start, end]),
    direction: selection.direction,
  };
}

export function cloneSelectionMap(
  selectionByLineId: Record<string, MathSelectionSnapshot>,
): Record<string, MathSelectionSnapshot> {
  return Object.fromEntries(
    Object.entries(selectionByLineId).map(([lineId, selection]) => [
      lineId,
      cloneSelection(selection),
    ]),
  );
}

export function cloneDocumentSnapshot(snapshot: DocumentSnapshot): DocumentSnapshot {
  return {
    ...snapshot,
    lines: snapshot.lines.map(line => ({ ...line })),
    selectionByLineId: cloneSelectionMap(snapshot.selectionByLineId),
    documentSelection: snapshot.documentSelection ? {
      anchor: { ...snapshot.documentSelection.anchor },
      focus: { ...snapshot.documentSelection.focus },
    } : null,
  };
}

export function documentSnapshotsEquivalent(
  left: DocumentSnapshot,
  right: DocumentSnapshot,
): boolean {
  return (
    left.title === right.title &&
    left.activeLineId === right.activeLineId &&
    (left.formulaAlignment ?? "left") ===
      (right.formulaAlignment ?? "left") &&
    left.lines.length === right.lines.length &&
    left.lines.every((line, index) => {
      const rightLine = right.lines[index];
      return (
        line.id === rightLine?.id &&
        line.latex === rightLine?.latex &&
        (line.mode === "inline" ? "inline" : "display") ===
          (rightLine?.mode === "inline" ? "inline" : "display") &&
        (line.displayStyle ?? "default") ===
          (rightLine?.displayStyle ?? "default")
      );
    })
  );
}
