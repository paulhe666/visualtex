import type { PendingMultilineOcrFormula } from "./multilineFormula.ts";

export type FormulaLineLike = {
  id?: string;
  latex?: string;
  [key: string]: unknown;
};

function createLineId(): string {
  try {
    if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
      return crypto.randomUUID();
    }
  } catch {
    // Fall through to a deterministic-shape random identifier.
  }
  return `ocr-${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}

export function expandFormulaLinesForPendingOcr(
  lines: FormulaLineLike[],
  pending: PendingMultilineOcrFormula,
): FormulaLineLike[] | null {
  if (pending.lines.length < 2) return null;

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const latex = typeof line.latex === "string" ? line.latex : "";
    const match = latex.indexOf(pending.joinedLatex);
    if (match < 0) continue;

    const prefix = latex.slice(0, match);
    const suffix = latex.slice(match + pending.joinedLatex.length);
    const expanded = pending.lines.map((row, rowIndex) => ({
      ...line,
      id: rowIndex === 0 && line.id ? line.id : createLineId(),
      latex:
        (rowIndex === 0 ? prefix : "") +
        row +
        (rowIndex === pending.lines.length - 1 ? suffix : ""),
    }));
    return [...lines.slice(0, index), ...expanded, ...lines.slice(index + 1)];
  }
  return null;
}
