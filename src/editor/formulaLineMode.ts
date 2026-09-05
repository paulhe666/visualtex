import type { FormulaLineMode, LatexCodeFormat } from "../types/formula";

export interface FormulaLineModeShortcutState {
  shiftKey: boolean;
  altKey: boolean;
}

export function resolveNewFormulaLineMode(
  format: LatexCodeFormat,
  currentMode: FormulaLineMode | undefined,
  shortcut: FormulaLineModeShortcutState,
): FormulaLineMode | undefined {
  if (format !== "mixed-inline-display") return undefined;
  if (shortcut.shiftKey) return "inline";
  if (shortcut.altKey) return "display";
  return currentMode === "inline" ? "inline" : "display";
}
