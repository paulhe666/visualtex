import type { FormulaHotkeyChord } from "../../../shared/visualtex-app/src/shortcuts/formulaHotkeys";

// VisualTeX Anywhere needs the desktop app (APP_FEATURES.anywhere is false); the rules exist for the shared
// shortcut code's types only.
export const anywhereKeyRules = {
  globalHotkeySupportsCode: (_code: string) => false,
  keepsStoredChord: (_chord: FormulaHotkeyChord) => true,
  systemKeyProblem: (_chord: FormulaHotkeyChord, _en: boolean): string | null => null,
  missingModifierText: (en: boolean) =>
    en ? "Add Ctrl, Alt or Command; Shift can be used in addition." : "请加入 Ctrl、Alt 或 Command；Shift 可作为附加修饰键。",
  isSilentOcrChord: (_chord: FormulaHotkeyChord) => false,
  openFormulaChordProblem: (_chord: FormulaHotkeyChord, _en: boolean): string | null => null,
};
