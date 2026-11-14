import type { FormulaHotkeyChord } from "../../../shared/visualtex-app/src/shortcuts/formulaHotkeys";
import { officeHotkeySupportsCode } from "../shortcuts/officeHotkeys";

// macOS' keyboard rules for VisualTeX Anywhere's shortcuts (shortcuts/anywhereHotkeys.ts). The input method gives
// the open formula every key but Command chords (those stay the application's), so those shortcuts use Control or
// Option.

export const anywhereKeyRules = {
  /** keys the app registers system-wide, as for the Office hotkeys */
  globalHotkeySupportsCode: officeHotkeySupportsCode,
  keepsStoredChord: (_chord: FormulaHotkeyChord) => true,
  systemKeyProblem: (_chord: FormulaHotkeyChord, _en: boolean): string | null => null,
  missingModifierText: (en: boolean) =>
    en
      ? "Add Control, Option or Command; Shift can be used in addition."
      : "请加入 Control、Option 或 Command；Shift 可作为附加修饰键。",
  /** macOS registers silent OCR itself (⌘⇧O); it does not collide with the formula's chords */
  isSilentOcrChord: (_chord: FormulaHotkeyChord) => false,
  openFormulaChordProblem: (chord: FormulaHotkeyChord, en: boolean) =>
    chord.metaKey
      ? en
        ? "While a formula is open, Command shortcuts stay the application's. Use Control or Option."
        : "公式模式中 Command 组合键仍归当前应用使用，请改用 Control 或 Option。"
      : null,
};
