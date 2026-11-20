import type { FormulaHotkeyChord } from "../../../shared/visualtex-app/src/shortcuts/formulaHotkeys";

// Windows' keyboard rules for VisualTeX Anywhere's shortcuts (shortcuts/anywhereHotkeys.ts). The input method
// gives the open formula every key but plain Ctrl chords (those stay the application's, except Ctrl+Z / Ctrl+Y /
// Ctrl+V / Ctrl+Enter, which act on the formula) and the Windows key, so those shortcuts use Alt.

/** Keys Windows can register as a system-wide shortcut (RegisterHotKey,
 *  src-tauri/src/anywhere/hotkey.rs maps them to virtual keys). */
const GLOBAL_HOTKEY_CODES = new Set([
  "KeyA", "KeyB", "KeyC", "KeyD", "KeyE", "KeyF", "KeyG", "KeyH",
  "KeyI", "KeyJ", "KeyK", "KeyL", "KeyM", "KeyN", "KeyO", "KeyP",
  "KeyQ", "KeyR", "KeyS", "KeyT", "KeyU", "KeyV", "KeyW", "KeyX",
  "KeyY", "KeyZ",
  "Digit0", "Digit1", "Digit2", "Digit3", "Digit4", "Digit5", "Digit6",
  "Digit7", "Digit8", "Digit9",
  "Backquote", "Backslash", "BracketLeft", "BracketRight", "Comma", "Equal",
  "Minus", "Period", "Quote", "Semicolon", "Slash", "Space",
  "Home", "End", "PageUp", "PageDown", "Insert", "Delete",
  "ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown",
  "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10",
  "F11", "F12", "F13", "F14", "F15", "F16", "F17", "F18", "F19",
  "F20", "F21", "F22", "F23", "F24",
]);

export const anywhereKeyRules = {
  globalHotkeySupportsCode: (code: string) => GLOBAL_HOTKEY_CODES.has(code),
  /** a chord with the Windows key (a configuration from a Mac) is none here */
  keepsStoredChord: (chord: FormulaHotkeyChord) => !chord.metaKey,
  systemKeyProblem: (chord: FormulaHotkeyChord, en: boolean) =>
    chord.metaKey
      ? en
        ? "The Windows key belongs to Windows. Use Ctrl and Alt."
        : "Windows 键归系统使用，请改用 Ctrl 和 Alt。"
      : null,
  missingModifierText: (en: boolean) =>
    en ? "Add Ctrl or Alt; Shift can be used in addition." : "请加入 Ctrl 或 Alt；Shift 可作为附加修饰键。",
  /** The screenshot OCR shortcut Windows reserves while silent OCR is on
   *  (src-tauri/src/windows_silent_ocr_hotkey.rs: Ctrl+Alt+O, or
   *  Ctrl+Alt+Shift+O when that is taken): a system-wide shortcut wins over
   *  the open formula's keys. */
  isSilentOcrChord: (chord: FormulaHotkeyChord) =>
    chord.code === "KeyO" && chord.ctrlKey && chord.altKey && !chord.metaKey,
  openFormulaChordProblem: (chord: FormulaHotkeyChord, en: boolean) =>
    chord.altKey
      ? null
      : en
        ? "While a formula is open, Ctrl shortcuts stay the application's. Add Alt (Ctrl+Alt or Alt)."
        : "公式模式中 Ctrl 组合键仍归当前应用使用，请加入 Alt（Ctrl+Alt 或 Alt）。",
};
