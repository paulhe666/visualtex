import {
  formulaHotkeyHasModifier,
  protectedFormulaHotkeyAction,
  type FormulaHotkeyChord,
} from "./formulaHotkeys";

/**
 * The key that confirms a quick escape: it turns only the name right before
 * the caret into its symbol (alpha -> α); Backspace right after undoes it.
 * Any other key (Tab, operators, arrows) never escapes. It can be a plain
 * key such as Space, so it is kept apart from the formula-row hotkeys, which
 * need a modifier.
 */
export const CONFIRM_ESCAPE_KEY_STORAGE_KEY = "visualtex-confirm-escape-key";

export const DEFAULT_CONFIRM_ESCAPE_CHORD: FormulaHotkeyChord = {
  code: "Space",
  key: " ",
  ctrlKey: false,
  altKey: false,
  shiftKey: false,
  metaKey: false,
};

// Keys whose own job inside a formula must stay: typing, Backspace (it
// undoes an escape), Enter, Esc.
const RESERVED_CODES = new Set([
  "Backspace",
  "Delete",
  "Enter",
  "NumpadEnter",
  "Escape",
]);

export function confirmEscapeChordProblem(
  chord: FormulaHotkeyChord,
  language: "cn" | "en" = "cn",
): string | null {
  if (RESERVED_CODES.has(chord.code)) {
    return language === "en"
      ? "Backspace, Delete, Enter and Esc keep their own jobs in a formula."
      : "Backspace、Delete、Enter 和 Esc 在公式中有固定用途，不能用作确认转义键。";
  }
  if (
    !formulaHotkeyHasModifier(chord) &&
    /^(?:Key[A-Z]|Digit[0-9]|Numpad[0-9])$/.test(chord.code)
  ) {
    return language === "en"
      ? "Letters and digits are typed into the formula; pick another key or add Ctrl, Option/Alt or Command."
      : "字母和数字会输入到公式中；请换一个键，或加入 Ctrl、Option/Alt 或 Command。";
  }
  if (formulaHotkeyHasModifier(chord)) {
    const protectedAction = protectedFormulaHotkeyAction(chord, language);
    if (protectedAction) {
      return language === "en"
        ? `This shortcut is reserved for ${protectedAction} and cannot be overridden.`
        : `该快捷键已保留用于“${protectedAction}”，不能覆盖。`;
    }
  }
  return null;
}

export function sameConfirmEscapeChord(
  a: FormulaHotkeyChord | null,
  b: FormulaHotkeyChord | null,
) {
  if (!a || !b) return a === b;
  return (
    a.code === b.code &&
    a.ctrlKey === b.ctrlKey &&
    a.altKey === b.altKey &&
    a.shiftKey === b.shiftKey &&
    a.metaKey === b.metaKey
  );
}

/** What the editor kernel matches a keystroke against. */
export function confirmEscapeKernelKey(chord: FormulaHotkeyChord | null) {
  if (!chord) return null;
  const { code, ctrlKey, altKey, shiftKey, metaKey } = chord;
  return { code, ctrlKey, altKey, shiftKey, metaKey };
}
