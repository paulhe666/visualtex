import {
  formulaHotkeyChordFromEvent,
  formulaHotkeyChordId,
  formulaHotkeyHasModifier,
  isMacKeyboardPlatform,
  protectedFormulaHotkeyAction,
  type FormulaHotkeyChord,
} from "./formulaHotkeys";

/**
 * Formula-row actions of the main editor: split the row at the caret into a
 * new inline or display row, or switch the current row between inline and
 * display (what the row's $ / $$ marker does with the pointer). Each action
 * has a default chord that can be changed or turned off in the hotkey
 * settings.
 */
export type FormulaRowHotkeyActionId =
  | "row-new-inline"
  | "row-new-display"
  | "row-toggle-mode";

export interface FormulaRowHotkeyActionDefinition {
  id: FormulaRowHotkeyActionId;
  labelZh: string;
  labelEn: string;
  descriptionZh: string;
  descriptionEn: string;
}

export interface FormulaRowHotkeyBinding {
  actionId: FormulaRowHotkeyActionId;
  /** null: the user turned this shortcut off. */
  chord: FormulaHotkeyChord | null;
  updatedAt: number;
}

export const FORMULA_ROW_HOTKEY_STORAGE_KEY = "visualtex-formula-row-hotkeys-v1";

export const FORMULA_ROW_HOTKEY_ACTIONS: readonly FormulaRowHotkeyActionDefinition[] = [
  {
    id: "row-new-inline",
    labelZh: "新建行内公式行",
    labelEn: "New inline formula row",
    descriptionZh: "在光标处拆分，新行为行内公式",
    descriptionEn: "Splits the row at the caret; the new row is inline",
  },
  {
    id: "row-new-display",
    labelZh: "新建行间公式行",
    labelEn: "New display formula row",
    descriptionZh: "在光标处拆分，新行为行间公式",
    descriptionEn: "Splits the row at the caret; the new row is display",
  },
  {
    id: "row-toggle-mode",
    labelZh: "切换当前行的行内 / 行间",
    labelEn: "Switch row between inline and display",
    descriptionZh: "与行首的 $ / $$ 标记相同；多行公式固定为行间",
    descriptionEn: "Same as the row's $ / $$ marker; multi-line rows stay display",
  },
] as const;

const actionIds = new Set<FormulaRowHotkeyActionId>(
  FORMULA_ROW_HOTKEY_ACTIONS.map((action) => action.id),
);

export function isFormulaRowHotkeyActionId(
  value: unknown,
): value is FormulaRowHotkeyActionId {
  return (
    typeof value === "string" &&
    actionIds.has(value as FormulaRowHotkeyActionId)
  );
}

export function formulaRowHotkeyAction(actionId: FormulaRowHotkeyActionId) {
  const action = FORMULA_ROW_HOTKEY_ACTIONS.find((item) => item.id === actionId);
  if (!action) throw new Error(`Unknown formula row hotkey action: ${actionId}`);
  return action;
}

export function formulaRowHotkeyActionLabel(
  action: FormulaRowHotkeyActionDefinition,
  language: "cn" | "en",
) {
  return language === "en" ? action.labelEn : action.labelZh;
}

function enterChord(
  modifiers: Partial<
    Pick<FormulaHotkeyChord, "ctrlKey" | "altKey" | "shiftKey" | "metaKey">
  >,
): FormulaHotkeyChord {
  return {
    code: "Enter",
    key: "Enter",
    ctrlKey: false,
    altKey: false,
    shiftKey: false,
    metaKey: false,
    ...modifiers,
  };
}

/**
 * Ctrl+Enter and Option/Alt+Enter keep splitting the row as before. Switching
 * the current row is Command+Shift+Enter on macOS (Command is the Mac
 * counterpart of the Windows Ctrl+Shift+Enter).
 */
export function defaultFormulaRowHotkeyChord(
  actionId: FormulaRowHotkeyActionId,
  mac = isMacKeyboardPlatform(),
): FormulaHotkeyChord {
  switch (actionId) {
    case "row-new-inline":
      return enterChord({ ctrlKey: true });
    case "row-new-display":
      return enterChord({ altKey: true });
    case "row-toggle-mode":
      return mac
        ? enterChord({ metaKey: true, shiftKey: true })
        : enterChord({ ctrlKey: true, shiftKey: true });
  }
}

export function createDefaultFormulaRowHotkeyBindings(): FormulaRowHotkeyBinding[] {
  return FORMULA_ROW_HOTKEY_ACTIONS.map((action) => ({
    actionId: action.id,
    chord: defaultFormulaRowHotkeyChord(action.id),
    updatedAt: 0,
  }));
}

/** The keypad Enter key acts as Enter, as it always did for these actions. */
export function formulaRowHotkeyChordId(chord: FormulaHotkeyChord) {
  return formulaHotkeyChordId(
    chord.code === "NumpadEnter" ? { ...chord, code: "Enter" } : chord,
  );
}

export function isDefaultFormulaRowHotkeyChord(
  actionId: FormulaRowHotkeyActionId,
  chord: FormulaHotkeyChord | null,
) {
  return (
    chord !== null &&
    formulaRowHotkeyChordId(chord) ===
      formulaRowHotkeyChordId(defaultFormulaRowHotkeyChord(actionId))
  );
}

/**
 * Why a chord cannot drive a formula-row action, or null when it can. A
 * modifier is required: Enter alone splits the row and Shift+Enter breaks the
 * line inside the formula.
 */
export function formulaRowHotkeyChordProblem(
  chord: FormulaHotkeyChord,
  language: "cn" | "en" = "cn",
): string | null {
  if (!formulaHotkeyHasModifier(chord)) {
    return language === "en"
      ? "Add Ctrl, Option/Alt or Command; Shift can be used in addition."
      : "请加入 Ctrl、Option/Alt 或 Command；Shift 可作为附加修饰键。";
  }
  const protectedAction = protectedFormulaHotkeyAction(chord, language);
  if (protectedAction) {
    return language === "en"
      ? `This shortcut is reserved for ${protectedAction} and cannot be overridden.`
      : `该快捷键已保留用于“${protectedAction}”，不能覆盖。`;
  }
  return null;
}

export function matchFormulaRowHotkey(
  event: KeyboardEvent,
  bindings: readonly FormulaRowHotkeyBinding[],
): FormulaRowHotkeyActionId | null {
  if (event.isComposing || event.key === "Process") return null;
  const chord = formulaHotkeyChordFromEvent(event);
  if (!chord) return null;
  const chordId = formulaRowHotkeyChordId(chord);
  return (
    bindings.find(
      (binding) =>
        binding.chord !== null &&
        formulaRowHotkeyChordId(binding.chord) === chordId,
    )?.actionId ?? null
  );
}
