import {
  formatFormulaHotkeyChord,
  formulaHotkeyChordId,
  formulaHotkeyHasModifier,
  protectedFormulaHotkeyAction,
  type FormulaHotkeyChord,
} from "./formulaHotkeys";
import { officeHotkeySupportsCode } from "./officeHotkeys";

// VisualTeX Anywhere's shortcuts (Settings > Keyboard shortcuts): the
// system-wide ones that open and close a formula in any application (as an
// inline or as a display formula), and
// the ones that act while a formula is open (output format, the row's inline
// or display form, screenshot OCR).
// The input method gives the open formula every key but Command chords
// (those stay the application's), so those shortcuts use Control or Option.

export type AnywhereHotkeyAction =
  | "anywhere-toggle"
  | "anywhere-open-display"
  | "anywhere-next-output"
  | "anywhere-output-source"
  | "anywhere-output-markdown"
  | "anywhere-output-unicode"
  | "anywhere-output-image"
  | "anywhere-row-display"
  | "anywhere-row-inline"
  | "anywhere-row-toggle"
  | "anywhere-ocr";

export interface AnywhereHotkeyDefinition {
  id: AnywhereHotkeyAction;
  /** System-wide (registered with macOS), not only while a formula is open. */
  global: boolean;
  labelZh: string;
  labelEn: string;
  defaultChord: FormulaHotkeyChord | null;
}

const chord = (code: string, key: string, shiftKey = false): FormulaHotkeyChord => ({
  code,
  key,
  ctrlKey: true,
  altKey: true,
  shiftKey,
  metaKey: false,
});

export const ANYWHERE_HOTKEY_ACTIONS: readonly AnywhereHotkeyDefinition[] = [
  { id: "anywhere-toggle", global: true, labelZh: "进入 / 退出公式模式（全局）", labelEn: "Open / close a formula (system-wide)", defaultChord: chord("KeyM", "m") },
  { id: "anywhere-open-display", global: true, labelZh: "进入 / 退出独立公式（全局）", labelEn: "Open / close a display formula (system-wide)", defaultChord: chord("KeyM", "M", true) },
  { id: "anywhere-next-output", global: false, labelZh: "切换到下一种输出格式", labelEn: "Next output format", defaultChord: chord("KeyO", "o") },
  { id: "anywhere-output-source", global: false, labelZh: "输出格式：LaTeX", labelEn: "Output format: LaTeX", defaultChord: null },
  { id: "anywhere-output-markdown", global: false, labelZh: "输出格式：Markdown", labelEn: "Output format: Markdown", defaultChord: null },
  { id: "anywhere-output-unicode", global: false, labelZh: "输出格式：Unicode", labelEn: "Output format: Unicode", defaultChord: null },
  { id: "anywhere-output-image", global: false, labelZh: "输出格式：图片", labelEn: "Output format: image", defaultChord: null },
  { id: "anywhere-row-display", global: false, labelZh: "当前行设为独立公式", labelEn: "Make the row display math", defaultChord: chord("KeyD", "d") },
  { id: "anywhere-row-inline", global: false, labelZh: "当前行设为行内公式", labelEn: "Make the row inline math", defaultChord: chord("KeyI", "i") },
  { id: "anywhere-row-toggle", global: false, labelZh: "切换行内 / 独立公式", labelEn: "Switch the row between inline and display", defaultChord: null },
  { id: "anywhere-ocr", global: false, labelZh: "截图识别公式（OCR）", labelEn: "Screenshot OCR into the formula", defaultChord: chord("KeyS", "s") },
];

export const ANYWHERE_HOTKEY_STORAGE_KEY = "visualtex-anywhere-hotkeys-v1";

export type AnywhereHotkeyBindings = Record<AnywhereHotkeyAction, FormulaHotkeyChord | null>;

export function isAnywhereHotkeyAction(value: unknown): value is AnywhereHotkeyAction {
  return ANYWHERE_HOTKEY_ACTIONS.some((action) => action.id === value);
}

export function anywhereHotkeyDefinition(id: AnywhereHotkeyAction): AnywhereHotkeyDefinition {
  return ANYWHERE_HOTKEY_ACTIONS.find((action) => action.id === id)!;
}

export function defaultAnywhereHotkeyBindings(): AnywhereHotkeyBindings {
  return Object.fromEntries(
    ANYWHERE_HOTKEY_ACTIONS.map((action) => [action.id, action.defaultChord]),
  ) as AnywhereHotkeyBindings;
}

function normalizeChord(value: unknown): FormulaHotkeyChord | null {
  if (!value || typeof value !== "object") return null;
  const v = value as Record<string, unknown>;
  if (typeof v.code !== "string" || !v.code || v.code.length > 40) return null;
  const chord: FormulaHotkeyChord = {
    code: v.code,
    key: typeof v.key === "string" ? v.key.slice(0, 40) : "",
    ctrlKey: v.ctrlKey === true,
    altKey: v.altKey === true,
    shiftKey: v.shiftKey === true,
    metaKey: v.metaKey === true,
  };
  return formulaHotkeyHasModifier(chord) ? chord : null;
}

/** Every action, defaults for the ones not stored (an explicit null stays
 *  "none"), one action per chord. */
export function normalizeAnywhereHotkeyBindings(value: unknown): AnywhereHotkeyBindings {
  const stored = value && typeof value === "object" ? (value as Record<string, unknown>) : {};
  const bindings = defaultAnywhereHotkeyBindings();
  const used = new Set<string>();
  for (const action of ANYWHERE_HOTKEY_ACTIONS) {
    const next = action.id in stored ? normalizeChord(stored[action.id]) : action.defaultChord;
    const key = next ? formulaHotkeyChordId(next) : null;
    bindings[action.id] = key && !used.has(key) ? next : null;
    if (key) used.add(key);
  }
  return bindings;
}

/** Why `chord` cannot be `action`'s shortcut, or null when it can. */
export function anywhereHotkeyChordProblem(
  action: AnywhereHotkeyAction,
  chord: FormulaHotkeyChord,
  language: "cn" | "en" = "cn",
): string | null {
  const en = language === "en";
  if (!formulaHotkeyHasModifier(chord)) {
    return en
      ? "Add Control, Option or Command; Shift can be used in addition."
      : "请加入 Control、Option 或 Command；Shift 可作为附加修饰键。";
  }
  const reserved = protectedFormulaHotkeyAction(chord, language);
  if (reserved) {
    return en
      ? `This shortcut is reserved for ${reserved} and cannot be overridden.`
      : `该快捷键已保留用于“${reserved}”，不能覆盖。`;
  }
  if (anywhereHotkeyDefinition(action).global) {
    if (!officeHotkeySupportsCode(chord.code)) {
      return en ? "This key cannot be a system-wide shortcut." : "该按键不能用作全局快捷键。";
    }
    return null;
  }
  if (chord.metaKey) {
    return en
      ? "While a formula is open, Command shortcuts stay the application's. Use Control or Option."
      : "公式模式中 Command 组合键仍归当前应用使用，请改用 Control 或 Option。";
  }
  return null;
}

/** The action whose chord this key is (actions used while a formula is open). */
export function matchAnywhereHotkey(
  bindings: AnywhereHotkeyBindings,
  key: { code: string; ctrl?: boolean; alt?: boolean; shift?: boolean; meta?: boolean },
): AnywhereHotkeyAction | null {
  if (!key.ctrl && !key.alt && !key.meta) return null;
  const id = formulaHotkeyChordId({
    code: key.code,
    key: "",
    ctrlKey: Boolean(key.ctrl),
    altKey: Boolean(key.alt),
    shiftKey: Boolean(key.shift),
    metaKey: Boolean(key.meta),
  });
  for (const action of ANYWHERE_HOTKEY_ACTIONS) {
    const bound = bindings[action.id];
    if (!action.global && bound && formulaHotkeyChordId(bound) === id) return action.id;
  }
  return null;
}

export function anywhereHotkeyLabel(chord: FormulaHotkeyChord | null | undefined) {
  return chord ? formatFormulaHotkeyChord(chord) : "";
}
