import {
  formulaHotkeyChordFromEvent,
  formulaHotkeyChordId,
  formulaHotkeyHasModifier,
  protectedFormulaHotkeyAction,
  type FormulaHotkeyChord,
} from "./formulaHotkeys";

// Button shortcuts (Settings > Keyboard shortcuts): the user may give the
// common buttons of the main window and of the Office editor window a key
// chord. None is set at first. The chords live in the editor settings, so the
// Office editor window receives them with the other preferences, and they
// travel with configuration backups.

export type ButtonShortcutId =
  | "quick-ocr"
  | "ocr"
  | "copy-latex"
  | "copy-image"
  | "toggle-source"
  | "export"
  | "history"
  | "toggle-tiles"
  | "silent-ocr"
  | "settings"
  | "help"
  | "office-cancel";

/** Where a button is: the main window, the Office editor window, or both. */
export type ButtonShortcutScope = "main" | "office" | "both";

export interface ButtonShortcutDefinition {
  id: ButtonShortcutId;
  scope: ButtonShortcutScope;
  labelZh: string;
  labelEn: string;
}

export const BUTTON_SHORTCUT_ACTIONS: readonly ButtonShortcutDefinition[] = [
  { id: "quick-ocr", scope: "both", labelZh: "快捷截图识别（快捷 OCR）", labelEn: "Quick screenshot OCR" },
  { id: "ocr", scope: "both", labelZh: "打开 OCR 识别窗口", labelEn: "Open OCR" },
  { id: "copy-latex", scope: "both", labelZh: "复制 LaTeX", labelEn: "Copy LaTeX" },
  { id: "copy-image", scope: "both", labelZh: "复制为图片", labelEn: "Copy as image" },
  { id: "toggle-source", scope: "both", labelZh: "显示 / 隐藏源码", labelEn: "Show / hide source" },
  { id: "export", scope: "main", labelZh: "导出", labelEn: "Export" },
  { id: "history", scope: "main", labelZh: "历史记录", labelEn: "History" },
  { id: "toggle-tiles", scope: "main", labelZh: "显示 / 隐藏公式面板", labelEn: "Show / hide formula panel" },
  { id: "silent-ocr", scope: "main", labelZh: "开启 / 关闭静默 OCR", labelEn: "Turn silent OCR on / off" },
  { id: "settings", scope: "main", labelZh: "设置", labelEn: "Settings" },
  { id: "help", scope: "main", labelZh: "帮助", labelEn: "Help" },
  { id: "office-cancel", scope: "office", labelZh: "取消编辑（编辑器窗口）", labelEn: "Cancel editing (editor window)" },
];

const ACTION_IDS = new Set<string>(BUTTON_SHORTCUT_ACTIONS.map((a) => a.id));

export type ButtonShortcutSettings = Partial<Record<ButtonShortcutId, FormulaHotkeyChord>>;

/** The event a matched chord sends (detail: { action }); the window that owns the button acts on it. */
export const BUTTON_SHORTCUT_EVENT = "visualtex-button-shortcut";

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

/** Only known actions with valid chords, one action per chord. */
export function normalizeButtonShortcutSettings(value: unknown): ButtonShortcutSettings {
  const out: ButtonShortcutSettings = {};
  if (!value || typeof value !== "object") return out;
  const used = new Set<string>();
  for (const [id, raw] of Object.entries(value as Record<string, unknown>)) {
    if (!ACTION_IDS.has(id)) continue;
    const chord = normalizeChord(raw);
    if (!chord) continue;
    const key = formulaHotkeyChordId(chord);
    if (used.has(key)) continue;
    used.add(key);
    out[id as ButtonShortcutId] = chord;
  }
  return out;
}

/** Give `id` the chord (null: none); another button that had the chord loses it. */
export function withButtonShortcut(
  settings: ButtonShortcutSettings,
  id: ButtonShortcutId,
  chord: FormulaHotkeyChord | null,
): ButtonShortcutSettings {
  const next: ButtonShortcutSettings = {};
  const key = chord ? formulaHotkeyChordId(chord) : null;
  for (const [other, c] of Object.entries(settings) as [ButtonShortcutId, FormulaHotkeyChord][]) {
    if (other === id || !c) continue;
    if (key && formulaHotkeyChordId(c) === key) continue;
    next[other] = c;
  }
  if (chord) next[id] = chord;
  return next;
}

export function buttonShortcutLabel(id: ButtonShortcutId, language: "cn" | "en" = "cn"): string {
  const action = BUTTON_SHORTCUT_ACTIONS.find((a) => a.id === id);
  return action ? (language === "en" ? action.labelEn : action.labelZh) : id;
}

/** The button that has this chord, if any. */
export function buttonShortcutWithChord(
  settings: ButtonShortcutSettings,
  chord: FormulaHotkeyChord,
): ButtonShortcutId | null {
  const key = formulaHotkeyChordId(chord);
  for (const [id, c] of Object.entries(settings) as [ButtonShortcutId, FormulaHotkeyChord][]) {
    if (c && formulaHotkeyChordId(c) === key) return id;
  }
  return null;
}

/** Chords the app keeps for itself besides the protected editing ones. */
function reservedChordAction(chord: FormulaHotkeyChord, language: "cn" | "en"): string | null {
  const only = (ctrl: boolean, alt: boolean, shift: boolean) =>
    chord.ctrlKey === ctrl && chord.altKey === alt && chord.shiftKey === shift && !chord.metaKey;
  if (chord.code === "KeyK" && only(true, false, true)) return language === "en" ? "the keypad" : "小键盘";
  if (chord.code === "Backslash" && only(true, false, false)) return language === "en" ? "switching between formula and source" : "公式 / 源码切换";
  if (chord.code === "KeyO" && only(true, true, false)) return language === "en" ? "silent OCR (system-wide)" : "静默 OCR（全局）";
  if (chord.code === "KeyA" && only(true, false, false)) return language === "en" ? "Select all" : "全选";
  if (chord.code === "KeyG" && only(true, false, false)) return language === "en" ? "Greek letters" : "希腊字母";
  return null;
}

/** Why a chord cannot be a button shortcut, or null when it can. */
export function buttonShortcutChordProblem(
  chord: FormulaHotkeyChord,
  language: "cn" | "en" = "cn",
): string | null {
  if (!formulaHotkeyHasModifier(chord)) {
    return language === "en"
      ? "Add Ctrl, Option/Alt or Command; Shift can be used in addition."
      : "请加入 Ctrl、Option/Alt 或 Command；Shift 可作为附加修饰键。";
  }
  const reserved = protectedFormulaHotkeyAction(chord, language) ?? reservedChordAction(chord, language);
  if (reserved) {
    return language === "en"
      ? `This shortcut is reserved for ${reserved} and cannot be overridden.`
      : `该快捷键已保留用于“${reserved}”，不能覆盖。`;
  }
  return null;
}

function belongsTo(scope: ButtonShortcutScope, window: "main" | "office") {
  return scope === "both" || scope === window;
}

/** True while a shortcut is being recorded (the keys belong to the recorder). */
function recordingShortcut() {
  return !!document.querySelector(
    ".formula-hotkey-recorder-dialog, .office-formula-shortcut-capture.is-capturing, [data-button-shortcut-capturing]",
  );
}

/**
 * Watch the keyboard of this window (`main` or `office`): a chord the user gave
 * one of its buttons sends BUTTON_SHORTCUT_EVENT. It runs before the editor's
 * own keys (capture phase on window), so a chosen chord takes precedence.
 */
export function installButtonShortcuts(
  window_: "main" | "office",
  settings: () => ButtonShortcutSettings,
): () => void {
  if (typeof window === "undefined") return () => {};
  const onKeyDown = (event: KeyboardEvent) => {
    if (event.repeat || event.isComposing || event.keyCode === 229) return;
    if (!event.ctrlKey && !event.altKey && !event.metaKey) return;
    const chord = formulaHotkeyChordFromEvent(event);
    if (!chord || recordingShortcut()) return;
    const id = buttonShortcutWithChord(settings(), chord);
    const action = id ? BUTTON_SHORTCUT_ACTIONS.find((a) => a.id === id) : undefined;
    if (!action || !belongsTo(action.scope, window_)) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    window.dispatchEvent(new CustomEvent(BUTTON_SHORTCUT_EVENT, { detail: { action: action.id } }));
  };
  window.addEventListener("keydown", onKeyDown, true);
  return () => window.removeEventListener("keydown", onKeyDown, true);
}

/** Run `handler` for the button shortcuts of this window's buttons. */
export function onButtonShortcut(handler: (action: ButtonShortcutId) => void): () => void {
  const listener = (event: Event) => {
    const action = (event as CustomEvent<{ action?: ButtonShortcutId }>).detail?.action;
    if (action) handler(action);
  };
  window.addEventListener(BUTTON_SHORTCUT_EVENT, listener);
  return () => window.removeEventListener(BUTTON_SHORTCUT_EVENT, listener);
}
