import type { FormulaHotkeyChord } from "./formulaHotkeys";

export type OfficeHotkeyActionId =
  | "word-image-inline"
  | "word-image-display"
  | "word-omml-inline"
  | "word-omml-display"
  | "powerpoint-svg-new"
  | "powerpoint-omml-inline"
  | "powerpoint-omml-display";

export type OfficeHotkeyHost = "word" | "powerpoint";

export interface OfficeHotkeyActionDefinition {
  id: OfficeHotkeyActionId;
  host: OfficeHotkeyHost;
  labelZh: string;
  labelEn: string;
  ribbonLabelZh: string;
  ribbonLabelEn: string;
}

export interface OfficeHotkeyBinding {
  actionId: OfficeHotkeyActionId;
  chord: FormulaHotkeyChord;
  updatedAt: number;
}

export const OFFICE_HOTKEY_STORAGE_KEY = "visualtex-office-hotkeys-v1";

// These entries intentionally mirror the creation buttons shipped in the
// macOS Word DOTM and PowerPoint PPAM. Native execution maps the ids to the
// same public VBA procedures used by those Ribbon callbacks.
export const OFFICE_HOTKEY_ACTIONS: readonly OfficeHotkeyActionDefinition[] = [
  {
    id: "word-image-inline",
    host: "word",
    labelZh: "Word OLE 行内公式",
    labelEn: "Word inline OLE formula",
    ribbonLabelZh: "OLE 行内公式",
    ribbonLabelEn: "Inline OLE",
  },
  {
    id: "word-image-display",
    host: "word",
    labelZh: "Word OLE 行间公式",
    labelEn: "Word display OLE formula",
    ribbonLabelZh: "OLE 行间公式",
    ribbonLabelEn: "Display OLE",
  },
  {
    id: "word-omml-inline",
    host: "word",
    labelZh: "Word OMML 行内公式",
    labelEn: "Word inline OMML formula",
    ribbonLabelZh: "OMML 行内公式",
    ribbonLabelEn: "Inline OMML",
  },
  {
    id: "word-omml-display",
    host: "word",
    labelZh: "Word OMML 行间公式",
    labelEn: "Word display OMML formula",
    ribbonLabelZh: "OMML 行间公式",
    ribbonLabelEn: "Display OMML",
  },
  {
    id: "powerpoint-svg-new",
    host: "powerpoint",
    labelZh: "PowerPoint 新建 SVG 公式",
    labelEn: "PowerPoint new SVG formula",
    ribbonLabelZh: "新建公式",
    ribbonLabelEn: "New formula",
  },
  {
    id: "powerpoint-omml-inline",
    host: "powerpoint",
    labelZh: "PowerPoint 行内 OMML",
    labelEn: "PowerPoint inline OMML",
    ribbonLabelZh: "行内 OMML",
    ribbonLabelEn: "Inline OMML",
  },
  {
    id: "powerpoint-omml-display",
    host: "powerpoint",
    labelZh: "PowerPoint 行间 OMML",
    labelEn: "PowerPoint display OMML",
    ribbonLabelZh: "行间 OMML",
    ribbonLabelEn: "Display OMML",
  },
] as const;

const actionIds = new Set<OfficeHotkeyActionId>(
  OFFICE_HOTKEY_ACTIONS.map((action) => action.id),
);

const supportedCodes = new Set([
  "KeyA", "KeyB", "KeyC", "KeyD", "KeyE", "KeyF", "KeyG", "KeyH",
  "KeyI", "KeyJ", "KeyK", "KeyL", "KeyM", "KeyN", "KeyO", "KeyP",
  "KeyQ", "KeyR", "KeyS", "KeyT", "KeyU", "KeyV", "KeyW", "KeyX",
  "KeyY", "KeyZ",
  "Digit0", "Digit1", "Digit2", "Digit3", "Digit4", "Digit5", "Digit6",
  "Digit7", "Digit8", "Digit9",
  "Backquote", "Backslash", "BracketLeft", "BracketRight", "Comma", "Equal",
  "Minus", "Period", "Quote", "Semicolon", "Slash", "Space", "Tab",
  "Enter", "Backspace", "Delete", "Escape", "Home", "End", "PageUp",
  "PageDown", "ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown",
  "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10",
  "F11", "F12", "F13", "F14", "F15", "F16", "F17", "F18", "F19",
  "F20",
]);

export function isOfficeHotkeyActionId(
  value: unknown,
): value is OfficeHotkeyActionId {
  return typeof value === "string" && actionIds.has(value as OfficeHotkeyActionId);
}

export function officeHotkeyAction(
  actionId: OfficeHotkeyActionId,
): OfficeHotkeyActionDefinition {
  const action = OFFICE_HOTKEY_ACTIONS.find((item) => item.id === actionId);
  if (!action) throw new Error(`Missing Office hotkey action: ${actionId}`);
  return action;
}

export function officeHotkeyActionLabel(
  action: OfficeHotkeyActionDefinition,
  language: "cn" | "en",
) {
  return language === "en" ? action.labelEn : action.labelZh;
}

export function officeHotkeySupportsCode(code: string) {
  return supportedCodes.has(code);
}
