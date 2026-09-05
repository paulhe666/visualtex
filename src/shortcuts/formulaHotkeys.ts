import { commandRegistry } from "../autocomplete/commandRegistry";
import type { LatexCommand } from "../types/command";

export type FormulaHotkeyTargetKind =
  | "command"
  | "common-tile"
  | "custom-tile"
  | "matrix";

export interface FormulaHotkeyChord {
  code: string;
  key: string;
  ctrlKey: boolean;
  altKey: boolean;
  shiftKey: boolean;
  metaKey: boolean;
}

export interface FormulaHotkeyTarget {
  id: string;
  kind: FormulaHotkeyTargetKind;
  command: LatexCommand;
}

export interface FormulaHotkeyBinding {
  id: string;
  target: FormulaHotkeyTarget;
  chord: FormulaHotkeyChord;
  updatedAt: number;
}

function commandById(commandId: string) {
  const command = commandRegistry.find((item) => item.id === commandId);
  if (!command) throw new Error(`Missing formula hotkey command: ${commandId}`);
  return command;
}

function defaultHotkeyCommand(
  id: string,
  command: string,
  insertTemplate: string,
  previewLatex: string,
  labelZh: string,
  labelEn: string,
): LatexCommand {
  return {
    id,
    command,
    insertTemplate,
    previewLatex,
    labelZh,
    labelEn,
    aliases: [],
    keywords: [],
    category: "common",
    defaultPriority: 0,
    supportedInMathMode: true,
  };
}

function controlChord(code: string, shiftKey = false): FormulaHotkeyChord {
  return {
    code,
    key: /^Key[A-Z]$/.test(code) ? code.slice(3).toLowerCase() : code,
    ctrlKey: true,
    altKey: false,
    shiftKey,
    metaKey: false,
  };
}

function altShiftChord(code: string): FormulaHotkeyChord {
  return {
    code,
    key: /^Key[A-Z]$/.test(code) ? code.slice(3).toLowerCase() : code,
    ctrlKey: false,
    altKey: true,
    shiftKey: true,
    metaKey: false,
  };
}

const partialSymbolCommand = defaultHotkeyCommand(
  "default-partial-symbol",
  "\\partial",
  "\\partial",
  "\\partial",
  "偏微分符号",
  "Partial differential symbol",
);

const boundedIntegralTemplateCommand = defaultHotkeyCommand(
  "default-bounded-integral-template",
  "\\int",
  "\\int_{\\placeholder{}}^{\\placeholder{}} \\placeholder{}",
  "\\int_a^b f(x)",
  "定积分模板",
  "Definite integral template",
);

const underlinedXCommand = defaultHotkeyCommand(
  "default-underlined-x",
  "\\underline{X}",
  "\\underline{X}",
  "\\underline{X}",
  "下划线 X",
  "Underlined X",
);

const underlinedYCommand = defaultHotkeyCommand(
  "default-underlined-y",
  "\\underline{Y}",
  "\\underline{Y}",
  "\\underline{Y}",
  "下划线 Y",
  "Underlined Y",
);

const defaultFormulaHotkeyDefinitions: Array<{
  command: LatexCommand;
  chord: FormulaHotkeyChord;
}> = [
  { command: commandById("sqrt"), chord: controlChord("KeyR") },
  { command: commandById("frac"), chord: controlChord("KeyF") },
  { command: commandById("scripts"), chord: controlChord("KeyJ") },
  { command: commandById("subscript"), chord: controlChord("KeyL") },
  { command: commandById("degree"), chord: controlChord("KeyD") },
  { command: partialSymbolCommand, chord: controlChord("KeyP", true) },
  {
    command: boundedIntegralTemplateCommand,
    chord: controlChord("KeyI", true),
  },
  { command: commandById("sum"), chord: controlChord("KeyS", true) },
  { command: underlinedXCommand, chord: altShiftChord("KeyX") },
  { command: underlinedYCommand, chord: altShiftChord("KeyY") },
];

export function createDefaultFormulaHotkeyBindings(): FormulaHotkeyBinding[] {
  return defaultFormulaHotkeyDefinitions.map(({ command, chord }) => ({
    id: `default:${command.id}`,
    target: createFormulaHotkeyTarget(
      formulaHotkeyTargetIdForCommand(command.id),
      "command",
      command,
    ),
    chord: { ...chord },
    updatedAt: 0,
  }));
}

const modifierCodes = new Set([
  "AltLeft",
  "AltRight",
  "ControlLeft",
  "ControlRight",
  "MetaLeft",
  "MetaRight",
  "ShiftLeft",
  "ShiftRight",
]);

const keyLabelsByCode: Record<string, string> = {
  ArrowDown: "↓",
  ArrowLeft: "←",
  ArrowRight: "→",
  ArrowUp: "↑",
  Backquote: "`",
  Backslash: "\\",
  Backspace: "Backspace",
  BracketLeft: "[",
  BracketRight: "]",
  Comma: ",",
  Delete: "Delete",
  End: "End",
  Enter: "Enter",
  Equal: "=",
  Escape: "Esc",
  Home: "Home",
  Minus: "-",
  PageDown: "Page Down",
  PageUp: "Page Up",
  Period: ".",
  Quote: "'",
  Semicolon: ";",
  Slash: "/",
  Space: "Space",
  Tab: "Tab",
};

export function isMacKeyboardPlatform() {
  if (typeof navigator === "undefined") return false;
  const platform = navigator.platform || navigator.userAgent;
  return /Mac|iPhone|iPad|iPod/i.test(platform);
}

export function formulaHotkeyChordId(chord: FormulaHotkeyChord) {
  return [
    chord.ctrlKey ? "ctrl" : "",
    chord.altKey ? "alt" : "",
    chord.shiftKey ? "shift" : "",
    chord.metaKey ? "meta" : "",
    chord.code,
  ].join(":");
}

export function formulaHotkeyTargetIdForCommand(commandId: string) {
  return `command:${commandId}`;
}

function stableStringHash(value: string) {
  let hash = 2166136261;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(36);
}

export function formulaHotkeyTargetIdForTile(
  kind: "common-tile" | "custom-tile",
  tileId: string,
  latex: string,
) {
  return kind === "common-tile"
    ? `tile:common:${tileId}`
    : `tile:custom:${stableStringHash(latex)}`;
}

export function createFormulaHotkeyTarget(
  id: string,
  kind: FormulaHotkeyTargetKind,
  command: LatexCommand,
): FormulaHotkeyTarget {
  return {
    id,
    kind,
    command: { ...command },
  };
}

export function resolveFormulaHotkeyCommand(target: FormulaHotkeyTarget) {
  if (target.kind === "command") {
    return (
      commandRegistry.find((command) => command.id === target.command.id) ??
      target.command
    );
  }
  return target.command;
}

export function formulaHotkeyTargetLabel(
  target: FormulaHotkeyTarget,
  language: "cn" | "en",
) {
  return language === "en"
    ? target.command.labelEn
    : target.command.labelZh;
}

export function formulaHotkeyTargetKindLabel(
  target: FormulaHotkeyTarget,
  language: "cn" | "en",
) {
  const labels = language === "en"
    ? {
        command: "Formula tool",
        "common-tile": "Common tile",
        "custom-tile": "Custom tile",
        matrix: "Custom matrix",
      }
    : {
        command: "公式工具",
        "common-tile": "常用磁贴",
        "custom-tile": "自定义磁贴",
        matrix: "自定义矩阵",
      };
  return labels[target.kind];
}

export function formulaHotkeyChordFromEvent(
  event: Pick<
    KeyboardEvent,
    "code" | "key" | "ctrlKey" | "altKey" | "shiftKey" | "metaKey"
  >,
): FormulaHotkeyChord | null {
  if (!event.code || modifierCodes.has(event.code)) return null;
  return {
    code: event.code,
    key: event.key,
    ctrlKey: event.ctrlKey,
    altKey: event.altKey,
    shiftKey: event.shiftKey,
    metaKey: event.metaKey,
  };
}

export function formulaHotkeyHasModifier(chord: FormulaHotkeyChord) {
  // Shift by itself is normal text input (for example, Shift+F types an
  // uppercase F), so a formula hotkey must include Ctrl, Option/Alt or Command.
  return chord.ctrlKey || chord.altKey || chord.metaKey;
}

function displayKeyForChord(chord: FormulaHotkeyChord) {
  if (keyLabelsByCode[chord.code]) return keyLabelsByCode[chord.code];
  if (/^Key[A-Z]$/.test(chord.code)) return chord.code.slice(3);
  if (/^Digit[0-9]$/.test(chord.code)) return chord.code.slice(5);
  if (/^Numpad[0-9]$/.test(chord.code)) return `Num ${chord.code.slice(6)}`;
  if (/^F(?:[1-9]|1[0-9]|2[0-4])$/.test(chord.code)) return chord.code;
  const key = chord.key.length === 1 ? chord.key.toLocaleUpperCase() : chord.key;
  return key || chord.code;
}

export function formatFormulaHotkeyChord(
  chord: FormulaHotkeyChord,
  mac = isMacKeyboardPlatform(),
) {
  const key = displayKeyForChord(chord);
  if (mac) {
    return [
      chord.ctrlKey ? "⌃" : "",
      chord.altKey ? "⌥" : "",
      chord.shiftKey ? "⇧" : "",
      chord.metaKey ? "⌘" : "",
      key,
    ].join("");
  }
  return [
    chord.ctrlKey ? "Ctrl" : "",
    chord.altKey ? "Alt" : "",
    chord.shiftKey ? "Shift" : "",
    chord.metaKey ? "Meta" : "",
    key,
  ]
    .filter(Boolean)
    .join("+");
}

function normalizedChordKey(chord: FormulaHotkeyChord) {
  const key = chord.key.toLocaleLowerCase();
  if (key && key !== "unidentified") return key;
  if (/^Key[A-Z]$/.test(chord.code)) return chord.code.slice(3).toLowerCase();
  if (/^Digit[0-9]$/.test(chord.code)) return chord.code.slice(5);
  if (chord.code === "Comma") return ",";
  if (chord.code === "Equal") return chord.shiftKey ? "+" : "=";
  if (chord.code === "Minus") return chord.shiftKey ? "_" : "-";
  return "";
}

export function protectedFormulaHotkeyAction(
  chord: FormulaHotkeyChord,
  language: "cn" | "en" = "cn",
): string | null {
  const primaryModifier =
    (chord.metaKey || chord.ctrlKey) && !chord.altKey;
  if (!primaryModifier) return null;

  const key = normalizedChordKey(chord);
  const labels = language === "en"
    ? {
        copy: "Copy",
        cut: "Cut",
        paste: "Paste",
        undo: "Undo",
        redo: "Redo",
        new: "New",
        open: "Open",
        save: "Save",
        greekMode: "Greek letter mode",
        settings: "Settings",
        resetZoom: "Reset zoom",
        zoomIn: "Zoom in",
        zoomOut: "Zoom out",
      }
    : {
        copy: "复制",
        cut: "剪切",
        paste: "粘贴",
        undo: "撤销",
        redo: "重做",
        new: "新建",
        open: "打开",
        save: "保存",
        greekMode: "希腊字母输入",
        settings: "设置",
        resetZoom: "恢复缩放",
        zoomIn: "放大",
        zoomOut: "缩小",
      };

  if (!chord.shiftKey && key === "c") return labels.copy;
  if (!chord.shiftKey && key === "x") return labels.cut;
  if (!chord.shiftKey && key === "v") return labels.paste;
  if (key === "z") return chord.shiftKey ? labels.redo : labels.undo;
  if (!chord.shiftKey && key === "y") return labels.redo;
  if (key === "n") return labels.new;
  if (key === "o") return labels.open;
  if (!chord.shiftKey && key === "s") return labels.save;
  if (
    chord.ctrlKey &&
    !chord.metaKey &&
    !chord.altKey &&
    !chord.shiftKey &&
    key === "g"
  ) {
    return labels.greekMode;
  }
  if (key === ",") return labels.settings;
  if (key === "0") return labels.resetZoom;
  if (key === "=" || key === "+") return labels.zoomIn;
  if (key === "-") return labels.zoomOut;
  return null;
}

export function matchFormulaHotkey(
  event: KeyboardEvent,
  bindings: readonly FormulaHotkeyBinding[],
) {
  if (
    event.repeat ||
    event.isComposing ||
    event.key === "Process" ||
    event.getModifierState("AltGraph")
  ) {
    return null;
  }
  const chord = formulaHotkeyChordFromEvent(event);
  if (
    !chord ||
    !formulaHotkeyHasModifier(chord) ||
    protectedFormulaHotkeyAction(chord)
  ) {
    return null;
  }
  const chordId = formulaHotkeyChordId(chord);
  return (
    bindings.find(
      (binding) => formulaHotkeyChordId(binding.chord) === chordId,
    ) ?? null
  );
}
