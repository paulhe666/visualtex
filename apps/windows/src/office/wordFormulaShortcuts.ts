import type {
  WordFormulaInsertFormat,
  WordFormulaShortcutSettings,
} from "../types/formula";

export const DEFAULT_WORD_FORMULA_SHORTCUTS: WordFormulaShortcutSettings = {
  inline: { shortcut: null, format: "ole" },
  display: { shortcut: null, format: "ole" },
};

const supportedCodeKeys: Readonly<Record<string, string>> = {
  ...Object.fromEntries(
    Array.from({ length: 26 }, (_, index) => {
      const letter = String.fromCharCode(65 + index);
      return [`Key${letter}`, letter];
    }),
  ),
  ...Object.fromEntries(
    Array.from({ length: 10 }, (_, index) => [`Digit${index}`, String(index)]),
  ),
  ...Object.fromEntries(
    Array.from({ length: 24 }, (_, index) => [`F${index + 1}`, `F${index + 1}`]),
  ),
};

const canonicalShortcutPattern =
  /^(?:(Ctrl)\+)?(?:(Alt)\+)?(?:(Shift)\+)?([A-Z]|[0-9]|F(?:[1-9]|1[0-9]|2[0-4]))$/;

export type WordFormulaShortcutCapture =
  | { shortcut: string }
  | { error: "modifier" | "unsupported" | "reserved" };

export function normalizeWordFormulaShortcut(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const normalized = value.trim();
  const match = canonicalShortcutPattern.exec(normalized);
  if (!match || (!match[1] && !match[2])) return null;
  const parts = [match[1], match[2], match[3], match[4]].filter(Boolean);
  return parts.join("+");
}

function normalizeFormat(value: unknown): WordFormulaInsertFormat {
  return value === "omml" ? "omml" : "ole";
}

export function normalizeWordFormulaShortcutSettings(
  value: unknown,
): WordFormulaShortcutSettings {
  if (!value || typeof value !== "object") {
    return {
      inline: { ...DEFAULT_WORD_FORMULA_SHORTCUTS.inline },
      display: { ...DEFAULT_WORD_FORMULA_SHORTCUTS.display },
    };
  }
  const candidate = value as Partial<WordFormulaShortcutSettings>;
  const inlineShortcut = normalizeWordFormulaShortcut(
    candidate.inline?.shortcut,
  );
  let displayShortcut = normalizeWordFormulaShortcut(
    candidate.display?.shortcut,
  );
  if (inlineShortcut && displayShortcut === inlineShortcut) {
    displayShortcut = null;
  }
  return {
    inline: {
      shortcut: inlineShortcut,
      format: normalizeFormat(candidate.inline?.format),
    },
    display: {
      shortcut: displayShortcut,
      format: normalizeFormat(candidate.display?.format),
    },
  };
}

export function wordFormulaShortcutFromEvent(
  event: Pick<
    KeyboardEvent,
    "code" | "ctrlKey" | "altKey" | "shiftKey" | "metaKey"
  >,
): WordFormulaShortcutCapture {
  if (event.metaKey) return { error: "unsupported" };
  if (!event.ctrlKey && !event.altKey) return { error: "modifier" };
  const key = supportedCodeKeys[event.code];
  if (!key) return { error: "unsupported" };
  const shortcut = [
    event.ctrlKey ? "Ctrl" : "",
    event.altKey ? "Alt" : "",
    event.shiftKey ? "Shift" : "",
    key,
  ]
    .filter(Boolean)
    .join("+");
  if (shortcut === "Ctrl+Alt+O") return { error: "reserved" };
  return { shortcut };
}
