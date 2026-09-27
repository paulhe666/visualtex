import type {
  PowerPointFormulaShortcutAction,
  PowerPointFormulaShortcutSettings,
  WordFormulaInsertFormat,
  WordFormulaShortcutSettings,
} from "../types/formula";

export const DEFAULT_WORD_FORMULA_SHORTCUTS: WordFormulaShortcutSettings = {
  inline: { shortcut: null, format: "ole" },
  display: { shortcut: null, format: "ole" },
};

export const DEFAULT_POWERPOINT_FORMULA_SHORTCUTS: PowerPointFormulaShortcutSettings = {
  svg: null,
  ole: null,
  inlineOmml: null,
  displayOmml: null,
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

export type OfficeFormulaShortcutCapture =
  | { shortcut: string }
  | { error: "modifier" | "unsupported" | "reserved" };

export function normalizeOfficeFormulaShortcut(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const normalized = value.trim();
  const match = canonicalShortcutPattern.exec(normalized);
  if (!match || (!match[1] && !match[2])) return null;
  const parts = [match[1], match[2], match[3], match[4]].filter(Boolean);
  return parts.join("+");
}

function normalizeWordFormat(value: unknown): WordFormulaInsertFormat {
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
  const inlineShortcut = normalizeOfficeFormulaShortcut(
    candidate.inline?.shortcut,
  );
  let displayShortcut = normalizeOfficeFormulaShortcut(
    candidate.display?.shortcut,
  );
  if (inlineShortcut && displayShortcut === inlineShortcut) {
    displayShortcut = null;
  }
  return {
    inline: {
      shortcut: inlineShortcut,
      format: normalizeWordFormat(candidate.inline?.format),
    },
    display: {
      shortcut: displayShortcut,
      format: normalizeWordFormat(candidate.display?.format),
    },
  };
}

const powerPointActionOrder: readonly PowerPointFormulaShortcutAction[] = [
  "svg",
  "ole",
  "inlineOmml",
  "displayOmml",
];

function distinctPowerPointShortcuts(
  candidate: Partial<PowerPointFormulaShortcutSettings>,
): PowerPointFormulaShortcutSettings {
  const result = { ...DEFAULT_POWERPOINT_FORMULA_SHORTCUTS };
  const used = new Set<string>();
  for (const action of powerPointActionOrder) {
    const shortcut = normalizeOfficeFormulaShortcut(candidate[action]);
    if (!shortcut || used.has(shortcut)) continue;
    result[action] = shortcut;
    used.add(shortcut);
  }
  return result;
}

export function normalizePowerPointFormulaShortcutSettings(
  value: unknown,
): PowerPointFormulaShortcutSettings {
  if (!value || typeof value !== "object") {
    return { ...DEFAULT_POWERPOINT_FORMULA_SHORTCUTS };
  }
  const candidate = value as Record<string, unknown>;
  if (powerPointActionOrder.some((action) => action in candidate)) {
    return distinctPowerPointShortcuts(
      candidate as Partial<PowerPointFormulaShortcutSettings>,
    );
  }

  // One-time semantic migration from the old Word-shaped PowerPoint settings:
  // OMML keeps its inline/display meaning; the two fake OLE modes collapse into
  // the one real "new OLE object" Ribbon action.
  const legacy = candidate as {
    inline?: { shortcut?: unknown; format?: unknown };
    display?: { shortcut?: unknown; format?: unknown };
  };
  const inlineShortcut = normalizeOfficeFormulaShortcut(legacy.inline?.shortcut);
  const displayShortcut = normalizeOfficeFormulaShortcut(legacy.display?.shortcut);
  const migrated: Partial<PowerPointFormulaShortcutSettings> = {};
  if (legacy.inline?.format === "omml") migrated.inlineOmml = inlineShortcut;
  else migrated.ole = inlineShortcut;
  if (legacy.display?.format === "omml") migrated.displayOmml = displayShortcut;
  else if (!migrated.ole) migrated.ole = displayShortcut;
  return distinctPowerPointShortcuts(migrated);
}

export function officeFormulaShortcutFromEvent(
  event: Pick<
    KeyboardEvent,
    "code" | "ctrlKey" | "altKey" | "shiftKey" | "metaKey"
  >,
): OfficeFormulaShortcutCapture {
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
