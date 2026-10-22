import { safeStorage } from "../runtime/safeStorage.ts";
import {
  SPACING_DEFINITIONS,
  spacingPercent,
  type Spacing,
  type SpacingKey,
} from "../../../shared/tex-layout/src/spacing.ts";
import type { FormulaLetterFont } from "./formulaFontPreferences.ts";

// Formula spacing (Settings > Interface customization): for each formula font
// a list of named configurations of the layout engine's spacing adjustments
// (percentages of LaTeX's spacing, tex-layout spacing.ts) and the one in use.
// The editor, previews, caret and Office window all lay out with the active
// configuration of the chosen font.

export interface FormulaSpacingConfig {
  id: string;
  /** "" for the built-in first configuration (shown as "Default") */
  name: string;
  values: Spacing;
}

export interface FormulaSpacingFontState {
  activeId: string;
  configs: FormulaSpacingConfig[];
}

export type FormulaSpacingPreferences = Record<FormulaLetterFont, FormulaSpacingFontState>;

export const DEFAULT_FORMULA_SPACING_CONFIG_ID = "default";
const SPACING_FONTS: readonly FormulaLetterFont[] = ["katex", "times"];
const MAX_CONFIGS = 30;
const MAX_NAME_LENGTH = 60;
const FORMULA_SPACING_STORAGE_KEY = "visualtex.formula-spacing";
export const FORMULA_SPACING_CHANGED_EVENT = "visualtex-formula-spacing-changed";

function defaultFontState(): FormulaSpacingFontState {
  return {
    activeId: DEFAULT_FORMULA_SPACING_CONFIG_ID,
    configs: [{ id: DEFAULT_FORMULA_SPACING_CONFIG_ID, name: "", values: {} }],
  };
}

export function defaultFormulaSpacingPreferences(): FormulaSpacingPreferences {
  return { katex: defaultFontState(), times: defaultFontState() };
}

/** Only known adjustments, within their ranges; 100 (LaTeX's) is left out. */
export function normalizeSpacingValues(value: unknown): Spacing {
  const out: Spacing = {};
  if (!value || typeof value !== "object") return out;
  for (const d of SPACING_DEFINITIONS) {
    const raw = (value as Record<string, unknown>)[d.key];
    if (typeof raw !== "number") continue;
    const p = spacingPercent({ [d.key]: raw }, d.key);
    if (p !== 100) out[d.key] = p;
  }
  return out;
}

function normalizeFontState(value: unknown): FormulaSpacingFontState {
  if (!value || typeof value !== "object") return defaultFontState();
  const raw = value as { activeId?: unknown; configs?: unknown };
  const configs: FormulaSpacingConfig[] = [];
  const seen = new Set<string>();
  for (const item of Array.isArray(raw.configs) ? raw.configs : []) {
    if (!item || typeof item !== "object") continue;
    const c = item as { id?: unknown; name?: unknown; values?: unknown };
    if (typeof c.id !== "string" || !c.id || c.id.length > 64 || seen.has(c.id)) continue;
    seen.add(c.id);
    configs.push({
      id: c.id,
      name: typeof c.name === "string" ? c.name.slice(0, MAX_NAME_LENGTH) : "",
      values: normalizeSpacingValues(c.values),
    });
    if (configs.length >= MAX_CONFIGS) break;
  }
  if (configs.length === 0) return defaultFontState();
  const activeId = typeof raw.activeId === "string" && seen.has(raw.activeId) ? raw.activeId : configs[0].id;
  return { activeId, configs };
}

export function normalizeFormulaSpacingPreferences(value: unknown): FormulaSpacingPreferences {
  const raw = value && typeof value === "object" ? (value as Record<string, unknown>) : {};
  return { katex: normalizeFontState(raw.katex), times: normalizeFontState(raw.times) };
}

/**
 * Configurations of `incoming` (a document's or an imported configuration's)
 * take the place of those with the same id in `current` and are added
 * otherwise; their active configurations become the active ones. The user's
 * other configurations stay.
 */
export function mergeFormulaSpacingPreferences(
  current: FormulaSpacingPreferences,
  incoming: unknown,
): FormulaSpacingPreferences {
  const next = normalizeFormulaSpacingPreferences(incoming);
  const out = {} as FormulaSpacingPreferences;
  for (const font of SPACING_FONTS) {
    const configs = current[font].configs.map((c) => next[font].configs.find((n) => n.id === c.id) ?? c);
    for (const c of next[font].configs) if (!configs.some((x) => x.id === c.id)) configs.push(c);
    out[font] = { activeId: next[font].activeId, configs: configs.slice(0, MAX_CONFIGS) };
  }
  return out;
}

export function activeFormulaSpacingConfig(
  preferences: FormulaSpacingPreferences,
  font: FormulaLetterFont,
): FormulaSpacingConfig {
  const state = preferences[font];
  return state.configs.find((c) => c.id === state.activeId) ?? state.configs[0];
}

/** The adjustments the layout engine uses for a font (null: LaTeX's spacing). */
export function activeFormulaSpacing(preferences: FormulaSpacingPreferences, font: FormulaLetterFont): Spacing | null {
  const values = activeFormulaSpacingConfig(preferences, font).values;
  return Object.keys(values).length ? values : null;
}

export function readPersistedFormulaSpacing(): FormulaSpacingPreferences {
  const text = safeStorage.getItem(FORMULA_SPACING_STORAGE_KEY);
  if (!text) return defaultFormulaSpacingPreferences();
  try {
    return normalizeFormulaSpacingPreferences(JSON.parse(text));
  } catch {
    return defaultFormulaSpacingPreferences();
  }
}

/** Store the preferences; formulas are laid out again only when they changed. */
export function persistFormulaSpacing(preferences: FormulaSpacingPreferences): void {
  const text = JSON.stringify(normalizeFormulaSpacingPreferences(preferences));
  if (safeStorage.getItem(FORMULA_SPACING_STORAGE_KEY) === text) return;
  safeStorage.setItem(FORMULA_SPACING_STORAGE_KEY, text);
  if (typeof window !== "undefined") window.dispatchEvent(new CustomEvent(FORMULA_SPACING_CHANGED_EVENT));
}

// --- editing (the settings dialog) ---------------------------------------

function newConfigId(): string {
  return `c${Date.now().toString(36)}${Math.floor(Math.random() * 1e6).toString(36)}`;
}

function withFont(
  preferences: FormulaSpacingPreferences,
  font: FormulaLetterFont,
  state: FormulaSpacingFontState,
): FormulaSpacingPreferences {
  return { ...preferences, [font]: state };
}

/** Set one adjustment of the font's active configuration. */
export function setSpacingValue(
  preferences: FormulaSpacingPreferences,
  font: FormulaLetterFont,
  key: SpacingKey,
  percent: number,
): FormulaSpacingPreferences {
  const state = preferences[font];
  const active = activeFormulaSpacingConfig(preferences, font);
  const values = normalizeSpacingValues({ ...active.values, [key]: percent });
  return withFont(preferences, font, {
    ...state,
    configs: state.configs.map((c) => (c.id === active.id ? { ...c, values } : c)),
  });
}

/** Every adjustment of the font's active configuration back to LaTeX's (100%). */
export function resetActiveSpacing(preferences: FormulaSpacingPreferences, font: FormulaLetterFont): FormulaSpacingPreferences {
  const state = preferences[font];
  return withFont(preferences, font, {
    ...state,
    configs: state.configs.map((c) => (c.id === state.activeId ? { ...c, values: {} } : c)),
  });
}

export function selectSpacingConfig(
  preferences: FormulaSpacingPreferences,
  font: FormulaLetterFont,
  id: string,
): FormulaSpacingPreferences {
  const state = preferences[font];
  if (!state.configs.some((c) => c.id === id)) return preferences;
  return withFont(preferences, font, { ...state, activeId: id });
}

/** A copy of the active configuration under a new name, made active; null when the list is full. */
export function saveSpacingConfigAs(
  preferences: FormulaSpacingPreferences,
  font: FormulaLetterFont,
  name: string,
): FormulaSpacingPreferences | null {
  const state = preferences[font];
  if (state.configs.length >= MAX_CONFIGS) return null;
  const active = activeFormulaSpacingConfig(preferences, font);
  const config: FormulaSpacingConfig = { id: newConfigId(), name: name.trim().slice(0, MAX_NAME_LENGTH), values: { ...active.values } };
  return withFont(preferences, font, { activeId: config.id, configs: [...state.configs, config] });
}

export function renameSpacingConfig(
  preferences: FormulaSpacingPreferences,
  font: FormulaLetterFont,
  id: string,
  name: string,
): FormulaSpacingPreferences {
  const state = preferences[font];
  return withFont(preferences, font, {
    ...state,
    configs: state.configs.map((c) => (c.id === id ? { ...c, name: name.trim().slice(0, MAX_NAME_LENGTH) } : c)),
  });
}

/** Remove a configuration (the last one stays); the first remaining becomes active if it was. */
export function deleteSpacingConfig(
  preferences: FormulaSpacingPreferences,
  font: FormulaLetterFont,
  id: string,
): FormulaSpacingPreferences {
  const state = preferences[font];
  if (state.configs.length <= 1) return preferences;
  const configs = state.configs.filter((c) => c.id !== id);
  return withFont(preferences, font, { activeId: state.activeId === id ? configs[0].id : state.activeId, configs });
}

export function formulaSpacingConfigCount(preferences: FormulaSpacingPreferences, font: FormulaLetterFont): number {
  return preferences[font].configs.length;
}

export const FORMULA_SPACING_MAX_CONFIGS = MAX_CONFIGS;
