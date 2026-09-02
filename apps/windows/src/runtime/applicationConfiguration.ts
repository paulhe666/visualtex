import { invoke, isTauri } from "@tauri-apps/api/core";
import { useEditorStore } from "../stores/editorStore";
import { safeStorage } from "./safeStorage";
import { publishSynchronizedTheme } from "../themeSync";
import { OCR_MODELS } from "../ocr/ocrService";
import { normalizeQuickOcrCaptureMode } from "../ocr/quickOcr";
import type { CommandUsage } from "../types/command";
import type { FormulaDocument, FormulaHistoryItem } from "../types/formula";
import {
  CUSTOM_SYMBOL_STORAGE_KEY,
  readCustomSymbolLibrary,
  replaceCustomSymbolLibrary,
} from "../math/customSymbolRegistry";

export const VISUALTEX_CONFIGURATION_SCHEMA = "visualtex-user-configuration";
export const VISUALTEX_CONFIGURATION_VERSION = 1;
export const VISUALTEX_CONFIGURATION_EXTENSION = "vtxconfig";

export interface VisualTexConfigurationWindowSize {
  width: number;
  height: number;
}

export interface VisualTexConfigurationWindowState {
  main?: VisualTexConfigurationWindowSize | null;
  keypad?: VisualTexConfigurationWindowSize | null;
  officeEditor?: VisualTexConfigurationWindowSize | null;
}

export interface VisualTexConfigurationPersonalization {
  usage?: Record<string, CommandUsage>;
  history?: FormulaHistoryItem[];
}

export interface VisualTexConfigurationWordPreferences {
  defaultDisplayEquationNumbered?: boolean;
  defaultEquationNumberFormat?: string;
}

export interface VisualTexUserConfiguration {
  schema: typeof VISUALTEX_CONFIGURATION_SCHEMA;
  version: typeof VISUALTEX_CONFIGURATION_VERSION;
  exportedAt: string;
  editorSettings: Partial<FormulaDocument["settings"]>;
  storage: Record<string, string>;
  capturedStorageKeys?: string[];
  personalization?: VisualTexConfigurationPersonalization;
  word?: VisualTexConfigurationWordPreferences;
  windows?: VisualTexConfigurationWindowState;
}

const configurationStorageKeys = [
  "visualtex-custom-formula-tiles",
  "visualtex-common-toolbar-command-ids-v1",
  "visualtex-common-toolbar-command-ids-v2",
  "visualtex-formula-hotkeys-v1",
  "visualtex-custom-formula-text-colors",
  "visualtex-custom-formula-background-colors",
  "visualtex-desktop-editor-toolbar-open",
  "visualtex-desktop-editor-tiles-open",
  "visualtex-desktop-editor-source-open",
  "visualtex-office-editor-toolbar-open",
  "visualtex-office-editor-tiles-open",
  "visualtex-office-editor-source-open",
  "visualtex.ocr.model",
  "visualtex.silent-ocr.enabled",
  "visualtex.quick-ocr.capture-mode",
  "visualtex.custom-theme.v1",
  "visualtex.formula-letter-font",
  "visualtex.formula-chinese-font",
  "visualtex-classic-tile-width",
  "visualtex-classic-dock-height",
  CUSTOM_SYMBOL_STORAGE_KEY,
] as const;

const booleanConfigurationStorageKeys = new Set<string>([
  "visualtex-desktop-editor-toolbar-open",
  "visualtex-desktop-editor-tiles-open",
  "visualtex-desktop-editor-source-open",
  "visualtex-office-editor-toolbar-open",
  "visualtex-office-editor-tiles-open",
  "visualtex-office-editor-source-open",
  "visualtex.silent-ocr.enabled",
]);

const jsonConfigurationStorageKeys = new Set<string>([
  "visualtex-custom-formula-tiles",
  "visualtex-common-toolbar-command-ids-v1",
  "visualtex-common-toolbar-command-ids-v2",
  "visualtex-formula-hotkeys-v1",
  "visualtex-custom-formula-text-colors",
  "visualtex-custom-formula-background-colors",
  "visualtex.custom-theme.v1",
  CUSTOM_SYMBOL_STORAGE_KEY,
]);

const editorSettingKeys = [
  "theme",
  "zoom",
  "formulaAlignment",
  "latexCodeFormat",
  "editorLayout",
  "language",
  "sourceOpen",
  "autoPairDelimiters",
  "showLineNumbers",
  "highlightActiveLine",
  "formulaInsetLeft",
  "formulaInsetRight",
  "formulaToolButtonSize",
  "formulaToolButtonPadding",
  "formulaRowVerticalInset",
  "pngExportBackground",
  "formulaLetterFont",
  "formulaChineseFont",
  "inputBehavior",
  "personalize",
  "suggestionCount",
  "checkUpdatesOnStartup",
  "powerPointDefaultFontSizePt",
  "classicTileWidth",
  "classicDockHeight",
  "keypadMinimizeOnCopy",
] as const satisfies readonly (keyof FormulaDocument["settings"])[];

const maximumStorageEntryLength = 2_000_000;
const maximumCustomSymbolConfigurationLength = 64_000_000;
const maximumConfigurationLength = 96_000_000;
const maximumUsageCommands = 10_000;
const maximumUsageMapEntries = 256;
const maximumHistoryItems = 30;
const maximumHistoryLatexLength = 1_000_000;
const validWordNumberFormats = new Set([
  "continuous",
  "heading1-dot",
  "heading1-dash",
  "heading2-dot",
  "heading2-dash",
]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function normalizeEditorSettings(value: unknown) {
  if (!isRecord(value)) return {};
  const settings: Partial<FormulaDocument["settings"]> = {};
  for (const key of editorSettingKeys) {
    if (Object.prototype.hasOwnProperty.call(value, key)) {
      (settings as Record<string, unknown>)[key] = value[key];
    }
  }
  return settings;
}

function normalizeWindowSize(value: unknown): VisualTexConfigurationWindowSize | null {
  if (!isRecord(value)) return null;
  const width = Number(value.width);
  const height = Number(value.height);
  if (
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    width < 300 ||
    height < 200 ||
    width > 5000 ||
    height > 4000
  ) {
    return null;
  }
  return { width, height };
}

function normalizeWindowState(value: unknown): VisualTexConfigurationWindowState | undefined {
  if (!isRecord(value)) return undefined;
  const main = normalizeWindowSize(value.main);
  const keypad = normalizeWindowSize(value.keypad);
  const officeEditor = normalizeWindowSize(value.officeEditor);
  if (!main && !keypad && !officeEditor) return undefined;
  return { main, keypad, officeEditor };
}

function normalizeCapturedStorageKeys(value: unknown) {
  if (!Array.isArray(value)) return undefined;
  const allowed = new Set<string>(configurationStorageKeys);
  return Array.from(
    new Set(value.filter((key): key is string => typeof key === "string" && allowed.has(key))),
  );
}

function normalizeCountMap(value: unknown) {
  const result: Record<string, number> = {};
  if (!isRecord(value)) return result;
  for (const [key, raw] of Object.entries(value).slice(0, maximumUsageMapEntries)) {
    const count = Number(raw);
    if (!key || !Number.isFinite(count) || count < 0) continue;
    result[key] = Math.min(Number.MAX_SAFE_INTEGER, Math.floor(count));
  }
  return result;
}

function normalizeUsage(value: unknown): Record<string, CommandUsage> | undefined {
  if (!isRecord(value)) return undefined;
  const result: Record<string, CommandUsage> = {};
  for (const [key, raw] of Object.entries(value).slice(0, maximumUsageCommands)) {
    if (!isRecord(raw)) continue;
    const commandId =
      typeof raw.commandId === "string" && raw.commandId.trim()
        ? raw.commandId.trim()
        : key.trim();
    if (!commandId) continue;
    const useCount = Number(raw.useCount);
    const lastUsedAt = Number(raw.lastUsedAt);
    const recentUses = Array.isArray(raw.recentUses)
      ? raw.recentUses
          .map(Number)
          .filter((item) => Number.isFinite(item) && item >= 0)
          .slice(-12)
      : [];
    result[commandId] = {
      commandId,
      useCount:
        Number.isFinite(useCount) && useCount >= 0
          ? Math.min(Number.MAX_SAFE_INTEGER, Math.floor(useCount))
          : 0,
      lastUsedAt:
        Number.isFinite(lastUsedAt) && lastUsedAt >= 0 ? lastUsedAt : 0,
      recentUses,
      acceptedPrefixes: normalizeCountMap(raw.acceptedPrefixes),
      contextCounts: normalizeCountMap(raw.contextCounts),
      pinned: raw.pinned === true,
    };
  }
  return result;
}

function normalizeHistory(value: unknown): FormulaHistoryItem[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const result: FormulaHistoryItem[] = [];
  for (let index = 0; index < value.length && result.length < maximumHistoryItems; index += 1) {
    const raw = value[index];
    if (!isRecord(raw) || typeof raw.latex !== "string") continue;
    const latex = raw.latex.slice(0, maximumHistoryLatexLength);
    if (!latex.trim()) continue;
    const createdAt = Number(raw.createdAt);
    result.push({
      id:
        typeof raw.id === "string" && raw.id.trim()
          ? raw.id.trim()
          : `imported-history-${index}`,
      latex,
      createdAt:
        Number.isFinite(createdAt) && createdAt >= 0 ? createdAt : 0,
    });
  }
  return result;
}

function normalizePersonalization(value: unknown): VisualTexConfigurationPersonalization | undefined {
  if (!isRecord(value)) return undefined;
  const usage = normalizeUsage(value.usage);
  const history = normalizeHistory(value.history);
  if (usage === undefined && history === undefined) return undefined;
  return { usage, history };
}

function normalizeWordPreferences(value: unknown): VisualTexConfigurationWordPreferences | undefined {
  if (!isRecord(value)) return undefined;
  const result: VisualTexConfigurationWordPreferences = {};
  if (typeof value.defaultDisplayEquationNumbered === "boolean") {
    result.defaultDisplayEquationNumbered = value.defaultDisplayEquationNumbered;
  }
  if (
    typeof value.defaultEquationNumberFormat === "string" &&
    validWordNumberFormats.has(value.defaultEquationNumberFormat)
  ) {
    result.defaultEquationNumberFormat = value.defaultEquationNumberFormat;
  }
  return Object.keys(result).length ? result : undefined;
}

function normalizeStorage(value: unknown) {
  const result: Record<string, string> = {};
  if (!isRecord(value)) return result;
  for (const key of configurationStorageKeys) {
    const raw = value[key];
    const maximumLength =
      key === CUSTOM_SYMBOL_STORAGE_KEY
        ? maximumCustomSymbolConfigurationLength
        : maximumStorageEntryLength;
    if (typeof raw !== "string" || raw.length > maximumLength) continue;
    if (jsonConfigurationStorageKeys.has(key)) {
      try {
        JSON.parse(raw);
      } catch {
        continue;
      }
    } else if (
      booleanConfigurationStorageKeys.has(key) &&
      !["true", "false", "1", "0"].includes(raw)
    ) {
      continue;
    } else if (
      key === "visualtex.ocr.model" &&
      !OCR_MODELS.some((item) => item.id === raw)
    ) {
      continue;
    } else if (key === "visualtex.quick-ocr.capture-mode") {
      const captureMode = normalizeQuickOcrCaptureMode(raw);
      if (!captureMode) continue;
      result[key] = captureMode;
      continue;
    }
    result[key] = raw;
  }
  return result;
}

export function parseVisualTexConfiguration(source: string): VisualTexUserConfiguration {
  if (!source.trim() || source.length > maximumConfigurationLength) {
    throw new Error("The configuration file is empty or too large.");
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(source);
  } catch {
    throw new Error("The configuration file is not valid JSON.");
  }
  if (!isRecord(parsed)) {
    throw new Error("The configuration file has an invalid structure.");
  }
  if (parsed.schema !== VISUALTEX_CONFIGURATION_SCHEMA) {
    throw new Error("This file is not a VisualTeX configuration file.");
  }
  const sourceVersion = parsed.version === undefined ? 1 : Number(parsed.version);
  if (!Number.isInteger(sourceVersion) || sourceVersion < 1) {
    throw new Error("This VisualTeX configuration version is not supported.");
  }
  return {
    schema: VISUALTEX_CONFIGURATION_SCHEMA,
    version: VISUALTEX_CONFIGURATION_VERSION,
    exportedAt:
      typeof parsed.exportedAt === "string" ? parsed.exportedAt : new Date(0).toISOString(),
    editorSettings: normalizeEditorSettings(parsed.editorSettings),
    storage: normalizeStorage(parsed.storage),
    capturedStorageKeys: normalizeCapturedStorageKeys(parsed.capturedStorageKeys),
    personalization: normalizePersonalization(parsed.personalization),
    word: normalizeWordPreferences(parsed.word),
    windows: normalizeWindowState(parsed.windows),
  };
}

async function readWindowConfiguration(): Promise<VisualTexConfigurationWindowState | undefined> {
  if (!isTauri()) {
    return {
      main: {
        width: window.innerWidth,
        height: window.innerHeight,
      },
    };
  }
  try {
    return normalizeWindowState(
      await invoke<unknown>("get_app_window_configuration"),
    );
  } catch {
    return undefined;
  }
}

async function readWordPreferences(): Promise<VisualTexConfigurationWordPreferences | undefined> {
  if (!isTauri()) return undefined;
  try {
    return normalizeWordPreferences(
      await invoke<VisualTexConfigurationWordPreferences>(
        "get_word_numbering_user_configuration",
      ),
    );
  } catch {
    return undefined;
  }
}

async function applyWordPreferences(
  word: VisualTexConfigurationWordPreferences | undefined,
) {
  if (!word || !isTauri()) return;
  const current = await readWordPreferences();
  const merged: VisualTexConfigurationWordPreferences = {
    defaultDisplayEquationNumbered:
      word.defaultDisplayEquationNumbered ??
      current?.defaultDisplayEquationNumbered ??
      false,
    defaultEquationNumberFormat:
      word.defaultEquationNumberFormat ??
      current?.defaultEquationNumberFormat ??
      "continuous",
  };
  try {
    await invoke("apply_word_numbering_user_configuration", {
      configuration: merged,
    });
  } catch {
    // Older runtimes do not know this command. Cross-version imports should
    // still restore every preference that the running version understands.
  }
}

async function applyWindowConfiguration(
  windows: VisualTexConfigurationWindowState | undefined,
) {
  if (!windows || !isTauri()) return;
  try {
    await invoke("apply_app_window_configuration", {
      configuration: windows,
    });
  } catch {
    // Window geometry backup must not make importing the remaining settings fail.
  }
}

export async function buildVisualTexConfiguration(): Promise<VisualTexUserConfiguration> {
  const editorState = useEditorStore.getState();
  const editorSettings = editorState.toDocument().settings;
  const storage: Record<string, string> = {};
  for (const key of configurationStorageKeys) {
    if (key === CUSTOM_SYMBOL_STORAGE_KEY) {
      storage[key] = JSON.stringify(readCustomSymbolLibrary());
      continue;
    }
    const value = safeStorage.getItem(key);
    if (value !== null) storage[key] = value;
  }
  return {
    schema: VISUALTEX_CONFIGURATION_SCHEMA,
    version: VISUALTEX_CONFIGURATION_VERSION,
    exportedAt: new Date().toISOString(),
    editorSettings: { ...editorSettings },
    storage,
    capturedStorageKeys: [...configurationStorageKeys],
    personalization: {
      usage: normalizeUsage(editorState.usage) ?? {},
      history: normalizeHistory(editorState.history) ?? [],
    },
    word: await readWordPreferences(),
    windows: await readWindowConfiguration(),
  };
}

export async function applyVisualTexConfiguration(
  input: VisualTexUserConfiguration,
): Promise<void> {
  const configuration = parseVisualTexConfiguration(JSON.stringify(input));
  const editor = useEditorStore.getState();
  const currentDocument = editor.toDocument();
  const mergedSettings = {
    ...currentDocument.settings,
    ...configuration.editorSettings,
  };
  editor.loadDocument({
    ...currentDocument,
    settings: mergedSettings,
  });

  const capturedStorageKeys = configuration.capturedStorageKeys
    ? new Set(configuration.capturedStorageKeys)
    : null;
  for (const key of configurationStorageKeys) {
    if (key === CUSTOM_SYMBOL_STORAGE_KEY) continue;
    const value = configuration.storage[key];
    if (typeof value === "string") {
      safeStorage.setItem(key, value);
    } else if (capturedStorageKeys?.has(key)) {
      safeStorage.removeItem(key);
    }
  }
  const customSymbols = configuration.storage[CUSTOM_SYMBOL_STORAGE_KEY];
  if (typeof customSymbols === "string") {
    replaceCustomSymbolLibrary(JSON.parse(customSymbols));
  } else if (capturedStorageKeys?.has(CUSTOM_SYMBOL_STORAGE_KEY)) {
    replaceCustomSymbolLibrary({ version: 1, symbols: [] });
  }

  if (configuration.personalization) {
    const personalizationUpdate: {
      usage?: Record<string, CommandUsage>;
      history?: FormulaHistoryItem[];
    } = {};
    if (configuration.personalization.usage !== undefined) {
      personalizationUpdate.usage = configuration.personalization.usage;
    }
    if (configuration.personalization.history !== undefined) {
      personalizationUpdate.history = configuration.personalization.history;
    }
    if (Object.keys(personalizationUpdate).length) {
      useEditorStore.setState(personalizationUpdate);
    }
  }

  publishSynchronizedTheme(useEditorStore.getState().theme);
  if (isTauri()) {
    await invoke("set_app_theme", {
      theme: useEditorStore.getState().theme,
    }).catch(() => undefined);
  }
  await applyWordPreferences(configuration.word);
  await applyWindowConfiguration(configuration.windows);
}
