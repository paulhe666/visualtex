import { invoke, isTauri } from "@tauri-apps/api/core";
import { useEditorStore } from "../stores/editorStore";
import { safeStorage } from "./safeStorage";
import { publishSynchronizedTheme } from "../themeSync";
import { OCR_MODELS } from "../ocr/ocrService";
import type { FormulaDocument } from "../types/formula";
import type { CommandUsage } from "../types/command";
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
  officeEditor?: VisualTexConfigurationWindowSize | null;
}

export interface VisualTexUserConfiguration {
  schema: typeof VISUALTEX_CONFIGURATION_SCHEMA;
  version: typeof VISUALTEX_CONFIGURATION_VERSION;
  exportedAt: string;
  editorSettings: Partial<FormulaDocument["settings"]>;
  storage: Record<string, string>;
  usage?: Record<string, CommandUsage>;
  windows?: VisualTexConfigurationWindowState;
}

const configurationStorageKeys = [
  "visualtex-custom-formula-tiles",
  "visualtex-common-toolbar-command-ids-v1",
  "visualtex-formula-hotkeys-v1",
  "visualtex-custom-formula-text-colors",
  "visualtex-custom-formula-background-colors",
  "visualtex-desktop-editor-toolbar-open",
  "visualtex-desktop-editor-tiles-open",
  "visualtex-office-editor-toolbar-open",
  "visualtex-office-editor-tiles-open",
  "visualtex.ocr.model",
  "visualtex.silent-ocr.enabled",
  "visualtex.quick-ocr.capture-mode",
  "visualtex.custom-theme.v1",
  CUSTOM_SYMBOL_STORAGE_KEY,
] as const;

const booleanConfigurationStorageKeys = new Set<string>([
  "visualtex-desktop-editor-toolbar-open",
  "visualtex-desktop-editor-tiles-open",
  "visualtex-office-editor-toolbar-open",
  "visualtex-office-editor-tiles-open",
  "visualtex.silent-ocr.enabled",
]);

const jsonConfigurationStorageKeys = new Set<string>([
  "visualtex-custom-formula-tiles",
  "visualtex-common-toolbar-command-ids-v1",
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
const maximumUsageEntries = 4096;
const maximumUsageMapEntries = 512;
const maximumUsageKeyLength = 256;

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function normalizeEditorSettings(value: unknown) {
  if (!isRecord(value)) {
    throw new Error("The configuration is missing editor settings.");
  }
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
  const officeEditor = normalizeWindowSize(value.officeEditor);
  if (!main && !officeEditor) return undefined;
  return { main, officeEditor };
}

function normalizeUsageCounter(value: unknown) {
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) return 0;
  return Math.min(Number.MAX_SAFE_INTEGER, Math.floor(value));
}

function normalizeUsageCounterMap(value: unknown) {
  const result: Record<string, number> = {};
  if (!isRecord(value)) return result;
  for (const [key, rawCount] of Object.entries(value).slice(0, maximumUsageMapEntries)) {
    if (!key || key.length > maximumUsageKeyLength) continue;
    const count = normalizeUsageCounter(rawCount);
    if (count > 0) result[key] = count;
  }
  return result;
}

function normalizeCommandUsage(value: unknown): Record<string, CommandUsage> | undefined {
  if (value === undefined) return undefined;
  if (!isRecord(value)) return {};
  const result: Record<string, CommandUsage> = {};
  for (const [storageId, rawUsage] of Object.entries(value).slice(0, maximumUsageEntries)) {
    if (!storageId || storageId.length > maximumUsageKeyLength || !isRecord(rawUsage)) {
      continue;
    }
    const commandId =
      typeof rawUsage.commandId === "string" &&
      rawUsage.commandId.length > 0 &&
      rawUsage.commandId.length <= maximumUsageKeyLength
        ? rawUsage.commandId
        : storageId;
    const recentUses = Array.isArray(rawUsage.recentUses)
      ? rawUsage.recentUses
          .filter(
            (item): item is number =>
              typeof item === "number" && Number.isFinite(item) && item >= 0,
          )
          .map((item) => Math.floor(item))
          .slice(-12)
      : [];
    result[storageId] = {
      commandId,
      useCount: normalizeUsageCounter(rawUsage.useCount),
      lastUsedAt: normalizeUsageCounter(rawUsage.lastUsedAt),
      recentUses,
      acceptedPrefixes: normalizeUsageCounterMap(rawUsage.acceptedPrefixes),
      contextCounts: normalizeUsageCounterMap(rawUsage.contextCounts),
      pinned: rawUsage.pinned === true,
    };
  }
  return result;
}

function cloneCommandUsage(usage: Record<string, CommandUsage>) {
  return Object.fromEntries(
    Object.entries(usage).map(([commandId, item]) => [
      commandId,
      {
        ...item,
        recentUses: [...item.recentUses],
        acceptedPrefixes: { ...item.acceptedPrefixes },
        contextCounts: { ...item.contextCounts },
      },
    ]),
  );
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
    } else if (
      key === "visualtex.quick-ocr.capture-mode" &&
      !["immediate", "system-screenshot"].includes(raw)
    ) {
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
  if (parsed.version !== VISUALTEX_CONFIGURATION_VERSION) {
    throw new Error("This VisualTeX configuration version is not supported.");
  }
  return {
    schema: VISUALTEX_CONFIGURATION_SCHEMA,
    version: VISUALTEX_CONFIGURATION_VERSION,
    exportedAt:
      typeof parsed.exportedAt === "string" ? parsed.exportedAt : new Date(0).toISOString(),
    editorSettings: normalizeEditorSettings(parsed.editorSettings),
    storage: normalizeStorage(parsed.storage),
    usage: normalizeCommandUsage(parsed.usage),
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
    usage: cloneCommandUsage(editorState.usage),
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
  if (configuration.usage !== undefined) {
    useEditorStore.setState({ usage: cloneCommandUsage(configuration.usage) });
  }

  for (const key of configurationStorageKeys) {
    if (key === CUSTOM_SYMBOL_STORAGE_KEY) continue;
    const value = configuration.storage[key];
    if (typeof value === "string") safeStorage.setItem(key, value);
    else safeStorage.removeItem(key);
  }
  const customSymbols = configuration.storage[CUSTOM_SYMBOL_STORAGE_KEY];
  replaceCustomSymbolLibrary(
    customSymbols ? JSON.parse(customSymbols) : { version: 1, symbols: [] },
  );

  publishSynchronizedTheme(useEditorStore.getState().theme);
  if (isTauri()) {
    await invoke("set_app_theme", {
      theme: useEditorStore.getState().theme,
    }).catch(() => undefined);
    if (configuration.windows) {
      await invoke("apply_app_window_configuration", {
        configuration: configuration.windows,
      });
    }
  }
}
