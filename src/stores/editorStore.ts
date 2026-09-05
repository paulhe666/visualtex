import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import type { CommandSource, CommandUsage } from "../types/command";
import type {
  FormulaDocument,
  FormulaAlignment,
  FormulaHistoryItem,
  FormulaLine,
  FormulaLineMode,
  InputBehaviorSettingKey,
  InputBehaviorSettings,
  LatexCodeFormat,
  Theme,
} from "../types/formula";
import type { DocumentSnapshot } from "../history/historyTypes";
import {
  DEFAULT_LATEX_CODE_FORMAT,
  isLatexCodeFormat,
} from "../clipboard/LatexCopyService";
import { normalizeChineseLatex } from "../editor/normalizeChineseLatex";
import { normalizeMultilineLatex } from "../editor/normalizeChineseLatex";
import { normalizeFormulaLinePhysicalWhitespace } from "../math/formulaLineLatex";
import { isSingleCompleteLatexEnvironment } from "../math/latexEnvironment";
import { createUuid } from "../runtime/browserCompatibility";
import { safeStorage } from "../runtime/safeStorage";
import { isLandingPreview, LANDING_PREVIEW_ZOOM } from "../runtime/landingPreview";
import {
  DEFAULT_PNG_EXPORT_BACKGROUND,
  normalizePngExportBackground,
  type PngExportBackground,
} from "../export/pngBackground";
import { isTheme } from "../themeCustomization";
import {
  DEFAULT_FORMULA_CHINESE_FONT,
  DEFAULT_FORMULA_LETTER_FONT,
  normalizeFormulaChineseFont,
  normalizeFormulaLetterFont,
  persistFormulaChineseFontPreference,
  persistFormulaFontPreferences,
  persistFormulaLetterFontPreference,
  readPersistedFormulaFontPreferences,
  type FormulaChineseFont,
  type FormulaLetterFont,
} from "../editor/formulaFontPreferences";

export type Language = "cn" | "en";
export type EditorLayout = "standard" | "classic";
export const DEFAULT_EDITOR_LAYOUT: EditorLayout = "classic";
export const DEFAULT_THEME: Theme = "light";
export const MIN_EDITOR_ZOOM = 0.2;
export const MAX_EDITOR_ZOOM = 1.6;
export const EDITOR_ZOOM_STEP = 0.05;
export const DEFAULT_EDITOR_ZOOM = 0.45;
export const DEFAULT_FORMULA_INSET = 34;
export const MIN_FORMULA_INSET = 0;
export const MAX_FORMULA_INSET = 96;
export const DEFAULT_FORMULA_TOOL_BUTTON_SIZE = 52;
export const MIN_FORMULA_TOOL_BUTTON_SIZE = 38;
export const MAX_FORMULA_TOOL_BUTTON_SIZE = 72;
export const DEFAULT_FORMULA_TOOL_BUTTON_PADDING = 2;
export const MIN_FORMULA_TOOL_BUTTON_PADDING = 0;
export const MAX_FORMULA_TOOL_BUTTON_PADDING = 12;
export const DEFAULT_FORMULA_ROW_VERTICAL_INSET = 5;
export const MIN_FORMULA_ROW_VERTICAL_INSET = 0;
export const MAX_FORMULA_ROW_VERTICAL_INSET = 24;
export const DEFAULT_CLASSIC_TILE_WIDTH = 300;
export const MIN_CLASSIC_TILE_WIDTH = 220;
export const MAX_CLASSIC_TILE_WIDTH = 2000;
export const DEFAULT_CLASSIC_DOCK_HEIGHT = 240;
export const MIN_CLASSIC_DOCK_HEIGHT = 132;
export const MAX_CLASSIC_DOCK_HEIGHT = 2000;

const legacyClassicTileWidthStorageKey = "visualtex-classic-tile-width";
const legacyClassicDockHeightStorageKey = "visualtex-classic-dock-height";

export const DEFAULT_INPUT_BEHAVIOR_SETTINGS: InputBehaviorSettings = {
  autoEscapeShortcuts: true,
  autoExitSuperscript: true,
  autoExitSubscript: true,
  autoExitAccent: true,
  autoExitWrapperCommand: true,
  showStructuredCommandSuggestions: true,
  showOtherCommandSuggestions: false,
};

function normalizeInputBehaviorSettings(
  value: unknown,
): InputBehaviorSettings {
  const candidate =
    value && typeof value === "object"
      ? (value as Partial<InputBehaviorSettings>)
      : {};
  return {
    autoEscapeShortcuts:
      typeof candidate.autoEscapeShortcuts === "boolean"
        ? candidate.autoEscapeShortcuts
        : true,
    autoExitSuperscript:
      typeof candidate.autoExitSuperscript === "boolean"
        ? candidate.autoExitSuperscript
        : true,
    autoExitSubscript:
      typeof candidate.autoExitSubscript === "boolean"
        ? candidate.autoExitSubscript
        : true,
    autoExitAccent:
      typeof candidate.autoExitAccent === "boolean"
        ? candidate.autoExitAccent
        : true,
    autoExitWrapperCommand:
      typeof candidate.autoExitWrapperCommand === "boolean"
        ? candidate.autoExitWrapperCommand
        : true,
    showStructuredCommandSuggestions:
      typeof candidate.showStructuredCommandSuggestions === "boolean"
        ? candidate.showStructuredCommandSuggestions
        : true,
    showOtherCommandSuggestions:
      typeof candidate.showOtherCommandSuggestions === "boolean"
        ? candidate.showOtherCommandSuggestions
        : false,
  };
}

function normalizeFormulaAlignment(value: unknown): FormulaAlignment {
  return value === "center" || value === "right" ? value : "left";
}

export function normalizeEditorLayout(value: unknown): EditorLayout {
  return value === "standard" ? "standard" : DEFAULT_EDITOR_LAYOUT;
}

function normalizeTheme(value: unknown): Theme {
  return isTheme(value) ? value : DEFAULT_THEME;
}

function normalizeEditorZoom(value: unknown) {
  const zoom =
    typeof value === "number" && Number.isFinite(value)
      ? value
      : DEFAULT_EDITOR_ZOOM;
  const steppedZoom =
    Math.round(
      Math.round(zoom / EDITOR_ZOOM_STEP) * EDITOR_ZOOM_STEP * 100,
    ) / 100;
  return Math.min(MAX_EDITOR_ZOOM, Math.max(MIN_EDITOR_ZOOM, steppedZoom));
}

function normalizeFormulaInset(value: unknown) {
  const inset =
    typeof value === "number" && Number.isFinite(value)
      ? Math.round(value)
      : DEFAULT_FORMULA_INSET;
  return Math.min(MAX_FORMULA_INSET, Math.max(MIN_FORMULA_INSET, inset));
}

function normalizeFormulaToolButtonSize(value: unknown) {
  const size =
    typeof value === "number" && Number.isFinite(value)
      ? Math.round(value)
      : DEFAULT_FORMULA_TOOL_BUTTON_SIZE;
  return Math.min(
    MAX_FORMULA_TOOL_BUTTON_SIZE,
    Math.max(MIN_FORMULA_TOOL_BUTTON_SIZE, size),
  );
}

function normalizeFormulaToolButtonPadding(value: unknown) {
  const padding =
    typeof value === "number" && Number.isFinite(value)
      ? Math.round(value)
      : DEFAULT_FORMULA_TOOL_BUTTON_PADDING;
  return Math.min(
    MAX_FORMULA_TOOL_BUTTON_PADDING,
    Math.max(MIN_FORMULA_TOOL_BUTTON_PADDING, padding),
  );
}

function normalizeFormulaRowVerticalInset(value: unknown) {
  const inset =
    typeof value === "number" && Number.isFinite(value)
      ? Math.round(value)
      : DEFAULT_FORMULA_ROW_VERTICAL_INSET;
  return Math.min(
    MAX_FORMULA_ROW_VERTICAL_INSET,
    Math.max(MIN_FORMULA_ROW_VERTICAL_INSET, inset),
  );
}

function normalizeClassicPanelSize(
  value: unknown,
  fallback: number,
  minimum: number,
  maximum: number,
) {
  const size =
    typeof value === "number" && Number.isFinite(value)
      ? Math.round(value)
      : fallback;
  return Math.min(maximum, Math.max(minimum, size));
}

function legacyClassicPanelSize(
  storageKey: string,
  fallback: number,
  minimum: number,
  maximum: number,
) {
  const stored = Number.parseFloat(safeStorage.getItem(storageKey) ?? "");
  return normalizeClassicPanelSize(
    Number.isFinite(stored) ? stored : fallback,
    fallback,
    minimum,
    maximum,
  );
}

function normalizeClassicTileWidth(value: unknown) {
  return normalizeClassicPanelSize(
    value,
    DEFAULT_CLASSIC_TILE_WIDTH,
    MIN_CLASSIC_TILE_WIDTH,
    MAX_CLASSIC_TILE_WIDTH,
  );
}

function normalizeClassicDockHeight(value: unknown) {
  return normalizeClassicPanelSize(
    value,
    DEFAULT_CLASSIC_DOCK_HEIGHT,
    MIN_CLASSIC_DOCK_HEIGHT,
    MAX_CLASSIC_DOCK_HEIGHT,
  );
}

function normalizeFormulaLineLatex(latex: string) {
  const normalized = latex.replace(/\r\n?/g, "\n");
  const trimmed = normalized.trim();
  return normalizeChineseLatex(
    isSingleCompleteLatexEnvironment(trimmed)
      ? trimmed
      : normalizeFormulaLinePhysicalWhitespace(normalized),
  );
}

export function createFormulaLine(
  latex = "",
  id: string = createUuid(),
  mode: FormulaLineMode = "display",
): FormulaLine {
  return {
    id,
    latex: normalizeFormulaLineLatex(latex.replace(/\r\n?/g, "\n")),
    mode: mode === "inline" ? "inline" : "display",
  };
}

function uniqueLineId(candidate: unknown, usedIds: Set<string>) {
  const normalized = typeof candidate === "string" ? candidate.trim() : "";
  if (normalized && !usedIds.has(normalized)) {
    usedIds.add(normalized);
    return normalized;
  }
  let nextId: string = createUuid();
  while (usedIds.has(nextId)) nextId = createUuid();
  usedIds.add(nextId);
  return nextId;
}

export function normalizeFormulaLines(
  value: unknown,
  legacyLatex?: unknown,
): FormulaLine[] {
  const usedIds = new Set<string>();
  if (Array.isArray(value) && value.length) {
    const normalized = value
      .map((item) => {
        if (!item || typeof item !== "object") return null;
        const candidate = item as Partial<FormulaLine>;
        return {
          id: uniqueLineId(candidate.id, usedIds),
          latex: normalizeFormulaLineLatex(
            typeof candidate.latex === "string"
              ? candidate.latex.replace(/\r\n?/g, "\n")
              : "",
          ),
          mode: candidate.mode === "inline" ? "inline" : "display",
        } satisfies FormulaLine;
      })
      .filter((line): line is NonNullable<typeof line> => line !== null);
    if (normalized.length) return normalized;
  }

  const normalizedLatex = normalizeMultilineLatex(
    typeof legacyLatex === "string" ? legacyLatex : "",
  );
  const values = normalizedLatex.split("\n");
  return (values.length ? values : [""]).map((latex) =>
    createFormulaLine(latex, uniqueLineId(undefined, usedIds)),
  );
}

export function joinFormulaLines(lines: readonly FormulaLine[]): string {
  return lines.map((line) => line.latex).join("\n");
}

export function cloneFormulaLines(lines: readonly FormulaLine[]): FormulaLine[] {
  return lines.map((line) => ({ ...line }));
}

function validActiveLineId(
  lines: readonly FormulaLine[],
  candidate: unknown,
): string | null {
  if (
    typeof candidate === "string" &&
    lines.some((line) => line.id === candidate)
  ) {
    return candidate;
  }
  return lines[0]?.id ?? null;
}

interface EditorState {
  title: string;
  lines: FormulaLine[];
  activeLineId: string | null;
  formulaAlignment: FormulaAlignment;
  editorLayout: EditorLayout;
  theme: Theme;
  language: Language;
  zoom: number;
  sourceOpen: boolean;
  latexCodeFormat: LatexCodeFormat;
  autoPairDelimiters: boolean;
  showLineNumbers: boolean;
  highlightActiveLine: boolean;
  formulaInsetLeft: number;
  formulaInsetRight: number;
  formulaToolButtonSize: number;
  formulaToolButtonPadding: number;
  formulaRowVerticalInset: number;
  pngExportBackground: PngExportBackground;
  formulaLetterFont: FormulaLetterFont;
  formulaChineseFont: FormulaChineseFont;
  classicTileWidth: number;
  classicDockHeight: number;
  keypadMinimizeOnCopy: boolean;
  inputBehavior: InputBehaviorSettings;
  personalize: boolean;
  suggestionCount: number;
  checkUpdatesOnStartup: boolean;
  powerPointDefaultFontSizePt: number;
  usage: Record<string, CommandUsage>;
  history: FormulaHistoryItem[];
  setTitle: (title: string) => void;
  setActiveLineId: (lineId: string | null) => void;
  replaceFormulaLine: (lineId: string, latex: string) => void;
  setFormulaAlignment: (alignment: FormulaAlignment) => void;
  setEditorLayout: (layout: EditorLayout) => void;
  insertFormulaLine: (line: FormulaLine, index: number) => void;
  removeFormulaLine: (lineId: string) => void;
  replaceDocumentState: (snapshot: DocumentSnapshot) => void;
  setTheme: (theme: Theme) => void;
  setLanguage: (language: Language) => void;
  setZoom: (zoom: number) => void;
  setSourceOpen: (open: boolean) => void;
  setLatexCodeFormat: (format: LatexCodeFormat) => void;
  setAutoPairDelimiters: (enabled: boolean) => void;
  setShowLineNumbers: (enabled: boolean) => void;
  setHighlightActiveLine: (enabled: boolean) => void;
  setFormulaInsetLeft: (inset: number) => void;
  setFormulaInsetRight: (inset: number) => void;
  setFormulaToolButtonSize: (size: number) => void;
  setFormulaToolButtonPadding: (padding: number) => void;
  setFormulaRowVerticalInset: (inset: number) => void;
  setPngExportBackground: (background: PngExportBackground) => void;
  setFormulaLetterFont: (font: FormulaLetterFont) => void;
  setFormulaChineseFont: (font: FormulaChineseFont) => void;
  setClassicTileWidth: (width: number) => void;
  setClassicDockHeight: (height: number) => void;
  setKeypadMinimizeOnCopy: (enabled: boolean) => void;
  setInputBehavior: (
    setting: InputBehaviorSettingKey,
    enabled: boolean,
  ) => void;
  setPersonalize: (enabled: boolean) => void;
  setSuggestionCount: (count: number) => void;
  setCheckUpdatesOnStartup: (enabled: boolean) => void;
  setPowerPointDefaultFontSizePt: (fontSizePt: number) => void;
  recordCommand: (commandId: string, prefix: string, source: CommandSource) => void;
  resetUsage: () => void;
  addHistory: (latex?: string) => void;
  clearHistory: () => void;
  loadDocument: (document: FormulaDocument) => void;
  toDocument: () => FormulaDocument;
}

const initialLatex = "\\int_{-\\infty}^{\\infty} e^{-x^2}\\,\\mathrm{d}x = \\sqrt{\\pi}";
const initialLines = [createFormulaLine(initialLatex)];

export const useEditorStore = create<EditorState>()(
  persist(
    (set, get) => ({
      title: "未命名公式",
      lines: initialLines,
      activeLineId: initialLines[0].id,
      formulaAlignment: "left",
      editorLayout: DEFAULT_EDITOR_LAYOUT,
      theme: DEFAULT_THEME,
      language: "cn",
      zoom: isLandingPreview ? LANDING_PREVIEW_ZOOM : DEFAULT_EDITOR_ZOOM,
      sourceOpen: false,
      latexCodeFormat: DEFAULT_LATEX_CODE_FORMAT,
      autoPairDelimiters: true,
      showLineNumbers: false,
      highlightActiveLine: false,
      formulaInsetLeft: DEFAULT_FORMULA_INSET,
      formulaInsetRight: DEFAULT_FORMULA_INSET,
      formulaToolButtonSize: DEFAULT_FORMULA_TOOL_BUTTON_SIZE,
      formulaToolButtonPadding: DEFAULT_FORMULA_TOOL_BUTTON_PADDING,
      formulaRowVerticalInset: DEFAULT_FORMULA_ROW_VERTICAL_INSET,
      pngExportBackground: DEFAULT_PNG_EXPORT_BACKGROUND,
      formulaLetterFont: DEFAULT_FORMULA_LETTER_FONT,
      formulaChineseFont: DEFAULT_FORMULA_CHINESE_FONT,
      classicTileWidth: legacyClassicPanelSize(
        legacyClassicTileWidthStorageKey,
        DEFAULT_CLASSIC_TILE_WIDTH,
        MIN_CLASSIC_TILE_WIDTH,
        MAX_CLASSIC_TILE_WIDTH,
      ),
      classicDockHeight: legacyClassicPanelSize(
        legacyClassicDockHeightStorageKey,
        DEFAULT_CLASSIC_DOCK_HEIGHT,
        MIN_CLASSIC_DOCK_HEIGHT,
        MAX_CLASSIC_DOCK_HEIGHT,
      ),
      keypadMinimizeOnCopy: true,
      inputBehavior: { ...DEFAULT_INPUT_BEHAVIOR_SETTINGS },
      personalize: true,
      suggestionCount: 6,
      checkUpdatesOnStartup: true,
      powerPointDefaultFontSizePt: 20,
      usage: {},
      history: [],
      setTitle: (title) => set({ title }),
      setActiveLineId: (activeLineId) =>
        set((state) => ({
          activeLineId: validActiveLineId(state.lines, activeLineId),
        })),
      replaceFormulaLine: (lineId, latex) =>
        set((state) => ({
          lines: state.lines.map((line) =>
            line.id === lineId
              ? { ...line, latex: normalizeFormulaLineLatex(latex) }
              : line,
          ),
        })),
      setFormulaAlignment: (formulaAlignment) =>
        set({ formulaAlignment: normalizeFormulaAlignment(formulaAlignment) }),
      setEditorLayout: (editorLayout) =>
        set({ editorLayout: normalizeEditorLayout(editorLayout) }),
      insertFormulaLine: (line, index) =>
        set((state) => {
          const nextLines = state.lines.filter((item) => item.id !== line.id);
          const targetIndex = Math.max(0, Math.min(index, nextLines.length));
          nextLines.splice(targetIndex, 0, {
            id: line.id,
            latex: normalizeFormulaLineLatex(line.latex),
            mode: line.mode === "inline" ? "inline" : "display",
          });
          return {
            lines: nextLines,
            activeLineId: validActiveLineId(nextLines, state.activeLineId),
          };
        }),
      removeFormulaLine: (lineId) =>
        set((state) => {
          const nextLines = state.lines.filter((line) => line.id !== lineId);
          const safeLines = nextLines.length ? nextLines : [createFormulaLine("")];
          return {
            lines: safeLines,
            activeLineId: validActiveLineId(safeLines, state.activeLineId),
          };
        }),
      replaceDocumentState: (snapshot) =>
        set(() => {
          const lines = normalizeFormulaLines(snapshot.lines);
          return {
            title: snapshot.title,
            lines,
            activeLineId: validActiveLineId(lines, snapshot.activeLineId),
            formulaAlignment: normalizeFormulaAlignment(
              snapshot.formulaAlignment,
            ),
          };
        }),
      setTheme: (theme) => set({ theme: normalizeTheme(theme) }),
      setLanguage: (language) => set({ language }),
      setZoom: (zoom) => set({
        zoom: isLandingPreview ? LANDING_PREVIEW_ZOOM : normalizeEditorZoom(zoom),
      }),
      setSourceOpen: (sourceOpen) => set({ sourceOpen }),
      setLatexCodeFormat: (latexCodeFormat) =>
        set({
          latexCodeFormat: isLatexCodeFormat(latexCodeFormat)
            ? latexCodeFormat
            : DEFAULT_LATEX_CODE_FORMAT,
        }),
      setAutoPairDelimiters: (autoPairDelimiters) =>
        set({ autoPairDelimiters }),
      setShowLineNumbers: (showLineNumbers) => set({ showLineNumbers }),
      setHighlightActiveLine: (highlightActiveLine) =>
        set({ highlightActiveLine }),
      setFormulaInsetLeft: (formulaInsetLeft) =>
        set({ formulaInsetLeft: normalizeFormulaInset(formulaInsetLeft) }),
      setFormulaInsetRight: (formulaInsetRight) =>
        set({ formulaInsetRight: normalizeFormulaInset(formulaInsetRight) }),
      setFormulaToolButtonSize: (formulaToolButtonSize) =>
        set({
          formulaToolButtonSize: normalizeFormulaToolButtonSize(
            formulaToolButtonSize,
          ),
        }),
      setFormulaToolButtonPadding: (formulaToolButtonPadding) =>
        set({
          formulaToolButtonPadding: normalizeFormulaToolButtonPadding(
            formulaToolButtonPadding,
          ),
        }),
      setFormulaRowVerticalInset: (formulaRowVerticalInset) =>
        set({
          formulaRowVerticalInset: normalizeFormulaRowVerticalInset(
            formulaRowVerticalInset,
          ),
        }),
      setPngExportBackground: (pngExportBackground) =>
        set({
          pngExportBackground: normalizePngExportBackground(
            pngExportBackground,
          ),
        }),
      setFormulaLetterFont: (formulaLetterFont) => {
        const normalized = normalizeFormulaLetterFont(formulaLetterFont);
        persistFormulaLetterFontPreference(normalized);
        set({ formulaLetterFont: normalized });
      },
      setFormulaChineseFont: (formulaChineseFont) => {
        const normalized = normalizeFormulaChineseFont(formulaChineseFont);
        persistFormulaChineseFontPreference(normalized);
        set({ formulaChineseFont: normalized });
      },
      setClassicTileWidth: (classicTileWidth) => {
        const normalized = normalizeClassicTileWidth(classicTileWidth);
        safeStorage.setItem(legacyClassicTileWidthStorageKey, String(normalized));
        set({ classicTileWidth: normalized });
      },
      setClassicDockHeight: (classicDockHeight) => {
        const normalized = normalizeClassicDockHeight(classicDockHeight);
        safeStorage.setItem(legacyClassicDockHeightStorageKey, String(normalized));
        set({ classicDockHeight: normalized });
      },
      setKeypadMinimizeOnCopy: (keypadMinimizeOnCopy) =>
        set({ keypadMinimizeOnCopy }),
      setInputBehavior: (setting, enabled) =>
        set((state) => ({
          inputBehavior: {
            ...state.inputBehavior,
            [setting]: enabled,
          },
        })),
      setPersonalize: (personalize) => set({ personalize }),
      setSuggestionCount: (suggestionCount) =>
        set({ suggestionCount: Math.min(10, Math.max(3, suggestionCount)) }),
      setCheckUpdatesOnStartup: (checkUpdatesOnStartup) =>
        set({ checkUpdatesOnStartup }),
      setPowerPointDefaultFontSizePt: (powerPointDefaultFontSizePt) =>
        set({
          powerPointDefaultFontSizePt:
            Math.round(Math.min(200, Math.max(5, powerPointDefaultFontSizePt)) * 2) /
            2,
        }),
      recordCommand: (commandId, prefix, source) =>
        set((state) => {
          const now = Date.now();
          const normalizedPrefix = prefix.replace(/^\\/, "").toLocaleLowerCase();
          const previous = state.usage[commandId] ?? {
            commandId,
            useCount: 0,
            lastUsedAt: 0,
            recentUses: [],
            acceptedPrefixes: {},
            contextCounts: {},
            pinned: false,
          };
          return {
            usage: {
              ...state.usage,
              [commandId]: {
                ...previous,
                useCount: previous.useCount + 1,
                lastUsedAt: now,
                recentUses: [...previous.recentUses, now].slice(-12),
                acceptedPrefixes: {
                  ...previous.acceptedPrefixes,
                  [normalizedPrefix]: (previous.acceptedPrefixes[normalizedPrefix] ?? 0) + 1,
                },
                contextCounts: {
                  ...(previous.contextCounts ?? {}),
                  [source]: (previous.contextCounts?.[source] ?? 0) + 1,
                },
              },
            },
          };
        }),
      resetUsage: () => set({ usage: {} }),
      addHistory: (latexOverride) =>
        set((state) => {
          const latex = normalizeMultilineLatex(
            latexOverride ?? joinFormulaLines(state.lines),
          );
          if (!latex.trim() || state.history[0]?.latex === latex) return state;
          const next: FormulaHistoryItem = {
            id: createUuid(),
            latex,
            createdAt: Date.now(),
            lines: cloneFormulaLines(state.lines),
          };
          return { history: [next, ...state.history].slice(0, 30) };
        }),
      clearHistory: () => set({ history: [] }),
      loadDocument: (document) =>
        set((state) => {
          const lines = normalizeFormulaLines(
            document.formulas.map((formula) => ({
              id: formula.id,
              latex: formula.latex,
              mode: formula.displayMode === "inline" ? "inline" : "display",
            })),
          );
          const settings = document.settings ?? {};
          const formulaLetterFont =
            settings.formulaLetterFont === undefined
              ? state.formulaLetterFont
              : normalizeFormulaLetterFont(settings.formulaLetterFont);
          const formulaChineseFont =
            settings.formulaChineseFont === undefined
              ? state.formulaChineseFont
              : normalizeFormulaChineseFont(settings.formulaChineseFont);
          if (
            settings.formulaLetterFont !== undefined ||
            settings.formulaChineseFont !== undefined
          ) {
            persistFormulaFontPreferences(
              formulaLetterFont,
              formulaChineseFont,
            );
          }
          return {
            title: document.title,
            lines,
            activeLineId: lines[0]?.id ?? null,
            formulaAlignment: normalizeFormulaAlignment(
              settings.formulaAlignment ?? document.formulas[0]?.alignment,
            ),
            editorLayout:
              settings.editorLayout === undefined
                ? state.editorLayout
                : normalizeEditorLayout(settings.editorLayout),
            theme:
              settings.theme === undefined
                ? state.theme
                : normalizeTheme(settings.theme),
            language:
              settings.language === "en"
                ? "en"
                : settings.language === "cn"
                  ? "cn"
                  : state.language,
            zoom:
              settings.zoom === undefined
                ? state.zoom
                : normalizeEditorZoom(settings.zoom),
            // Tools/source is workspace UI state, not formula-document state.
            // Keep the user's current workspace choice when opening old files
            // that still carry the legacy settings.sourceOpen field.
            sourceOpen: state.sourceOpen,
            latexCodeFormat: isLatexCodeFormat(settings.latexCodeFormat)
              ? settings.latexCodeFormat
              : state.latexCodeFormat,
            autoPairDelimiters:
              typeof settings.autoPairDelimiters === "boolean"
                ? settings.autoPairDelimiters
                : state.autoPairDelimiters,
            showLineNumbers:
              typeof settings.showLineNumbers === "boolean"
                ? settings.showLineNumbers
                : state.showLineNumbers,
            highlightActiveLine:
              typeof settings.highlightActiveLine === "boolean"
                ? settings.highlightActiveLine
                : state.highlightActiveLine,
            formulaInsetLeft:
              settings.formulaInsetLeft === undefined
                ? state.formulaInsetLeft
                : normalizeFormulaInset(settings.formulaInsetLeft),
            formulaInsetRight:
              settings.formulaInsetRight === undefined
                ? state.formulaInsetRight
                : normalizeFormulaInset(settings.formulaInsetRight),
            formulaToolButtonSize:
              settings.formulaToolButtonSize === undefined
                ? state.formulaToolButtonSize
                : normalizeFormulaToolButtonSize(settings.formulaToolButtonSize),
            formulaToolButtonPadding:
              settings.formulaToolButtonPadding === undefined
                ? state.formulaToolButtonPadding
                : normalizeFormulaToolButtonPadding(
                    settings.formulaToolButtonPadding,
                  ),
            formulaRowVerticalInset:
              settings.formulaRowVerticalInset === undefined
                ? state.formulaRowVerticalInset
                : normalizeFormulaRowVerticalInset(
                    settings.formulaRowVerticalInset,
                  ),
            pngExportBackground:
              settings.pngExportBackground === undefined
                ? state.pngExportBackground
                : normalizePngExportBackground(settings.pngExportBackground),
            formulaLetterFont,
            formulaChineseFont,
            inputBehavior:
              settings.inputBehavior === undefined
                ? state.inputBehavior
                : normalizeInputBehaviorSettings(settings.inputBehavior),
            personalize:
              typeof settings.personalize === "boolean"
                ? settings.personalize
                : state.personalize,
            suggestionCount:
              typeof settings.suggestionCount === "number" &&
              Number.isFinite(settings.suggestionCount)
                ? Math.min(10, Math.max(3, Math.round(settings.suggestionCount)))
                : state.suggestionCount,
            checkUpdatesOnStartup:
              typeof settings.checkUpdatesOnStartup === "boolean"
                ? settings.checkUpdatesOnStartup
                : state.checkUpdatesOnStartup,
            powerPointDefaultFontSizePt:
              typeof settings.powerPointDefaultFontSizePt === "number" &&
              Number.isFinite(settings.powerPointDefaultFontSizePt)
                ? Math.round(
                    Math.min(
                      200,
                      Math.max(5, settings.powerPointDefaultFontSizePt),
                    ) * 2,
                  ) / 2
                : state.powerPointDefaultFontSizePt,
            classicTileWidth:
              settings.classicTileWidth === undefined
                ? state.classicTileWidth
                : normalizeClassicTileWidth(settings.classicTileWidth),
            classicDockHeight:
              settings.classicDockHeight === undefined
                ? state.classicDockHeight
                : normalizeClassicDockHeight(settings.classicDockHeight),
            keypadMinimizeOnCopy:
              typeof settings.keypadMinimizeOnCopy === "boolean"
                ? settings.keypadMinimizeOnCopy
                : state.keypadMinimizeOnCopy,
          };
        }),
      toDocument: () => {
        const state = get();
        const now = Date.now();
        return {
          version: 3,
          title: state.title,
          formulas: state.lines.map((line) => ({
            id: line.id,
            latex: line.latex,
            displayMode: line.mode === "inline" ? "inline" : "block",
            alignment: state.formulaAlignment,
            fontSize: Math.round(36 * state.zoom),
            createdAt: now,
            updatedAt: now,
          })),
          macros: {},
          settings: {
            theme: state.theme,
            zoom: state.zoom,
            formulaAlignment: state.formulaAlignment,
            latexCodeFormat: state.latexCodeFormat,
            editorLayout: state.editorLayout,
            language: state.language,
            sourceOpen: state.sourceOpen,
            autoPairDelimiters: state.autoPairDelimiters,
            showLineNumbers: state.showLineNumbers,
            highlightActiveLine: state.highlightActiveLine,
            formulaInsetLeft: state.formulaInsetLeft,
            formulaInsetRight: state.formulaInsetRight,
            formulaToolButtonSize: state.formulaToolButtonSize,
            formulaToolButtonPadding: state.formulaToolButtonPadding,
            formulaRowVerticalInset: state.formulaRowVerticalInset,
            pngExportBackground: state.pngExportBackground,
            formulaLetterFont: state.formulaLetterFont,
            formulaChineseFont: state.formulaChineseFont,
            inputBehavior: { ...state.inputBehavior },
            personalize: state.personalize,
            suggestionCount: state.suggestionCount,
            checkUpdatesOnStartup: state.checkUpdatesOnStartup,
            powerPointDefaultFontSizePt: state.powerPointDefaultFontSizePt,
            classicTileWidth: state.classicTileWidth,
            classicDockHeight: state.classicDockHeight,
            keypadMinimizeOnCopy: state.keypadMinimizeOnCopy,
          },
        };
      },
    }),
    {
      name: "visualtex-editor",
      storage: createJSONStorage(() => safeStorage),
      partialize: (state) => ({
        title: state.title,
        lines: state.lines,
        activeLineId: state.activeLineId,
        formulaAlignment: state.formulaAlignment,
        editorLayout: state.editorLayout,
        theme: state.theme,
        language: state.language,
        zoom: state.zoom,
        sourceOpen: state.sourceOpen,
        latexCodeFormat: state.latexCodeFormat,
        autoPairDelimiters: state.autoPairDelimiters,
        showLineNumbers: state.showLineNumbers,
        highlightActiveLine: state.highlightActiveLine,
        formulaInsetLeft: state.formulaInsetLeft,
        formulaInsetRight: state.formulaInsetRight,
        formulaToolButtonSize: state.formulaToolButtonSize,
        formulaToolButtonPadding: state.formulaToolButtonPadding,
        formulaRowVerticalInset: state.formulaRowVerticalInset,
        pngExportBackground: state.pngExportBackground,
        formulaLetterFont: state.formulaLetterFont,
        formulaChineseFont: state.formulaChineseFont,
        classicTileWidth: state.classicTileWidth,
        classicDockHeight: state.classicDockHeight,
        keypadMinimizeOnCopy: state.keypadMinimizeOnCopy,
        inputBehavior: state.inputBehavior,
        personalize: state.personalize,
        suggestionCount: state.suggestionCount,
        checkUpdatesOnStartup: state.checkUpdatesOnStartup,
        powerPointDefaultFontSizePt: state.powerPointDefaultFontSizePt,
        usage: state.usage,
        history: state.history,
      }),
      merge: (persistedState, currentState) => {
        const persisted = persistedState as Partial<EditorState> & {
          latex?: string;
        };
        const { latex: legacyLatex, ...currentPersisted } = persisted;
        const storedFormulaFonts = readPersistedFormulaFontPreferences();
        const lines = normalizeFormulaLines(persisted.lines, legacyLatex);
        const legacyLineAlignment = Array.isArray(persisted.lines)
          ? (persisted.lines[0] as { alignment?: unknown } | undefined)?.alignment
          : undefined;
        return {
          ...currentState,
          ...currentPersisted,
          lines,
          activeLineId: validActiveLineId(lines, persisted.activeLineId),
          formulaAlignment: normalizeFormulaAlignment(
            persisted.formulaAlignment ?? legacyLineAlignment,
          ),
          editorLayout: normalizeEditorLayout(persisted.editorLayout),
          theme: normalizeTheme(persisted.theme),
          zoom: isLandingPreview ? LANDING_PREVIEW_ZOOM : normalizeEditorZoom(persisted.zoom),
          latexCodeFormat: isLatexCodeFormat(persisted.latexCodeFormat)
            ? persisted.latexCodeFormat
            : DEFAULT_LATEX_CODE_FORMAT,
          autoPairDelimiters:
            typeof persisted.autoPairDelimiters === "boolean"
              ? persisted.autoPairDelimiters
              : true,
          showLineNumbers:
            typeof persisted.showLineNumbers === "boolean"
              ? persisted.showLineNumbers
              : false,
          highlightActiveLine:
            typeof persisted.highlightActiveLine === "boolean"
              ? persisted.highlightActiveLine
              : false,
          formulaInsetLeft: normalizeFormulaInset(persisted.formulaInsetLeft),
          formulaInsetRight: normalizeFormulaInset(persisted.formulaInsetRight),
          formulaToolButtonSize: normalizeFormulaToolButtonSize(
            persisted.formulaToolButtonSize,
          ),
          formulaToolButtonPadding: normalizeFormulaToolButtonPadding(
            persisted.formulaToolButtonPadding,
          ),
          formulaRowVerticalInset: normalizeFormulaRowVerticalInset(
            persisted.formulaRowVerticalInset,
          ),
          pngExportBackground: normalizePngExportBackground(
            persisted.pngExportBackground,
          ),
          formulaLetterFont:
            storedFormulaFonts.formulaLetterFont ??
            normalizeFormulaLetterFont(persisted.formulaLetterFont),
          formulaChineseFont:
            storedFormulaFonts.formulaChineseFont ??
            normalizeFormulaChineseFont(persisted.formulaChineseFont),
          classicTileWidth:
            persisted.classicTileWidth === undefined
              ? legacyClassicPanelSize(
                  legacyClassicTileWidthStorageKey,
                  DEFAULT_CLASSIC_TILE_WIDTH,
                  MIN_CLASSIC_TILE_WIDTH,
                  MAX_CLASSIC_TILE_WIDTH,
                )
              : normalizeClassicTileWidth(persisted.classicTileWidth),
          classicDockHeight:
            persisted.classicDockHeight === undefined
              ? legacyClassicPanelSize(
                  legacyClassicDockHeightStorageKey,
                  DEFAULT_CLASSIC_DOCK_HEIGHT,
                  MIN_CLASSIC_DOCK_HEIGHT,
                  MAX_CLASSIC_DOCK_HEIGHT,
                )
              : normalizeClassicDockHeight(persisted.classicDockHeight),
          keypadMinimizeOnCopy:
            typeof persisted.keypadMinimizeOnCopy === "boolean"
              ? persisted.keypadMinimizeOnCopy
              : true,
          inputBehavior: normalizeInputBehaviorSettings(
            persisted.inputBehavior,
          ),
          powerPointDefaultFontSizePt:
            typeof persisted.powerPointDefaultFontSizePt === "number" &&
            Number.isFinite(persisted.powerPointDefaultFontSizePt)
              ? Math.round(
                  Math.min(200, Math.max(5, persisted.powerPointDefaultFontSizePt)) *
                    2,
                ) / 2
              : 20,
        };
      },
    },
  ),
);
