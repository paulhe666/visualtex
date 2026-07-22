import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import type { CommandSource, CommandUsage } from "../types/command";
import type {
  FormulaDocument,
  FormulaHistoryItem,
  FormulaLine,
  InputBehaviorSettingKey,
  InputBehaviorSettings,
  LatexCodeFormat,
} from "../types/formula";
import type { DocumentSnapshot } from "../history/historyTypes";
import {
  DEFAULT_LATEX_CODE_FORMAT,
  isLatexCodeFormat,
} from "../clipboard/LatexCopyService";
import { normalizeChineseLatex } from "../editor/normalizeChineseLatex";
import { normalizeMultilineLatex } from "../editor/normalizeChineseLatex";

type Theme = "light" | "dark";
export type Language = "cn" | "en";
export const MIN_EDITOR_ZOOM = 0.2;
export const MAX_EDITOR_ZOOM = 1.6;

export const DEFAULT_INPUT_BEHAVIOR_SETTINGS: InputBehaviorSettings = {
  autoExitSuperscript: true,
  autoExitSubscript: true,
  autoExitAccent: true,
};

function normalizeInputBehaviorSettings(
  value: unknown,
): InputBehaviorSettings {
  const candidate =
    value && typeof value === "object"
      ? (value as Partial<InputBehaviorSettings>)
      : {};
  return {
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
  };
}

function normalizeEditorZoom(value: unknown) {
  const zoom = typeof value === "number" && Number.isFinite(value) ? value : 1;
  return Math.min(
    MAX_EDITOR_ZOOM,
    Math.max(MIN_EDITOR_ZOOM, Math.round(zoom * 10) / 10),
  );
}

export function createFormulaLine(
  latex = "",
  id: string = crypto.randomUUID(),
): FormulaLine {
  return {
    id,
    latex: normalizeChineseLatex(latex.replace(/\r\n?/g, "\n").split("\n")[0] ?? ""),
  };
}

function uniqueLineId(candidate: unknown, usedIds: Set<string>) {
  const normalized = typeof candidate === "string" ? candidate.trim() : "";
  if (normalized && !usedIds.has(normalized)) {
    usedIds.add(normalized);
    return normalized;
  }
  let nextId: string = crypto.randomUUID();
  while (usedIds.has(nextId)) nextId = crypto.randomUUID();
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
          latex: normalizeChineseLatex(
            typeof candidate.latex === "string"
              ? candidate.latex.replace(/\r\n?/g, "\n").split("\n")[0] ?? ""
              : "",
          ),
        } satisfies FormulaLine;
      })
      .filter((line): line is FormulaLine => line !== null);
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
  theme: Theme;
  language: Language;
  zoom: number;
  sourceOpen: boolean;
  latexCodeFormat: LatexCodeFormat;
  autoPairDelimiters: boolean;
  inputBehavior: InputBehaviorSettings;
  personalize: boolean;
  suggestionCount: number;
  checkUpdatesOnStartup: boolean;
  usage: Record<string, CommandUsage>;
  history: FormulaHistoryItem[];
  setTitle: (title: string) => void;
  setActiveLineId: (lineId: string | null) => void;
  replaceFormulaLine: (lineId: string, latex: string) => void;
  insertFormulaLine: (line: FormulaLine, index: number) => void;
  removeFormulaLine: (lineId: string) => void;
  replaceDocumentState: (snapshot: DocumentSnapshot) => void;
  setTheme: (theme: Theme) => void;
  setLanguage: (language: Language) => void;
  setZoom: (zoom: number) => void;
  setSourceOpen: (open: boolean) => void;
  setLatexCodeFormat: (format: LatexCodeFormat) => void;
  setAutoPairDelimiters: (enabled: boolean) => void;
  setInputBehavior: (
    setting: InputBehaviorSettingKey,
    enabled: boolean,
  ) => void;
  setPersonalize: (enabled: boolean) => void;
  setSuggestionCount: (count: number) => void;
  setCheckUpdatesOnStartup: (enabled: boolean) => void;
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
      theme: "light",
      language: "cn",
      zoom: 1,
      sourceOpen: false,
      latexCodeFormat: DEFAULT_LATEX_CODE_FORMAT,
      autoPairDelimiters: true,
      inputBehavior: { ...DEFAULT_INPUT_BEHAVIOR_SETTINGS },
      personalize: true,
      suggestionCount: 6,
      checkUpdatesOnStartup: true,
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
              ? { ...line, latex: normalizeChineseLatex(latex) }
              : line,
          ),
        })),
      insertFormulaLine: (line, index) =>
        set((state) => {
          const nextLines = state.lines.filter((item) => item.id !== line.id);
          const targetIndex = Math.max(0, Math.min(index, nextLines.length));
          nextLines.splice(targetIndex, 0, {
            id: line.id,
            latex: normalizeChineseLatex(line.latex),
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
          };
        }),
      setTheme: (theme) => set({ theme }),
      setLanguage: (language) => set({ language }),
      setZoom: (zoom) => set({ zoom: normalizeEditorZoom(zoom) }),
      setSourceOpen: (sourceOpen) => set({ sourceOpen }),
      setLatexCodeFormat: (latexCodeFormat) =>
        set({
          latexCodeFormat: isLatexCodeFormat(latexCodeFormat)
            ? latexCodeFormat
            : DEFAULT_LATEX_CODE_FORMAT,
        }),
      setAutoPairDelimiters: (autoPairDelimiters) =>
        set({ autoPairDelimiters }),
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
      recordCommand: (commandId, prefix) =>
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
            id: crypto.randomUUID(),
            latex,
            createdAt: Date.now(),
          };
          return { history: [next, ...state.history].slice(0, 30) };
        }),
      clearHistory: () => set({ history: [] }),
      loadDocument: (document) =>
        set(() => {
          const lines = normalizeFormulaLines(
            document.formulas.map((formula) => ({
              id: formula.id,
              latex: formula.latex,
            })),
          );
          return {
            title: document.title,
            lines,
            activeLineId: lines[0]?.id ?? null,
            theme: document.settings.theme,
            zoom: normalizeEditorZoom(document.settings.zoom),
            latexCodeFormat: isLatexCodeFormat(document.settings.latexCodeFormat)
              ? document.settings.latexCodeFormat
              : DEFAULT_LATEX_CODE_FORMAT,
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
            displayMode: "block",
            alignment: "center",
            fontSize: Math.round(36 * state.zoom),
            createdAt: now,
            updatedAt: now,
          })),
          macros: {},
          settings: {
            theme: state.theme,
            zoom: state.zoom,
            latexCodeFormat: state.latexCodeFormat,
          },
        };
      },
    }),
    {
      name: "visualtex-editor",
      storage: createJSONStorage(() => localStorage),
      partialize: (state) => ({
        title: state.title,
        lines: state.lines,
        activeLineId: state.activeLineId,
        theme: state.theme,
        language: state.language,
        zoom: state.zoom,
        sourceOpen: state.sourceOpen,
        latexCodeFormat: state.latexCodeFormat,
        autoPairDelimiters: state.autoPairDelimiters,
        inputBehavior: state.inputBehavior,
        personalize: state.personalize,
        suggestionCount: state.suggestionCount,
        checkUpdatesOnStartup: state.checkUpdatesOnStartup,
        usage: state.usage,
        history: state.history,
      }),
      merge: (persistedState, currentState) => {
        const persisted = persistedState as Partial<EditorState> & {
          latex?: string;
        };
        const { latex: legacyLatex, ...currentPersisted } = persisted;
        const lines = normalizeFormulaLines(persisted.lines, legacyLatex);
        return {
          ...currentState,
          ...currentPersisted,
          lines,
          activeLineId: validActiveLineId(lines, persisted.activeLineId),
          zoom: normalizeEditorZoom(persisted.zoom),
          latexCodeFormat: isLatexCodeFormat(persisted.latexCodeFormat)
            ? persisted.latexCodeFormat
            : DEFAULT_LATEX_CODE_FORMAT,
          autoPairDelimiters:
            typeof persisted.autoPairDelimiters === "boolean"
              ? persisted.autoPairDelimiters
              : true,
          inputBehavior: normalizeInputBehaviorSettings(
            persisted.inputBehavior,
          ),
        };
      },
    },
  ),
);
