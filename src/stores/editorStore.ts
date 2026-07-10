import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import type { CommandSource, CommandUsage } from "../types/command";
import type { FormulaDocument, FormulaHistoryItem } from "../types/formula";
import { normalizeMultilineLatex } from "../editor/normalizeChineseLatex";

type Theme = "light" | "dark";
export type Language = "cn" | "en";

interface EditorState {
  title: string;
  latex: string;
  theme: Theme;
  language: Language;
  zoom: number;
  sourceOpen: boolean;
  personalize: boolean;
  suggestionCount: number;
  usage: Record<string, CommandUsage>;
  history: FormulaHistoryItem[];
  setTitle: (title: string) => void;
  setLatex: (latex: string) => void;
  setTheme: (theme: Theme) => void;
  setLanguage: (language: Language) => void;
  setZoom: (zoom: number) => void;
  setSourceOpen: (open: boolean) => void;
  setPersonalize: (enabled: boolean) => void;
  setSuggestionCount: (count: number) => void;
  recordCommand: (commandId: string, prefix: string, source: CommandSource) => void;
  resetUsage: () => void;
  addHistory: (latex?: string) => void;
  clearHistory: () => void;
  loadDocument: (document: FormulaDocument) => void;
  toDocument: () => FormulaDocument;
}

const initialLatex = "\\int_{-\\infty}^{\\infty} e^{-x^2}\\,\\mathrm{d}x = \\sqrt{\\pi}";

export const useEditorStore = create<EditorState>()(
  persist(
    (set, get) => ({
      title: "未命名公式",
      latex: initialLatex,
      theme: "light",
      language: "cn",
      zoom: 1,
      sourceOpen: true,
      personalize: true,
      suggestionCount: 6,
      usage: {},
      history: [],
      setTitle: (title) => set({ title }),
      setLatex: (latex) => set({ latex: normalizeMultilineLatex(latex) }),
      setTheme: (theme) => set({ theme }),
      setLanguage: (language) => set({ language }),
      setZoom: (zoom) => set({ zoom: Math.min(1.6, Math.max(0.7, zoom)) }),
      setSourceOpen: (sourceOpen) => set({ sourceOpen }),
      setPersonalize: (personalize) => set({ personalize }),
      setSuggestionCount: (suggestionCount) =>
        set({ suggestionCount: Math.min(10, Math.max(3, suggestionCount)) }),
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
          const latex = normalizeMultilineLatex(latexOverride ?? state.latex);
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
        set({
          title: document.title,
          latex: normalizeMultilineLatex(
            document.formulas.length
              ? document.formulas.map((formula) => formula.latex).join("\n")
              : "",
          ),
          theme: document.settings.theme,
          zoom: document.settings.zoom,
        }),
      toDocument: () => {
        const state = get();
        const now = Date.now();
        return {
          version: 2,
          title: state.title,
          formulas: state.latex.split("\n").map((latex) => ({
            id: crypto.randomUUID(),
            latex,
            displayMode: "block",
            alignment: "center",
            fontSize: Math.round(36 * state.zoom),
            createdAt: now,
            updatedAt: now,
          })),
          macros: {},
          settings: { theme: state.theme, zoom: state.zoom },
        };
      },
    }),
    {
      name: "visualtex-editor",
      storage: createJSONStorage(() => localStorage),
      partialize: (state) => ({
        title: state.title,
        latex: state.latex,
        theme: state.theme,
        language: state.language,
        zoom: state.zoom,
        sourceOpen: state.sourceOpen,
        personalize: state.personalize,
        suggestionCount: state.suggestionCount,
        usage: state.usage,
        history: state.history,
      }),
    },
  ),
);
