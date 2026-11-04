// The document's input and output formats (docs/document-editor/DESIGN.md §7): LaTeX or Markdown input (what is
// typed converts as LaTeX commands or as Markdown; the candidate menu offers LaTeX or Markdown items) and LaTeX or
// Markdown output (the source tab, the copy button, Copy/Export).
import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import { safeStorage } from "../runtime/safeStorage";

export type DocumentFormat = "latex" | "markdown";

/** How formulas are written into Word: Office equations, VisualTeX formula pictures, MathType objects (Windows). */
export type WordFormulaFormat = "omml" | "visualtex" | "mathtype";

/** Groups of Markdown input rules that can be turned off (SiYuan had to add such switches after false positives). */
export interface MarkdownRules {
  /** "# ", "- ", "1. ", "> ", "```" at the start of a line; "---", "| a | b |" on Enter */
  blocks: boolean;
  /** **bold**, *italic*, `code`, ~~strike~~, ==highlight==, [text](url) */
  marks: boolean;
  /** a Chinese input method's 》 1。 ··· 、 ／ in place of > 1. ``` / */
  chinese: boolean;
}

interface DocumentFormatState {
  inputMode: DocumentFormat;
  outputFormat: DocumentFormat;
  /** Markdown output: references to numbered formulas as $\eqref{label}$ (links in Typora) or as plain "(n)" */
  markdownReferences: "eqref" | "plain";
  markdownRules: MarkdownRules;
  /** Word export: the formulas' form (the last one chosen) */
  wordFormulas: WordFormulaFormat;
  setWordFormulas: (format: WordFormulaFormat) => void;
  setInputMode: (mode: DocumentFormat) => void;
  setOutputFormat: (format: DocumentFormat) => void;
  setMarkdownReferences: (style: "eqref" | "plain") => void;
  setMarkdownRule: (rule: keyof MarkdownRules, on: boolean) => void;
}

export const useDocumentFormatStore = create<DocumentFormatState>()(
  persist(
    (set) => ({
      inputMode: "latex",
      outputFormat: "latex",
      markdownReferences: "eqref",
      markdownRules: { blocks: true, marks: true, chinese: true },
      wordFormulas: "omml",
      setWordFormulas: (wordFormulas) => set({ wordFormulas }),
      setInputMode: (inputMode) => set({ inputMode }),
      setOutputFormat: (outputFormat) => set({ outputFormat }),
      setMarkdownReferences: (markdownReferences) => set({ markdownReferences }),
      setMarkdownRule: (rule, on) => set((state) => ({ markdownRules: { ...state.markdownRules, [rule]: on } })),
    }),
    {
      name: "visualtex.document-format.v1",
      storage: createJSONStorage(() => safeStorage),
      partialize: (state) => ({
        inputMode: state.inputMode,
        outputFormat: state.outputFormat,
        markdownReferences: state.markdownReferences,
        markdownRules: state.markdownRules,
        wordFormulas: state.wordFormulas,
      }),
    },
  ),
);
