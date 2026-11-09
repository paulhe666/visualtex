// Document mode's Markdown settings (docs/document-editor/DESIGN.md §7): how references to numbered formulas are
// written and which groups of Markdown input rules are on. Documents are typed and written as Markdown.
import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import { safeStorage } from "../runtime/safeStorage";

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
  /** Markdown output: references to numbered formulas as $\eqref{label}$ (links in Typora) or as plain "(n)" */
  markdownReferences: "eqref" | "plain";
  markdownRules: MarkdownRules;
  setMarkdownReferences: (style: "eqref" | "plain") => void;
  setMarkdownRule: (rule: keyof MarkdownRules, on: boolean) => void;
}

export const useDocumentFormatStore = create<DocumentFormatState>()(
  persist(
    (set) => ({
      markdownReferences: "eqref",
      markdownRules: { blocks: true, marks: true, chinese: true },
      setMarkdownReferences: (markdownReferences) => set({ markdownReferences }),
      setMarkdownRule: (rule, on) => set((state) => ({ markdownRules: { ...state.markdownRules, [rule]: on } })),
    }),
    {
      name: "visualtex.document-format.v1",
      storage: createJSONStorage(() => safeStorage),
      partialize: (state) => ({ markdownReferences: state.markdownReferences, markdownRules: state.markdownRules }),
      // (the input and output formats an earlier version stored here are left out)
      merge: (persisted, current) => {
        const stored = (persisted ?? {}) as Partial<DocumentFormatState>;
        return {
          ...current,
          markdownReferences: stored.markdownReferences ?? current.markdownReferences,
          markdownRules: { ...current.markdownRules, ...stored.markdownRules },
        };
      },
    },
  ),
);
