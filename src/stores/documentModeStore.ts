// The main window shows the formula rows or a document (docs/document-editor/DESIGN.md §6.5).
import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import { safeStorage } from "../runtime/safeStorage";

export type WorkspaceContent = "formula" | "document";

interface DocumentModeState {
  content: WorkspaceContent;
  setContent: (content: WorkspaceContent) => void;
}

export const useDocumentModeStore = create<DocumentModeState>()(
  persist(
    (set) => ({
      content: "formula",
      setContent: (content) => set({ content }),
    }),
    {
      name: "visualtex.workspace-content.v1",
      storage: createJSONStorage(() => safeStorage),
      partialize: (state) => ({ content: state.content }),
    },
  ),
);

