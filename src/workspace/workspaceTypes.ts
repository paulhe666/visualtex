import type { ReactNode, RefObject } from "react";
import type {
  MathEditorHandle,
  MathEditorInsertionTarget,
} from "../editor/MathEditor";
import type { DocumentSnapshot, ReplaceDocumentEntry } from "../history/historyTypes";

export type WorkspaceExportFormat = "markdown" | "svg" | "png";

export interface EditorWorkspaceProps {
  showFileActions: boolean;
  desktopHeaderControls?: ReactNode;
  onOpenExport?: () => void;

  editorRef: RefObject<MathEditorHandle | null>;
  sidebarOpen: boolean;
  onSidebarOpenChange: (open: boolean) => void;
  onHistoryBusyChange: (busy: boolean) => void;
  onPasteImage?: (
    file: File,
    target: MathEditorInsertionTarget,
  ) => Promise<void>;
  onCopyPng?: () => Promise<void>;
  onCopy: () => Promise<void>;
  onReplaceDocument: (
    snapshot: DocumentSnapshot,
    source: ReplaceDocumentEntry["source"],
  ) => boolean;
}
