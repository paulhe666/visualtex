import type { ReactNode, RefObject } from "react";
import type { MathEditorHandle } from "../editor/MathEditor";
import type { DocumentSnapshot, ReplaceDocumentEntry } from "../history/historyTypes";

export type WorkspaceMode =
  | "web"
  | "office-create"
  | "office-edit";

export type WorkspaceExportFormat = "markdown" | "svg" | "png";

export interface EditorWorkspaceProps {
  mode: WorkspaceMode;

  showFileActions: boolean;
  showUpdateActions: boolean;
  showOfficeActions: boolean;

  primaryActionLabel?: string;
  officeHeaderLeadingControls?: ReactNode;
  officeHeaderTrailingActions?: ReactNode;
  desktopHeaderControls?: ReactNode;

  onPrimaryAction?: () => Promise<void>;
  onCancel?: () => Promise<void>;
  onOpenExport?: () => void;

  editorRef: RefObject<MathEditorHandle | null>;
  editorInstanceKey?: string;
  reuseEditorLineSlots?: boolean;
  sidebarOpen: boolean;
  onSidebarOpenChange: (open: boolean) => void;
  onHistoryBusyChange: (busy: boolean) => void;
  onCopyPng?: () => Promise<void>;
  onCopy: () => Promise<void>;
  onReplaceDocument: (
    snapshot: DocumentSnapshot,
    source: ReplaceDocumentEntry["source"],
  ) => boolean;
}
