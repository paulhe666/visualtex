import type { ReactNode, RefObject } from "react";
import type {
  MathEditorHandle,
  MathEditorInsertionTarget,
} from "../editor/MathEditor";
import type { ReplaceDocumentEntry } from "../history/historyTypes";
import type { DocumentSnapshot } from "../history/historyTypes";

export type WorkspaceMode =
  | "desktop"
  | "office-create"
  | "office-edit";

export interface WorkspaceOcrModelOption {
  id: string;
  labelZh: string;
  labelEn: string;
}

export interface EditorWorkspaceProps {
  mode: WorkspaceMode;

  showFileActions: boolean;
  showUpdateActions: boolean;
  showOfficeActions: boolean;
  showOcrActions: boolean;

  primaryActionLabel?: string;

  onPrimaryAction?: () => Promise<void>;
  onCancel?: () => Promise<void>;
  onExportMarkdown?: () => void;

  editorRef: RefObject<MathEditorHandle | null>;
  sidebarOpen: boolean;
  onSidebarOpenChange: (open: boolean) => void;
  onHistoryBusyChange: (busy: boolean) => void;
  onPasteImage?: (
    file: File,
    target: MathEditorInsertionTarget,
  ) => Promise<void>;
  onCopy: () => Promise<void>;
  onReplaceDocument: (
    snapshot: DocumentSnapshot,
    source: ReplaceDocumentEntry["source"],
  ) => boolean;

  ocrModel?: string;
  ocrModels?: readonly WorkspaceOcrModelOption[];
  ocrBusy?: boolean;
  onOcrModelChange?: (model: string) => void;
  ocrOverlay?: ReactNode;
}
