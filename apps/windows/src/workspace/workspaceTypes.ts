import type { ReactNode, RefObject } from "react";
import type { QuickOcrCaptureMode } from "../ocr/quickOcr";
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

export type WorkspaceExportFormat = "markdown" | "svg" | "png";

export interface WorkspaceOcrRecognizerOption {
  id: string;
  group?: "local" | "api";
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
  officeHeaderLeadingControls?: ReactNode;
  officeHeaderTrailingActions?: ReactNode;
  desktopHeaderControls?: ReactNode;
  keypadMode?: boolean;

  onPrimaryAction?: () => Promise<void>;
  onCancel?: () => Promise<void>;
  onOpenExport?: () => void;

  editorRef: RefObject<MathEditorHandle | null>;
  editorInstanceKey?: string;
  reuseEditorLineSlots?: boolean;
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

  ocrRecognizer?: string;
  ocrRecognizers?: readonly WorkspaceOcrRecognizerOption[];
  ocrBusy?: boolean;
  onOcrRecognizerChange?: (recognizer: string) => void;
  onQuickOcr?: () => void;
  quickOcrCaptureMode?: QuickOcrCaptureMode;
  onQuickOcrCaptureModeChange?: (mode: QuickOcrCaptureMode) => void;
  silentOcrEnabled?: boolean;
  silentOcrShortcut?: string;
  onSilentOcrEnabledChange?: (enabled: boolean) => void;
  ocrOverlay?: ReactNode;
}
