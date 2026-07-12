import type {
  CreateOfficeSessionInput,
  OfficeFormulaSession,
  OfficeHost,
  OfficeSessionMode,
} from "../api/sessionClient";

export interface OfficeSelectionContext {
  sourceDocumentId: string | null;
  sourceObjectId: string | null;
  sessionSeed: Partial<CreateOfficeSessionInput>;
}

export interface OfficeHostAdapter {
  readonly host: OfficeHost;

  readSelection(mode: OfficeSessionMode): Promise<OfficeSelectionContext>;
  applySession(session: OfficeFormulaSession): Promise<void>;
  openDesktopApp(): Promise<void>;
  showMessage(message: string): void;
}

class PendingOfficeHostAdapter implements OfficeHostAdapter {
  constructor(readonly host: OfficeHost) {}

  async readSelection(mode: OfficeSessionMode): Promise<OfficeSelectionContext> {
    if (mode === "edit") {
      throw new Error(
        "当前选中的对象不是 VisualTeX 公式。请先选择由 VisualTeX 插入的公式。",
      );
    }
    return {
      sourceDocumentId: null,
      sourceObjectId: null,
      sessionSeed: {},
    };
  }

  async applySession(): Promise<void> {
    throw new Error(`${this.host} adapter is not installed yet`);
  }

  async openDesktopApp(): Promise<void> {
    window.location.href = "visualtex://office/start";
  }

  showMessage(message: string) {
    const status = document.getElementById("bridge-status");
    if (status) status.textContent = message;
  }
}

export function officeHostFromReadyInfo(host: Office.HostType): OfficeHost {
  if (host === Office.HostType.Word) return "word";
  if (host === Office.HostType.PowerPoint) return "powerpoint";
  throw new Error("VisualTeX Office integration supports Word and PowerPoint only.");
}

export function createOfficeHostAdapter(host: OfficeHost): OfficeHostAdapter {
  return new PendingOfficeHostAdapter(host);
}
