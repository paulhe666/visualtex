import { PowerPointAdapter } from "./PowerPointAdapter";
import { WordAdapter } from "./WordAdapter";
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
  readonly requiredExportFormat?: "svg" | "png";

  readSelection(mode: OfficeSessionMode): Promise<OfficeSelectionContext>;
  applySession(session: OfficeFormulaSession): Promise<void>;
  openDesktopApp(): Promise<void>;
  showMessage(message: string): void;
}

export function officeHostFromReadyInfo(host: Office.HostType): OfficeHost {
  if (host === Office.HostType.Word) return "word";
  if (host === Office.HostType.PowerPoint) return "powerpoint";
  throw new Error("VisualTeX Office integration supports Word and PowerPoint only.");
}

export function createOfficeHostAdapter(host: OfficeHost): OfficeHostAdapter {
  return host === "word" ? new WordAdapter() : new PowerPointAdapter();
}
