import { PowerPointAdapter } from "./PowerPointAdapter";
import { WordAdapter } from "./WordAdapter";
import type {
  CreateOfficeSessionInput,
  NativePowerPointCommitSelection,
  OfficeFormulaSession,
  OfficeHost,
  OfficeSessionMode,
} from "../api/sessionClient";

export interface OfficeSelectionContext {
  sourceDocumentId: string | null;
  sourceObjectId: string | null;
  sessionSeed: Partial<CreateOfficeSessionInput>;
}

export interface OfficeInteractionTarget {
  host: "word" | "powerpoint";
  formulaId: string;
  shapeName: string;
  slideIndex?: number;
  slideId?: number;
  presentationIdentity?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

export interface OfficeHostAdapter {
  readonly host: OfficeHost;
  readonly requiredExportFormat?: "svg" | "png";

  readSelection(mode: OfficeSessionMode): Promise<OfficeSelectionContext>;
  /** Supplies the immutable shape snapshot captured at double-click time.
   * PowerPoint can move focus to its native format UI before readSelection
   * runs, so the adapter must not depend on a second selection query. */
  prepareInteractionTarget?(target: OfficeInteractionTarget): void;
  applySession(session: OfficeFormulaSession): Promise<void>;
  /** macOS Word uses the exact persisted alternative-text payload to find the
   * inserted picture after Office.js releases the mutable selection. */
  getNativeWordFormulaMarker?(sessionId: string): string | null;
  /** macOS PowerPoint native paste is only the preparation phase. The hidden
   * Office.js command page must decorate and verify the exact selected shape
   * before the companion is allowed to confirm the Session. */
  finalizeNativePowerPointCommit?(
    session: OfficeFormulaSession,
    selection: NativePowerPointCommitSelection,
  ): Promise<void>;
  /** Word-only command. Kept optional so PowerPoint adapters don't expose a
   * command that their manifest never registers. */
  updateEquationNumbers?(): Promise<number>;
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
