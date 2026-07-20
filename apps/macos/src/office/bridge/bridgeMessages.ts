export interface DialogCommitMessage {
  type: "visualtex-commit";
  sessionId: string;
}

export interface DialogCancelMessage {
  type: "visualtex-cancel";
  sessionId: string;
}

export interface DialogReadyMessage {
  type: "visualtex-ready";
  sessionId: string;
}

export interface DialogErrorMessage {
  type: "visualtex-error";
  sessionId: string;
  message: string;
}

export type VisualTeXDialogMessage =
  | DialogCommitMessage
  | DialogCancelMessage
  | DialogReadyMessage
  | DialogErrorMessage;

export function parseDialogMessage(value: string): VisualTeXDialogMessage | null {
  try {
    const parsed = JSON.parse(value) as Partial<VisualTeXDialogMessage>;
    if (typeof parsed.type !== "string" || typeof parsed.sessionId !== "string") {
      return null;
    }
    if (
      parsed.type === "visualtex-commit" ||
      parsed.type === "visualtex-cancel" ||
      parsed.type === "visualtex-ready"
    ) {
      return parsed as VisualTeXDialogMessage;
    }
    if (parsed.type === "visualtex-error" && typeof parsed.message === "string") {
      return parsed as DialogErrorMessage;
    }
    return null;
  } catch {
    return null;
  }
}
