import {
  parseDialogMessage,
  type VisualTeXDialogMessage,
} from "./bridgeMessages";

export const OFFICE_COMPANION_ORIGIN = "https://127.0.0.1:43127";

export interface DialogControllerCallbacks {
  onMessage: (message: VisualTeXDialogMessage) => void | Promise<void>;
  onClosed: (errorCode: number) => void | Promise<void>;
}

export class DialogController {
  private dialog: Office.Dialog | null = null;

  get isOpen() {
    return this.dialog !== null;
  }

  async open(
    sessionId: string,
    callbacks: DialogControllerCallbacks,
  ): Promise<void> {
    if (this.dialog) {
      throw new Error("A VisualTeX Office editor window is already open.");
    }

    const dialogUrl = `${OFFICE_COMPANION_ORIGIN}/dialog/${encodeURIComponent(
      sessionId,
    )}`;

    await new Promise<void>((resolve, reject) => {
      Office.context.ui.displayDialogAsync(
        dialogUrl,
        {
          width: 90,
          height: 90,
          displayInIframe: false,
        },
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            reject(
              new Error(
                result.error?.message ?? "Unable to open the VisualTeX editor.",
              ),
            );
            return;
          }

          const dialog = result.value;
          this.dialog = dialog;
          dialog.addEventHandler(
            Office.EventType.DialogMessageReceived,
            (event) => {
              if (!("message" in event)) return;
              const parsed = parseDialogMessage(event.message);
              if (parsed) void callbacks.onMessage(parsed);
            },
          );
          dialog.addEventHandler(
            Office.EventType.DialogEventReceived,
            (event) => {
              if (!("error" in event)) return;
              this.dialog = null;
              void callbacks.onClosed(event.error);
            },
          );
          resolve();
        },
      );
    });
  }

  close() {
    const dialog = this.dialog;
    this.dialog = null;
    dialog?.close();
  }
}
