import {
  parseDialogMessage,
  type VisualTeXDialogMessage,
} from "./bridgeMessages";

export const OFFICE_COMPANION_ORIGIN = "https://127.0.0.1:43127";

export interface DialogControllerCallbacks {
  onMessage: (message: VisualTeXDialogMessage) => void | Promise<void>;
  onClosed: (errorCode: number) => void | Promise<void>;
}

function runDialogCallback(label: string, callback: () => void | Promise<void>) {
  try {
    void Promise.resolve(callback()).catch((error) => {
      console.error(`VisualTeX Office dialog ${label} callback failed`, error);
    });
  } catch (error) {
    console.error(`VisualTeX Office dialog ${label} callback failed`, error);
  }
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
          try {
            this.dialog = dialog;
            dialog.addEventHandler(
              Office.EventType.DialogMessageReceived,
              (event) => {
                if (!("message" in event)) return;
                const parsed = parseDialogMessage(event.message);
                if (parsed) {
                  runDialogCallback("message", () => callbacks.onMessage(parsed));
                }
              },
            );
            dialog.addEventHandler(
              Office.EventType.DialogEventReceived,
              (event) => {
                if (!("error" in event)) return;
                this.dialog = null;
                runDialogCallback("closed", () => callbacks.onClosed(event.error));
              },
            );
            resolve();
          } catch (error) {
            this.dialog = null;
            try {
              dialog?.close();
            } catch {
              // The host may already have invalidated a partially-created dialog.
            }
            reject(
              error instanceof Error
                ? error
                : new Error("Unable to initialize the VisualTeX editor dialog."),
            );
          }
        },
      );
    });
  }

  close() {
    const dialog = this.dialog;
    this.dialog = null;
    if (!dialog) return;
    try {
      dialog.close();
    } catch (error) {
      // Office may invalidate the dialog handle before the companion observes
      // the final Session state. Closing is cleanup only and must never turn a
      // successfully committed formula into a failed Session.
      console.warn("VisualTeX Office dialog was already unavailable while closing", error);
    }
  }
}
