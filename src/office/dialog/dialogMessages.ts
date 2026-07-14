import type { VisualTeXDialogMessage } from "../bridge/bridgeMessages";

export function messageOfficeParent(message: VisualTeXDialogMessage) {
  const serialized = JSON.stringify(message);
  const ui = Office.context.ui;
  try {
    ui.messageParent(serialized, { targetOrigin: window.location.origin });
  } catch {
    ui.messageParent(serialized);
  }
}
