import type { VisualTeXDialogMessage } from "../bridge/bridgeMessages";

export function messageOfficeParent(message: VisualTeXDialogMessage) {
  const office = globalThis.Office;
  const ui = office?.context?.ui;
  if (!ui || typeof ui.messageParent !== "function") return;

  const serialized = JSON.stringify(message);
  try {
    ui.messageParent(serialized, { targetOrigin: window.location.origin });
  } catch {
    ui.messageParent(serialized);
  }
}
