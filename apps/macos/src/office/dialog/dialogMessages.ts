import type { VisualTeXDialogMessage } from "../bridge/bridgeMessages";

interface OptionalOfficeUi {
  messageParent(message: string, options?: { targetOrigin: string }): void;
}

export function messageOfficeParent(message: VisualTeXDialogMessage) {
  const office = (globalThis as typeof globalThis & {
    Office?: { context?: { ui?: OptionalOfficeUi } };
  }).Office;
  const ui = office?.context?.ui;
  if (!ui || typeof ui.messageParent !== "function") return;

  const serialized = JSON.stringify(message);
  try {
    ui.messageParent(serialized, { targetOrigin: window.location.origin });
  } catch {
    ui.messageParent(serialized);
  }
}
