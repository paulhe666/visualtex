import type { VisualTeXDialogMessage } from "../bridge/bridgeMessages";

export function messageOfficeParent(message: VisualTeXDialogMessage) {
  const isVstoDesktopRuntime =
    new URLSearchParams(window.location.search).get("runtime") === "vsto-desktop";
  if (
    isVstoDesktopRuntime ||
    typeof Office === "undefined" ||
    !Office.context?.ui?.messageParent
  ) {
    return false;
  }
  const serialized = JSON.stringify(message);
  const ui = Office.context.ui;
  try {
    ui.messageParent(serialized, { targetOrigin: window.location.origin });
  } catch {
    ui.messageParent(serialized);
  }
  return true;
}
