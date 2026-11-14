import type { AnywhereActivation } from "../../../shared/visualtex-app/src/anywhere/anywhereBridge";

/** What VisualTeX Anywhere's formula window does the Windows way (anywhere/AnywhereApp.tsx). */
export const anywhereWindow = {
  /** the note shown when the formula opens: the input method is not connected to the field */
  activationNotice: (activation: AnywhereActivation): string | null =>
    activation.accessibilityTrusted ? null : "输入法未连接",
  copiedNotice: "已复制",
  /** The shortcuts that act while a formula is open are handed to the native side, which gives them to the page
   *  whatever the host application does with modifier chords (Chrome keeps them from the input method). */
  sharesFormulaShortcuts: true,
};
