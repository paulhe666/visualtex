import type { AnywhereActivation } from "../../../shared/visualtex-app/src/anywhere/anywhereBridge";

/** VisualTeX Anywhere's formula window: not part of the web edition (APP_FEATURES.anywhere is false). */
export const anywhereWindow = {
  activationNotice: (_activation: AnywhereActivation): string | null => null,
  copiedNotice: "已复制",
  sharesFormulaShortcuts: false,
};
