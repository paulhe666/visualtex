import type { AnywhereActivation } from "../../../shared/visualtex-app/src/anywhere/anywhereBridge";

/** What VisualTeX Anywhere's formula window does the macOS way (anywhere/AnywhereApp.tsx). */
export const anywhereWindow = {
  /** the note shown when the formula opens: Accessibility not granted, the field's application to be reopened,
   *  or VisualTeX missing among the input sources */
  activationNotice: (activation: AnywhereActivation): string | null =>
    !activation.accessibilityTrusted
      ? "未授权辅助功能 · ↩ 复制"
      : activation.reopenApp
        ? `${activation.reopenApp} 需重新打开 · 之后公式写在输入框里`
        : null,
  copiedNotice: "已复制 · ⌘V 粘贴",
  /** the input method gives the open formula its keys itself */
  sharesFormulaShortcuts: false,
};
