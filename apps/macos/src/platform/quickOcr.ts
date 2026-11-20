import { invoke } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import type { LatexCodeFormat } from "../../../shared/visualtex-app/src/types/formula";
import type { OcrModelName } from "../../../shared/visualtex-app/src/ocr/ocrService";
import { decodeQuickOcrCapture } from "../../../shared/visualtex-app/src/ocr/quickOcrPayloadValidation";
import type { QuickOcrCaptureModeTexts } from "../../../shared/visualtex-app/src/ocr/quickOcrFile";
export { quickOcrCaptureToFile, type QuickOcrCapture } from "../../../shared/visualtex-app/src/ocr/quickOcrFile";

export const SILENT_OCR_STORAGE_KEY = "visualtex.silent-ocr.enabled";
export const SILENT_OCR_SHORTCUT = "⌘⇧O";
export const QUICK_OCR_CAPTURE_MODE_STORAGE_KEY = "visualtex.quick-ocr.capture-mode";

export type QuickOcrCaptureMode = "immediate" | "system-screenshot";

/** macOS' capture modes, the default first: select at once, or wait for a system screenshot (⌘⇧3 / 4 / 5). */
export const QUICK_OCR_CAPTURE_MODES: readonly ({ id: QuickOcrCaptureMode } & QuickOcrCaptureModeTexts)[] = [
  {
    id: "immediate",
    option: ["立即框选（默认）", "Immediate selection (default)"],
    current: ["立即框选", "Immediate selection"],
    waiting: ["框选公式", "Select the formula"],
    missed: ["已取消截图", "Screenshot cancelled"],
    changed: ["快捷 OCR：立即框选", "Quick OCR: select at once"],
  },
  {
    id: "system-screenshot",
    option: ["等待系统截图（⌘⇧3 / 4 / 5）", "Wait for system screenshot (⌘⇧3 / 4 / 5)"],
    current: ["等待 macOS 系统截图", "Wait for macOS screenshot"],
    waiting: ["请在 60 秒内截图", "Take a screenshot within 60 seconds"],
    missed: ["60 秒内未收到系统截图", "No screenshot within 60 seconds"],
    changed: ["快捷 OCR：等待系统截图", "Quick OCR: wait for a screenshot"],
  },
];

export function normalizeQuickOcrCaptureMode(
  value: unknown,
): QuickOcrCaptureMode | null {
  if (value === "immediate" || value === "windows") return "immediate";
  if (value === "system-screenshot" || value === "clipboard") {
    return "system-screenshot";
  }
  return null;
}

export function isQuickOcrCaptureMode(value: unknown): value is QuickOcrCaptureMode {
  return normalizeQuickOcrCaptureMode(value) !== null;
}


export async function captureQuickOcrScreenshot(
  captureMode: QuickOcrCaptureMode = "immediate",
) {
  return captureMode === "system-screenshot"
    ? waitForQuickOcrSystemScreenshot()
    : decodeQuickOcrCapture(
        await invoke<unknown>("capture_quick_ocr_screenshot"),
      );
}

export async function restoreQuickOcrWindow() {
  try {
    const current = getCurrentWindow();
    await current.show();
    await current.unminimize();
    await current.setFocus();
  } catch {
    // Recognition/insertion remains valid even when macOS refuses focus.
  }
}

export async function waitForQuickOcrSystemScreenshot() {
  return decodeQuickOcrCapture(
    await invoke<unknown>("wait_for_quick_ocr_system_screenshot"),
  );
}

export async function configureSilentOcr(
  enabled: boolean,
  model: OcrModelName,
  copyFormat: LatexCodeFormat,
  _captureMode?: QuickOcrCaptureMode,
) {
  await invoke("configure_silent_ocr", { enabled, model, copyFormat });
  return enabled ? SILENT_OCR_SHORTCUT : "";
}
