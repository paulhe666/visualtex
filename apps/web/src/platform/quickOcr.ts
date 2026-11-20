// Screenshot OCR (quick OCR, silent OCR) needs the desktop app: a browser cannot capture the screen on a shortcut
// (APP_FEATURES.screenOcr is false, so the editor shows neither). The names exist for the shared code that reads
// them: the capture mode setting kept in configuration backups, and the stored silent OCR switch.
import type { LatexCodeFormat } from "../../../shared/visualtex-app/src/types/formula";
import type { OcrModelName } from "../../../shared/visualtex-app/src/ocr/ocrService";
import type { QuickOcrCaptureModeTexts } from "../../../shared/visualtex-app/src/ocr/quickOcrFile";
import type { QuickOcrCapture } from "../../../shared/visualtex-app/src/ocr/quickOcrFile";
export { quickOcrCaptureToFile, type QuickOcrCapture } from "../../../shared/visualtex-app/src/ocr/quickOcrFile";

export const SILENT_OCR_STORAGE_KEY = "visualtex.silent-ocr.enabled";
export const SILENT_OCR_SHORTCUT = "";
export const QUICK_OCR_CAPTURE_MODE_STORAGE_KEY = "visualtex.quick-ocr.capture-mode";

export type QuickOcrCaptureMode = "immediate";

export const QUICK_OCR_CAPTURE_MODES: readonly ({ id: QuickOcrCaptureMode } & QuickOcrCaptureModeTexts)[] = [
  {
    id: "immediate",
    option: ["立即框选", "Immediate selection"],
    current: ["立即框选", "Immediate selection"],
    waiting: ["框选公式", "Select the formula"],
    missed: ["已取消截图", "Screenshot cancelled"],
    changed: ["快捷 OCR：立即框选", "Quick OCR: select at once"],
  },
];

export function normalizeQuickOcrCaptureMode(value: unknown): QuickOcrCaptureMode | null {
  return typeof value === "string" && value ? "immediate" : null;
}

export function isQuickOcrCaptureMode(value: unknown): value is QuickOcrCaptureMode {
  return value === "immediate";
}

export async function captureQuickOcrScreenshot(_captureMode: QuickOcrCaptureMode = "immediate"): Promise<QuickOcrCapture | null> {
  return null;
}

export async function restoreQuickOcrWindow() {}

export async function configureSilentOcr(
  _enabled: boolean,
  _model: OcrModelName,
  _copyFormat: LatexCodeFormat,
  _captureMode?: QuickOcrCaptureMode,
) {
  return "";
}
