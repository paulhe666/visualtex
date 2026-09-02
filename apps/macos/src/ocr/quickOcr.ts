import { invoke } from "@tauri-apps/api/core";
import type { LatexCodeFormat } from "../types/formula";
import type { OcrModelName } from "./ocrService";
import { decodeQuickOcrCapture } from "./quickOcrPayloadValidation";

export const SILENT_OCR_STORAGE_KEY = "visualtex.silent-ocr.enabled";
export const SILENT_OCR_SHORTCUT = "⌘⇧O";
export const QUICK_OCR_CAPTURE_MODE_STORAGE_KEY = "visualtex.quick-ocr.capture-mode";

export type QuickOcrCaptureMode = "immediate" | "system-screenshot";

export function isQuickOcrCaptureMode(value: unknown): value is QuickOcrCaptureMode {
  return value === "immediate" || value === "system-screenshot";
}

export interface QuickOcrCapture {
  dataBase64: string;
  extension: string;
}

export async function captureQuickOcrScreenshot() {
  return decodeQuickOcrCapture(
    await invoke<unknown>("capture_quick_ocr_screenshot"),
  );
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
) {
  await invoke("configure_silent_ocr", { enabled, model, copyFormat });
}

export function quickOcrCaptureToFile(capture: QuickOcrCapture) {
  const binary = atob(capture.dataBase64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new File([bytes], `VisualTeX-Quick-OCR.${capture.extension}`, {
    type: capture.extension === "png" ? "image/png" : "application/octet-stream",
  });
}
