import { invoke } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import type { LatexCodeFormat } from "../types/formula";
import type { OcrModelName } from "./ocrService";
import { decodeQuickOcrCapture } from "./quickOcrPayloadValidation";

export const SILENT_OCR_STORAGE_KEY = "visualtex.silent-ocr.enabled";
export const SILENT_OCR_SHORTCUT = "⌘⇧O";
export const QUICK_OCR_CAPTURE_MODE_STORAGE_KEY = "visualtex.quick-ocr.capture-mode";

export type QuickOcrCaptureMode = "immediate" | "system-screenshot";

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

export interface QuickOcrCapture {
  dataBase64: string;
  extension: string;
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
