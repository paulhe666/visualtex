import { invoke } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import type { LatexCodeFormat } from "../types/formula";
import type { OcrModelName } from "./ocrService";

export const SILENT_OCR_STORAGE_KEY = "visualtex.silent-ocr.enabled";
export const SILENT_OCR_SHORTCUT = "Ctrl+Alt+O";
export const QUICK_OCR_CAPTURE_MODE_STORAGE_KEY =
  "visualtex.quick-ocr.capture-mode";

export type QuickOcrCaptureMode = "windows" | "pixpin" | "clipboard";

/**
 * Keep configuration imports from older 1.2.x builds working while exposing
 * names that describe the actual capture provider. "clipboard" accepts the
 * next image written by PixPin, ShareX, Snipping Tool, or any other utility.
 */
export function normalizeQuickOcrCaptureMode(
  value: unknown,
): QuickOcrCaptureMode | null {
  if (value === "windows" || value === "immediate") return "windows";
  if (value === "pixpin") return "pixpin";
  if (value === "clipboard" || value === "system-screenshot") return "clipboard";
  return null;
}

export function isQuickOcrCaptureMode(
  value: unknown,
): value is QuickOcrCaptureMode {
  return value === "windows" || value === "pixpin" || value === "clipboard";
}

export interface QuickOcrCapture {
  dataBase64: string;
  extension: string;
}

function hasTauriRuntime() {
  return Boolean(
    (window as Window & { __TAURI_INTERNALS__?: { metadata?: unknown } })
      .__TAURI_INTERNALS__?.metadata,
  );
}

async function captureWindowsClipboardImage(captureMode: QuickOcrCaptureMode) {
  if (!hasTauriRuntime()) {
    throw new Error("Windows Quick OCR is available in the desktop app only.");
  }
  return invoke<QuickOcrCapture | null>("capture_windows_quick_ocr", {
    captureMode,
    timeoutMs: 60_000,
  });
}

export async function restoreQuickOcrWindow() {
  if (!hasTauriRuntime()) return;
  try {
    const current = getCurrentWindow();
    await current.show();
    await current.unminimize();
    await current.setFocus();
  } catch {
    // OCR insertion can still finish even if Windows refuses focus restoration.
  }
}

export async function writeSilentOcrClipboardText(text: string) {
  if (!hasTauriRuntime()) {
    throw new Error("Windows silent OCR clipboard access is unavailable.");
  }
  await invoke("write_windows_ocr_clipboard_text", { text });
}

/**
 * Capture with the selected provider, then read the newly committed Windows
 * clipboard image directly. Windows Snipping Tool remains the default; PixPin
 * uses its official scripting command, while clipboard mode waits for any tool.
 */
export async function captureQuickOcrScreenshot(
  captureMode: QuickOcrCaptureMode = "windows",
) {
  return captureWindowsClipboardImage(captureMode);
}

/** Register the Windows global Ctrl+Alt+O bridge while OCR stays in React. */
export async function configureSilentOcr(
  enabled: boolean,
  model: OcrModelName,
  copyFormat: LatexCodeFormat,
  captureMode: QuickOcrCaptureMode,
) {
  if (!hasTauriRuntime()) return enabled ? SILENT_OCR_SHORTCUT : "";
  return invoke<string>("configure_silent_ocr", {
    enabled,
    model,
    copyFormat,
    captureMode,
  });
}

export function quickOcrCaptureToFile(capture: QuickOcrCapture) {
  const binary = atob(capture.dataBase64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  const extension = capture.extension || "png";
  return new File([bytes], `VisualTeX-Quick-OCR.${extension}`, {
    type:
      extension === "jpg" || extension === "jpeg"
        ? "image/jpeg"
        : extension === "webp"
          ? "image/webp"
          : "image/png",
  });
}
