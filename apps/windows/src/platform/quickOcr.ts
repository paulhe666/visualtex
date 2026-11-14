import { invoke } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import type { LatexCodeFormat } from "../../../shared/visualtex-app/src/types/formula";
import type { OcrModelName } from "../../../shared/visualtex-app/src/ocr/ocrService";
import {
  decodeQuickOcrCapture,
  decodeSilentOcrShortcut,
} from "../../../shared/visualtex-app/src/ocr/quickOcrPayloadValidation";
import type { QuickOcrCaptureModeTexts } from "../../../shared/visualtex-app/src/ocr/quickOcrFile";
export { quickOcrCaptureToFile, type QuickOcrCapture } from "../../../shared/visualtex-app/src/ocr/quickOcrFile";

export const SILENT_OCR_STORAGE_KEY = "visualtex.silent-ocr.enabled";
export const SILENT_OCR_SHORTCUT = "Ctrl+Alt+O";
export const QUICK_OCR_CAPTURE_MODE_STORAGE_KEY =
  "visualtex.quick-ocr.capture-mode";

export type QuickOcrCaptureMode = "windows" | "pixpin" | "clipboard";

/** Windows' capture modes, the default first: Snipping Tool, PixPin (its scripting command), or the next image
 *  any tool writes to the clipboard. */
export const QUICK_OCR_CAPTURE_MODES: readonly ({ id: QuickOcrCaptureMode } & QuickOcrCaptureModeTexts)[] = [
  {
    id: "windows",
    option: ["Windows 截图工具（默认）", "Windows Snipping Tool (default)"],
    current: ["Windows 截图工具", "Windows Snipping Tool"],
    waiting: ["在截图工具中框选公式", "Select the formula in Snipping Tool"],
    missed: ["已取消截图", "Screenshot cancelled"],
    changed: ["截图来源：截图工具", "Capture source: Snipping Tool"],
  },
  {
    id: "pixpin",
    option: ["PixPin", "PixPin"],
    current: ["PixPin", "PixPin"],
    waiting: ["在 PixPin 中框选公式", "Select the formula in PixPin"],
    missed: ["已取消截图", "Screenshot cancelled"],
    changed: ["截图来源：PixPin", "Capture source: PixPin"],
  },
  {
    id: "clipboard",
    option: ["其他截图工具 / 剪贴板", "Other screenshot tool / clipboard"],
    current: ["剪贴板", "Clipboard"],
    waiting: ["等待剪贴板图片（60 秒）", "Waiting for a clipboard image (60 s)"],
    missed: ["60 秒内未收到剪贴板图片", "No new clipboard image was detected within 60 seconds"],
    changed: ["截图来源：剪贴板", "Capture source: clipboard"],
  },
];

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
  return decodeQuickOcrCapture(
    await invoke<unknown>("capture_windows_quick_ocr", {
      captureMode,
      timeoutMs: 60_000,
    }),
  );
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
  return decodeSilentOcrShortcut(
    await invoke<unknown>("configure_silent_ocr", {
      enabled,
      model,
      copyFormat,
      captureMode,
    }),
  );
}
