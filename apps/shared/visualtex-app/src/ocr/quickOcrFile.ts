// A screenshot the quick OCR capture returned (base64 and file extension), as a File for recognition.

export interface QuickOcrCapture {
  dataBase64: string;
  extension: string;
}

/** The texts of one capture mode (each [Chinese, English]), given by the app (src/platform/quickOcr.ts). */
export interface QuickOcrCaptureModeTexts {
  /** the mode in the capture menu */
  option: readonly [string, string];
  /** under "Run Quick OCR" while it is the current mode */
  current: readonly [string, string];
  /** the notice while waiting for the capture */
  waiting: readonly [string, string];
  /** the notice when nothing was captured */
  missed: readonly [string, string];
  /** the notice after choosing the mode */
  changed: readonly [string, string];
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
