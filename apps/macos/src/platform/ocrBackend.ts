import type { OcrBackendFeatures } from "../../../shared/visualtex-app/src/ocr/ocrService";

/**
 * The macOS OCR backend installs the bundled runtime and imports model packages; it has no install recovery, no
 * storage folder of the user's choice and no model downloads, so the OCR dialog leaves those parts out.
 */
export const OCR_BACKEND: OcrBackendFeatures = {
  warmupCommand: "prewarm_ocr_model",
  installRecovery: false,
  storageLocation: false,
  modelDownloads: false,
};
