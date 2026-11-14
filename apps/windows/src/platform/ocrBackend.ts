import type { OcrBackendFeatures } from "../../../shared/visualtex-app/src/ocr/ocrService";

/**
 * The Windows OCR backend: besides installing the runtime and importing model packages it keeps the installation
 * state (cancel, retry from the failed step, the install log), stores everything in a folder the user chooses and
 * downloads verified models from its catalog.
 */
export const OCR_BACKEND: OcrBackendFeatures = {
  warmupCommand: "warmup_ocr_model",
  installRecovery: true,
  storageLocation: true,
  modelDownloads: true,
};
