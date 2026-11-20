import type { OcrBackendFeatures } from "../../../shared/visualtex-app/src/ocr/ocrService";

/**
 * The web edition recognizes formulas through online services only (src/ocr/webOcrTransport.ts: the browser, and
 * the site's relay worker for the services that do not accept browser requests). No local runtime, no models.
 */
export const OCR_BACKEND: OcrBackendFeatures = {
  warmupCommand: "prewarm_ocr_model",
  installRecovery: false,
  storageLocation: false,
  modelDownloads: false,
  remoteProviders: ["openai-compatible", "mathpix", "paddleocr", "simpletex"],
};
