import assert from "node:assert/strict";
import {
  decodeBooleanOcrResult,
  decodeNullableOcrModelDownloadSnapshot,
  decodeOcrInstallProgress,
  decodeOcrInstallStatus,
  decodeOcrModelCatalog,
  decodeOcrModelDownloadSnapshot,
  decodeOcrProviderConfiguration,
  decodeOcrRecognitionProgress,
  decodeOcrRecognitionResult,
  decodeOcrRuntimeStatus,
} from "../src/ocr/ocrPayloadValidation.ts";

const provider = {
  activeProvider: "local",
  openAiCompatible: {
    protocol: "responses",
    baseUrl: "https://example.invalid/v1",
    model: "vision",
    prompt: "read formula",
    hasApiKey: false,
  },
  ollama: {
    baseUrl: "http://127.0.0.1:11434",
    model: "vision",
    prompt: "read formula",
  },
  mathpix: {
    baseUrl: "https://api.mathpix.com",
    appId: "",
    hasAppKey: false,
  },
  paddleOcr: {
    model: "PaddleOCR-VL-1.6",
    hasAccessToken: true,
  },
};
assert.equal(decodeOcrProviderConfiguration(provider), provider);
assert.throws(
  () => decodeOcrProviderConfiguration({ ...provider, activeProvider: "bad" }),
  /activeProvider/,
);
assert.throws(
  () => decodeOcrProviderConfiguration({
    ...provider,
    paddleOcr: { ...provider.paddleOcr, model: "PP-OCRv5" },
  }),
  /paddleOcr\.model/,
);

const runtime = {
  installed: true,
  pythonPath: "C:/VisualTeX/python.exe",
  pythonVersion: "3.10",
  paddleVersion: "3.3.1",
  paddleocrVersion: "3.7.0",
  runtimePath: "C:/VisualTeX/OCR",
  storageConfigPath: "C:/VisualTeX/ocr-storage.json",
  storageSource: "default",
  storageManaged: true,
  storageAvailableBytes: 10_000,
  storagePersistentAcrossUninstall: true,
  runtimeBundleAvailable: true,
  offlineBundleAvailable: true,
  installedModels: ["PP-FormulaNet_plus-M"],
  damagedModels: [],
  modelCatalogAvailable: true,
  defaultModel: "PP-FormulaNet_plus-M",
  message: "ready",
};
assert.equal(decodeOcrRuntimeStatus(runtime), runtime);
assert.throws(
  () => decodeOcrRuntimeStatus({ ...runtime, damagedModels: [null] }),
  /damagedModels\[0\]/,
);

const installStatus = {
  schemaVersion: 1,
  state: "complete",
  currentStep: null,
  completedSteps: ["runtime"],
  percent: 100,
  message: "done",
  detail: null,
  error: null,
  logPath: "C:/VisualTeX/install.log",
  updatedAtMs: 1_788_278_400_000,
};
assert.equal(decodeOcrInstallStatus(installStatus), installStatus);
assert.throws(
  () => decodeOcrInstallStatus({ ...installStatus, state: "unknown" }),
  /\.state/,
);

const installProgress = {
  stage: "download",
  state: "installing",
  percent: 50,
  message: "halfway",
  detail: null,
  error: null,
  logPath: null,
};
assert.equal(decodeOcrInstallProgress(installProgress), installProgress);

const recognition = {
  provider: "local",
  model: "PP-FormulaNet_plus-M",
  elapsedMs: 120,
  processedWidth: 800,
  processedHeight: 400,
  backgroundInverted: false,
  backgroundLuminance: 0.9,
  formulas: [{ latex: "x^2" }],
};
assert.equal(decodeOcrRecognitionResult(recognition), recognition);
assert.throws(
  () => decodeOcrRecognitionResult({ ...recognition, formulas: [{}] }),
  /formulas\[0\]\.latex/,
);

const recognitionProgress = {
  event: "progress",
  id: "run-1",
  stage: "inference",
  message: "running",
  model: "PP-FormulaNet_plus-M",
};
assert.equal(decodeOcrRecognitionProgress(recognitionProgress), recognitionProgress);
assert.throws(
  () => decodeOcrRecognitionProgress({ ...recognitionProgress, model: "bad" }),
  /\.model/,
);

const catalog = {
  schemaVersion: 1,
  platform: "windows",
  architecture: "x64",
  entries: [
    {
      model: "PP-FormulaNet_plus-M",
      url: "https://example.invalid/model",
      size: 100,
      sha256: "a".repeat(64),
    },
  ],
};
assert.equal(decodeOcrModelCatalog(catalog), catalog);
assert.throws(
  () =>
    decodeOcrModelCatalog({
      ...catalog,
      entries: [{ ...catalog.entries[0], sha256: "bad" }],
    }),
  /sha256/,
);

const download = {
  model: "PP-FormulaNet_plus-M",
  state: "downloading",
  downloadedBytes: 50,
  totalBytes: 100,
  percent: 50,
  speedBytesPerSecond: 10,
  etaSeconds: 5,
  message: "downloading",
  error: null,
};
assert.equal(decodeOcrModelDownloadSnapshot(download), download);
assert.equal(decodeNullableOcrModelDownloadSnapshot(null), null);
assert.throws(
  () => decodeOcrModelDownloadSnapshot({ ...download, percent: 101 }),
  /\.percent/,
);
assert.equal(decodeBooleanOcrResult(false, "cancelled"), false);
assert.throws(() => decodeBooleanOcrResult("false", "cancelled"), /cancelled/);

console.log("VisualTeX Windows OCR payload safety regression passed");
