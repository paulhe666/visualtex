import assert from "node:assert/strict";
import { PADDLE_OCR_API_MODELS } from "../src/ocr/ocrService.ts";
import {
  decodeOcrInstallProgress,
  decodeOcrProviderConfiguration,
  decodeOcrRecognitionProgress,
  decodeOcrRecognitionResult,
  decodeOcrRuntimeStatus,
} from "../src/ocr/ocrPayloadValidation.ts";

assert.deepEqual(
  PADDLE_OCR_API_MODELS,
  ["PaddleOCR-VL-1.6"],
  "Paddle API UI must expose only the validated formula model",
);

const runtime = {
  installed: true,
  pythonPath: "/tmp/python",
  pythonVersion: "3.10",
  paddleVersion: "3.3.1",
  paddleocrVersion: "3.7.0",
  runtimePath: "/tmp/runtime",
  offlineBundleAvailable: true,
  installedModels: ["PP-FormulaNet_plus-M"],
  defaultModel: "PP-FormulaNet_plus-M",
  message: "ready",
};
assert.equal(decodeOcrRuntimeStatus(runtime), runtime);
assert.throws(
  () => decodeOcrRuntimeStatus({ ...runtime, installedModels: [null] }),
  /installedModels\[0\]/,
);

const recognition = {
  provider: "local",
  model: "PP-FormulaNet_plus-M",
  elapsedMs: 123,
  processedWidth: 640,
  processedHeight: 320,
  backgroundInverted: false,
  backgroundLuminance: 0.95,
  formulas: [{ latex: "x^2" }],
};
assert.equal(decodeOcrRecognitionResult(recognition), recognition);
assert.throws(
  () => decodeOcrRecognitionResult({ ...recognition, formulas: [{ latex: 1 }] }),
  /formulas\[0\]\.latex/,
);
assert.throws(
  () => decodeOcrRecognitionResult({ ...recognition, provider: "" }),
  /\.provider/,
);

const providers = {
  activeProvider: "openai-compatible",
  openAiCompatible: {
    protocol: "responses",
    baseUrl: "https://api.openai.com/v1",
    model: "vision-model",
    prompt: "Return formula JSON",
    hasApiKey: true,
  },
  ollama: {
    baseUrl: "http://127.0.0.1:11434",
    model: "vision-model",
    prompt: "Return formula JSON",
  },
  mathpix: {
    baseUrl: "https://api.mathpix.com",
    appId: "app-id",
    hasAppKey: false,
  },
  paddleOcr: {
    model: "PaddleOCR-VL-1.6",
    hasAccessToken: true,
  },
  simpleTex: {
    model: "standard",
    hasAccessToken: true,
  },
};
assert.equal(decodeOcrProviderConfiguration(providers), providers);
assert.throws(
  () => decodeOcrProviderConfiguration({ ...providers, activeProvider: "unknown" }),
  /activeProvider/,
);
assert.throws(
  () =>
    decodeOcrProviderConfiguration({
      ...providers,
      paddleOcr: { ...providers.paddleOcr, model: "PP-OCRv5" },
    }),
  /paddleOcr\.model/,
);
assert.throws(
  () =>
    decodeOcrProviderConfiguration({
      ...providers,
      simpleTex: { ...providers.simpleTex, model: "invalid" },
    }),
  /simpleTex\.model/,
);

const progress = {
  event: "progress",
  id: "run-1",
  stage: "inference",
  message: "running",
  model: "PP-FormulaNet_plus-M",
};
assert.equal(decodeOcrRecognitionProgress(progress), progress);
assert.throws(
  () => decodeOcrRecognitionProgress({ ...progress, model: "unknown" }),
  /\.model/,
);

const install = {
  stage: "download",
  percent: 50,
  message: "halfway",
  detail: null,
};
assert.equal(decodeOcrInstallProgress(install), install);
assert.throws(
  () => decodeOcrInstallProgress({ ...install, percent: Number.NaN }),
  /\.percent/,
);

console.log("VisualTeX macOS OCR payload safety regression passed");
