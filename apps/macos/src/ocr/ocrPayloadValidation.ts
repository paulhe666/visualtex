import type {
  OcrInstallProgress,
  OcrProviderConfiguration,
  OcrRecognitionProgress,
  OcrRecognitionResult,
  OcrRuntimeStatus,
} from "./ocrService";

type JsonRecord = Record<string, unknown>;

const OCR_MODEL_NAMES = new Set([
  "PP-FormulaNet_plus-S",
  "PP-FormulaNet_plus-M",
  "PP-FormulaNet_plus-L",
]);
const OCR_PROVIDER_IDS = new Set([
  "local",
  "openai-compatible",
  "ollama",
  "mathpix",
  "paddleocr",
]);
const PADDLE_OCR_API_MODELS = new Set([
  "PaddleOCR-VL-1.6",
  "PP-StructureV3",
]);

function invalid(path: string, expectation: string): never {
  throw new Error(
    `VisualTeX OCR returned invalid data at ${path}; expected ${expectation}.`,
  );
}

function record(value: unknown, path: string): JsonRecord {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    invalid(path, "an object");
  }
  return value as JsonRecord;
}

function stringValue(value: unknown, path: string) {
  if (typeof value !== "string") invalid(path, "a string");
  return value;
}

function nonEmptyString(value: unknown, path: string) {
  const result = stringValue(value, path);
  if (!result.trim()) invalid(path, "a non-empty string");
  return result;
}

function nullableString(value: unknown, path: string) {
  if (value === null) return null;
  return stringValue(value, path);
}

function booleanValue(value: unknown, path: string) {
  if (typeof value !== "boolean") invalid(path, "a boolean");
  return value;
}

function finiteNumber(value: unknown, path: string) {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    invalid(path, "a finite number");
  }
  return value;
}

function nonNegativeNumber(value: unknown, path: string) {
  const result = finiteNumber(value, path);
  if (result < 0) invalid(path, "a non-negative finite number");
  return result;
}

function stringArray(value: unknown, path: string) {
  if (!Array.isArray(value)) invalid(path, "an array of strings");
  value.forEach((entry, index) => stringValue(entry, `${path}[${index}]`));
  return value as string[];
}

export function decodeOcrProviderConfiguration(
  value: unknown,
): OcrProviderConfiguration {
  const config = record(value, "ocrProviderConfiguration");
  if (
    typeof config.activeProvider !== "string" ||
    !OCR_PROVIDER_IDS.has(config.activeProvider)
  ) {
    invalid("ocrProviderConfiguration.activeProvider", "a supported OCR provider id");
  }
  const openAi = record(
    config.openAiCompatible,
    "ocrProviderConfiguration.openAiCompatible",
  );
  if (openAi.protocol !== "responses" && openAi.protocol !== "chat-completions") {
    invalid(
      "ocrProviderConfiguration.openAiCompatible.protocol",
      '"responses" or "chat-completions"',
    );
  }
  stringValue(openAi.baseUrl, "ocrProviderConfiguration.openAiCompatible.baseUrl");
  stringValue(openAi.model, "ocrProviderConfiguration.openAiCompatible.model");
  stringValue(openAi.prompt, "ocrProviderConfiguration.openAiCompatible.prompt");
  booleanValue(openAi.hasApiKey, "ocrProviderConfiguration.openAiCompatible.hasApiKey");

  const ollama = record(config.ollama, "ocrProviderConfiguration.ollama");
  stringValue(ollama.baseUrl, "ocrProviderConfiguration.ollama.baseUrl");
  stringValue(ollama.model, "ocrProviderConfiguration.ollama.model");
  stringValue(ollama.prompt, "ocrProviderConfiguration.ollama.prompt");

  const mathpix = record(config.mathpix, "ocrProviderConfiguration.mathpix");
  stringValue(mathpix.baseUrl, "ocrProviderConfiguration.mathpix.baseUrl");
  stringValue(mathpix.appId, "ocrProviderConfiguration.mathpix.appId");
  booleanValue(mathpix.hasAppKey, "ocrProviderConfiguration.mathpix.hasAppKey");

  const paddleOcr = record(config.paddleOcr, "ocrProviderConfiguration.paddleOcr");
  if (
    typeof paddleOcr.model !== "string" ||
    !PADDLE_OCR_API_MODELS.has(paddleOcr.model)
  ) {
    invalid(
      "ocrProviderConfiguration.paddleOcr.model",
      "a supported PaddleOCR AI Studio model",
    );
  }
  booleanValue(
    paddleOcr.hasAccessToken,
    "ocrProviderConfiguration.paddleOcr.hasAccessToken",
  );
  return config as unknown as OcrProviderConfiguration;
}

export function decodeOcrRuntimeStatus(value: unknown): OcrRuntimeStatus {
  const status = record(value, "ocrRuntimeStatus");
  booleanValue(status.installed, "ocrRuntimeStatus.installed");
  nullableString(status.pythonPath, "ocrRuntimeStatus.pythonPath");
  nullableString(status.pythonVersion, "ocrRuntimeStatus.pythonVersion");
  nullableString(status.paddleVersion, "ocrRuntimeStatus.paddleVersion");
  nullableString(status.paddleocrVersion, "ocrRuntimeStatus.paddleocrVersion");
  stringValue(status.runtimePath, "ocrRuntimeStatus.runtimePath");
  booleanValue(
    status.offlineBundleAvailable,
    "ocrRuntimeStatus.offlineBundleAvailable",
  );
  stringArray(status.installedModels, "ocrRuntimeStatus.installedModels");
  stringValue(status.defaultModel, "ocrRuntimeStatus.defaultModel");
  stringValue(status.message, "ocrRuntimeStatus.message");
  return status as unknown as OcrRuntimeStatus;
}

export function decodeOcrRecognitionResult(value: unknown): OcrRecognitionResult {
  const result = record(value, "ocrRecognitionResult");
  nonEmptyString(result.provider, "ocrRecognitionResult.provider");
  stringValue(result.model, "ocrRecognitionResult.model");
  nonNegativeNumber(result.elapsedMs, "ocrRecognitionResult.elapsedMs");
  nonNegativeNumber(result.processedWidth, "ocrRecognitionResult.processedWidth");
  nonNegativeNumber(result.processedHeight, "ocrRecognitionResult.processedHeight");
  booleanValue(
    result.backgroundInverted,
    "ocrRecognitionResult.backgroundInverted",
  );
  finiteNumber(
    result.backgroundLuminance,
    "ocrRecognitionResult.backgroundLuminance",
  );
  if (!Array.isArray(result.formulas)) {
    invalid("ocrRecognitionResult.formulas", "an array of formula results");
  }
  result.formulas.forEach((entry, index) => {
    const formula = record(entry, `ocrRecognitionResult.formulas[${index}]`);
    stringValue(formula.latex, `ocrRecognitionResult.formulas[${index}].latex`);
  });
  return result as unknown as OcrRecognitionResult;
}

export function decodeOcrRecognitionProgress(value: unknown): OcrRecognitionProgress {
  const progress = record(value, "ocrRecognitionProgress");
  if (progress.event !== "progress") {
    invalid("ocrRecognitionProgress.event", '"progress"');
  }
  nonEmptyString(progress.id, "ocrRecognitionProgress.id");
  stringValue(progress.stage, "ocrRecognitionProgress.stage");
  stringValue(progress.message, "ocrRecognitionProgress.message");
  if (typeof progress.model !== "string" || !OCR_MODEL_NAMES.has(progress.model)) {
    invalid("ocrRecognitionProgress.model", "a supported OCR model name");
  }
  return progress as unknown as OcrRecognitionProgress;
}

export function decodeOcrInstallProgress(value: unknown): OcrInstallProgress {
  const progress = record(value, "ocrInstallProgress");
  stringValue(progress.stage, "ocrInstallProgress.stage");
  const percent = finiteNumber(progress.percent, "ocrInstallProgress.percent");
  if (percent < 0 || percent > 100) {
    invalid("ocrInstallProgress.percent", "a percentage from 0 through 100");
  }
  stringValue(progress.message, "ocrInstallProgress.message");
  nullableString(progress.detail, "ocrInstallProgress.detail");
  return progress as unknown as OcrInstallProgress;
}
