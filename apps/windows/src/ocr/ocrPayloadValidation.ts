import type {
  OcrInstallProgress,
  OcrInstallState,
  OcrInstallStatus,
  OcrModelCatalog,
  OcrModelDownloadSnapshot,
  OcrModelDownloadState,
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
const OCR_INSTALL_STATES = new Set<OcrInstallState>([
  "notInstalled",
  "installing",
  "installFailed",
  "dependenciesInstalled",
  "verifying",
  "verificationFailed",
  "complete",
  "cancelled",
]);
const OCR_DOWNLOAD_STATES = new Set<OcrModelDownloadState>([
  "idle",
  "downloading",
  "verifying",
  "installing",
  "complete",
  "cancelled",
  "failed",
]);
const OCR_PROVIDER_IDS = new Set([
  "local",
  "openai-compatible",
  "ollama",
  "mathpix",
  "paddleocr",
  "simpletex",
]);
const PADDLE_OCR_API_MODELS = new Set([
  "PaddleOCR-VL-1.6",
]);
const SIMPLETEX_API_MODELS = new Set(["standard", "turbo"]);

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

function nonNegativeInteger(value: unknown, path: string) {
  const result = nonNegativeNumber(value, path);
  if (!Number.isSafeInteger(result)) invalid(path, "a non-negative integer");
  return result;
}

function positiveInteger(value: unknown, path: string) {
  const result = nonNegativeInteger(value, path);
  if (result < 1) invalid(path, "a positive integer");
  return result;
}

function percentage(value: unknown, path: string) {
  const result = finiteNumber(value, path);
  if (result < 0 || result > 100) {
    invalid(path, "a percentage from 0 through 100");
  }
  return result;
}

function stringArray(value: unknown, path: string) {
  if (!Array.isArray(value)) invalid(path, "an array of strings");
  value.forEach((entry, index) => stringValue(entry, `${path}[${index}]`));
  return value as string[];
}

function supportedModel(value: unknown, path: string) {
  if (typeof value !== "string" || !OCR_MODEL_NAMES.has(value)) {
    invalid(path, "a supported OCR model name");
  }
  return value;
}

export function decodeOcrProviderConfiguration(
  value: unknown,
): OcrProviderConfiguration {
  const config = record(value, "ocrProviderConfiguration");
  if (typeof config.activeProvider !== "string" || !OCR_PROVIDER_IDS.has(config.activeProvider)) {
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
  const simpleTex = record(
    config.simpleTex,
    "ocrProviderConfiguration.simpleTex",
  );
  if (
    typeof simpleTex.model !== "string" ||
    !SIMPLETEX_API_MODELS.has(simpleTex.model)
  ) {
    invalid(
      "ocrProviderConfiguration.simpleTex.model",
      "a supported SimpleTex formula model",
    );
  }
  booleanValue(
    simpleTex.hasAccessToken,
    "ocrProviderConfiguration.simpleTex.hasAccessToken",
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
  stringValue(status.storageConfigPath, "ocrRuntimeStatus.storageConfigPath");
  stringValue(status.storageSource, "ocrRuntimeStatus.storageSource");
  booleanValue(status.storageManaged, "ocrRuntimeStatus.storageManaged");
  if (status.storageAvailableBytes !== null) {
    nonNegativeNumber(
      status.storageAvailableBytes,
      "ocrRuntimeStatus.storageAvailableBytes",
    );
  }
  booleanValue(
    status.storagePersistentAcrossUninstall,
    "ocrRuntimeStatus.storagePersistentAcrossUninstall",
  );
  booleanValue(status.runtimeBundleAvailable, "ocrRuntimeStatus.runtimeBundleAvailable");
  booleanValue(status.offlineBundleAvailable, "ocrRuntimeStatus.offlineBundleAvailable");
  stringArray(status.installedModels, "ocrRuntimeStatus.installedModels");
  stringArray(status.damagedModels, "ocrRuntimeStatus.damagedModels");
  booleanValue(status.modelCatalogAvailable, "ocrRuntimeStatus.modelCatalogAvailable");
  stringValue(status.defaultModel, "ocrRuntimeStatus.defaultModel");
  stringValue(status.message, "ocrRuntimeStatus.message");
  return status as unknown as OcrRuntimeStatus;
}

function installState(value: unknown, path: string): OcrInstallState {
  if (typeof value !== "string" || !OCR_INSTALL_STATES.has(value as OcrInstallState)) {
    invalid(path, "a supported OCR install state");
  }
  return value as OcrInstallState;
}

export function decodeOcrInstallStatus(value: unknown): OcrInstallStatus {
  const status = record(value, "ocrInstallStatus");
  positiveInteger(status.schemaVersion, "ocrInstallStatus.schemaVersion");
  installState(status.state, "ocrInstallStatus.state");
  nullableString(status.currentStep, "ocrInstallStatus.currentStep");
  stringArray(status.completedSteps, "ocrInstallStatus.completedSteps");
  percentage(status.percent, "ocrInstallStatus.percent");
  stringValue(status.message, "ocrInstallStatus.message");
  nullableString(status.detail, "ocrInstallStatus.detail");
  nullableString(status.error, "ocrInstallStatus.error");
  stringValue(status.logPath, "ocrInstallStatus.logPath");
  nonNegativeInteger(status.updatedAtMs, "ocrInstallStatus.updatedAtMs");
  return status as unknown as OcrInstallStatus;
}

export function decodeOcrInstallProgress(value: unknown): OcrInstallProgress {
  const progress = record(value, "ocrInstallProgress");
  stringValue(progress.stage, "ocrInstallProgress.stage");
  installState(progress.state, "ocrInstallProgress.state");
  percentage(progress.percent, "ocrInstallProgress.percent");
  stringValue(progress.message, "ocrInstallProgress.message");
  nullableString(progress.detail, "ocrInstallProgress.detail");
  nullableString(progress.error, "ocrInstallProgress.error");
  nullableString(progress.logPath, "ocrInstallProgress.logPath");
  return progress as unknown as OcrInstallProgress;
}

export function decodeOcrRecognitionResult(value: unknown): OcrRecognitionResult {
  const result = record(value, "ocrRecognitionResult");
  nonEmptyString(result.provider, "ocrRecognitionResult.provider");
  stringValue(result.model, "ocrRecognitionResult.model");
  nonNegativeNumber(result.elapsedMs, "ocrRecognitionResult.elapsedMs");
  nonNegativeNumber(result.processedWidth, "ocrRecognitionResult.processedWidth");
  nonNegativeNumber(result.processedHeight, "ocrRecognitionResult.processedHeight");
  booleanValue(result.backgroundInverted, "ocrRecognitionResult.backgroundInverted");
  finiteNumber(result.backgroundLuminance, "ocrRecognitionResult.backgroundLuminance");
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
  supportedModel(progress.model, "ocrRecognitionProgress.model");
  return progress as unknown as OcrRecognitionProgress;
}

export function decodeOcrModelCatalog(value: unknown): OcrModelCatalog {
  const catalog = record(value, "ocrModelCatalog");
  positiveInteger(catalog.schemaVersion, "ocrModelCatalog.schemaVersion");
  if (catalog.platform !== "windows") invalid("ocrModelCatalog.platform", '"windows"');
  if (catalog.architecture !== "x64") invalid("ocrModelCatalog.architecture", '"x64"');
  if (!Array.isArray(catalog.entries)) invalid("ocrModelCatalog.entries", "an array");
  catalog.entries.forEach((entry, index) => {
    const item = record(entry, `ocrModelCatalog.entries[${index}]`);
    supportedModel(item.model, `ocrModelCatalog.entries[${index}].model`);
    nonEmptyString(item.url, `ocrModelCatalog.entries[${index}].url`);
    nonNegativeInteger(item.size, `ocrModelCatalog.entries[${index}].size`);
    if (typeof item.sha256 !== "string" || !/^[0-9a-f]{64}$/i.test(item.sha256)) {
      invalid(`ocrModelCatalog.entries[${index}].sha256`, "a SHA-256 digest");
    }
  });
  return catalog as unknown as OcrModelCatalog;
}

function downloadState(value: unknown, path: string): OcrModelDownloadState {
  if (
    typeof value !== "string" ||
    !OCR_DOWNLOAD_STATES.has(value as OcrModelDownloadState)
  ) {
    invalid(path, "a supported OCR model download state");
  }
  return value as OcrModelDownloadState;
}

export function decodeOcrModelDownloadSnapshot(
  value: unknown,
): OcrModelDownloadSnapshot {
  const status = record(value, "ocrModelDownload");
  supportedModel(status.model, "ocrModelDownload.model");
  downloadState(status.state, "ocrModelDownload.state");
  nonNegativeNumber(status.downloadedBytes, "ocrModelDownload.downloadedBytes");
  nonNegativeNumber(status.totalBytes, "ocrModelDownload.totalBytes");
  percentage(status.percent, "ocrModelDownload.percent");
  nonNegativeNumber(
    status.speedBytesPerSecond,
    "ocrModelDownload.speedBytesPerSecond",
  );
  if (status.etaSeconds !== null) {
    nonNegativeNumber(status.etaSeconds, "ocrModelDownload.etaSeconds");
  }
  stringValue(status.message, "ocrModelDownload.message");
  nullableString(status.error, "ocrModelDownload.error");
  return status as unknown as OcrModelDownloadSnapshot;
}

export function decodeNullableOcrModelDownloadSnapshot(
  value: unknown,
): OcrModelDownloadSnapshot | null {
  return value === null ? null : decodeOcrModelDownloadSnapshot(value);
}

export function decodeBooleanOcrResult(value: unknown, path: string): boolean {
  return booleanValue(value, path);
}
