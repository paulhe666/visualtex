import { VISUALTEX_ALIGNMENT_MARKER_LATEX } from "../editor/alignmentMarkers.ts";
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
} from "./ocrPayloadValidation";

export type UnlistenFn = () => void;

export interface OcrTransportEvent<T> {
  event: string;
  id: number;
  payload: T;
}

export interface OcrTransport {
  environment: "desktop" | "office";
  invoke<T>(command: string, args?: Record<string, unknown>): Promise<T>;
  listen<T>(
    eventName: string,
    handler: (event: OcrTransportEvent<T>) => void,
  ): Promise<UnlistenFn>;
}

let configuredTransport: OcrTransport | null = null;

export function configureOcrTransport(transport: OcrTransport) {
  configuredTransport = transport;
}

function activeTransport() {
  if (!configuredTransport) {
    throw new Error("VisualTeX OCR transport has not been initialized.");
  }
  return configuredTransport;
}

function invoke<T>(command: string, args?: Record<string, unknown>) {
  return activeTransport().invoke<T>(command, args);
}

function listen<T>(
  eventName: string,
  handler: (event: OcrTransportEvent<T>) => void,
) {
  return activeTransport().listen(eventName, handler);
}

export const OCR_MODELS = [
  {
    id: "PP-FormulaNet_plus-S",
    labelZh: "高速版 S",
    labelEn: "Fast S",
    hintZh: "独立模型包，安装后约 248 MB；速度最快，主要适合英文公式",
    hintEn: "Separate model package, about 248 MB installed; fastest for English formulas",
    downloadMb: 259.6,
    storageMb: 248,
    cpuBenchmarkMs: 260.99,
  },
  {
    id: "PP-FormulaNet_plus-M",
    labelZh: "均衡版 M（推荐）",
    labelEn: "Balanced M (recommended)",
    hintZh: "独立模型包；兼顾中文、复杂公式与速度，推荐首次安装",
    hintEn: "Separate model package; balanced for Chinese and complex formulas, recommended first",
    downloadMb: 620.5,
    storageMb: 592,
    cpuBenchmarkMs: 1615.8,
  },
  {
    id: "PP-FormulaNet_plus-L",
    labelZh: "高精度版 L",
    labelEn: "High accuracy L",
    hintZh: "独立模型包，安装后约 698 MB；首次加载较久，并会占用数 GB 内存",
    hintEn: "Separate model package, about 698 MB installed; first load is slow and may use several GB of memory",
    downloadMb: 731.5,
    storageMb: 698,
    cpuBenchmarkMs: 3125.58,
  },
] as const;

export type OcrModelName = (typeof OCR_MODELS)[number]["id"];
export const DEFAULT_OCR_MODEL: OcrModelName = "PP-FormulaNet_plus-M";

export interface OcrRuntimeStatus {
  installed: boolean;
  pythonPath: string | null;
  pythonVersion: string | null;
  paddleVersion: string | null;
  paddleocrVersion: string | null;
  runtimePath: string;
  storageConfigPath: string;
  storageSource: "configured" | "legacy" | "default" | string;
  storageManaged: boolean;
  storageAvailableBytes: number | null;
  storagePersistentAcrossUninstall: boolean;
  runtimeBundleAvailable: boolean;
  offlineBundleAvailable: boolean;
  installedModels: string[];
  damagedModels: string[];
  modelCatalogAvailable: boolean;
  defaultModel: string;
  message: string;
}

export function resolveAvailableOcrModel(
  runtime: Pick<OcrRuntimeStatus, "installedModels" | "defaultModel">,
  requested: OcrModelName,
): OcrModelName {
  const installed = new Set(runtime.installedModels);
  if (installed.has(requested)) return requested;
  if (installed.has(runtime.defaultModel)) {
    return runtime.defaultModel as OcrModelName;
  }
  const fallback = OCR_MODELS.find((item) => installed.has(item.id));
  return fallback?.id ?? requested;
}

export type OcrInstallState =
  | "notInstalled"
  | "installing"
  | "installFailed"
  | "dependenciesInstalled"
  | "verifying"
  | "verificationFailed"
  | "complete"
  | "cancelled";

export interface OcrInstallProgress {
  stage: string;
  state: OcrInstallState;
  percent: number;
  message: string;
  detail: string | null;
  error: string | null;
  logPath: string | null;
}

export interface OcrInstallStatus {
  schemaVersion: number;
  state: OcrInstallState;
  currentStep: string | null;
  completedSteps: string[];
  percent: number;
  message: string;
  detail: string | null;
  error: string | null;
  logPath: string;
  updatedAtMs: number;
}

export interface OcrModelCatalogEntry {
  model: OcrModelName;
  url: string;
  size: number;
  sha256: string;
}

export interface OcrModelCatalog {
  schemaVersion: number;
  platform: "windows";
  architecture: "x64";
  entries: OcrModelCatalogEntry[];
}

export type OcrModelDownloadState =
  | "idle"
  | "downloading"
  | "verifying"
  | "installing"
  | "complete"
  | "cancelled"
  | "failed";

export interface OcrModelDownloadSnapshot {
  model: OcrModelName;
  state: OcrModelDownloadState;
  downloadedBytes: number;
  totalBytes: number;
  percent: number;
  speedBytesPerSecond: number;
  etaSeconds: number | null;
  message: string;
  error: string | null;
}

export interface OcrFormulaResult {
  latex: string;
}

export interface OcrRecognitionProgress {
  event: "progress";
  id: string;
  stage: "preprocess" | "model" | "inference" | string;
  message: string;
  model: OcrModelName;
}

export type OcrProviderId =
  | "local"
  | "openai-compatible"
  | "ollama"
  | "mathpix"
  | "paddleocr";

export type OpenAiCompatibleProtocol = "responses" | "chat-completions";
export type PaddleOcrApiModel =
  | "PaddleOCR-VL-1.6";

export const PADDLE_OCR_API_MODELS: readonly PaddleOcrApiModel[] = [
  "PaddleOCR-VL-1.6",
] as const;

export interface OcrProviderConfiguration {
  activeProvider: OcrProviderId;
  openAiCompatible: {
    protocol: OpenAiCompatibleProtocol;
    baseUrl: string;
    model: string;
    prompt: string;
    hasApiKey: boolean;
  };
  ollama: {
    baseUrl: string;
    model: string;
    prompt: string;
  };
  mathpix: {
    baseUrl: string;
    appId: string;
    hasAppKey: boolean;
  };
  paddleOcr: {
    model: PaddleOcrApiModel;
    hasAccessToken: boolean;
  };
}

export interface OcrProviderConfigurationUpdate {
  activeProvider: OcrProviderId;
  openAiCompatible: {
    protocol: OpenAiCompatibleProtocol;
    baseUrl: string;
    model: string;
    prompt: string;
    apiKey?: string;
    clearApiKey?: boolean;
  };
  ollama: {
    baseUrl: string;
    model: string;
    prompt: string;
  };
  mathpix: {
    baseUrl: string;
    appId: string;
    appKey?: string;
    clearAppKey?: boolean;
  };
  paddleOcr: {
    model: PaddleOcrApiModel;
    accessToken?: string;
    clearAccessToken?: boolean;
  };
}

export interface OcrRecognitionResult {
  provider: OcrProviderId | string;
  model: string;
  elapsedMs: number;
  processedWidth: number;
  processedHeight: number;
  backgroundInverted: boolean;
  backgroundLuminance: number;
  formulas: OcrFormulaResult[];
}

const OCR_MULTILINE_ENVIRONMENTS = new Set([
  "align",
  "align*",
  "aligned",
  "alignedat",
  "gather",
  "gather*",
  "gathered",
  "split",
  "multline",
  "multline*",
]);

const OCR_TRANSPARENT_DISPLAY_ENVIRONMENTS = new Set([
  "equation",
  "equation*",
  "displaymath",
]);

function isEscapedLatexCharacter(value: string, index: number) {
  let slashCount = 0;
  for (let cursor = index - 1; cursor >= 0 && value[cursor] === "\\"; cursor -= 1) {
    slashCount += 1;
  }
  return slashCount % 2 === 1;
}

function findUnescapedDollar(value: string, startIndex = 0) {
  for (let index = Math.max(0, startIndex); index < value.length; index += 1) {
    if (value[index] !== "$" || isEscapedLatexCharacter(value, index)) continue;
    return index;
  }
  return -1;
}

function stripOcrOuterMathDelimiter(value: string) {
  const trimmed = value.trim();
  const pairs = [
    ["$$", "$$"],
    ["\\[", "\\]"],
    ["\\(", "\\)"],
    ["$", "$"],
  ] as const;
  for (const [opening, closing] of pairs) {
    if (
      !trimmed.startsWith(opening) ||
      !trimmed.endsWith(closing) ||
      trimmed.length <= opening.length + closing.length
    ) {
      continue;
    }
    const inner = trimmed.slice(opening.length, -closing.length);
    if (
      opening.includes("$") &&
      findUnescapedDollar(inner, 0) >= 0
    ) {
      // Multiple independently delimited rows must be split first; do not treat
      // the first opening dollar and last closing dollar as one giant formula.
      continue;
    }
    return inner.trim();
  }
  return trimmed;
}

function readOcrOuterEnvironment(value: string) {
  const match = value
    .trim()
    .match(/^\\begin\{([^{}]+)\}([\s\S]*)\\end\{\1\}$/);
  if (!match) return null;
  return { name: match[1].trim(), body: match[2].trim() };
}

function encodeTopLevelOcrAlignmentMarkers(value: string) {
  let result = "";
  let braceDepth = 0;
  const environments: string[] = [];
  for (let index = 0; index < value.length; index += 1) {
    const rest = value.slice(index);
    const environment = rest.match(/^\\(begin|end)\{([^{}]+)\}/);
    if (environment) {
      const token = environment[0];
      const name = environment[2];
      result += token;
      if (environment[1] === "begin") environments.push(name);
      else {
        const position = environments.lastIndexOf(name);
        if (position >= 0) environments.splice(position, 1);
      }
      index += token.length - 1;
      continue;
    }
    const character = value[index];
    if (character === "{" && !isEscapedLatexCharacter(value, index)) braceDepth += 1;
    else if (character === "}" && !isEscapedLatexCharacter(value, index)) {
      braceDepth = Math.max(0, braceDepth - 1);
    }
    if (
      character === "&" &&
      !isEscapedLatexCharacter(value, index) &&
      braceDepth === 0 &&
      environments.length === 0
    ) {
      result += VISUALTEX_ALIGNMENT_MARKER_LATEX;
      continue;
    }
    result += character;
  }
  return result.trim();
}

function splitTopLevelOcrRows(value: string) {
  const rows: string[] = [];
  const environments: string[] = [];
  let braceDepth = 0;
  let current = "";

  const flush = () => {
    const row = encodeTopLevelOcrAlignmentMarkers(current);
    if (row) rows.push(row);
    current = "";
  };

  for (let index = 0; index < value.length; index += 1) {
    const rest = value.slice(index);
    const environment = rest.match(/^\\(begin|end)\{([^{}]+)\}/);
    if (environment) {
      const token = environment[0];
      const name = environment[2];
      current += token;
      if (environment[1] === "begin") environments.push(name);
      else {
        const position = environments.lastIndexOf(name);
        if (position >= 0) environments.splice(position, 1);
      }
      index += token.length - 1;
      continue;
    }

    const character = value[index];
    if (character === "{" && !isEscapedLatexCharacter(value, index)) {
      braceDepth += 1;
      current += character;
      continue;
    }
    if (character === "}" && !isEscapedLatexCharacter(value, index)) {
      braceDepth = Math.max(0, braceDepth - 1);
      current += character;
      continue;
    }

    const atTopLevel = braceDepth === 0 && environments.length === 0;
    if (
      atTopLevel &&
      character === "\\" &&
      value[index + 1] === "\\" &&
      !isEscapedLatexCharacter(value, index)
    ) {
      flush();
      index += 1;
      // TeX permits an optional row-spacing argument after \\; it is layout
      // metadata and must not become part of the next independently editable row.
      let cursor = index + 1;
      while (cursor < value.length && /\s/.test(value[cursor])) cursor += 1;
      if (value[cursor] === "[") {
        let depth = 1;
        cursor += 1;
        while (cursor < value.length && depth > 0) {
          if (value[cursor] === "[" && !isEscapedLatexCharacter(value, cursor)) depth += 1;
          else if (value[cursor] === "]" && !isEscapedLatexCharacter(value, cursor)) depth -= 1;
          cursor += 1;
        }
      }
      index = cursor - 1;
      continue;
    }
    if (character === "\n" || character === "\r") {
      if (atTopLevel) flush();
      else if (current && !/\s$/.test(current)) current += " ";
      if (character === "\r" && value[index + 1] === "\n") index += 1;
      continue;
    }
    current += character;
  }
  flush();
  return rows;
}

export function splitOcrLatexIntoFormulaLines(value: string): string[] {
  let current = String(value ?? "").replace(/\r\n?/g, "\n").trim();
  if (!current) return [];

  for (let attempt = 0; attempt < 4; attempt += 1) {
    const withoutDelimiter = stripOcrOuterMathDelimiter(current);
    if (withoutDelimiter !== current) {
      current = withoutDelimiter;
      continue;
    }
    const environment = readOcrOuterEnvironment(current);
    if (!environment) break;
    if (OCR_TRANSPARENT_DISPLAY_ENVIRONMENTS.has(environment.name)) {
      current = environment.body;
      continue;
    }
    if (OCR_MULTILINE_ENVIRONMENTS.has(environment.name)) {
      return splitTopLevelOcrRows(environment.body);
    }
    break;
  }

  return splitTopLevelOcrRows(current)
    .map((row) => stripOcrOuterMathDelimiter(row))
    .filter(Boolean);
}

export function normalizeOcrFormulaLines(
  formulas: readonly (OcrFormulaResult | string)[],
): string[] {
  return formulas.flatMap((formula) =>
    splitOcrLatexIntoFormulaLines(
      typeof formula === "string" ? formula : formula.latex,
    ),
  );
}

export function normalizeOcrFormulaText(value: string): string[] {
  return splitOcrLatexIntoFormulaLines(value);
}

export interface OcrImageRequest {
  bytes: number[];
  extension: string;
  model: OcrModelName;
}

const SUPPORTED_EXTENSIONS = new Set([
  "png",
  "jpg",
  "jpeg",
  "webp",
  "bmp",
  "tif",
  "tiff",
]);

export const isTauriEnvironment = () =>
  configuredTransport?.environment === "desktop";

export const isOfficeCompanionEnvironment = () => {
  if (typeof window === "undefined" || typeof document === "undefined") {
    return false;
  }
  const token =
    window.__VISUALTEX_INSTALL_TOKEN__ ??
    document
      .querySelector<HTMLMetaElement>('meta[name="visualtex-install-token"]')
      ?.content;
  return (
    configuredTransport?.environment === "office" &&
    window.location.protocol === "https:" &&
    window.location.hostname === "127.0.0.1" &&
    window.location.port === "43127" &&
    typeof token === "string" &&
    token.length >= 32
  );
};

export function getImageExtension(file: File): string {
  const fromName = file.name.split(".").pop()?.toLocaleLowerCase() ?? "";
  if (SUPPORTED_EXTENSIONS.has(fromName)) return fromName;

  const mimeMap: Record<string, string> = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/webp": "webp",
    "image/bmp": "bmp",
    "image/tiff": "tiff",
  };
  const fromMime = mimeMap[file.type];
  if (fromMime) return fromMime;
  throw new Error("不支持该图片格式，请使用 PNG、JPEG、WebP、BMP 或 TIFF");
}

export function validateOcrImage(file: File) {
  getImageExtension(file);
  if (file.size <= 0) throw new Error("图片文件为空");
  if (file.size > 20 * 1024 * 1024) {
    throw new Error("图片不能超过 20 MB");
  }
}

export async function fileToOcrRequest(
  file: File,
  model: OcrModelName,
): Promise<OcrImageRequest> {
  validateOcrImage(file);
  const buffer = await file.arrayBuffer();
  return {
    bytes: Array.from(new Uint8Array(buffer)),
    extension: getImageExtension(file),
    model,
  };
}

function requireOcrEnvironment() {
  if (!isTauriEnvironment() && !isOfficeCompanionEnvironment()) {
    throw new Error(
      "OCR 只在 VisualTeX 桌面应用或本地 Office 编辑器中可用。",
    );
  }
}

function requireDesktopOcrEnvironment() {
  if (!isTauriEnvironment()) {
    throw new Error("可选 OCR 模型包只能在 VisualTeX 桌面应用中管理。");
  }
}

export async function getOcrProviderConfiguration(): Promise<OcrProviderConfiguration> {
  requireOcrEnvironment();
  return decodeOcrProviderConfiguration(
    await invoke<unknown>("get_ocr_provider_configuration"),
  );
}

export async function saveOcrProviderConfiguration(
  configuration: OcrProviderConfigurationUpdate,
): Promise<OcrProviderConfiguration> {
  requireOcrEnvironment();
  return decodeOcrProviderConfiguration(
    await invoke<unknown>("save_ocr_provider_configuration", { configuration }),
  );
}

export async function getOcrRuntimeStatus(
  forceRefresh = false,
): Promise<OcrRuntimeStatus> {
  requireOcrEnvironment();
  return decodeOcrRuntimeStatus(
    await invoke<unknown>("get_ocr_runtime_status", { forceRefresh }),
  );
}

export async function configureOcrStorageLocation(
  selectedDirectory: string,
): Promise<OcrRuntimeStatus> {
  requireDesktopOcrEnvironment();
  return decodeOcrRuntimeStatus(
    await invoke<unknown>("configure_ocr_storage_location", { selectedDirectory }),
  );
}

export async function openOcrStorageLocation(): Promise<void> {
  requireDesktopOcrEnvironment();
  return invoke("open_ocr_storage_location");
}

export async function installOcrRuntime(): Promise<OcrRuntimeStatus> {
  requireOcrEnvironment();
  return decodeOcrRuntimeStatus(await invoke<unknown>("install_ocr_runtime"));
}

export async function getOcrInstallStatus(): Promise<OcrInstallStatus> {
  requireOcrEnvironment();
  return decodeOcrInstallStatus(await invoke<unknown>("get_ocr_install_status"));
}

export async function cancelOcrInstall(): Promise<void> {
  requireOcrEnvironment();
  return invoke("cancel_ocr_install");
}

export async function openOcrInstallLogs(): Promise<void> {
  requireOcrEnvironment();
  return invoke("open_ocr_install_logs");
}

export async function recognizeFormulaImage(
  request: OcrImageRequest,
): Promise<OcrRecognitionResult> {
  requireOcrEnvironment();
  return decodeOcrRecognitionResult(
    await invoke<unknown>("recognize_formula_image", { request }),
  );
}

export async function cancelOcrRecognition(): Promise<void> {
  requireOcrEnvironment();
  return invoke("cancel_ocr_recognition");
}

export async function restartOcrWorker(): Promise<void> {
  requireOcrEnvironment();
  return invoke("restart_ocr_worker");
}

export async function warmupOcrModel(model: OcrModelName): Promise<void> {
  requireOcrEnvironment();
  return invoke("warmup_ocr_model", { model });
}

export async function resetOcrRuntime(): Promise<OcrRuntimeStatus> {
  requireOcrEnvironment();
  return decodeOcrRuntimeStatus(await invoke<unknown>("reset_ocr_runtime"));
}

export async function installOptionalOcrModel(
  packagePath: string,
): Promise<OcrRuntimeStatus> {
  requireDesktopOcrEnvironment();
  return decodeOcrRuntimeStatus(
    await invoke<unknown>("install_optional_ocr_model", { packagePath }),
  );
}

export async function removeOptionalOcrModel(
  model: OcrModelName,
): Promise<OcrRuntimeStatus> {
  requireDesktopOcrEnvironment();
  return decodeOcrRuntimeStatus(
    await invoke<unknown>("remove_optional_ocr_model", { model }),
  );
}

export async function getOcrModelCatalog(): Promise<OcrModelCatalog> {
  requireDesktopOcrEnvironment();
  return decodeOcrModelCatalog(await invoke<unknown>("get_ocr_model_catalog"));
}

export async function getOcrModelDownloadStatus(): Promise<OcrModelDownloadSnapshot | null> {
  requireDesktopOcrEnvironment();
  return decodeNullableOcrModelDownloadSnapshot(
    await invoke<unknown>("get_ocr_model_download_status"),
  );
}

export async function downloadOcrModel(
  model: OcrModelName,
): Promise<OcrRuntimeStatus> {
  requireDesktopOcrEnvironment();
  return decodeOcrRuntimeStatus(
    await invoke<unknown>("download_ocr_model", { model }),
  );
}

export async function cancelOcrModelDownload(): Promise<boolean> {
  requireDesktopOcrEnvironment();
  return decodeBooleanOcrResult(
    await invoke<unknown>("cancel_ocr_model_download"),
    "cancelOcrModelDownload",
  );
}

export async function listenOcrModelDownloadProgress(
  listener: (progress: OcrModelDownloadSnapshot) => void,
): Promise<UnlistenFn> {
  requireDesktopOcrEnvironment();
  return listen<unknown>("ocr-model-download-progress", (event) => {
    listener(decodeOcrModelDownloadSnapshot(event.payload));
  });
}

export async function listenOcrRecognitionProgress(
  listener: (progress: OcrRecognitionProgress) => void,
): Promise<UnlistenFn> {
  requireOcrEnvironment();
  return listen<unknown>("ocr-recognition-progress", (event) => {
    listener(decodeOcrRecognitionProgress(event.payload));
  });
}

export async function listenOcrInstallProgress(
  listener: (progress: OcrInstallProgress) => void,
): Promise<UnlistenFn> {
  requireOcrEnvironment();
  return listen<unknown>("ocr-install-progress", (event) => {
    listener(decodeOcrInstallProgress(event.payload));
  });
}
