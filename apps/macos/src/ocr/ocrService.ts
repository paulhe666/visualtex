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
    hintZh: "可选离线模型包，安装后约 248 MB；速度最快，主要适合英文公式",
    hintEn: "Optional offline model pack, about 248 MB installed; fastest for English formulas",
    downloadMb: 259.6,
    storageMb: 248,
    cpuBenchmarkMs: 260.99,
  },
  {
    id: "PP-FormulaNet_plus-M",
    labelZh: "均衡版 M（推荐）",
    labelEn: "Balanced M (recommended)",
    hintZh: "随 VisualTeX 离线资源内置；兼顾中文、复杂公式与速度",
    hintEn: "Included in the VisualTeX offline bundle; balanced for Chinese and complex formulas",
    downloadMb: 620.5,
    storageMb: 592,
    cpuBenchmarkMs: 1615.8,
  },
  {
    id: "PP-FormulaNet_plus-L",
    labelZh: "高精度版 L",
    labelEn: "High accuracy L",
    hintZh: "可选离线模型包，安装后约 698 MB；首次加载较久，并会占用数 GB 内存",
    hintEn: "Optional offline model pack, about 698 MB installed; first load is slow and may use several GB of memory",
    downloadMb: 731.5,
    storageMb: 698,
    cpuBenchmarkMs: 3125.58,
  },
] as const;

export type OcrModelName = (typeof OCR_MODELS)[number]["id"];

export interface OcrRuntimeStatus {
  installed: boolean;
  pythonPath: string | null;
  pythonVersion: string | null;
  paddleVersion: string | null;
  paddleocrVersion: string | null;
  runtimePath: string;
  offlineBundleAvailable: boolean;
  installedModels: string[];
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

export interface OcrInstallProgress {
  stage: string;
  percent: number;
  message: string;
  detail: string | null;
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

export interface OcrRecognitionResult {
  model: string;
  elapsedMs: number;
  processedWidth: number;
  processedHeight: number;
  backgroundInverted: boolean;
  backgroundLuminance: number;
  formulas: OcrFormulaResult[];
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

export async function getOcrRuntimeStatus(
  forceRefresh = false,
): Promise<OcrRuntimeStatus> {
  requireOcrEnvironment();
  return invoke<OcrRuntimeStatus>("get_ocr_runtime_status", { forceRefresh });
}

export async function installOcrRuntime(): Promise<OcrRuntimeStatus> {
  requireOcrEnvironment();
  return invoke<OcrRuntimeStatus>("install_ocr_runtime");
}

export async function recognizeFormulaImage(
  request: OcrImageRequest,
): Promise<OcrRecognitionResult> {
  requireOcrEnvironment();
  return invoke<OcrRecognitionResult>("recognize_formula_image", { request });
}

export async function prewarmOcrModel(model: OcrModelName): Promise<void> {
  requireOcrEnvironment();
  return invoke("prewarm_ocr_model", { model });
}

export async function cancelOcrRecognition(): Promise<void> {
  requireOcrEnvironment();
  return invoke("cancel_ocr_recognition");
}

export async function restartOcrWorker(): Promise<void> {
  requireOcrEnvironment();
  return invoke("restart_ocr_worker");
}

export async function resetOcrRuntime(): Promise<OcrRuntimeStatus> {
  requireOcrEnvironment();
  return invoke<OcrRuntimeStatus>("reset_ocr_runtime");
}

export async function installOptionalOcrModel(
  packagePath: string,
): Promise<OcrRuntimeStatus> {
  requireDesktopOcrEnvironment();
  return invoke<OcrRuntimeStatus>("install_optional_ocr_model", {
    packagePath,
  });
}

export async function removeOptionalOcrModel(
  model: OcrModelName,
): Promise<OcrRuntimeStatus> {
  requireDesktopOcrEnvironment();
  return invoke<OcrRuntimeStatus>("remove_optional_ocr_model", { model });
}

export async function listenOcrRecognitionProgress(
  listener: (progress: OcrRecognitionProgress) => void,
): Promise<UnlistenFn> {
  requireOcrEnvironment();
  return listen<OcrRecognitionProgress>("ocr-recognition-progress", (event) => {
    listener(event.payload);
  });
}

export async function listenOcrInstallProgress(
  listener: (progress: OcrInstallProgress) => void,
): Promise<UnlistenFn> {
  requireOcrEnvironment();
  return listen<OcrInstallProgress>("ocr-install-progress", (event) => {
    listener(event.payload);
  });
}
