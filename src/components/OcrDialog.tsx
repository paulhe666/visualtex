import { getCurrentWebview } from "@tauri-apps/api/webview";
import { open as openDialog } from "@tauri-apps/plugin-dialog";
import {
  AlertCircle,
  Check,
  CheckCircle2,
  ClipboardPaste,
  Cloud,
  Copy,
  Cpu,
  Download,
  ImagePlus,
  KeyRound,
  LoaderCircle,
  Plus,
  RefreshCw,
  Save,
  ScanLine,
  Server,
  Trash2,
  Upload,
  X,
} from "lucide-react";
import {
  type ChangeEvent,
  type DragEvent,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { MathPreview } from "./MathPreview";
import {
  DEFAULT_OCR_MODEL,
  OCR_MODELS,
  PADDLE_OCR_API_MODELS,
  SIMPLETEX_API_MODELS,
  cancelOcrRecognition,
  type OcrInstallProgress,
  type OcrModelName,
  type OcrProviderConfiguration,
  type OcrProviderId,
  type PaddleOcrApiModel,
  type SimpleTexApiModel,
  type OcrRecognitionProgress,
  type OcrRecognitionResult,
  type OcrRuntimeStatus,
  fileToOcrRequest,
  formatOcrError,
  getOcrProviderConfiguration,
  getOcrRuntimeStatus,
  installOcrRuntime,
  installOptionalOcrModel,
  isOfficeCompanionEnvironment,
  isTauriEnvironment,
  listenOcrInstallProgress,
  listenOcrRecognitionProgress,
  normalizeOcrFormulaLines,
  normalizeOcrFormulaText,
  recognizeFormulaImage,
  resolveAvailableOcrModel,
  prewarmOcrModel,
  removeOptionalOcrModel,
  resetOcrRuntime,
  restartOcrWorker,
  saveOcrProviderConfiguration,
  validateOcrImage,
} from "../ocr/ocrService";

interface OcrDialogProps {
  open: boolean;
  language: "cn" | "en";
  model: OcrModelName;
  onModelChange: (model: OcrModelName) => void;
  onClose: () => void;
  onInsert: (latex: string) => void;
  onAppend: (latex: string) => void;
  onNotify: (message: string) => void;
  onProviderConfigurationChange?: (
    configuration: OcrProviderConfiguration,
  ) => void;
}

function readableBytes(bytes: number) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  if (bytes < 1024 * 1024 * 1024) return (bytes / 1024 / 1024).toFixed(1) + " MB";
  return (bytes / 1024 / 1024 / 1024).toFixed(2) + " GB";
}

function readError(error: unknown) {
  return formatOcrError(error);
}

function credentialStatusLabel(status: boolean | null, isEn: boolean, clearing: boolean) {
  if (clearing || status === false) return "";
  if (status === true) return isEn ? " · saved" : " · 已保存";
  return isEn ? " · not checked" : " · 未确认";
}

const OCR_MODEL_PACKAGE_EXTENSION = ".vtxocrmodel";

const FALLBACK_OCR_PROVIDER_CONFIGURATION: OcrProviderConfiguration = {
  activeProvider: "local",
  openAiCompatible: {
    protocol: "responses",
    baseUrl: "https://api.openai.com/v1",
    model: "",
    prompt:
      'Read every mathematical formula in this image in visual order. Return JSON only in the exact form {"formulas":[{"latex":"..."}]}. Return each independent visual formula row as a separate formulas-array item; keep matrices and cases as one item. Use valid LaTeX without markdown fences or surrounding dollar delimiters.',
    hasApiKey: false,
  },
  ollama: {
    baseUrl: "http://127.0.0.1:11434",
    model: "",
    prompt:
      'Read every mathematical formula in this image in visual order. Return JSON only in the exact form {"formulas":[{"latex":"..."}]}. Return each independent visual formula row as a separate formulas-array item; keep matrices and cases as one item. Use valid LaTeX without markdown fences or surrounding dollar delimiters.',
  },
  mathpix: {
    baseUrl: "https://api.mathpix.com",
    appId: "",
    hasAppKey: false,
  },
  paddleOcr: {
    model: "PaddleOCR-VL-1.6",
    hasAccessToken: false,
  },
  simpleTex: {
    model: "standard",
    hasAccessToken: false,
  },
};

function isOcrModelPackagePath(path: string) {
  return path.trim().toLowerCase().endsWith(OCR_MODEL_PACKAGE_EXTENSION);
}

function modelNameFromPackagePath(path: string): OcrModelName | null {
  const normalized = path.replaceAll("\\", "/").toLowerCase();
  return (
    OCR_MODELS.find((candidate) =>
      normalized.includes(candidate.id.toLowerCase()),
    )?.id ?? null
  );
}

function isDropPositionInsideElement(
  position: { x: number; y: number },
  element: HTMLElement | null,
) {
  if (!element) return false;
  const scale = window.devicePixelRatio || 1;
  const x = position.x / scale;
  const y = position.y / scale;
  const bounds = element.getBoundingClientRect();
  return x >= bounds.left && x <= bounds.right && y >= bounds.top && y <= bounds.bottom;
}

function normalizeResultLatex(value: string) {
  return normalizeOcrFormulaText(value).join("\n");
}

function providerLabel(provider: OcrProviderId, isEn: boolean) {
  switch (provider) {
    case "openai-compatible":
      return isEn ? "OpenAI-compatible API" : "OpenAI 兼容 API";
    case "ollama":
      return "Ollama";
    case "mathpix":
      return "Mathpix";
    case "paddleocr":
      return isEn ? "PaddleOCR AI Studio" : "PaddleOCR 星河 API";
    case "simpletex":
      return "SimpleTex";
    default:
      return isEn ? "Local PP-FormulaNet" : "本地 PP-FormulaNet";
  }
}

export function OcrDialog({
  open,
  language,
  model,
  onModelChange,
  onClose,
  onInsert,
  onAppend,
  onNotify,
  onProviderConfigurationChange,
}: OcrDialogProps) {
  const isEn = language === "en";
  const dialogRef = useRef<HTMLElement>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const modelDropZoneRef = useRef<HTMLDivElement>(null);
  const modelDragPathsRef = useRef<string[]>([]);
  const recognizingRef = useRef(false);
  const cancellingRef = useRef(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const objectUrlRef = useRef<string | null>(null);
  const [runtime, setRuntime] = useState<OcrRuntimeStatus | null>(null);
  const [providerConfiguration, setProviderConfiguration] =
    useState<OcrProviderConfiguration>(FALLBACK_OCR_PROVIDER_CONFIGURATION);
  const [providerLoaded, setProviderLoaded] = useState(false);
  const [providerDirty, setProviderDirty] = useState(false);
  const [savingProvider, setSavingProvider] = useState(false);
  const [openAiApiKey, setOpenAiApiKey] = useState("");
  const [clearOpenAiApiKey, setClearOpenAiApiKey] = useState(false);
  const [mathpixAppKey, setMathpixAppKey] = useState("");
  const [clearMathpixAppKey, setClearMathpixAppKey] = useState(false);
  const [paddleOcrAccessToken, setPaddleOcrAccessToken] = useState("");
  const [clearPaddleOcrAccessToken, setClearPaddleOcrAccessToken] = useState(false);
  const [simpleTexAccessToken, setSimpleTexAccessToken] = useState("");
  const [clearSimpleTexAccessToken, setClearSimpleTexAccessToken] = useState(false);
  const [modelBusy, setModelBusy] = useState(false);
  const [modelPackageDragging, setModelPackageDragging] = useState(false);
  const [checkingRuntime, setCheckingRuntime] = useState(false);
  const [installing, setInstalling] = useState(false);
  const [installProgress, setInstallProgress] = useState<OcrInstallProgress | null>(null);
  const [file, setFile] = useState<File | null>(null);
  const [previewUrl, setPreviewUrl] = useState("");
  const [imageSize, setImageSize] = useState({ width: 0, height: 0 });
  const [dragging, setDragging] = useState(false);
  const [recognizing, setRecognizing] = useState(false);
  const [cancelling, setCancelling] = useState(false);
  const [recognitionSeconds, setRecognitionSeconds] = useState(0);
  const [recognitionProgress, setRecognitionProgress] =
    useState<OcrRecognitionProgress | null>(null);
  const [result, setResult] = useState<OcrRecognitionResult | null>(null);
  const [latex, setLatex] = useState("");
  const [error, setError] = useState("");
  const [copied, setCopied] = useState(false);

  const selectedModel = useMemo(
    () =>
      OCR_MODELS.find((item) => item.id === model) ??
      OCR_MODELS.find((item) => item.id === DEFAULT_OCR_MODEL)!,
    [model],
  );
  const defaultModel = runtime?.defaultModel ?? DEFAULT_OCR_MODEL;
  const installedModels = runtime?.installedModels ?? [];
  const selectedModelInstalled = installedModels.includes(model);
  const selectedModelRemovable =
    selectedModelInstalled && model !== defaultModel && isTauriEnvironment();
  const activeProvider = providerConfiguration.activeProvider;
  const usingLocalProvider = activeProvider === "local";
  const activeProviderLabel = providerLabel(activeProvider, isEn);

  const clearObjectUrl = useCallback(() => {
    if (objectUrlRef.current) {
      URL.revokeObjectURL(objectUrlRef.current);
      objectUrlRef.current = null;
    }
  }, []);

  const refreshRuntime = useCallback(async (forceRefresh = false) => {
    if (!isTauriEnvironment() && !isOfficeCompanionEnvironment()) {
      setRuntime({
        installed: false,
        pythonPath: null,
        pythonVersion: null,
        paddleVersion: null,
        paddleocrVersion: null,
        runtimePath: "",
        offlineBundleAvailable: false,
        installedModels: [],
        defaultModel: DEFAULT_OCR_MODEL,
        message: isEn
          ? "OCR is only available in the desktop app."
          : "OCR 仅在桌面应用中可用。",
      });
      return;
    }

    setCheckingRuntime(true);
    try {
      setRuntime(await getOcrRuntimeStatus(forceRefresh));
    } catch (runtimeError) {
      setError(readError(runtimeError));
    } finally {
      setCheckingRuntime(false);
    }
  }, [isEn]);

  const refreshProviderConfiguration = useCallback(async () => {
    if (!isTauriEnvironment() && !isOfficeCompanionEnvironment()) {
      setProviderConfiguration(FALLBACK_OCR_PROVIDER_CONFIGURATION);
      setProviderLoaded(true);
      setProviderDirty(false);
      return;
    }
    try {
      const configuration = await getOcrProviderConfiguration();
      setProviderConfiguration(configuration);
      onProviderConfigurationChange?.(configuration);
      setOpenAiApiKey("");
      setClearOpenAiApiKey(false);
      setMathpixAppKey("");
      setClearMathpixAppKey(false);
      setPaddleOcrAccessToken("");
      setClearPaddleOcrAccessToken(false);
      setSimpleTexAccessToken("");
      setClearSimpleTexAccessToken(false);
      setProviderDirty(false);
    } catch (providerError) {
      setProviderConfiguration(FALLBACK_OCR_PROVIDER_CONFIGURATION);
      setOpenAiApiKey("");
      setClearOpenAiApiKey(false);
      setMathpixAppKey("");
      setClearMathpixAppKey(false);
      setPaddleOcrAccessToken("");
      setClearPaddleOcrAccessToken(false);
      setSimpleTexAccessToken("");
      setClearSimpleTexAccessToken(false);
      setProviderDirty(true);
      setError(readError(providerError));
    } finally {
      setProviderLoaded(true);
    }
  }, [onProviderConfigurationChange]);

  const updateProviderConfiguration = useCallback(
    (update: (current: OcrProviderConfiguration) => OcrProviderConfiguration) => {
      setProviderConfiguration(update);
      setProviderDirty(true);
    },
    [],
  );

  const handleSaveProviderConfiguration = useCallback(async () => {
    if (!isTauriEnvironment() && !isOfficeCompanionEnvironment()) return;
    setSavingProvider(true);
    setError("");
    try {
      const saved = await saveOcrProviderConfiguration({
        activeProvider: providerConfiguration.activeProvider,
        openAiCompatible: {
          protocol: providerConfiguration.openAiCompatible.protocol,
          baseUrl: providerConfiguration.openAiCompatible.baseUrl,
          model: providerConfiguration.openAiCompatible.model,
          prompt: providerConfiguration.openAiCompatible.prompt,
          apiKey: openAiApiKey || undefined,
          clearApiKey: clearOpenAiApiKey,
        },
        ollama: {
          baseUrl: providerConfiguration.ollama.baseUrl,
          model: providerConfiguration.ollama.model,
          prompt: providerConfiguration.ollama.prompt,
        },
        mathpix: {
          baseUrl: providerConfiguration.mathpix.baseUrl,
          appId: providerConfiguration.mathpix.appId,
          appKey: mathpixAppKey || undefined,
          clearAppKey: clearMathpixAppKey,
        },
        paddleOcr: {
          model: providerConfiguration.paddleOcr.model,
          accessToken: paddleOcrAccessToken || undefined,
          clearAccessToken: clearPaddleOcrAccessToken,
        },
        simpleTex: {
          model: providerConfiguration.simpleTex.model,
          accessToken: simpleTexAccessToken || undefined,
          clearAccessToken: clearSimpleTexAccessToken,
        },
      });
      setProviderConfiguration(saved);
      onProviderConfigurationChange?.(saved);
      setOpenAiApiKey("");
      setClearOpenAiApiKey(false);
      setMathpixAppKey("");
      setClearMathpixAppKey(false);
      setPaddleOcrAccessToken("");
      setClearPaddleOcrAccessToken(false);
      setSimpleTexAccessToken("");
      setClearSimpleTexAccessToken(false);
      setProviderDirty(false);
      onNotify(
        isEn
          ? `OCR provider saved: ${providerLabel(saved.activeProvider, true)}`
          : `OCR 提供器已保存：${providerLabel(saved.activeProvider, false)}`,
      );
    } catch (providerError) {
      setError(readError(providerError));
    } finally {
      setSavingProvider(false);
    }
  }, [
    clearMathpixAppKey,
    clearOpenAiApiKey,
    clearPaddleOcrAccessToken,
    clearSimpleTexAccessToken,
    isEn,
    mathpixAppKey,
    onNotify,
    openAiApiKey,
    paddleOcrAccessToken,
    providerConfiguration,
    simpleTexAccessToken,
  ]);

  const importModelPackage = useCallback(
    async (packagePath: string) => {
      if (!isTauriEnvironment() || modelBusy) return;
      if (!isOcrModelPackagePath(packagePath)) {
        setError(
          isEn
            ? "Choose a .vtxocrmodel package."
            : "请选择 .vtxocrmodel 模型包。",
        );
        return;
      }

      const previouslyInstalled = new Set(runtime?.installedModels ?? []);
      const packageModel = modelNameFromPackagePath(packagePath);
      setModelBusy(true);
      setError("");
      try {
        const nextRuntime = await installOptionalOcrModel(packagePath);
        setRuntime(nextRuntime);
        const newlyInstalled = nextRuntime.installedModels.find(
          (candidate) => !previouslyInstalled.has(candidate),
        ) as OcrModelName | undefined;
        const imported =
          newlyInstalled ??
          (packageModel && nextRuntime.installedModels.includes(packageModel)
            ? packageModel
            : undefined) ??
          (nextRuntime.installedModels.includes(model) ? model : undefined) ??
          (nextRuntime.installedModels.at(-1) as OcrModelName | undefined);
        if (imported) onModelChange(imported);
        onNotify(isEn ? "Verified OCR model imported" : "OCR 模型已校验并导入");
      } catch (importError) {
        setError(readError(importError));
      } finally {
        setModelBusy(false);
      }
    },
    [isEn, model, modelBusy, onModelChange, onNotify, runtime?.installedModels],
  );

  useEffect(() => {
    if (!open) return;
    setError("");
    let cancelled = false;
    const frame = window.requestAnimationFrame(() => {
      if (cancelled) return;
      void refreshProviderConfiguration();
      void refreshRuntime();
    });
    return () => {
      cancelled = true;
      window.cancelAnimationFrame(frame);
    };
  }, [open, refreshProviderConfiguration, refreshRuntime]);

  useEffect(() => {
    if (!open || !isTauriEnvironment()) return;
    let unlisten: (() => void) | undefined;
    let disposed = false;

    const updateDropHighlight = (
      paths: string[],
      position: { x: number; y: number },
    ) => {
      setModelPackageDragging(
        !modelBusy &&
          paths.some(isOcrModelPackagePath) &&
          isDropPositionInsideElement(position, modelDropZoneRef.current),
      );
    };

    void getCurrentWebview()
      .onDragDropEvent((event) => {
        if (disposed) return;
        const payload = event.payload;
        if (payload.type === "enter") {
          modelDragPathsRef.current = payload.paths;
          updateDropHighlight(payload.paths, payload.position);
          return;
        }
        if (payload.type === "over") {
          updateDropHighlight(modelDragPathsRef.current, payload.position);
          return;
        }
        if (payload.type === "leave") {
          modelDragPathsRef.current = [];
          setModelPackageDragging(false);
          return;
        }

        const droppedInsideModelArea = isDropPositionInsideElement(
          payload.position,
          modelDropZoneRef.current,
        );
        const packages = payload.paths.filter(isOcrModelPackagePath);
        modelDragPathsRef.current = [];
        setModelPackageDragging(false);
        if (!droppedInsideModelArea) return;
        if (packages.length !== 1) {
          setError(
            packages.length > 1
              ? isEn
                ? "Import one .vtxocrmodel package at a time."
                : "每次只能导入一个 .vtxocrmodel 模型包。"
              : isEn
                ? "Only .vtxocrmodel packages can be imported."
                : "仅支持 .vtxocrmodel 模型包。",
          );
          return;
        }
        void importModelPackage(packages[0]);
      })
      .then((dispose) => {
        if (disposed) dispose();
        else unlisten = dispose;
      })
      .catch((dragDropError) => {
        if (!disposed) setError(readError(dragDropError));
      });

    return () => {
      disposed = true;
      modelDragPathsRef.current = [];
      setModelPackageDragging(false);
      unlisten?.();
    };
  }, [importModelPackage, isEn, modelBusy, open]);

  useEffect(() => {
    if (!open) return;

    const handlePaste = (event: ClipboardEvent) => {
      const item = Array.from(event.clipboardData?.items ?? []).find((candidate) =>
        candidate.type.startsWith("image/"),
      );
      const pastedFile = item?.getAsFile();
      if (!pastedFile) return;
      event.preventDefault();
      try {
        validateOcrImage(pastedFile);
        clearObjectUrl();
        const nextUrl = URL.createObjectURL(pastedFile);
        objectUrlRef.current = nextUrl;
        setPreviewUrl(nextUrl);
        setFile(pastedFile);
        setImageSize({ width: 0, height: 0 });
        setResult(null);
        setLatex("");
        setError("");
      } catch (pasteError) {
        setError(readError(pasteError));
      }
    };

    window.addEventListener("paste", handlePaste);
    return () => window.removeEventListener("paste", handlePaste);
  }, [open, clearObjectUrl]);

  useEffect(
    () => () => {
      clearObjectUrl();
    },
    [clearObjectUrl],
  );

  useEffect(() => {
    if (!recognizing) return;
    setRecognitionSeconds(0);
    const startedAt = Date.now();
    const timer = window.setInterval(() => {
      setRecognitionSeconds(Math.floor((Date.now() - startedAt) / 1000));
    }, 1000);
    return () => window.clearInterval(timer);
  }, [recognizing]);

  useEffect(() => {
    recognizingRef.current = recognizing;
  }, [recognizing]);

  useEffect(() => {
    cancellingRef.current = cancelling;
  }, [cancelling]);

  const selectFile = useCallback(
    (nextFile: File) => {
      try {
        validateOcrImage(nextFile);
        clearObjectUrl();
        const nextUrl = URL.createObjectURL(nextFile);
        objectUrlRef.current = nextUrl;
        setPreviewUrl(nextUrl);
        setFile(nextFile);
        setImageSize({ width: 0, height: 0 });
        setResult(null);
        setLatex("");
        setError("");
      } catch (selectionError) {
        setError(readError(selectionError));
      }
    },
    [clearObjectUrl],
  );

  const handleFileInput = (event: ChangeEvent<HTMLInputElement>) => {
    const nextFile = event.target.files?.[0];
    if (nextFile) selectFile(nextFile);
    event.target.value = "";
  };

  const handleDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragging(false);
    const nextFile = Array.from(event.dataTransfer.files).find((candidate) =>
      candidate.type.startsWith("image/"),
    );
    if (nextFile) selectFile(nextFile);
    else setError(isEn ? "Drop an image file here." : "请拖入图片文件。");
  };

  const handleInstall = async () => {
    setInstalling(true);
    setError("");
    setInstallProgress({
      stage: "start",
      percent: 1,
      message: isEn ? "Starting OCR installation" : "正在启动 OCR 安装",
      detail: null,
    });

    let unlisten: (() => void) | undefined;
    try {
      unlisten = await listenOcrInstallProgress(setInstallProgress);
      const nextRuntime = await installOcrRuntime();
      setRuntime(nextRuntime);
      const availableModel = resolveAvailableOcrModel(nextRuntime, model);
      void prewarmOcrModel(availableModel).catch(() => undefined);
      onNotify(isEn ? "OCR runtime installed" : "OCR 运行环境安装完成");
    } catch (installError) {
      setError(readError(installError));
      await refreshRuntime(true);
    } finally {
      unlisten?.();
      setInstalling(false);
    }
  };

  const handleImportModel = async () => {
    if (!isTauriEnvironment() || modelBusy) return;
    try {
      const selected = await openDialog({
        multiple: false,
        directory: false,
        filters: [
          {
            name: "VisualTeX OCR model",
            extensions: ["vtxocrmodel"],
          },
        ],
      });
      if (typeof selected === "string") await importModelPackage(selected);
    } catch (importError) {
      setError(readError(importError));
    }
  };

  const handleRemoveModel = async () => {
    if (!selectedModelRemovable || modelBusy) return;
    const confirmed = window.confirm(
      isEn
        ? `Remove ${selectedModel.labelEn}?`
        : `删除${selectedModel.labelZh}？`,
    );
    if (!confirmed) return;
    setModelBusy(true);
    setError("");
    try {
      const nextRuntime = await removeOptionalOcrModel(model);
      setRuntime(nextRuntime);
      const fallback = resolveAvailableOcrModel(nextRuntime, DEFAULT_OCR_MODEL);
      onModelChange(fallback);
      onNotify(isEn ? "Optional OCR model removed" : "可选 OCR 模型已删除");
    } catch (removeError) {
      setError(readError(removeError));
    } finally {
      setModelBusy(false);
    }
  };

  const handleRecognize = async () => {
    if (!file) {
      setError(isEn ? "Choose or paste a formula image first." : "请先选择或粘贴一张公式图片。");
      return;
    }
    if (!providerLoaded) {
      setError(isEn ? "OCR provider configuration is still loading." : "OCR 提供器配置仍在加载中。");
      return;
    }
    if (providerDirty) {
      setError(
        isEn
          ? "Save the provider configuration first."
          : "请先保存提供器配置。",
      );
      return;
    }
    if (usingLocalProvider && !runtime?.installed) {
      setError(isEn ? "Install the OCR runtime first." : "请先安装 OCR 运行环境。");
      return;
    }
    if (usingLocalProvider && !selectedModelInstalled) {
      setError(
        isEn
          ? "Import the selected model package first."
          : "请先导入所选模型包。",
      );
      return;
    }

    setRecognizing(true);
    cancellingRef.current = false;
    setCancelling(false);
    setRecognitionProgress({
      event: "progress",
      id: "pending",
      stage: "preprocess",
      model,
      message: usingLocalProvider
        ? isEn
          ? "Preparing the formula image"
          : "正在准备公式图片"
        : isEn
          ? `Preparing request for ${activeProviderLabel}`
          : `正在准备发送到${activeProviderLabel}`,
    });
    setResult(null);
    setLatex("");
    setError("");

    let unlisten: (() => void) | undefined;
    try {
      unlisten = await listenOcrRecognitionProgress((progress) => {
        if (progress.model === model) setRecognitionProgress(progress);
      });
      const request = await fileToOcrRequest(file, model);
      const nextResult = await recognizeFormulaImage(request);
      setResult(nextResult);
      setLatex(normalizeOcrFormulaLines(nextResult.formulas).join("\n"));
    } catch (recognitionError) {
      const message = readError(recognitionError);
      if (cancellingRef.current || message.includes("OCR_CANCELLED")) {
        onNotify(isEn ? "OCR recognition cancelled" : "OCR 识别已取消");
      } else {
        setError(message);
      }
    } finally {
      unlisten?.();
      setRecognizing(false);
      cancellingRef.current = false;
      setCancelling(false);
      setRecognitionProgress(null);
    }
  };

  const handleCancelRecognition = async () => {
    if (!recognizing || cancelling) return;
    cancellingRef.current = true;
    setCancelling(true);
    setRecognitionProgress((current) => ({
      event: "progress",
      id: current?.id ?? "pending",
      stage: "cancelling",
      model,
      message: isEn ? "Stopping the OCR worker…" : "正在停止 OCR 进程…",
    }));
    try {
      await cancelOcrRecognition();
    } catch (cancelError) {
      setError(readError(cancelError));
      cancellingRef.current = false;
      setCancelling(false);
    }
  };

  const requestClose = () => {
    if (recognizingRef.current) void handleCancelRecognition();
    onClose();
  };

  useEffect(() => {
    if (!open) return;
    previousFocusRef.current = document.activeElement as HTMLElement | null;
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current?.querySelector<HTMLElement>("button, input, select")?.focus();
    });

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        requestClose();
        return;
      }
      if (event.key !== "Tab" || !dialogRef.current) return;

      const focusable = Array.from(
        dialogRef.current.querySelectorAll<HTMLElement>(
          'button:not(:disabled), input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])',
        ),
      );
      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (!first || !last) return;

      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    };

    document.addEventListener("keydown", handleKeyDown);
    return () => {
      window.cancelAnimationFrame(frame);
      document.removeEventListener("keydown", handleKeyDown);
      previousFocusRef.current?.focus({ preventScroll: true });
    };
  }, [open]);

  const handleCopy = async () => {
    const value = normalizeResultLatex(latex);
    if (!value) return;
    try {
      await navigator.clipboard.writeText(value);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1200);
    } catch (copyError) {
      setError(readError(copyError));
    }
  };

  const handleInsert = () => {
    const value = normalizeResultLatex(latex);
    if (!value) return;
    onInsert(value);
    onNotify(isEn ? "OCR formula inserted at the cursor" : "OCR 公式已插入当前光标");
    onClose();
  };

  const handleAppend = () => {
    const value = normalizeResultLatex(latex);
    if (!value) return;
    onAppend(value);
    onNotify(isEn ? "OCR formula appended as a new line" : "OCR 公式已追加为新公式行");
    onClose();
  };

  const handleRestartWorker = async () => {
    try {
      await restartOcrWorker();
      await prewarmOcrModel(model);
      setResult(null);
      setLatex("");
      setError("");
      onNotify(isEn ? "OCR worker restarted" : "OCR 识别进程已重启");
    } catch (restartError) {
      setError(readError(restartError));
    }
  };

  const handleResetRuntime = async () => {
    const confirmed = window.confirm(
      isEn
        ? "Remove the OCR runtime?"
        : "删除 OCR 运行环境？",
    );
    if (!confirmed) return;

    setCheckingRuntime(true);
    setError("");
    try {
      setRuntime(await resetOcrRuntime());
      setResult(null);
      setLatex("");
    } catch (resetError) {
      setError(readError(resetError));
    } finally {
      setCheckingRuntime(false);
    }
  };

  if (!open) return null;

  return (
    <div
      className="modal-backdrop ocr-modal-backdrop"
      role="presentation"
      onMouseDown={requestClose}
    >
      <section
        ref={dialogRef}
        className="ocr-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="ocr-dialog-title"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="dialog-header ocr-dialog-header">
          <div className="ocr-heading">
            <span className="ocr-heading-icon">
              <ScanLine size={20} />
            </span>
            <div>
              <span className="eyebrow">VISUALTEX OCR</span>
              <h2 id="ocr-dialog-title">{isEn ? "Formula image recognition" : "图片公式识别"}</h2>
            </div>
          </div>
          <button
            type="button"
            className="icon-button"
            onClick={requestClose}
            aria-label={isEn ? "Close OCR" : "关闭 OCR"}
          >
            <X size={18} />
          </button>
        </header>

        <div className="ocr-dialog-body">
          <div className="ocr-input-column">
            <input
              ref={fileInputRef}
              type="file"
              className="visually-hidden"
              accept="image/png,image/jpeg,image/webp,image/bmp,image/tiff"
              onChange={handleFileInput}
            />

            <div
              className={
                "ocr-drop-zone" +
                (dragging ? " is-dragging" : "") +
                (previewUrl ? " has-image" : "")
              }
              onDragEnter={(event) => {
                event.preventDefault();
                setDragging(true);
              }}
              onDragOver={(event) => event.preventDefault()}
              onDragLeave={(event) => {
                if (event.currentTarget === event.target) setDragging(false);
              }}
              onDrop={handleDrop}
            >
              {previewUrl ? (
                <>
                  <img
                    src={previewUrl}
                    alt={isEn ? "Formula source preview" : "公式原图预览"}
                    onLoad={(event) =>
                      setImageSize({
                        width: event.currentTarget.naturalWidth,
                        height: event.currentTarget.naturalHeight,
                      })
                    }
                  />
                  <div className="ocr-image-actions">
                    <button type="button" onClick={() => fileInputRef.current?.click()}>
                      <RefreshCw size={14} />
                      {isEn ? "Replace" : "更换图片"}
                    </button>
                  </div>
                </>
              ) : (
                <div className="ocr-drop-empty">
                  <span className="ocr-drop-icon">
                    <ImagePlus size={28} />
                  </span>
                  <strong>{isEn ? "Drop an image" : "拖入图片"}</strong>
                  <button type="button" onClick={() => fileInputRef.current?.click()}>
                    <Upload size={15} />
                    {isEn ? "Choose image" : "选择图片"}
                  </button>
                  <small>
                    <ClipboardPaste size={13} />
                    {isEn ? "Paste (⌘V)" : "粘贴（⌘V）"}
                  </small>
                </div>
              )}
            </div>

            {file && (
              <div className="ocr-file-meta">
                <span>{file.name || (isEn ? "Clipboard image" : "剪贴板图片")}</span>
                <span>
                  {imageSize.width > 0 ? imageSize.width + "×" + imageSize.height + " · " : ""}
                  {readableBytes(file.size)}
                </span>
              </div>
            )}

            <section className="ocr-provider-card" aria-label={isEn ? "OCR provider" : "OCR 提供器"}>
              <div className="ocr-provider-heading">
                <span className={`ocr-provider-icon${usingLocalProvider ? " is-local" : " is-api"}`}>
                  {usingLocalProvider ? <Server size={17} /> : <Cloud size={17} />}
                </span>
                <div>
                  <strong>{isEn ? "Recognition provider" : "识别提供器"}</strong>
                  <span>{activeProviderLabel}</span>
                </div>
              </div>

              <label className="ocr-provider-field">
                <span>{isEn ? "Provider" : "提供器"}</span>
                <select
                  value={activeProvider}
                  disabled={recognizing || savingProvider}
                  onChange={(event) =>
                    updateProviderConfiguration((current) => ({
                      ...current,
                      activeProvider: event.target.value as OcrProviderId,
                    }))
                  }
                >
                  <option value="local">{isEn ? "Local PP-FormulaNet" : "本地 PP-FormulaNet"}</option>
                  <option value="openai-compatible">{isEn ? "OpenAI-compatible API" : "OpenAI 兼容 API"}</option>
                  <option value="ollama">Ollama</option>
                  <option value="mathpix">Mathpix</option>
                  <option value="paddleocr">{isEn ? "PaddleOCR AI Studio" : "PaddleOCR 星河 API"}</option>
                  <option value="simpletex">SimpleTex</option>
                </select>
              </label>

              {activeProvider === "openai-compatible" && (
                <div className="ocr-provider-fields">
                  <label className="ocr-provider-field">
                    <span>{isEn ? "Protocol" : "协议"}</span>
                    <select
                      value={providerConfiguration.openAiCompatible.protocol}
                      disabled={recognizing || savingProvider}
                      onChange={(event) =>
                        updateProviderConfiguration((current) => ({
                          ...current,
                          openAiCompatible: {
                            ...current.openAiCompatible,
                            protocol: event.target.value as "responses" | "chat-completions",
                          },
                        }))
                      }
                    >
                      <option value="responses">Responses API</option>
                      <option value="chat-completions">Chat Completions</option>
                    </select>
                  </label>
                  <label className="ocr-provider-field is-wide">
                    <span>{isEn ? "Base URL" : "基础地址"}</span>
                    <input
                      value={providerConfiguration.openAiCompatible.baseUrl}
                      disabled={recognizing || savingProvider}
                      spellCheck={false}
                      placeholder="https://api.openai.com/v1"
                      onChange={(event) =>
                        updateProviderConfiguration((current) => ({
                          ...current,
                          openAiCompatible: {
                            ...current.openAiCompatible,
                            baseUrl: event.target.value,
                          },
                        }))
                      }
                    />
                  </label>
                  <label className="ocr-provider-field">
                    <span>{isEn ? "Vision model" : "视觉模型"}</span>
                    <input
                      value={providerConfiguration.openAiCompatible.model}
                      disabled={recognizing || savingProvider}
                      spellCheck={false}
                      onChange={(event) =>
                        updateProviderConfiguration((current) => ({
                          ...current,
                          openAiCompatible: {
                            ...current.openAiCompatible,
                            model: event.target.value,
                          },
                        }))
                      }
                    />
                  </label>
                  <label className="ocr-provider-field">
                    <span>
                      {isEn ? "API key" : "API 密钥"}
                      {credentialStatusLabel(providerConfiguration.openAiCompatible.hasApiKey, isEn, clearOpenAiApiKey)}
                    </span>
                    <input
                      type="password"
                      value={openAiApiKey}
                      disabled={recognizing || savingProvider || clearOpenAiApiKey}
                      autoComplete="new-password"
                      placeholder={
                        providerConfiguration.openAiCompatible.hasApiKey === true
                          ? isEn
                            ? "Unchanged if blank"
                            : "留空则不更改"
                          : providerConfiguration.openAiCompatible.hasApiKey === null
                            ? isEn
                              ? "Unchanged if blank"
                              : "留空则不更改"
                          : isEn
                            ? "Optional"
                            : "可选"
                      }
                      onChange={(event) => {
                        setOpenAiApiKey(event.target.value);
                        setProviderDirty(true);
                      }}
                    />
                  </label>
                  {providerConfiguration.openAiCompatible.hasApiKey !== false && (
                    <label className="ocr-provider-clear-secret">
                      <input
                        type="checkbox"
                        checked={clearOpenAiApiKey}
                        disabled={recognizing || savingProvider}
                        onChange={(event) => {
                          setClearOpenAiApiKey(event.target.checked);
                          setProviderDirty(true);
                        }}
                      />
                      <span>{isEn ? "Remove the saved API key" : "删除已保存的 API 密钥"}</span>
                    </label>
                  )}
                  <label className="ocr-provider-field is-wide">
                    <span>{isEn ? "Recognition prompt" : "识别提示词"}</span>
                    <textarea
                      value={providerConfiguration.openAiCompatible.prompt}
                      disabled={recognizing || savingProvider}
                      spellCheck={false}
                      onChange={(event) =>
                        updateProviderConfiguration((current) => ({
                          ...current,
                          openAiCompatible: {
                            ...current.openAiCompatible,
                            prompt: event.target.value,
                          },
                        }))
                      }
                    />
                  </label>
                </div>
              )}

              {activeProvider === "ollama" && (
                <div className="ocr-provider-fields">
                  <label className="ocr-provider-field is-wide">
                    <span>{isEn ? "Ollama address" : "Ollama 地址"}</span>
                    <input
                      value={providerConfiguration.ollama.baseUrl}
                      disabled={recognizing || savingProvider}
                      spellCheck={false}
                      placeholder="http://127.0.0.1:11434"
                      onChange={(event) =>
                        updateProviderConfiguration((current) => ({
                          ...current,
                          ollama: { ...current.ollama, baseUrl: event.target.value },
                        }))
                      }
                    />
                  </label>
                  <label className="ocr-provider-field is-wide">
                    <span>{isEn ? "Vision model" : "视觉模型"}</span>
                    <input
                      value={providerConfiguration.ollama.model}
                      disabled={recognizing || savingProvider}
                      spellCheck={false}
                      onChange={(event) =>
                        updateProviderConfiguration((current) => ({
                          ...current,
                          ollama: { ...current.ollama, model: event.target.value },
                        }))
                      }
                    />
                  </label>
                  <label className="ocr-provider-field is-wide">
                    <span>{isEn ? "Recognition prompt" : "识别提示词"}</span>
                    <textarea
                      value={providerConfiguration.ollama.prompt}
                      disabled={recognizing || savingProvider}
                      spellCheck={false}
                      onChange={(event) =>
                        updateProviderConfiguration((current) => ({
                          ...current,
                          ollama: { ...current.ollama, prompt: event.target.value },
                        }))
                      }
                    />
                  </label>
                </div>
              )}

              {activeProvider === "mathpix" && (
                <div className="ocr-provider-fields">
                  <label className="ocr-provider-field is-wide">
                    <span>{isEn ? "Mathpix address" : "Mathpix 地址"}</span>
                    <input
                      value={providerConfiguration.mathpix.baseUrl}
                      disabled={recognizing || savingProvider}
                      spellCheck={false}
                      placeholder="https://api.mathpix.com"
                      onChange={(event) =>
                        updateProviderConfiguration((current) => ({
                          ...current,
                          mathpix: { ...current.mathpix, baseUrl: event.target.value },
                        }))
                      }
                    />
                  </label>
                  <label className="ocr-provider-field">
                    <span>app_id</span>
                    <input
                      value={providerConfiguration.mathpix.appId}
                      disabled={recognizing || savingProvider}
                      spellCheck={false}
                      onChange={(event) =>
                        updateProviderConfiguration((current) => ({
                          ...current,
                          mathpix: { ...current.mathpix, appId: event.target.value },
                        }))
                      }
                    />
                  </label>
                  <label className="ocr-provider-field">
                    <span>
                      app_key
                      {credentialStatusLabel(providerConfiguration.mathpix.hasAppKey, isEn, clearMathpixAppKey)}
                    </span>
                    <input
                      type="password"
                      value={mathpixAppKey}
                      disabled={recognizing || savingProvider || clearMathpixAppKey}
                      autoComplete="new-password"
                      placeholder={
                        providerConfiguration.mathpix.hasAppKey === true
                          ? isEn
                            ? "Unchanged if blank"
                            : "留空则不更改"
                          : providerConfiguration.mathpix.hasAppKey === null
                            ? isEn
                              ? "Unchanged if blank"
                              : "留空则不更改"
                          : "app_key"
                      }
                      onChange={(event) => {
                        setMathpixAppKey(event.target.value);
                        setProviderDirty(true);
                      }}
                    />
                  </label>
                  {providerConfiguration.mathpix.hasAppKey !== false && (
                    <label className="ocr-provider-clear-secret">
                      <input
                        type="checkbox"
                        checked={clearMathpixAppKey}
                        disabled={recognizing || savingProvider}
                        onChange={(event) => {
                          setClearMathpixAppKey(event.target.checked);
                          setProviderDirty(true);
                        }}
                      />
                      <span>{isEn ? "Remove the saved app_key" : "删除已保存的 app_key"}</span>
                    </label>
                  )}
                </div>
              )}

              {activeProvider === "paddleocr" && (
                <div className="ocr-provider-fields">
                  <label className="ocr-provider-field">
                    <span>{isEn ? "PaddleOCR model" : "PaddleOCR 模型"}</span>
                    <select
                      value={providerConfiguration.paddleOcr.model}
                      disabled={recognizing || savingProvider}
                      onChange={(event) =>
                        updateProviderConfiguration((current) => ({
                          ...current,
                          paddleOcr: {
                            ...current.paddleOcr,
                            model: event.target.value as PaddleOcrApiModel,
                          },
                        }))
                      }
                    >
                      {PADDLE_OCR_API_MODELS.map((candidate) => (
                        <option value={candidate} key={candidate}>{candidate}</option>
                      ))}
                    </select>
                  </label>
                  <label className="ocr-provider-field is-wide">
                    <span>
                      {isEn ? "Access token" : "访问令牌"}
                      {credentialStatusLabel(providerConfiguration.paddleOcr.hasAccessToken, isEn, clearPaddleOcrAccessToken)}
                    </span>
                    <input
                      type="password"
                      value={paddleOcrAccessToken}
                      disabled={recognizing || savingProvider || clearPaddleOcrAccessToken}
                      autoComplete="new-password"
                      placeholder={
                        providerConfiguration.paddleOcr.hasAccessToken === true
                          ? isEn ? "Unchanged if blank" : "留空则不更改"
                          : providerConfiguration.paddleOcr.hasAccessToken === null
                            ? isEn ? "Unchanged if blank" : "留空则不更改"
                          : isEn ? "AI Studio TOKEN" : "AI Studio TOKEN"
                      }
                      onChange={(event) => {
                        setPaddleOcrAccessToken(event.target.value);
                        setProviderDirty(true);
                      }}
                    />
                  </label>
                  {providerConfiguration.paddleOcr.hasAccessToken !== false && (
                    <label className="ocr-provider-clear-secret">
                      <input
                        type="checkbox"
                        checked={clearPaddleOcrAccessToken}
                        disabled={recognizing || savingProvider}
                        onChange={(event) => {
                          setClearPaddleOcrAccessToken(event.target.checked);
                          setProviderDirty(true);
                        }}
                      />
                      <span>{isEn ? "Remove the saved access token" : "删除已保存的访问令牌"}</span>
                    </label>
                  )}
                </div>
              )}

              {activeProvider === "simpletex" && (
                <div className="ocr-provider-fields">
                  <label className="ocr-provider-field">
                    <span>{isEn ? "Formula model" : "公式模型"}</span>
                    <select
                      value={providerConfiguration.simpleTex.model}
                      disabled={recognizing || savingProvider}
                      onChange={(event) =>
                        updateProviderConfiguration((current) => ({
                          ...current,
                          simpleTex: {
                            ...current.simpleTex,
                            model: event.target.value as SimpleTexApiModel,
                          },
                        }))
                      }
                    >
                      {SIMPLETEX_API_MODELS.map((candidate) => (
                        <option key={candidate} value={candidate}>
                          {candidate === "standard"
                            ? isEn
                              ? "Standard"
                              : "标准"
                            : isEn
                              ? "Turbo"
                              : "轻量"}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label className="ocr-provider-field is-wide">
                    <span>
                      {isEn ? "User access token (UAT)" : "用户授权令牌（UAT）"}
                      {credentialStatusLabel(providerConfiguration.simpleTex.hasAccessToken, isEn, clearSimpleTexAccessToken)}
                    </span>
                    <input
                      type="password"
                      value={simpleTexAccessToken}
                      disabled={recognizing || savingProvider || clearSimpleTexAccessToken}
                      autoComplete="new-password"
                      placeholder={
                        providerConfiguration.simpleTex.hasAccessToken === true
                          ? isEn
                            ? "Unchanged if blank"
                            : "留空则不更改"
                          : providerConfiguration.simpleTex.hasAccessToken === null
                            ? isEn
                              ? "Unchanged if blank"
                              : "留空则不更改"
                          : "SimpleTex UAT"
                      }
                      onChange={(event) => {
                        setSimpleTexAccessToken(event.target.value);
                        setProviderDirty(true);
                      }}
                    />
                  </label>
                  {providerConfiguration.simpleTex.hasAccessToken !== false && (
                    <label className="ocr-provider-clear-secret">
                      <input
                        type="checkbox"
                        checked={clearSimpleTexAccessToken}
                        disabled={recognizing || savingProvider}
                        onChange={(event) => {
                          setClearSimpleTexAccessToken(event.target.checked);
                          setProviderDirty(true);
                        }}
                      />
                      <span>
                        {isEn
                          ? "Remove the saved SimpleTex token"
                          : "删除已保存的 SimpleTex 令牌"}
                      </span>
                    </label>
                  )}
                </div>
              )}


              <div className="ocr-provider-actions">
                <span>
                  <KeyRound size={13} />
                  {isEn
                    ? ""
                    : ""}
                </span>
                <button
                  type="button"
                  className={providerDirty ? "primary-button" : "secondary-button"}
                  disabled={!providerLoaded || !providerDirty || savingProvider || recognizing}
                  onClick={() => void handleSaveProviderConfiguration()}
                >
                  {savingProvider ? (
                    <LoaderCircle size={14} className="is-spinning" />
                  ) : (
                    <Save size={14} />
                  )}
                  {savingProvider
                    ? isEn
                      ? "Saving…"
                      : "正在保存…"
                    : isEn
                      ? "Save provider"
                      : "保存提供器"}
                </button>
              </div>
            </section>

            {usingLocalProvider && (
              <>
            <label className="ocr-model-field">
              <span>{isEn ? "Recognition model" : "识别模型"}</span>
              <select
                value={model}
                disabled={recognizing || cancelling || modelBusy}
                onChange={(event) =>
                  onModelChange(event.target.value as OcrModelName)
                }
              >
                {OCR_MODELS.map((item) => {
                  const available = installedModels.includes(item.id);
                  return (
                    <option value={item.id} key={item.id}>
                      {isEn ? item.labelEn : item.labelZh}
                      {available
                        ? isEn
                          ? " · installed"
                          : " · 已安装"
                        : isEn
                          ? " · not installed"
                          : " · 未安装"}
                    </option>
                  );
                })}
              </select>
              <small>{isEn ? selectedModel.hintEn : selectedModel.hintZh}</small>
            </label>

            <div
              ref={modelDropZoneRef}
              className={
                "ocr-model-warning ocr-model-drop-zone" +
                (modelPackageDragging ? " is-dragging" : "")
              }
              role="status"
            >
              {selectedModelInstalled ? (
                <CheckCircle2 size={15} />
              ) : (
                <AlertCircle size={15} />
              )}
              <div>
                <strong>
                  {selectedModelInstalled
                    ? isEn
                      ? `${selectedModel.labelEn} is installed`
                      : `${selectedModel.labelZh}已安装`
                    : isEn
                      ? `${selectedModel.labelEn} is not installed`
                      : `${selectedModel.labelZh}尚未安装`}
                </strong>
                <span>
                  {selectedModelInstalled
                    ? isEn
                      ? ""
                      : ""
                    : isEn
                      ? "Import a model package."
                      : "请导入模型包。"}
                </span>
                <small className="ocr-model-drop-hint">
                  {modelPackageDragging
                    ? isEn
                      ? "Release to import"
                      : "松开以导入"
                    : isEn
                      ? ""
                      : ""}
                </small>
                <div className="ocr-install-actions">
                  {!selectedModelInstalled && (
                    <button
                      type="button"
                      onClick={() => void handleImportModel()}
                      disabled={!isTauriEnvironment() || modelBusy}
                    >
                      {modelBusy ? (
                        <LoaderCircle size={14} className="is-spinning" />
                      ) : (
                        <Upload size={14} />
                      )}
                      {isEn ? "Import package" : "导入模型包"}
                    </button>
                  )}
                  {selectedModelRemovable && (
                    <button
                      type="button"
                      className="is-danger"
                      onClick={() => void handleRemoveModel()}
                      disabled={modelBusy || recognizing}
                    >
                      <Trash2 size={14} />
                      {isEn ? "Remove model" : "删除模型"}
                    </button>
                  )}
                </div>
              </div>
            </div>
              </>
            )}
          </div>

          <div className="ocr-output-column">
            {usingLocalProvider ? (
            <section className="ocr-runtime-card">
              <div className="ocr-runtime-summary">
                <span className={"ocr-runtime-icon " + (runtime?.installed ? "is-ready" : "")}>
                  {checkingRuntime ? (
                    <LoaderCircle size={17} className="is-spinning" />
                  ) : runtime?.installed ? (
                    <CheckCircle2 size={17} />
                  ) : (
                    <Cpu size={17} />
                  )}
                </span>
                <div>
                  <strong>
                    {runtime?.installed
                      ? isEn
                        ? "Local OCR runtime ready"
                        : "本地 OCR 环境已就绪"
                      : isEn
                        ? "OCR runtime is not installed"
                        : "尚未安装 OCR 运行环境"}
                  </strong>
                  <span>{runtime?.message ?? (isEn ? "Checking runtime…" : "正在检查运行环境…")}</span>
                </div>
              </div>

              {runtime?.installed ? (
                <div className="ocr-runtime-details">
                  <span>Python {runtime.pythonVersion}</span>
                  <span>Paddle {runtime.paddleVersion}</span>
                  <span>PaddleOCR {runtime.paddleocrVersion}</span>
                  <span>
                    {isEn ? "Models" : "模型"}: {installedModels.length > 0
                      ? installedModels.join(", ")
                      : isEn
                        ? "none"
                        : "未安装"}
                  </span>
                  <button
                    type="button"
                    onClick={handleRestartWorker}
                    disabled={!selectedModelInstalled}
                  >
                    <RefreshCw size={13} />
                    {isEn ? "Restart" : "重启进程"}
                  </button>
                  <button type="button" className="is-danger" onClick={handleResetRuntime}>
                    <Trash2 size={13} />
                    {isEn ? "Reset" : "重置环境"}
                  </button>
                </div>
              ) : (
                <div className="ocr-install-panel">
                  {installing && installProgress ? (
                    <>
                      <div className="ocr-progress-label">
                        <span>{installProgress.message}</span>
                        <strong>{installProgress.percent}%</strong>
                      </div>
                      <div className="ocr-progress-track">
                        <span style={{ width: installProgress.percent + "%" }} />
                      </div>
                      {installProgress.detail && <small>{installProgress.detail}</small>}
                    </>
                  ) : (
                    <>
                      <button
                        type="button"
                        className="primary-button"
                        onClick={handleInstall}
                        disabled={
                          (!isTauriEnvironment() &&
                            !isOfficeCompanionEnvironment()) ||
                          checkingRuntime
                        }
                      >
                        <Download size={15} />
                        {isEn ? "Install OCR runtime" : "安装 OCR 运行环境"}
                      </button>
                    </>
                  )}
                </div>
              )}
            </section>
            ) : (
              <section className="ocr-runtime-card ocr-api-runtime-card">
                <div className="ocr-runtime-summary">
                  <span className={`ocr-runtime-icon${providerDirty ? "" : " is-ready"}`}>
                    {providerDirty ? <AlertCircle size={17} /> : <Cloud size={17} />}
                  </span>
                  <div>
                    <strong>
                      {providerDirty
                        ? isEn
                          ? "OCR API changes are not saved"
                          : "OCR API 修改尚未保存"
                        : isEn
                          ? `${activeProviderLabel} is active`
                          : `当前使用${activeProviderLabel}`}
                    </strong>
                    <span>
                      {providerDirty
                        ? isEn
                          ? "Configuration not saved"
                          : "配置未保存"
                        : isEn
                          ? ""
                          : ""}
                    </span>
                  </div>
                </div>
                <div className="ocr-runtime-details ocr-api-provider-details">
                  {activeProvider === "openai-compatible" && (
                    <>
                      <span>{providerConfiguration.openAiCompatible.protocol === "responses" ? "Responses API" : "Chat Completions"}</span>
                      <span>{providerConfiguration.openAiCompatible.model || (isEn ? "model not set" : "未填写模型")}</span>
                      <code>{providerConfiguration.openAiCompatible.baseUrl}</code>
                    </>
                  )}
                  {activeProvider === "ollama" && (
                    <>
                      <span>{providerConfiguration.ollama.model || (isEn ? "model not set" : "未填写模型")}</span>
                      <span>/api/chat</span>
                      <code>{providerConfiguration.ollama.baseUrl}</code>
                    </>
                  )}
                  {activeProvider === "mathpix" && (
                    <>
                      <span>{providerConfiguration.mathpix.appId || "app_id"}</span>
                      <span>POST /v3/text</span>
                      <code>{providerConfiguration.mathpix.baseUrl}</code>
                    </>
                  )}
                  {activeProvider === "paddleocr" && (
                    <>
                      <span>{providerConfiguration.paddleOcr.model}</span>
                      <span>{isEn ? "Official asynchronous API" : "官方统一异步 API"}</span>
                      <code>paddleocr.aistudio-app.com/api/v2/ocr/jobs</code>
                    </>
                  )}
                  {activeProvider === "simpletex" && (
                    <>
                      <span>
                        {providerConfiguration.simpleTex.model === "standard"
                          ? isEn
                            ? "Standard formula model"
                            : "标准公式模型"
                          : isEn
                            ? "Turbo formula model"
                            : "轻量公式模型"}
                      </span>
                      <span>SimpleTex V2.5</span>
                      <code>
                        {providerConfiguration.simpleTex.model === "standard"
                          ? "server.simpletex.cn/api/latex_ocr"
                          : "server.simpletex.cn/api/latex_ocr_turbo"}
                      </code>
                    </>
                  )}
                </div>
              </section>
            )}

            <section className="ocr-result-card">
              <div className="ocr-result-heading">
                <div>
                  <span className="eyebrow">LATEX RESULT</span>
                  <strong>{isEn ? "Recognition result" : "识别结果"}</strong>
                </div>
                {result && (
                  <span>
                    {providerLabel(result.provider as OcrProviderId, isEn)} · {result.backgroundInverted
                      ? isEn
                        ? "Background inverted · "
                        : "已反色 · "
                      : ""}
                    {result.elapsedMs} ms · {result.processedWidth}×{result.processedHeight}
                  </span>
                )}
              </div>

              {recognizing ? (
                <div className="ocr-recognizing-state">
                  <LoaderCircle size={24} className="is-spinning" />
                  <strong>
                    {recognitionProgress?.message ??
                      (isEn ? "Recognizing formula…" : "正在识别公式…")}
                  </strong>
                  <span>
                    {usingLocalProvider
                      ? isEn
                        ? `${selectedModel.labelEn} · ${recognitionSeconds}s elapsed`
                        : `${selectedModel.labelZh} · 已等待 ${recognitionSeconds} 秒`
                      : isEn
                        ? `${activeProviderLabel} · ${recognitionSeconds}s elapsed`
                        : `${activeProviderLabel} · 已等待 ${recognitionSeconds} 秒`}
                  </span>
                </div>
              ) : latex ? (
                <>
                  <div className="ocr-formula-preview">
                    <MathPreview latex={latex.split("\n")[0]} />
                  </div>
                  <label className="ocr-latex-editor">
                    <span>{isEn ? "Editable LaTeX" : "可编辑 LaTeX"}</span>
                    <textarea value={latex} onChange={(event) => setLatex(event.target.value)} spellCheck={false} />
                  </label>
                </>
              ) : (
                <div className="ocr-empty-result">
                  <ScanLine size={24} />
                  <span>
                    {isEn
                      ? ""
                      : ""}
                  </span>
                </div>
              )}
            </section>

            {error && (
              <div className="ocr-error-box" role="alert">
                <AlertCircle size={16} />
                <pre>{error}</pre>
              </div>
            )}
          </div>
        </div>

        <footer className="dialog-footer ocr-dialog-footer">
          {recognizing ? (
            <button
              type="button"
              className="secondary-button is-danger"
              onClick={handleCancelRecognition}
              disabled={cancelling}
            >
              {cancelling ? (
                <LoaderCircle size={15} className="is-spinning" />
              ) : (
                <X size={15} />
              )}
              {cancelling
                ? isEn
                  ? "Stopping…"
                  : "正在停止…"
                : isEn
                  ? "Cancel recognition"
                  : "取消识别"}
            </button>
          ) : (
            <button
              type="button"
              className="secondary-button"
              onClick={handleRecognize}
              disabled={
                !file ||
                !providerLoaded ||
                providerDirty ||
                (usingLocalProvider && (!runtime?.installed || !selectedModelInstalled)) ||
                (usingLocalProvider && (installing || modelBusy)) ||
                savingProvider
              }
            >
              <ScanLine size={15} />
              {isEn ? "Recognize" : "开始识别"}
            </button>
          )}
          <div className="ocr-result-actions">
            <button type="button" className="secondary-button" onClick={handleCopy} disabled={!latex.trim()}>
              {copied ? <Check size={15} /> : <Copy size={15} />}
              {copied ? (isEn ? "Copied" : "已复制") : isEn ? "Copy LaTeX" : "复制 LaTeX"}
            </button>
            <button type="button" className="secondary-button" onClick={handleAppend} disabled={!latex.trim()}>
              <Plus size={15} />
              {isEn ? "Append line" : "追加为新行"}
            </button>
            <button type="button" className="primary-button" onClick={handleInsert} disabled={!latex.trim()}>
              <ScanLine size={15} />
              {isEn ? "Insert at cursor" : "插入当前光标"}
            </button>
          </div>
        </footer>
      </section>
    </div>
  );
}
