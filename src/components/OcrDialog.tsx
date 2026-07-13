import {
  AlertCircle,
  Check,
  CheckCircle2,
  ClipboardPaste,
  Copy,
  Cpu,
  Download,
  ImagePlus,
  LoaderCircle,
  Plus,
  RefreshCw,
  ScanLine,
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
  OCR_MODELS,
  cancelOcrRecognition,
  type OcrInstallProgress,
  type OcrModelName,
  type OcrRecognitionProgress,
  type OcrRecognitionResult,
  type OcrRuntimeStatus,
  fileToOcrRequest,
  getOcrRuntimeStatus,
  installOcrRuntime,
  isOfficeCompanionEnvironment,
  isTauriEnvironment,
  listenOcrInstallProgress,
  listenOcrRecognitionProgress,
  recognizeFormulaImage,
  resolveAvailableOcrModel,
  resetOcrRuntime,
  restartOcrWorker,
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
}

function readableBytes(bytes: number) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1024 / 1024).toFixed(1) + " MB";
}

function readError(error: unknown) {
  if (error instanceof Error) return error.message;
  if (typeof error === "string") return error;
  try {
    return JSON.stringify(error);
  } catch {
    return "Unknown OCR error";
  }
}

function normalizeResultLatex(value: string) {
  return value
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .join("\n");
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
}: OcrDialogProps) {
  const isEn = language === "en";
  const dialogRef = useRef<HTMLElement>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const recognizingRef = useRef(false);
  const cancellingRef = useRef(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const objectUrlRef = useRef<string | null>(null);
  const [runtime, setRuntime] = useState<OcrRuntimeStatus | null>(null);
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
    () => OCR_MODELS.find((item) => item.id === model) ?? OCR_MODELS[1],
    [model],
  );
  const defaultModel = runtime?.defaultModel ?? "PP-FormulaNet_plus-M";
  const installedModels = runtime?.installedModels ?? [];
  const selectedModelInstalled = installedModels.includes(model);
  const optionalModelMissing = model !== defaultModel && !selectedModelInstalled;

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
        defaultModel: "PP-FormulaNet_plus-M",
        message: isEn
          ? "OCR is available in the VisualTeX desktop app, not in the browser preview."
          : "OCR 只能在 VisualTeX 桌面应用中运行，浏览器预览无法调用本地模型。",
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

  useEffect(() => {
    if (!runtime?.installed) return;
    const availableModel = resolveAvailableOcrModel(runtime, model);
    if (availableModel !== model) onModelChange(availableModel);
  }, [defaultModel, model, onModelChange, runtime]);

  useEffect(() => {
    if (!open) return;
    setError("");
    if (runtime) return;

    let cancelled = false;
    const frame = window.requestAnimationFrame(() => {
      if (!cancelled) void refreshRuntime();
    });
    return () => {
      cancelled = true;
      window.cancelAnimationFrame(frame);
    };
  }, [open, runtime, refreshRuntime]);

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
      onNotify(isEn ? "OCR runtime installed" : "OCR 运行环境安装完成");
    } catch (installError) {
      setError(readError(installError));
      await refreshRuntime(true);
    } finally {
      unlisten?.();
      setInstalling(false);
    }
  };

  const handleRecognize = async () => {
    if (!file) {
      setError(isEn ? "Choose or paste a formula image first." : "请先选择或粘贴一张公式图片。");
      return;
    }
    if (!runtime?.installed) {
      setError(isEn ? "Install the OCR runtime first." : "请先安装 OCR 运行环境。");
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
      message: isEn ? "Preparing the formula image" : "正在准备公式图片",
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
      setLatex(
        normalizeResultLatex(nextResult.formulas.map((formula) => formula.latex).join("\n")),
      );
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
    await navigator.clipboard.writeText(value);
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1200);
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
        ? "Remove the OCR runtime and its installed packages?"
        : "确定删除 OCR 运行环境和已经安装的依赖吗？",
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
              <span className="eyebrow">PP-FORMULANET OCR</span>
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
                  <strong>{isEn ? "Drop a formula image here" : "将公式图片拖到这里"}</strong>
                  <span>{isEn ? "Choose a file or paste an image" : "选择文件，或直接粘贴剪贴板图片"}</span>
                  <button type="button" onClick={() => fileInputRef.current?.click()}>
                    <Upload size={15} />
                    {isEn ? "Choose image" : "选择图片"}
                  </button>
                  <small>
                    <ClipboardPaste size={13} />
                    {isEn ? "Paste with ⌘V while this dialog is open" : "窗口打开时可直接按 ⌘V 粘贴"}
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

            <label className="ocr-model-field">
              <span>{isEn ? "Recognition model" : "识别模型"}</span>
              <select
                value={model}
                disabled={recognizing || cancelling}
                onChange={(event) =>
                  onModelChange(event.target.value as OcrModelName)
                }
              >
                {OCR_MODELS.map((item) => {
                  const available =
                    installedModels.includes(item.id);
                  return (
                    <option value={item.id} key={item.id} disabled={!available}>
                      {isEn ? item.labelEn : item.labelZh}
                      {!available
                        ? isEn
                          ? " · optional offline pack required"
                          : " · 需要可选离线模型包"
                        : ""}
                    </option>
                  );
                })}
              </select>
              <small>{isEn ? selectedModel.hintEn : selectedModel.hintZh}</small>
            </label>

            {(model === "PP-FormulaNet_plus-L" || optionalModelMissing) && (
              <div className="ocr-model-warning" role="note">
                <AlertCircle size={15} />
                <span>
                  {optionalModelMissing
                    ? isEn
                      ? `${selectedModel.labelEn} is not installed. Import the matching VisualTeX offline model pack before selecting it.`
                      : `${selectedModel.labelZh}尚未安装，请先导入对应的 VisualTeX 离线模型包。`
                    : isEn
                      ? "The L model occupies about 698 MB and can use several GB of memory. Use the bundled M model unless L accuracy is necessary."
                      : "L 模型约占 698 MB，并可能占用数 GB 内存；没有明确精度需求时建议使用内置 M 模型。"}
                </span>
              </div>
            )}

            <div className="ocr-input-tip">
              <AlertCircle size={14} />
              <span>
                {isEn
                  ? "Use a tight crop around one formula. Avoid blur, shadows, and perspective distortion."
                  : "建议只截取一条公式并尽量裁紧，避免模糊、阴影和明显透视变形。"}
              </span>
            </div>
          </div>

          <div className="ocr-output-column">
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
                  <button type="button" onClick={handleRestartWorker}>
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
                      <p>
                        {isEn
                          ? "VisualTeX will verify and extract the bundled Python 3.10, PaddlePaddle 3.3.1, PaddleOCR 3.7.0, and the default M model entirely on this Mac. No network or pip installation is used."
                          : "VisualTeX 会在本机校验并解压应用内置的 Python 3.10、PaddlePaddle 3.3.1、PaddleOCR 3.7.0 与默认 M 模型；全程不联网，也不会运行 pip 安装。"}
                      </p>
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

            <section className="ocr-result-card">
              <div className="ocr-result-heading">
                <div>
                  <span className="eyebrow">LATEX RESULT</span>
                  <strong>{isEn ? "Recognition result" : "识别结果"}</strong>
                </div>
                {result && (
                  <span>
                    {result.backgroundInverted
                      ? isEn
                        ? "Dark background normalized · "
                        : "已自动反色 · "
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
                    {isEn
                      ? `${selectedModel.labelEn} · ${recognitionSeconds}s elapsed`
                      : `${selectedModel.labelZh} · 已等待 ${recognitionSeconds} 秒`}
                  </span>
                  <small className="ocr-recognition-meta">
                    {isEn
                      ? `First use may download ${selectedModel.downloadMb.toFixed(1)} MB. You can cancel without closing VisualTeX.`
                      : `首次使用可能需要下载 ${selectedModel.downloadMb.toFixed(1)} MB；现在可以随时取消，不会卡住 VisualTeX。`}
                  </small>
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
                      ? "Choose an image and run recognition."
                      : "选择图片并开始识别后，结果会显示在这里。"}
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
              disabled={!file || !runtime?.installed || installing}
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
