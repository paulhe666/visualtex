import {
  AlertCircle,
  Check,
  ClipboardPaste,
  Cloud,
  Copy,
  ImagePlus,
  LoaderCircle,
  Plus,
  RefreshCw,
  Save,
  ScanLine,
  ShieldCheck,
  Upload,
  X,
} from "lucide-react";
import {
  type ChangeEvent,
  type DragEvent,
  useCallback,
  useEffect,
  useRef,
  useState,
} from "react";
import { MathPreview } from "./MathPreview";
import {
  loadWebOcrConfiguration,
  recognizeFormulaWithWebApi,
  saveWebOcrConfiguration,
  type WebOcrConfiguration,
  type WebOcrProgress,
  type WebOcrProvider,
} from "../ocr/webOcrService";

interface WebOcrDialogProps {
  open: boolean;
  language: "cn" | "en";
  onClose: () => void;
  onInsert: (latex: string) => void;
  onAppend: (latex: string) => void;
  onNotify: (message: string) => void;
}

function readError(error: unknown) {
  if (error instanceof Error) return error.message;
  if (typeof error === "string") return error;
  return "Unknown OCR error";
}

function readableBytes(bytes: number) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function providerLabel(provider: WebOcrProvider) {
  switch (provider) {
    case "openai-compatible":
      return "OpenAI Compatible";
    case "mathpix":
      return "Mathpix";
    case "paddleocr":
      return "PaddleOCR AI Studio";
    case "simpletex":
      return "SimpleTex";
  }
}

function normalizeLatex(value: string) {
  return value
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .join("\n");
}

export function WebOcrDialog({
  open,
  language,
  onClose,
  onInsert,
  onAppend,
  onNotify,
}: WebOcrDialogProps) {
  const isEn = language === "en";
  const dialogRef = useRef<HTMLElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const objectUrlRef = useRef("");
  const [configuration, setConfiguration] = useState<WebOcrConfiguration>(
    loadWebOcrConfiguration,
  );
  const [file, setFile] = useState<File | null>(null);
  const [previewUrl, setPreviewUrl] = useState("");
  const [dragging, setDragging] = useState(false);
  const [imageSize, setImageSize] = useState({ width: 0, height: 0 });
  const [recognizing, setRecognizing] = useState(false);
  const [progress, setProgress] = useState<WebOcrProgress | null>(null);
  const [elapsedSeconds, setElapsedSeconds] = useState(0);
  const [latex, setLatex] = useState("");
  const [resultMeta, setResultMeta] = useState("");
  const [error, setError] = useState("");
  const [copied, setCopied] = useState(false);
  const [saved, setSaved] = useState(false);

  const revokePreview = useCallback(() => {
    if (!objectUrlRef.current) return;
    URL.revokeObjectURL(objectUrlRef.current);
    objectUrlRef.current = "";
  }, []);

  const setImage = useCallback(
    (nextFile: File) => {
      if (!nextFile.type.startsWith("image/")) {
        setError(
          isEn
            ? "Choose a PNG, JPEG, WebP, BMP, or TIFF image."
            : "请选择 PNG、JPEG、WebP、BMP 或 TIFF 图片。",
        );
        return;
      }
      if (nextFile.size > 20 * 1024 * 1024) {
        setError(isEn ? "Images must not exceed 20 MB." : "图片不能超过 20 MB。");
        return;
      }
      revokePreview();
      const url = URL.createObjectURL(nextFile);
      objectUrlRef.current = url;
      setFile(nextFile);
      setPreviewUrl(url);
      setImageSize({ width: 0, height: 0 });
      setLatex("");
      setResultMeta("");
      setError("");
    },
    [isEn, revokePreview],
  );

  useEffect(() => () => revokePreview(), [revokePreview]);

  useEffect(() => {
    if (!open) return;
    setConfiguration(loadWebOcrConfiguration());
    previousFocusRef.current =
      document.activeElement instanceof HTMLElement
        ? document.activeElement
        : null;
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current
        ?.querySelector<HTMLElement>("button, input, select, textarea")
        ?.focus({ preventScroll: true });
    });
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape" && !recognizing) {
        event.preventDefault();
        onClose();
      }
      if (
        (event.metaKey || event.ctrlKey) &&
        event.key.toLowerCase() === "v"
      ) {
        return;
      }
      if (event.key !== "Tab") return;
      const focusable = Array.from(
        dialogRef.current?.querySelectorAll<HTMLElement>(
          'button:not(:disabled), input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])',
        ) ?? [],
      );
      if (!focusable.length) return;
      const first = focusable[0];
      const last = focusable[focusable.length - 1];
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
  }, [onClose, open, recognizing]);

  useEffect(() => {
    if (!open) return;
    const handlePaste = (event: ClipboardEvent) => {
      const pasted = Array.from(event.clipboardData?.files ?? []).find((item) =>
        item.type.startsWith("image/"),
      );
      if (!pasted) return;
      event.preventDefault();
      setImage(pasted);
    };
    document.addEventListener("paste", handlePaste);
    return () => document.removeEventListener("paste", handlePaste);
  }, [open, setImage]);

  useEffect(() => {
    if (!recognizing) return;
    setElapsedSeconds(0);
    const timer = window.setInterval(
      () => setElapsedSeconds((value) => value + 1),
      1000,
    );
    return () => window.clearInterval(timer);
  }, [recognizing]);

  if (!open) return null;

  const activeProvider = configuration.activeProvider;
  const updateConfiguration = (
    update: (current: WebOcrConfiguration) => WebOcrConfiguration,
  ) => {
    setConfiguration((current) => update(current));
    setSaved(false);
  };

  const saveConfiguration = () => {
    try {
      saveWebOcrConfiguration(configuration);
      setSaved(true);
      setError("");
      onNotify(
        isEn
          ? "OCR API configuration saved for this browser session"
          : "OCR API 配置已保存到当前浏览器会话",
      );
    } catch (saveError) {
      setError(readError(saveError));
    }
  };

  const handleRecognize = async () => {
    if (!file || recognizing) return;
    setRecognizing(true);
    setError("");
    setLatex("");
    setResultMeta("");
    try {
      saveWebOcrConfiguration(configuration);
      setSaved(true);
      const result = await recognizeFormulaWithWebApi(
        file,
        configuration,
        setProgress,
      );
      const value = normalizeLatex(result.formulas.join("\n"));
      if (!value) throw new Error("OCR API returned no usable formula");
      setLatex(value);
      setResultMeta(
        `${providerLabel(result.provider)} · ${result.model} · ${result.elapsedMs} ms`,
      );
      setProgress(null);
    } catch (recognitionError) {
      setError(readError(recognitionError));
    } finally {
      setRecognizing(false);
    }
  };

  const handleCopy = async () => {
    const value = normalizeLatex(latex);
    if (!value) return;
    try {
      await navigator.clipboard.writeText(value);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1200);
    } catch {
      setError(
        isEn
          ? "The browser blocked clipboard access."
          : "浏览器阻止了剪贴板访问。",
      );
    }
  };

  const requestClose = () => {
    if (!recognizing) onClose();
  };

  return (
    <div
      className="modal-backdrop ocr-modal-backdrop"
      role="presentation"
      onMouseDown={requestClose}
    >
      <section
        ref={dialogRef}
        className="ocr-dialog web-ocr-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="web-ocr-dialog-title"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="dialog-header ocr-dialog-header">
          <div className="ocr-heading">
            <span className="ocr-heading-icon">
              <ScanLine size={20} />
            </span>
            <div>
              <span className="eyebrow">WEB OCR API</span>
              <h2 id="web-ocr-dialog-title">
                {isEn ? "Formula image recognition" : "图片公式识别"}
              </h2>
            </div>
          </div>
          <button
            type="button"
            className="icon-button"
            onClick={requestClose}
            disabled={recognizing}
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
              onChange={(event: ChangeEvent<HTMLInputElement>) => {
                const selected = event.target.files?.[0];
                if (selected) setImage(selected);
                event.target.value = "";
              }}
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
              onDrop={(event: DragEvent<HTMLDivElement>) => {
                event.preventDefault();
                setDragging(false);
                const selected = Array.from(event.dataTransfer.files).find(
                  (item) => item.type.startsWith("image/"),
                );
                if (selected) setImage(selected);
              }}
            >
              {previewUrl ? (
                <>
                  <img
                    src={previewUrl}
                    alt={isEn ? "Formula image preview" : "公式图片预览"}
                    onLoad={(event) =>
                      setImageSize({
                        width: event.currentTarget.naturalWidth,
                        height: event.currentTarget.naturalHeight,
                      })
                    }
                  />
                  <div className="ocr-image-actions">
                    <button
                      type="button"
                      onClick={() => fileInputRef.current?.click()}
                      disabled={recognizing}
                    >
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
                  <strong>
                    {isEn
                      ? "Drop a formula image here"
                      : "将公式图片拖到这里"}
                  </strong>
                  <span>
                    {isEn
                      ? "Choose a file or paste an image"
                      : "选择文件，或直接粘贴剪贴板图片"}
                  </span>
                  <button
                    type="button"
                    onClick={() => fileInputRef.current?.click()}
                  >
                    <Upload size={15} />
                    {isEn ? "Choose image" : "选择图片"}
                  </button>
                  <small>
                    <ClipboardPaste size={13} />
                    {isEn
                      ? "Paste with Ctrl/⌘V while this dialog is open"
                      : "窗口打开时可直接按 Ctrl/⌘V 粘贴"}
                  </small>
                </div>
              )}
            </div>
            {file && (
              <div className="ocr-file-meta">
                <span>
                  {file.name || (isEn ? "Clipboard image" : "剪贴板图片")}
                </span>
                <span>
                  {imageSize.width
                    ? `${imageSize.width}×${imageSize.height} · `
                    : ""}
                  {readableBytes(file.size)}
                </span>
              </div>
            )}

            <section
              className="ocr-provider-card"
              aria-label={isEn ? "OCR provider" : "OCR 提供器"}
            >
              <div className="ocr-provider-heading">
                <span className="ocr-provider-icon is-api">
                  <Cloud size={17} />
                </span>
                <div>
                  <strong>
                    {isEn ? "OCR API provider" : "OCR API 提供器"}
                  </strong>
                  <span>{providerLabel(activeProvider)}</span>
                </div>
              </div>
              <label className="ocr-provider-field is-wide">
                <span>{isEn ? "Provider" : "提供器"}</span>
                <select
                  value={activeProvider}
                  disabled={recognizing}
                  onChange={(event) =>
                    updateConfiguration((current) => ({
                      ...current,
                      activeProvider: event.target.value as WebOcrProvider,
                    }))
                  }
                >
                  <option value="simpletex">SimpleTex</option>
                  <option value="paddleocr">PaddleOCR AI Studio</option>
                  <option value="mathpix">Mathpix</option>
                  <option value="openai-compatible">
                    OpenAI Compatible
                  </option>
                </select>
              </label>
            </section>
          </div>

          <div className="ocr-output-column">
            <section className="ocr-provider-card web-ocr-provider-settings">
              <div className="ocr-provider-heading">
                <span className="ocr-provider-icon is-api">
                  <ShieldCheck size={17} />
                </span>
                <div>
                  <strong>
                    {isEn ? "API configuration" : "API 配置"}
                  </strong>
                  <span>
                    {isEn
                      ? "Secrets stay in this tab session"
                      : "密钥只保留在当前标签页会话"}
                  </span>
                </div>
              </div>

              <div className="ocr-provider-fields">
                {activeProvider === "simpletex" && (
                  <>
                    <label className="ocr-provider-field">
                      <span>{isEn ? "Model" : "模型"}</span>
                      <select
                        value={configuration.simpleTex.model}
                        disabled={recognizing}
                        onChange={(event) =>
                          updateConfiguration((current) => ({
                            ...current,
                            simpleTex: {
                              ...current.simpleTex,
                              model: event.target.value as "standard" | "turbo",
                            },
                          }))
                        }
                      >
                        <option value="standard">
                          {isEn ? "Standard · accuracy" : "标准模型 · 精度优先"}
                        </option>
                        <option value="turbo">
                          {isEn ? "Turbo · speed" : "轻量模型 · 速度优先"}
                        </option>
                      </select>
                    </label>
                    <label className="ocr-provider-field">
                      <span>SimpleTex UAT</span>
                      <input
                        type="password"
                        autoComplete="off"
                        value={configuration.simpleTex.accessToken}
                        disabled={recognizing}
                        placeholder={isEn ? "User access token" : "用户访问令牌"}
                        onChange={(event) =>
                          updateConfiguration((current) => ({
                            ...current,
                            simpleTex: {
                              ...current.simpleTex,
                              accessToken: event.target.value,
                            },
                          }))
                        }
                      />
                    </label>
                    <p className="ocr-provider-protocol-note">
                      {isEn
                        ? "Uploads one image directly to SimpleTex V2.5 and reads res.latex."
                        : "图片由浏览器直接上传到 SimpleTex V2.5，并读取 res.latex。"}
                    </p>
                  </>
                )}

                {activeProvider === "paddleocr" && (
                  <>
                    <label className="ocr-provider-field">
                      <span>{isEn ? "Model" : "模型"}</span>
                      <input value="PaddleOCR-VL-1.6" disabled />
                    </label>
                    <label className="ocr-provider-field">
                      <span>Access Token</span>
                      <input
                        type="password"
                        autoComplete="off"
                        value={configuration.paddleOcr.accessToken}
                        disabled={recognizing}
                        placeholder="AI Studio Access Token"
                        onChange={(event) =>
                          updateConfiguration((current) => ({
                            ...current,
                            paddleOcr: {
                              ...current.paddleOcr,
                              accessToken: event.target.value,
                            },
                          }))
                        }
                      />
                    </label>
                    <p className="ocr-provider-protocol-note">
                      {isEn
                        ? "Uses PaddleOCR-VL-1.6 and waits for the normal asynchronous queue for up to 120 seconds."
                        : "使用 PaddleOCR-VL-1.6，并正常等待异步队列，最长 120 秒。"}
                    </p>
                  </>
                )}

                {activeProvider === "mathpix" && (
                  <>
                    <label className="ocr-provider-field is-wide">
                      <span>{isEn ? "Mathpix address" : "Mathpix 地址"}</span>
                      <input
                        value={configuration.mathpix.baseUrl}
                        disabled={recognizing}
                        onChange={(event) =>
                          updateConfiguration((current) => ({
                            ...current,
                            mathpix: {
                              ...current.mathpix,
                              baseUrl: event.target.value,
                            },
                          }))
                        }
                      />
                    </label>
                    <label className="ocr-provider-field">
                      <span>app_id</span>
                      <input
                        value={configuration.mathpix.appId}
                        disabled={recognizing}
                        autoComplete="off"
                        onChange={(event) =>
                          updateConfiguration((current) => ({
                            ...current,
                            mathpix: {
                              ...current.mathpix,
                              appId: event.target.value,
                            },
                          }))
                        }
                      />
                    </label>
                    <label className="ocr-provider-field">
                      <span>app_key</span>
                      <input
                        type="password"
                        value={configuration.mathpix.appKey}
                        disabled={recognizing}
                        autoComplete="off"
                        onChange={(event) =>
                          updateConfiguration((current) => ({
                            ...current,
                            mathpix: {
                              ...current.mathpix,
                              appKey: event.target.value,
                            },
                          }))
                        }
                      />
                    </label>
                    <p className="ocr-provider-protocol-note">
                      {isEn
                        ? "Uses POST /v3/text. Mathpix limits base64 images to 2 MB."
                        : "使用 POST /v3/text；Mathpix 的 base64 图片上限为 2 MB。"}
                    </p>
                  </>
                )}

                {activeProvider === "openai-compatible" && (
                  <>
                    <label className="ocr-provider-field">
                      <span>{isEn ? "Protocol" : "协议"}</span>
                      <select
                        value={configuration.openAiCompatible.protocol}
                        disabled={recognizing}
                        onChange={(event) =>
                          updateConfiguration((current) => ({
                            ...current,
                            openAiCompatible: {
                              ...current.openAiCompatible,
                              protocol: event.target.value as
                                | "responses"
                                | "chat-completions",
                            },
                          }))
                        }
                      >
                        <option value="responses">Responses API</option>
                        <option value="chat-completions">
                          Chat Completions
                        </option>
                      </select>
                    </label>
                    <label className="ocr-provider-field">
                      <span>{isEn ? "Model" : "模型"}</span>
                      <input
                        value={configuration.openAiCompatible.model}
                        disabled={recognizing}
                        autoComplete="off"
                        onChange={(event) =>
                          updateConfiguration((current) => ({
                            ...current,
                            openAiCompatible: {
                              ...current.openAiCompatible,
                              model: event.target.value,
                            },
                          }))
                        }
                      />
                    </label>
                    <label className="ocr-provider-field is-wide">
                      <span>Base URL</span>
                      <input
                        value={configuration.openAiCompatible.baseUrl}
                        disabled={recognizing}
                        autoComplete="off"
                        onChange={(event) =>
                          updateConfiguration((current) => ({
                            ...current,
                            openAiCompatible: {
                              ...current.openAiCompatible,
                              baseUrl: event.target.value,
                            },
                          }))
                        }
                      />
                    </label>
                    <label className="ocr-provider-field is-wide">
                      <span>API Key</span>
                      <input
                        type="password"
                        value={configuration.openAiCompatible.apiKey}
                        disabled={recognizing}
                        autoComplete="off"
                        onChange={(event) =>
                          updateConfiguration((current) => ({
                            ...current,
                            openAiCompatible: {
                              ...current.openAiCompatible,
                              apiKey: event.target.value,
                            },
                          }))
                        }
                      />
                    </label>
                    <label className="ocr-provider-field is-wide">
                      <span>{isEn ? "Recognition prompt" : "识别提示词"}</span>
                      <textarea
                        value={configuration.openAiCompatible.prompt}
                        disabled={recognizing}
                        spellCheck={false}
                        onChange={(event) =>
                          updateConfiguration((current) => ({
                            ...current,
                            openAiCompatible: {
                              ...current.openAiCompatible,
                              prompt: event.target.value,
                            },
                          }))
                        }
                      />
                    </label>
                  </>
                )}
              </div>

              <div className="ocr-provider-actions">
                <span>
                  <ShieldCheck size={14} />
                  {activeProvider === "simpletex" || activeProvider === "paddleocr"
                    ? isEn
                      ? "This provider blocks browser CORS, so requests use VisualTeX's fixed-target relay. The relay accepts no arbitrary URL and does not log or store images, formulas, or credentials."
                      : "该服务商阻止浏览器跨域调用，因此请求经 VisualTeX 固定目标转发；转发层不接受任意网址，也不记录或存储图片、公式与密钥。"
                    : isEn
                      ? "Images and credentials are sent directly from your browser to the selected provider and are not uploaded to VisualTeX servers."
                      : "图片和密钥由浏览器直接发送到所选服务商，不上传到 VisualTeX 服务器。"}
                </span>
                <button
                  type="button"
                  className="secondary-button"
                  onClick={saveConfiguration}
                  disabled={recognizing}
                >
                  {saved ? <Check size={15} /> : <Save size={15} />}
                  {saved
                    ? isEn
                      ? "Saved"
                      : "已保存"
                    : isEn
                      ? "Save session"
                      : "保存本次会话"}
                </button>
              </div>
            </section>

            <section className="ocr-result-card">
              <div className="ocr-result-heading">
                <div>
                  <span className="eyebrow">LATEX RESULT</span>
                  <strong>
                    {isEn ? "Recognition result" : "识别结果"}
                  </strong>
                </div>
                {resultMeta && <span>{resultMeta}</span>}
              </div>
              {recognizing ? (
                <div className="ocr-recognizing-state">
                  <LoaderCircle size={24} className="is-spinning" />
                  <strong>
                    {progress
                      ? isEn
                        ? progress.messageEn
                        : progress.messageZh
                      : isEn
                        ? "Preparing OCR request…"
                        : "正在准备 OCR 请求…"}
                  </strong>
                  <span>
                    {providerLabel(activeProvider)} · {elapsedSeconds}s
                  </span>
                </div>
              ) : latex ? (
                <>
                  <div className="ocr-formula-preview">
                    <MathPreview latex={latex.split("\n")[0]} />
                  </div>
                  <label className="ocr-latex-editor">
                    <span>{isEn ? "Editable LaTeX" : "可编辑 LaTeX"}</span>
                    <textarea
                      value={latex}
                      onChange={(event) => setLatex(event.target.value)}
                      spellCheck={false}
                    />
                  </label>
                </>
              ) : (
                <div className="ocr-empty-result">
                  <ScanLine size={24} />
                  <span>
                    {isEn
                      ? "Configure an API, choose an image, and run recognition."
                      : "配置 API 并选择图片后，即可开始识别。"}
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
          <button
            type="button"
            className="secondary-button"
            onClick={() => void handleRecognize()}
            disabled={!file || recognizing}
          >
            {recognizing ? (
              <LoaderCircle size={15} className="is-spinning" />
            ) : (
              <ScanLine size={15} />
            )}
            {recognizing
              ? isEn
                ? "Recognizing…"
                : "正在识别…"
              : isEn
                ? "Recognize"
                : "开始识别"}
          </button>
          <div className="ocr-result-actions">
            <button
              type="button"
              className="secondary-button"
              onClick={() => void handleCopy()}
              disabled={!latex.trim() || recognizing}
            >
              {copied ? <Check size={15} /> : <Copy size={15} />}
              {copied
                ? isEn
                  ? "Copied"
                  : "已复制"
                : isEn
                  ? "Copy LaTeX"
                  : "复制 LaTeX"}
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={!latex.trim() || recognizing}
              onClick={() => {
                const value = normalizeLatex(latex);
                if (!value) return;
                onAppend(value);
                onNotify(
                  isEn
                    ? "OCR formula appended as new formula rows"
                    : "OCR 公式已追加为新公式行",
                );
                onClose();
              }}
            >
              <Plus size={15} />
              {isEn ? "Append rows" : "追加为新行"}
            </button>
            <button
              type="button"
              className="primary-button"
              disabled={!latex.trim() || recognizing}
              onClick={() => {
                const value = normalizeLatex(latex);
                if (!value) return;
                onInsert(value);
                onNotify(
                  isEn
                    ? "OCR formula inserted at the cursor"
                    : "OCR 公式已插入当前光标",
                );
                onClose();
              }}
            >
              <ScanLine size={15} />
              {isEn ? "Insert at cursor" : "插入当前光标"}
            </button>
          </div>
        </footer>
      </section>
    </div>
  );
}
