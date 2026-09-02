import {
  Component,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent,
  type ErrorInfo,
  type ReactNode,
} from "react";
import {
  AlertCircle,
  Braces,
  Check,
  CheckCircle2,
  Eye,
  FileText,
  Image as ImageIcon,
  Upload,
  LoaderCircle,
  Sigma,
  X,
} from "lucide-react";
import { MathPreview } from "../../components/MathPreview";
import { onCurrentTauriWindowCloseRequested } from "../shared/tauriTransport";
import {
  createFormulaMetadata,
  type VisualTeXFormulaMetadata,
} from "../shared/formulaMetadata";
import {
  normalizeFormulaEditorDocument,
} from "../shared/formulaEditorDocument";
import {
  OFFICE_FORMULA_REFERENCE_FONT_SIZE_PT,
  renderOfficeFormulaArtifacts,
} from "../shared/formulaRenderArtifacts";
import { createUuid } from "../../runtime/browserCompatibility";
import { useEditorStore } from "../../stores/editorStore";
import { documentImportErrorMessage } from "./documentImportErrors";
import {
  cancelMacosDocumentImport,
  closeMacosDocumentImportWindow,
  commitMacosDocumentImport,
  focusMacosDocumentImportTarget,
  getMacosDocumentImportProgress,
  getMacosDocumentImportRequest,
  restoreMacosDocumentImportWindow,
  type DocumentImportCommitItem,
  type MacosDocumentImportRequest,
} from "./documentImportClient";
import {
  type ImportedDocumentFile,
  readDocumentImportFile,
} from "./documentImportFile.ts";
import {
  mergeDocumentImportBlocks,
  parseLatexMarkdownDocument,
  type DocumentFormulaBlock,
  type DocumentFormulaOutputKind,
  type DocumentImportBlock,
  type DocumentImportSourceKind,
} from "./documentImportParser";

const MAX_WORD_REFERENCE_WIDTH_PT = 500;
const WORD_IMAGE_VISUAL_SCALE = 1.1;
const REFERENCE_FONT_SIZE_PT = OFFICE_FORMULA_REFERENCE_FONT_SIZE_PT;

type ImportedFileState = Pick<
  ImportedDocumentFile,
  "name" | "encoding" | "size"
> & { modified: boolean };

function formatFileSize(bytes: number) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

class FormulaPreviewBoundary extends Component<
  { children: ReactNode; message: string },
  { failed: boolean }
> {
  state = { failed: false };

  static getDerivedStateFromError() {
    return { failed: true };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    console.warn("VisualTeX document formula preview failed", error, info);
  }

  componentDidUpdate(previous: { children: ReactNode; message: string }) {
    if (this.state.failed && previous.children !== this.props.children) {
      this.setState({ failed: false });
    }
  }

  render() {
    if (this.state.failed) {
      return <span className="document-import-formula-error">{this.props.message}</span>;
    }
    return this.props.children;
  }
}

function clampFontSize(value: number, fallback = 12) {
  if (!Number.isFinite(value)) return fallback;
  return Math.min(512, Math.max(1, Math.round(value * 2) / 2));
}

function formulaCount(blocks: DocumentImportBlock[]) {
  return blocks.reduce((count, block) => count + (block.kind === "formula" ? 1 : 0), 0);
}

function textCharacterCount(blocks: DocumentImportBlock[]) {
  return blocks.reduce(
    (count, block) => count + (block.kind === "text" ? block.text.length : 0),
    0,
  );
}

function decodeUrlSafeBase64Utf8(value: string) {
  const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
  const binary = atob(padded);
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

function ommlRetainsLiteralLatexCommand(ommlBase64: string, latex: string) {
  const commands = [...latex.matchAll(/\\([A-Za-z@]+)\b/g)]
    .map((match) => match[1])
    .filter((command, index, values) => values.indexOf(command) === index);
  if (!commands.length) return false;
  const omml = decodeUrlSafeBase64Utf8(ommlBase64);
  return commands.some((command) => omml.includes(`\\${command}`));
}

function calculateReferenceGeometry(widthPx: number, heightPx: number, baselinePx: number) {
  const naturalWidthPt = widthPx * 0.75 * WORD_IMAGE_VISUAL_SCALE;
  const naturalHeightPt = heightPx * 0.75 * WORD_IMAGE_VISUAL_SCALE;
  const scale = Math.min(1, MAX_WORD_REFERENCE_WIDTH_PT / naturalWidthPt);
  const referenceWidthPt = naturalWidthPt * scale;
  const referenceHeightPt = naturalHeightPt * scale;
  const descentRatio = Math.max(0, Math.min(1, (heightPx - baselinePx) / heightPx));
  // Keep the exported descent as a fractional 14 pt reference. Word accepts
  // only an integer Font.Position, so rounding here and again after scaling to
  // the requested font size makes short subscript formulas (for example L_z)
  // lose almost a full point relative to superscript formulas such as L^2.
  // Round exactly once at the final Word dispatch boundary instead.
  const referenceBaselinePt = -Math.max(
    0,
    referenceHeightPt * descentRatio,
  );
  return { referenceWidthPt, referenceHeightPt, referenceBaselinePt };
}

async function prepareFormulaArtifactCommitItem(
  block: DocumentFormulaBlock,
  outputKind: DocumentFormulaOutputKind,
): Promise<DocumentImportCommitItem> {
  const formulaId = createUuid();
  const line = { id: createUuid(), latex: block.latex.trim() };
  if (!line.latex) throw new Error("存在空公式，请填写或删除后再插入。");
  const editorDocument = normalizeFormulaEditorDocument([line], "raw");
  const { formulaLetterFont, formulaChineseFont } = useEditorStore.getState();
  const artifacts = renderOfficeFormulaArtifacts({
    lines: editorDocument.lines,
    codeFormat: editorDocument.codeFormat,
    displayMode: block.displayMode,
    host: "word",
    formulaLetterFont,
    formulaChineseFont,
  });
  const { canonicalLatex, svg } = artifacts;
  if (!artifacts.omml) {
    throw new Error("无法生成 Word OMML 公式制品。");
  }
  const omml = artifacts.omml;
  if (ommlRetainsLiteralLatexCommand(omml.ommlBase64, canonicalLatex)) {
    throw new Error("公式包含未被 Word 公式转换器识别的自定义命令。");
  }

  const paragraphMetadata = {
    paragraphId: block.paragraphId,
    paragraphStyle: block.paragraphStyle,
    paragraphAlignment: block.paragraphAlignment,
    listKind: block.listKind,
    listLevel: block.listLevel,
    paragraphStart: block.paragraphStart,
    paragraphEnd: block.paragraphEnd,
  };

  let metadata: VisualTeXFormulaMetadata;
  if (outputKind === "image") {
    const { svgToPng } = await import("../../export/svgToPng");
    const png = await svgToPng(svg, { scale: 2, background: "transparent" });
    const pngBase64 = png.base64;
    const resolvedBaseline = svg.baseline;
    const reference = calculateReferenceGeometry(
      svg.width,
      svg.height,
      resolvedBaseline,
    );
    metadata = createFormulaMetadata({
      formulaId,
      title: block.displayMode === "inline" ? "Imported inline formula" : "Imported display formula",
      lines: editorDocument.lines,
      codeFormat: editorDocument.codeFormat,
      sourceLatex: canonicalLatex,
      displayMode: block.displayMode,
      numbered: block.displayMode === "block" && block.numbered,
      fontSizePt: block.fontSizePt,
      formulaLetterFont,
      formulaChineseFont,
      renderWidthPx: svg.width,
      renderHeightPx: svg.height,
      imageInkCenterYRatio: png.inkCenterYRatio,
      ...reference,
    });
    return {
      kind: "formula",
      formulaId,
      latex: canonicalLatex,
      displayMode: block.displayMode,
      numbered: block.displayMode === "block" && block.numbered,
      fontSizePt: block.fontSizePt,
      metadata,
      ommlBase64: omml.ommlBase64,
      ommlDocxBase64: omml.ommlDocxBase64,
      svgBase64: svg.base64,
      pngBase64,
      width: svg.width,
      height: svg.height,
      baseline: resolvedBaseline,
      inkCenterYRatio: png.inkCenterYRatio,
      ...paragraphMetadata,
    };
  }

  metadata = createFormulaMetadata({
    formulaId,
    title: block.displayMode === "inline" ? "Imported inline formula" : "Imported display formula",
    lines: editorDocument.lines,
    codeFormat: editorDocument.codeFormat,
    sourceLatex: canonicalLatex,
    displayMode: block.displayMode,
    numbered: block.displayMode === "block" && block.numbered,
    fontSizePt: block.fontSizePt,
    formulaLetterFont,
    formulaChineseFont,
  });
  return {
    kind: "formula",
    formulaId,
    latex: canonicalLatex,
    displayMode: block.displayMode,
    numbered: block.displayMode === "block" && block.numbered,
    fontSizePt: block.fontSizePt,
    metadata,
    ommlBase64: omml.ommlBase64,
    ommlDocxBase64: omml.ommlDocxBase64,
    ...paragraphMetadata,
  };
}

function formulaLiteralFallbackText(block: DocumentFormulaBlock) {
  const original = block.sourceText?.trim();
  if (original) return original;
  const latex = block.latex.trim();
  if (block.displayMode === "inline") return `\\(${latex}\\)`;
  if (/^\\begin\s*\{[^{}]+\}/.test(latex)) return latex;
  return `\\[\n${latex}\n\\]`;
}

async function prepareFormulaCommitItem(
  block: DocumentFormulaBlock,
  outputKind: DocumentFormulaOutputKind,
): Promise<DocumentImportCommitItem> {
  try {
    return await prepareFormulaArtifactCommitItem(block, outputKind);
  } catch (reason) {
    console.warn(
      "VisualTeX preserved an unsupported document formula as literal text",
      reason,
      block.latex,
    );
    const paragraphMetadata = block.paragraphId
      ? {
          paragraphId: block.paragraphId,
          paragraphStyle: block.paragraphStyle,
          paragraphAlignment: block.paragraphAlignment,
          listKind: block.listKind,
          listLevel: block.listLevel,
          paragraphStart: block.paragraphStart,
          paragraphEnd: block.paragraphEnd,
        }
      : {
          paragraphId: createUuid(),
          paragraphStyle: "code" as const,
          paragraphAlignment: "left" as const,
          listKind: "none" as const,
          listLevel: 0,
          paragraphStart: true,
          paragraphEnd: true,
        };
    return {
      kind: "text",
      text: formulaLiteralFallbackText(block),
      ...paragraphMetadata,
    };
  }
}

export function OfficeDocumentImportApp() {
  const sessionId = useMemo(
    () => new URLSearchParams(window.location.search).get("sessionId") ?? "",
    [],
  );
  const [request, setRequest] = useState<MacosDocumentImportRequest | null>(null);
  const [source, setSource] = useState("");
  const [sourceKind, setSourceKind] = useState<DocumentImportSourceKind>("auto");
  const [outputKind, setOutputKind] = useState<DocumentFormulaOutputKind>("omml");
  const [blocks, setBlocks] = useState<DocumentImportBlock[]>([]);
  const [importedFile, setImportedFile] = useState<ImportedFileState | null>(null);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [toast, setToast] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const sourceRef = useRef<HTMLTextAreaElement>(null);
  const allowNativeCloseRef = useRef(false);
  const nativeCloseInFlightRef = useRef(false);

  useEffect(() => {
    if (!sessionId) {
      setError("缺少 Word 文档导入会话标识。");
      setLoading(false);
      return;
    }
    void getMacosDocumentImportRequest(sessionId)
      .then((value) => setRequest(value))
      .catch((reason) =>
        setError(documentImportErrorMessage(reason, "无法读取 Word 文档导入请求。")),
      )
      .finally(() => setLoading(false));
  }, [sessionId]);

  useEffect(() => {
    if (loading) return;
    const frame = window.requestAnimationFrame(() => sourceRef.current?.focus());
    return () => window.cancelAnimationFrame(frame);
  }, [loading]);

  const reparse = useCallback(
    (nextSource: string, nextKind = sourceKind) => {
      const parsed = parseLatexMarkdownDocument(
        nextSource,
        nextKind,
        request?.defaultFontSizePt ?? 12,
      );
      setBlocks((previous) => mergeDocumentImportBlocks(previous, parsed));
    },
    [request?.defaultFontSizePt, sourceKind],
  );

  const updateFormula = useCallback(
    (id: string, update: Partial<Omit<DocumentFormulaBlock, "id" | "kind">>) => {
      setBlocks((current) =>
        current.map((block) => {
          if (block.kind !== "formula" || block.id !== id) return block;
          const updated = { ...block, ...update };
          if (
            ("latex" in update && update.latex !== block.latex) ||
            ("displayMode" in update && update.displayMode !== block.displayMode)
          ) {
            delete updated.sourceText;
          }
          return updated;
        }),
      );
    },
    [],
  );

  const handleSourceChange = (value: string) => {
    setSource(value);
    setImportedFile((current) =>
      current ? { ...current, modified: true } : current,
    );
    setError("");
    reparse(value);
  };

  const handleDocumentFile = async (file: File | null) => {
    if (!file || busy) return;
    setError("");
    setToast("正在读取文档源码…");
    try {
      const imported = await readDocumentImportFile(file);
      setSource(imported.source);
      setSourceKind(imported.format);
      setImportedFile({
        name: imported.name,
        encoding: imported.encoding,
        size: imported.size,
        modified: false,
      });
      const parsed = parseLatexMarkdownDocument(
        imported.source,
        imported.format,
        request?.defaultFontSizePt ?? 12,
      );
      setBlocks((previous) => mergeDocumentImportBlocks(previous, parsed));
      setToast(
        `已加载 ${imported.name} · ${imported.encoding} · ${formatFileSize(imported.size)}`,
      );
      window.requestAnimationFrame(() => sourceRef.current?.focus());
    } catch (reason) {
      setError(documentImportErrorMessage(reason, "无法读取文档源码文件。"));
      setToast("");
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  };

  const handleFileInput = (event: ChangeEvent<HTMLInputElement>) => {
    void handleDocumentFile(event.currentTarget.files?.[0] ?? null);
  };

  const handleSourceKindChange = (value: DocumentImportSourceKind) => {
    setSourceKind(value);
    setError("");
    const parsed = parseLatexMarkdownDocument(
      source,
      value,
      request?.defaultFontSizePt ?? 12,
    );
    setBlocks((previous) => mergeDocumentImportBlocks(previous, parsed));
  };

  const cancel = async () => {
    if (busy || nativeCloseInFlightRef.current) return;
    nativeCloseInFlightRef.current = true;
    setBusy(true);
    setError("");
    try {
      if (sessionId) await cancelMacosDocumentImport(sessionId);
      allowNativeCloseRef.current = true;
      await closeMacosDocumentImportWindow();
    } catch (reason) {
      setError(documentImportErrorMessage(reason, "无法取消文档导入。"));
      setBusy(false);
      nativeCloseInFlightRef.current = false;
    }
  };

  const commit = async () => {
    if (!request || busy) return;
    if (!blocks.length || blocks.every((block) => block.kind === "text" && !block.text.trim())) {
      setError("请先粘贴包含文字或公式的 LaTeX/Markdown 内容。");
      return;
    }
    const formulas = blocks.filter(
      (block): block is DocumentFormulaBlock => block.kind === "formula",
    );
    if (formulas.some((block) => !block.latex.trim())) {
      setError("存在空公式，请填写公式内容后再插入。");
      return;
    }

    setBusy(true);
    setError("");
    setToast(outputKind === "omml" ? "正在生成 Word 原生公式…" : "正在生成 SVG 图片公式…");
    let importerHidden = false;
    let progressTimer: number | undefined;
    let progressRequestInFlight = false;
    try {
      await focusMacosDocumentImportTarget();
      importerHidden = true;
      const preparedFormulas = await Promise.all(
        formulas.map(async (block, index) => {
          try {
            return await prepareFormulaCommitItem(block, outputKind);
          } catch (reason) {
            const detail = documentImportErrorMessage(
              reason,
              "未知公式转换错误。",
            );
            const preview = block.latex.trim().replace(/\s+/g, " ").slice(0, 120);
            throw new Error(
              `公式 ${index + 1} 生成失败：${detail}${preview ? `（${preview}）` : ""}`,
              { cause: reason },
            );
          }
        }),
      );
      let formulaIndex = 0;
      const items: DocumentImportCommitItem[] = blocks.map((block) => {
        if (block.kind === "text") {
          return {
            kind: "text",
            text: block.text,
            paragraphId: block.paragraphId,
            paragraphStyle: block.paragraphStyle,
            paragraphAlignment: block.paragraphAlignment,
            listKind: block.listKind,
            listLevel: block.listLevel,
            paragraphStart: block.paragraphStart,
            paragraphEnd: block.paragraphEnd,
          };
        }
        const prepared = preparedFormulas[formulaIndex];
        formulaIndex += 1;
        return prepared;
      });
      const literalFallbackCount = preparedFormulas.filter(
        (item) => item.kind === "text",
      ).length;
      setToast(
        literalFallbackCount > 0
          ? `正在写入 Word（${literalFallbackCount} 个不支持片段按原文保留）…`
          : `正在写入 Word：0/${items.length}`,
      );
      progressTimer = window.setInterval(() => {
        if (progressRequestInFlight) return;
        progressRequestInFlight = true;
        void getMacosDocumentImportProgress(sessionId)
          .then((progress) => {
            if (progress.total > 0 && progress.stage === "inserting") {
              setToast(`正在写入 Word：${progress.current}/${progress.total}`);
            }
          })
          .catch(() => undefined)
          .finally(() => {
            progressRequestInFlight = false;
          });
      }, 120);
      await commitMacosDocumentImport(sessionId, { outputKind, items });
      if (progressTimer !== undefined) window.clearInterval(progressTimer);
      setToast(`已完成：${items.length}/${items.length}`);
      allowNativeCloseRef.current = true;
      await closeMacosDocumentImportWindow();
    } catch (reason) {
      if (progressTimer !== undefined) window.clearInterval(progressTimer);
      if (importerHidden) {
        await restoreMacosDocumentImportWindow().catch(() => undefined);
      }
      setError(documentImportErrorMessage(reason, "无法将内容插入 Word。"));
      setToast("");
      setBusy(false);
    }
  };

  useEffect(() => {
    if (!sessionId) return;
    let disposed = false;
    let unlisten: (() => void) | undefined;
    void onCurrentTauriWindowCloseRequested((event) => {
      if (disposed || allowNativeCloseRef.current) return;
      event.preventDefault();
      void cancel();
    })
      .then((dispose) => {
        if (disposed) dispose();
        else unlisten = dispose;
      })
      .catch((reason) => {
        setError(documentImportErrorMessage(reason, "无法注册文档导入窗口关闭处理。"));
      });
    return () => {
      disposed = true;
      unlisten?.();
    };
  }, [sessionId, busy]);

  if (loading) {
    return (
      <main className="document-import-state">
        <LoaderCircle className="is-spinning" />
        <span>正在准备 Word 文档导入器…</span>
      </main>
    );
  }

  if (!request) {
    return (
      <main className="document-import-state is-error" role="alert">
        <AlertCircle />
        <strong>无法打开文档导入器</strong>
        <p>{error || "Word 文档导入请求不存在或已经失效。"}</p>
      </main>
    );
  }

  const formulas = formulaCount(blocks);
  const textCharacters = textCharacterCount(blocks);

  return (
    <main className="doc-import-shell macos-doc-import">
      <header className="doc-import-toolbar">
        <div className="doc-import-title-block">
          <FileText size={20} />
          <div>
            <strong>Word 文档批量导入</strong>
            <span>左侧编辑源码，右侧实时查看并调整最终 Word 结构</span>
          </div>
        </div>
        <div className="doc-import-options">
          <label>
            <span>源格式</span>
            <select
              value={sourceKind}
              onChange={(event) =>
                handleSourceKindChange(event.target.value as DocumentImportSourceKind)
              }
              disabled={busy}
            >
              <option value="auto">自动识别</option>
              <option value="latex">LaTeX</option>
              <option value="markdown">Markdown</option>
            </select>
          </label>
          <label>
            <span>公式格式</span>
            <select
              value={outputKind}
              onChange={(event) =>
                setOutputKind(event.target.value as DocumentFormulaOutputKind)
              }
              disabled={busy}
            >
              <option value="omml">Word 原生 OMML</option>
              <option value="image">SVG 图片公式</option>
            </select>
          </label>
          <button
            type="button"
            className="doc-import-secondary doc-import-file-button"
            onClick={() => fileInputRef.current?.click()}
            disabled={busy}
            title="导入单个 LaTeX 或 Markdown 文件"
          >
            <Upload size={16} />
            导入 .tex / .md
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".tex,.md,.markdown,text/x-tex,text/markdown"
            aria-label="导入 LaTeX 或 Markdown 文件"
            hidden
            onChange={handleFileInput}
          />
        </div>
      </header>

      <section className="doc-import-workspace">
        <article className="doc-import-pane source-pane">
          <div className="doc-import-pane-header">
            <div className="doc-import-pane-heading">
              <span className="doc-import-pane-icon" aria-hidden="true">
                <Braces size={16} />
              </span>
              <div>
                <strong>LaTeX / Markdown 源码</strong>
                <small>支持正文、标题、列表、定理、引用、代码块和混合公式</small>
              </div>
            </div>
            <div className="doc-import-source-meta">
              {importedFile ? (
                <span
                  className="doc-import-file-chip"
                  title={`${importedFile.name} · ${importedFile.encoding} · ${formatFileSize(importedFile.size)}`}
                >
                  <FileText size={12} />
                  <span>{importedFile.name}</span>
                  <small>
                    {importedFile.encoding}
                    {importedFile.modified ? " · 已编辑" : ""}
                  </small>
                </span>
              ) : null}
              <span className="doc-import-pane-stat">
                {source.length.toLocaleString()} 字符
              </span>
            </div>
          </div>
          <textarea
            ref={sourceRef}
            value={source}
            onChange={(event) => handleSourceChange(event.target.value)}
            placeholder={String.raw`在这里粘贴 LaTeX 或 Markdown，例如：

正文中的行内公式 $E=mc^2$。

\begin{equation}
\begin{aligned}
a&=b\\
c&=d
\end{aligned}
\end{equation}`}
            spellCheck={false}
            autoCapitalize="off"
            autoCorrect="off"
            disabled={busy}
            aria-label="文档源码"
          />
        </article>

        <article className="doc-import-pane preview-pane">
          <div className="doc-import-pane-header">
            <div className="doc-import-pane-heading">
              <span className="doc-import-pane-icon is-preview" aria-hidden="true">
                <Eye size={16} />
              </span>
              <div>
                <strong>Word 结构预览</strong>
                <small>公式卡片可单独调整行内/行间、编号和字号</small>
              </div>
            </div>
            <div className="doc-import-preview-counts" aria-label="预览统计">
              <span>{blocks.length} 块</span>
              <span>{textCharacters} 字</span>
              <span>{formulas} 公式</span>
            </div>
          </div>
          <div className="doc-import-preview-scroll">
            <div className="doc-import-preview-document">
              {!blocks.length ? (
                <div className="document-import-empty">
                  <FileText size={34} />
                  <strong>等待文档内容</strong>
                  <span>在左侧粘贴内容后，这里会实时生成 Word 结构预览。</span>
                </div>
              ) : (
                blocks.map((block, index) =>
                  block.kind === "text" ? (
                    <div
                      key={block.id}
                      className={`document-import-text-preview is-${block.paragraphStyle ?? "normal"} is-${block.listKind ?? "none"}`}
                      data-paragraph-start={block.paragraphStart ? "true" : undefined}
                    >
                      {block.paragraphStart && block.listKind === "bullet" ? (
                        <span className="document-import-list-marker">•</span>
                      ) : block.paragraphStart && block.listKind === "number" ? (
                        <span className="document-import-list-marker">1.</span>
                      ) : null}
                      {block.text}
                    </div>
                  ) : (
                    <article
                      key={block.id}
                      className={`document-import-formula-card is-${block.displayMode}`}
                    >
                      <header>
                        <span>公式 {index + 1}</span>
                        <div>
                          <select
                            value={block.displayMode}
                            onChange={(event) =>
                              updateFormula(block.id, {
                                displayMode: event.target.value as "inline" | "block",
                                numbered:
                                  event.target.value === "block" ? block.numbered : false,
                              })
                            }
                            disabled={busy}
                            aria-label="公式显示模式"
                          >
                            <option value="inline">行内公式</option>
                            <option value="block">行间公式</option>
                          </select>
                          {block.displayMode === "block" ? (
                            <label className="document-import-number-toggle">
                              <input
                                type="checkbox"
                                checked={block.numbered}
                                onChange={(event) =>
                                  updateFormula(block.id, { numbered: event.target.checked })
                                }
                                disabled={busy}
                              />
                              <span>编号</span>
                            </label>
                          ) : null}
                          <label>
                            <span>字号</span>
                            <input
                              type="number"
                              min="1"
                              max="512"
                              step="0.5"
                              value={block.fontSizePt}
                              onChange={(event) =>
                                updateFormula(block.id, {
                                  fontSizePt: clampFontSize(
                                    Number(event.target.value),
                                    block.fontSizePt,
                                  ),
                                })
                              }
                              disabled={busy}
                            />
                            <span>pt</span>
                          </label>
                        </div>
                      </header>
                      <div className="document-import-formula-preview">
                        <FormulaPreviewBoundary message="公式暂时无法预览，请检查 LaTeX。">
                          <MathPreview latex={block.latex || "\\placeholder{}"} />
                        </FormulaPreviewBoundary>
                      </div>
                      <textarea
                        value={block.latex}
                        onChange={(event) => updateFormula(block.id, { latex: event.target.value })}
                        spellCheck={false}
                        disabled={busy}
                        aria-label="编辑公式 LaTeX"
                      />
                    </article>
                  ),
                )
              )}
            </div>
          </div>
        </article>
      </section>

      <footer className="doc-import-footer">
        <div className="doc-import-messages">
          {error ? (
            <span className="error" role="alert"><AlertCircle size={15} />{error}</span>
          ) : toast ? (
            <span><LoaderCircle size={15} className="is-spinning" />{toast}</span>
          ) : (
            <span className="ok">
              <CheckCircle2 size={15} />
              预览解析正常；点击导入后将切回 Word 并实时显示插入进度。
            </span>
          )}
        </div>
        <div className="doc-import-actions">
          <button type="button" className="doc-import-secondary" onClick={() => void cancel()} disabled={busy}>
            <X size={16} />取消
          </button>
          <button
            type="button"
            className="doc-import-primary"
            onClick={() => void commit()}
            disabled={busy || !blocks.length}
          >
            {busy ? <LoaderCircle size={16} className="is-spinning" /> : <Check size={16} />}
            {busy ? "正在导入…" : "导入到 Word"}
          </button>
        </div>
      </footer>
    </main>
  );
}
