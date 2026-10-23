import { wordImageReferenceGeometry } from "../shared/wordImageGeometry";
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
  Upload,
  LoaderCircle,
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
import { renderOfficeFormulaArtifacts } from "../shared/formulaRenderArtifacts";
import { createUuid } from "../../runtime/browserCompatibility";
import { useEditorStore } from "../../stores/editorStore";
import { documentImportErrorMessage } from "./documentImportErrors";
import { describeConversionFailure } from "../shared/conversionErrors";
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

const isEnglish = () => useEditorStore.getState().language === "en";
const t = (zh: string, en: string) => (isEnglish() ? en : zh);

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

async function prepareFormulaArtifactCommitItem(
  block: DocumentFormulaBlock,
  _outputKind: DocumentFormulaOutputKind,
): Promise<DocumentImportCommitItem> {
  const formulaId = createUuid();
  const line = { id: createUuid(), latex: block.latex.trim() };
  if (!line.latex) throw new Error(t("存在空公式，请填写或删除后再插入。", "There is an empty formula. Fill it in or delete it before inserting."));
  const editorDocument = normalizeFormulaEditorDocument([line], "raw");
  const { formulaLetterFont, formulaChineseFont } = useEditorStore.getState();
  const artifacts = await renderOfficeFormulaArtifacts({
    lines: editorDocument.lines,
    codeFormat: editorDocument.codeFormat,
    displayMode: block.displayMode,
    host: "word",
    fontSizePt: block.fontSizePt,
    numbered: block.displayMode === "block" && block.numbered,
    formulaLetterFont,
    formulaChineseFont,
  });
  const { canonicalLatex, svg } = artifacts;
  // Windows' MathML step rejects unknown commands while parsing, so a formula
  // that reaches this point has no literal LaTeX command left in it.
  if (!artifacts.wordMathMl) {
    throw new Error(t("无法生成 Word OMML 公式制品。", "The Word equation (OMML) could not be generated."));
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

  // Prepare both Word representations once. OMML formulas need the real image
  // geometry and SVG/DOCX cache too, otherwise their first format conversion
  // falls back to the historical 14 x 14 placeholder square.
  const { svgToPng } = await import("../../export/svgToPng");
  const png = await svgToPng(svg, { scale: 2, background: "transparent" });
  const pngBase64 = png.base64;
  const resolvedBaseline = svg.baseline;
  const reference = wordImageReferenceGeometry(
    svg.width,
    svg.height,
    resolvedBaseline,
    formulaLetterFont,
  );
  const metadata: VisualTeXFormulaMetadata = createFormulaMetadata({
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
    wordMathMl: artifacts.wordMathMl,
    svgBase64: svg.base64,
    pngBase64,
    width: svg.width,
    height: svg.height,
    baseline: resolvedBaseline,
    inkCenterYRatio: png.inkCenterYRatio,
    ...paragraphMetadata,
  };
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
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const explicitFontSizeIds = useRef(new Set<string>());
  // Formulas take the size set in Word for converting their kind (inline or
  // display), otherwise the text size at the Word insertion point, unless the
  // user set a size for that formula here.
  const defaultFontSizePt = request?.defaultFontSizePt ?? 12;
  const inlineFontSizePt = request?.inlineFontSizePt ?? defaultFontSizePt;
  const displayFontSizePt = request?.displayFontSizePt ?? defaultFontSizePt;
  const conversionFontSize = useCallback(
    (displayMode: DocumentFormulaBlock["displayMode"]) =>
      displayMode === "block" ? displayFontSizePt : inlineFontSizePt,
    [displayFontSizePt, inlineFontSizePt],
  );
  const withConversionSizes = useCallback(
    (current: DocumentImportBlock[]) => current.map(block =>
      block.kind === "formula" && !explicitFontSizeIds.current.has(block.id)
        ? { ...block, fontSizePt: conversionFontSize(block.displayMode) } : block),
    [conversionFontSize],
  );
  useEffect(() => {
    setBlocks(withConversionSizes);
  }, [withConversionSizes]);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const sourceRef = useRef<HTMLTextAreaElement>(null);
  const allowNativeCloseRef = useRef(false);
  const nativeCloseInFlightRef = useRef(false);

  useEffect(() => {
    if (!sessionId) {
      setError(t("缺少 Word 文档导入会话标识。", "The Word document import session ID is missing."));
      setLoading(false);
      return;
    }
    void getMacosDocumentImportRequest(sessionId)
      .then((value) => setRequest(value))
      .catch((reason) =>
        setError(documentImportErrorMessage(reason, t("无法读取 Word 文档导入请求。", "Could not read the Word document import request."))),
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
        defaultFontSizePt,
      );
      setBlocks((previous) => withConversionSizes(mergeDocumentImportBlocks(previous, parsed)));
    },
    [defaultFontSizePt, sourceKind, withConversionSizes],
  );

  const updateFormula = useCallback(
    (id: string, update: Partial<Omit<DocumentFormulaBlock, "id" | "kind">>) => {
      if ("fontSizePt" in update) explicitFontSizeIds.current.add(id);
      setBlocks((current) =>
        current.map((block) => {
          if (block.kind !== "formula" || block.id !== id) return block;
          const updated = { ...block, ...update };
          // A formula switched between inline and display takes that kind's size.
          if (
            "displayMode" in update &&
            update.displayMode !== block.displayMode &&
            !explicitFontSizeIds.current.has(id)
          ) {
            updated.fontSizePt = conversionFontSize(updated.displayMode);
          }
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
    [conversionFontSize],
  );

  const setAllDisplayFormulaNumbering = useCallback((numbered: boolean) => {
    setBlocks((current) =>
      current.map((block) =>
        block.kind === "formula" && block.displayMode === "block"
          ? { ...block, numbered }
          : block,
      ),
    );
  }, []);

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
    setToast(t("正在读取文档源码…", "Reading the source file…"));
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
        defaultFontSizePt,
      );
      setBlocks((previous) => withConversionSizes(mergeDocumentImportBlocks(previous, parsed)));
      setToast(
        t(
          `已加载 ${imported.name} · ${imported.encoding} · ${formatFileSize(imported.size)}`,
          `Loaded ${imported.name} · ${imported.encoding} · ${formatFileSize(imported.size)}`,
        ),
      );
      window.requestAnimationFrame(() => sourceRef.current?.focus());
    } catch (reason) {
      setError(documentImportErrorMessage(reason, t("无法读取文档源码文件。", "Could not read the source file.")));
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
      defaultFontSizePt,
    );
    setBlocks((previous) => withConversionSizes(mergeDocumentImportBlocks(previous, parsed)));
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
      setError(documentImportErrorMessage(reason, t("无法取消文档导入。", "Could not cancel the import.")));
      setBusy(false);
      nativeCloseInFlightRef.current = false;
    }
  };

  const commit = async () => {
    if (!request || busy) return;
    if (!blocks.length || blocks.every((block) => block.kind === "text" && !block.text.trim())) {
      setError(t("请先粘贴包含文字或公式的 LaTeX/Markdown 内容。", "Paste LaTeX or Markdown that contains text or formulas first."));
      return;
    }
    const formulas = blocks.filter(
      (block): block is DocumentFormulaBlock => block.kind === "formula",
    );
    if (formulas.some((block) => !block.latex.trim())) {
      setError(t("存在空公式，请填写公式内容后再插入。", "There is an empty formula. Fill it in before inserting."));
      return;
    }

    setBusy(true);
    setError("");
    setToast(outputKind === "omml"
      ? t("正在生成 Word 原生公式…", "Generating native Word equations…")
      : t("正在生成 SVG 图片公式…", "Generating SVG image formulas…"));
    let importerHidden = false;
    let progressTimer: number | undefined;
    let progressRequestInFlight = false;
    try {
      await focusMacosDocumentImportTarget();
      importerHidden = true;
      const preparedFormulas = await Promise.all(
        formulas.map(async (block, index) => {
          try {
            return await prepareFormulaArtifactCommitItem(block, outputKind);
          } catch (reason) {
            const detail = describeConversionFailure(reason, isEnglish());
            const preview = block.latex.trim().replace(/\s+/g, " ").slice(0, 120);
            throw new Error(
              t(
                `公式 ${index + 1} 生成失败：${detail}${preview ? `（${preview}）` : ""}`,
                `Formula ${index + 1} could not be generated: ${detail}${preview ? ` (${preview})` : ""}`,
              ),
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
      setToast(t(`正在写入 Word：0/${items.length}`, `Writing to Word: 0/${items.length}`));
      progressTimer = window.setInterval(() => {
        if (progressRequestInFlight) return;
        progressRequestInFlight = true;
        void getMacosDocumentImportProgress(sessionId)
          .then((progress) => {
            if (progress.total > 0 && progress.stage === "inserting") {
              setToast(t(
                `正在写入 Word：${progress.current}/${progress.total}`,
                `Writing to Word: ${progress.current}/${progress.total}`,
              ));
            }
          })
          .catch(() => undefined)
          .finally(() => {
            progressRequestInFlight = false;
          });
      }, 120);
      await commitMacosDocumentImport(sessionId, { outputKind, items });
      if (progressTimer !== undefined) window.clearInterval(progressTimer);
      setToast(t(`已完成：${items.length}/${items.length}`, `Done: ${items.length}/${items.length}`));
      allowNativeCloseRef.current = true;
      await closeMacosDocumentImportWindow();
    } catch (reason) {
      if (progressTimer !== undefined) window.clearInterval(progressTimer);
      if (importerHidden) {
        await restoreMacosDocumentImportWindow().catch(() => undefined);
      }
      setError(documentImportErrorMessage(reason, t("无法将内容插入 Word。", "Could not insert the content into Word.")));
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
        setError(documentImportErrorMessage(reason, t("无法注册文档导入窗口关闭处理。", "Could not register the import window close handler.")));
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
        <span>{t("正在准备 Word 文档导入器…", "Preparing the Word document importer…")}</span>
      </main>
    );
  }

  if (!request) {
    return (
      <main className="document-import-state is-error" role="alert">
        <AlertCircle />
        <strong>{t("无法打开文档导入器", "Could not open the document importer")}</strong>
        <p>{error || t("Word 文档导入请求不存在或已经失效。", "The Word document import request does not exist or has expired.")}</p>
      </main>
    );
  }

  const formulas = formulaCount(blocks);
  const displayFormulaCount = blocks.reduce(
    (count, block) =>
      count + (block.kind === "formula" && block.displayMode === "block" ? 1 : 0),
    0,
  );
  const allDisplayFormulasNumbered =
    displayFormulaCount > 0 &&
    blocks.every(
      (block) =>
        block.kind !== "formula" ||
        block.displayMode !== "block" ||
        block.numbered,
    );
  const textCharacters = textCharacterCount(blocks);

  return (
    <main className="doc-import-shell macos-doc-import">
      <header className="doc-import-toolbar">
        <div className="doc-import-title-block">
          <FileText size={20} />
          <div>
            <strong>{t("Word 文档批量导入", "Word document import")}</strong>
            <span>{t("左侧编辑源码，右侧实时查看并调整最终 Word 结构", "Edit the source on the left; review and adjust the final Word structure on the right")}</span>
          </div>
        </div>
        <div className="doc-import-options">
          <label>
            <span>{t("源格式", "Source format")}</span>
            <select
              value={sourceKind}
              onChange={(event) =>
                handleSourceKindChange(event.target.value as DocumentImportSourceKind)
              }
              disabled={busy}
            >
              <option value="auto">{t("自动识别", "Auto-detect")}</option>
              <option value="latex">LaTeX</option>
              <option value="markdown">Markdown</option>
            </select>
          </label>
          <label>
            <span>{t("公式格式", "Formula format")}</span>
            <select
              value={outputKind}
              onChange={(event) =>
                setOutputKind(event.target.value as DocumentFormulaOutputKind)
              }
              disabled={busy}
            >
              <option value="omml">{t("Word 原生 OMML", "Native Word (OMML)")}</option>
              <option value="image">{t("SVG 图片公式", "SVG image")}</option>
            </select>
          </label>
          <label
            className="doc-import-global-number-toggle"
            title={t("统一设置当前所有行间公式的编号", "Set numbering for all display formulas at once")}
          >
            <input
              type="checkbox"
              checked={allDisplayFormulasNumbered}
              onChange={(event) =>
                setAllDisplayFormulaNumbering(event.target.checked)
              }
              disabled={busy || displayFormulaCount === 0}
              aria-label={t("所有行间公式添加编号", "Number all display formulas")}
            />
            <span>{t("行间公式全部编号", "Number all display formulas")}</span>
          </label>
          <button
            type="button"
            className="doc-import-secondary doc-import-file-button"
            onClick={() => fileInputRef.current?.click()}
            disabled={busy}
            title={t("导入单个 LaTeX 或 Markdown 文件", "Import a single LaTeX or Markdown file")}
          >
            <Upload size={16} />
            {t("导入 .tex / .md", "Import .tex / .md")}
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".tex,.md,.markdown,text/x-tex,text/markdown"
            aria-label={t("导入 LaTeX 或 Markdown 文件", "Import a LaTeX or Markdown file")}
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
                <strong>{t("LaTeX / Markdown 源码", "LaTeX / Markdown source")}</strong>
                <small>{t("支持正文、标题、列表、定理、引用、代码块和混合公式", "Supports body text, headings, lists, theorems, quotes, code blocks, and mixed formulas")}</small>
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
                    {importedFile.modified ? t(" · 已编辑", " · edited") : ""}
                  </small>
                </span>
              ) : null}
              <span className="doc-import-pane-stat">
                {source.length.toLocaleString(isEn ? "en-US" : "zh-CN")} {t("字符", "characters")}
              </span>
            </div>
          </div>
          <textarea
            ref={sourceRef}
            value={source}
            onChange={(event) => handleSourceChange(event.target.value)}
            placeholder={t(
              String.raw`在这里粘贴 LaTeX 或 Markdown，例如：

正文中的行内公式 $E=mc^2$。

\begin{equation}
\begin{aligned}
a&=b\\
c&=d
\end{aligned}
\end{equation}`,
              String.raw`Paste LaTeX or Markdown here, for example:

An inline formula in body text: $E=mc^2$.

\begin{equation}
\begin{aligned}
a&=b\\
c&=d
\end{aligned}
\end{equation}`,
            )}
            spellCheck={false}
            autoCapitalize="off"
            autoCorrect="off"
            disabled={busy}
            aria-label={t("文档源码", "Document source")}
          />
        </article>

        <article className="doc-import-pane preview-pane">
          <div className="doc-import-pane-header">
            <div className="doc-import-pane-heading">
              <span className="doc-import-pane-icon is-preview" aria-hidden="true">
                <Eye size={16} />
              </span>
              <div>
                <strong>{t("Word 结构预览", "Word structure preview")}</strong>
                <small>{t("公式卡片可单独调整行内/行间、编号和字号", "Adjust inline/display mode, numbering, and font size on each formula card")}</small>
              </div>
            </div>
            <div className="doc-import-preview-counts" aria-label={t("预览统计", "Preview statistics")}>
              <span>{blocks.length} {t("块", "blocks")}</span>
              <span>{textCharacters} {t("字", "characters")}</span>
              <span>{formulas} {t("公式", "formulas")}</span>
            </div>
          </div>
          <div className="doc-import-preview-scroll">
            <div className="doc-import-preview-document">
              {!blocks.length ? (
                <div className="document-import-empty">
                  <FileText size={34} />
                  <strong>{t("等待文档内容", "Waiting for document content")}</strong>
                  <span>{t("在左侧粘贴内容后，这里会实时生成 Word 结构预览。", "Paste content on the left and a live Word structure preview appears here.")}</span>
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
                        <span>{t("公式", "Formula")} {index + 1}</span>
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
                            aria-label={t("公式显示模式", "Formula display mode")}
                          >
                            <option value="inline">{t("行内公式", "Inline formula")}</option>
                            <option value="block">{t("行间公式", "Display formula")}</option>
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
                              <span>{t("编号", "Numbered")}</span>
                            </label>
                          ) : null}
                          <label>
                            <span>{t("字号", "Font size")}</span>
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
                        <FormulaPreviewBoundary message={t("公式暂时无法预览，请检查 LaTeX。", "The formula cannot be previewed right now. Check the LaTeX.")}>
                          <MathPreview latex={block.latex || "\\placeholder{}"} />
                        </FormulaPreviewBoundary>
                      </div>
                      <textarea
                        value={block.latex}
                        onChange={(event) => updateFormula(block.id, { latex: event.target.value })}
                        spellCheck={false}
                        disabled={busy}
                        aria-label={t("编辑公式 LaTeX", "Edit formula LaTeX")}
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
              {t("预览解析正常；点击导入后将切回 Word 并实时显示插入进度。", "Preview looks good. After you click Import, the app switches back to Word and shows insertion progress.")}
            </span>
          )}
        </div>
        <div className="doc-import-actions">
          <button type="button" className="doc-import-secondary" onClick={() => void cancel()} disabled={busy}>
            <X size={16} />{t("取消", "Cancel")}
          </button>
          <button
            type="button"
            className="doc-import-primary"
            onClick={() => void commit()}
            disabled={busy || !blocks.length}
          >
            {busy ? <LoaderCircle size={16} className="is-spinning" /> : <Check size={16} />}
            {busy ? t("正在导入…", "Importing…") : t("导入到 Word", "Import to Word")}
          </button>
        </div>
      </footer>
    </main>
  );
}
