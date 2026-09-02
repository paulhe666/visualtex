import { useEffect, useMemo, useRef, useState } from "react";
import { createUuid } from "../../runtime/browserCompatibility";
import {
  Braces,
  CheckCircle2,
  Eye,
  FileText,
  FolderOpen,
  LoaderCircle,
  TriangleAlert,
  X,
} from "lucide-react";
import { readErrorMessage } from "../../errors/readErrorMessage";
import { latexToSvg } from "../../export/latexToSvg";
import {
  applyDocumentTheme,
  normalizeSynchronizedTheme,
  readSynchronizedTheme,
  subscribeSynchronizedTheme,
} from "../../themeSync";
import {
  closeOfficeSessionWindow,
  getOfficeSession,
  getOfficeTheme,
  saveOfficeSessionKeepalive,
  updateOfficeSession,
  type OfficeFormulaSession,
} from "../api/sessionClient";
import {
  readDocumentImportFile,
  type ImportedDocumentFile,
} from "./documentImportFile";
import {
  parseDocumentImport,
  type DocumentImportBlock,
  type DocumentImportRun,
  type DocumentObjectMode,
  type DocumentSourceFormat,
  type ParsedDocumentImport,
} from "./documentImportParser";
import "./documentImport.css";

function sessionIdFromLocation() {
  const match = window.location.pathname.match(/\/dialog\/([0-9a-f-]{36})/i);
  if (!match) throw new Error("批量导入窗口缺少有效的 Session id。");
  return match[1].toLowerCase();
}

function formatFromSession(session: OfficeFormulaSession): DocumentSourceFormat {
  if (session.codeFormat === "markdown-document") return "markdown";
  if (session.codeFormat === "latex-document") return "latex";
  return "auto";
}

function FormulaPreview({ latex, display }: { latex: string; display: boolean }) {
  const rendered = useMemo(() => {
    try {
      return {
        svg: latexToSvg(latex, {
          displayMode: display,
          fontSizePt: 13,
          paddingPx: display ? 4 : 1,
          background: "transparent",
        }).svg,
        error: "",
      };
    } catch (error) {
      return {
        svg: "",
        error: readErrorMessage(error, "公式预览失败。"),
      };
    }
  }, [display, latex]);

  if (rendered.error) {
    return (
      <span className={display ? "doc-import-formula-error display" : "doc-import-formula-error"}>
        <code>{latex}</code>
        <small>{rendered.error}</small>
      </span>
    );
  }
  return (
    <span
      className={display ? "doc-import-formula display" : "doc-import-formula"}
      dangerouslySetInnerHTML={{ __html: rendered.svg }}
    />
  );
}

function InlineRuns({ runs }: { runs: DocumentImportRun[] }) {
  return (
    <>
      {runs.map((run, index) => {
        if (run.kind === "formula") {
          return <FormulaPreview key={index} latex={run.latex} display={run.display} />;
        }
        const className = [
          run.bold ? "bold" : "",
          run.italic ? "italic" : "",
          run.code ? "code" : "",
          run.strike ? "strike" : "",
          run.underline ? "underline" : "",
        ]
          .filter(Boolean)
          .join(" ");
        return (
          <span key={index} className={className || undefined}>
            {run.text}
          </span>
        );
      })}
    </>
  );
}

function PreviewBlock({ block }: { block: DocumentImportBlock }) {
  if (block.kind === "display") {
    const formula = block.runs.find((run) => run.kind === "formula");
    return formula?.kind === "formula" ? (
      <div className="doc-import-display-row">
        <FormulaPreview latex={formula.latex} display />
      </div>
    ) : null;
  }
  if (block.kind === "code") {
    const text = block.runs.map((run) => (run.kind === "text" ? run.text : run.latex)).join("");
    return <pre className="doc-import-code-block">{text}</pre>;
  }
  if (block.kind === "heading") {
    const content = <InlineRuns runs={block.runs} />;
    switch (Math.min(6, Math.max(1, block.level))) {
      case 1:
        return <h1 className="doc-import-heading">{content}</h1>;
      case 2:
        return <h2 className="doc-import-heading">{content}</h2>;
      case 3:
        return <h3 className="doc-import-heading">{content}</h3>;
      case 4:
        return <h4 className="doc-import-heading">{content}</h4>;
      case 5:
        return <h5 className="doc-import-heading">{content}</h5>;
      default:
        return <h6 className="doc-import-heading">{content}</h6>;
    }
  }
  if (block.kind === "quote") {
    return (
      <blockquote className="doc-import-quote">
        <InlineRuns runs={block.runs} />
      </blockquote>
    );
  }
  if (block.kind === "bullet" || block.kind === "numbered") {
    return (
      <div
        className={`doc-import-list-row ${block.kind}`}
        style={{ paddingInlineStart: `${18 + block.level * 24}px` }}
      >
        <span className="doc-import-list-marker">
          {block.kind === "bullet" ? "•" : `${block.level + 1}.`}
        </span>
        <span><InlineRuns runs={block.runs} /></span>
      </div>
    );
  }
  return (
    <p className="doc-import-paragraph">
      <InlineRuns runs={block.runs} />
    </p>
  );
}

function PreviewPane({ parsed }: { parsed: ParsedDocumentImport }) {
  return (
    <div className="doc-import-preview-stage">
      <div className="doc-import-preview-caption" aria-hidden="true">
        <span>
          <Eye size={14} />
          Word 页面预览
        </span>
        <span>A4 · 实时结构</span>
      </div>
      <div className="doc-import-preview-document" role="document">
        <div className="doc-import-paper-content">
          {parsed.blocks.length > 0 ? (
            parsed.blocks.map((block) => (
              <PreviewBlock key={block.id} block={block} />
            ))
          ) : (
            <div className="doc-import-paper-empty">
              <FileText size={28} />
              <strong>等待文档内容</strong>
              <span>在左侧输入或粘贴 LaTeX、Markdown 后，这里会实时生成 Word 结构预览。</span>
            </div>
          )}
        </div>
        <div className="doc-import-page-footer" aria-hidden="true">
          <span>VisualTeX 文档预览</span>
          <span>1</span>
        </div>
      </div>
    </div>
  );
}

type ImportedFileState = Pick<ImportedDocumentFile, "name" | "encoding" | "size"> & {
  modified: boolean;
};

function formatFileSize(size: number) {
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

export function DocumentImportApp() {
  const sessionId = useMemo(sessionIdFromLocation, []);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const finalizedRef = useRef(false);
  const [session, setSession] = useState<OfficeFormulaSession | null>(null);
  const [source, setSource] = useState("");
  const [format, setFormat] = useState<DocumentSourceFormat>("auto");
  const [objectMode, setObjectMode] = useState<DocumentObjectMode>("wordOmml");
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [fileBusy, setFileBusy] = useState(false);
  const [importedFile, setImportedFile] = useState<ImportedFileState | null>(null);
  const [loadError, setLoadError] = useState("");

  useEffect(() => {
    let disposed = false;
    const applyTheme = (value: unknown) => applyDocumentTheme(normalizeSynchronizedTheme(value));
    applyTheme(readSynchronizedTheme());
    const unsubscribe = subscribeSynchronizedTheme(applyTheme);
    let syncInFlight = false;
    const sync = async () => {
      if (disposed || syncInFlight) return;
      syncInFlight = true;
      try {
        const status = await getOfficeTheme();
        if (!disposed) applyTheme(status.theme);
      } catch {
        // Keep the last synchronized theme while the companion restarts.
      } finally {
        syncInFlight = false;
      }
    };
    void sync();
    const interval = window.setInterval(() => void sync(), 500);
    return () => {
      disposed = true;
      unsubscribe();
      window.clearInterval(interval);
    };
  }, []);

  useEffect(() => {
    let disposed = false;
    void getOfficeSession(sessionId)
      .then((next) => {
        if (disposed) return;
        setSession(next);
        setSource(next.lines[0]?.latex ?? "");
        setFormat(formatFromSession(next));
        setObjectMode(
          next.objectMode === "nativeOle"
            ? "nativeOle"
            : next.objectMode === "mathTypeOle"
              ? "mathTypeOle"
              : "wordOmml",
        );
        setLoading(false);
      })
      .catch((error) => {
        if (disposed) return;
        setLoadError(readErrorMessage(error, "无法读取批量导入 Session。"));
        setLoading(false);
      });
    return () => {
      disposed = true;
    };
  }, [sessionId]);

  const preview = useMemo(() => {
    try {
      return { parsed: parseDocumentImport(source, format), error: "" };
    } catch (error) {
      return { parsed: null, error: readErrorMessage(error, "无法解析当前文档。") };
    }
  }, [format, source]);

  useEffect(() => {
    const cancelOnClose = () => {
      if (finalizedRef.current || !session) return;
      void saveOfficeSessionKeepalive(sessionId, {
        status: "cancelled",
        explicitCancel: true,
        error: null,
      }).catch(() => undefined);
    };
    window.addEventListener("beforeunload", cancelOnClose);
    return () => window.removeEventListener("beforeunload", cancelOnClose);
  }, [session, sessionId]);

  const cancel = async () => {
    if (busy) return;
    finalizedRef.current = true;
    try {
      await updateOfficeSession(sessionId, {
        status: "cancelled",
        explicitCancel: true,
        error: null,
      });
    } finally {
      await closeOfficeSessionWindow(sessionId).catch(() => undefined);
    }
  };

  const commit = async () => {
    if (!session || !preview.parsed || preview.error || busy) return;
    setBusy(true);
    try {
      const lineId = session.lines[0]?.id || createUuid();
      const serializedDocument = JSON.stringify(preview.parsed);
      if (serializedDocument.length > 5_000_000) {
        throw new Error("解析后的文档结构超过 5 MB，无法提交给 Word。请拆分后导入。");
      }
      await updateOfficeSession(sessionId, {
        title: "Word 文档批量导入",
        lines: [{ id: lineId, latex: serializedDocument }],
        activeLineId: lineId,
        codeFormat: "visualtex-document-json",
        objectMode,
        displayMode: "block",
        numbered: false,
        dirty: true,
        status: "committing",
        explicitCancel: false,
        error: null,
      });
      finalizedRef.current = true;
      await closeOfficeSessionWindow(sessionId).catch(() => undefined);
    } catch (error) {
      setLoadError(readErrorMessage(error, "无法把文档提交给 Word。"));
      setBusy(false);
    }
  };

  const openFile = async (file: File) => {
    if (fileBusy || busy) return;
    setFileBusy(true);
    setLoadError("");
    try {
      const imported = await readDocumentImportFile(file);
      setSource(imported.source);
      setFormat(imported.format);
      setImportedFile({
        name: imported.name,
        encoding: imported.encoding,
        size: imported.size,
        modified: false,
      });
    } catch (error) {
      setLoadError(readErrorMessage(error, "无法读取所选文件。"));
    } finally {
      setFileBusy(false);
    }
  };

  if (loading) {
    return (
      <main className="doc-import-loading">
        <LoaderCircle className="spin" />
        <span>正在打开 Word 文档导入器…</span>
      </main>
    );
  }
  if (!session) {
    return (
      <main className="doc-import-loading error">
        <TriangleAlert />
        <span>{loadError || "无法读取批量导入 Session。"}</span>
      </main>
    );
  }

  return (
    <main className="doc-import-shell">
      <header className="doc-import-toolbar">
        <div className="doc-import-title-block">
          <FileText size={20} />
          <div>
            <strong>Word 文档批量导入</strong>
            <span>左侧编辑源码，右侧实时查看 Word 导入结构</span>
          </div>
        </div>
        <div className="doc-import-options">
          <label>
            <span>源格式</span>
            <select value={format} onChange={(event) => setFormat(event.target.value as DocumentSourceFormat)}>
              <option value="auto">自动识别</option>
              <option value="latex">LaTeX</option>
              <option value="markdown">Markdown</option>
            </select>
          </label>
          <label>
            <span>公式格式</span>
            <select value={objectMode} onChange={(event) => setObjectMode(event.target.value as DocumentObjectMode)}>
              <option value="wordOmml">Word 原生 OMML</option>
              <option value="nativeOle">VisualTeX OLE</option>
              <option value="mathTypeOle">MathType OLE</option>
            </select>
          </label>
          <button
            className="doc-import-secondary doc-import-file-button"
            onClick={() => fileInputRef.current?.click()}
            disabled={busy || fileBusy}
            title="导入单个 LaTeX 或 Markdown 文件"
          >
            {fileBusy ? <LoaderCircle size={16} className="spin" /> : <FolderOpen size={16} />}
            {fileBusy ? "正在读取…" : "导入 .tex / .md"}
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".tex,.md,.markdown,text/x-tex,text/markdown"
            aria-label="导入 LaTeX 或 Markdown 文件"
            hidden
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) void openFile(file);
              event.currentTarget.value = "";
            }}
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
                <small>支持正文、标题、列表、引用、代码块和混合公式</small>
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
                  <small>{importedFile.encoding}{importedFile.modified ? " · 已编辑" : ""}</small>
                </span>
              ) : null}
              <span className="doc-import-pane-stat">
                {source.length.toLocaleString()} 字符
              </span>
            </div>
          </div>
          <textarea
            value={source}
            placeholder={String.raw`在这里粘贴 LaTeX 或 Markdown，例如：

正文中的行内公式 $E=mc^2$。

正文后直接接行间公式： \[\frac{1}{2\pi\tau}\]

\begin{itemize}
\item 第一项
\item 第二项
\end{itemize}`}
            onChange={(event) => {
              setSource(event.target.value);
              setImportedFile((current) =>
                current && !current.modified ? { ...current, modified: true } : current,
              );
            }}
            spellCheck={false}
            autoCapitalize="off"
            autoCorrect="off"
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
                <small>按最终导入层级模拟正文、公式与段落间距</small>
              </div>
            </div>
            {preview.parsed ? (
              <div className="doc-import-preview-counts" aria-label="预览统计">
                <span>{preview.parsed.blocks.length} 块</span>
                <span>{preview.parsed.inlineFormulaCount} 行内</span>
                <span>{preview.parsed.displayFormulaCount} 行间</span>
              </div>
            ) : (
              <span className="doc-import-pane-stat">等待有效内容</span>
            )}
          </div>
          <div className="doc-import-preview-scroll">
            {preview.parsed ? <PreviewPane parsed={preview.parsed} /> : (
              <div className="doc-import-preview-error">
                <TriangleAlert size={20} />
                <span>{preview.error}</span>
              </div>
            )}
          </div>
        </article>
      </section>

      <footer className="doc-import-footer">
        <div className="doc-import-messages">
          {loadError ? <span className="error"><TriangleAlert size={15} />{loadError}</span> : null}
          {preview.parsed?.warnings.map((warning, index) => (
            <span className="warning" key={index}><TriangleAlert size={15} />{warning}</span>
          ))}
          {!loadError && preview.parsed && preview.parsed.warnings.length === 0 ? (
            <span className="ok">
              <CheckCircle2 size={15} />
              预览解析正常；Word 将按当前结构化预览结果插入。
            </span>
          ) : null}
        </div>
        <div className="doc-import-actions">
          <button className="doc-import-secondary" onClick={() => void cancel()} disabled={busy}>
            <X size={16} />取消
          </button>
          <button className="doc-import-primary" onClick={() => void commit()} disabled={busy || !preview.parsed || Boolean(preview.error)}>
            {busy ? <LoaderCircle size={16} className="spin" /> : <FileText size={16} />}
            {busy ? "正在提交…" : "导入到 Word"}
          </button>
        </div>
      </footer>
    </main>
  );
}
