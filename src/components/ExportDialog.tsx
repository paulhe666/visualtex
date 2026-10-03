import { useEffect, useMemo, useState } from "react";
import { invoke, isTauri } from "@tauri-apps/api/core";
import { save } from "@tauri-apps/plugin-dialog";
import {
  ClipboardCopy,
  FileCode2,
  FileImage,
  FileText,
  FolderOpen,
  LoaderCircle,
  X,
} from "lucide-react";
import { buildMarkdownDocument } from "../export/markdownExport";
import { latexToSvg } from "../export/runtime";
import {
  copyFormulaDocumentPngToClipboard,
  renderFormulaDocumentPng,
} from "../export/pngClipboard";
import { useEditorStore } from "../stores/editorStore";

export type ExportFormat = "markdown" | "svg" | "png";

interface ExportDialogProps {
  open: boolean;
  title: string;
  formulas: readonly string[];
  language: "cn" | "en";
  onClose: () => void;
  onNotify: (message: string) => void;
}

interface ExportFormatDefinition {
  extension: "md" | "svg" | "png";
  mime: string;
  labelZh: string;
  labelEn: string;
}

const EXPORT_FORMATS: Record<ExportFormat, ExportFormatDefinition> = {
  markdown: {
    extension: "md",
    mime: "text/markdown;charset=utf-8",
    labelZh: "Markdown",
    labelEn: "Markdown",
  },
  svg: {
    extension: "svg",
    mime: "image/svg+xml;charset=utf-8",
    labelZh: "SVG",
    labelEn: "SVG",
  },
  png: {
    extension: "png",
    mime: "image/png",
    labelZh: "PNG",
    labelEn: "PNG",
  },
};

function safeFilename(title: string) {
  return title.trim().replace(/[\\/:*?"<>|]/g, "-") || "VisualTeX-Formula";
}

function basename(path: string) {
  const normalized = path.replaceAll("\\", "/");
  return normalized.slice(normalized.lastIndexOf("/") + 1);
}

function withExtension(path: string, extension: string) {
  const trimmed = path.trim();
  if (!trimmed) return trimmed;
  const name = basename(trimmed);
  if (!name) return trimmed;
  const dotIndex = name.lastIndexOf(".");
  if (dotIndex <= 0) return `${trimmed}.${extension}`;
  return `${trimmed.slice(0, trimmed.length - name.length)}${name.slice(0, dotIndex)}.${extension}`;
}

function encodeUtf8Base64(value: string) {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return btoa(binary);
}

function downloadInBrowser(content: string | Blob, filename: string, mime: string) {
  const blob = content instanceof Blob ? content : new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

export function ExportDialog({
  open,
  title,
  formulas,
  language,
  onClose,
  onNotify,
}: ExportDialogProps) {
  const [format, setFormat] = useState<ExportFormat>("markdown");
  const [path, setPath] = useState("");
  const [busy, setBusy] = useState(false);
  const [copyBusy, setCopyBusy] = useState(false);
  const [error, setError] = useState("");
  const isEn = language === "en";
  const pngExportBackground = useEditorStore(
    (state) => state.pngExportBackground,
  );
  const formulaLetterFont = useEditorStore((state) => state.formulaLetterFont);
  const formulaChineseFont = useEditorStore((state) => state.formulaChineseFont);
  const nativeTauri = isTauri();
  const definition = EXPORT_FORMATS[format];
  const suggestedFilename = `${safeFilename(title)}.${definition.extension}`;
  const nonEmptyFormulas = useMemo(
    () => formulas.map((formula) => formula.trim()).filter(Boolean),
    [formulas],
  );

  useEffect(() => {
    if (!open) return;
    setError("");
  }, [open]);

  useEffect(() => {
    setPath((current) =>
      current.trim() ? withExtension(current, definition.extension) : current,
    );
    setError("");
  }, [definition.extension]);

  if (!open) return null;

  const choosePath = async () => {
    setError("");
    if (!nativeTauri) return path;
    try {
      const selected = await save({
        title: isEn ? "Choose export path" : "选择导出路径",
        defaultPath: path.trim() || suggestedFilename,
        filters: [
          {
            name: definition.labelEn,
            extensions: [definition.extension],
          },
        ],
      });
      if (!selected) return null;
      const normalized = withExtension(selected, definition.extension);
      setPath(normalized);
      return normalized;
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
      return null;
    }
  };

  const writeExport = async () => {
    if (!nonEmptyFormulas.length || busy) {
      if (!nonEmptyFormulas.length) {
        setError(isEn ? "There is no formula to export." : "没有可导出的公式。");
      }
      return;
    }

    setBusy(true);
    setError("");
    try {
      let targetPath = path.trim();
      if (nativeTauri && !targetPath) {
        targetPath = (await choosePath()) ?? "";
        if (!targetPath) return;
      }
      targetPath = withExtension(targetPath, definition.extension);
      if (targetPath && targetPath !== path) setPath(targetPath);

      let dataBase64: string;
      let browserPayload: string | Blob;

      if (format === "markdown") {
        const markdown = buildMarkdownDocument(title, nonEmptyFormulas);
        dataBase64 = encodeUtf8Base64(markdown);
        browserPayload = markdown;
      } else if (format === "svg") {
        const svg = latexToSvg(nonEmptyFormulas.join("\n"), {
          displayMode: true,
          fontSizePt: 18,
          paddingPx: 18,
          background: "transparent",
          formulaLetterFont,
          formulaChineseFont,
        });
        dataBase64 = svg.base64;
        browserPayload = svg.svg;
      } else {
        const png = await renderFormulaDocumentPng(nonEmptyFormulas, {
          background: pngExportBackground,
          formulaLetterFont,
          formulaChineseFont,
        });
        dataBase64 = png.base64;
        browserPayload = png.blob;
      }

      if (nativeTauri) {
        if (!targetPath) throw new Error("Export path is empty");
        await invoke("write_export_file", {
          path: targetPath,
          dataBase64,
        });
      } else {
        downloadInBrowser(
          browserPayload,
          suggestedFilename,
          definition.mime,
        );
      }

      onNotify(
        isEn
          ? `${definition.labelEn} exported successfully`
          : `${definition.labelZh} 已成功导出`,
      );
      onClose();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    } finally {
      setBusy(false);
    }
  };

  const copyPng = async () => {
    if (!nonEmptyFormulas.length || copyBusy) {
      if (!nonEmptyFormulas.length) {
        setError(isEn ? "There is no formula to export." : "没有可导出的公式。");
      }
      return;
    }
    setCopyBusy(true);
    setError("");
    try {
      await copyFormulaDocumentPngToClipboard(nonEmptyFormulas, {
        background: pngExportBackground,
        formulaLetterFont,
        formulaChineseFont,
      });
      onNotify(isEn ? "PNG copied to Clipboard" : "PNG 已复制到剪贴板");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    } finally {
      setCopyBusy(false);
    }
  };

  return (
    <div className="modal-backdrop export-dialog-backdrop" role="presentation">
      <section
        className="dialog export-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="export-dialog-title"
      >
        <header>
          <div>
            <strong id="export-dialog-title">{isEn ? "Export formula" : "导出公式"}</strong>
          </div>
          <button
            type="button"
            className="icon-button compact"
            onClick={onClose}
            disabled={busy}
            aria-label={isEn ? "Close export dialog" : "关闭导出窗口"}
          >
            <X size={17} />
          </button>
        </header>

        <div className="export-format-grid" role="radiogroup" aria-label={isEn ? "Export format" : "导出格式"}>
          {(
            [
              ["markdown", FileText],
              ["svg", FileCode2],
              ["png", FileImage],
            ] as const
          ).map(([id, Icon]) => {
            const item = EXPORT_FORMATS[id];
            return (
              <button
                type="button"
                role="radio"
                aria-checked={format === id}
                className={`export-format-option${format === id ? " is-active" : ""}`}
                key={id}
                onClick={() => setFormat(id)}
                disabled={busy}
              >
                <Icon size={22} />
                <span>
                  <strong>{isEn ? item.labelEn : item.labelZh}</strong>
                </span>
              </button>
            );
          })}
        </div>

        <div className="export-copy-png-row">
          <button
            type="button"
            className="export-copy-png-button"
            onClick={() => void copyPng()}
            disabled={busy || copyBusy}
            data-copy-png-from-export
          >
            {copyBusy ? (
              <LoaderCircle size={15} className="spin" />
            ) : (
              <ClipboardCopy size={15} />
            )}
            <span>
              {copyBusy
                ? isEn
                  ? "Copying PNG…"
                  : "正在复制 PNG…"
                : isEn
                  ? "Copy PNG to Clipboard"
                  : "复制 PNG 到剪贴板"}
            </span>
          </button>
        </div>

        <label className="export-path-field">
          <span>{isEn ? "Export path" : "导出路径"}</span>
          <div>
            <input
              value={path}
              disabled={busy}
              onChange={(event) => setPath(event.target.value)}
              placeholder={`/Users/name/Documents/${suggestedFilename}`}
              spellCheck={false}
              autoComplete="off"
            />
            <button
              type="button"
              className="secondary-button export-browse-button"
              onClick={() => void choosePath()}
              disabled={busy || !nativeTauri}
            >
              <FolderOpen size={15} />
              {isEn ? "Browse" : "浏览"}
            </button>
          </div>
        </label>

        {error && <div className="export-error" role="alert">{error}</div>}

        <footer>
          <button
            type="button"
            className="secondary-button"
            onClick={onClose}
            disabled={busy}
          >
            {isEn ? "Cancel" : "取消"}
          </button>
          <button
            type="button"
            className="primary-button export-confirm-button"
            onClick={() => void writeExport()}
            disabled={busy || !nonEmptyFormulas.length}
          >
            {busy && <LoaderCircle className="spin" size={15} />}
            {busy
              ? isEn
                ? "Exporting…"
                : "正在导出…"
              : isEn
                ? `Export ${definition.labelEn}`
                : `导出 ${definition.labelZh}`}
          </button>
        </footer>
      </section>
    </div>
  );
}
