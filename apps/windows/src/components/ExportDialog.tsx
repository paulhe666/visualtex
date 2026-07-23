import { useEffect, useMemo, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { save } from "@tauri-apps/plugin-dialog";
import {
  FileCode2,
  FileImage,
  FileText,
  FolderOpen,
  LoaderCircle,
  X,
} from "lucide-react";
import { buildMarkdownDocument } from "../export/markdownExport";
import { latexToSvg, svgToPng } from "../export/runtime";
import { isTauriEnvironment } from "../ocr/ocrService";

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
  descriptionZh: string;
  descriptionEn: string;
}

const EXPORT_FORMATS: Record<ExportFormat, ExportFormatDefinition> = {
  markdown: {
    extension: "md",
    mime: "text/markdown;charset=utf-8",
    labelZh: "Markdown",
    labelEn: "Markdown",
    descriptionZh: "按行导出为独立的块级 LaTeX 公式",
    descriptionEn: "Export every line as a separate display LaTeX block",
  },
  svg: {
    extension: "svg",
    mime: "image/svg+xml;charset=utf-8",
    labelZh: "SVG",
    labelEn: "SVG",
    descriptionZh: "自包含矢量图，适合排版和继续缩放",
    descriptionEn: "Self-contained vector artwork for publishing and scaling",
  },
  png: {
    extension: "png",
    mime: "image/png",
    labelZh: "PNG",
    labelEn: "PNG",
    descriptionZh: "透明背景高分辨率位图",
    descriptionEn: "High-resolution bitmap with a transparent background",
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
  const [error, setError] = useState("");
  const isEn = language === "en";
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
    if (!isTauriEnvironment()) return path;
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
      if (isTauriEnvironment() && !targetPath) {
        targetPath = (await choosePath()) ?? "";
        if (!targetPath) return;
      }
      targetPath = withExtension(targetPath, definition.extension);
      if (targetPath && targetPath !== path) setPath(targetPath);

      const joinedLatex = nonEmptyFormulas.join("\n");
      let text: string | undefined;
      let base64: string | undefined;
      let browserPayload: string | Blob;

      if (format === "markdown") {
        text = buildMarkdownDocument(title, nonEmptyFormulas);
        browserPayload = text;
      } else {
        const svg = latexToSvg(joinedLatex, {
          displayMode: true,
          fontSizePt: 18,
          paddingPx: 18,
          background: "transparent",
        });
        if (format === "svg") {
          text = svg.svg;
          browserPayload = text;
        } else {
          const png = await svgToPng(svg, {
            scale: 3,
            background: "transparent",
          });
          base64 = png.base64;
          browserPayload = png.blob;
        }
      }

      if (isTauriEnvironment()) {
        if (!targetPath) throw new Error("Export path is empty");
        await invoke("write_export_file", {
          request: {
            path: targetPath,
            text,
            base64,
          },
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
            <span>
              {isEn
                ? "Choose a format and enter an exact destination path."
                : "选择导出格式，并输入准确的保存路径。"}
            </span>
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
                  <small>{isEn ? item.descriptionEn : item.descriptionZh}</small>
                </span>
              </button>
            );
          })}
        </div>

        <label className="export-path-field">
          <span>{isEn ? "Export path" : "导出路径"}</span>
          <div>
            <input
              value={path}
              disabled={busy}
              onChange={(event) => setPath(event.target.value)}
              placeholder={
                isEn
                  ? `For example: C:\\Users\\Name\\Documents\\${suggestedFilename}`
                  : `例如：C:\\Users\\用户名\\Documents\\${suggestedFilename}`
              }
              spellCheck={false}
              autoComplete="off"
            />
            <button
              type="button"
              className="secondary-button export-browse-button"
              onClick={() => void choosePath()}
              disabled={busy || !isTauriEnvironment()}
            >
              <FolderOpen size={15} />
              {isEn ? "Browse" : "浏览"}
            </button>
          </div>
          <small>
            {isEn
              ? `The filename will use .${definition.extension}. Missing folders are created automatically.`
              : `文件将使用 .${definition.extension} 扩展名；不存在的文件夹会自动创建。`}
          </small>
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
