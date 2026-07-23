import { useEffect, useRef, useState } from "react";
import {
  ChevronDown,
  Download,
  FileCode2,
  FileImage,
  FileText,
  FolderOpen,
  LoaderCircle,
} from "lucide-react";
import type { WorkspaceExportFormat } from "../workspace/workspaceTypes";

interface ExportMenuProps {
  isEn: boolean;
  directory?: string;
  busy?: boolean;
  onChooseDirectory: () => Promise<void>;
  onExport: (format: WorkspaceExportFormat) => Promise<void>;
}

function compactPath(path: string, fallback: string) {
  const normalized = path.trim().replace(/\/+$/, "");
  if (!normalized) return fallback;
  const parts = normalized.split("/").filter(Boolean);
  if (parts.length <= 3) return normalized;
  return `…/${parts.slice(-3).join("/")}`;
}

export function ExportMenu({
  isEn,
  directory = "",
  busy = false,
  onChooseDirectory,
  onExport,
}: ExportMenuProps) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const handlePointerDown = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setOpen(false);
    };
    window.addEventListener("pointerdown", handlePointerDown);
    window.addEventListener("keydown", handleKeyDown);
    return () => {
      window.removeEventListener("pointerdown", handlePointerDown);
      window.removeEventListener("keydown", handleKeyDown);
    };
  }, [open]);

  const runExport = async (format: WorkspaceExportFormat) => {
    if (busy) return;
    try {
      await onExport(format);
    } finally {
      setOpen(false);
    }
  };

  const pathLabel = compactPath(
    directory,
    isEn ? "Choose on first export" : "首次导出时选择",
  );

  return (
    <div className="export-menu" ref={rootRef}>
      <button
        type="button"
        className={`export-menu-trigger${open ? " is-active" : ""}`}
        aria-haspopup="menu"
        aria-expanded={open}
        onClick={() => setOpen((value) => !value)}
      >
        {busy ? <LoaderCircle size={16} className="is-spinning" /> : <Download size={16} />}
        <span>{isEn ? "Export" : "导出"}</span>
        <ChevronDown size={14} aria-hidden="true" />
      </button>

      {open && (
        <div
          className="export-menu-popover"
          role="menu"
          aria-label={isEn ? "Export formula" : "导出公式"}
        >
          <div className="export-menu-heading">
            <strong>{isEn ? "Export format" : "导出格式"}</strong>
            <span>
              {isEn
                ? "Export the current formula document"
                : "导出当前公式文档"}
            </span>
          </div>

          <div className="export-format-options">
            <button
              type="button"
              role="menuitem"
              disabled={busy}
              onClick={() => void runExport("markdown")}
            >
              <FileText size={17} />
              <span>
                <strong>Markdown</strong>
                <small>.md</small>
              </span>
            </button>
            <button
              type="button"
              role="menuitem"
              disabled={busy}
              onClick={() => void runExport("svg")}
            >
              <FileCode2 size={17} />
              <span>
                <strong>SVG</strong>
                <small>.svg</small>
              </span>
            </button>
            <button
              type="button"
              role="menuitem"
              disabled={busy}
              onClick={() => void runExport("png")}
            >
              <FileImage size={17} />
              <span>
                <strong>PNG</strong>
                <small>.png</small>
              </span>
            </button>
          </div>

          <div className="export-path-section">
            <div className="export-path-copy">
              <span>{isEn ? "Export location" : "导出路径"}</span>
              <strong title={directory || pathLabel}>{pathLabel}</strong>
            </div>
            <button
              type="button"
              className="export-path-button"
              disabled={busy}
              onClick={() => void onChooseDirectory()}
            >
              <FolderOpen size={15} />
              <span>{isEn ? "Choose" : "选择"}</span>
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
