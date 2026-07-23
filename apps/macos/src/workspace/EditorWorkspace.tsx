import { useState } from "react";
import {
  Braces,
  Code2,
  Minus,
  PanelBottomClose,
  PanelBottomOpen,
  Plus,
  ScanLine,
} from "lucide-react";
import { MathEditor } from "../editor/MathEditor";
import { InputBehaviorMenu } from "../components/InputBehaviorMenu";
import { ExportMenu } from "../components/ExportMenu";
import { FormulaToolbar } from "../toolbar/FormulaToolbar";
import { LatexSourceEditor } from "../source-editor/LatexSourceEditor";
import {
  MAX_EDITOR_ZOOM,
  MIN_EDITOR_ZOOM,
  joinFormulaLines,
  useEditorStore,
} from "../stores/editorStore";
import {
  formatLatex,
  parseLatexSource,
} from "../clipboard/LatexCopyService";
import { normalizeChineseLatex } from "../editor/normalizeChineseLatex";
import { reconcileFormulaLines } from "../history/documentHistory";
import type { EditorWorkspaceProps } from "./workspaceTypes";

export function EditorWorkspace({
  mode,
  showFileActions,
  showOfficeActions,
  showOcrActions,
  primaryActionLabel,
  onPrimaryAction,
  onCancel,
  onExport,
  onChooseExportDirectory,
  exportDirectory,
  exportBusy = false,
  editorRef,
  sidebarOpen,
  onSidebarOpenChange,
  onHistoryBusyChange,
  onPasteImage,
  onCopy,
  onReplaceDocument,
  ocrModel,
  ocrModels = [],
  ocrBusy = false,
  onOcrModelChange,
  ocrOverlay,
}: EditorWorkspaceProps) {
  const [primaryBusy, setPrimaryBusy] = useState(false);
  const title = useEditorStore((state) => state.title);
  const lines = useEditorStore((state) => state.lines);
  const activeLineId = useEditorStore((state) => state.activeLineId);
  const language = useEditorStore((state) => state.language);
  const theme = useEditorStore((state) => state.theme);
  const zoom = useEditorStore((state) => state.zoom);
  const setZoom = useEditorStore((state) => state.setZoom);
  const sourceOpen = useEditorStore((state) => state.sourceOpen);
  const setSourceOpen = useEditorStore((state) => state.setSourceOpen);
  const latexCodeFormat = useEditorStore((state) => state.latexCodeFormat);
  const isEn = language === "en";
  const latex = joinFormulaLines(lines);
  const sourceLatex = formatLatex(latex, latexCodeFormat);

  const runPrimaryAction = async () => {
    if (!onPrimaryAction || primaryBusy) return;
    setPrimaryBusy(true);
    try {
      await onPrimaryAction();
    } finally {
      setPrimaryBusy(false);
    }
  };

  return (
    <>
      {showOfficeActions && (
        <div className="office-workspace-actions" data-workspace-mode={mode}>
          <div>
            <strong>
              {mode === "office-edit"
                ? isEn
                  ? "Edit selected formula"
                  : "编辑所选公式"
                : isEn
                  ? "Create Office formula"
                  : "新建 Office 公式"}
            </strong>
            <span>
              {isEn
                ? "The document is updated only after you finish or close this window."
                : "点击完成或关闭本窗口后，公式才会写入 Office 文档。"}
            </span>
          </div>
          <div>
            {onCancel && (
              <button
                type="button"
                className="secondary-button"
                onClick={() => void onCancel()}
                disabled={primaryBusy}
              >
                {isEn ? "Cancel" : "取消"}
              </button>
            )}
            {onPrimaryAction && (
              <button
                type="button"
                className="primary-button"
                onClick={() => void runPrimaryAction()}
                disabled={primaryBusy}
              >
                {primaryBusy
                  ? isEn
                    ? "Applying…"
                    : "正在应用…"
                  : primaryActionLabel ??
                    (mode === "office-edit"
                      ? isEn
                        ? "Update formula"
                        : "更新公式"
                      : isEn
                        ? "Finish and insert"
                        : "完成并插入")}
              </button>
            )}
          </div>
        </div>
      )}

      <main className={`workspace${sidebarOpen ? " has-sidebar" : ""}`}>
        {sidebarOpen && (
          <FormulaToolbar
            onInsert={(command) => editorRef.current?.insertCommand(command)}
            onClose={() => onSidebarOpenChange(false)}
          />
        )}

        <section className="formula-workspace editor-pane">
          <header className="workspace-heading pane-header editor-pane-header">
            <div className="pane-title-group">
              <span className="pane-icon" aria-hidden="true">
                <Braces size={16} />
              </span>
              <div className="pane-title-copy">
                <h1>{isEn ? "Visual editor" : "可视化编辑"}</h1>
              </div>
            </div>
            <div className="canvas-tool-group">
              {showFileActions && onExport && onChooseExportDirectory && (
                <ExportMenu
                  isEn={isEn}
                  directory={exportDirectory}
                  busy={exportBusy}
                  onChooseDirectory={onChooseExportDirectory}
                  onExport={onExport}
                />
              )}
              <InputBehaviorMenu />
              {showOcrActions && ocrModels.length > 0 && ocrModel && (
                <label
                  className="canvas-ocr-model"
                  title={
                    isEn
                      ? "Model used when an image is pasted into a formula field"
                      : "在公式输入框中粘贴图片时使用的 OCR 模型"
                  }
                >
                  <ScanLine size={14} />
                  <select
                    value={ocrModel}
                    disabled={ocrBusy}
                    onChange={(event) =>
                      onOcrModelChange?.(event.target.value)
                    }
                    aria-label={isEn ? "OCR recognition model" : "OCR 识别模型"}
                  >
                    {ocrModels.map((item) => (
                      <option key={item.id} value={item.id}>
                        {isEn ? item.labelEn : item.labelZh}
                      </option>
                    ))}
                  </select>
                </label>
              )}
              <div className="canvas-controls">
                <button
                  type="button"
                  className="icon-button compact"
                  onClick={() => setZoom(zoom - 0.1)}
                  disabled={zoom <= MIN_EDITOR_ZOOM + 0.0001}
                  aria-label={isEn ? "Zoom out" : "缩小公式"}
                  title={
                    zoom <= MIN_EDITOR_ZOOM + 0.0001
                      ? isEn
                        ? "Minimum zoom: 20%"
                        : "最小缩放：20%"
                      : undefined
                  }
                >
                  <Minus size={15} />
                </button>
                <span aria-live="polite" aria-atomic="true">
                  {Math.round(zoom * 100)}%
                </span>
                <button
                  type="button"
                  className="icon-button compact"
                  onClick={() => setZoom(zoom + 0.1)}
                  disabled={zoom >= MAX_EDITOR_ZOOM - 0.0001}
                  aria-label={isEn ? "Zoom in" : "放大公式"}
                  title={
                    zoom >= MAX_EDITOR_ZOOM - 0.0001
                      ? isEn
                        ? "Maximum zoom: 160%"
                        : "最大缩放：160%"
                      : undefined
                  }
                >
                  <Plus size={15} />
                </button>
              </div>
            </div>
          </header>

          <div className="editor-pane-scroll">
            <MathEditor
              ref={editorRef}
              lines={lines}
              activeLineId={activeLineId}
              zoom={zoom}
              onPasteImage={showOcrActions ? onPasteImage : undefined}
              onHistoryBusyChange={onHistoryBusyChange}
              overlay={ocrOverlay}
            />

            <div className="source-toggle-row">
              <button
                type="button"
                className="source-toggle"
                onClick={() => setSourceOpen(!sourceOpen)}
                aria-label={
                  sourceOpen
                    ? isEn
                      ? "Hide LaTeX source"
                      : "收起 LaTeX 源码"
                    : isEn
                      ? "Show LaTeX source"
                      : "展开 LaTeX 源码"
                }
                title={
                  sourceOpen
                    ? isEn
                      ? "Hide LaTeX source"
                      : "收起 LaTeX 源码"
                    : isEn
                      ? "Show LaTeX source"
                      : "展开 LaTeX 源码"
                }
              >
                <Code2 size={15} />
                {sourceOpen ? (
                  <PanelBottomClose size={15} />
                ) : (
                  <PanelBottomOpen size={15} />
                )}
              </button>
            </div>

            {sourceOpen && (
              <LatexSourceEditor
                latex={sourceLatex}
                theme={theme}
                format={latexCodeFormat}
                onApply={(source, sourceFormat) => {
                  const values = parseLatexSource(source, sourceFormat).map(
                    normalizeChineseLatex,
                  );
                  const nextLines = reconcileFormulaLines(values, lines);
                  const nextActiveLineId = nextLines.some(
                    (line) => line.id === activeLineId,
                  )
                    ? activeLineId
                    : nextLines[0]?.id ?? null;
                  onReplaceDocument(
                    {
                      title,
                      lines: nextLines,
                      activeLineId: nextActiveLineId,
                      selectionByLineId:
                        editorRef.current?.getSelectionMap() ?? {},
                    },
                    "source-apply",
                  );
                }}
                onCopy={() => void onCopy()}
              />
            )}
          </div>
        </section>
      </main>

      <footer className="status-bar">
        <div>
          <span className="status-live-dot" />
          {isEn ? "Ready" : "就绪"}
        </div>
        <div>
          <span>
            {lines.length} {isEn ? "lines" : "行"}
          </span>
          <span>
            · {latex.length} {isEn ? "characters" : "字符"}
          </span>
        </div>
      </footer>
    </>
  );
}
