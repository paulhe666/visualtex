import { useState } from "react";
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Braces,
  Code2,
  Copy,
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
import type { FormulaAlignment } from "../types/formula";
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
  const [classicDockOpen, setClassicDockOpen] = useState(true);
  const title = useEditorStore((state) => state.title);
  const lines = useEditorStore((state) => state.lines);
  const activeLineId = useEditorStore((state) => state.activeLineId);
  const language = useEditorStore((state) => state.language);
  const theme = useEditorStore((state) => state.theme);
  const zoom = useEditorStore((state) => state.zoom);
  const setZoom = useEditorStore((state) => state.setZoom);
  const formulaAlignment = useEditorStore((state) => state.formulaAlignment);
  const setFormulaAlignment = useEditorStore(
    (state) => state.setFormulaAlignment,
  );
  const editorLayout = useEditorStore((state) => state.editorLayout);
  const sourceOpen = useEditorStore((state) => state.sourceOpen);
  const setSourceOpen = useEditorStore((state) => state.setSourceOpen);
  const latexCodeFormat = useEditorStore((state) => state.latexCodeFormat);
  const isEn = language === "en";
  const latex = joinFormulaLines(lines);
  const sourceLatex = formatLatex(latex, latexCodeFormat);
  const applyFormulaAlignment = (alignment: FormulaAlignment) => {
    setFormulaAlignment(alignment);
    editorRef.current?.focus();
  };
  const applySource = (source: string, sourceFormat: typeof latexCodeFormat) => {
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
        formulaAlignment,
        selectionByLineId: editorRef.current?.getSelectionMap() ?? {},
      },
      "source-apply",
    );
  };
  const renderSourceEditor = ({
    showCollapseAction = true,
    showCopyAction = true,
    compact = false,
  }: {
    showCollapseAction?: boolean;
    showCopyAction?: boolean;
    compact?: boolean;
  } = {}) => (
    <LatexSourceEditor
      latex={sourceLatex}
      theme={theme}
      format={latexCodeFormat}
      onCollapse={() => setSourceOpen(false)}
      showCollapseAction={showCollapseAction}
      showCopyAction={showCopyAction}
      compact={compact}
      onApply={applySource}
      onCopy={() => void onCopy()}
    />
  );

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

      <main
        className={
          `workspace ${editorLayout === "classic" ? "is-classic-layout" : "is-standard-layout"}` +
          (sidebarOpen ? " has-sidebar" : "")
        }
        data-editor-layout={editorLayout}
      >
        {editorLayout === "standard" && sidebarOpen && (
          <FormulaToolbar
            stabilizeTileLayout
            onInsert={(command) => editorRef.current?.insertCommand(command)}
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
              <div
                className="formula-alignment-controls"
                role="toolbar"
                aria-label={isEn ? "Formula alignment" : "公式对齐方式"}
              >
                {(
                  [
                    ["left", AlignLeft, isEn ? "Align left" : "左对齐"],
                    ["center", AlignCenter, isEn ? "Centre" : "居中"],
                    ["right", AlignRight, isEn ? "Align right" : "右对齐"],
                  ] as const
                ).map(([alignment, Icon, label]) => (
                  <button
                    key={alignment}
                    type="button"
                    className={
                      "icon-button compact formula-alignment-button" +
                      (formulaAlignment === alignment ? " is-active" : "")
                    }
                    aria-label={label}
                    title={label}
                    aria-pressed={formulaAlignment === alignment}
                    data-formula-alignment={alignment}
                    onMouseDown={(event) => event.preventDefault()}
                    onClick={() => applyFormulaAlignment(alignment)}
                  >
                    <Icon size={16} strokeWidth={2} />
                  </button>
                ))}
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
                  <Minus size={14} />
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
                  <Plus size={14} />
                </button>
              </div>
            </div>
          </header>

          {editorLayout === "classic" ? (
            <div
              className={
                "classic-editor-pane-body" +
                (classicDockOpen ? "" : " is-dock-collapsed")
              }
            >
              <div className="editor-pane-scroll">
                <MathEditor
                  ref={editorRef}
                  lines={lines}
                  activeLineId={activeLineId}
                  formulaAlignment={formulaAlignment}
                  zoom={zoom}
                  onPasteImage={showOcrActions ? onPasteImage : undefined}
                  onHistoryBusyChange={onHistoryBusyChange}
                  overlay={ocrOverlay}
                />
              </div>

              <section
                className={
                  "classic-bottom-dock" +
                  (classicDockOpen ? "" : " is-collapsed") +
                  (sourceOpen ? " is-source-panel" : " is-tools-panel")
                }
                aria-label={
                  isEn ? "Formula tools and LaTeX source" : "公式工具与 LaTeX 源码"
                }
              >
                <nav
                  className="classic-bottom-tabs"
                  aria-label={isEn ? "Bottom editor panel" : "底部编辑面板"}
                >
                  <span className="classic-bottom-tab-spacer" aria-hidden="true" />
                  <div
                    className="classic-bottom-tab-group"
                    role="tablist"
                    aria-label={isEn ? "Bottom editor panel" : "底部编辑面板"}
                  >
                    <button
                      type="button"
                      role="tab"
                      className={!sourceOpen ? "is-active" : ""}
                      aria-selected={!sourceOpen}
                      data-classic-bottom-view="tools"
                      onClick={() => {
                        setSourceOpen(false);
                        setClassicDockOpen(true);
                      }}
                    >
                      <Braces size={16} />
                      {isEn ? "Formula tools" : "公式工具"}
                    </button>
                    <button
                      type="button"
                      role="tab"
                      className={sourceOpen ? "is-active" : ""}
                      aria-selected={sourceOpen}
                      data-classic-bottom-view="source"
                      onClick={() => {
                        setSourceOpen(true);
                        setClassicDockOpen(true);
                      }}
                    >
                      <Code2 size={16} />
                      {isEn ? "LaTeX source" : "LaTeX 源码"}
                    </button>
                  </div>
                  <div className="classic-bottom-actions">
                    {sourceOpen && classicDockOpen && (
                      <button
                        type="button"
                        className="icon-button compact classic-bottom-copy"
                        data-classic-bottom-copy
                        onClick={() => void onCopy()}
                        aria-label={isEn ? "Copy LaTeX source" : "复制 LaTeX 源码"}
                        title={isEn ? "Copy LaTeX source" : "复制 LaTeX 源码"}
                      >
                        <Copy size={14} />
                      </button>
                    )}
                    <button
                      type="button"
                      className="icon-button compact classic-bottom-collapse"
                      data-classic-bottom-collapse
                      aria-expanded={classicDockOpen}
                      aria-label={
                        classicDockOpen
                          ? isEn
                            ? "Collapse formula tools and source"
                            : "收起公式工具与源码"
                          : isEn
                            ? "Expand formula tools and source"
                            : "展开公式工具与源码"
                      }
                      title={
                        classicDockOpen
                          ? isEn
                            ? "Collapse bottom panel"
                            : "收起底部面板"
                          : isEn
                            ? "Expand bottom panel"
                            : "展开底部面板"
                      }
                      onClick={() => setClassicDockOpen((open) => !open)}
                    >
                      {classicDockOpen ? (
                        <PanelBottomClose size={14} />
                      ) : (
                        <PanelBottomOpen size={14} />
                      )}
                    </button>
                  </div>
                </nav>
                {classicDockOpen && (
                  <div className="classic-bottom-content">
                    {sourceOpen ? (
                      <div className="source-pane-slot classic-source-pane-slot">
                        {renderSourceEditor({
                          showCollapseAction: false,
                          showCopyAction: false,
                          compact: true,
                        })}
                      </div>
                    ) : (
                      <FormulaToolbar
                        view="tools"
                        layout="horizontal"
                        className="classic-bottom-toolbar"
                        onInsert={(command) =>
                          editorRef.current?.insertCommand(command)
                        }
                      />
                    )}
                  </div>
                )}
              </section>
            </div>
          ) : (
            <div className={`editor-pane-body${sourceOpen ? " has-source" : ""}`}>
              <div className="editor-pane-scroll">
                <MathEditor
                  ref={editorRef}
                  lines={lines}
                  activeLineId={activeLineId}
                  formulaAlignment={formulaAlignment}
                  zoom={zoom}
                  onPasteImage={showOcrActions ? onPasteImage : undefined}
                  onHistoryBusyChange={onHistoryBusyChange}
                  overlay={ocrOverlay}
                />
              </div>

              {sourceOpen ? (
                <div className="source-pane-slot">{renderSourceEditor()}</div>
              ) : (
                <div className="source-toggle-row">
                  <span className="source-toggle-label" aria-hidden="true">
                    <Code2 size={15} />
                  </span>
                  <button
                    type="button"
                    className="source-toggle"
                    onClick={() => setSourceOpen(true)}
                    aria-label={isEn ? "Show LaTeX source" : "展开 LaTeX 源码"}
                    title={isEn ? "Show LaTeX source" : "展开 LaTeX 源码"}
                  >
                    <PanelBottomOpen size={15} />
                  </button>
                </div>
              )}
            </div>
          )}
        </section>

        {editorLayout === "classic" && sidebarOpen && (
          <FormulaToolbar
            view="tiles"
            className="classic-tile-toolbar"
            stabilizeTileLayout
            onInsert={(command) => editorRef.current?.insertCommand(command)}
          />
        )}
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
