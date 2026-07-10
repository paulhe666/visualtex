import { ChangeEvent, useEffect, useRef, useState } from "react";
import {
  Braces,
  Check,
  ChevronDown,
  Code2,
  Copy,
  FilePlus2,
  FolderOpen,
  History,
  Languages,
  Menu,
  Minus,
  Moon,
  PanelBottomClose,
  PanelBottomOpen,
  Plus,
  Redo2,
  Save,
  Settings2,
  Sun,
  Undo2,
} from "lucide-react";
import { MathEditor, type MathEditorHandle } from "./editor/MathEditor";
import { FormulaToolbar } from "./toolbar/FormulaToolbar";
import { LatexSourceEditor } from "./source-editor/LatexSourceEditor";
import { SettingsDialog } from "./components/SettingsDialog";
import { HistoryPanel } from "./components/HistoryPanel";
import { useEditorStore } from "./stores/editorStore";
import {
  copyLatex,
  formatLatex,
  parseLatexSource,
  splitLatexLines,
  type CopyFormat,
} from "./clipboard/LatexCopyService";
import { normalizeChineseLatex } from "./editor/normalizeChineseLatex";
import type { FormulaDocument } from "./types/formula";

function App() {
  const editorRef = useRef<MathEditorHandle>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [copyMenuOpen, setCopyMenuOpen] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [toast, setToast] = useState("");
  const [savedPulse, setSavedPulse] = useState(false);

  const title = useEditorStore((state) => state.title);
  const setTitle = useEditorStore((state) => state.setTitle);
  const latex = useEditorStore((state) => state.latex);
  const setLatex = useEditorStore((state) => state.setLatex);
  const theme = useEditorStore((state) => state.theme);
  const setTheme = useEditorStore((state) => state.setTheme);
  const language = useEditorStore((state) => state.language);
  const setLanguage = useEditorStore((state) => state.setLanguage);
  const zoom = useEditorStore((state) => state.zoom);
  const setZoom = useEditorStore((state) => state.setZoom);
  const sourceOpen = useEditorStore((state) => state.sourceOpen);
  const setSourceOpen = useEditorStore((state) => state.setSourceOpen);
  const addHistory = useEditorStore((state) => state.addHistory);
  const loadDocument = useEditorStore((state) => state.loadDocument);
  const toDocument = useEditorStore((state) => state.toDocument);
  const isEn = language === "en";
  const latexLines = splitLatexLines(latex);
  const sourceLatex = formatLatex(latex, "display");

  const copyLabels: Record<CopyFormat, { title: string; hint: string }> = isEn
    ? {
        display: { title: "Display math (recommended)", hint: "$$ ... $$" },
        plain: { title: "Raw LaTeX", hint: "\\frac{x}{y}" },
        inline: { title: "Inline math", hint: "\\( ... \\)" },
        equation: { title: "equation environment", hint: "\\begin{equation}" },
      }
    : {
        display: { title: "独立公式（推荐）", hint: "$$ ... $$" },
        plain: { title: "纯公式源码", hint: "\\frac{x}{y}" },
        inline: { title: "行内公式", hint: "\\( ... \\)" },
        equation: { title: "equation 环境", hint: "\\begin{equation}" },
      };

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
  }, [theme]);

  useEffect(() => {
    document.documentElement.lang = isEn ? "en" : "zh-CN";
  }, [isEn]);

  useEffect(() => {
    if (!latex.trim()) return;
    const timeout = window.setTimeout(() => addHistory(latex), 1800);
    return () => window.clearTimeout(timeout);
  }, [latex, addHistory]);

  useEffect(() => {
    if (!toast) return;
    const timeout = window.setTimeout(() => setToast(""), 1800);
    return () => window.clearTimeout(timeout);
  }, [toast]);

  const handleCopy = async (format: CopyFormat = "display") => {
    try {
      await copyLatex(latex, format);
      addHistory(latex);
      setToast(
        isEn
          ? "Copied " + copyLabels[format].title
          : "已复制" + copyLabels[format].title,
      );
      setCopyMenuOpen(false);
    } catch {
      setToast(
        isEn
          ? "Copy failed. Check clipboard permission."
          : "复制失败，请检查系统剪贴板权限",
      );
    }
  };

  const saveDocument = () => {
    const document = toDocument();
    const blob = new Blob([JSON.stringify(document, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const link = window.document.createElement("a");
    const safeTitle =
      title.trim().replace(/[\\/:*?"<>|]/g, "-") ||
      (isEn ? "Untitled Formula" : "未命名公式");
    link.href = url;
    link.download = safeTitle + ".visualtex.json";
    link.click();
    URL.revokeObjectURL(url);
    setSavedPulse(true);
    setToast(isEn ? "Formula document saved" : "公式文档已保存");
    window.setTimeout(() => setSavedPulse(false), 900);
  };

  const openDocument = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const parsed = JSON.parse(await file.text()) as FormulaDocument;
      if (!parsed.formulas || !Array.isArray(parsed.formulas)) {
        throw new Error("invalid");
      }
      loadDocument(parsed);
      setToast(isEn ? "Formula document opened" : "公式文档已打开");
    } catch {
      setToast(
        isEn
          ? "Unable to open: invalid file format"
          : "无法打开：文件格式不正确",
      );
    } finally {
      event.target.value = "";
    }
  };

  const newFormula = () => {
    addHistory(latex);
    setTitle(isEn ? "Untitled Formula" : "未命名公式");
    setLatex("");
    editorRef.current?.focus();
    setToast(isEn ? "Created a blank formula" : "已新建空白公式");
  };

  const runMenuAction = (action: () => void) => {
    setMenuOpen(false);
    action();
  };

  return (
    <div className="app-shell">
      <input
        ref={fileInputRef}
        type="file"
        accept=".json,.visualtex"
        className="visually-hidden"
        onChange={openDocument}
      />

      <header className="app-header">
        <div className="brand-area">
          <button
            type="button"
            className={"menu-button " + (menuOpen ? "is-active" : "")}
            aria-label={isEn ? "Main menu" : "主菜单"}
            aria-expanded={menuOpen}
            onClick={() => setMenuOpen((open) => !open)}
          >
            <Menu size={18} />
          </button>
          <div className="brand-mark">
            <Braces size={19} strokeWidth={2.3} />
          </div>
          <div className="brand-copy">
            <strong>VisualTeX</strong>
            <span>Formula Studio</span>
          </div>

          {menuOpen && (
            <div className="app-menu-popover">
              <div className="app-menu-heading">
                <strong>VisualTeX</strong>
                <span>{isEn ? "Formula workspace" : "公式工作区"}</span>
              </div>
              <button type="button" onClick={() => runMenuAction(newFormula)}>
                <FilePlus2 size={16} />
                <span>{isEn ? "New formula" : "新建公式"}</span>
                <kbd>⌘N</kbd>
              </button>
              <button
                type="button"
                onClick={() =>
                  runMenuAction(() => fileInputRef.current?.click())
                }
              >
                <FolderOpen size={16} />
                <span>{isEn ? "Open document" : "打开文档"}</span>
                <kbd>⌘O</kbd>
              </button>
              <button type="button" onClick={() => runMenuAction(saveDocument)}>
                <Save size={16} />
                <span>{isEn ? "Save document" : "保存文档"}</span>
                <kbd>⌘S</kbd>
              </button>
              <div className="app-menu-divider" />
              <button
                type="button"
                onClick={() => runMenuAction(() => setHistoryOpen(true))}
              >
                <History size={16} />
                <span>{isEn ? "Formula history" : "公式历史"}</span>
              </button>
              <button
                type="button"
                onClick={() => runMenuAction(() => setSettingsOpen(true))}
              >
                <Settings2 size={16} />
                <span>{isEn ? "Settings" : "设置"}</span>
              </button>
              <div className="app-menu-divider" />
              <div className="app-menu-language">
                <span>
                  <Languages size={15} />
                  {isEn ? "Language" : "语言"}
                </span>
                <div>
                  <button
                    type="button"
                    className={language === "cn" ? "is-active" : ""}
                    onClick={() => setLanguage("cn")}
                  >
                    CN
                  </button>
                  <button
                    type="button"
                    className={language === "en" ? "is-active" : ""}
                    onClick={() => setLanguage("en")}
                  >
                    ENG
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>

        <div className="document-title-area">
          <input
            value={title}
            onChange={(event) => setTitle(event.target.value)}
            aria-label={isEn ? "Formula document title" : "公式文档标题"}
          />
          <span className={"save-state " + (savedPulse ? "is-saved" : "")}>
            <Check size={12} /> {isEn ? "Auto saved" : "自动保存"}
          </span>
        </div>

        <div className="header-actions">
          <div className="action-group">
            <button type="button" className="icon-button" onClick={newFormula} title={isEn ? "New" : "新建"}>
              <FilePlus2 size={17} />
            </button>
            <button type="button" className="icon-button" onClick={() => fileInputRef.current?.click()} title={isEn ? "Open" : "打开"}>
              <FolderOpen size={17} />
            </button>
            <button type="button" className="icon-button" onClick={saveDocument} title={isEn ? "Save" : "保存到本地"}>
              <Save size={17} />
            </button>
          </div>
          <div className="action-group">
            <button type="button" className="icon-button" onClick={() => editorRef.current?.undo()} title={isEn ? "Undo" : "撤销"}>
              <Undo2 size={17} />
            </button>
            <button type="button" className="icon-button" onClick={() => editorRef.current?.redo()} title={isEn ? "Redo" : "重做"}>
              <Redo2 size={17} />
            </button>
          </div>
          <button type="button" className="icon-button" onClick={() => setHistoryOpen(true)} title={isEn ? "Formula history" : "公式历史"}>
            <History size={17} />
          </button>
          <button
            type="button"
            className="icon-button"
            onClick={() => setTheme(theme === "light" ? "dark" : "light")}
            title={
              isEn
                ? theme === "light"
                  ? "Switch to dark mode"
                  : "Switch to light mode"
                : theme === "light"
                  ? "切换深色模式"
                  : "切换浅色模式"
            }
          >
            {theme === "light" ? <Moon size={17} /> : <Sun size={17} />}
          </button>
          <button type="button" className="icon-button" onClick={() => setSettingsOpen(true)} title={isEn ? "Settings" : "设置"}>
            <Settings2 size={17} />
          </button>
          <div className="copy-control">
            <button type="button" className="copy-primary" onClick={() => handleCopy("display")}>
              <Copy size={16} /> {isEn ? "Copy LaTeX" : "复制 LaTeX"}
            </button>
            <button
              type="button"
              className="copy-chevron"
              aria-label={isEn ? "Choose copy format" : "选择复制格式"}
              onClick={() => setCopyMenuOpen((open) => !open)}
            >
              <ChevronDown size={15} />
            </button>
            {copyMenuOpen && (
              <div className="copy-menu">
                <span className="copy-menu-label">
                  {isEn ? "Copy format" : "复制格式"}
                </span>
                {(Object.keys(copyLabels) as CopyFormat[]).map((format) => (
                  <button type="button" key={format} onClick={() => handleCopy(format)}>
                    <span>
                      <strong>{copyLabels[format].title}</strong>
                      <small>{copyLabels[format].hint}</small>
                    </span>
                    {format === "display" && <Check size={14} />}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
      </header>

      {menuOpen && (
        <button
          type="button"
          className="menu-dismiss-layer"
          aria-label={isEn ? "Close menu" : "关闭菜单"}
          onClick={() => setMenuOpen(false)}
        />
      )}

      <FormulaToolbar
        onInsert={(command) => editorRef.current?.insertCommand(command)}
      />

      <main className="workspace">
        <section className="formula-workspace">
          <div className="workspace-heading">
            <div>
              <span className="eyebrow">FORMULA CANVAS</span>
              <h1>{isEn ? "Visual formula" : "可视化公式"}</h1>
            </div>
            <div className="canvas-controls">
              <button type="button" className="icon-button small" onClick={() => setZoom(zoom - 0.1)} aria-label={isEn ? "Zoom out" : "缩小公式"}>
                <Minus size={15} />
              </button>
              <span>{Math.round(zoom * 100)}%</span>
              <button type="button" className="icon-button small" onClick={() => setZoom(zoom + 0.1)} aria-label={isEn ? "Zoom in" : "放大公式"}>
                <Plus size={15} />
              </button>
            </div>
          </div>

          <MathEditor
            ref={editorRef}
            lines={latexLines}
            zoom={zoom}
            onChange={(nextLines) => setLatex(nextLines.join("\n"))}
          />

          <div className="source-toggle-row">
            <button
              type="button"
              className="source-toggle"
              onClick={() => setSourceOpen(!sourceOpen)}
            >
              <Code2 size={15} />
              {sourceOpen
                ? isEn
                  ? "Hide LaTeX source"
                  : "收起 LaTeX 源码"
                : isEn
                  ? "Show LaTeX source"
                  : "展开 LaTeX 源码"}
              {sourceOpen ? <PanelBottomClose size={15} /> : <PanelBottomOpen size={15} />}
            </button>
            <span>
              {isEn ? "Single source of truth · Two-way sync" : "单一数据源 · 双向同步"}
            </span>
          </div>

          {sourceOpen && (
            <LatexSourceEditor
              latex={sourceLatex}
              theme={theme}
              onApply={(source) =>
                setLatex(
                  parseLatexSource(source)
                    .map(normalizeChineseLatex)
                    .join("\n"),
                )
              }
              onCopy={() => handleCopy("display")}
            />
          )}
        </section>
      </main>

      <footer className="status-bar">
        <div>
          <span className="status-live-dot" />
          {isEn ? "MathLive ready" : "MathLive 就绪"}
        </div>
        <div>
          <span>
            {latexLines.length} {isEn ? "formula lines" : "行公式"}
          </span>
          <span className="status-divider" />
          <span>
            {latex.length} {isEn ? "characters" : "字符"}
          </span>
          <span className="status-divider" />
          <span>{isEn ? "No TeX Live required" : "无需 TeX Live"}</span>
          <span className="status-divider" />
          <span>UTF-8</span>
        </div>
      </footer>

      <SettingsDialog open={settingsOpen} onClose={() => setSettingsOpen(false)} />
      <HistoryPanel
        open={historyOpen}
        onClose={() => setHistoryOpen(false)}
        onRestore={(value) => {
          setLatex(value);
          setHistoryOpen(false);
          setToast(isEn ? "Formula restored" : "已恢复历史公式");
        }}
      />

      {historyOpen && (
        <div className="panel-backdrop" onClick={() => setHistoryOpen(false)} />
      )}
      {toast && (
        <div className="toast">
          <Check size={15} />
          {toast}
        </div>
      )}
    </div>
  );
}

export default App;
