import { ChangeEvent, useEffect, useRef, useState } from "react";
import {
  Braces,
  Check,
  ChevronDown,
  CircleHelp,
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
  PanelLeftClose,
  PanelLeftOpen,
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
import { OnboardingTour } from "./components/OnboardingTour";
import { VisualTeXLogo } from "./components/VisualTeXLogo";
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
const ONBOARDING_STORAGE_KEY = "visualtex.onboarding.web.v1.completed";

function App() {
  const editorRef = useRef<MathEditorHandle>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const menuButtonRef = useRef<HTMLButtonElement>(null);
  const appMenuRef = useRef<HTMLDivElement>(null);
  const copyMenuButtonRef = useRef<HTMLButtonElement>(null);
  const copyMenuRef = useRef<HTMLDivElement>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(() => window.innerWidth >= 1040);
  const [onboardingOpen, setOnboardingOpen] = useState(
    () => window.localStorage.getItem(ONBOARDING_STORAGE_KEY) !== "true",
  );
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
  const shortcutModifier = /Mac|iPhone|iPad/.test(navigator.platform) ? "⌘" : "Ctrl+";
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
    const compactWindow = window.matchMedia("(max-width: 1040px)");
    const handleCompactWindow = (event: MediaQueryListEvent) => {
      if (event.matches) setSidebarOpen(false);
    };
    compactWindow.addEventListener("change", handleCompactWindow);
    return () => compactWindow.removeEventListener("change", handleCompactWindow);
  }, []);

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

  useEffect(() => {
    const menu = menuOpen ? appMenuRef.current : copyMenuOpen ? copyMenuRef.current : null;
    const trigger = menuOpen ? menuButtonRef.current : copyMenuButtonRef.current;
    if (!menu || !trigger) return;

    const items = Array.from(
      menu.querySelectorAll<HTMLButtonElement>('button:not(:disabled)'),
    );
    const frame = window.requestAnimationFrame(() => items[0]?.focus());

    const handleMenuKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        setMenuOpen(false);
        setCopyMenuOpen(false);
        trigger.focus({ preventScroll: true });
        return;
      }

      if (event.key !== "ArrowDown" && event.key !== "ArrowUp") return;
      event.preventDefault();
      const currentIndex = items.indexOf(document.activeElement as HTMLButtonElement);
      const direction = event.key === "ArrowDown" ? 1 : -1;
      const nextIndex = currentIndex < 0
        ? 0
        : (currentIndex + direction + items.length) % items.length;
      items[nextIndex]?.focus();
    };

    menu.addEventListener("keydown", handleMenuKeyDown);
    return () => {
      window.cancelAnimationFrame(frame);
      menu.removeEventListener("keydown", handleMenuKeyDown);
    };
  }, [menuOpen, copyMenuOpen]);

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

  const finishOnboarding = () => {
    window.localStorage.setItem(ONBOARDING_STORAGE_KEY, "true");
    setOnboardingOpen(false);
    window.requestAnimationFrame(() => editorRef.current?.focus());
  };

  useEffect(() => {
    const handleWindowKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setMenuOpen(false);
        setCopyMenuOpen(false);
        return;
      }

      if (settingsOpen || historyOpen || onboardingOpen) return;
      if ((!event.metaKey && !event.ctrlKey) || event.altKey) return;
      const key = event.key.toLowerCase();

      if (key === "n") {
        event.preventDefault();
        newFormula();
      } else if (key === "o") {
        event.preventDefault();
        fileInputRef.current?.click();
      } else if (key === "s") {
        event.preventDefault();
        saveDocument();
      } else if (key === ",") {
        event.preventDefault();
        setSettingsOpen(true);
      } else if (key === "0") {
        event.preventDefault();
        setZoom(1);
      } else if (key === "=" || key === "+") {
        event.preventDefault();
        setZoom(zoom + 0.1);
      } else if (key === "-") {
        event.preventDefault();
        setZoom(zoom - 0.1);
      }
    };

    window.addEventListener("keydown", handleWindowKeyDown);
    return () => window.removeEventListener("keydown", handleWindowKeyDown);
  }, [latex, title, isEn, zoom, settingsOpen, historyOpen, onboardingOpen]);

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
            ref={menuButtonRef}
            type="button"
            className={"menu-button " + (menuOpen ? "is-active" : "")}
            aria-label={isEn ? "Main menu" : "主菜单"}
            aria-expanded={menuOpen}
            aria-haspopup="menu"
            aria-controls="app-main-menu"
            onClick={() => {
              setCopyMenuOpen(false);
              setMenuOpen((open) => !open);
            }}
          >
            <Menu size={18} />
          </button>
          <button
            type="button"
            className={"icon-button sidebar-toggle " + (sidebarOpen ? "is-active" : "")}
            aria-label={sidebarOpen ? (isEn ? "Hide formula tools" : "隐藏公式工具") : (isEn ? "Show formula tools" : "显示公式工具")}
            aria-pressed={sidebarOpen}
            onClick={() => setSidebarOpen((open) => !open)}
          >
            {sidebarOpen ? <PanelLeftClose size={17} /> : <PanelLeftOpen size={17} />}
          </button>
          <div className="brand-mark" aria-hidden="true">
            <VisualTeXLogo className="visualtex-brand-logo" />
          </div>
          <div className="brand-copy">
            <strong>VisualTeX</strong>
          </div>

          {menuOpen && (
            <div
              ref={appMenuRef}
              id="app-main-menu"
              className="app-menu-popover"
              role="menu"
              aria-label={isEn ? "VisualTeX menu" : "VisualTeX 菜单"}
            >
              <div className="app-menu-heading">
                <strong>VisualTeX</strong>
                <span>{isEn ? "Web formula workspace" : "网页版公式工作区"}</span>
              </div>
              <button type="button" role="menuitem" onClick={() => runMenuAction(newFormula)}>
                <FilePlus2 size={16} />
                <span>{isEn ? "New formula" : "新建公式"}</span>
                <kbd>{shortcutModifier}N</kbd>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() =>
                  runMenuAction(() => fileInputRef.current?.click())
                }
              >
                <FolderOpen size={16} />
                <span>{isEn ? "Open document" : "打开文档"}</span>
                <kbd>{shortcutModifier}O</kbd>
              </button>
              <button type="button" role="menuitem" onClick={() => runMenuAction(saveDocument)}>
                <Save size={16} />
                <span>{isEn ? "Save document" : "保存文档"}</span>
                <kbd>{shortcutModifier}S</kbd>
              </button>
              <div className="app-menu-divider" />
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => setHistoryOpen(true))}
              >
                <History size={16} />
                <span>{isEn ? "Formula history" : "公式历史"}</span>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => setSettingsOpen(true))}
              >
                <Settings2 size={16} />
                <span>{isEn ? "Settings" : "设置"}</span>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => setOnboardingOpen(true))}
              >
                <CircleHelp size={16} />
                <span>{isEn ? "Quick tour" : "新手教程"}</span>
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
                    role="menuitemradio"
                    aria-checked={language === "cn"}
                    className={language === "cn" ? "is-active" : ""}
                    onClick={() => setLanguage("cn")}
                  >
                    CN
                  </button>
                  <button
                    type="button"
                    role="menuitemradio"
                    aria-checked={language === "en"}
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
          <span
            className={"save-state " + (savedPulse ? "is-saved" : "")}
            aria-label={isEn ? "Saved" : "已保存"}
            title={isEn ? "Saved" : "已保存"}
          >
            <Check size={13} />
          </span>
        </div>

        <div className="header-actions">
          <div className="action-group file-actions">
            <button type="button" className="icon-button" onClick={newFormula} aria-label={isEn ? "New" : "新建"} title={`${isEn ? "New" : "新建"} · ${shortcutModifier}N`}>
              <FilePlus2 size={17} />
            </button>
            <button type="button" className="icon-button" onClick={() => fileInputRef.current?.click()} aria-label={isEn ? "Open" : "打开"} title={`${isEn ? "Open" : "打开"} · ${shortcutModifier}O`}>
              <FolderOpen size={17} />
            </button>
            <button type="button" className="icon-button" onClick={saveDocument} aria-label={isEn ? "Save" : "保存到本地"} title={`${isEn ? "Save" : "保存到本地"} · ${shortcutModifier}S`}>
              <Save size={17} />
            </button>
          </div>
          <div className="action-group edit-actions">
            <button type="button" className="icon-button" onClick={() => editorRef.current?.undo()} aria-label={isEn ? "Undo" : "撤销"} title={isEn ? "Undo · ⌘Z" : "撤销 · ⌘Z"}>
              <Undo2 size={17} />
            </button>
            <button type="button" className="icon-button" onClick={() => editorRef.current?.redo()} aria-label={isEn ? "Redo" : "重做"} title={isEn ? "Redo · ⇧⌘Z" : "重做 · ⇧⌘Z"}>
              <Redo2 size={17} />
            </button>
          </div>
          <button type="button" className="icon-button workspace-action" onClick={() => setHistoryOpen(true)} aria-label={isEn ? "Formula history" : "公式历史"} title={isEn ? "Formula history" : "公式历史"}>
            <History size={17} />
          </button>
          <button
            type="button"
            className="icon-button theme-toggle"
            onClick={() => setTheme(theme === "light" ? "dark" : "light")}
            aria-label={theme === "light" ? (isEn ? "Switch to dark mode" : "切换深色模式") : (isEn ? "Switch to light mode" : "切换浅色模式")}
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
          <button type="button" className="icon-button settings-toggle" onClick={() => setSettingsOpen(true)} aria-label={isEn ? "Settings" : "设置"} title={isEn ? "Settings · ⌘," : "设置 · ⌘,"}>
            <Settings2 size={17} />
          </button>
          <div className="copy-control">
            <button type="button" className="copy-primary" onClick={() => handleCopy("display")}>
              <Copy size={16} /> {isEn ? "Copy LaTeX" : "复制 LaTeX"}
            </button>
            <button
              ref={copyMenuButtonRef}
              type="button"
              className="copy-chevron"
              aria-label={isEn ? "Choose copy format" : "选择复制格式"}
              aria-expanded={copyMenuOpen}
              aria-haspopup="menu"
              aria-controls="copy-format-menu"
              onClick={() => {
                setMenuOpen(false);
                setCopyMenuOpen((open) => !open);
              }}
            >
              <ChevronDown size={15} />
            </button>
            {copyMenuOpen && (
              <div
                ref={copyMenuRef}
                id="copy-format-menu"
                className="copy-menu"
                role="menu"
                aria-label={isEn ? "Copy format" : "复制格式"}
              >
                <span className="copy-menu-label">
                  {isEn ? "Copy format" : "复制格式"}
                </span>
                {(Object.keys(copyLabels) as CopyFormat[]).map((format) => (
                  <button type="button" role="menuitem" key={format} onClick={() => handleCopy(format)}>
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

      {(menuOpen || copyMenuOpen) && (
        <button
          type="button"
          className="menu-dismiss-layer"
          aria-label={isEn ? "Close menu" : "关闭菜单"}
          onClick={() => {
            setMenuOpen(false);
            setCopyMenuOpen(false);
          }}
        />
      )}

      <main
        className={`workspace${sidebarOpen ? " has-sidebar" : ""}`}
      >
        {sidebarOpen && (
          <FormulaToolbar
            onInsert={(command) => editorRef.current?.insertCommand(command)}
            onClose={() => setSidebarOpen(false)}
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
              <div className="canvas-controls">
                <button
                  type="button"
                  className="icon-button compact"
                  onClick={() => setZoom(zoom - 0.1)}
                  disabled={zoom <= 0.5001}
                  aria-label={isEn ? "Zoom out" : "缩小公式"}
                  title={zoom <= 0.5001 ? (isEn ? "Minimum zoom: 50%" : "最小缩放：50%") : undefined}
                >
                  <Minus size={15} />
                </button>
                <span aria-live="polite" aria-atomic="true">{Math.round(zoom * 100)}%</span>
                <button
                  type="button"
                  className="icon-button compact"
                  onClick={() => setZoom(zoom + 0.1)}
                  disabled={zoom >= 1.5999}
                  aria-label={isEn ? "Zoom in" : "放大公式"}
                  title={zoom >= 1.5999 ? (isEn ? "Maximum zoom: 160%" : "最大缩放：160%") : undefined}
                >
                  <Plus size={15} />
                </button>
              </div>
            </div>
          </header>

          <div className="editor-pane-scroll">
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
              aria-label={sourceOpen ? (isEn ? "Hide LaTeX source" : "收起 LaTeX 源码") : (isEn ? "Show LaTeX source" : "展开 LaTeX 源码")}
              title={sourceOpen ? (isEn ? "Hide LaTeX source" : "收起 LaTeX 源码") : (isEn ? "Show LaTeX source" : "展开 LaTeX 源码")}
            >
              <Code2 size={15} />
              {sourceOpen ? <PanelBottomClose size={15} /> : <PanelBottomOpen size={15} />}
            </button>
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
          </div>
        </section>

      </main>

      <footer className="status-bar">
        <div>
          <span className="status-live-dot" />
          {isEn ? "Web · saved in this browser" : "网页版 · 数据保存在当前浏览器"}
        </div>
        <div>
          <span>
            {latexLines.length} {isEn ? "lines" : "行"}
          </span>
          <span>
            · {latex.length} {isEn ? "characters" : "字符"}
          </span>
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
      <OnboardingTour
        open={onboardingOpen}
        language={language}
        onFinish={finishOnboarding}
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
