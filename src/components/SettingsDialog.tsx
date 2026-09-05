import { useEffect, useRef, useState, type ChangeEvent } from "react";
import {
  BrainCircuit,
  Download,
  Image,
  Keyboard,
  Languages,
  Palette,
  RefreshCw,
  RotateCcw,
  Shapes,
  SlidersHorizontal,
  Type,
  Upload,
  X,
} from "lucide-react";
import { CustomSymbolDesignerDialog } from "./CustomSymbolDesignerDialog";
import {
  FORMULA_CHINESE_FONT_OPTIONS,
  FORMULA_LETTER_FONT_OPTIONS,
} from "../editor/formulaFontPreferences";
import {
  normalizePngExportBackground,
  pngExportBackgroundPickerValue,
} from "../export/pngBackground";
import {
  createDefaultCustomTheme,
  publishCustomTheme,
  readCustomTheme,
  THEME_DEFINITIONS,
  type CustomThemeState,
  type ThemePaletteColors,
} from "../themeCustomization";
import {
  readCustomSymbolLibrary,
  replaceCustomSymbolLibrary,
} from "../math/customSymbolRegistry";
import { useEditorStore } from "../stores/editorStore";
import type {
  InputBehaviorSettingKey,
  LatexCodeFormat,
  Theme,
} from "../types/formula";

interface Props {
  open: boolean;
  showApplicationUpdates?: boolean;
  onClose: () => void;
  onCheckForUpdates: () => void;
  onOpenFormulaHotkeys: () => void;
}

const DEFAULT_FORMULA_INSET = 34;
const MIN_FORMULA_INSET = 0;
const MAX_FORMULA_INSET = 160;
const DEFAULT_FORMULA_ROW_VERTICAL_INSET = 8;
const MIN_FORMULA_ROW_VERTICAL_INSET = 0;
const MAX_FORMULA_ROW_VERTICAL_INSET = 24;
const DEFAULT_FORMULA_TOOL_BUTTON_SIZE = 52;
const MIN_FORMULA_TOOL_BUTTON_SIZE = 36;
const MAX_FORMULA_TOOL_BUTTON_SIZE = 84;
const DEFAULT_FORMULA_TOOL_BUTTON_PADDING = 2;
const MIN_FORMULA_TOOL_BUTTON_PADDING = 0;
const MAX_FORMULA_TOOL_BUTTON_PADDING = 12;
const DEFAULT_CLASSIC_TILE_WIDTH = 300;
const MIN_CLASSIC_TILE_WIDTH = 220;
const MAX_CLASSIC_TILE_WIDTH = 720;
const DEFAULT_CLASSIC_DOCK_HEIGHT = 240;
const MIN_CLASSIC_DOCK_HEIGHT = 140;
const MAX_CLASSIC_DOCK_HEIGHT = 560;

const paletteKeys: Array<{
  key: keyof ThemePaletteColors;
  zh: string;
  en: string;
}> = [
  { key: "accent", zh: "主强调色", en: "Accent" },
  { key: "background", zh: "全局背景", en: "Background" },
  { key: "surface", zh: "面板/纸张", en: "Surface" },
  { key: "elevated", zh: "抬升层", en: "Elevated" },
  { key: "sunken", zh: "凹陷层", en: "Sunken" },
  { key: "foreground", zh: "主文字", en: "Foreground" },
  { key: "textMuted", zh: "次级文字", en: "Muted text" },
  { key: "border", zh: "边框", en: "Border" },
  { key: "formulaSurface", zh: "公式背景", en: "Formula surface" },
  { key: "formulaPlaceholder", zh: "占位符", en: "Placeholder" },
  { key: "formulaCaret", zh: "公式光标", en: "Formula caret" },
  { key: "toolbarStructure", zh: "结构类工具", en: "Structure tools" },
  { key: "toolbarCalculus", zh: "微积分工具", en: "Calculus tools" },
  { key: "toolbarMatrix", zh: "矩阵工具", en: "Matrix tools" },
  { key: "toolbarGreek", zh: "希腊字母工具", en: "Greek tools" },
];

const inputBehaviorRows: Array<{
  key: InputBehaviorSettingKey;
  zh: string;
  en: string;
}> = [
  {
    key: "autoEscapeShortcuts",
    zh: "自动识别快捷输入",
    en: "Automatic shortcut expansion",
  },
  {
    key: "autoExitSuperscript",
    zh: "自动退出上标",
    en: "Auto-exit superscript",
  },
  {
    key: "autoExitSubscript",
    zh: "自动退出下标",
    en: "Auto-exit subscript",
  },
  {
    key: "autoExitAccent",
    zh: "自动退出重音结构",
    en: "Auto-exit accents",
  },
  {
    key: "autoExitWrapperCommand",
    zh: "自动退出包裹命令",
    en: "Auto-exit wrapper commands",
  },
  {
    key: "showStructuredCommandSuggestions",
    zh: "显示结构命令建议",
    en: "Structured command suggestions",
  },
  {
    key: "showOtherCommandSuggestions",
    zh: "显示其他命令建议",
    en: "Other command suggestions",
  },
];

interface WebEditorConfiguration {
  version: 1;
  exportedAt: string;
  editor: {
    theme: Theme;
    language: "cn" | "en";
    editorLayout: "standard" | "classic";
    zoom: number;
    sourceOpen: boolean;
    latexCodeFormat: LatexCodeFormat;
    autoPairDelimiters: boolean;
    showLineNumbers: boolean;
    highlightActiveLine: boolean;
    formulaInsetLeft: number;
    formulaInsetRight: number;
    formulaToolButtonSize: number;
    formulaToolButtonPadding: number;
    formulaRowVerticalInset: number;
    pngExportBackground: string;
    formulaLetterFont: string;
    formulaChineseFont: string;
    classicTileWidth: number;
    classicDockHeight: number;
    inputBehavior: Record<string, boolean>;
    personalize: boolean;
    suggestionCount: number;
    checkUpdatesOnStartup: boolean;
  };
  customTheme: CustomThemeState;
  customSymbols: ReturnType<typeof readCustomSymbolLibrary>;
}

function ToggleRow({
  title,
  description,
  checked,
  onChange,
}: {
  title: string;
  description?: string;
  checked: boolean;
  onChange: (enabled: boolean) => void;
}) {
  return (
    <label className="switch-row">
      <span>
        <strong>{title}</strong>
        {description ? <small>{description}</small> : null}
      </span>
      <input
        type="checkbox"
        checked={checked}
        onChange={(event) => onChange(event.currentTarget.checked)}
      />
      <span className="switch-control" />
    </label>
  );
}

function RangeRow({
  title,
  valueLabel,
  min,
  max,
  step = 1,
  value,
  onChange,
}: {
  title: string;
  valueLabel: string;
  min: number;
  max: number;
  step?: number;
  value: number;
  onChange: (value: number) => void;
}) {
  return (
    <label className="range-setting">
      <span>
        <strong>{title}</strong>
        <small>{valueLabel}</small>
      </span>
      <input
        type="range"
        min={min}
        max={max}
        step={step}
        value={value}
        onChange={(event) => onChange(Number(event.currentTarget.value))}
      />
    </label>
  );
}

function SelectRow({
  title,
  description,
  value,
  onChange,
  children,
}: {
  title: string;
  description?: string;
  value: string;
  onChange: (value: string) => void;
  children: React.ReactNode;
}) {
  return (
    <label className="range-setting visualtex-settings-select-row">
      <span>
        <strong>{title}</strong>
        {description ? <small>{description}</small> : null}
      </span>
      <select value={value} onChange={(event) => onChange(event.currentTarget.value)}>
        {children}
      </select>
    </label>
  );
}

export function SettingsDialog({
  open,
  showApplicationUpdates = true,
  onClose,
  onCheckForUpdates,
  onOpenFormulaHotkeys,
}: Props) {
  const dialogRef = useRef<HTMLElement>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const configInputRef = useRef<HTMLInputElement>(null);
  const state = useEditorStore();
  const isEn = state.language === "en";
  const [customTheme, setCustomTheme] = useState<CustomThemeState>(() =>
    readCustomTheme(),
  );
  const [customSymbolDesignerOpen, setCustomSymbolDesignerOpen] = useState(false);
  const [backupStatus, setBackupStatus] = useState("");

  useEffect(() => {
    if (!open) return;
    previousFocusRef.current = document.activeElement as HTMLElement | null;
    setCustomTheme(readCustomTheme());
    setBackupStatus("");
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current
        ?.querySelector<HTMLElement>("button, input, select")
        ?.focus();
    });

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        if (customSymbolDesignerOpen) {
          setCustomSymbolDesignerOpen(false);
        } else {
          onClose();
        }
        return;
      }
      if (event.key !== "Tab" || !dialogRef.current) return;

      const focusable = Array.from(
        dialogRef.current.querySelectorAll<HTMLElement>(
          'button:not(:disabled), input:not(:disabled), select:not(:disabled), [tabindex]:not([tabindex="-1"])',
        ),
      );
      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (!first || !last) return;
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    };

    document.addEventListener("keydown", handleKeyDown);
    return () => {
      window.cancelAnimationFrame(frame);
      document.removeEventListener("keydown", handleKeyDown);
      previousFocusRef.current?.focus({ preventScroll: true });
    };
  }, [open, customSymbolDesignerOpen, onClose]);

  if (!open) return null;

  const updateCustomTheme = (
    updater: (current: CustomThemeState) => CustomThemeState,
  ) => {
    setCustomTheme((current) => {
      const next = updater(current);
      publishCustomTheme(next);
      return next;
    });
    if (state.theme !== "custom") state.setTheme("custom");
  };

  const resetCustomTheme = () => {
    const next = createDefaultCustomTheme();
    setCustomTheme(next);
    publishCustomTheme(next);
    state.setTheme("custom");
  };

  const exportConfiguration = () => {
    const current = useEditorStore.getState();
    const payload: WebEditorConfiguration = {
      version: 1,
      exportedAt: new Date().toISOString(),
      editor: {
        theme: current.theme,
        language: current.language,
        editorLayout: current.editorLayout,
        zoom: current.zoom,
        sourceOpen: current.sourceOpen,
        latexCodeFormat: current.latexCodeFormat,
        autoPairDelimiters: current.autoPairDelimiters,
        showLineNumbers: current.showLineNumbers,
        highlightActiveLine: current.highlightActiveLine,
        formulaInsetLeft: current.formulaInsetLeft,
        formulaInsetRight: current.formulaInsetRight,
        formulaToolButtonSize: current.formulaToolButtonSize,
        formulaToolButtonPadding: current.formulaToolButtonPadding,
        formulaRowVerticalInset: current.formulaRowVerticalInset,
        pngExportBackground: current.pngExportBackground,
        formulaLetterFont: current.formulaLetterFont,
        formulaChineseFont: current.formulaChineseFont,
        classicTileWidth: current.classicTileWidth,
        classicDockHeight: current.classicDockHeight,
        inputBehavior: { ...current.inputBehavior },
        personalize: current.personalize,
        suggestionCount: current.suggestionCount,
        checkUpdatesOnStartup: current.checkUpdatesOnStartup,
      },
      customTheme: readCustomTheme(),
      customSymbols: readCustomSymbolLibrary(),
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `visualtex-web-config-${new Date()
      .toISOString()
      .slice(0, 10)}.json`;
    anchor.click();
    URL.revokeObjectURL(url);
    setBackupStatus(isEn ? "Configuration exported." : "配置已导出。" );
  };

  const importConfiguration = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.currentTarget.files?.[0];
    event.currentTarget.value = "";
    if (!file) return;
    try {
      const parsed = JSON.parse(
        await file.text(),
      ) as Partial<WebEditorConfiguration>;
      if (parsed.version !== 1 || !parsed.editor) {
        throw new Error(
          isEn ? "Unsupported configuration file." : "不支持的配置文件。",
        );
      }
      const editor = parsed.editor;
      const current = useEditorStore.getState();
      if (typeof editor.theme === "string") current.setTheme(editor.theme as Theme);
      if (editor.language === "cn" || editor.language === "en") {
        current.setLanguage(editor.language);
      }
      if (
        editor.editorLayout === "standard" ||
        editor.editorLayout === "classic"
      ) {
        current.setEditorLayout(editor.editorLayout);
      }
      if (typeof editor.zoom === "number") current.setZoom(editor.zoom);
      if (typeof editor.sourceOpen === "boolean") current.setSourceOpen(editor.sourceOpen);
      if (typeof editor.latexCodeFormat === "string") {
        current.setLatexCodeFormat(editor.latexCodeFormat as LatexCodeFormat);
      }
      if (typeof editor.autoPairDelimiters === "boolean") {
        current.setAutoPairDelimiters(editor.autoPairDelimiters);
      }
      if (typeof editor.showLineNumbers === "boolean") {
        current.setShowLineNumbers(editor.showLineNumbers);
      }
      if (typeof editor.highlightActiveLine === "boolean") {
        current.setHighlightActiveLine(editor.highlightActiveLine);
      }
      if (typeof editor.formulaInsetLeft === "number") {
        current.setFormulaInsetLeft(editor.formulaInsetLeft);
      }
      if (typeof editor.formulaInsetRight === "number") {
        current.setFormulaInsetRight(editor.formulaInsetRight);
      }
      if (typeof editor.formulaToolButtonSize === "number") {
        current.setFormulaToolButtonSize(editor.formulaToolButtonSize);
      }
      if (typeof editor.formulaToolButtonPadding === "number") {
        current.setFormulaToolButtonPadding(editor.formulaToolButtonPadding);
      }
      if (typeof editor.formulaRowVerticalInset === "number") {
        current.setFormulaRowVerticalInset(editor.formulaRowVerticalInset);
      }
      if (typeof editor.pngExportBackground === "string") {
        current.setPngExportBackground(
          normalizePngExportBackground(editor.pngExportBackground),
        );
      }
      if (typeof editor.formulaLetterFont === "string") {
        current.setFormulaLetterFont(
          editor.formulaLetterFont as typeof current.formulaLetterFont,
        );
      }
      if (typeof editor.formulaChineseFont === "string") {
        current.setFormulaChineseFont(
          editor.formulaChineseFont as typeof current.formulaChineseFont,
        );
      }
      if (typeof editor.classicTileWidth === "number") {
        current.setClassicTileWidth(editor.classicTileWidth);
      }
      if (typeof editor.classicDockHeight === "number") {
        current.setClassicDockHeight(editor.classicDockHeight);
      }
      if (typeof editor.personalize === "boolean") {
        current.setPersonalize(editor.personalize);
      }
      if (typeof editor.suggestionCount === "number") {
        current.setSuggestionCount(editor.suggestionCount);
      }
      if (typeof editor.checkUpdatesOnStartup === "boolean") {
        current.setCheckUpdatesOnStartup(editor.checkUpdatesOnStartup);
      }
      if (editor.inputBehavior && typeof editor.inputBehavior === "object") {
        for (const row of inputBehaviorRows) {
          const value = editor.inputBehavior[row.key];
          if (typeof value === "boolean") {
            current.setInputBehavior(row.key, value);
          }
        }
      }
      if (parsed.customTheme) {
        setCustomTheme(parsed.customTheme);
        publishCustomTheme(parsed.customTheme);
      }
      if (parsed.customSymbols) replaceCustomSymbolLibrary(parsed.customSymbols);
      setBackupStatus(isEn ? "Configuration imported." : "配置已导入。" );
    } catch (error) {
      setBackupStatus(error instanceof Error ? error.message : String(error));
    }
  };

  return (
    <div
      className="modal-backdrop"
      role="presentation"
      onMouseDown={(event) => {
        if (event.currentTarget === event.target) onClose();
      }}
    >
      <section
        ref={dialogRef}
        className="settings-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="settings-title"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="dialog-header">
          <div>
            <span className="eyebrow">PREFERENCES</span>
            <h2 id="settings-title">{isEn ? "Settings" : "设置"}</h2>
          </div>
          <button
            type="button"
            className="icon-button"
            onClick={onClose}
            aria-label={isEn ? "Close settings" : "关闭设置"}
          >
            <X size={18} />
          </button>
        </header>

        <div className="settings-content">
          <div className="settings-section">
            <div className="settings-section-title">
              <BrainCircuit size={18} />
              <div>
                <h3>{isEn ? "Personalized commands" : "个性化命令推荐"}</h3>
                <p>
                  {isEn
                    ? "Rank suggestions using frequency, accepted prefixes and recency."
                    : "根据使用频率、前缀选择和最近使用时间调整候选顺序。"}
                </p>
              </div>
            </div>
            <ToggleRow
              title={
                isEn ? "Enable personalized ranking" : "启用个性化排序"
              }
              description={
                isEn
                  ? "Turn off to restore the default order"
                  : "关闭后恢复系统默认顺序"
              }
              checked={state.personalize}
              onChange={state.setPersonalize}
            />
            <RangeRow
              title={isEn ? "Suggestion count" : "候选项数量"}
              valueLabel={`${state.suggestionCount} ${isEn ? "items" : "项"}`}
              min={3}
              max={10}
              value={state.suggestionCount}
              onChange={state.setSuggestionCount}
            />
            <button
              type="button"
              className="secondary-button danger-subtle"
              onClick={state.resetUsage}
            >
              <RotateCcw size={15} />
              {isEn ? "Reset recommendation history" : "重置推荐记录"}
            </button>
          </div>

          <div className="settings-section">
            <div className="settings-section-title">
              <Keyboard size={18} />
              <div>
                <h3>{isEn ? "Formula hotkeys" : "公式快捷键"}</h3>
                <p>
                  {isEn
                    ? "Manage formula shortcuts. Ctrl+G (Command+G on macOS) arms one-shot Greek input."
                    : "管理公式快捷键。Ctrl+G（macOS 为 Command+G）可开启一次性希腊字母输入。"}
                </p>
              </div>
            </div>
            <button
              type="button"
              className="secondary-button settings-hotkey-button"
              onClick={onOpenFormulaHotkeys}
            >
              <Keyboard size={15} />
              {isEn ? "Manage formula hotkeys" : "管理公式快捷键"}
            </button>
          </div>

          <div className="settings-section">
            <div className="settings-section-title">
              <SlidersHorizontal size={18} />
              <div>
                <h3>{isEn ? "Appearance & editor" : "外观与编辑"}</h3>
                <p>
                  {isEn
                    ? "Keep the original Web layout while configuring the new editor capabilities."
                    : "保持 Web 原有布局，同时配置新增编辑能力。"}
                </p>
              </div>
            </div>

            <div className="editor-layout-setting">
              <span>
                <strong>{isEn ? "Editor layout" : "编辑器布局"}</strong>
                <small>
                  {isEn
                    ? "Switch between the existing sidebar layout and the classic bottom-tools layout."
                    : "在原有侧栏布局与经典底部工具栏布局之间切换。"}
                </small>
              </span>
              <div
                className="theme-segment editor-layout-segment"
                role="group"
                aria-label={isEn ? "Editor layout" : "编辑器布局"}
              >
                <button
                  type="button"
                  className={state.editorLayout === "standard" ? "is-active" : ""}
                  aria-pressed={state.editorLayout === "standard"}
                  data-editor-layout-choice="standard"
                  onClick={() => state.setEditorLayout("standard")}
                >
                  {isEn ? "Standard" : "标准布局"}
                </button>
                <button
                  type="button"
                  className={state.editorLayout === "classic" ? "is-active" : ""}
                  aria-pressed={state.editorLayout === "classic"}
                  data-editor-layout-choice="classic"
                  onClick={() => state.setEditorLayout("classic")}
                >
                  {isEn ? "Classic" : "经典布局"}
                </button>
              </div>
            </div>

            <div className="theme-choice-setting visualtex-settings-block-gap">
              <span>
                <strong>{isEn ? "Colour theme" : "界面配色"}</strong>
                <small>
                  {isEn
                    ? "The original five Web themes remain available together with the new presets."
                    : "保留原 Web 五套主题，并补充新增主题预设。"}
                </small>
              </span>
              <div
                className="theme-segment theme-choice-segment"
                role="group"
                aria-label={isEn ? "Colour theme" : "界面配色"}
              >
                {THEME_DEFINITIONS.map((definition) => (
                  <button
                    key={definition.id}
                    type="button"
                    className={state.theme === definition.id ? "is-active" : ""}
                    aria-pressed={state.theme === definition.id}
                    data-theme-choice={definition.id}
                    onClick={() => state.setTheme(definition.id)}
                  >
                    <span className="theme-choice-swatch" aria-hidden="true">
                      {definition.swatches.map((color) => (
                        <i key={color} style={{ background: color }} />
                      ))}
                    </span>
                    <span>{isEn ? definition.labelEn : definition.labelZh}</span>
                  </button>
                ))}
              </div>
            </div>

            {state.theme === "custom" ? (
              <div className="visualtex-custom-palette-panel">
                <div className="visualtex-settings-inline-heading">
                  <strong>{isEn ? "Custom palette" : "自定义配色"}</strong>
                  <button
                    type="button"
                    className="secondary-button"
                    onClick={resetCustomTheme}
                  >
                    <RotateCcw size={14} />
                    {isEn ? "Reset" : "重置"}
                  </button>
                </div>
                <SelectRow
                  title={isEn ? "Palette mode" : "配色模式"}
                  value={customTheme.mode}
                  onChange={(value) =>
                    updateCustomTheme((current) => ({
                      ...current,
                      mode: value === "dark" ? "dark" : "light",
                    }))
                  }
                >
                  <option value="light">{isEn ? "Light" : "浅色"}</option>
                  <option value="dark">{isEn ? "Dark" : "深色"}</option>
                </SelectRow>
                <div className="visualtex-custom-palette-grid">
                  {paletteKeys.map((item) => (
                    <label key={item.key}>
                      <span>{isEn ? item.en : item.zh}</span>
                      <input
                        type="color"
                        value={customTheme.colors[item.key]}
                        onChange={(event) => {
                          const value = event.currentTarget.value.toUpperCase();
                          updateCustomTheme((current) => ({
                            ...current,
                            colors: {
                              ...current.colors,
                              [item.key]: value,
                            },
                          }));
                        }}
                      />
                    </label>
                  ))}
                </div>
              </div>
            ) : null}

            <ToggleRow
              title={isEn ? "Auto-pair delimiters" : "自动补全成对符号"}
              description={
                isEn
                  ? "Automatically add the matching bracket, brace or vertical bar"
                  : "输入括号、花括号或竖线时自动添加匹配符号"
              }
              checked={state.autoPairDelimiters}
              onChange={state.setAutoPairDelimiters}
            />
            <ToggleRow
              title={isEn ? "Show line numbers" : "显示行号"}
              checked={state.showLineNumbers}
              onChange={state.setShowLineNumbers}
            />
            <ToggleRow
              title={isEn ? "Highlight active line" : "高亮当前行"}
              checked={state.highlightActiveLine}
              onChange={state.setHighlightActiveLine}
            />
            <RangeRow
              title={isEn ? "Formula zoom" : "公式显示缩放"}
              valueLabel={`${Math.round(state.zoom * 100)}%`}
              min={0.5}
              max={1.6}
              step={0.05}
              value={state.zoom}
              onChange={state.setZoom}
            />
            <RangeRow
              title={isEn ? "Left formula inset" : "公式区左侧留白"}
              valueLabel={`${state.formulaInsetLeft}px`}
              min={MIN_FORMULA_INSET}
              max={MAX_FORMULA_INSET}
              value={state.formulaInsetLeft}
              onChange={state.setFormulaInsetLeft}
            />
            <RangeRow
              title={isEn ? "Right formula inset" : "公式区右侧留白"}
              valueLabel={`${state.formulaInsetRight}px`}
              min={MIN_FORMULA_INSET}
              max={MAX_FORMULA_INSET}
              value={state.formulaInsetRight}
              onChange={state.setFormulaInsetRight}
            />
            <RangeRow
              title={isEn ? "Formula row spacing" : "公式行上下留白"}
              valueLabel={`${state.formulaRowVerticalInset}px`}
              min={MIN_FORMULA_ROW_VERTICAL_INSET}
              max={MAX_FORMULA_ROW_VERTICAL_INSET}
              value={state.formulaRowVerticalInset}
              onChange={state.setFormulaRowVerticalInset}
            />
            <button
              type="button"
              className="secondary-button"
              onClick={() => {
                state.setFormulaInsetLeft(DEFAULT_FORMULA_INSET);
                state.setFormulaInsetRight(DEFAULT_FORMULA_INSET);
                state.setFormulaRowVerticalInset(
                  DEFAULT_FORMULA_ROW_VERTICAL_INSET,
                );
              }}
            >
              <RotateCcw size={14} />
              {isEn ? "Reset formula spacing" : "恢复公式区默认间距"}
            </button>

            <RangeRow
              title={isEn ? "Formula tool button size" : "公式工具按钮尺寸"}
              valueLabel={`${state.formulaToolButtonSize}px`}
              min={MIN_FORMULA_TOOL_BUTTON_SIZE}
              max={MAX_FORMULA_TOOL_BUTTON_SIZE}
              value={state.formulaToolButtonSize}
              onChange={state.setFormulaToolButtonSize}
            />
            <RangeRow
              title={isEn ? "Formula tool gap" : "公式工具按钮间距"}
              valueLabel={`${state.formulaToolButtonPadding}px`}
              min={MIN_FORMULA_TOOL_BUTTON_PADDING}
              max={MAX_FORMULA_TOOL_BUTTON_PADDING}
              value={state.formulaToolButtonPadding}
              onChange={state.setFormulaToolButtonPadding}
            />
            <button
              type="button"
              className="secondary-button"
              onClick={() => {
                state.setFormulaToolButtonSize(DEFAULT_FORMULA_TOOL_BUTTON_SIZE);
                state.setFormulaToolButtonPadding(
                  DEFAULT_FORMULA_TOOL_BUTTON_PADDING,
                );
              }}
            >
              <RotateCcw size={14} />
              {isEn ? "Reset toolbar sizing" : "恢复工具栏默认尺寸"}
            </button>

            {state.editorLayout === "classic" ? (
              <>
                <RangeRow
                  title={isEn ? "Classic tile width" : "经典布局磁贴栏宽度"}
                  valueLabel={`${state.classicTileWidth}px`}
                  min={MIN_CLASSIC_TILE_WIDTH}
                  max={MAX_CLASSIC_TILE_WIDTH}
                  value={state.classicTileWidth}
                  onChange={state.setClassicTileWidth}
                />
                <RangeRow
                  title={isEn ? "Classic dock height" : "经典布局底部栏高度"}
                  valueLabel={`${state.classicDockHeight}px`}
                  min={MIN_CLASSIC_DOCK_HEIGHT}
                  max={MAX_CLASSIC_DOCK_HEIGHT}
                  value={state.classicDockHeight}
                  onChange={state.setClassicDockHeight}
                />
                <button
                  type="button"
                  className="secondary-button"
                  onClick={() => {
                    state.setClassicTileWidth(DEFAULT_CLASSIC_TILE_WIDTH);
                    state.setClassicDockHeight(DEFAULT_CLASSIC_DOCK_HEIGHT);
                  }}
                >
                  <RotateCcw size={14} />
                  {isEn ? "Reset classic layout size" : "恢复经典布局默认尺寸"}
                </button>
              </>
            ) : null}
          </div>

          <div className="settings-section">
            <div className="settings-section-title">
              <Type size={18} />
              <div>
                <h3>{isEn ? "Formula output" : "公式输出"}</h3>
                <p>
                  {isEn
                    ? "Choose formula fonts and the PNG background without changing the Web page theme."
                    : "设置公式字体和 PNG 背景，不改变 Web 页面主题。"}
                </p>
              </div>
            </div>
            <SelectRow
              title={isEn ? "Letters and numbers" : "字母与数字字体"}
              value={state.formulaLetterFont}
              onChange={(value) =>
                state.setFormulaLetterFont(
                  value as typeof state.formulaLetterFont,
                )
              }
            >
              {FORMULA_LETTER_FONT_OPTIONS.map((option) => (
                <option key={option.id} value={option.id}>
                  {option.label}
                </option>
              ))}
            </SelectRow>
            <SelectRow
              title={isEn ? "Chinese formula text" : "公式中文字体"}
              value={state.formulaChineseFont}
              onChange={(value) =>
                state.setFormulaChineseFont(
                  value as typeof state.formulaChineseFont,
                )
              }
            >
              {FORMULA_CHINESE_FONT_OPTIONS.map((option) => (
                <option key={option.id} value={option.id}>
                  {isEn ? option.labelEn : option.labelZh}
                </option>
              ))}
            </SelectRow>
            <SelectRow
              title={isEn ? "PNG background" : "PNG 背景"}
              value={
                state.pngExportBackground === "transparent"
                  ? "transparent"
                  : "solid"
              }
              onChange={(value) =>
                state.setPngExportBackground(
                  value === "transparent"
                    ? "transparent"
                    : pngExportBackgroundPickerValue(
                        state.pngExportBackground,
                      ),
                )
              }
            >
              <option value="transparent">
                {isEn ? "Transparent" : "透明"}
              </option>
              <option value="solid">{isEn ? "Solid colour" : "纯色"}</option>
            </SelectRow>
            {state.pngExportBackground !== "transparent" ? (
              <label className="range-setting visualtex-settings-color-row">
                <span>
                  <strong>{isEn ? "PNG background colour" : "PNG 背景颜色"}</strong>
                  <small>{state.pngExportBackground}</small>
                </span>
                <input
                  type="color"
                  value={pngExportBackgroundPickerValue(
                    state.pngExportBackground,
                  )}
                  onChange={(event) =>
                    state.setPngExportBackground(
                      normalizePngExportBackground(event.currentTarget.value),
                    )
                  }
                />
              </label>
            ) : null}
          </div>

          <div className="settings-section">
            <div className="settings-section-title">
              <Keyboard size={18} />
              <div>
                <h3>{isEn ? "Input behavior" : "输入行为"}</h3>
                <p>
                  {isEn
                    ? "The new input rules can be changed independently."
                    : "各项新增输入逻辑可独立开关。"}
                </p>
              </div>
            </div>
            {inputBehaviorRows.map((row) => (
              <ToggleRow
                key={row.key}
                title={isEn ? row.en : row.zh}
                checked={state.inputBehavior[row.key]}
                onChange={(enabled) => state.setInputBehavior(row.key, enabled)}
              />
            ))}
          </div>

          <div className="settings-section">
            <div className="settings-section-title">
              <Shapes size={18} />
              <div>
                <h3>{isEn ? "Custom symbols" : "自定义字符"}</h3>
                <p>
                  {isEn
                    ? "Build reusable vector symbols and register them as LaTeX commands."
                    : "组合矢量字符并注册为可直接输入的 LaTeX 命令。"}
                </p>
              </div>
            </div>
            <button
              type="button"
              className="secondary-button"
              onClick={() => setCustomSymbolDesignerOpen(true)}
            >
              <Shapes size={15} />
              {isEn ? "Open custom symbol designer" : "打开自定义字符设计器"}
            </button>
          </div>

          <div className="settings-section">
            <div className="settings-section-title">
              <Languages size={18} />
              <div>
                <h3>{isEn ? "Interface language" : "界面语言"}</h3>
                <p>
                  {isEn
                    ? "Switch between English and Chinese."
                    : "切换中文或英文界面。"}
                </p>
              </div>
            </div>
            <div className="theme-segment">
              <button
                type="button"
                className={state.language === "cn" ? "is-active" : ""}
                aria-pressed={state.language === "cn"}
                onClick={() => state.setLanguage("cn")}
              >
                中文
              </button>
              <button
                type="button"
                className={state.language === "en" ? "is-active" : ""}
                aria-pressed={state.language === "en"}
                onClick={() => state.setLanguage("en")}
              >
                English
              </button>
            </div>
          </div>

          <div className="settings-section">
            <div className="settings-section-title">
              <Palette size={18} />
              <div>
                <h3>{isEn ? "Configuration backup" : "配置备份"}</h3>
                <p>
                  {isEn
                    ? "Export or restore editor preferences, custom themes and custom symbols."
                    : "导出或恢复编辑器设置、自定义配色和自定义字符。"}
                </p>
              </div>
            </div>
            <div className="visualtex-settings-action-row">
              <button
                type="button"
                className="secondary-button"
                onClick={exportConfiguration}
              >
                <Download size={15} />
                {isEn ? "Export configuration" : "导出配置"}
              </button>
              <button
                type="button"
                className="secondary-button"
                onClick={() => configInputRef.current?.click()}
              >
                <Upload size={15} />
                {isEn ? "Import configuration" : "导入配置"}
              </button>
              <input
                ref={configInputRef}
                type="file"
                accept="application/json,.json"
                hidden
                onChange={importConfiguration}
              />
            </div>
            {backupStatus ? (
              <div className="visualtex-settings-status" role="status">
                {backupStatus}
              </div>
            ) : null}
          </div>

          {showApplicationUpdates ? (
            <div className="settings-section">
              <div className="settings-section-title">
                <RefreshCw size={18} />
                <div>
                  <h3>{isEn ? "Application updates" : "应用更新"}</h3>
                  <p>
                    {isEn
                      ? "Automatically check GitHub Releases and show localized update details when a newer stable version is published."
                      : "自动检查 GitHub Releases；发布新稳定版本时，按当前语言显示更新内容。"}
                  </p>
                </div>
              </div>
              <ToggleRow
                title={
                  isEn
                    ? "Automatic update notifications"
                    : "自动更新提醒"
                }
                description={
                  isEn
                    ? "Manual checks remain available when automatic checks are disabled."
                    : "关闭自动检查后仍可手动检查更新。"
                }
                checked={state.checkUpdatesOnStartup}
                onChange={state.setCheckUpdatesOnStartup}
              />
              <button
                type="button"
                className="secondary-button settings-update-button"
                onClick={onCheckForUpdates}
              >
                <RefreshCw size={15} />
                {isEn ? "Check now" : "立即检查"}
              </button>
            </div>
          ) : null}
        </div>

        <footer className="dialog-footer">
          <span>{isEn ? "Settings saved automatically" : "设置已自动保存"}</span>
          <button type="button" className="primary-button" onClick={onClose}>
            {isEn ? "Done" : "完成"}
          </button>
        </footer>
      </section>

      <CustomSymbolDesignerDialog
        open={customSymbolDesignerOpen}
        language={state.language}
        onClose={() => setCustomSymbolDesignerOpen(false)}
      />
    </div>
  );
}
