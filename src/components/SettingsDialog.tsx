import { useEffect, useMemo, useRef, useState, type ChangeEvent } from "react";
import { createPortal } from "react-dom";
import { Download, Shapes, Upload, X } from "lucide-react";
import { CustomSymbolDesignerDialog } from "./CustomSymbolDesignerDialog";
import {
  FORMULA_CHINESE_FONT_OPTIONS,
  FORMULA_LETTER_FONT_OPTIONS,
} from "../editor/formulaFontPreferences";
import {
  pngExportBackgroundPickerValue,
  normalizePngExportBackground,
} from "../export/pngBackground";
import {
  createDefaultCustomTheme,
  getThemeDefinition,
  publishCustomTheme,
  readCustomTheme,
  THEME_DEFINITIONS,
  type CustomThemeState,
  type ThemePaletteColors,
} from "../themeCustomization";
import { useEditorStore } from "../stores/editorStore";
import {
  readCustomSymbolLibrary,
  replaceCustomSymbolLibrary,
} from "../math/customSymbolRegistry";
import type { InputBehaviorSettingKey, Theme } from "../types/formula";
import {
  readWebKeypadMode,
  subscribeWebKeypadMode,
  writeWebKeypadMode,
} from "../runtime/webKeypadMode";

interface Props {
  open: boolean;
  onClose: () => void;
  [key: string]: unknown;
}

type SettingsTab = "appearance" | "editor" | "input" | "backup";

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
const DEFAULT_CLASSIC_TILE_WIDTH = 320;
const MIN_CLASSIC_TILE_WIDTH = 220;
const MAX_CLASSIC_TILE_WIDTH = 720;
const DEFAULT_CLASSIC_DOCK_HEIGHT = 240;
const MIN_CLASSIC_DOCK_HEIGHT = 140;
const MAX_CLASSIC_DOCK_HEIGHT = 560;

interface WebEditorConfiguration {
  version: 1;
  exportedAt: string;
  editor: {
    theme: Theme;
    language: "cn" | "en";
    editorLayout: "standard" | "classic";
    zoom: number;
    sourceOpen: boolean;
    latexCodeFormat: string;
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
    webKeypadMode: boolean;
  };
  customTheme: CustomThemeState;
  customSymbols: ReturnType<typeof readCustomSymbolLibrary>;
}

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
  { key: "autoEscapeShortcuts", zh: "自动识别快捷输入", en: "Automatic shortcut expansion" },
  { key: "autoExitSuperscript", zh: "自动退出上标", en: "Auto-exit superscript" },
  { key: "autoExitSubscript", zh: "自动退出下标", en: "Auto-exit subscript" },
  { key: "autoExitAccent", zh: "自动退出重音结构", en: "Auto-exit accents" },
  { key: "autoExitWrapperCommand", zh: "自动退出包裹命令", en: "Auto-exit wrapper commands" },
  { key: "showStructuredCommandSuggestions", zh: "显示结构命令建议", en: "Structured command suggestions" },
  { key: "showOtherCommandSuggestions", zh: "显示其他命令建议", en: "Other command suggestions" },
];

function RangeSetting({
  label,
  value,
  min,
  max,
  step = 1,
  suffix = "px",
  onChange,
}: {
  label: string;
  value: number;
  min: number;
  max: number;
  step?: number;
  suffix?: string;
  onChange: (value: number) => void;
}) {
  return (
    <label className="web-settings-range">
      <span>{label}</span>
      <input
        type="range"
        min={min}
        max={max}
        step={step}
        value={value}
        onChange={(event) => onChange(Number(event.currentTarget.value))}
      />
      <input
        type="number"
        min={min}
        max={max}
        step={step}
        value={value}
        onChange={(event) => {
          const next = Number(event.currentTarget.value);
          if (Number.isFinite(next)) onChange(next);
        }}
      />
      <em>{suffix}</em>
    </label>
  );
}

function ToggleSetting({
  label,
  checked,
  onChange,
}: {
  label: string;
  checked: boolean;
  onChange: (value: boolean) => void;
}) {
  return (
    <label className="web-settings-toggle">
      <span>{label}</span>
      <input
        type="checkbox"
        checked={checked}
        onChange={(event) => onChange(event.currentTarget.checked)}
      />
    </label>
  );
}

export function SettingsDialog({ open, onClose }: Props) {
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const [tab, setTab] = useState<SettingsTab>("appearance");
  const [customTheme, setCustomTheme] = useState<CustomThemeState>(() => readCustomTheme());
  const [backupStatus, setBackupStatus] = useState("");
  const [customSymbolDesignerOpen, setCustomSymbolDesignerOpen] = useState(false);
  const [webKeypadMode, setWebKeypadMode] = useState(readWebKeypadMode);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const state = useEditorStore();
  useEffect(
    () => subscribeWebKeypadMode(setWebKeypadMode),
    [],
  );

  const selectedThemeDefinition = useMemo(
    () => getThemeDefinition(state.theme),
    [state.theme],
  );

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
        webKeypadMode: readWebKeypadMode(),
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
    anchor.download = `visualtex-web-config-${new Date().toISOString().slice(0, 10)}.json`;
    anchor.click();
    URL.revokeObjectURL(url);
    setBackupStatus(isEn ? "Configuration exported." : "配置已导出。" );
  };

  const importConfiguration = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.currentTarget.files?.[0];
    event.currentTarget.value = "";
    if (!file) return;
    try {
      const parsed = JSON.parse(await file.text()) as Partial<WebEditorConfiguration>;
      if (parsed.version !== 1 || !parsed.editor) {
        throw new Error(isEn ? "Unsupported configuration file." : "不支持的配置文件。" );
      }
      const editor = parsed.editor;
      const current = useEditorStore.getState();
      if (typeof editor.theme === "string") current.setTheme(editor.theme as Theme);
      if (editor.language === "cn" || editor.language === "en") current.setLanguage(editor.language);
      if (editor.editorLayout === "standard" || editor.editorLayout === "classic") current.setEditorLayout(editor.editorLayout);
      if (typeof editor.zoom === "number") current.setZoom(editor.zoom);
      if (typeof editor.sourceOpen === "boolean") current.setSourceOpen(editor.sourceOpen);
      if (typeof editor.autoPairDelimiters === "boolean") current.setAutoPairDelimiters(editor.autoPairDelimiters);
      if (typeof editor.showLineNumbers === "boolean") current.setShowLineNumbers(editor.showLineNumbers);
      if (typeof editor.highlightActiveLine === "boolean") current.setHighlightActiveLine(editor.highlightActiveLine);
      if (typeof editor.formulaInsetLeft === "number") current.setFormulaInsetLeft(editor.formulaInsetLeft);
      if (typeof editor.formulaInsetRight === "number") current.setFormulaInsetRight(editor.formulaInsetRight);
      if (typeof editor.formulaToolButtonSize === "number") current.setFormulaToolButtonSize(editor.formulaToolButtonSize);
      if (typeof editor.formulaToolButtonPadding === "number") current.setFormulaToolButtonPadding(editor.formulaToolButtonPadding);
      if (typeof editor.formulaRowVerticalInset === "number") current.setFormulaRowVerticalInset(editor.formulaRowVerticalInset);
      if (typeof editor.pngExportBackground === "string") current.setPngExportBackground(normalizePngExportBackground(editor.pngExportBackground));
      if (typeof editor.formulaLetterFont === "string") current.setFormulaLetterFont(editor.formulaLetterFont as typeof current.formulaLetterFont);
      if (typeof editor.formulaChineseFont === "string") current.setFormulaChineseFont(editor.formulaChineseFont as typeof current.formulaChineseFont);
      if (typeof editor.classicTileWidth === "number") current.setClassicTileWidth(editor.classicTileWidth);
      if (typeof editor.classicDockHeight === "number") current.setClassicDockHeight(editor.classicDockHeight);
      if (typeof editor.personalize === "boolean") current.setPersonalize(editor.personalize);
      if (typeof editor.suggestionCount === "number") current.setSuggestionCount(editor.suggestionCount);
      if (typeof editor.checkUpdatesOnStartup === "boolean") current.setCheckUpdatesOnStartup(editor.checkUpdatesOnStartup);
      if (typeof editor.webKeypadMode === "boolean") writeWebKeypadMode(editor.webKeypadMode);
      if (editor.inputBehavior && typeof editor.inputBehavior === "object") {
        for (const row of inputBehaviorRows) {
          const value = editor.inputBehavior[row.key];
          if (typeof value === "boolean") current.setInputBehavior(row.key, value);
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

  const resetCustomTheme = () => {
    const next = createDefaultCustomTheme();
    setCustomTheme(next);
    publishCustomTheme(next);
    state.setTheme("custom");
  };

  return createPortal(
    <div
      className="modal-backdrop web-settings-backdrop"
      onMouseDown={(event) => {
        if (event.currentTarget === event.target) onClose();
      }}
    >
      <section
        className="web-settings-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={isEn ? "Settings" : "设置"}
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="web-settings-header">
          <strong>{isEn ? "Settings" : "设置"}</strong>
          <button type="button" className="icon-button compact" onClick={onClose}>
            <X size={16} />
          </button>
        </header>
        <div className="web-settings-content">
          <nav className="web-settings-nav">
            {([
              ["appearance", isEn ? "Appearance" : "界面"],
              ["editor", isEn ? "Editor" : "编辑器"],
              ["input", isEn ? "Input" : "输入"],
              ["backup", isEn ? "Configuration" : "配置备份"],
            ] as const).map(([id, label]) => (
              <button
                key={id}
                type="button"
                className={tab === id ? "is-active" : ""}
                onClick={() => setTab(id)}
              >
                {label}
              </button>
            ))}
          </nav>
          <main className="web-settings-main">
            {tab === "appearance" ? (
              <>
                <section className="web-settings-section">
                  <h3>{isEn ? "Theme" : "主题"}</h3>
                  <div className="web-theme-grid">
                    {THEME_DEFINITIONS.map((definition) => (
                      <button
                        type="button"
                        key={definition.id}
                        className={state.theme === definition.id ? "is-active" : ""}
                        onClick={() => state.setTheme(definition.id)}
                      >
                        <span className="web-theme-swatches">
                          {definition.swatches.map((color) => (
                            <i key={color} style={{ backgroundColor: color }} />
                          ))}
                        </span>
                        <span>{isEn ? definition.labelEn : definition.labelZh}</span>
                      </button>
                    ))}
                  </div>
                  <p className="web-settings-meta">
                    {isEn ? selectedThemeDefinition.labelEn : selectedThemeDefinition.labelZh}
                  </p>
                </section>
                <section className="web-settings-section">
                  <div className="web-settings-section-title-row">
                    <h3>{isEn ? "Custom palette" : "自定义配色"}</h3>
                    <button type="button" onClick={resetCustomTheme}>
                      {isEn ? "Reset" : "重置"}
                    </button>
                  </div>
                  <label className="web-settings-select-row">
                    <span>{isEn ? "Palette mode" : "配色模式"}</span>
                    <select
                      value={customTheme.mode}
                      onChange={(event) =>
                        updateCustomTheme((current) => ({
                          ...current,
                          mode: event.currentTarget.value === "dark" ? "dark" : "light",
                        }))
                      }
                    >
                      <option value="light">{isEn ? "Light" : "浅色"}</option>
                      <option value="dark">{isEn ? "Dark" : "深色"}</option>
                    </select>
                  </label>
                  <div className="web-palette-grid">
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
                              colors: { ...current.colors, [item.key]: value },
                            }));
                          }}
                        />
                        <code>{customTheme.colors[item.key]}</code>
                      </label>
                    ))}
                  </div>
                </section>
                <section className="web-settings-section">
                  <h3>{isEn ? "Custom symbols" : "自定义字符"}</h3>
                  <p className="web-settings-meta">
                    {isEn
                      ? "Build reusable vector symbols from LaTeX glyphs and geometry, then use them as normal LaTeX commands."
                      : "用 LaTeX 字形和几何图形组合可复用的矢量字符，并注册为普通 LaTeX 命令。"}
                  </p>
                  <button
                    type="button"
                    className="web-settings-open-designer"
                    onClick={() => setCustomSymbolDesignerOpen(true)}
                  >
                    <Shapes size={15} />
                    {isEn ? "Open custom symbol designer" : "打开自定义字符设计器"}
                  </button>
                </section>
                <section className="web-settings-section">
                  <h3>{isEn ? "Formula fonts" : "公式字体"}</h3>
                  <label className="web-settings-select-row">
                    <span>{isEn ? "Letters and numbers" : "字母与数字"}</span>
                    <select
                      value={state.formulaLetterFont}
                      onChange={(event) =>
                        state.setFormulaLetterFont(
                          event.currentTarget.value as typeof state.formulaLetterFont,
                        )
                      }
                    >
                      {FORMULA_LETTER_FONT_OPTIONS.map((option) => (
                        <option key={option.id} value={option.id}>{option.label}</option>
                      ))}
                    </select>
                  </label>
                  <label className="web-settings-select-row">
                    <span>{isEn ? "Chinese text" : "中文字体"}</span>
                    <select
                      value={state.formulaChineseFont}
                      onChange={(event) =>
                        state.setFormulaChineseFont(
                          event.currentTarget.value as typeof state.formulaChineseFont,
                        )
                      }
                    >
                      {FORMULA_CHINESE_FONT_OPTIONS.map((option) => (
                        <option key={option.id} value={option.id}>
                          {isEn ? option.labelEn : option.labelZh}
                        </option>
                      ))}
                    </select>
                  </label>
                </section>
              </>
            ) : null}

            {tab === "editor" ? (
              <>
                <section className="web-settings-section">
                  <h3>{isEn ? "Formula area" : "公式编辑区"}</h3>
                  <ToggleSetting
                    label={isEn ? "Show line numbers" : "显示行号"}
                    checked={state.showLineNumbers}
                    onChange={state.setShowLineNumbers}
                  />
                  <ToggleSetting
                    label={isEn ? "Highlight active line" : "高亮当前行"}
                    checked={state.highlightActiveLine}
                    onChange={state.setHighlightActiveLine}
                  />
                  <RangeSetting
                    label={isEn ? "Left inset" : "左侧留白"}
                    value={state.formulaInsetLeft}
                    min={MIN_FORMULA_INSET}
                    max={MAX_FORMULA_INSET}
                    onChange={state.setFormulaInsetLeft}
                  />
                  <RangeSetting
                    label={isEn ? "Right inset" : "右侧留白"}
                    value={state.formulaInsetRight}
                    min={MIN_FORMULA_INSET}
                    max={MAX_FORMULA_INSET}
                    onChange={state.setFormulaInsetRight}
                  />
                  <RangeSetting
                    label={isEn ? "Row vertical inset" : "公式行上下留白"}
                    value={state.formulaRowVerticalInset}
                    min={MIN_FORMULA_ROW_VERTICAL_INSET}
                    max={MAX_FORMULA_ROW_VERTICAL_INSET}
                    onChange={state.setFormulaRowVerticalInset}
                  />
                  <button
                    type="button"
                    className="web-settings-reset-row"
                    onClick={() => {
                      state.setFormulaInsetLeft(DEFAULT_FORMULA_INSET);
                      state.setFormulaInsetRight(DEFAULT_FORMULA_INSET);
                      state.setFormulaRowVerticalInset(DEFAULT_FORMULA_ROW_VERTICAL_INSET);
                    }}
                  >
                    {isEn ? "Reset formula spacing" : "恢复公式区默认间距"}
                  </button>
                </section>
                <section className="web-settings-section">
                  <h3>{isEn ? "Formula toolbar" : "公式工具栏"}</h3>
                  <RangeSetting
                    label={isEn ? "Button size" : "按钮尺寸"}
                    value={state.formulaToolButtonSize}
                    min={MIN_FORMULA_TOOL_BUTTON_SIZE}
                    max={MAX_FORMULA_TOOL_BUTTON_SIZE}
                    onChange={state.setFormulaToolButtonSize}
                  />
                  <RangeSetting
                    label={isEn ? "Button padding" : "按钮间距"}
                    value={state.formulaToolButtonPadding}
                    min={MIN_FORMULA_TOOL_BUTTON_PADDING}
                    max={MAX_FORMULA_TOOL_BUTTON_PADDING}
                    onChange={state.setFormulaToolButtonPadding}
                  />
                  <button
                    type="button"
                    className="web-settings-reset-row"
                    onClick={() => {
                      state.setFormulaToolButtonSize(DEFAULT_FORMULA_TOOL_BUTTON_SIZE);
                      state.setFormulaToolButtonPadding(DEFAULT_FORMULA_TOOL_BUTTON_PADDING);
                    }}
                  >
                    {isEn ? "Reset toolbar sizing" : "恢复工具栏默认尺寸"}
                  </button>
                </section>
                <section className="web-settings-section">
                  <h3>{isEn ? "Classic layout" : "经典布局"}</h3>
                  <RangeSetting
                    label={isEn ? "Tile column width" : "磁贴栏宽度"}
                    value={state.classicTileWidth}
                    min={MIN_CLASSIC_TILE_WIDTH}
                    max={MAX_CLASSIC_TILE_WIDTH}
                    onChange={state.setClassicTileWidth}
                  />
                  <RangeSetting
                    label={isEn ? "Bottom dock height" : "底部工具栏高度"}
                    value={state.classicDockHeight}
                    min={MIN_CLASSIC_DOCK_HEIGHT}
                    max={MAX_CLASSIC_DOCK_HEIGHT}
                    onChange={state.setClassicDockHeight}
                  />
                  <button
                    type="button"
                    className="web-settings-reset-row"
                    onClick={() => {
                      state.setClassicTileWidth(DEFAULT_CLASSIC_TILE_WIDTH);
                      state.setClassicDockHeight(DEFAULT_CLASSIC_DOCK_HEIGHT);
                    }}
                  >
                    {isEn ? "Reset classic layout" : "恢复经典布局尺寸"}
                  </button>
                </section>
                <section className="web-settings-section">
                  <h3>{isEn ? "PNG export" : "PNG 导出"}</h3>
                  <label className="web-settings-select-row">
                    <span>{isEn ? "Background" : "背景"}</span>
                    <select
                      value={state.pngExportBackground === "transparent" ? "transparent" : "custom"}
                      onChange={(event) => {
                        state.setPngExportBackground(
                          event.currentTarget.value === "transparent"
                            ? "transparent"
                            : pngExportBackgroundPickerValue(state.pngExportBackground),
                        );
                      }}
                    >
                      <option value="transparent">{isEn ? "Transparent" : "透明"}</option>
                      <option value="custom">{isEn ? "Solid color" : "纯色"}</option>
                    </select>
                  </label>
                  {state.pngExportBackground !== "transparent" ? (
                    <label className="web-settings-color-row">
                      <span>{isEn ? "Background color" : "背景颜色"}</span>
                      <input
                        type="color"
                        value={pngExportBackgroundPickerValue(state.pngExportBackground)}
                        onChange={(event) =>
                          state.setPngExportBackground(
                            normalizePngExportBackground(event.currentTarget.value),
                          )
                        }
                      />
                      <code>{state.pngExportBackground}</code>
                    </label>
                  ) : null}
                </section>
              </>
            ) : null}

            {tab === "input" ? (
              <>
                <section className="web-settings-section">
                  <h3>{isEn ? "Input behavior" : "输入行为"}</h3>
                  <ToggleSetting
                    label={isEn ? "Pair delimiters automatically" : "自动补全括号与定界符"}
                    checked={state.autoPairDelimiters}
                    onChange={state.setAutoPairDelimiters}
                  />
                  {inputBehaviorRows.map((row) => (
                    <ToggleSetting
                      key={row.key}
                      label={isEn ? row.en : row.zh}
                      checked={state.inputBehavior[row.key]}
                      onChange={(enabled) => state.setInputBehavior(row.key, enabled)}
                    />
                  ))}
                </section>
                <section className="web-settings-section">
                  <h3>{isEn ? "Suggestions" : "命令建议"}</h3>
                  <ToggleSetting
                    label={isEn ? "Personalized ranking" : "根据使用习惯排序"}
                    checked={state.personalize}
                    onChange={state.setPersonalize}
                  />
                  <RangeSetting
                    label={isEn ? "Suggestion count" : "候选数量"}
                    value={state.suggestionCount}
                    min={3}
                    max={10}
                    suffix=""
                    onChange={state.setSuggestionCount}
                  />
                </section>
                <section className="web-settings-section">
                  <h3>{isEn ? "Keypad mode" : "小键盘模式"}</h3>
                  <ToggleSetting
                    label={
                      isEn
                        ? "Ctrl/Cmd+S copies the current LaTeX source"
                        : "Ctrl/Cmd+S 复制当前 LaTeX 源码"
                    }
                    checked={webKeypadMode}
                    onChange={(enabled) => {
                      setWebKeypadMode(enabled);
                      writeWebKeypadMode(enabled);
                    }}
                  />
                  <p className="web-settings-meta">
                    {isEn
                      ? "The browser version does not minimize the window after copying."
                      : "浏览器版复制后不会执行桌面端的窗口最小化。"}
                  </p>
                </section>
                <section className="web-settings-section">
                  <h3>{isEn ? "Greek one-shot input" : "希腊字母一次性输入"}</h3>
                  <p className="web-settings-meta">
                    {isEn
                      ? "Press Ctrl+G (Command+G on macOS), then a Latin letter. Hold Shift for capitals."
                      : "按 Ctrl+G（macOS 为 Command+G），再按对应拉丁字母；按住 Shift 输入大写希腊字母。"}
                  </p>
                </section>
              </>
            ) : null}

            {tab === "backup" ? (
              <section className="web-settings-section">
                <h3>{isEn ? "Configuration backup" : "配置备份"}</h3>
                <p className="web-settings-meta">
                  {isEn
                    ? "Exports editor preferences, the custom palette and custom symbols. Document contents are not included."
                    : "导出编辑器设置、自定义配色和自定义字符；不会包含当前文档内容。"}
                </p>
                <div className="web-settings-backup-actions">
                  <button type="button" onClick={exportConfiguration}>
                    <Download size={15} />
                    {isEn ? "Export configuration" : "导出配置"}
                  </button>
                  <button type="button" onClick={() => fileInputRef.current?.click()}>
                    <Upload size={15} />
                    {isEn ? "Import configuration" : "导入配置"}
                  </button>
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept="application/json,.json"
                    hidden
                    onChange={importConfiguration}
                  />
                </div>
                {backupStatus ? <div className="web-settings-status">{backupStatus}</div> : null}
              </section>
            ) : null}
          </main>
        </div>
      </section>
      <CustomSymbolDesignerDialog
        open={customSymbolDesignerOpen}
        language={language}
        onClose={() => setCustomSymbolDesignerOpen(false)}
      />
    </div>,
    document.body,
  );
}
