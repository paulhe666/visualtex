import {
  useEffect,
  useRef,
  useState,
  type ChangeEvent,
  type CSSProperties,
} from "react";
import { invoke, isTauri } from "@tauri-apps/api/core";
import { save } from "@tauri-apps/plugin-dialog";
import {
  BrainCircuit,
  ChevronRight,
  Download,
  Eye,
  Keyboard,
  Languages,
  RotateCcw,
  SlidersHorizontal,
  Upload,
  X,
} from "lucide-react";
import {
  DEFAULT_FORMULA_INSET,
  DEFAULT_FORMULA_ROW_VERTICAL_INSET,
  DEFAULT_FORMULA_TOOL_BUTTON_PADDING,
  DEFAULT_FORMULA_TOOL_BUTTON_SIZE,
  DEFAULT_SOURCE_EDITOR_FONT_SIZE,
  EDITOR_ZOOM_STEP,
  MAX_FORMULA_INSET,
  MAX_FORMULA_ROW_VERTICAL_INSET,
  MAX_FORMULA_TOOL_BUTTON_PADDING,
  MAX_FORMULA_TOOL_BUTTON_SIZE,
  MAX_SOURCE_EDITOR_FONT_SIZE,
  MIN_FORMULA_INSET,
  MIN_FORMULA_ROW_VERTICAL_INSET,
  MIN_FORMULA_TOOL_BUTTON_PADDING,
  MIN_FORMULA_TOOL_BUTTON_SIZE,
  MIN_SOURCE_EDITOR_FONT_SIZE,
  useEditorStore,
} from "../stores/editorStore";
import { MathPreview } from "./MathPreview";
import { pngExportBackgroundPickerValue } from "../export/pngBackground";
import {
  DEFAULT_FORMULA_CHINESE_FONT,
  DEFAULT_FORMULA_LETTER_FONT,
  FORMULA_CHINESE_FONT_OPTIONS,
  FORMULA_LETTER_FONT_OPTIONS,
  formulaChineseFontFamily,
  formulaLetterFontFamilies,
  type FormulaChineseFont,
  type FormulaLetterFont,
} from "../editor/formulaFontPreferences";
import {
  applyVisualTexConfiguration,
  buildVisualTexConfiguration,
  parseVisualTexConfiguration,
  VISUALTEX_CONFIGURATION_EXTENSION,
} from "../runtime/applicationConfiguration";
import {
  createDefaultCustomTheme,
  publishCustomTheme,
  readCustomTheme,
  THEME_DEFINITIONS,
  type CustomThemeState,
  type ThemePaletteColors,
} from "../themeCustomization";

const THEME_CORE_COLOR_FIELDS: readonly [keyof ThemePaletteColors, string, string][] = [
  ["accent", "Accent", "强调色"],
  ["accentHover", "Accent hover", "强调色悬停"],
  ["accentSoft", "Accent soft", "强调浅色"],
  ["background", "App background", "应用背景"],
  ["elevated", "Raised background", "抬升背景"],
  ["surface", "Panel surface", "面板背景"],
  ["sunken", "Sunken surface", "凹陷背景"],
  ["hover", "Hover surface", "悬停背景"],
  ["active", "Selected surface", "选中背景"],
  ["foreground", "Primary text", "主要前景"],
  ["textMuted", "Secondary text", "次要前景"],
  ["textFaint", "Muted text", "弱化前景"],
  ["border", "Border", "边框"],
  ["borderStrong", "Strong border", "强调边框"],
  ["focusRing", "Focus ring", "焦点描边"],
  ["info", "Info", "信息状态"],
  ["success", "Success", "成功状态"],
  ["warning", "Warning", "警告状态"],
  ["danger", "Danger", "危险状态"],
];

const THEME_FORMULA_COLOR_FIELDS: readonly [keyof ThemePaletteColors, string, string][] = [
  ["formulaSurface", "Formula canvas", "公式画布"],
  ["formulaPlaceholder", "Placeholder", "公式占位符"],
  ["formulaPlaceholderSelected", "Selected placeholder", "选中占位符"],
  ["formulaCaret", "Formula caret", "公式光标"],
];

const THEME_SYNTAX_COLOR_FIELDS: readonly [keyof ThemePaletteColors, string, string][] = [
  ["syntaxCommand", "Command", "命令"],
  ["syntaxKeyword", "Keyword", "关键字"],
  ["syntaxOperator", "Operator", "运算符"],
  ["syntaxNumber", "Number", "数字"],
  ["syntaxBracket", "Bracket", "括号"],
  ["syntaxString", "String", "字符串"],
  ["syntaxComment", "Comment", "注释"],
  ["syntaxVariable", "Variable", "变量"],
  ["syntaxFunction", "Function", "函数"],
  ["syntaxError", "Error", "错误"],
];

const THEME_TOOLBAR_COLOR_FIELDS: readonly [keyof ThemePaletteColors, string, string][] = [
  ["toolbarCommon", "Common", "常用"],
  ["toolbarStructure", "Structure", "结构"],
  ["toolbarCalculus", "Calculus", "微积分"],
  ["toolbarMatrix", "Matrix", "矩阵"],
  ["toolbarRelation", "Relation", "关系"],
  ["toolbarGreek", "Greek", "希腊字母"],
  ["toolbarArrow", "Arrow", "箭头"],
  ["toolbarPhysics", "Physics", "物理"],
  ["toolbarSet", "Set", "集合"],
];

interface Props {
  open: boolean;
  onClose: () => void;
  onOpenFormulaHotkeys: () => void;
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

function configurationFilename() {
  return `VisualTeX-Configuration-${new Date().toISOString().slice(0, 10)}.${VISUALTEX_CONFIGURATION_EXTENSION}`;
}

function downloadConfigurationInBrowser(contents: string, filename: string) {
  const blob = new Blob([contents], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

export function SettingsDialog({
  open,
  onClose,
  onOpenFormulaHotkeys,
}: Props) {
  const dialogRef = useRef<HTMLElement>(null);
  const interfaceCustomizationDialogRef = useRef<HTMLElement>(null);
  const interfaceCustomizationOpenRef = useRef(false);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const configurationInputRef = useRef<HTMLInputElement>(null);
  const [interfaceCustomizationOpen, setInterfaceCustomizationOpen] =
    useState(false);
  const [configurationBusy, setConfigurationBusy] = useState(false);
  const [configurationStatus, setConfigurationStatus] = useState("");
  const [customTheme, setCustomTheme] =
    useState<CustomThemeState>(() => readCustomTheme());
  const theme = useEditorStore((state) => state.theme);
  const setTheme = useEditorStore((state) => state.setTheme);
  const language = useEditorStore((state) => state.language);
  const setLanguage = useEditorStore((state) => state.setLanguage);
  const zoom = useEditorStore((state) => state.zoom);
  const setZoom = useEditorStore((state) => state.setZoom);
  const editorLayout = useEditorStore((state) => state.editorLayout);
  const setEditorLayout = useEditorStore((state) => state.setEditorLayout);
  const autoPairDelimiters = useEditorStore(
    (state) => state.autoPairDelimiters,
  );
  const showLineNumbers = useEditorStore((state) => state.showLineNumbers);
  const highlightActiveLine = useEditorStore(
    (state) => state.highlightActiveLine,
  );
  const sourceEditorFontSize = useEditorStore(
    (state) => state.sourceEditorFontSize,
  );
  const formulaInsetLeft = useEditorStore((state) => state.formulaInsetLeft);
  const formulaInsetRight = useEditorStore((state) => state.formulaInsetRight);
  const formulaToolButtonSize = useEditorStore(
    (state) => state.formulaToolButtonSize,
  );
  const formulaToolButtonPadding = useEditorStore(
    (state) => state.formulaToolButtonPadding,
  );
  const formulaRowVerticalInset = useEditorStore(
    (state) => state.formulaRowVerticalInset,
  );
  const pngExportBackground = useEditorStore(
    (state) => state.pngExportBackground,
  );
  const formulaLetterFont = useEditorStore((state) => state.formulaLetterFont);
  const formulaChineseFont = useEditorStore((state) => state.formulaChineseFont);
  const setAutoPairDelimiters = useEditorStore(
    (state) => state.setAutoPairDelimiters,
  );
  const setShowLineNumbers = useEditorStore(
    (state) => state.setShowLineNumbers,
  );
  const setHighlightActiveLine = useEditorStore(
    (state) => state.setHighlightActiveLine,
  );
  const setSourceEditorFontSize = useEditorStore(
    (state) => state.setSourceEditorFontSize,
  );
  const setFormulaInsetLeft = useEditorStore(
    (state) => state.setFormulaInsetLeft,
  );
  const setFormulaInsetRight = useEditorStore(
    (state) => state.setFormulaInsetRight,
  );
  const setFormulaToolButtonSize = useEditorStore(
    (state) => state.setFormulaToolButtonSize,
  );
  const setFormulaToolButtonPadding = useEditorStore(
    (state) => state.setFormulaToolButtonPadding,
  );
  const setFormulaRowVerticalInset = useEditorStore(
    (state) => state.setFormulaRowVerticalInset,
  );
  const setPngExportBackground = useEditorStore(
    (state) => state.setPngExportBackground,
  );
  const setFormulaLetterFont = useEditorStore(
    (state) => state.setFormulaLetterFont,
  );
  const setFormulaChineseFont = useEditorStore(
    (state) => state.setFormulaChineseFont,
  );
  const personalize = useEditorStore((state) => state.personalize);
  const setPersonalize = useEditorStore((state) => state.setPersonalize);
  const suggestionCount = useEditorStore((state) => state.suggestionCount);
  const setSuggestionCount = useEditorStore((state) => state.setSuggestionCount);
  const resetUsage = useEditorStore((state) => state.resetUsage);
  const isEn = language === "en";
  const formulaLetterFamilies = formulaLetterFontFamilies(formulaLetterFont);
  const formulaChineseFamily = formulaChineseFontFamily(formulaChineseFont);

  useEffect(() => {
    if (!open) return;
    previousFocusRef.current = document.activeElement as HTMLElement | null;
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current?.querySelector<HTMLElement>("button, input, select")?.focus();
    });

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        if (interfaceCustomizationOpenRef.current) {
          interfaceCustomizationOpenRef.current = false;
          setInterfaceCustomizationOpen(false);
        } else {
          onClose();
        }
        return;
      }
      const activeDialog = interfaceCustomizationOpenRef.current
        ? interfaceCustomizationDialogRef.current
        : dialogRef.current;
      if (event.key !== "Tab" || !activeDialog) return;

      const focusable = Array.from(
        activeDialog.querySelectorAll<HTMLElement>(
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
  }, [open]);

  useEffect(() => {
    interfaceCustomizationOpenRef.current = interfaceCustomizationOpen;
    if (!interfaceCustomizationOpen) return;
    const frame = window.requestAnimationFrame(() => {
      interfaceCustomizationDialogRef.current
        ?.querySelector<HTMLElement>("button, input")
        ?.focus();
    });
    return () => window.cancelAnimationFrame(frame);
  }, [interfaceCustomizationOpen]);

  useEffect(() => {
    if (open) return;
    interfaceCustomizationOpenRef.current = false;
    setInterfaceCustomizationOpen(false);
  }, [open]);

  const openInterfaceCustomization = () => {
    interfaceCustomizationOpenRef.current = true;
    setInterfaceCustomizationOpen(true);
  };
  const closeInterfaceCustomization = () => {
    interfaceCustomizationOpenRef.current = false;
    setInterfaceCustomizationOpen(false);
  };

  const updateCustomTheme = (
    update: (current: CustomThemeState) => CustomThemeState,
  ) => {
    setCustomTheme((current) => {
      const next = update(current);
      publishCustomTheme(next);
      setTheme("custom");
      return next;
    });
  };

  const updateThemeColor = (key: keyof ThemePaletteColors, color: string) => {
    if (!/^#[0-9a-f]{6}$/i.test(color)) return;
    updateCustomTheme((current) => ({
      ...current,
      colors: {
        ...current.colors,
        [key]: color.toUpperCase(),
      },
    }));
  };

  const resetThemePalette = () => {
    const next = createDefaultCustomTheme();
    setCustomTheme(next);
    publishCustomTheme(next);
    setTheme("custom");
  };

  const renderThemeColorFields = (
    fields: readonly [keyof ThemePaletteColors, string, string][],
  ) => (
    <div className="theme-color-grid">
      {fields.map(([key, labelEn, labelZh]) => {
        const color = customTheme.colors[key];
        return (
          <label className="theme-color-row" key={key}>
            <strong>{isEn ? labelEn : labelZh}</strong>
            <span className="theme-color-control">
              <input
                type="color"
                value={color}
                aria-label={isEn ? `${labelEn} color` : `${labelZh}颜色`}
                data-theme-color-setting={key}
                onChange={(event) =>
                  updateThemeColor(key, event.currentTarget.value)
                }
              />
              <input
                key={`${key}-${color}`}
                type="text"
                defaultValue={color}
                spellCheck={false}
                aria-label={isEn ? `${labelEn} hex value` : `${labelZh}十六进制值`}
                onBlur={(event) => {
                  const value = event.currentTarget.value.trim();
                  if (/^#[0-9a-f]{6}$/i.test(value)) {
                    updateThemeColor(key, value);
                  } else {
                    event.currentTarget.value = color;
                  }
                }}
                onKeyDown={(event) => {
                  if (event.key === "Enter") event.currentTarget.blur();
                }}
              />
            </span>
          </label>
        );
      })}
    </div>
  );

  const saveConfiguration = async () => {
    if (configurationBusy) return;
    setConfigurationBusy(true);
    setConfigurationStatus("");
    try {
      const configuration = await buildVisualTexConfiguration();
      const contents = JSON.stringify(configuration, null, 2);
      const filename = configurationFilename();
      if (isTauri()) {
        const selectedPath = await save({
          defaultPath: filename,
          filters: [
            {
              name: "VisualTeX Configuration",
              extensions: [VISUALTEX_CONFIGURATION_EXTENSION],
            },
          ],
        });
        if (!selectedPath) return;
        const path = selectedPath.toLowerCase().endsWith(`.${VISUALTEX_CONFIGURATION_EXTENSION}`)
          ? selectedPath
          : `${selectedPath}.${VISUALTEX_CONFIGURATION_EXTENSION}`;
        await invoke("write_export_file", {
          path,
          dataBase64: encodeUtf8Base64(contents),
        });
      } else {
        downloadConfigurationInBrowser(contents, filename);
      }
      setConfigurationStatus(
        isEn ? "Configuration saved." : "配置文件已保存。",
      );
    } catch (reason) {
      const message =
        reason instanceof Error
          ? reason.message
          : typeof reason === "string"
            ? reason
            : "";
      setConfigurationStatus(
        message ||
          (isEn
            ? "Unable to save the configuration."
            : "无法保存配置文件。"),
      );
    } finally {
      setConfigurationBusy(false);
    }
  };

  const importConfigurationFile = async (
    event: ChangeEvent<HTMLInputElement>,
  ) => {
    const file = event.currentTarget.files?.[0];
    event.currentTarget.value = "";
    if (!file || configurationBusy) return;
    setConfigurationBusy(true);
    setConfigurationStatus("");
    try {
      const configuration = parseVisualTexConfiguration(await file.text());
      await applyVisualTexConfiguration(configuration);
      setConfigurationStatus(
        isEn
          ? "Configuration imported. Reloading VisualTeX…"
          : "配置已导入，正在重新载入 VisualTeX…",
      );
      window.setTimeout(() => window.location.reload(), 450);
    } catch (reason) {
      setConfigurationStatus(
        reason instanceof Error
          ? reason.message
          : isEn
            ? "Unable to import the configuration."
            : "无法导入配置文件。",
      );
      setConfigurationBusy(false);
    }
  };

  if (!open) return null;

  return (
    <div className="modal-backdrop" role="presentation" onMouseDown={onClose}>
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
              </div>
            </div>
            <label className="switch-row">
              <span>
                <strong>{isEn ? "Enable personalized ranking" : "启用个性化排序"}</strong>
              </span>
              <input
                type="checkbox"
                checked={personalize}
                onChange={(event) => setPersonalize(event.target.checked)}
              />
              <span className="switch-control" />
            </label>
            <label className="range-setting">
              <span>
                <strong>{isEn ? "Suggestion count" : "候选项数量"}</strong>
                <small>
                  {suggestionCount} {isEn ? "items" : "项"}
                </small>
              </span>
              <input
                type="range"
                min="3"
                max="10"
                value={suggestionCount}
                onChange={(event) => setSuggestionCount(Number(event.target.value))}
              />
            </label>
            <button
              type="button"
              className="secondary-button danger-subtle"
              onClick={resetUsage}
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
              </div>
            </div>
            <div className="editor-layout-setting">
              <span>
                <strong>{isEn ? "Editor layout" : "编辑器布局"}</strong>
              </span>
              <div
                className="theme-segment editor-layout-segment"
                role="group"
                aria-label={isEn ? "Editor layout" : "编辑器布局"}
              >
                <button
                  type="button"
                  className={editorLayout === "standard" ? "is-active" : ""}
                  aria-pressed={editorLayout === "standard"}
                  data-editor-layout-choice="standard"
                  onClick={() => setEditorLayout("standard")}
                >
                  {isEn ? "Standard" : "标准布局"}
                </button>
                <button
                  type="button"
                  className={editorLayout === "classic" ? "is-active" : ""}
                  aria-pressed={editorLayout === "classic"}
                  data-editor-layout-choice="classic"
                  onClick={() => setEditorLayout("classic")}
                >
                  {isEn ? "Classic" : "经典布局"}
                </button>
              </div>
            </div>
            <div className="theme-choice-setting">
              <span>
                <strong>{isEn ? "Colour theme" : "界面配色"}</strong>
              </span>
              <div
                className="theme-segment theme-choice-segment"
                role="group"
                aria-label={isEn ? "Colour theme" : "界面配色"}
              >
                {THEME_DEFINITIONS.map((definition) => {
                  const swatches =
                    definition.id === "custom"
                      ? ([
                          customTheme.colors.background,
                          customTheme.colors.surface,
                          customTheme.colors.accent,
                        ] as const)
                      : definition.swatches;
                  return (
                    <button
                      key={definition.id}
                      type="button"
                      className={theme === definition.id ? "is-active" : ""}
                      aria-pressed={theme === definition.id}
                      data-theme-choice={definition.id}
                      onClick={() => setTheme(definition.id)}
                    >
                      <span className="theme-choice-swatch" aria-hidden="true">
                        {swatches.map((color) => (
                          <i key={color} style={{ background: color }} />
                        ))}
                      </span>
                      <span>{isEn ? definition.labelEn : definition.labelZh}</span>
                    </button>
                  );
                })}
              </div>
            </div>
            <label className="switch-row">
              <span>
                <strong>
                  {isEn ? "Auto-pair delimiters" : "自动补全成对符号"}
                </strong>
              </span>
              <input
                type="checkbox"
                checked={autoPairDelimiters}
                onChange={(event) =>
                  setAutoPairDelimiters(event.target.checked)
                }
              />
              <span className="switch-control" />
            </label>
            <label className="range-setting">
              <span>
                <strong>{isEn ? "Formula zoom" : "公式显示缩放"}</strong>
                <small>{Math.round(zoom * 100)}%</small>
              </span>
              <input
                type="range"
                min="0.5"
                max="1.6"
                step={EDITOR_ZOOM_STEP}
                value={zoom}
                onChange={(event) => setZoom(Number(event.target.value))}
              />
            </label>
          </div>

          <div className="settings-section">
            <button
              type="button"
              className="settings-subdialog-trigger"
              data-interface-customization-trigger
              onClick={openInterfaceCustomization}
            >
              <Eye size={18} />
              <span>
                <strong>{isEn ? "Interface customization" : "界面自定义"}</strong>
              </span>
              <ChevronRight size={18} aria-hidden="true" />
            </button>
          </div>

          <div className="settings-section" data-configuration-transfer>
            <div className="settings-section-title">
              <Download size={18} />
              <div>
                <h3>{isEn ? "Configuration backup" : "配置备份与迁移"}</h3>
              </div>
            </div>
            <div className="configuration-transfer-actions">
              <button
                type="button"
                className="secondary-button"
                data-save-configuration
                disabled={configurationBusy}
                onClick={() => void saveConfiguration()}
              >
                <Download size={15} />
                {isEn ? "Save current configuration" : "保存目前配置"}
              </button>
              <button
                type="button"
                className="secondary-button"
                data-import-configuration
                disabled={configurationBusy}
                onClick={() => configurationInputRef.current?.click()}
              >
                <Upload size={15} />
                {isEn ? "Import configuration" : "导入配置"}
              </button>
              <input
                ref={configurationInputRef}
                className="configuration-file-input"
                type="file"
                accept={`.${VISUALTEX_CONFIGURATION_EXTENSION},application/json`}
                onChange={importConfigurationFile}
                tabIndex={-1}
              />
            </div>
            {configurationStatus && (
              <p className="configuration-transfer-status" role="status">
                {configurationStatus}
              </p>
            )}
          </div>

          <div className="settings-section">
            <div className="settings-section-title">
              <Languages size={18} />
              <div>
                <h3>{isEn ? "Interface language" : "界面语言"}</h3>
              </div>
            </div>
            <div className="theme-segment">
              <button
                type="button"
                className={language === "cn" ? "is-active" : ""}
                aria-pressed={language === "cn"}
                onClick={() => setLanguage("cn")}
              >
                中文
              </button>
              <button
                type="button"
                className={language === "en" ? "is-active" : ""}
                aria-pressed={language === "en"}
                onClick={() => setLanguage("en")}
              >
                English
              </button>
            </div>
          </div>
        </div>

        <footer className="dialog-footer">
          <span>{isEn ? "Settings saved automatically" : "设置已自动保存"}</span>
          <button type="button" className="primary-button" onClick={onClose}>
            {isEn ? "Done" : "完成"}
          </button>
        </footer>

        {interfaceCustomizationOpen && (
          <div
            className="settings-subdialog-backdrop"
            role="presentation"
            onMouseDown={(event) => {
              event.stopPropagation();
              closeInterfaceCustomization();
            }}
          >
            <section
              ref={interfaceCustomizationDialogRef}
              className="settings-subdialog"
              role="dialog"
              aria-modal="true"
              aria-labelledby="interface-customization-title"
              data-interface-customization-dialog
              onMouseDown={(event) => event.stopPropagation()}
            >
              <header className="dialog-header">
                <div>
                  <span className="eyebrow">INTERFACE</span>
                  <h2 id="interface-customization-title">
                    {isEn ? "Interface customization" : "界面自定义"}
                  </h2>
                </div>
                <button
                  type="button"
                  className="icon-button"
                  data-interface-customization-close
                  onClick={closeInterfaceCustomization}
                  aria-label={
                    isEn
                      ? "Close interface customization"
                      : "关闭界面自定义"
                  }
                >
                  <X size={18} />
                </button>
              </header>

              <div className="settings-subdialog-content">
                <section
                  className="theme-studio-customization"
                  aria-labelledby="theme-studio-title"
                  data-theme-studio
                >
                  <header className="theme-studio-header">
                    <div>
                      <strong id="theme-studio-title">
                        {isEn ? "Theme studio" : "界面配色工作室"}
                      </strong>
                    </div>
                    <div className="theme-studio-actions">
                      <button
                        type="button"
                        className="secondary-button"
                        data-theme-reset-preset
                        onClick={resetThemePalette}
                      >
                        {isEn ? "Reset custom" : "重置自定义"}
                      </button>
                      <button
                        type="button"
                        className={theme === "custom" ? "primary-button" : "secondary-button"}
                        data-theme-use-custom
                        onClick={() => setTheme("custom")}
                      >
                        {isEn ? "Use Custom" : "使用自定义"}
                      </button>
                    </div>
                  </header>

                  <div className="theme-preset-section">
                    <strong>{isEn ? "Ready-made palettes" : "现成配色方案"}</strong>
                    <div className="theme-preset-grid" role="group">
                      {THEME_DEFINITIONS.filter((preset) => preset.id !== "custom").map(
                        (preset) => (
                          <button
                            type="button"
                            key={preset.id}
                            className={theme === preset.id ? "is-active" : ""}
                            aria-pressed={theme === preset.id}
                            data-theme-preset={preset.id}
                            onClick={() => setTheme(preset.id)}
                          >
                            <span
                              className="theme-preset-aa"
                              style={{
                                background: preset.swatches[0],
                                color: preset.swatches[2],
                                borderColor: preset.swatches[1],
                              }}
                              aria-hidden="true"
                            >
                              Aa
                            </span>
                            <span className="theme-preset-name">
                              {isEn ? preset.labelEn : preset.labelZh}
                            </span>
                            <span className="theme-preset-swatches" aria-hidden="true">
                              {preset.swatches.map((color) => (
                                <i key={color} style={{ background: color }} />
                              ))}
                            </span>
                          </button>
                        ),
                      )}
                    </div>
                  </div>

                  <div className="theme-mode-row">
                    <strong>{isEn ? "Control appearance" : "控件明暗模式"}</strong>
                    <div className="theme-segment theme-mode-segment" role="group">
                      <button
                        type="button"
                        className={customTheme.mode === "light" ? "is-active" : ""}
                        aria-pressed={customTheme.mode === "light"}
                        onClick={() =>
                          updateCustomTheme((current) => ({
                            ...current,
                            mode: "light",
                          }))
                        }
                      >
                        {isEn ? "Light" : "浅色"}
                      </button>
                      <button
                        type="button"
                        className={customTheme.mode === "dark" ? "is-active" : ""}
                        aria-pressed={customTheme.mode === "dark"}
                        onClick={() =>
                          updateCustomTheme((current) => ({
                            ...current,
                            mode: "dark",
                          }))
                        }
                      >
                        {isEn ? "Dark" : "深色"}
                      </button>
                    </div>
                  </div>

                  <details className="theme-color-details" open>
                    <summary>{isEn ? "Interface colors" : "界面颜色"}</summary>
                    {renderThemeColorFields(THEME_CORE_COLOR_FIELDS)}
                  </details>

                  <details className="theme-color-details" open>
                    <summary>{isEn ? "Formula colors" : "公式区域颜色"}</summary>
                    {renderThemeColorFields(THEME_FORMULA_COLOR_FIELDS)}
                  </details>

                  <details className="theme-color-details">
                    <summary>{isEn ? "Source editor colors" : "源码编辑器颜色"}</summary>
                    {renderThemeColorFields(THEME_SYNTAX_COLOR_FIELDS)}
                  </details>

                  <details className="theme-color-details">
                    <summary>{isEn ? "Toolbar category colors" : "工具栏分类颜色"}</summary>
                    {renderThemeColorFields(THEME_TOOLBAR_COLOR_FIELDS)}
                  </details>

                  <div
                    className="theme-studio-preview"
                    data-theme-studio-preview
                    style={
                      {
                        "--theme-preview-accent": customTheme.colors.accent,
                        "--theme-preview-accent-hover": customTheme.colors.accentHover,
                        "--theme-preview-accent-soft": customTheme.colors.accentSoft,
                        "--theme-preview-bg": customTheme.colors.background,
                        "--theme-preview-elevated": customTheme.colors.elevated,
                        "--theme-preview-surface": customTheme.colors.surface,
                        "--theme-preview-sunken": customTheme.colors.sunken,
                        "--theme-preview-hover": customTheme.colors.hover,
                        "--theme-preview-active": customTheme.colors.active,
                        "--theme-preview-text": customTheme.colors.foreground,
                        "--theme-preview-muted": customTheme.colors.textMuted,
                        "--theme-preview-faint": customTheme.colors.textFaint,
                        "--theme-preview-border": customTheme.colors.border,
                        "--theme-preview-border-strong": customTheme.colors.borderStrong,
                        "--theme-preview-formula": customTheme.colors.formulaSurface,
                        "--theme-preview-placeholder": customTheme.colors.formulaPlaceholder,
                        "--theme-preview-danger": customTheme.colors.danger,
                        "--theme-preview-warning": customTheme.colors.warning,
                        "--theme-preview-success": customTheme.colors.success,
                      } as CSSProperties
                    }
                  >
                    <div className="theme-preview-titlebar">
                      <span className="theme-preview-window-dots" aria-hidden="true">
                        <i />
                        <i />
                        <i />
                      </span>
                      <strong>{isEn ? "Custom" : "自定义"}</strong>
                      <span className="theme-preview-status">
                        {theme === "custom"
                          ? isEn
                            ? "Active"
                            : "当前主题"
                          : isEn
                            ? "Preview"
                            : "预览"}
                      </span>
                    </div>
                    <div className="theme-preview-body">
                      <aside className="theme-preview-sidebar" aria-hidden="true">
                        <span className="is-active" />
                        <span />
                        <span />
                        <span />
                      </aside>
                      <div className="theme-preview-main">
                        <div className="theme-preview-toolbar">
                          <span />
                          <span />
                          <span />
                          <button type="button" tabIndex={-1}>Aa</button>
                        </div>
                        <div className="theme-preview-editor">
                          <div className="theme-preview-formula-card">
                            <span className="theme-preview-formula-text">x² + y² = r²</span>
                            <i />
                          </div>
                          <div className="theme-preview-controls">
                            <button type="button" tabIndex={-1}>
                              {isEn ? "Primary" : "主要操作"}
                            </button>
                            <span className="success" />
                            <span className="warning" />
                            <span className="danger" />
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </section>

                <section
                  className="formula-inset-customization source-editor-customization"
                  aria-labelledby="source-editor-customization-title"
                >
                  <header className="formula-inset-customization-header">
                    <div>
                      <strong id="source-editor-customization-title">
                        {isEn ? "Source editor" : "源码编辑器"}
                      </strong>
                      <small>
                        {isEn
                          ? "Adjust the LaTeX source text without changing formula size."
                          : "仅调整 LaTeX 源码文字，不影响公式字号。"}
                      </small>
                    </div>
                    <button
                      type="button"
                      className="secondary-button formula-inset-reset"
                      data-source-editor-font-size-reset
                      onClick={() =>
                        setSourceEditorFontSize(
                          DEFAULT_SOURCE_EDITOR_FONT_SIZE,
                        )
                      }
                    >
                      {isEn ? "Reset" : "恢复默认"}
                    </button>
                  </header>
                  <label className="range-setting formula-inset-range">
                    <span>
                      <strong>{isEn ? "Text size" : "文字大小"}</strong>
                      <small>{sourceEditorFontSize}px</small>
                    </span>
                    <input
                      type="range"
                      min={MIN_SOURCE_EDITOR_FONT_SIZE}
                      max={MAX_SOURCE_EDITOR_FONT_SIZE}
                      step="1"
                      value={sourceEditorFontSize}
                      data-source-editor-font-size-setting
                      aria-label={
                        isEn ? "Source editor text size" : "源码编辑器文字大小"
                      }
                      onChange={(event) =>
                        setSourceEditorFontSize(Number(event.target.value))
                      }
                    />
                  </label>
                </section>

                <label className="switch-row">
                  <span>
                    <strong>
                      {isEn
                        ? "Highlight formula rows"
                        : "高亮当前公式行"}
                    </strong>
                  </span>
                  <input
                    type="checkbox"
                    checked={highlightActiveLine}
                    data-highlight-active-line-setting
                    onChange={(event) =>
                      setHighlightActiveLine(event.target.checked)
                    }
                  />
                  <span className="switch-control" />
                </label>

                <label className="switch-row">
                  <span>
                    <strong>
                      {isEn
                        ? "Show formula row numbers"
                        : "显示公式行序号"}
                    </strong>
                  </span>
                  <input
                    type="checkbox"
                    checked={showLineNumbers}
                    data-show-line-numbers-setting
                    onChange={(event) =>
                      setShowLineNumbers(event.target.checked)
                    }
                  />
                  <span className="switch-control" />
                </label>

                <section
                  className="formula-toolbar-customization"
                  aria-labelledby="formula-toolbar-customization-title"
                >
                  <header className="formula-inset-customization-header">
                    <div>
                      <strong id="formula-toolbar-customization-title">
                        {isEn
                          ? "Formula toolbar buttons"
                          : "公式工具栏按钮"}
                      </strong>
                    </div>
                    <button
                      type="button"
                      className="secondary-button formula-inset-reset"
                      data-formula-tool-button-reset="true"
                      onClick={() => {
                        setFormulaToolButtonSize(
                          DEFAULT_FORMULA_TOOL_BUTTON_SIZE,
                        );
                        setFormulaToolButtonPadding(
                          DEFAULT_FORMULA_TOOL_BUTTON_PADDING,
                        );
                      }}
                    >
                      {isEn ? "Reset" : "恢复默认"}
                    </button>
                  </header>

                  <div className="formula-inset-range-grid">
                    <label className="range-setting formula-inset-range">
                      <span>
                        <strong>{isEn ? "Button size" : "按钮尺寸"}</strong>
                        <small>{formulaToolButtonSize}px</small>
                      </span>
                      <input
                        type="range"
                        min={MIN_FORMULA_TOOL_BUTTON_SIZE}
                        max={MAX_FORMULA_TOOL_BUTTON_SIZE}
                        step="1"
                        value={formulaToolButtonSize}
                        data-formula-tool-button-size-setting
                        onChange={(event) =>
                          setFormulaToolButtonSize(Number(event.target.value))
                        }
                      />
                    </label>
                    <label className="range-setting formula-inset-range">
                      <span>
                        <strong>
                          {isEn ? "Content inset" : "字符边距"}
                        </strong>
                        <small>{formulaToolButtonPadding}px</small>
                      </span>
                      <input
                        type="range"
                        min={MIN_FORMULA_TOOL_BUTTON_PADDING}
                        max={MAX_FORMULA_TOOL_BUTTON_PADDING}
                        step="1"
                        value={formulaToolButtonPadding}
                        data-formula-tool-button-padding-setting
                        onChange={(event) =>
                          setFormulaToolButtonPadding(
                            Number(event.target.value),
                          )
                        }
                      />
                    </label>
                  </div>
                </section>

                <section
                  className="formula-inset-customization"
                  aria-labelledby="formula-inset-customization-title"
                >
                  <header className="formula-inset-customization-header">
                    <div>
                      <strong id="formula-inset-customization-title">
                        {isEn
                          ? "Formula area spacing"
                          : "公式区间距"}
                      </strong>
                    </div>
                    <button
                      type="button"
                      className="secondary-button formula-inset-reset"
                      data-formula-inset-reset="true"
                      onClick={() => {
                        setFormulaInsetLeft(DEFAULT_FORMULA_INSET);
                        setFormulaInsetRight(DEFAULT_FORMULA_INSET);
                        setFormulaRowVerticalInset(
                          DEFAULT_FORMULA_ROW_VERTICAL_INSET,
                        );
                      }}
                    >
                      {isEn ? "Reset" : "恢复默认"}
                    </button>
                  </header>

                  <div className="formula-inset-range-grid">
                    <label className="range-setting formula-inset-range">
                      <span>
                        <strong>{isEn ? "Left spacing" : "左侧距离"}</strong>
                        <small>{formulaInsetLeft}px</small>
                      </span>
                      <input
                        type="range"
                        min={MIN_FORMULA_INSET}
                        max={MAX_FORMULA_INSET}
                        step="1"
                        value={formulaInsetLeft}
                        data-formula-inset-left-setting
                        onChange={(event) =>
                          setFormulaInsetLeft(Number(event.target.value))
                        }
                      />
                    </label>
                    <label className="range-setting formula-inset-range">
                      <span>
                        <strong>{isEn ? "Right spacing" : "右侧距离"}</strong>
                        <small>{formulaInsetRight}px</small>
                      </span>
                      <input
                        type="range"
                        min={MIN_FORMULA_INSET}
                        max={MAX_FORMULA_INSET}
                        step="1"
                        value={formulaInsetRight}
                        data-formula-inset-right-setting
                        onChange={(event) =>
                          setFormulaInsetRight(Number(event.target.value))
                        }
                      />
                    </label>
                    <label className="range-setting formula-inset-range">
                      <span>
                        <strong>
                          {isEn ? "Formula row vertical spacing" : "公式行上下距离"}
                        </strong>
                        <small>{formulaRowVerticalInset}px</small>
                      </span>
                      <input
                        type="range"
                        min={MIN_FORMULA_ROW_VERTICAL_INSET}
                        max={MAX_FORMULA_ROW_VERTICAL_INSET}
                        step="1"
                        value={formulaRowVerticalInset}
                        data-formula-row-vertical-inset-setting
                        onChange={(event) =>
                          setFormulaRowVerticalInset(Number(event.target.value))
                        }
                      />
                    </label>
                  </div>

                  <div
                    className="formula-inset-preview-window"
                    data-formula-inset-preview
                    style={
                      {
                        "--preview-formula-inset-left": `${formulaInsetLeft}px`,
                        "--preview-formula-inset-right": `${formulaInsetRight}px`,
                        "--preview-toolbar-button-size": `${formulaToolButtonSize}px`,
                        "--preview-toolbar-button-padding": `${formulaToolButtonPadding}px`,
                        "--preview-formula-row-vertical-inset": `${formulaRowVerticalInset}px`,
                      } as CSSProperties
                    }
                  >
                    <div className="formula-inset-preview-titlebar">
                      <i />
                      <i />
                      <i />
                      <span>{isEn ? "Live preview" : "实时预览"}</span>
                    </div>
                    <div className="formula-inset-preview-toolbar">
                      <span>
                        <MathPreview
                          latex={String.raw`\frac{a}{b}`}
                          fit
                          maximumFitScale={1}
                          fitInsetRatio={0.9}
                        />
                      </span>
                      <span>
                        <MathPreview
                          latex={String.raw`\sqrt{x}`}
                          fit
                          maximumFitScale={1}
                          fitInsetRatio={0.9}
                        />
                      </span>
                      <span>
                        <MathPreview
                          latex={String.raw`\int_a^b`}
                          fit
                          maximumFitScale={1}
                          fitInsetRatio={0.9}
                        />
                      </span>
                    </div>
                    <div className="formula-inset-preview-canvas">
                      <div className="formula-inset-preview-row">
                        <MathPreview
                          latex={String.raw`x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}`}
                          fit
                          maximumFitScale={1}
                          fitInsetRatio={0.88}
                        />
                      </div>
                    </div>
                  </div>
                </section>

                <section
                  className="formula-font-customization"
                  aria-labelledby="formula-font-customization-title"
                >
                  <header className="formula-inset-customization-header">
                    <div>
                      <strong id="formula-font-customization-title">
                        {isEn ? "Formula fonts" : "可视化公式字体"}
                      </strong>
                    </div>
                    <button
                      type="button"
                      className="secondary-button formula-inset-reset"
                      data-formula-font-reset
                      onClick={() => {
                        setFormulaLetterFont(DEFAULT_FORMULA_LETTER_FONT);
                        setFormulaChineseFont(DEFAULT_FORMULA_CHINESE_FONT);
                      }}
                    >
                      {isEn ? "Reset" : "恢复默认"}
                    </button>
                  </header>

                  <div className="formula-font-select-grid">
                    <label className="formula-font-setting">
                      <span>
                        <strong>{isEn ? "Math letter font" : "数学字母字体"}</strong>
                      </span>
                      <select
                        value={formulaLetterFont}
                        data-formula-letter-font-setting
                        onChange={(event) =>
                          setFormulaLetterFont(
                            event.currentTarget.value as FormulaLetterFont,
                          )
                        }
                      >
                        {FORMULA_LETTER_FONT_OPTIONS.map((item) => (
                          <option key={item.id} value={item.id}>
                            {item.label}
                          </option>
                        ))}
                      </select>
                    </label>

                    <label className="formula-font-setting">
                      <span>
                        <strong>{isEn ? "Chinese font" : "中文字体"}</strong>
                      </span>
                      <select
                        value={formulaChineseFont}
                        data-formula-chinese-font-setting
                        onChange={(event) =>
                          setFormulaChineseFont(
                            event.currentTarget.value as FormulaChineseFont,
                          )
                        }
                      >
                        {FORMULA_CHINESE_FONT_OPTIONS.map((item) => (
                          <option key={item.id} value={item.id}>
                            {isEn ? item.labelEn : item.labelZh}
                          </option>
                        ))}
                      </select>
                    </label>
                  </div>

                  <div
                    className="formula-font-live-preview"
                    data-formula-font-preview
                    style={
                      {
                        "--visualtex-formula-upright-font-family":
                          formulaLetterFamilies.upright,
                        "--visualtex-formula-italic-font-family":
                          formulaLetterFamilies.italic,
                        "--visualtex-formula-chinese-font-family":
                          formulaChineseFamily,
                      } as CSSProperties
                    }
                  >
                    <span>{isEn ? "Live preview" : "实时预览"}</span>
                    <MathPreview
                      latex={String.raw`E=mc^2,\quad \alpha+\beta=\gamma,\quad \text{中文公式}`}
                      staticLayout
                    />
                  </div>
                </section>

                <section
                  className="png-background-customization"
                  aria-labelledby="png-background-customization-title"
                >
                  <header className="formula-inset-customization-header">
                    <div>
                      <strong id="png-background-customization-title">
                        {isEn ? "PNG background" : "PNG 背景颜色"}
                      </strong>
                    </div>
                  </header>
                  <div className="png-background-controls">
                    <button
                      type="button"
                      className={`png-background-transparent${pngExportBackground === "transparent" ? " is-active" : ""}`}
                      data-png-background-transparent
                      aria-pressed={pngExportBackground === "transparent"}
                      onClick={() => setPngExportBackground("transparent")}
                    >
                      <span className="png-transparent-swatch" aria-hidden="true" />
                      <span>
                        <strong>{isEn ? "Transparent" : "透明"}</strong>
                      </span>
                    </button>
                    <label
                      className={`png-background-color${pngExportBackground !== "transparent" ? " is-active" : ""}`}
                    >
                      <input
                        type="color"
                        value={pngExportBackgroundPickerValue(pngExportBackground)}
                        data-png-background-color-setting
                        onChange={(event) =>
                          setPngExportBackground(event.currentTarget.value as `#${string}`)
                        }
                      />
                      <span>
                        <strong>{isEn ? "Custom colour" : "自定义颜色"}</strong>
                      </span>
                    </label>
                  </div>
                </section>
              </div>

              <footer className="dialog-footer">
                <span>
                  {isEn
                    ? "Changes apply immediately"
                    : "修改会立即生效"}
                </span>
                <button
                  type="button"
                  className="primary-button"
                  onClick={closeInterfaceCustomization}
                >
                  {isEn ? "Done" : "完成"}
                </button>
              </footer>
            </section>
          </div>
        )}
      </section>
    </div>
  );
}
