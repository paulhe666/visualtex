import { useEffect, useRef, useState } from "react";
import {
  ArrowLeft,
  ArrowRight,
  Check,
  Code2,
  Download,
  FileCode2,
  FileImage,
  FileText,
  FolderOpen,
  Keyboard,
  MousePointerClick,
  PanelLeft,
  Settings2,
  X,
} from "lucide-react";
import { MathPreview } from "./MathPreview";
import { VisualTeXLogo } from "./VisualTeXLogo";
import type { Language } from "../stores/editorStore";

interface Props {
  open: boolean;
  language: Language;
  onFinish: () => void;
}

type StepId =
  | "welcome"
  | "library"
  | "keyboard"
  | "hotkeys-tiles"
  | "layouts-themes"
  | "code-format"
  | "export"
  | "input-behavior";

interface TutorialStep {
  id: StepId;
  title: string;
  description: string;
}

export function tutorialSteps(language: Language): TutorialStep[] {
  const isEn = language === "en";
  return [
    {
      id: "welcome",
      title: isEn ? "Welcome to VisualTeX" : "欢迎使用 VisualTeX",
      description: isEn
        ? "Write formulas naturally and inspect the source whenever you need it."
        : "用熟悉的方式输入公式，需要时随时查看源码。",
    },
    {
      id: "library",
      title: isEn ? "Start from the formula library" : "从公式库开始",
      description: isEn
        ? "Choose a structure or symbol to insert it at the cursor."
        : "选择结构或符号，它会直接插入当前光标。",
    },
    {
      id: "keyboard",
      title: isEn ? "Keep your hands on the keyboard" : "保持双手在键盘上",
      description: isEn
        ? "A few keys cover line creation, navigation, and deletion."
        : "几个按键就能完成换行、跳转和删除。",
    },
    {
      id: "hotkeys-tiles",
      title: isEn ? "Turn formulas into fast shortcuts" : "把常用公式变成快捷入口",
      description: isEn
        ? "Right-click any formula tool or tile to bind a hotkey. Save the current formula as a custom tile, then organize it by section, colour, and shortcut."
        : "右键公式工具或磁贴即可绑定快捷键；还可以把当前公式保存为自定义磁贴，再按分区、颜色和快捷键管理。",
    },
    {
      id: "layouts-themes",
      title: isEn ? "Choose your layout and colour theme" : "选择适合你的布局与主题",
      description: isEn
        ? "Open Settings → Appearance & editing to switch between the Standard and Classic layouts and five complete colour themes."
        : "在“设置 → 外观与编辑”中切换标准布局、经典布局和五套完整主题。",
    },
    {
      id: "code-format",
      title: isEn ? "Switch the LaTeX code format" : "切换 LaTeX 代码格式",
      description: isEn
        ? "Choose an independent or combined environment from the top bar. The source panel and copied output update immediately."
        : "从顶部选择单公式或多公式环境；下方源码区和复制结果会立即按所选格式更新。",
    },
    {
      id: "export",
      title: isEn ? "Export the current document" : "导出当前公式文档",
      description: isEn
        ? "Open Export in the top bar to download the current document as Markdown, SVG, or PNG through your browser."
        : "从顶部打开“导出”，即可通过浏览器将当前公式文档下载为 Markdown、SVG 或 PNG。",
    },
    {
      id: "input-behavior",
      title: isEn ? "Customize input behavior" : "自定义操作逻辑",
      description: isEn
        ? "Control plain-text math conversion, automatic exits from scripts, accents, and font commands, plus the large command suggestion panels."
        : "控制普通字符数学转义、上下标、重音和字体命令的自动跳出，以及大型命令候选框。",
    },
  ];
}

export function OnboardingTour({ open, language, onFinish }: Props) {
  const [step, setStep] = useState(0);
  const dialogRef = useRef<HTMLElement>(null);
  const isEn = language === "en";
  const steps = tutorialSteps(language);
  const current = steps[Math.min(step, steps.length - 1)];
  const lastStep = step === steps.length - 1;

  useEffect(() => {
    if (!open) return;
    setStep(0);
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current?.querySelector<HTMLElement>("button")?.focus();
    });
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onFinish();
        return;
      }
      if (event.key !== "Tab" || !dialogRef.current) return;
      const focusable = Array.from(
        dialogRef.current.querySelectorAll<HTMLElement>(
          'button:not(:disabled), [tabindex]:not([tabindex="-1"])',
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
    };
  }, [open, onFinish]);

  useEffect(() => {
    if (step >= steps.length) setStep(Math.max(steps.length - 1, 0));
  }, [step, steps.length]);

  if (!open || !current) return null;

  return (
    <div className="onboarding-backdrop">
      <section
        ref={dialogRef}
        className="onboarding-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="onboarding-title"
      >
        <header className="onboarding-header">
          <div className="onboarding-brand">
            <span><VisualTeXLogo className="onboarding-brand-logo" /></span>
            <strong>VisualTeX</strong>
          </div>
          <button
            type="button"
            className="icon-button compact"
            onClick={onFinish}
            aria-label={isEn ? "Close tutorial" : "关闭教程"}
          >
            <X size={16} />
          </button>
        </header>

        <div className="onboarding-content" aria-live="polite">
          <div className="onboarding-copy">
            <span>{String(step + 1).padStart(2, "0")}</span>
            <h2 id="onboarding-title">{current.title}</h2>
            <p>{current.description}</p>
          </div>

          <div className={`onboarding-stage step-${current.id}`}>
            {current.id === "welcome" && (
              <div className="onboarding-welcome-mark">
                <span><VisualTeXLogo className="onboarding-welcome-logo" /></span>
                <div>
                  <strong>VisualTeX</strong>
                  <small>{isEn ? "Formula workspace for the web" : "网页版公式工作台"}</small>
                </div>
              </div>
            )}

            {current.id === "library" && (
              <div className="onboarding-library-demo">
                <div className="onboarding-library-rail">
                  <PanelLeft size={15} />
                  <span>{isEn ? "Formula tools" : "公式工具"}</span>
                </div>
                <div className="onboarding-formula-grid">
                  {["\\frac{a}{b}", "\\sqrt{x}", "\\int_a^b f(x)\\,dx", "\\sum_{i=1}^{n} a_i"].map((latex) => (
                    <span key={latex}><MathPreview latex={latex} /></span>
                  ))}
                </div>
              </div>
            )}

            {current.id === "keyboard" && (
              <div className="onboarding-editor-demo">
                <div className="onboarding-formula-line">
                  <MathPreview latex="\\int_{-\\infty}^{\\infty} e^{-x^2}\\,dx = \\sqrt{\\pi}" />
                </div>
                <div className="onboarding-key-row">
                  <span><Keyboard size={14} /><kbd>Enter</kbd><small>{isEn ? "New line" : "新建一行"}</small></span>
                  <span><kbd>↑ ↓</kbd><small>{isEn ? "Switch rows" : "切换公式行"}</small></span>
                  <span><kbd>Tab</kbd><small>{isEn ? "Next field" : "下个位置"}</small></span>
                  <span><kbd>⌫</kbd><small>{isEn ? "Delete empty line" : "删除空行"}</small></span>
                </div>
              </div>
            )}

            {current.id === "hotkeys-tiles" && (
              <div className="onboarding-hotkeys-tiles-demo">
                <section className="onboarding-hotkey-guide">
                  <header>
                    <Keyboard size={17} />
                    <strong>{isEn ? "Formula hotkeys" : "公式快捷键"}</strong>
                  </header>
                  <div className="onboarding-hotkey-flow">
                    <span>
                      <b>1</b>
                      <small>{isEn ? "Right-click a tool or tile" : "右键公式工具或磁贴"}</small>
                    </span>
                    <i><ArrowRight size={14} /></i>
                    <span>
                      <b>2</b>
                      <small>{isEn ? "Press and assign a shortcut" : "按下组合键并完成绑定"}</small>
                    </span>
                    <i><ArrowRight size={14} /></i>
                    <span>
                      <kbd>⌥1</kbd>
                      <small>{isEn ? "Insert at the active cursor" : "在当前公式光标处插入"}</small>
                    </span>
                  </div>
                  <div className="onboarding-hotkey-manager-note">
                    <Settings2 size={14} />
                    <span>{isEn ? "Settings → Manage formula hotkeys" : "设置 → 管理公式快捷键，可统一修改或删除"}</span>
                  </div>
                </section>

                <section className="onboarding-custom-tile-guide">
                  <header>
                    <PanelLeft size={17} />
                    <strong>{isEn ? "Custom formula tiles" : "自定义公式磁贴"}</strong>
                  </header>
                  <div className="onboarding-custom-tile-preview">
                    <MathPreview latex="\\int_0^1 x^2\\,\\mathrm{d}x" />
                  </div>
                  <button type="button" tabIndex={-1}>
                    {isEn ? "Save current formula" : "保存当前公式"}
                  </button>
                  <small>
                    {isEn
                      ? "Create sections, then right-click a tile to change its shortcut, colour, or section."
                      : "先建立分区；保存后右键磁贴，可调整快捷键、颜色和所属分区。"}
                  </small>
                </section>
              </div>
            )}

            {current.id === "layouts-themes" && (
              <div className="onboarding-layout-theme-demo">
                <div className="onboarding-layout-choice-list">
                  <article>
                    <div className="onboarding-layout-mini is-standard" aria-hidden="true"><i /><span /><b /></div>
                    <strong>{isEn ? "Standard layout" : "标准布局"}</strong>
                    <small>{isEn ? "Editor with side tools and tiles" : "公式区配合侧边工具和磁贴"}</small>
                  </article>
                  <article className="is-selected">
                    <div className="onboarding-layout-mini is-classic" aria-hidden="true"><i /><span /><b /></div>
                    <strong>{isEn ? "Classic layout" : "经典布局"}</strong>
                    <small>{isEn ? "The default: bottom tools with a right tile rail" : "默认布局：底部公式工具＋右侧磁贴栏"}</small>
                  </article>
                </div>

                <div className="onboarding-theme-guide">
                  <header>
                    <Settings2 size={17} />
                    <strong>{isEn ? "Five complete themes" : "五套完整界面主题"}</strong>
                  </header>
                  <div className="onboarding-theme-swatches">
                    {[
                      ["light", isEn ? "Light" : "浅色"],
                      ["beige", isEn ? "Warm beige" : "暖米色"],
                      ["dark", isEn ? "Dark" : "深色"],
                      ["purple", isEn ? "Deep purple" : "深紫色"],
                      ["green", isEn ? "Deep green" : "深绿色"],
                    ].map(([themeId, label]) => (
                      <span className={`is-${themeId}`} key={themeId}>
                        <i><b /><b /><b /></i>
                        <small>{label}</small>
                      </span>
                    ))}
                  </div>
                  <div className="onboarding-theme-sync-note">
                    <Check size={14} />
                    <span>{isEn ? "The editor, panels, tiles, and source area switch together" : "编辑区、面板、磁贴和源码区会同步切换主题"}</span>
                  </div>
                </div>
              </div>
            )}

            {current.id === "code-format" && (
              <div className="onboarding-code-format-demo">
                <div className="onboarding-code-format-toolbar">
                  <Code2 size={16} />
                  <strong>{isEn ? "LaTeX code format" : "LaTeX 代码格式"}</strong>
                  <span>⌄</span>
                </div>
                <div className="onboarding-code-format-choice">
                  <span>
                    <small>{isEn ? "Independent" : "单公式环境"}</small>
                    <strong>\\[ ... \\]</strong>
                  </span>
                  <span className="is-selected">
                    <Check size={14} />
                    <small>{isEn ? "Combined" : "多公式环境"}</small>
                    <strong>align*</strong>
                  </span>
                </div>
                <i><ArrowRight size={15} /></i>
                <pre>{"\\begin{align*}\na &= b + c \\\\\\nd &= e - f\n\\end{align*}"}</pre>
              </div>
            )}

            {current.id === "export" && (
              <div className="onboarding-export-demo">
                <div className="onboarding-export-trigger">
                  <Download size={17} />
                  <strong>{isEn ? "Export" : "导出"}</strong>
                  <span>⌄</span>
                </div>
                <div className="onboarding-export-formats">
                  <span><FileText size={18} /><strong>Markdown</strong><small>.md</small></span>
                  <span><FileCode2 size={18} /><strong>SVG</strong><small>.svg</small></span>
                  <span><FileImage size={18} /><strong>PNG</strong><small>.png</small></span>
                </div>
                <div className="onboarding-export-path">
                  <span><FolderOpen size={16} /><small>{isEn ? "Browser destination" : "浏览器下载位置"}</small></span>
                  <strong>{isEn ? "Your default Downloads folder" : "浏览器默认下载目录"}</strong>
                </div>
              </div>
            )}

            {current.id === "input-behavior" && (
              <div className="onboarding-input-behavior-demo">
                <div className="onboarding-input-behavior-heading">
                  <MousePointerClick size={17} />
                  <strong>{isEn ? "Input behavior" : "操作逻辑"}</strong>
                </div>
                <div className="onboarding-input-behavior-options">
                  <span>
                    <div><strong>{isEn ? "Auto-convert common math input" : "常用数学输入自动转义"}</strong><small>{isEn ? "alpha, >=, pp, hat and more" : "支持 alpha、>=、pp、hat 等输入"}</small></div>
                    <i className="is-on" />
                  </span>
                  <span>
                    <div><strong>{isEn ? "Exit font command after input" : "字体命令输入后跳出"}</strong><small>{isEn ? "Off: type multiple characters" : "关闭后可连续输入多个字符"}</small></div>
                    <i />
                  </span>
                  <span>
                    <div><strong>{isEn ? "Structured command suggestions" : "求和、积分等结构候选框"}</strong><small>{isEn ? "Large VisualTeX command panel" : "VisualTeX 大型命令候选框"}</small></div>
                    <i className="is-on" />
                  </span>
                </div>
                <div className="onboarding-input-behavior-example">
                  <code>\\mathbb&#123;AB&#125;</code>
                  <kbd>Enter</kbd>
                  <small>{isEn ? "Confirm and leave the font scope" : "确认并退出字体作用域"}</small>
                </div>
              </div>
            )}
          </div>
        </div>

        <footer className="onboarding-footer">
          <button type="button" className="onboarding-skip" onClick={onFinish}>
            {isEn ? "Skip" : "跳过"}
          </button>
          <div className="onboarding-progress" aria-label={isEn ? "Tutorial progress" : "教程进度"}>
            {steps.map((item, index) => (
              <span key={item.id} className={index === step ? "is-active" : index < step ? "is-complete" : ""} />
            ))}
          </div>
          <div className="onboarding-actions">
            {step > 0 && (
              <button type="button" className="secondary-button" onClick={() => setStep((value) => value - 1)}>
                <ArrowLeft size={15} />
                {isEn ? "Back" : "上一步"}
              </button>
            )}
            <button
              type="button"
              className="primary-button"
              onClick={() => lastStep ? onFinish() : setStep((value) => value + 1)}
            >
              {lastStep ? <Check size={15} /> : null}
              {lastStep ? (isEn ? "Start editing" : "开始使用") : (isEn ? "Continue" : "继续")}
              {!lastStep ? <ArrowRight size={15} /> : null}
            </button>
          </div>
        </footer>
      </section>
    </div>
  );
}
