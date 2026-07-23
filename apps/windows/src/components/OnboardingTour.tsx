import { useEffect, useRef, useState } from "react";
import {
  ArrowLeft,
  ArrowRight,
  Check,
  Code2,
  Download,
  FileText,
  Grid3X3,
  Keyboard,
  Menu,
  PanelLeft,
  Presentation,
  Puzzle,
  RefreshCw,
  ScanLine,
  Settings2,
  Subscript,
  Superscript,
  ToggleLeft,
  Trash2,
  Type,
  X,
} from "lucide-react";
import { MathPreview } from "./MathPreview";
import { VisualTeXLogo } from "./VisualTeXLogo";
import type { Language } from "../stores/editorStore";
import type { DesktopPlatform } from "../platform";

interface Props {
  open: boolean;
  language: Language;
  platform: DesktopPlatform;
  onFinish: () => void;
}

type StepId =
  | "welcome"
  | "library"
  | "keyboard"
  | "matrix-fonts"
  | "input-behavior"
  | "code-format"
  | "export"
  | "ocr-setup"
  | "paste-image"
  | "mac-office-enable"
  | "mac-office-manage"
  | "windows-office-manage"
  | "updates";

interface TutorialStep {
  id: StepId;
  title: string;
  description: string;
}

export function tutorialSteps(language: Language, platform: DesktopPlatform): TutorialStep[] {
  const isEn = language === "en";
  const steps: TutorialStep[] = [
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
      id: "matrix-fonts",
      title: isEn ? "Build matrices and styled symbols" : "快速插入矩阵与字体变体",
      description: isEn
        ? "Choose matrix dimensions up to 10 × 10, then insert blackboard bold, calligraphic, Fraktur, bold, and other styled symbols from the formula tools."
        : "通过尺寸面板快速插入最高 10×10 的矩阵，并从公式工具中使用黑板粗体、花体、哥特体、粗体等字体变体。",
    },
    {
      id: "input-behavior",
      title: isEn ? "Control each input scope independently" : "独立控制每一种输入逻辑",
      description: isEn
        ? "Superscript and subscript auto-exit are independent. Styled-font input can exit after one character, or stay open and grow until Enter. Differential d is normalized upright in derivative and integral contexts."
        : "上标与下标的自动跳出可独立设置；字体变体既可输入一个字符后跳出，也可连续扩展输入框并按 Enter 结束。导数和积分语境中的微分 d 会自动规范为正体。",
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
      title: isEn ? "Export from one place" : "从一个入口完成导出",
      description: isEn
        ? "Use Export to save Markdown, SVG, or PNG. Choose the format, edit the file name, and select any permitted destination path."
        : "使用统一“导出”入口保存 Markdown、SVG 或 PNG；可选择格式、修改文件名并将结果保存到自选路径。",
    },
    {
      id: "ocr-setup",
      title: isEn ? "First-time OCR setup" : "第一次使用 OCR",
      description:
        platform === "macos"
          ? isEn
            ? "The complete macOS package includes Python, PaddleOCR, and the default M model. Setup verifies and extracts the local archives."
            : "完整 macOS 包已内置 Python、PaddleOCR 与默认 M 模型；首次安装只在本机校验并解压。"
          : platform === "windows"
            ? isEn
              ? "The Windows installer checks for a compatible 64-bit Python 3.9–3.13 runtime. OCR setup is available after that prerequisite is present."
              : "Windows 安装程序会检测兼容的 64 位 Python 3.9–3.13；具备该环境后即可安装 OCR。"
            : isEn
              ? "Open Formula image OCR from the app menu and follow the local runtime setup."
              : "从应用菜单打开“图片公式识别”，并按提示准备本地运行环境。",
    },
    {
      id: "paste-image",
      title: isEn ? "Paste images directly afterward" : "之后直接粘贴图片",
      description: isEn
        ? "Once OCR is ready, paste an image into a formula field and the result returns to the saved cursor."
        : "OCR 准备好后，把光标放进公式框即可粘贴图片识别，结果会插回原光标位置。",
    },
  ];

  if (platform === "macos") {
    steps.push(
      {
        id: "mac-office-enable",
        title: isEn ? "Enable VisualTeX in Office" : "在 Office 中添加 VisualTeX",
        description: isEn
          ? "In Word or PowerPoint, open Home → Add-ins → My Add-ins or Developer Add-ins, then choose VisualTeX. Repeat this after a restart if Office hides the sideloaded tab."
          : "在 Word 或 PowerPoint 中打开“开始 → 加载项 → 我的加载项/开发人员加载项”，再选择 VisualTeX。若重启后标签页消失，可按同一路径再次添加。",
      },
      {
        id: "mac-office-manage",
        title: isEn ? "Manage the macOS integration" : "管理或卸载 macOS 集成",
        description: isEn
          ? "Open Settings → macOS Office integration. Disable startup without removing the add-in, stop the current companion, or choose Uninstall Office integration to remove it."
          : "打开“设置 → macOS Office 集成”。可单独关闭开机启动、停止当前伴侣服务，或点击“卸载 Office 集成”完整移除。",
      },
    );
  } else if (platform === "windows") {
    steps.push({
      id: "windows-office-manage",
      title: isEn ? "Use VisualTeX in Word and PowerPoint" : "在 Word 和 PowerPoint 中使用 VisualTeX",
      description: isEn
        ? "The native Office add-in lets Word insert inline or display formulas as editable OLE or native OMML, convert formats, update equation numbers, and insert references. In PowerPoint, create or edit formulas, convert them to native OLE, or export them as pictures. Formulas stay with the document and can be reopened by double-clicking."
        : "新版原生 Office 插件支持在 Word 中插入 OLE 或 OMML 行内、行间公式，并可编辑、格式互转、更新公式编号和插入引用；在 PowerPoint 中可新建或编辑公式、转为原生 OLE，或导出为图片。公式会随文档保存，双击即可继续编辑。",
    });
  }

  steps.push({
    id: "updates",
    title: isEn ? "Check for updates anytime" : "随时检查更新",
    description: isEn
      ? "Open the top-left menu and choose Check for updates. The same action is also available in Settings."
      : "打开左上角菜单，选择“检查更新”；也可以在设置中执行同一操作。",
  });
  return steps;
}

export function OnboardingTour({ open, language, platform, onFinish }: Props) {
  const [step, setStep] = useState(0);
  const dialogRef = useRef<HTMLElement>(null);
  const isEn = language === "en";
  const steps = tutorialSteps(language, platform);
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

  const pasteShortcut = platform === "windows" ? "Ctrl+V" : "⌘V";
  const platformLabel =
    platform === "windows"
      ? isEn ? "Formula workspace for Windows" : "Windows 公式工作台"
      : platform === "macos"
        ? isEn ? "Formula workspace for macOS" : "macOS 公式工作台"
        : isEn ? "Visual formula workspace" : "可视化公式工作台";

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
                  <small>{platformLabel}</small>
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
                  <span><kbd>Tab</kbd><small>{isEn ? "Next field" : "下个位置"}</small></span>
                  <span><kbd>⌫</kbd><small>{isEn ? "Delete empty line" : "删除空行"}</small></span>
                </div>
              </div>
            )}

            {current.id === "matrix-fonts" && (
              <div className="onboarding-matrix-font-demo">
                <div className="onboarding-matrix-picker-preview">
                  <span className="onboarding-demo-heading">
                    <Grid3X3 size={16} />
                    <strong>{isEn ? "Matrix size" : "矩阵尺寸"}</strong>
                    <small>10 × 10</small>
                  </span>
                  <div className="onboarding-mini-matrix-grid">
                    {Array.from({ length: 16 }, (_, index) => (
                      <i key={index} className={index < 11 ? "is-selected" : ""} />
                    ))}
                  </div>
                  <b>3 × 4</b>
                </div>
                <i className="onboarding-feature-arrow"><ArrowRight size={15} /></i>
                <div className="onboarding-font-variants-preview">
                  <span className="onboarding-demo-heading">
                    <Type size={16} />
                    <strong>{isEn ? "Font variants" : "字体变体"}</strong>
                  </span>
                  <div>
                    <span><MathPreview latex="\\mathbb{R}" /><small>mathbb</small></span>
                    <span><MathPreview latex="\\mathcal{F}" /><small>mathcal</small></span>
                    <span><MathPreview latex="\\mathfrak{g}" /><small>mathfrak</small></span>
                  </div>
                </div>
              </div>
            )}

            {current.id === "input-behavior" && (
              <div className="onboarding-input-behavior-demo">
                <div className="onboarding-input-toggle-list">
                  <span>
                    <Superscript size={17} />
                    <strong>{isEn ? "Superscript auto-exit" : "上标输入后跳出"}</strong>
                    <i className="is-on"><b /></i>
                  </span>
                  <span>
                    <Subscript size={17} />
                    <strong>{isEn ? "Subscript auto-exit" : "下标输入后跳出"}</strong>
                    <i><b /></i>
                  </span>
                </div>
                <div className="onboarding-wrapper-input-preview">
                  <small>{isEn ? "Continuous styled input" : "字体变体连续输入"}</small>
                  <span><MathPreview latex="\\mathbb{AB}" /><i /></span>
                  <kbd>Enter</kbd>
                  <ArrowRight size={15} />
                  <MathPreview latex="\\mathbb{AB}C" />
                </div>
                <div className="onboarding-upright-preview">
                  <small>{isEn ? "Automatic upright differential" : "微分正体自动规范"}</small>
                  <MathPreview latex="\\frac{\\mathrm{d}\\Phi}{\\mathrm{d}\\theta}" />
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
                <div className="onboarding-export-formats">
                  {[
                    ["Markdown", ".md"],
                    ["SVG", ".svg"],
                    ["PNG", ".png"],
                  ].map(([name, extension], index) => (
                    <span key={name} className={index === 1 ? "is-selected" : ""}>
                      <Download size={18} />
                      <strong>{name}</strong>
                      <small>{extension}</small>
                    </span>
                  ))}
                </div>
                <div className="onboarding-export-path">
                  <FileText size={16} />
                  <span>
                    <small>{isEn ? "Save to" : "保存路径"}</small>
                    <strong>{isEn ? "Documents / formula.svg" : "文档 / formula.svg"}</strong>
                  </span>
                  <button type="button" tabIndex={-1}>
                    {isEn ? "Choose…" : "选择…"}
                  </button>
                </div>
              </div>
            )}

            {current.id === "ocr-setup" && (
              <div className="onboarding-ocr-setup-demo">
                <span>
                  <ScanLine size={20} />
                  <strong>{isEn ? "Open Formula image OCR" : "打开“图片公式识别”"}</strong>
                  <small>{isEn ? "From the app menu" : "从应用菜单进入"}</small>
                </span>
                <i><ArrowRight size={15} /></i>
                <span>
                  <Download size={20} />
                  <strong>{isEn ? "Prepare runtime" : "准备 OCR 环境"}</strong>
                  <small>{platform === "windows" ? "Python 3.9–3.13 x64" : isEn ? "One-time setup" : "只需安装一次"}</small>
                </span>
                <i><ArrowRight size={15} /></i>
                <span>
                  <Check size={20} />
                  <strong>{isEn ? "Verify locally" : "完成本地校验"}</strong>
                  <small>{isEn ? "Ready for recognition" : "随后即可识别"}</small>
                </span>
              </div>
            )}

            {current.id === "paste-image" && (
              <div className="onboarding-paste-demo">
                <div className="onboarding-paste-field">
                  <span className="onboarding-paste-caret" />
                  <small>{isEn ? "Formula field" : "公式输入框"}</small>
                </div>
                <span className="onboarding-paste-shortcut">
                  <ScanLine size={20} />
                  <strong>{isEn ? "Paste formula image" : "粘贴公式图片"}</strong>
                  <kbd>{pasteShortcut}</kbd>
                </span>
                <i><ArrowRight size={15} /></i>
                <span className="onboarding-paste-result">
                  <Code2 size={20} />
                  <strong>{isEn ? "Inserted at saved cursor" : "插回原光标位置"}</strong>
                  <MathPreview latex="\\frac{a+b}{c}" />
                </span>
              </div>
            )}

            {current.id === "mac-office-enable" && (
              <div className="onboarding-workflow-demo onboarding-office-demo">
                <span>
                  <FileText size={20} />
                  <Presentation size={20} />
                  <strong>Word / PowerPoint</strong>
                </span>
                <i><ArrowRight size={15} /></i>
                <span>
                  <Puzzle size={22} />
                  <strong>{isEn ? "Home → Add-ins" : "开始 → 加载项"}</strong>
                </span>
                <i><ArrowRight size={15} /></i>
                <span className="is-selected">
                  <Check size={22} />
                  <strong>VisualTeX</strong>
                </span>
              </div>
            )}

            {current.id === "mac-office-manage" && (
              <div className="onboarding-workflow-demo onboarding-office-demo">
                <span>
                  <Settings2 size={22} />
                  <strong>{isEn ? "Settings" : "设置"}</strong>
                </span>
                <i><ArrowRight size={15} /></i>
                <span>
                  <ToggleLeft size={22} />
                  <strong>{isEn ? "Disable startup" : "关闭开机启动"}</strong>
                </span>
                <i><ArrowRight size={15} /></i>
                <span className="is-danger">
                  <Trash2 size={22} />
                  <strong>{isEn ? "Uninstall integration" : "卸载 Office 集成"}</strong>
                </span>
              </div>
            )}

            {current.id === "windows-office-manage" && (
              <div className="onboarding-windows-office-demo">
                <span>
                  <FileText size={20} />
                  <Presentation size={20} />
                  <strong>{isEn ? "VisualTeX native Office add-in" : "VisualTeX 原生 Office 插件"}</strong>
                </span>
                <div>
                  <span><Code2 size={19} /><strong>{isEn ? "OLE / OMML · inline / display" : "OLE / OMML · 行内 / 行间"}</strong></span>
                  <span><RefreshCw size={19} /><strong>{isEn ? "Edit · convert · number · reference" : "编辑 · 互转 · 编号 · 引用"}</strong></span>
                  <span><Check size={19} /><strong>{isEn ? "Double-click edit · save · export" : "双击编辑 · 随文档保存 · 导出图片"}</strong></span>
                </div>
              </div>
            )}

            {current.id === "updates" && (
              <div className="onboarding-update-demo">
                <span>
                  <Menu size={20} />
                  <strong>{isEn ? "Open app menu" : "打开左上角菜单"}</strong>
                </span>
                <i><ArrowRight size={15} /></i>
                <span className="onboarding-update-menu-item">
                  <RefreshCw size={20} />
                  <strong>{isEn ? "Check for updates" : "检查更新"}</strong>
                </span>
                <i><ArrowRight size={15} /></i>
                <span>
                  <Check size={20} />
                  <strong>{isEn ? "Review the result" : "查看版本结果"}</strong>
                </span>
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
