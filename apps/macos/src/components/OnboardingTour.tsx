import { useEffect, useRef, useState } from "react";
import {
  ArrowLeft,
  ArrowRight,
  Check,
  Code2,
  Download,
  FileText,
  Keyboard,
  Menu,
  PanelLeft,
  Power,
  Presentation,
  RefreshCw,
  ScanLine,
  Settings2,
  ToggleLeft,
  Trash2,
  X,
} from "lucide-react";
import { MathPreview } from "./MathPreview";
import { VisualTeXLogo } from "./VisualTeXLogo";
import { PowerPointAddinGuide } from "./PowerPointAddinGuide";
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
  | "code-format"
  | "ocr-setup"
  | "paste-image"
  | "mac-word-plugin"
  | "mac-powerpoint-load"
  | "mac-powerpoint-use"
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
      id: "code-format",
      title: isEn ? "Switch the LaTeX code format" : "切换 LaTeX 代码格式",
      description: isEn
        ? "Choose an independent or combined environment from the top bar. The source panel and copied output update immediately."
        : "从顶部选择单公式或多公式环境；下方源码区和复制结果会立即按所选格式更新。",
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
        id: "mac-word-plugin",
        title: isEn ? "Use the native VisualTeX tab in Word" : "在 Word 中使用原生 VisualTeX 标签页",
        description: isEn
          ? "After installing the DOTM and restarting Word, open the VisualTeX tab. Insert picture or native OMML formulas, edit the selected formula, convert an image formula to Word math, and manage numbering or cross-references."
          : "安装 DOTM 并重启 Word 后，打开“VisualTeX”标签页。可以插入图片公式或原生 OMML 公式、编辑所选公式、把图片公式转换成 Word 原生公式，并管理编号与交叉引用。",
      },
      {
        id: "mac-powerpoint-load",
        title: isEn ? "Register the PPAM once in PowerPoint" : "在 PowerPoint 中登记一次 PPAM",
        description: isEn
          ? "Open Tools → PowerPoint Add-ins, click +, select the fixed VisualTeX.ppam file, keep VisualTeX checked, and restart PowerPoint. Later VisualTeX updates reuse the same registered path."
          : "打开“工具 → PowerPoint 加载项”，点击＋，选择固定路径下的 VisualTeX.ppam，保持 VisualTeX 勾选并重启 PowerPoint。后续 VisualTeX 更新会继续复用这个登记路径。",
      },
      {
        id: "mac-powerpoint-use",
        title: isEn ? "Create and edit formulas in PowerPoint" : "在 PowerPoint 中新建与编辑公式",
        description: isEn
          ? "Open the VisualTeX tab and choose New formula. Select an existing VisualTeX formula and use Edit selected formula or double-click it to reopen the editor; Delete selected formula removes it cleanly."
          : "打开“VisualTeX”标签页并点击“新建公式”。选中已有 VisualTeX 公式后，可以点击“编辑所选公式”或直接双击重新打开编辑器；“删除所选公式”会完整删除公式对象。",
      },
    );
  } else if (platform === "windows") {
    steps.push({
      id: "windows-office-manage",
      title: isEn ? "Manage the Windows OLE service" : "管理 Windows OLE 服务",
      description: isEn
        ? "When OLE is selected in the installer, setup completes the certificate, catalog, Ribbon cache, and background registration automatically. In Settings → Windows Office integration, stop the current companion, disable startup, or remove the OLE manifest."
        : "安装程序勾选 OLE 后，会自动完成证书、可信目录、Ribbon 缓存和后台注册，不需要额外配置。可在“设置 → Windows Office 集成”中停止当前伴侣服务、关闭开机启动，或移除 OLE manifest。",
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

            {current.id === "mac-word-plugin" && (
              <div className="onboarding-native-ribbon-demo is-word">
                <div className="onboarding-native-ribbon-title">
                  <FileText size={17} />
                  <strong>Microsoft Word</strong>
                  <span>VisualTeX</span>
                </div>
                <div className="onboarding-native-ribbon-tools">
                  <span><b>OMML</b><small>{isEn ? "Inline" : "行内公式"}</small></span>
                  <span><b>OMML</b><small>{isEn ? "Display" : "行间公式"}</small></span>
                  <span><Check size={16} /><small>{isEn ? "Edit selected" : "编辑所选公式"}</small></span>
                  <span><RefreshCw size={16} /><small>{isEn ? "Update numbers" : "更新公式编号"}</small></span>
                </div>
                <p>{isEn ? "The DOTM loads automatically after Word restarts." : "重启 Word 后，DOTM 会从 Startup 目录自动加载。"}</p>
              </div>
            )}

            {current.id === "mac-powerpoint-load" && (
              <PowerPointAddinGuide language={language} compact />
            )}

            {current.id === "mac-powerpoint-use" && (
              <div className="onboarding-native-ribbon-demo is-powerpoint">
                <div className="onboarding-native-ribbon-title">
                  <Presentation size={17} />
                  <strong>Microsoft PowerPoint</strong>
                  <span>VisualTeX</span>
                </div>
                <div className="onboarding-native-ribbon-tools">
                  <span><Presentation size={17} /><small>{isEn ? "New formula" : "新建公式"}</small></span>
                  <span><Check size={16} /><small>{isEn ? "Edit selected" : "编辑所选公式"}</small></span>
                  <span><Trash2 size={16} /><small>{isEn ? "Delete selected" : "删除所选公式"}</small></span>
                </div>
                <p>{isEn ? "Double-click an existing VisualTeX formula to edit it again." : "双击已有 VisualTeX 公式，也可以直接重新打开编辑器。"}</p>
              </div>
            )}

            {current.id === "windows-office-manage" && (
              <div className="onboarding-windows-office-demo">
                <span>
                  <Settings2 size={21} />
                  <strong>{isEn ? "Windows Office integration" : "Windows Office 集成"}</strong>
                </span>
                <div>
                  <span><Power size={19} /><strong>{isEn ? "Stop companion now" : "停止当前伴侣服务"}</strong></span>
                  <span><ToggleLeft size={19} /><strong>{isEn ? "Disable startup" : "关闭开机启动"}</strong></span>
                  <span><Trash2 size={19} /><strong>{isEn ? "Remove OLE manifest" : "移除 OLE manifest"}</strong></span>
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
