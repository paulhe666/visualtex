import { useEffect, useRef, useState } from "react";
import {
  ArrowLeft,
  ArrowRight,
  Check,
  Code2,
  Download,
  Keyboard,
  Menu,
  PanelLeft,
  RefreshCw,
  ScanLine,
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

interface StepCopy {
  title: string;
  description: string;
}

const copy: Record<Language, StepCopy[]> = {
  cn: [
    { title: "欢迎使用 VisualTeX", description: "用熟悉的方式输入公式，需要时随时查看源码。" },
    { title: "从公式库开始", description: "选择结构或符号，它会直接插入当前光标。" },
    { title: "保持双手在键盘上", description: "几个按键就能完成换行、跳转和删除。" },
    { title: "切换 LaTeX 代码格式", description: "从顶部选择单公式或多公式环境；下方源码区和复制结果会立即按所选格式更新。" },
    { title: "第一次使用 OCR", description: "完整 macOS 包已内置 Python、PaddleOCR 与默认 M 模型；首次安装只在本机校验并解压，不需要联网。" },
    { title: "之后直接粘贴图片", description: "环境准备好后，把光标放进公式框即可粘贴图片识别；高速 S 与高精度 L 可在设置中导入独立离线模型包。" },
    { title: "随时检查更新", description: "打开左上角菜单，选择“检查更新”；也可以在设置中执行同一操作。" },
  ],
  en: [
    { title: "Welcome to VisualTeX", description: "Write formulas naturally and inspect the source whenever you need it." },
    { title: "Start from the formula library", description: "Choose a structure or symbol to insert it at the cursor." },
    { title: "Keep your hands on the keyboard", description: "A few keys cover line creation, navigation, and deletion." },
    { title: "Switch the LaTeX code format", description: "Choose an independent or combined environment from the top bar. The source panel and copied output update immediately." },
    { title: "First-time OCR setup", description: "The complete macOS package includes Python, PaddleOCR, and the default M model. First-time setup only verifies and extracts local archives; no network is required." },
    { title: "Paste images directly afterward", description: "Once ready, paste an image into a formula field. Optional offline S and L model packs can be imported from Settings." },
    { title: "Check for updates anytime", description: "Open the top-left menu and choose “Check for updates”. The same action is also available in Settings." },
  ],
};

export function OnboardingTour({ open, language, onFinish }: Props) {
  const [step, setStep] = useState(0);
  const dialogRef = useRef<HTMLElement>(null);
  const isEn = language === "en";
  const steps = copy[language];
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

  if (!open) return null;

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
            <h2 id="onboarding-title">{steps[step].title}</h2>
            <p>{steps[step].description}</p>
          </div>

          <div className={`onboarding-stage step-${step}`}>
            {step === 0 && (
              <div className="onboarding-welcome-mark">
                <span><VisualTeXLogo className="onboarding-welcome-logo" /></span>
                <div>
                  <strong>VisualTeX</strong>
                  <small>{isEn ? "Formula workspace for macOS" : "macOS 公式工作台"}</small>
                </div>
              </div>
            )}

            {step === 1 && (
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

            {step === 2 && (
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

            {step === 3 && (
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

            {step === 4 && (
              <div className="onboarding-ocr-setup-demo">
                <span>
                  <ScanLine size={20} />
                  <strong>{isEn ? "Open Formula image OCR" : "打开“图片公式识别”"}</strong>
                  <small>{isEn ? "From the app menu" : "从应用菜单进入"}</small>
                </span>
                <i><ArrowRight size={15} /></i>
                <span>
                  <Download size={20} />
                  <strong>{isEn ? "Install runtime" : "安装 OCR 环境"}</strong>
                  <small>{isEn ? "One-time setup" : "只需安装一次"}</small>
                </span>
                <i><ArrowRight size={15} /></i>
                <span>
                  <Download size={20} />
                  <strong>{isEn ? "Download model" : "首次下载模型"}</strong>
                  <small>{isEn ? "On first recognition" : "第一次识别时进行"}</small>
                </span>
              </div>
            )}

            {step === 5 && (
              <div className="onboarding-paste-demo">
                <div className="onboarding-paste-field">
                  <span className="onboarding-paste-caret" />
                  <small>{isEn ? "Formula field" : "公式输入框"}</small>
                </div>
                <span className="onboarding-paste-shortcut">
                  <ScanLine size={20} />
                  <strong>{isEn ? "Paste formula image" : "粘贴公式图片"}</strong>
                  <kbd>⌘V</kbd>
                </span>
                <i><ArrowRight size={15} /></i>
                <span className="onboarding-paste-result">
                  <Code2 size={20} />
                  <strong>{isEn ? "Inserted at saved cursor" : "插回原光标位置"}</strong>
                  <MathPreview latex="\\frac{a+b}{c}" />
                </span>
              </div>
            )}

            {step === 6 && (
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
            {steps.map((_, index) => (
              <span key={index} className={index === step ? "is-active" : index < step ? "is-complete" : ""} />
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
