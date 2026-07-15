import { useEffect, useRef, useState } from "react";
import {
  ArrowLeft,
  ArrowRight,
  Check,
  Code2,
  Copy,
  Keyboard,
  PanelLeft,
  Save,
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
    {
      title: "欢迎使用 VisualTeX",
      description: "用熟悉的方式输入公式，需要时随时查看源码。",
    },
    {
      title: "从公式库开始",
      description: "选择结构或符号，它会直接插入当前光标。",
    },
    {
      title: "保持双手在键盘上",
      description: "通过键盘完成换行、候选选择、结构跳转和跨行移动。",
    },
    {
      title: "切换 LaTeX 代码格式",
      description:
        "从顶部选择单公式或多公式环境；下方源码区和复制结果会立即按所选格式更新。",
    },
    {
      title: "复制、保存与继续编辑",
      description:
        "复制所需的 LaTeX 格式，或把完整文档保存到本地浏览器设备中。",
    },
  ],
  en: [
    {
      title: "Welcome to VisualTeX",
      description:
        "Write formulas naturally and inspect the source whenever you need it.",
    },
    {
      title: "Start from the formula library",
      description: "Choose a structure or symbol to insert it at the cursor.",
    },
    {
      title: "Keep your hands on the keyboard",
      description:
        "Create rows, choose candidates, move through structures, and navigate between formula rows from the keyboard.",
    },
    {
      title: "Switch the LaTeX code format",
      description:
        "Choose an independent or combined environment from the top bar. The source panel and copied output update immediately.",
    },
    {
      title: "Copy, save, and keep editing",
      description:
        "Copy the LaTeX format you need or save the complete document to the current browser device.",
    },
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
            <span>
              <VisualTeXLogo className="onboarding-brand-logo" />
            </span>
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
                <span>
                  <VisualTeXLogo className="onboarding-welcome-logo" />
                </span>
                <div>
                  <strong>VisualTeX</strong>
                  <small>
                    {isEn
                      ? "Formula workspace for the web"
                      : "网页版公式工作台"}
                  </small>
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
                  {[
                    "\\frac{a}{b}",
                    "\\sqrt{x}",
                    "\\int_a^b f(x)\\,dx",
                    "\\sum_{i=1}^{n} a_i",
                  ].map((latex) => (
                    <span key={latex}>
                      <MathPreview latex={latex} />
                    </span>
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
                  <span>
                    <Keyboard size={14} />
                    <kbd>Enter</kbd>
                    <small>{isEn ? "New line" : "新建一行"}</small>
                  </span>
                  <span>
                    <kbd>↑ ↓</kbd>
                    <small>{isEn ? "Switch rows" : "切换公式行"}</small>
                  </span>
                  <span>
                    <kbd>Tab</kbd>
                    <small>{isEn ? "Next field" : "下个位置"}</small>
                  </span>
                </div>
              </div>
            )}

            {step === 3 && (
              <div className="onboarding-code-format-demo">
                <div className="onboarding-code-format-toolbar">
                  <Code2 size={16} />
                  <strong>
                    {isEn ? "LaTeX code format" : "LaTeX 代码格式"}
                  </strong>
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
                <i>
                  <ArrowRight size={15} />
                </i>
                <pre>{"\\begin{align*}\na &= b + c \\\\\\nd &= e - f\n\\end{align*}"}</pre>
              </div>
            )}

            {step === 4 && (
              <div className="onboarding-workflow-demo">
                <span>
                  <Code2 size={20} />
                  <strong>{isEn ? "Source" : "源码"}</strong>
                </span>
                <i>
                  <ArrowRight size={15} />
                </i>
                <span>
                  <Copy size={20} />
                  <strong>{isEn ? "Copy" : "复制"}</strong>
                </span>
                <i>
                  <ArrowRight size={15} />
                </i>
                <span>
                  <Save size={20} />
                  <strong>{isEn ? "Save" : "保存"}</strong>
                </span>
              </div>
            )}
          </div>
        </div>

        <footer className="onboarding-footer">
          <button
            type="button"
            className="onboarding-skip"
            onClick={onFinish}
          >
            {isEn ? "Skip" : "跳过"}
          </button>
          <div
            className="onboarding-progress"
            aria-label={isEn ? "Tutorial progress" : "教程进度"}
          >
            {steps.map((_, index) => (
              <span
                key={index}
                className={
                  index === step
                    ? "is-active"
                    : index < step
                      ? "is-complete"
                      : ""
                }
              />
            ))}
          </div>
          <div className="onboarding-actions">
            {step > 0 && (
              <button
                type="button"
                className="secondary-button"
                onClick={() => setStep((value) => value - 1)}
              >
                <ArrowLeft size={15} />
                {isEn ? "Back" : "上一步"}
              </button>
            )}
            <button
              type="button"
              className="primary-button"
              onClick={() =>
                lastStep ? onFinish() : setStep((value) => value + 1)
              }
            >
              {lastStep ? <Check size={15} /> : null}
              {lastStep
                ? isEn
                  ? "Start editing"
                  : "开始使用"
                : isEn
                  ? "Continue"
                  : "继续"}
              {!lastStep ? <ArrowRight size={15} /> : null}
            </button>
          </div>
        </footer>
      </section>
    </div>
  );
}
